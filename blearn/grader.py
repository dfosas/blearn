import argparse
import logging
from pathlib import Path
from typing import Callable, Optional
from zipfile import ZipFile, BadZipFile

import numpy as np
import pandas as pd
from parse import parse
from IPython.core.display_functions import display

from blearn import utils as u


def _read_fnames(path: Path, /, mode: str = "log") -> list[str]:
    """
    Read file names from Path.

    Parameters
    ----------
    path
        Can be a path to a zip archive or to its extracted folder.
    mode
        One of `['all', 'log', 'others']`.
        * 'log': retrieves data from the txt log file.
        * 'all': retrieves every file.
        * 'others': retrieves every file that is not a txt log file.

    """
    iszip = True if path.name.endswith(".zip") else False
    if iszip:
        with ZipFile(path) as z:
            fnames = z.namelist()
    else:
        fnames = [f.name for f in sorted(path.glob("*"))]
    logs = [fname for fname in fnames if parse(u.TXT_DEFAULT_PATH_TXT, fname)]
    if mode == "log":
        return logs
    elif mode == "all":
        return fnames
    elif mode == "others":
        return [fname for fname in fnames if fname not in logs]
    else:
        raise ValueError(f"`{mode=}` is not supported.")


def metadata_from_logs(fzip: Path, /) -> pd.DataFrame:
    iszip = True if fzip.name.endswith(".zip") else False
    fnames_txt = _read_fnames(fzip, mode="log")
    if iszip:
        md_files = []
        with ZipFile(fzip) as z:
            for fname_txt in fnames_txt:
                with z.open(fname_txt) as zf:
                    md_files.append(u.msg_loads(zf.read().decode()))
    else:
        paths = [fzip / fname_txt for fname_txt in fnames_txt]
        md_files = [u.msg_load(path, fname=path.name) for path in paths]
    df = pd.DataFrame.from_records(md_files).set_index("id").sort_index()
    # arrow datetimes are not supported in pandas
    df["datetime"] = pd.to_datetime(df["datetime"].apply(lambda x: x.isoformat()))
    return df


def metadata_from_filenames(fzip: Path, /, quiet: bool = False) -> pd.DataFrame:
    fnames_txt = _read_fnames(fzip, mode="log")
    fnames_other = _read_fnames(fzip, mode="others")
    # Metadata retrieved form path names
    md_paths = []
    for fname_txt in fnames_txt:
        metadata = parse(u.TXT_DEFAULT_PATH_TXT, fname_txt).named
        metadata["log"] = fname_txt
        mode = "warn" if not quiet else "halt"
        fnames_others = u._get_similar_files(fname_txt, fnames_other, mode=mode)
        metadata["submission"] = fnames_others
        md_paths.append(metadata)
    df = pd.DataFrame.from_records(md_paths).set_index("id")
    datetime_cols = ["year", "month", "day", "hour", "minute", "second"]
    df["datetime"] = pd.to_datetime(df[datetime_cols])
    df = df.drop(datetime_cols, axis=1).sort_index()
    if (x := df["assignment"].unique()).shape[0] != 1:
        msg = f"More than one candidate for assignment name: {x}"
        raise ValueError(msg)
    return df


def read_xls(
    f_xls: Path,
    /,
    drop_usernames: list[str] | None = None,
    auto_drop: str | None = "previewuser",
) -> pd.DataFrame:
    if drop_usernames is None:
        drop_usernames = []
    df = (
        pd.read_csv(
            f_xls,
            sep="\t",
            encoding="utf-16-le",
            dtype={
                "Last Name": str,
                "First Name": str,
                "Username": str,
                "Student ID": str,
                "Marking Notes": str,
                "Feedback to Learner": str,
            },
            parse_dates=["Last Access"],
            keep_default_na=False,
        )
        .assign(id=lambda x: x["Username"])
        .loc[lambda x: ~x["Username"].isin(drop_usernames)]
        .set_index("id")
    )
    if auto_drop and np.any(sel := df["Username"].str.contains(auto_drop)):
        msg = "{autodrop=} detected: dropping rows (use verbose=True to see affected rows)\n{}"
        logging.info(msg.format(df.loc[sel, ["Last Name", "First Name"]].reset_index()))
        df = df.loc[~sel, :]
    check = np.array([("s" + x) == y for x, y in zip(df["Student ID"], df["Username"])])
    if not np.all(check):
        raise ValueError(f"Unexpected entries\n{df[~check]}")
    return df.sort_index()


def prepare_project(
    ini_xls: str | Path,
    ini_zip: str | Path,
    root_end: str | Path,
    /,
    keep: str = "last",
    drop_usernames: Optional[list[str]] = None,
    drop_empty: bool = False,
    drop_callback: Optional[Callable] = None,
    safe: bool = True,
) -> tuple[Path, pd.DataFrame]:
    ini_xls, ini_zip, root_end = Path(ini_xls), Path(ini_zip), Path(root_end)
    if not (ini_xls.exists() or ini_zip.exists()):
        raise ValueError("ini_* file(s) do not exist")
    if not (root_end.exists() and root_end.is_dir()):
        raise ValueError(f"Not an existing folder: {str(root_end)}")
    if safe and any(root_end.iterdir()):
        raise ValueError("Project output path is not empty. Operation aborted.")

    # 1) Load grade template (xls from Learn offline marking)
    df_grades_tpl = read_xls(ini_xls, drop_usernames=drop_usernames)

    # 2) Unpack submissions (1 zip bundle to zip files (1 per submission))
    path_files = root_end / "submission_files"
    assignment_name, df_logs = u.unpack_submissions(
        path_files, ini_zip, f_md=metadata_from_logs
    )

    # 2.5) Handle multiple submissions
    df_logs = df_logs.sort_index().sort_values(by="log")
    is_duplicate = df_logs.index.duplicated(keep=keep)
    for _, row in df_logs.loc[is_duplicate].iterrows():
        for p in row["fnames_blearn"]:
            (path_files / p).unlink()
        (path_files / row["log"]).unlink()
    df_logs = df_logs.loc[~is_duplicate]
    # 3) Pack non-zip submissions into their own zip files (and add metadata)
    df_logs = u.pack_unexpected(df_logs, path_files)

    # Create macro table template
    # 1) Retrieve information from fixed submission folder file
    df_files = metadata_from_filenames(path_files)
    check = df_files["submission"].apply(
        lambda x: (isinstance(x, list) and len(x) == 1)
    )
    if not np.all(check):
        with pd.option_context("display.max_rows", None, "display.max_colwidth", None):
            display(df_files["submission"])
        msg = "BUG. Unexpected `submission` at this point: not a list of size 1."
        raise ValueError(msg)
    df_files.loc[:, "submission"] = df_files["submission"].apply(lambda x: x[0])
    # 2) Merge tables
    df_md = pd.merge(
        df_logs.reset_index().rename(columns={"pack": "submission"}),
        df_files.reset_index(),
        on=["id", "assignment", "log", "datetime", "submission"],
    ).set_index("id")
    if not (df_logs.shape[0] == df_files.shape[0] == df_md.shape[0]):
        raise ValueError("BUG: metadata merge gives unexpected results")
    df_all_tpl = pd.merge(
        df_grades_tpl, df_md, how="outer", left_index=True, right_index=True
    )
    if not (df_grades_tpl.shape[0] == df_all_tpl.shape[0]):
        raise ValueError("BUG: metadata merge gives unexpected results")
    drop_cols = [
        "Username",
        "First Name",
        "Student ID",
        "Last Access",
        "Availability",
        "Notes Format",
        "Feedback Format",
        "name",
        "assignment",
        "files",
        "datetime",
        "fnames_original",
        "fnames_blearn",
    ]
    df_all_tpl["datetime_lastaccess"] = df_all_tpl["Last Access"]
    df_all_tpl["datetime_log"] = df_all_tpl["datetime"]
    df_all_tpl = df_all_tpl.assign(
        full_name=lambda x: x["First Name"] + " " + x["Last Name"]
    ).sort_values(by=["Last Name", "full_name"])
    try:
        df_all_tpl = df_all_tpl.drop(drop_cols, axis=1)
    except KeyError as e:
        avail = df_all_tpl.columns.tolist()
        raise KeyError(f"Not all keys are available in axis: {avail}") from e
    cols = df_all_tpl.columns.tolist()
    cols_ini = [
        "Last Name",
        "full_name",
        "datetime_lastaccess",
        "datetime_log",
        "submission",
        "current_mark",
        "submission_field",
        "submission_comment",
    ]
    cols_end = ["Marking Notes", "Feedback to Learner"]
    cols_mid = [col for col in cols if (col not in cols_ini) and (col not in cols_end)]
    df_all_tpl = df_all_tpl[cols_ini + cols_mid + cols_end]

    # Erase logs
    df_aux = df_all_tpl.loc[lambda x: ~pd.isna(x["submission"]), :]
    for fname in df_aux["log"].tolist():
        (path_files / fname).unlink()

    # Extract zip files with naming convention
    submission = {}
    errors = 0
    for idx, fname in df_aux["submission"].to_dict().items():
        fzip = path_files / fname
        fdir = path_files / Path(idx).stem
        fdir.mkdir()
        try:
            with ZipFile(fzip, mode="r") as zip_ref:
                zip_ref.extractall(fdir)
        except BadZipFile:
            errors += 1
            logging.warning(f"BadZipFile at {idx=}.")
            (fdir / "corrupt_submission.txt").touch()
        fzip.unlink()
        submission[fname] = str(fdir.relative_to(root_end))
    df_all_tpl["submission"] = df_all_tpl["submission"].map(submission)

    # Enhance ease of use in Excel: hyperlink to folder
    df_all_tpl["submission"] = df_all_tpl["submission"].apply(
        lambda x: "" if pd.isna(x) else u.HYPERLINK_TPL.format(x)
    )

    # Wrap up and write final table
    df_all_tpl.drop(["log"], axis=1, inplace=True)
    if drop_empty:
        subset = [col for col in df_all_tpl if col not in ["Last Name", "full_name"]]
        df_all_tpl = (
            df_all_tpl.replace("", float("nan"))
            .dropna(how="all", axis=0, subset=subset)
            .fillna("")
        )
    if drop_callback:
        df_all_tpl = drop_callback(df_all_tpl)
    name = "template-" + assignment_name.lower().replace(" ", "_") + ".xlsx"
    f = root_end / name
    logging.debug(f"Writing DataFrame to {str(f)}")
    u._df_to_excel(df_all_tpl, f, group_icols=[2, 3])
    if errors > 0:
        print("error / warnings appeared processing submissions. See the log.")
    return f, df_all_tpl


def main():
    parser = argparse.ArgumentParser(
        prog="blearn",
        description="Prepare files and template for marking.",
    )
    parser.add_argument(
        "--root", type=Path, default=Path.cwd(), help="Root folder for marking project."
    )
    parser.add_argument("--log", type=Path, default=False, help="Log file.")
    parser.add_argument(
        "--force", action="store_true", help="Force overwriting output folder contents."
    )
    parser.add_argument(
        "--drop_empty", action="store_true", help="Remove entries without submissions."
    )
    args = parser.parse_args()

    u._setup_logger(path=args.root / "blearn.log", debug=True)
    logging.info("INI.")
    root_ini = args.root / "blearn-1_ini"
    if not root_ini.exists():
        raise ValueError(f"Cannot find {str(root_ini)}")
    root_end = args.root / "blearn-2_out"
    root_end.mkdir(exist_ok=True)
    ini_xls = root_ini / "a.xls"
    ini_zip = root_ini / "a.zip"
    prepare_project(
        ini_xls,
        ini_zip,
        root_end,
        drop_empty=True,
        safe=not args.force,
    )
    logging.info("END.")


if __name__ == "__main__":
    main()
