# coding=utf-8
import logging
from pathlib import Path
from typing import Callable
from zipfile import ZipFile, BadZipFile

import fire
import numpy as np
import pandas as pd
from parse import parse
from IPython.core.display_functions import display

from blearn import utils as u


def _read_file_names(path: Path, /, mode: str = "log") -> list[str]:
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
    is_zip = True if path.name.endswith(".zip") else False
    if is_zip:
        with ZipFile(path) as z:
            f_names = z.namelist()
    else:
        f_names = [f.name for f in sorted(path.glob("*"))]
    logs = [f_name for f_name in f_names if parse(u.TXT_DEFAULT_PATH_TXT, f_name)]
    if mode == "log":
        return logs
    elif mode == "all":
        return f_names
    elif mode == "others":
        return [f_name for f_name in f_names if f_name not in logs]
    else:
        raise ValueError(f"`{mode=}` is not supported.")


def metadata_from_logs(fzip: Path, /) -> pd.DataFrame:
    is_zip = True if fzip.name.endswith(".zip") else False
    f_names_txt = _read_file_names(fzip, mode="log")
    if is_zip:
        md_files = []
        with ZipFile(fzip) as z:
            for f_name_txt in f_names_txt:
                with z.open(f_name_txt) as zf:
                    md_files.append(u.msg_loads(zf.read().decode()))
    else:
        paths = [fzip / f_name_txt for f_name_txt in f_names_txt]
        md_files = [u.msg_load(path, fname=path.name) for path in paths]
    df = pd.DataFrame.from_records(md_files).set_index("id").sort_index()
    # arrow date-times are not supported in pandas
    df["datetime"] = pd.to_datetime(df["datetime"].apply(lambda x: x.isoformat()))
    return df


def metadata_from_filenames(fzip: Path, /, quiet: bool = False) -> pd.DataFrame:
    f_names_txt = _read_file_names(fzip, mode="log")
    f_names_other = _read_file_names(fzip, mode="others")
    # Metadata retrieved form path names
    md_paths = []
    for f_name_txt in f_names_txt:
        metadata = parse(u.TXT_DEFAULT_PATH_TXT, f_name_txt).named
        metadata["log"] = f_name_txt
        mode = "warn" if not quiet else "halt"
        f_names_others = u.get_similar_files(f_name_txt, f_names_other, mode=mode)
        metadata["submission"] = f_names_others
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
        msg = f"{auto_drop=} detected: dropping rows (use verbose=True to see affected rows)"
        msg += "\n{}"
        logging.info(msg.format(df.loc[sel, ["Last Name", "First Name"]].reset_index()))
        df = df.loc[~sel, :]
    check = np.array([("s" + x) == y for x, y in zip(df["Student ID"], df["Username"])])
    if not np.all(check):
        raise ValueError(f"Unexpected entries\n{df[~check]}")
    return df.sort_index()


def prepare_project(
    p_xls: str | Path,
    p_zip: str | Path,
    d_out: str | Path,
    keep: str = "last",
    drop_usernames: list[str] | None = None,
    drop_empty: bool = False,
    drop_callback: Callable | None = None,
    safe: bool = True,
) -> tuple[Path, pd.DataFrame]:
    """
    Prepare marking project (n.b. non-anonymous marking).

    Parameters
    ----------
    p_xls
        Path to xls grading template spreadsheet from Learn.
    p_zip
        Path to bulk zip file downloaded from Learn for the assignment.
    d_out
        Folder in which to save all output files.
    keep
        TODO
    drop_usernames
        TODO
    drop_empty
        TODO
    drop_callback
        TODO
    safe
        Halt execution if `d_out` is not empty.

    Returns
    -------
    Path
        Path to exported spreadsheet for marking.
    pd.DataFrame
        Table for marking.

    """
    p_xls, p_zip, d_out = Path(p_xls), Path(p_zip), Path(d_out)
    if not (p_xls.exists() or p_zip.exists()):
        raise ValueError("ini_* file(s) do not exist")
    if not (d_out.exists() and d_out.is_dir()):
        raise ValueError(f"Not an existing folder: {str(d_out)}")
    if safe and any(d_out.iterdir()):
        raise ValueError("Project output path is not empty. Operation aborted.")

    # 1) Load grade template (xls from Learn offline marking)
    df_grades_tpl = read_xls(p_xls, drop_usernames=drop_usernames)

    # 2) Unpack submissions (1 zip bundle to zip files (1 per submission))
    path_files = d_out / "submission_files"
    assignment_name, df_logs = u.unpack_submissions(
        path_files, p_zip, f_md=metadata_from_logs
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
    for f_name in df_aux["log"].tolist():
        (path_files / f_name).unlink()

    # Extract zip files with naming convention
    submission = {}
    errors = 0
    for idx, f_name in df_aux["submission"].to_dict().items():
        fzip = path_files / f_name
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
        submission[f_name] = str(fdir.relative_to(d_out))
    df_all_tpl["submission"] = df_all_tpl["submission"].map(submission)

    # Enhance ease of use in Excel: hyperlink to folder
    df_all_tpl["submission"] = df_all_tpl["submission"].apply(
        lambda x: "" if pd.isna(x) else u.XLSX_LINK.format(x)
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
    f = d_out / name
    logging.debug(f"Writing DataFrame to {str(f)}")
    u.df_to_excel(df_all_tpl, f, group_icols=[2, 3])
    if errors > 0:
        print("error / warnings appeared processing submissions. See the log.")
    return f, df_all_tpl


def main(
    p_xls: str | Path,
    p_zip: str | Path,
    d_out: str | Path,
    p_log: Path | None = None,
    safe: bool = True,
    debug: bool = False,
    **kwargs,
):
    """
    Prepare marking project (n.b., non-anonymous marking).

    Parameters
    ----------
    p_xls
        Path to xls grading template spreadsheet from Learn.
    p_zip
        Path to bulk zip file downloaded from Learn for the assignment.
    d_out
        Folder in which to save all output files.
    p_log
        Path to log file to write output to.
    safe
        Halt execution if `d_out` is not empty.
    debug
        Activate debug mode for the logger.

    """
    d_out = Path(d_out)
    d_out.mkdir(exist_ok=True)
    if p_log is None:
        p_log = d_out.parent / f"blearn-{Path(__file__).stem}.log"
    u.setup_logger(path=p_log, shout=True, debug=debug)
    logging.info("INI.")
    prepare_project(p_xls=p_xls, p_zip=p_zip, d_out=d_out, safe=safe, **kwargs)
    logging.info("END.")


def cli():
    fire.Fire(main)


if __name__ == "__main__":
    fire.Fire(main)
