import argparse
import logging
from pathlib import Path
from zipfile import ZipFile, BadZipFile

import pandas as pd

from blearn import utils as u


def metadata_from_logs_anon(root: Path, /) -> pd.DataFrame:
    md_files = [
        u.msg_load(path, fname=path.name) for path in sorted(root.glob("*.txt"))
    ]
    df = (
        pd.DataFrame.from_records(md_files)
        .assign(submission_id=lambda x: x["id"].str.replace("Receipt ID: ", ""))
        .drop("id", axis=1)
    )
    # arrow datetimes are not supported in pandas
    df["datetime"] = pd.to_datetime(df["datetime"].apply(lambda x: x.isoformat()))
    return df


def prepare_project_anon(
    ini_ids: str | Path,
    ini_zip: str | Path,
    root_end: str | Path,
    safe: bool = True,
) -> tuple[Path, pd.DataFrame]:
    ini_ids, ini_zip, root_end = Path(ini_ids), Path(ini_zip), Path(root_end)
    if not (ini_ids.exists() or ini_zip.exists()):
        raise ValueError("ini_* file(s) do not exist")
    if not (root_end.exists() and root_end.is_dir()):
        raise ValueError(f"Not an existing folder: {str(root_end)}")
    if safe and any(root_end.iterdir()):
        raise ValueError("Project output path is not empty. Operation aborted.")

    # 1) Load grade template (mapping of student_ids to submission_ids)
    df_grades_tpl = pd.read_excel(ini_ids, index_col=0).assign(
        student_id=lambda x: x["uid"].str.extract(r"(\d+)"),
        submission_id=lambda x: x["rcp"].str.extract(r"Receipt: (.+)"),
    )

    # 2) Unpack submissions (1 zip bundle to zip files (1 per submission))
    path_files = root_end / "submission_files"
    assignment_name, df_logs = u.unpack_submissions(
        path_files, ini_zip, f_md=metadata_from_logs_anon
    )

    # 3) Pack non-zip submissions into their own zip files (and add metadata)
    df_logs = u.pack_unexpected(df_logs, path_files)

    # 4) Merge tables
    keep_cols = [
        "student_id",
        "submission_id",
        "assignment",
        "submission_field",
        "submission_comment",
        "log",
        "zip",
        "current_mark",
    ]
    df_all_tpl = pd.merge(
        df_grades_tpl,
        df_logs.rename(columns={"pack": "zip"}),
        on=["submission_id"],
    ).loc[:, keep_cols]

    # 5) Erase logs
    for fname in df_all_tpl["log"].tolist():
        (path_files / fname).unlink()
    df_all_tpl.drop(["log"], axis=1, inplace=True)

    # 6) Extract zip files with naming convention
    submission = {}
    errors = 0
    for idx, fname in df_all_tpl.set_index("student_id")["zip"].to_dict().items():
        fzip = path_files / fname
        fdir = path_files / Path(idx)
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
    df_all_tpl["submission"] = df_all_tpl["zip"].map(submission)
    df_all_tpl.drop(["zip"], axis=1, inplace=True)

    # 7) Enhance ease of use in Excel: hyperlink to folder
    df_all_tpl["submission"] = df_all_tpl["submission"].apply(
        lambda x: "" if pd.isna(x) else u.HYPERLINK_TPL.format(x)
    )

    # 8) Wrap up and write final table
    name = "template-" + assignment_name.lower().replace(" ", "_") + ".xlsx"
    f = root_end / name
    logging.debug(f"Writing DataFrame to {str(f)}")
    u._df_to_excel(df_all_tpl, f, group_icols=[2, 3, 4, 5])
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
    prepare_project_anon(
        ini_xls,
        ini_zip,
        root_end,
        safe=not args.force,
    )
    logging.info("END.")


if __name__ == "__main__":
    main()
