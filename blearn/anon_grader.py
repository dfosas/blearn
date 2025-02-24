# coding=utf-8
import logging
from pathlib import Path
from zipfile import ZipFile, BadZipFile

import fire
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
    # arrow date-times are not supported in pandas
    df["datetime"] = pd.to_datetime(df["datetime"].apply(lambda x: x.isoformat()))
    return df


def prepare_project_anon(
    p_ids: str | Path,
    p_zip: str | Path,
    d_out: str | Path,
    safe: bool = True,
) -> tuple[Path, pd.DataFrame]:
    """
    Prepare anonymous marking project.

    Parameters
    ----------
    p_ids
        Path to xlsx spreadsheet with IDs mapping.
        See dedicated CLI tool to generate one from web automation.
    p_zip
        Path to bulk zip file downloaded from Learn for the assignment.
    d_out
        Folder in which to save all output files.
    safe
        Halt execution if `d_out` is not empty.

    Returns
    -------
    Path
        Path to exported spreadsheet for marking.
    pd.DataFrame
        Table for marking.

    """
    p_ids, p_zip, d_out = Path(p_ids), Path(p_zip), Path(d_out)
    if not (p_ids.exists() or p_zip.exists()):
        raise ValueError("ini_* file(s) do not exist")

    if not (d_out.exists() and d_out.is_dir()):
        raise ValueError(f"Not an existing folder: {str(d_out)}")

    if safe and any(d_out.iterdir()):
        raise ValueError("Project output path is not empty. Operation aborted.")

    # 1) Load grade template (mapping of student_ids to submission_ids)
    logging.info("loading grade template")
    df_grades_tpl = pd.read_excel(p_ids, index_col=0).assign(
        student_id=lambda x: x["uid"].str.extract(r"(\d+)"),
        submission_id=lambda x: x["rcp"].str.extract(r"Receipt: (.+)"),
    )

    # 2) Unpack submissions (1 zip bundle to zip files (1 per submission))
    logging.info("unpacking bulk submissions")
    path_files = d_out / "submission_files"
    assignment_name, df_logs = u.unpack_submissions(
        path_files, p_zip, f_md=metadata_from_logs_anon
    )

    # 3) Pack non-zip submissions into their own zip files (and add metadata)
    logging.info("retrieving logs from individual submissions")
    df_logs = u.pack_unexpected(df_logs, path_files)

    # 4) Merge tables
    logging.info("merging IDs mapping with info in submission logs")
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
        df_logs,
        on=["submission_id"],
    ).loc[:, keep_cols]

    # 5) Erase logs
    logging.info("removing submissions logs")
    for f_name in df_all_tpl["log"].tolist():
        (path_files / f_name).unlink()
    df_all_tpl.drop(["log"], axis=1, inplace=True)

    # 6) Extract zip files with naming convention
    logging.info("extracting individual submissions into normalised folders")
    submission = {}
    errors = 0
    for idx, f_name in df_all_tpl.set_index("student_id")["zip"].to_dict().items():
        fzip = path_files / f_name
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
        submission[f_name] = str(fdir.relative_to(d_out))
    df_all_tpl["submission"] = df_all_tpl["zip"].map(submission)
    df_all_tpl.drop(["zip"], axis=1, inplace=True)

    # 7) Enhance ease of use in Excel: hyperlink to folder
    logging.info("generating hyperlinks for Excel")
    df_all_tpl["submission"] = df_all_tpl["submission"].apply(
        lambda x: "" if pd.isna(x) else u.XLSX_LINK.format(x)
    )

    # 8) Wrap up and write final table
    logging.info("saving final table ")
    name = "template-" + assignment_name.lower().replace(" ", "_") + ".xlsx"
    f = d_out / name
    logging.info(f"Writing DataFrame to {str(f)}")
    u.df_to_excel(df_all_tpl, f, group_icols=[2, 3, 4, 5])
    if errors > 0:
        print("error / warnings appeared processing submissions. See the log.")
    return f, df_all_tpl


def main(
    p_ids: str | Path,
    p_zip: str | Path,
    d_out: str | Path,
    p_log: Path | None = None,
    safe: bool = True,
    debug: bool = False,
):
    """
    Prepare anonymous marking project.

    Parameters
    ----------
    p_ids
        Path to xlsx spreadsheet with IDs mapping.
        See dedicated CLI tool to generate one from web automation.
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
    prepare_project_anon(p_ids=p_ids, p_zip=p_zip, d_out=d_out, safe=safe)
    logging.info("END.")


def cli():
    fire.Fire(main)


if __name__ == "__main__":
    fire.Fire(main)
