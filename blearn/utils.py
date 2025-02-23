# coding=utf-8
import logging
import shutil
from contextlib import nullcontext
from pathlib import Path
from typing import ContextManager, Callable
from zipfile import ZipFile

import pandas as pd
import arrow
from bs4 import BeautifulSoup
from parse import parse

DEFAULT_ZIPFILE_SUFFIX = "-generated"  # n.b. this is without the extension
XLSX_LINK = (
    '=HYPERLINK(LEFT(CELL("filename",A1),FIND("[",CELL("filename",A1))-1)&"{}", "link")'
)
TXT_TPL = """\
Name: {name} ({id})
Assignment: {assignment}
Date Submitted: {datetime_raw}
Current Mark: {current_mark}

Submission Field:
{submission_field}

Comments:
{submission_comment}

Files:
{files}
"""
TXT_TPL_FILES = "\tOriginal filename: {fname_original}\n\tFilename: {fname_blearn}"
TXT_DEFAULT_COMMENT = "There are no student comments for this assignment."
TXT_DEFAULT_SUBMISSION_FIELD = (
    "There is no student submission text data for this assignment."
)
TXT_DEFAULT_FILES = "No files were attached to this submission."
TXT_DEFAULT_PATH_TXT = (
    "{assignment}_{id}_attempt_{year}-{month}-{day}-{hour}-{minute}-{second}.txt"
)
TXT_DEFAULT_PATH_SUBMISSION = TXT_DEFAULT_PATH_TXT.replace(".txt", "_{fname}")


def setup_logger(path: Path, shout: bool = True, debug: bool = False):
    """
    Set up logger and save to `path`.

    Parameters
    ----------
    path
        Path where to save log.
    shout
        Print path.
    debug
        Activate debug mode.

    Returns
    -------
    logging.FileHandler
        Logger's file handler

    """

    def filter_parse(record):
        if record.name == "parse":
            return False
        return True

    if shout:
        print(f"Setting up logger at {str(path)}")

    level = logging.DEBUG if debug else logging.INFO
    formatter = logging.Formatter(
        "{asctime} - {module} - {name} - {levelname} - {funcName}:{lineno} - "
        "{message}",
        style="{",
    )
    fh = logging.FileHandler(str(path), mode="w")
    fh.setLevel(level)
    fh.setFormatter(formatter)
    fh.addFilter(filter_parse)
    logging.basicConfig(handlers=[fh], level=level)
    return fh


def col_width_excel(df: pd.DataFrame, with_index: bool = True) -> list[int]:
    """
    Create a list of column widths based on column names.

    If the index is considered,
    then its width is based on its contents as well.

    Notes
    -----
    Code adapted from Cole Diamond at Stack Overflow
    ([here](https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter)).

    """
    if with_index:
        index = [max([len(str(v)) for v in df.index] + [len(str(df.index.name))])]
    else:
        index = []
    others = [len(str(col)) for col in df.columns]
    return index + others


def df_to_excel(
    df: pd.DataFrame,
    path: Path,
    adjust_colwidth: bool = True,
    with_index: bool = True,
    group_icols: list[int] | None = None,
):
    if group_icols is None:
        group_icols = []
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        sheet_name = "grades"
        df.to_excel(writer, sheet_name=sheet_name, index=with_index)
        writer.sheets[sheet_name].freeze_panes(1, 1)
        for i, width in enumerate(col_width_excel(df, with_index=with_index)):
            group = dict(level=1) if i in group_icols else None
            width_ = width if adjust_colwidth else None
            writer.sheets[sheet_name].set_column(i, i, width_, None, group)


def get_similar_files(
    template: str,
    candidates: list[str],
    mode: str = "warn",
) -> list[str]:
    template_stem, _ = template.rsplit(".", maxsplit=1)
    f_others = [
        f for f in candidates if (f.startswith(template_stem) and f != template)
    ]
    if not f_others and mode is not None:
        msg = f"Could not find similar file(s) for `{template}`: {f_others=}"
        if mode == "warn":
            logging.error(msg)
        elif mode == "halt":
            raise ValueError(msg)
        else:
            ValueError(f"{mode=} is not a valid option")
    return f_others


def pack_files(f_pack, /, root: Path, files: list[str | Path]):
    path_tmp = root / "_test"
    path_tmp.mkdir(exist_ok=True)
    path_zip = path_tmp / f_pack.name
    path_zip.mkdir(exist_ok=False)
    for f in files:
        shutil.move(root / f, path_zip / f)
    shutil.make_archive(f_pack, "zip", path_tmp)
    shutil.rmtree(path_tmp)


def parse_datetime(x: str, /) -> arrow.Arrow:
    """Parse datetime routine as Learn uses non-standard format and timezone info."""
    mapper = {"BST": "Europe/London"}
    dt_str = x.split(", ", maxsplit=1)[1].replace("o'clock ", "")
    dt_str, dt_tz = dt_str.rsplit(" ", maxsplit=1)
    if dt_tz in mapper:
        dt_tz = mapper[dt_tz]
    try:
        return arrow.get(dt_str, "D MMMM YYYY HH:mm:ss", tzinfo=dt_tz)
    except arrow.ParserError as e:
        raise ValueError(f"Could not parse datetime string: {dt_str=}, {dt_tz=}") from e


def msg_loads(
    txt: str,
    /,
    fname: str | None = None,
    remove_empty: bool = True,
    tpl: str = TXT_TPL,
) -> dict:
    err_tpl = "Cannot parse:\n---\n{txt}\n---\nwith\n---\n{tpl}\n---\n"
    try:
        data = parse(tpl, txt).named
    except AttributeError:
        raise ValueError(err_tpl.format(txt=txt, tpl=tpl))
    data["datetime"] = parse_datetime(data["datetime_raw"])
    del data["datetime_raw"]
    # Remove empty fields
    if remove_empty:
        if data["submission_field"] == TXT_DEFAULT_SUBMISSION_FIELD:
            data["submission_field"] = ""
        if data["submission_comment"] == TXT_DEFAULT_COMMENT:
            data["submission_comment"] = ""
        if data["files"] == TXT_DEFAULT_FILES:
            data["files"] = ""
    fnames_original = []
    fnames_blearn = []
    for part in data["files"].split("\n\n"):
        if part.strip() == "":
            continue
        try:
            data_part = parse(TXT_TPL_FILES, part).named
        except AttributeError:
            raise ValueError(err_tpl.format(txt=part, tpl=TXT_TPL_FILES))
        fnames_original.append(data_part["fname_original"].strip())
        fnames_blearn.append(data_part["fname_blearn"].strip())
    data["fnames_original"] = fnames_original
    data["fnames_blearn"] = fnames_blearn
    data["log"] = fname if fname is not None else None
    return data


def msg_load(path_or_buffer, /, **kwargs) -> dict:
    """
    Load txt message from a path or a buffer.

    Notes
    -----
    Code adapted from JL Pyret at Stack Overflow
    ([here](https://stackoverflow.com/questions/67416614/support-filename-path-and-buffer-input)).

    """
    if hasattr(path_or_buffer, "readline"):
        cm: ContextManager = nullcontext(path_or_buffer)
    else:
        cm = open(path_or_buffer)
    with cm as f:
        txt = f.read()
    return msg_loads(txt, **kwargs)


def extract_submission_field(txt: str, /) -> str:
    soup = BeautifulSoup(txt, features="html.parser")
    if soup.a:
        soup.a.decompose()  # remove link to submission file
    out = soup.get_text()
    return " ".join(out.splitlines())  # new lines to spaces


def pack_unexpected(df_logs: pd.DataFrame, path_files: Path) -> pd.DataFrame:
    """Pack non-zip submissions into their own zip file"""
    md = df_logs[["log", "fnames_blearn"]].to_dict(orient="index")
    for k, v in md.items():
        f_submission = v["fnames_blearn"]
        if len(f_submission) == 1 and f_submission[0].endswith(".zip"):
            md[k]["pack"] = f_submission[0]
            continue
        # Every other case needs to be packed
        basename = Path(v["log"]).stem
        if len(f_submission) == 0:
            # The submission file is missing: generate a mockup file for completeness
            f_shim = f"{basename}-generated-missing_submission_{k}.txt"
            f_shim = path_files / f_shim
            f_shim.write_text("File automatically generated.")
            f_submission = [f_shim.name]
        # note that shutil.make_archive adds the extension to the name
        f_pack = path_files / (basename + DEFAULT_ZIPFILE_SUFFIX)
        pack_files(f_pack, root=path_files, files=f_submission)
        md[k]["zip"] = f_pack.with_suffix(".zip").name
    df_logs["pack"] = df_logs.index.map({k: v["pack"] for k, v in md.items()})
    return df_logs


def unpack_submissions(
    path_files: Path, ini_zip: Path, f_md: Callable
) -> tuple[str, pd.DataFrame]:
    """Unpacks Learn's bulk submission zip file and extracts submission's log info"""

    if path_files.exists():
        shutil.rmtree(path_files)
    else:
        path_files.mkdir(exist_ok=True)

    # 1) Extract bundle of submissions
    with ZipFile(ini_zip, mode="r") as zip_ref:
        zip_ref.extractall(path_files)

    # 2) Get metadata from submission logs
    df_logs = f_md(path_files)
    df_logs["datetime"] = df_logs["datetime"].dt.tz_localize(None)
    df_logs["submission_field"] = df_logs["submission_field"].apply(
        extract_submission_field
    )
    try:
        [assignment_name] = df_logs["assignment"].unique()
    except ValueError as e:
        raise ValueError("Data seems to host more than one assignment") from e
    return assignment_name, df_logs
