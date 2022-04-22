import shutil
import warnings
from contextlib import nullcontext
from pathlib import Path
from typing import Optional, Callable
from zipfile import ZipFile, BadZipFile

import arrow
import numpy as np
import pandas as pd
from parse import parse
from IPython.core.display_functions import display


DEFAULT_ZIPFILE_SUFFIX = "-generated"  # note this is without the extension
HYPERLINK_TPL = (
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


def _col_width_excel(df: pd.DataFrame, with_index: bool = True) -> list[int]:
    """
    Create a list of column widths based on column names.

    If the index is considered,
    then its width is based on its contents as well.

    Notes
    -----
    Code adapted from [Cole Diamon@stackoverflow](https://stackoverflow.com/questions/29463274/simulate-autofit-column-in-xslxwriter).

    """
    if with_index:
        index = [max([len(str(v)) for v in df.index] + [len(str(df.index.name))])]
    else:
        index = []
    others = [len(str(col)) for col in df.columns]
    return index + others


def _df_to_excel(
    df: pd.DataFrame,
    path: Path,
    adjust_colwidth: bool = True,
    with_index: bool = True,
    group_icols: Optional[list[int]] = None,
):
    if group_icols is None:
        group_icols = []
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        sheet_name = "grades"
        df.to_excel(writer, sheet_name=sheet_name, index=with_index)
        writer.sheets[sheet_name].freeze_panes(1, 1)
        for i, width in enumerate(_col_width_excel(df, with_index=with_index)):
            group = dict(level=1) if i in group_icols else None
            width = width if adjust_colwidth else None
            writer.sheets[sheet_name].set_column(i, i, width, None, group)


def _get_similar_files(
    template: str,
    candidates: list[str],
    mode: Optional[str] = "warn",
) -> list[str]:
    template_stem, _ = template.rsplit(".", maxsplit=1)
    f_others = [
        f for f in candidates if (f.startswith(template_stem) and f != template)
    ]
    if not f_others and mode is not None:
        msg = f"Could not find similar file(s) for `{template}`: {f_others=}"
        if mode == "warn":
            warnings.warn(msg)
        elif mode == "halt":
            raise ValueError(msg)
        else:
            ValueError(f"{mode=} is not a valid option")
    return f_others


def _pack_files(f_pack, /, root: Path, files: list[str | Path]):
    path_tmp = root / "_test"
    path_tmp.mkdir(exist_ok=True)
    path_zip = path_tmp / f_pack.name
    path_zip.mkdir(exist_ok=False)
    for f in files:
        shutil.move(root / f, path_zip / f)
    shutil.make_archive(f_pack, "zip", path_tmp)
    shutil.rmtree(path_tmp)


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
    logs = [fname for fname in fnames if parse(TXT_DEFAULT_PATH_TXT, fname)]
    if mode == "log":
        return logs
    elif mode == "all":
        return fnames
    elif mode == "others":
        return [fname for fname in fnames if fname not in logs]
    else:
        raise ValueError(f"`{mode=}` is not supported.")


def _write_aloud(path: str | Path, how: Optional[Callable]):
    print(f"Writing file {str(path)} ...", end=" ")
    how(path)
    print("done.")


def msg_loads(
    txt: str, /, fname: Optional[str] = None, remove_empty: bool = True, tpl=TXT_TPL
) -> dict:
    err_tpl = "Cannot parse:\n---\n{txt}\n---\nwith\n---\n{tpl}\n---\n"
    try:
        data = parse(tpl, txt).named
    except AttributeError:
        raise ValueError(err_tpl.format(txt=txt, tpl=tpl))
    # Parse datetime
    dt_str = data["datetime_raw"].split(", ", maxsplit=1)[1].replace("o'clock ", "")
    dt_str, dt_tz = dt_str.rsplit(" ", maxsplit=1)
    data["datetime"] = arrow.get(dt_str, "D MMMM YYYY HH:mm:ss", tzinfo=dt_tz)
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
    Code adapted from [JL Pyret@stackoverflow](https://stackoverflow.com/questions/67416614/support-filename-path-and-buffer-input).

    """
    if hasattr(path_or_buffer, "readline"):
        cm = nullcontext(path_or_buffer)
    else:
        cm = open(path_or_buffer)
    with cm as f:
        txt = f.read()
    return msg_loads(txt, **kwargs)


def metadata_from_logs(fzip: Path, /) -> pd.DataFrame:
    iszip = True if fzip.name.endswith(".zip") else False
    fnames_txt = _read_fnames(fzip, mode="log")
    if iszip:
        md_files = []
        with ZipFile(fzip) as z:
            for fname_txt in fnames_txt:
                with z.open(fname_txt) as zf:
                    md_files.append(msg_loads(zf.read().decode()))
    else:
        paths = [fzip / fname_txt for fname_txt in fnames_txt]
        md_files = [msg_load(path, fname=path.name) for path in paths]
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
        metadata = parse(TXT_DEFAULT_PATH_TXT, fname_txt).named
        metadata["log"] = fname_txt
        mode = "warn" if not quiet else None
        fnames_others = _get_similar_files(fname_txt, fnames_other, mode=mode)
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


def read_xls(fxls: Path, /, quiet: bool = False, verbose: bool = False) -> pd.DataFrame:
    df = (
        pd.read_csv(
            fxls,
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
        .set_index("id")
    )
    if np.any(sel := df["Username"].str.contains("_previewuser")):
        msg = "preview user(s) detected: dropping rows (use verbose=True to see affected rows)"
        if not quiet:
            warnings.warn(msg)
        if verbose:
            print(msg)
            display(df.loc[sel, :])
        df = df.loc[~sel, :]
    assert np.all([("s" + x) == y for x, y in zip(df["Student ID"], df["Username"])])
    return df.sort_index()


def prepare_project(
    root_ini: str | Path,
    root_end: str | Path,
    /,
    drop_empty: bool = False,
    safe: bool = True,
    quiet: bool = True,
    verbose: bool = False,
) -> pd.DataFrame:
    root_ini, root_end = Path(root_ini), Path(root_end)
    if not (root_ini.exists() and root_ini.is_dir()):
        raise ValueError(f"Not an existing folder: {str(root_ini)}")
    if not (root_end.exists() and root_end.is_dir()):
        raise ValueError(f"Not an existing folder: {str(root_end)}")
    if safe:
        if any(root_end.iterdir()):
            raise ValueError(f"Project output path is not empty. Operation aborted.")

    # Grades
    [fxls] = list(root_ini.glob("*.xls"))
    df_grades_tpl = read_xls(fxls, quiet=quiet, verbose=verbose)
    if verbose:
        f = root_end / "debug-df_grades_tpl.xlsx"
        _write_aloud(f, lambda x: _df_to_excel(df_grades_tpl, x))

    # Submissions
    [fzip] = list(root_ini.glob("*.zip"))
    path_files = root_end / "submission_files"
    if path_files.exists():
        shutil.rmtree(path_files)
    else:
        path_files.mkdir(exist_ok=True)
    # 1) Extract
    with ZipFile(fzip, "r") as zip_ref:
        zip_ref.extractall(path_files)
    # 2) Get metadata from submission logs
    df_logs = metadata_from_logs(path_files)
    df_logs["datetime"] = df_logs["datetime"].dt.tz_localize(None)
    try:
        [assignment] = df_logs["assignment"].unique()
    except ValueError as e:
        raise ValueError(f"Data seems to host more than one assignment") from e
    if verbose:
        f = root_end / "debug-df_logs.xlsx"
        _write_aloud(f, lambda x: _df_to_excel(df_logs, x))
    # 3) Pack unexpected cases into zip files and reflect that in the metadata
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
        _pack_files(f_pack, root=path_files, files=f_submission)
        md[k]["pack"] = f_pack.with_suffix(".zip").name
    df_logs["pack"] = df_logs.index.map({k: v["pack"] for k, v in md.items()})
    if verbose:
        f = root_end / "debug-df_logs_pack.xlsx"
        _write_aloud(f, lambda x: _df_to_excel(df_logs, x))

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
    if verbose:
        f = root_end / "debug-df_files.xlsx"
        _write_aloud(f, lambda x: _df_to_excel(df_files, x))
    # 2) Merge information
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
        "Last Name",
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
    df_all_tpl = df_all_tpl.drop(drop_cols, axis=1)
    cols = df_all_tpl.columns.tolist()
    cols_ini = ["current_mark", "submission_field", "submission_comment", "submission"]
    cols_end = ["Marking Notes", "Feedback to Learner"]
    cols_mid = [col for col in cols if (col not in cols_ini) and (col not in cols_end)]
    df_all_tpl = df_all_tpl[cols_ini + cols_mid + cols_end]
    # 3) Erase logs and extract zip files
    df_aux = df_all_tpl.loc[lambda x: ~pd.isna(x["submission"]), :]
    for fname in df_aux["log"].tolist():
        (path_files / fname).unlink()
    submission = {}
    for id, fname in df_aux["submission"].to_dict().items():
        fzip = path_files / fname
        fdir = path_files / Path(id).stem
        fdir.mkdir()
        try:
            with ZipFile(fzip, "r") as zip_ref:
                zip_ref.extractall(fdir)
        except BadZipFile:
            warnings.warn(f"Something happened at {fname=}. Carrying on.")
            (fdir / "corrupt_submission.txt").touch()
        fzip.unlink()
        submission[fname] = str(fdir.relative_to(root_end))
    df_all_tpl["submission"] = df_all_tpl["submission"].map(submission)
    # 4) Enhance ease of use in Excel: hyperlink to folder
    df_all_tpl["submission"] = df_all_tpl["submission"].apply(
        lambda x: "" if pd.isna(x) else HYPERLINK_TPL.format(x)
    )

    # 5) Wrap up and write final table
    df_all_tpl.drop(["log"], axis=1, inplace=True)
    if drop_empty:
        df_all_tpl = (
            df_all_tpl.replace("", float("nan")).dropna(how="all", axis=0).fillna("")
        )
    name = "template-" + assignment.lower().replace(" ", "_") + ".xlsx"
    f = root_end / name
    _write_aloud(f, lambda x: _df_to_excel(df_all_tpl, x, group_icols=[2, 3]))
    return df_all_tpl
