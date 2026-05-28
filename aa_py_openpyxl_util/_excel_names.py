from __future__ import annotations

import re
from typing import Callable, Iterable, Literal, TypeAlias

ExcelNameKind: TypeAlias = Literal["name", "table", "defined_name"]

MAX_EXCEL_NAME_LENGTH = 255
MAX_EXCEL_SHEET_TITLE_LENGTH = 31
MAX_EXCEL_ROWS = 1_048_576
MAX_EXCEL_COLUMNS = 16_384

_A1_REFERENCE_RE = re.compile(r"^([A-Za-z]{1,3})([0-9]+)$")
_R1C1_REFERENCE_RE = re.compile(r"^R([0-9]*)C([0-9]*)$", re.IGNORECASE)
_UNDERSCORES_RE = re.compile(r"_+")
_INVALID_SHEET_TITLE_CHARS = set("[]:*?/\\")


class ExcelNameError(ValueError):
    """
    Base class for Excel-name validation errors raised by this package.
    """


class InvalidExcelNameError(ExcelNameError):
    """
    Raised when a table name or defined name is not accepted by Excel.
    """

    def __init__(
        self,
        name: object,
        reason: str,
        *,
        kind: ExcelNameKind = "name",
    ) -> None:
        self.name = name
        self.reason = reason
        self.kind = kind
        super().__init__(f"Invalid Excel {_name_kind_label(kind)} {name!r}: {reason}")


class DuplicateExcelNameError(ExcelNameError):
    """
    Raised when an Excel name conflicts with another name in the same scope.
    """

    def __init__(
        self,
        name: str,
        existing_name: str,
        *,
        kind: ExcelNameKind = "name",
        scope_label: str = "scope",
    ) -> None:
        self.name = name
        self.existing_name = existing_name
        self.kind = kind
        self.scope_label = scope_label
        super().__init__(
            f"Duplicate Excel {_name_kind_label(kind)} {name!r}: Excel names are "
            f"case-insensitive in {scope_label} and this conflicts with "
            f"{existing_name!r}."
        )


class InvalidExcelSheetTitleError(ExcelNameError):
    """
    Raised when a worksheet title is not accepted by Excel.
    """

    def __init__(self, title: object, reason: str) -> None:
        self.title = title
        self.reason = reason
        super().__init__(f"Invalid Excel sheet title {title!r}: {reason}")


def validate_excel_name(
    name: object,
    *,
    kind: ExcelNameKind = "name",
) -> str:
    """
    Validate an Excel table name or defined name.

    Returns:
        The original name, typed as `str`, for convenient inline use.

    Raises:
        InvalidExcelNameError: If Excel would reject or misinterpret the name.
    """
    if not isinstance(name, str):
        raise InvalidExcelNameError(
            name,
            f"expected a string, got {type(name).__name__}.",
            kind=kind,
        )

    if not name:
        raise InvalidExcelNameError(
            name,
            "names cannot be blank.",
            kind=kind,
        )

    if len(name) > MAX_EXCEL_NAME_LENGTH:
        raise InvalidExcelNameError(
            name,
            f"names can be at most {MAX_EXCEL_NAME_LENGTH} characters long; "
            f"got {len(name)}.",
            kind=kind,
        )

    first_char = name[0]
    if not _is_valid_first_name_char(first_char):
        raise InvalidExcelNameError(
            name,
            "the first character must be a letter, underscore, or backslash.",
            kind=kind,
        )

    for i, char in enumerate(name[1:], start=2):
        if not _is_valid_name_body_char(char):
            raise InvalidExcelNameError(
                name,
                f"character {char!r} at position {i} is not allowed; after the "
                "first character, use only letters, numbers, underscores, or "
                "periods.",
                kind=kind,
            )

    if name.casefold() in {"r", "c"}:
        raise InvalidExcelNameError(
            name,
            "`R` and `C` are reserved by Excel's R1C1 reference style.",
            kind=kind,
        )

    if is_a1_reference_like(name):
        raise InvalidExcelNameError(
            name,
            "names cannot look like an A1-style cell reference inside Excel's "
            "grid (A1:XFD1048576). Prefix the name with text such as `Name_`.",
            kind=kind,
        )

    if is_r1c1_reference_like(name):
        raise InvalidExcelNameError(
            name,
            "names cannot look like an R1C1-style cell reference.",
            kind=kind,
        )

    if name.casefold().startswith("_xl"):
        raise InvalidExcelNameError(
            name,
            "names starting with `_xl` are reserved by Excel and can make "
            "workbooks unreadable.",
            kind=kind,
        )

    return name


def validate_excel_table_name(name: object) -> str:
    """
    Validate an Excel table/ListObject name.
    """
    return validate_excel_name(name, kind="table")


def validate_excel_defined_name(name: object) -> str:
    """
    Validate an Excel defined-name or named-range name.
    """
    return validate_excel_name(name, kind="defined_name")


def is_valid_excel_name(
    name: object,
    *,
    kind: ExcelNameKind = "name",
) -> bool:
    """
    Return True if `name` is accepted by `validate_excel_name`.
    """
    try:
        validate_excel_name(name, kind=kind)
    except InvalidExcelNameError:
        return False
    return True


def is_valid_excel_table_name(name: object) -> bool:
    """
    Return True if `name` is a valid Excel table/ListObject name.
    """
    return is_valid_excel_name(name, kind="table")


def is_valid_excel_defined_name(name: object) -> bool:
    """
    Return True if `name` is a valid Excel defined-name or named-range name.
    """
    return is_valid_excel_name(name, kind="defined_name")


def make_safe_excel_name(
    name: object,
    *,
    kind: ExcelNameKind = "name",
    fallback: str | None = None,
    existing_names: Iterable[str] = (),
) -> str:
    """
    Convert arbitrary text to a valid Excel table/defined name.

    This function does not mutate workbooks. It returns a deterministic, valid
    name and appends a numeric suffix if `existing_names` already contains the
    candidate case-insensitively.
    """
    fallback = _safe_excel_name_fallback(kind=kind, fallback=fallback)
    candidate = _basic_safe_excel_name(name, fallback=fallback)

    if not is_valid_excel_name(candidate, kind=kind):
        candidate = _prepend_excel_name_fallback(candidate, fallback=fallback)

    if not is_valid_excel_name(candidate, kind=kind):
        candidate = fallback

    return _make_unique(
        candidate,
        existing_names=existing_names,
        max_length=MAX_EXCEL_NAME_LENGTH,
        is_valid=lambda value: is_valid_excel_name(value, kind=kind),
    )


def make_safe_excel_table_name(
    name: object,
    *,
    fallback: str = "Table",
    existing_names: Iterable[str] = (),
) -> str:
    """
    Convert arbitrary text to a valid Excel table/ListObject name.
    """
    return make_safe_excel_name(
        name,
        kind="table",
        fallback=fallback,
        existing_names=existing_names,
    )


def make_safe_excel_defined_name(
    name: object,
    *,
    fallback: str = "Name",
    existing_names: Iterable[str] = (),
) -> str:
    """
    Convert arbitrary text to a valid Excel defined-name or named-range name.
    """
    return make_safe_excel_name(
        name,
        kind="defined_name",
        fallback=fallback,
        existing_names=existing_names,
    )


def validate_unique_excel_names(
    names: Iterable[object],
    *,
    kind: ExcelNameKind = "name",
    existing_names: Iterable[str] = (),
    scope_label: str = "scope",
) -> None:
    """
    Validate names and ensure they are unique within an Excel name scope.
    """
    seen = {
        existing_name.casefold(): existing_name
        for existing_name in existing_names
        if isinstance(existing_name, str)
    }

    for name in names:
        valid_name = validate_excel_name(name, kind=kind)
        key = valid_name.casefold()
        if key in seen:
            raise DuplicateExcelNameError(
                valid_name,
                seen[key],
                kind=kind,
                scope_label=scope_label,
            )
        seen[key] = valid_name


def validate_excel_sheet_title(title: object) -> str:
    """
    Validate a worksheet title before creating an Excel sheet.
    """
    if not isinstance(title, str):
        raise InvalidExcelSheetTitleError(
            title,
            f"expected a string, got {type(title).__name__}.",
        )

    if not title:
        raise InvalidExcelSheetTitleError(title, "sheet titles cannot be blank.")

    if len(title) > MAX_EXCEL_SHEET_TITLE_LENGTH:
        raise InvalidExcelSheetTitleError(
            title,
            f"sheet titles can be at most {MAX_EXCEL_SHEET_TITLE_LENGTH} "
            f"characters long; got {len(title)}.",
        )

    invalid_chars = sorted(set(title).intersection(_INVALID_SHEET_TITLE_CHARS))
    if invalid_chars:
        raise InvalidExcelSheetTitleError(
            title,
            "sheet titles cannot contain these characters: "
            + ", ".join(repr(char) for char in invalid_chars)
            + ".",
        )

    if title.startswith("'") or title.endswith("'"):
        raise InvalidExcelSheetTitleError(
            title,
            "sheet titles cannot start or end with an apostrophe; Excel refuses "
            "to open files with those sheet titles.",
        )

    return title


def is_valid_excel_sheet_title(title: object) -> bool:
    """
    Return True if `title` is accepted by `validate_excel_sheet_title`.
    """
    try:
        validate_excel_sheet_title(title)
    except InvalidExcelSheetTitleError:
        return False
    return True


def make_safe_excel_sheet_title(
    title: object,
    *,
    fallback: str = "Sheet",
    existing_titles: Iterable[str] = (),
) -> str:
    """
    Convert arbitrary text to a valid Excel worksheet title.
    """
    fallback = _safe_sheet_title_fallback(fallback)
    candidate = _basic_safe_sheet_title(title, fallback=fallback)

    if not is_valid_excel_sheet_title(candidate):
        candidate = fallback

    return _make_unique(
        candidate,
        existing_names=existing_titles,
        max_length=MAX_EXCEL_SHEET_TITLE_LENGTH,
        is_valid=is_valid_excel_sheet_title,
    )


def is_a1_reference_like(name: str) -> bool:
    """
    Return True if `name` can be parsed as a cell reference in Excel's A1 grid.
    """
    match = _A1_REFERENCE_RE.fullmatch(name)
    if not match:
        return False

    column_letters, row_digits = match.groups()
    column_index = _column_index_from_letters(column_letters)
    row_index = int(row_digits)

    return 1 <= column_index <= MAX_EXCEL_COLUMNS and 1 <= row_index <= MAX_EXCEL_ROWS


def is_r1c1_reference_like(name: str) -> bool:
    """
    Return True if `name` looks like an R1C1-style cell reference.

    Excel treats some edge cases around omitted/current row/column references
    oddly. These rules match the cases that cause Excel to reject names in
    practice: RC, RC1, R1C, R1C1, and any R<valid-row>C... form.
    """
    match = _R1C1_REFERENCE_RE.fullmatch(name)
    if not match:
        return False

    row_digits, column_digits = match.groups()
    if row_digits == "":
        return column_digits == "" or int(column_digits) > 0

    row_index = int(row_digits)
    return 1 <= row_index <= MAX_EXCEL_ROWS


def iter_workbook_scope_excel_names(book: object) -> Iterable[str]:
    """
    Iterate names that occupy workbook-global Excel name scope.

    This includes workbook-scoped defined names and all ListObject/table names.
    """
    defined_names = getattr(book, "defined_names", None)
    if defined_names is not None:
        yield from defined_names.keys()

    for sheet in getattr(book, "worksheets", ()):
        tables = getattr(sheet, "tables", None)
        if tables is not None:
            yield from tables.keys()


def _name_kind_label(kind: ExcelNameKind) -> str:
    if kind == "table":
        return "table name"
    if kind == "defined_name":
        return "defined name"
    return "name"


def _is_valid_first_name_char(char: str) -> bool:
    return char == "_" or char == "\\" or char.isalpha()


def _is_valid_name_body_char(char: str) -> bool:
    return char == "_" or char == "." or char.isalpha() or char.isdecimal()


def _column_index_from_letters(column_letters: str) -> int:
    result = 0
    for char in column_letters.upper():
        result = result * 26 + (ord(char) - ord("A") + 1)
    return result


def _safe_excel_name_fallback(
    *,
    kind: ExcelNameKind,
    fallback: str | None,
) -> str:
    default = "Table" if kind == "table" else "Name"
    candidate = _basic_safe_excel_name(
        default if fallback is None else fallback, default
    )

    if is_valid_excel_name(candidate, kind=kind):
        return candidate

    return default


def _basic_safe_excel_name(name: object, fallback: str) -> str:
    text = "" if name is None else str(name).strip()
    chars: list[str] = []

    for i, char in enumerate(text):
        if i == 0 and _is_valid_first_name_char(char):
            chars.append(char)
        elif _is_valid_name_body_char(char):
            chars.append(char)
        else:
            chars.append("_")

    candidate = _UNDERSCORES_RE.sub("_", "".join(chars))
    if not candidate or set(candidate) == {"_"}:
        candidate = fallback
    elif not _is_valid_first_name_char(candidate[0]):
        candidate = f"{fallback}_{candidate}"

    return candidate[:MAX_EXCEL_NAME_LENGTH]


def _prepend_excel_name_fallback(candidate: str, *, fallback: str) -> str:
    separator = "_"
    max_tail_length = MAX_EXCEL_NAME_LENGTH - len(fallback) - len(separator)
    if max_tail_length <= 0:
        return fallback[:MAX_EXCEL_NAME_LENGTH]

    tail = candidate[:max_tail_length]
    if not tail:
        return fallback

    return f"{fallback}{separator}{tail}"


def _safe_sheet_title_fallback(fallback: str) -> str:
    candidate = _basic_safe_sheet_title(fallback, fallback="Sheet")
    if is_valid_excel_sheet_title(candidate):
        return candidate
    return "Sheet"


def _basic_safe_sheet_title(title: object, *, fallback: str) -> str:
    text = "" if title is None else str(title).strip()
    chars = ["_" if char in _INVALID_SHEET_TITLE_CHARS else char for char in text]
    candidate = "".join(chars).strip("'")
    if not candidate:
        candidate = fallback

    candidate = candidate[:MAX_EXCEL_SHEET_TITLE_LENGTH].strip("'")
    if not candidate:
        candidate = fallback

    return candidate


def _make_unique(
    candidate: str,
    *,
    existing_names: Iterable[str],
    max_length: int,
    is_valid: Callable[[str], bool],
) -> str:
    seen = {
        existing_name.casefold()
        for existing_name in existing_names
        if isinstance(existing_name, str)
    }

    candidate = candidate[:max_length]
    if is_valid(candidate) and candidate.casefold() not in seen:
        return candidate

    for i in range(2, 10_000):
        suffix = f"_{i}"
        base = candidate[: max_length - len(suffix)]
        next_candidate = f"{base}{suffix}"
        if is_valid(next_candidate) and next_candidate.casefold() not in seen:
            return next_candidate

    raise ExcelNameError(
        f"Could not create a unique Excel name based on {candidate!r}."
    )
