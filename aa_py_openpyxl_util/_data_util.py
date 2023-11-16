from typing import (
    Iterable,
    Optional,
    List,
    Callable,
    Any,
    Generator,
    OrderedDict,
    TypeVar,
)

from pydicti import odicti

T = TypeVar("T")
U = TypeVar("U")


def data_to_dicts(
    *,
    data: Iterable[Iterable[T]],
    value_callback: Callable[[T], U],
    header_callback: Callable[[T], str],
    columns: Optional[List[str]] = None,
) -> Generator[OrderedDict[str, U], None, None]:
    """
    Convert 2D data (rows and columns) into a generator, assuming the first line is a header.

    Args:
        data: The data containing the rows and columns. The first dimension should represent the rows.
        columns: A list of column names to keep. All other columns are ignored. If not specified, use all columns.
        header_callback: A callback function used to process each header value.
        value_callback: A callback function used to process each cell value.

    Yields:
        One case-insensitive ordered dictionary for each row, using keys from the header.

    Examples:
        >>> d1 = (("a", "b", "c"), (1, 2, 3), (4, 5, 6), (7, 8, 9))

        Use all columns in original order.
        >>> g = data_to_dicts(data=d1, value_callback=lambda x:x, header_callback=lambda x:x)
        >>> next(g)
        odicti({'a': 1, 'b': 2, 'c': 3})
        >>> next(g)
        odicti({'a': 4, 'b': 5, 'c': 6})
        >>> next(g)
        odicti({'a': 7, 'b': 8, 'c': 9})

        Use specific columns in specified order.
        >>> g = data_to_dicts(data=d1, value_callback=lambda x:x, header_callback=lambda x:x, columns=['B', 'a'])
        >>> next(g)
        odicti({'B': 2, 'a': 1})
        >>> next(g)
        odicti({'B': 5, 'a': 4})
        >>> next(g)
        odicti({'B': 8, 'a': 7})
    """
    it = iter(data)
    header = [header_callback(c) for c in next(it)]
    for row in it:
        if all_none(row):
            # This is an empty row. Skip it.
            continue

        d: OrderedDict[str, T] = odicti(zip(header, row))
        if columns:
            # Yield row with only the specified columns and in the specified order.
            yield odicti(((k, value_callback(d[k])) for k in columns))
        else:
            # Yield row with all columns and in table order.
            yield odicti(((k, value_callback(d[k])) for k in d))


def skip_empty_rows(
    data: Iterable[OrderedDict[str, T]]
) -> Generator[OrderedDict[str, T], None, None]:
    """
    Skip empty rows.
    """
    for row in data:
        if all_none(row.values()):
            # This is an empty row. Skip it.
            continue
        yield row


def all_none(it: Iterable[Any]) -> bool:
    """
    Check whether all items in the given iterable are None.
    """
    for item in it:
        if item is not None:
            return False
    return True
