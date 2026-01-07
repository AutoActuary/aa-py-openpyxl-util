from typing import Iterable, TYPE_CHECKING

if TYPE_CHECKING:
    from openpyxl.descriptors.excel import CellRange
    from openpyxl.worksheet.worksheet import Worksheet


def set_data_validation_input_message(
    *,
    worksheet: "Worksheet",
    ranges: Iterable["str | CellRange"],
    title: str,
    input_message: str,
) -> None:
    """
    Set an "input message" for the given ranges.

    This is like using the "Input Message" tab of the "Data Validation" dialog in Excel.

    Args:
        worksheet: The worksheet to set the data validation input message on.
        ranges: The ranges on which to set the data validation input message.
        title: The title. It will be shortened if longer than 32 characters.
        input_message: The message. It will be shortened if longer than 255 characters.
    """
    from openpyxl.worksheet.datavalidation import DataValidation

    validation = DataValidation(
        showInputMessage=True,
        promptTitle=title[:32],
        prompt=input_message[:255],
    )

    for rng in ranges:
        validation.sqref.add(rng)

    worksheet.add_data_validation(validation)
