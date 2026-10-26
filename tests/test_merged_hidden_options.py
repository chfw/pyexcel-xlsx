from io import BytesIO

import pytest
from openpyxl import Workbook
from pyexcel_xlsx import get_data


@pytest.mark.parametrize("source", ["stream", "file"])
@pytest.mark.parametrize("skip_hidden", [False, True])
def test_merged_cells_respect_hidden_filter_option(
    tmp_path, source, skip_hidden
):
    workbook = Workbook()
    sheet = workbook.active
    sheet["A1"] = "merged"
    sheet.merge_cells("A1:B2")
    sheet["C1"] = "visible"
    sheet["C2"] = "hidden row"
    sheet.column_dimensions["B"].hidden = True
    sheet.row_dimensions[2].hidden = True
    if source == "stream":
        target = BytesIO()
        workbook.save(target)
        target.seek(0)
    else:
        target = str(tmp_path / "merged-hidden.xlsx")
        workbook.save(target)

    data = get_data(
        target,
        detect_merged_cells=True,
        skip_hidden_row_and_column=skip_hidden,
        library="pyexcel-xlsx",
    )

    if skip_hidden:
        assert data["Sheet"] == [["merged", "visible"]]
    else:
        assert data["Sheet"] == [
            ["merged", "merged", "visible"],
            ["merged", "merged", "hidden row"],
        ]
