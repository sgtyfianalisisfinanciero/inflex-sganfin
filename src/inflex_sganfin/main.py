from pathlib import Path

import docx
import yaml
from data.series import Series
from representations.table import Table, TableBlock, TableColumn

if __name__ == "__main__":
    with open(Path("src") / "inflex_sganfin" / "test_table.yaml", "r") as f:
        full_dict = yaml.safe_load(f)
    blocks_dict_list = full_dict["blocks"]
    blocks: list[TableBlock] = []
    for block_dict in blocks_dict_list:
        heading = block_dict["heading"]
        rows_dict_list = block_dict["rows"]
        rows_list: list[Series] = []
        columns_dict_list = block_dict["columns"]
        columns_list: list[TableColumn] = []
        for row_dict in rows_dict_list:
            rows_list.append(Series(**row_dict))
        for column_dict in columns_dict_list:
            columns_list.append(TableColumn(**column_dict))
        blocks.append(TableBlock(heading=heading, columns=columns_list, rows=rows_list))

    table = Table(
        blocks=blocks,
        block_separation=full_dict["block_separation"],
        stack_axis=full_dict["stack_axis"],
    )
    table.render_docx(docx.Document())
