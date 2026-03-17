"""
手動動作確認スクリプト。

実行方法::

    python test_excel_table.py

出力ファイル ``test_output.xlsx`` が生成される。
Excel で開いて目視確認するか、read_sheet で読み直して assert する。
"""
from excel_table.models import (
    Table2DFloat,
    FormattedTable2D,
    TableKeyValue,
)
from excel_table.models.chart_format import (
    ChartConfig,
    LineSeriesConfig,
    ColorScale,
)
from excel_table.writer import SheetWriteSchema, write_sheet
from excel_table.reader import SheetReadSchema, read_sheet
from excel_table.models.table_format import (
    FormattedTable2DSchema,
    TableKeyValueSchema,
)

OUTPUT = "test_output.xlsx"
SHEET = "Results"

# ---------------------------------------------------------------------------
# テストデータ
# ---------------------------------------------------------------------------

# 速度(rpm) x トルク(Nm) の損失マップ (単位: W)
loss_map = Table2DFloat(
    title="Motor Loss Map",
    column_label="Speed [rpm]",
    row_label="Torque [Nm]",
    column=["1000", "2000", "3000", "4000"],
    row=["10", "20", "30", "40"],
    values=[
        [10.0, 15.0, 25.0, 40.0],
        [18.0, 28.0, 45.0, 70.0],
        [30.0, 48.0, 72.0, 110.0],
        [50.0, 78.0, 115.0, 160.0],
    ],
)

fmt_loss_map = FormattedTable2D(
    table=loss_map,
    column_location="top",
    row_location="left",
    column_color="#4472C4",
    row_color="#ED7D31",
    value_conditional_formats=[
        {
            "type": "3_color_scale",
            "min_color": "#FFFFFF",
            "mid_color": "#FFFF00",
            "max_color": "#FF0000",
        }
    ],
    row_label_direction="vertical"
)

# 動作条件テーブル
conditions = TableKeyValue(
    title="Operating Conditions",
    column=["Voltage [V]", "Temp [degC]", "Mode"],
    value=[48, 25, "continuous"],
)

# チャート設定
chart_cfg = ChartConfig(
    chart_type="line",
    width=480,
    height=288,
    x_label="Speed [rpm]",
    y_label="Loss [W]",
    x_axis="column",
    y_axis="value",
    series=[
        LineSeriesConfig(
            label="Loss",
            source_block="Motor Loss Map",
            style="both",
            color_axis="row",
            series_colorscale=ColorScale(
                min_color="#0000FF",
                mid_color="#FFFFFF",
                max_color="#FF0000",
            ),
        )
    ],
)

# ---------------------------------------------------------------------------
# Write
# ---------------------------------------------------------------------------

schema = SheetWriteSchema(rows=[
    [fmt_loss_map, chart_cfg, conditions] for _ in range(3)
])

write_sheet(OUTPUT, SHEET, schema)
print(f"Written: {OUTPUT}")

# ---------------------------------------------------------------------------
# Read back & assert
# ---------------------------------------------------------------------------

read_schema = SheetReadSchema(
    columns=[
        FormattedTable2DSchema(
            title="Motor Loss Map",
            table_type=Table2DFloat,
            column_location="top",
            row_location="left",
        ),
        TableKeyValueSchema(title="Operating Conditions"),
    ],
)

result = read_sheet(OUTPUT, SHEET, read_schema)
assert len(result) == 3, f"Expected 3 row, got {len(result)}"

read_map: Table2DFloat = result[0][0]
read_cond: TableKeyValue = result[0][1]

assert read_map.title == "Motor Loss Map"
assert read_map.column == ["1000", "2000", "3000", "4000"]
assert read_map.row == ["10", "20", "30", "40"]
assert read_map.values[0][0] == 10.0
assert read_map.values[3][3] == 160.0

assert read_cond.title == "Operating Conditions"
assert read_cond.column == ["Voltage [V]", "Temp [degC]", "Mode"]
assert read_cond.value[0] == 48
assert read_cond.value[2] == "continuous"

print("All assertions passed.")