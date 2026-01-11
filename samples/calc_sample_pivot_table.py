# PivotTable (DataPilot) sample

from excellikeuno.connection import connect_calc
from excellikeuno.typing.structs import Rectangle

# Connect to Calc and get the first sheet
(desktop, doc, sheet) = connect_calc()

# Prepare sample data
# Columns: Category, Label, Value
values = [
    ["Category", "Label", "Value"],
    ["Fruit", "Apple", 10],
    ["Fruit", "Banana", 8],
    ["Fruit", "Apple", 12],
    ["Vegetable", "Carrot", 5],
    ["Vegetable", "Onion", 7],
    ["Vegetable", "Carrot", 9],
]

# Write data to the sheet
source = sheet.range("A1:C7")
source.value = values

# Optional: resize columns a bit for readability (best-effort)
try:
    sheet.columns.getByIndex(0).Width = 3000
    sheet.columns.getByIndex(1).Width = 3000
    sheet.columns.getByIndex(2).Width = 2500
except Exception:
    pass

# Create a pivot table starting at E2
pivot = sheet.pivot_tables.add(
    name="Sample_Pivot",
    source_range=source,
    target_cell=(4, 1),  # E2 (0-based col/row)
    row_fields=["Category", "Label"],
    column_fields=None,
    data_fields=["Value"],
    filter_fields=None,
)

# Refresh to ensure data is populated
try:
    pivot.refresh()
except Exception:
    pass

print("Pivot table created:", pivot.name)

# Add a chart for the pivot result (optional, best-effort)
try:
    output_range = getattr(pivot, "output_range", None)
    if output_range is not None:
        chart = sheet.charts.add_bar_diagram(
            name="PivotChart",
            data_range=output_range,
            rectangle=Rectangle(12000, 2000, 10000, 8000),
            column_headers=True,
            row_headers=True,
        )
        chart.has_legend = True
        chart.title = "Pivot Summary"
except Exception:
    pass
