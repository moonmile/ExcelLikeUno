from pathlib import Path
from excellikeuno import new_calc_document, open_calc_document

OUTPUT_PATH = Path(__file__).with_name("output_save_load.ods")


# Create a new hidden document, write a couple of values, and save it.
_, doc, sheet = new_calc_document(hidden=True)
sheet.cell(0, 0).value = 123.45
sheet.cell(1, 0).text = "saved via save_as"
doc.save_as(str(OUTPUT_PATH))
print(f"Saved to {OUTPUT_PATH}")

# Re-open the saved document to verify round-trip.
_, reopened_doc, reopened_sheet = open_calc_document(str(OUTPUT_PATH), hidden=True)
value = reopened_sheet.cell(0, 0).value
text = reopened_sheet.cell(1, 0).text
print(f"Reopened {OUTPUT_PATH.name}: A1={value}, B1='{text}'")

# Make a copy using storeToURL for demonstration.
copy_path = OUTPUT_PATH.with_stem("output_save_copy")
reopened_doc.save_copy(str(copy_path))
print(f"Copied to {copy_path}")
