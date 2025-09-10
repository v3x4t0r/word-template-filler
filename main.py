from openpyxl import load_workbook
from docx import Document

# === CONFIGURATION ===

# Path to the Excel file where customer and product data is stored
EXCEL_FILE = "source.xlsx"

# Path to the Word template with placeholders like ${kunde}, ${Produkt1_name}, ${Produkt1_sum}, etc.
TEMPLATE_FILE = "template.docx"

# Path to the output Word file (result will be saved here)
OUTPUT_FILE = "filled.docx"

# Sheet name in Excel where data is stored
SHEET_NAME = "Sheet1"

# The cell that contains the customer name (you can change this)
CUSTOMER_CELL = "B1"

# Start reading products from this row downward (row 2 = first product)
PRODUCTS_START_ROW = 2

# === FUNCTION: Read all replacements from Excel ===

def get_replacements_from_excel(excel_path, sheet_name, customer_cell, start_row):
    """
    Reads customer name and a list of products from Excel.
    For each product row, adds two placeholders:
      - ${ProductName_name} → product name
      - ${ProductName_sum}  → product value

    Expected Excel format:
    ------------------------------------
    B1         → Customer name (e.g., "Ola Nordmann")
    A2:B2      → Product name, product sum
    A3:B3      → ...
    until column A is empty
    ------------------------------------

    Only products with a non-zero value are included.
    """
    wb = load_workbook(excel_path, data_only=True)
    ws = wb[sheet_name]

    # Start replacements dict with customer name
    replacements = {
        "${kunde}": ws[customer_cell].value
    }

    row = start_row
    while True:
        name = ws[f"A{row}"].value
        value = ws[f"B{row}"].value
        if name is None:
            break
        if value and value != 0:
            replacements[f"${{{name}_name}}"] = name
            replacements[f"${{{name}_sum}}"] = value
        row += 1

    return replacements


# === FUNCTION: Replace or remove placeholders in Word ===

def fill_template(template_path, output_path, replacements):
    """
    Opens the Word template, replaces placeholders with data from Excel,
    and removes any paragraph that contains an unresolved placeholder.
    """
    doc = Document(template_path)
    to_remove = []

    for para in doc.paragraphs:
        full_text = ''.join(run.text for run in para.runs)
        new_text = full_text

        # Replace all placeholders
        for key, value in replacements.items():
            new_text = new_text.replace(key, str(value))

        # If nothing was replaced and text contains ${...}, mark for deletion
        if new_text == full_text and "${" in full_text:
            to_remove.append(para)
        else:
            # Clear all runs and insert new text as a single run
            for run in para.runs:
                run.text = ''
            if para.runs:
                para.runs[0].text = new_text
            else:
                para.add_run(new_text)

    # Remove paragraphs with unresolved placeholders
    for para in to_remove:
        p = para._element
        p.getparent().remove(p)
        para._p = para._element = None

    # Save the new document
    doc.save(output_path)
    print(f"✔ Document created: {output_path}")


# === RUN THE SCRIPT ===

if __name__ == "__main__":
    replacements = get_replacements_from_excel(
        EXCEL_FILE,
        SHEET_NAME,
        CUSTOMER_CELL,
        PRODUCTS_START_ROW
    )

    fill_template(TEMPLATE_FILE, OUTPUT_FILE, replacements)
