# word-template-filler
A simple Python script that fills Word .docx templates with data from Excel, replacing placeholders and removing unused lines automatically.


## How to prepare your files

### 1. Excel file (`source.xlsx`)

- Open your Excel file.
- On **Sheet1**, enter the **customer name** in cell **B1**.
- Starting from row 2:
  - Put each **product name** in column **A** (one product per row).
  - Put the corresponding **product amount** in column **B**.
- Example:

| A (Product) | B (Amount) |
|-------------|------------|
| Product1    | 90000      |
| Product2    | 45000      |
| Product3    | 11000      |

- Leave empty or zero amounts for products you want to exclude.

---

### 2. Word template file (`template.docx`)

- Write your letter or document as usual.
- Use placeholders to mark where the data should go.
- Placeholders must match this format:

  - `${kunde}` — will be replaced by the customer name from cell B1.
  - For products, use two placeholders per product:

    - `${ProductName_name}` — replaced with the product name.
    - `${ProductName_sum}` — replaced with the product amount.




- Any line with a product whose amount is zero or missing will be automatically removed from the final document.

---

## How to use

1. Make sure your Excel and Word files are saved in the same folder as the program.
2. Run the program (or the `.exe`) — it will create a new Word file named `filled.docx` with your data filled in.
3. Open `filled.docx` and review your personalized document.

---

## Tips

- Double-check that product names in Excel exactly match the placeholders in Word (case-sensitive).
- If you add new products, update both Excel and Word placeholders accordingly.
- Leave amounts empty or zero in Excel to exclude those product lines from the final document.

---




## To modify the script:

# Install python
https://www.python.org/downloads/windows/

# Build Instructions Powershell
```
py -m venv venv
.\venv\Scripts\activate
pip install --upgrade pip
pip install pyinstaller openpyxl python-docx
pyinstaller --onefile main.py
```
