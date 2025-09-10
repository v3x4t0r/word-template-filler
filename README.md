# word-template-filler
A simple Python script that fills Word .docx templates with data from Excel, replacing placeholders and removing unused lines automatically.

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
