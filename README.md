# PDF Generator

This project is a Python-based PDF generator that can create PDF files from different sources like CSV files, DOCX files, and manual text input. It also supports simple page formatting such as plain pages and ruled pages with page numbers.

## What this project does

The program gives four options:

* Convert a DOCX file into PDF
* Generate PDF pages from CSV data
* Combine CSV topics with DOCX content
* Create PDF from manually entered text

## Features

* Reads topic and page count from CSV
* Reads text content from Word documents
* Converts DOCX to PDF using LibreOffice
* Adds page numbers automatically
* Supports plain and ruled page designs

## Technologies used

* Python
* ReportLab
* Pandas
* python-docx
* LibreOffice

## Example CSV format

```csv id="1d2t8m"
Topic,Pages
Introduction,2
Python Basics,3
```

## How to run

Install required libraries:

```bash id="zqk8ot"
pip install reportlab pandas python-docx
```

Run the program:

```bash id="g31x3o"
python3 main.py
```

## Output

The generated file is saved as:

```text id="d73yka"
output.pdf
```

## Notes

LibreOffice must be installed because DOCX to PDF conversion uses LibreOffice in headless mode.

## Project purpose

This project was built to understand file handling, PDF generation, CSV reading, DOCX reading, and automation in Python.
