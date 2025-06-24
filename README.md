# SapScript – SAP Automation Script

**SapScript** is a Python-based automation tool designed to streamline repetitive tasks in SAP. It reads an Excel file with input data and executes a defined SAP operation automatically, handling dozens of entries in seconds.

## 🔄 Overview

In SAP, certain operations require manual repetition for tens of entries (e.g. 40+ rows). This script automates the process end-to-end:

1. Reads rows from a given Excel file  
2. Opens the relevant SAP transaction  
3. Inputs each row automatically  
4. Executes the operation and repeats for all entries  
5. Logs success and errors for each row

## 🚀 Benefits

- Saves significant time by automating repetitive SAP tasks  
- Reduces human errors in data entry  
- Logs each execution for traceability and audit purposes

## ⚙️ How It Works

- Uses **Python** with an SAP GUI automation library (e.g. `pywinauto`, `win32com`)  
- Reads input data via **Pandas** from an `.xlsx` file  
- Structured logging for results and errors

*Note: Before running, ensure you have access to SAP GUI and the correct Excel format.*

## 🛠️ Tech Stack

- Python 3  
- Pandas  
- SAP GUI automation (`pywinauto`, `win32com`, or similar)  
- Excel `.xlsx` input

## 🚧 Status

- Fully functional for the target SAP transaction  
- Tested with batch inputs (>40 rows)  
- Extendable to other SAP tasks or transactions
