# Grade Conversion App

A desktop app built using Python and Tkinter to calculate and visualize final student grades from a CSV file.

## Features
- Read CSV or Excel files
- Calculate final grades with weighted formula
- Export to Excel
- Plot grade distribution
- Simple and interactive UI

## Requirements
- Python 3.8+
- See `requirements.txt` for Python packages

## How to Run

1. Clone the repository:
```
git clone https://github.com/azariadr/Grade-Conversion.git
cd grade-conversion-app
```
2. Install dependencies:
```
pip install -r requirements.txt
```
3. Run the app:
```
python "Grade Conversion.py"
```
### Sample Input Format
Your CSV should contain these columns:
Homework 1, Homework 2, Homework 3, Quiz 1, Quiz 2, Exam 1, and Exam 2

## Output
- Excel file with calculated letter grades
- Bar chart showing grade distribution
