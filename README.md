# Tkinter Data Entry Form

A desktop application built with Python and Tkinter that collects student registration data and saves it to an Excel file.

## Features

- Enter student personal information (first name, last name, title, age, nationality)
- Track course and semester counts
- Registration status checkbox
- Terms & conditions acceptance
- Saves all data to an `.xlsx` file on the desktop
- Warns user if required fields are missing

## Requirements

- Python 3.x
- openpyxl

Install dependencies:

```bash
pip install openpyxl
```

## Usage

1. Clone the repository:

```bash
git clone https://github.com/salihdincel5679-jpg/Tkinter-Data-Entry-and-Write-to-Excel-File.git
```

2. Run the application:

```bash
python main.py
```

3. Fill in the form fields and click **Enter Data** to save to `data.xlsx`.

## Project Structure

```
├── main.py       # Main application file
└── data.xlsx     # Generated automatically on first submit
```

## Notes

- The `data.xlsx` file is created automatically on first run at the path defined in `main.py`.
- If you want to change the save location, update the `filepath` variable inside the `enter_data()` function.
- First name and last name are required fields. All other fields must also be filled before submitting.
- The Terms & Conditions checkbox must be accepted before data can be saved.
