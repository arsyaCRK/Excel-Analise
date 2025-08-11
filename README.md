# Excel Filter Utility

A simple desktop application for filtering data in Excel files.

## Description

This application provides a graphical user interface (GUI) to help users filter data from an `.xlsx` file based on specific column values and dates. The filtered data is then saved to a new `.xlsx` file. The application's interface is in Russian.

## Features

*   Select an Excel file (`.xlsx`) using a file dialog.
*   Filter data by a specific column and value.
*   Optionally, filter by a date column and a specific date string.
*   Save the filtered data to a new Excel file.
*   User-friendly interface.

## Requirements

To run this application from the source code, you will need Python 3 and the following libraries:

*   PyQt6
*   openpyxl
*   openpyxl-dictreader

You can install these dependencies using pip:

```bash
pip install PyQt6 openpyxl openpyxl-dictreader
```

## How to Run from Source

1.  Clone this repository or download the source code.
2.  Install the required dependencies (see Requirements section).
3.  Run the `main.py` script:

```bash
python main.py
```

## How to Use the Application

1.  **Выберите файл (Select file):** Click the "..." button to open a file dialog and select the `.xlsx` file you want to filter.
2.  **Введите название столбца для фильра (Enter column name for filter):** In this text field, enter the exact name of the column you want to use for filtering.
3.  **Введите значение фильра (Enter filter value):** Enter the value you want to filter for in the specified column.
4.  **Название столбца дата (Date column name):** (Optional) If you want to filter by date, enter the name of the date column.
5.  **Введите дату (Enter date):** (Optional) If you are filtering by date, enter the date string to search for.
6.  **Выполнить (Execute):** Click this button to start the filtering process.
7.  You will be prompted to choose a location and name for the new, filtered `.xlsx` file.

## Building from Source

This project uses PyInstaller to create a standalone executable.

1.  Install PyInstaller:
    ```bash
    pip install pyinstaller
    ```
2.  Navigate to the project's root directory in your terminal.
3.  Run the following command to build the executable:
    ```bash
    pyinstaller main.spec
    ```
4.  The executable will be located in the `dist` directory.
