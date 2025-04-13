# Nielsen Formatter

## Important Notice: Dummy Data Disclaimer

Please note that all Nielsen data presented here is dummy data.
This means it is fictional and created solely for illustrative or testing purposes.
It does not represent real market data, consumer behavior, or actual Nielsen insights.

Use this data only for demonstration or training purposes — not for decision-making or reporting.

##
Project for Maspex GTM Employees that Automatically Formats Nielsen Data Into Excell Sheet with Calculations.
This project processes Nielsen data from Excel files, formats it, and generates a structured Excel output with calculations, filters, and visual enhancements.

![Formatter](https://github.com/user-attachments/assets/4eb880f4-3b72-4029-a7d6-ff4c535fa473)

## Features

- **Automatic Library Installation**: Ensures all required Python libraries are installed.
- **Data Import**: Reads data from an Excel file located in the `input` directory.
- **Date Formatting**: Extracts and reformats date columns for consistency.
- **Filtering**:
  - Filters out unnecessary years and facts.
  - Removes specific UPC patterns.
- **Including XLSXWRITER logic**
- **Searching Item By UPC**: UPC is used to identify specific item
- **Calculating Needed Informations**: Used Calculations:
    - "Sales Units",
    - "Sales Value",
    - "Average Price per Unit",
    - "Share in Units Sold",
    - "Value Share",
    - "Numeric Distribution TOTAL",
    - "Weighted Value Distribution",
    - "UNIT INDEX",
    - "VALUE INDEX"
- **Excel Output**:
  - Writes processed data to an Excel file in the `output` directory.
  - Adds multiple sheets (`DATA`, `LISTS`, `LINE`) with structured data and calculations.
  - Includes dynamic formulas, conditional formatting, and dropdown lists for interactivity.


## Requirements

- Python 3.7+
- Libraries:
  - `pandas`
  - `os`
  - `re`
  - `glob`
  - `xlsxwriter`
  - `datetime`

## Installation

1. Clone the repository.
2. Installation proces will begin automatically when you run first code line in `nielsen_formatter.ipynb`

## Usage

1. Place the input Excel file in the input directory.
2. !!! Specify the `SHEET_NUMBER` to format correct sheet in Excel workbook. ATTENTION: COUNT FROM 0 !!!
3. Open and run the nielsen_formatter.ipynb notebook in Jupyter Notebook, JupyterLab or Visual Studio Code.
4. The processed Excel file will be saved in the output directory as output.xlsx.
5. Open the `output.xlsx` file. Now you can paste UPC numbers below UPC row

## Contact

For any questions, please contact:

* **Author**: Milan Wcisło
* **Email**: milanwcislo@gmail.com
* **GitHub**: [Milan-Wcislo](https://github.com/Milan-Wcislo)

Thank you for checking out this project! Your support and feedback are greatly appreciated.

