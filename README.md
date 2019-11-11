# Python Excel Concatenation Version 2.0

RELEASE DATE: 11/11/2019

PyXL Concatenation is a lightweight python program to combine data from any number of templated excel sheets into one spreadsheet (.xslx format).

The program:

* Opens all files listed in the .config

* Selects the specified columns from each file

* Places the entries into a new excel sheet containing only the specified columns

To use, simply add the desired sheets to the Project\Sheets directory and list them in the project.config, add the column names to the project.config, and run the Concatenate_XLSX.py file.

## Dependencies

PyXL requires Python Modules: xlrd, pandas, & numpy to run properly

---

## Settings

### Overview

In order to work properly, PyXL requires a project.config file be placed in the Configurations folder of the project. An example config is provided, along with sample data, for your convenence.

All settings are loaded into the project with the creation of the Settings data object class. Attributes are accessed directly from the instantiated object.

### Formatting

Each entry in the .config file takes the general format of "**Setting: Associated Data**" where associated data can be a single string ("content"), or multiple comma-separated strings

### Default Options

The sample program runs with the following options:

* Export Path: The path where the resulting excel spreadsheet is saved to, relative to the project
* Template Path: The path to the template spreadsheet from which the result will be built
* Sheets Path: The directory where each excel spreadsheet can be found
* All Headers: List of all column names from the top of the spreadsheet
* Headers to Extract: The columns you wish to be included in the final result (all others will be excluded)
* Included Files: List of spreadsheets in the Sheets Path directory that you wish to be included in the result

---

## Project Structure

* Configurations: Stores the project.config file
* Imports: Stores all custom files accessed by the project
    * Sheets: Default Directory for the Spreadsheets to Concatenate
    * Template: Default Directory for the Template Spreadsheet
* Output: Default Directory to Save to
* Concatenate_XLSX.py: Main project file

---

## Upcoming Changes

* Default arguments for the headers (include all columns by default, removing dependency on column names being listed)
* Individual spreadsheet settings (select only certain rows)

---