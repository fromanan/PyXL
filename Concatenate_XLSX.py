# Dependencies (modules): xlrd, pandas, numpy
from pandas import read_excel
from pandas import Series
from pandas import ExcelWriter
import time

# TODO: FIX https://stackoverflow.com/questions/47975866/pandas-read-excel-parameter-sheet-name-not-working?noredirect=1&lq=1

class ConfigReadError(Exception): pass


class Settings:
    def __init__(self, config_path):
        self.num_entries = 0

        file = open(config_path, "r")

        for line in file:
            tokens = line.split(":")
            setting = tokens[0]
            data = [item[1:-1] for item in tokens[1].strip(" \n").split(", ")]

            if setting == "Export Path":
                self.export_path = data[0]
            elif setting == "Template Path":
                self.template_path = data[0]
            elif setting == "Sheets Path":
                self.sheets_path = data[0]
            elif setting == "All Headers":
                self.full_categories = data
            elif setting == "Headers to Extract":
                self.desired_categories = data
            elif setting == "Included Files":
                self.included_files = data
            else:
                raise ConfigReadError(f"Error: Missed Config Data, Tag [{setting}]")

        file.close()


def LoadExcelSheet(filepath, full_categories, desired_categories):
    """
    Extract Data from Excel Sheet
    :param filepath: Filepath to Spreadsheet
    :param full_categories: List of Full Categories on Spreadsheet
    :param desired_categories: Categories to Extract
    :return: Pandas Excel Sheet Data Object
    """
    excel_sheet = LoadWriter(filepath, full_categories)
    for data in excel_sheet:
        # Ignore Undesired Categories (Remove from Sheet)
        if data not in desired_categories:
            excel_sheet = excel_sheet.drop(data, 1)

        # Debugging: Print Included Columns for Verification
        """
        else:
            print(data)
        """

    # Adds Source Spreadsheet as a Data Field
    excel_sheet["List Name"] = Series([filepath[:-5]] * len(excel_sheet.index), index=excel_sheet.index)

    return excel_sheet


def LoadWriter(filepath, categories):
    """
    Load an Excel Sheet to a Pandas Excel Data Object
    :param filepath: Path to the XLSX spreadsheet file
    :param categories: Categories used in the spreadsheet
    :return: Pandas Excelsheet Data Object
    """
    return read_excel(io=filepath, header=0, skiprows=0, names=categories)


def ConcatenateSheets(src_xlsx, dest_xlsx, settings):
    """
    Iterate through Source Spreadsheet (src_xlsx)
        and add its contents to Destination Spreadsheet (dest_xlsx)
    :param src_xlsx: Source Spreadsheet
    :param dest_xlsx: Spreadsheet you are appending to
    :param settings: Settings object used to track number of concatenated entries
    :return: Appended Spreadsheet
    """
    # For debugging, you can print row["CATEGORY NAME"]
    for index, row in src_xlsx.iterrows():
        dest_xlsx = dest_xlsx.append(row)
        settings.num_entries += 1

    return dest_xlsx


def WriteToExcelSheet(template, filepath):
    """
    Write a Pandas Spreadsheet Data Object to File
    :param template: Template to start from (with header)
    :param filepath: Location to write to
    """
    writer = ExcelWriter(filepath)
    template.to_excel(writer, index=False)
    writer.save()
    writer.close()


def Main():
    # Start Timer
    start_time = time.time()

    # Load Project Configurations
    config_path = "Configurations/project.config"
    settings = Settings(config_path=config_path)

    # Load Template to Start from
    template_sheet = LoadWriter(settings.template_path, settings.desired_categories)

    # Read Each Sheet and Add it to a Master Sheet
    for file in settings.included_files:
        file_path = settings.sheets_path + file
        my_sheet = LoadExcelSheet(file_path, settings.full_categories, settings.desired_categories)
        template_sheet = ConcatenateSheets(my_sheet, template_sheet, settings)

    # Export the Finalized Master Sheet
    WriteToExcelSheet(template_sheet, settings.export_path)

    # End Timer, Calculate Resulting Time
    elapsed_time = time.time() - start_time

    # Output Resulting Statistics for Running
    print(f"\nConcatenated {settings.num_entries} Entries " \
            f"from {len(settings.included_files)} Spreadsheets " \
            f"with {len(settings.desired_categories)} Categories Each " \
            f"(from {len(settings.full_categories)} Total) " \
            f"in {round(elapsed_time, 2)} seconds")


if __name__ == '__main__':
    Main()
