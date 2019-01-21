import pandas as pd
from pandas import ExcelWriter
import time

def loadExcelSheet(xlsx):
    """
    Extract Data from Excel Sheet
    :param xlsx:
    :return:
    """
    ExcelSheet = pd.read_excel(io=xlsx, sheet_name="Sheet1", header=0, skiprows=0,
                               names=["ACCT_NBR", "NAME_PRFX", "FIRST_NM", "MIDDLE_NM", "LAST_NM", "NAME_SFX",
                                      "DESIG", "FULL_NAME", "COMMON_NN", "SEX", "BIRTH_DT", "WHICH_INT",
                                      "MARTL_STS", "SPOUSE", "MBR_TYP", "CERT_TYP", "STATUS", "TIMELINE_NAME",
                                      "TIMELINE_START_DATE", "TIMELINE_END_DATE", "EMAIL", "MAIL_TO", "MAIL_COMP",
                                      "MAIL_TTLE", "MAIL_LIN1", "MAIL_LIN2", "MAIL_CITY", "MAIL_ST", "MAIL_ZIP",
                                      "MAIL_CTRY", "MAIL_CNTY", "BUS_COMP", "BUS_TTLE", "BUS_LIN1", "BUS_LIN2",
                                      "BUS_CITY", "BUS_ST", "BUS_ZIP", "BUS_CTRY", "BUS_CNTRY", "BUS_ALT", "BUS_PH",
                                      "BUS_FAX", "HOME_COMP", "HOME_TTLE", "HOME_LIN1", "HOME_LIN2", "HOME_CITY",
                                      "HOME_ST", "HOME_ZIP", "HOME_CTRY", "HOME_CNTY", "HOME_ALT", "HOME_PH",
                                      "HOME_FAX", "JOIN_DT", "DESIG_DT", "CAND_DT", "AFFIL_DT", "LIFE_DT",
                                      "RETRD_DT", "LF_RT_DT", "SEMRET_DT", "TEMPNP_DT", "REGION", "PRMY_CHP",
                                      "SECD_CHP", "OTH_CHP", "WEB_ADRS"])
    categories = ["ACCT_NBR", "FIRST_NM", "LAST_NM", "MBR_TYP", "EMAIL", "MAIL_CNTY", "BUS_PH", "CERT_TYP"]
    for data in ExcelSheet:
        if data not in categories:
            ExcelSheet = ExcelSheet.drop(data, 1)
        ### USE TO TEST IF THE PROPER COLUMNS ARE BEING REMOVED ###
        """
        else:
            print(data)
        """

    # Adds List Name as a Column at the End
    ExcelSheet["List Name"] = pd.Series([xlsx[:-5]]*len(ExcelSheet.index), index=ExcelSheet.index)

    return ExcelSheet


def loadWriter(xlsx):
    """

    :param xlsx:
    :return:
    """
    ExcelSheet = pd.read_excel(io=xlsx, sheet_name="Sheet1", header=0, skiprows=0,
                               names=["Account Number", "First Name", "Last Name", "Member Type", "Email",
                                      "County", "Phone", "Tag (Cert Type, List Name)", "Cert Type", "List Name"])
    return ExcelSheet


def concatenateSheets(src_xlsx, dest_xlsx):
    """

    :param src_xlsx:
    :param dest_xlsx:
    :return:
    """
    for index, row in src_xlsx.iterrows():
        #print(row["ACCT_NBR"])
        dest_xlsx = dest_xlsx.append(row)

    return dest_xlsx


def writeToExcelSheet(src_xlsx, dest_xlsx):
    """

    :param src_xlsx:
    :param dest_xlsx:
    :return:
    """
    writer = ExcelWriter(dest_xlsx)
    src_xlsx.to_excel(writer, 'Sheet1', index=False)
    writer.save()
    writer.close()
    return None


def main():

    io_files = ["GreatLakes-C-AI-GRS-Affiliate.xlsx", "GreatLakes-C-AI-GRS-Practicing.xlsx",
                "GreatLakes-C-AI-RRS-Affiliate.xlsx", "GreatLakes-C-AI-RRS-Practicing.xlsx",
                "GreatLakes-C-MAI-Affiliate.xlsx", "GreatLakes-C-MAI-Practicing.xlsx",
                "GreatLakes-C-SRA-Affiliate.xlsx", "GreatLakes-C-SRA-Practicing.xlsx",
                "GreatLakes-D-AI-GRS-Affiliate.xlsx", "GreatLakes-D-AI-GRS-Practicing.xlsx",
                "GreatLakes-D-AI-RRS-Affiliate.xlsx", "GreatLakes-D-AI-RRS-Practicing.xlsx",
                "GreatLakes-D-MAI-Affiliate.xlsx", "GreatLakes-D-MAI-Practicing.xlsx",
                "GreatLakes-D-SRA-Affiliate.xlsx", "GreatLakes-D-SRA-Practicing.xlsx",
                "Northern-C-AI-GRS-Affiliate.xlsx", "Northern-C-AI-GRS-Practicing.xlsx",
                "Northern-C-AI-RRS-Affiliate.xlsx", "Northern-C-AI-RRS-Practicing.xlsx",
                "Northern-C-MAI-Affiliate.xlsx", "Northern-C-MAI-Practicing.xlsx",
                "Northern-C-SRA-Affiliate.xlsx", "Northern-C-SRA-Practicing.xlsx",
                "Northern-D-AI-GRS-Affiliate.xlsx", "Northern-D-AI-GRS-Practicing.xlsx",
                "Northern-D-AI-RRS-Affiliate.xlsx", "Northern-D-AI-RRS-Practicing.xlsx",
                "Northern-D-MAI-Affiliate.xlsx", "Northern-D-MAI-Practicing.xlsx",
                "Northern-D-SRA-Affiliate.xlsx", "Northern-D-SRA-Practicing.xlsx",
                "NorthStar-C-AI-GRS-Affiliate.xlsx", "NorthStar-C-AI-GRS-Practicing.xlsx",
                "NorthStar-C-AI-RRS-Affiliate.xlsx", "NorthStar-C-AI-RRS-Practicing.xlsx",
                "NorthStar-C-MAI-Affiliate.xlsx", "NorthStar-C-MAI-Practicing.xlsx",
                "NorthStar-C-SRA-Affiliate.xlsx", "NorthStar-C-SRA-Practicing.xlsx",
                "NorthStar-D-AI-GRS-Affiliate.xlsx", "NorthStar-D-AI-GRS-Practicing.xlsx",
                "NorthStar-D-AI-RRS-Affiliate.xlsx", "NorthStar-D-AI-RRS-Practicing.xlsx",
                "NorthStar-D-MAI-Affiliate.xlsx", "NorthStar-D-MAI-Practicing.xlsx",
                "NorthStar-D-SRA-Affiliate.xlsx", "NorthStar-D-SRA-Practicing.xlsx",
                "Wisconsin-C-AI-GRS-Affiliate.xlsx", "Wisconsin-C-AI-GRS-Practicing.xlsx",
                "Wisconsin-C-AI-RRS-Affiliate.xlsx", "Wisconsin-C-AI-RRS-Practicing.xlsx",
                "Wisconsin-C-MAI-Affiliate.xlsx", "Wisconsin-C-MAI-Practicing.xlsx",
                "Wisconsin-C-SRA-Affiliate.xlsx", "Wisconsin-C-SRA-Practicing.xlsx",
                "Wisconsin-D-AI-GRS-Affiliate.xlsx", "Wisconsin-D-AI-GRS-Practicing.xlsx",
                "Wisconsin-D-AI-RRS-Affiliate.xlsx", "Wisconsin-D-AI-RRS-Practicing.xlsx",
                "Wisconsin-D-MAI-Affiliate.xlsx", "Wisconsin-D-MAI-Practicing.xlsx",
                "Wisconsin-D-SRA-Affiliate.xlsx", "Wisconsin-D-SRA-Practicing.xlsx"]

    out_file = "Out_Sheet.xlsx"
    template_sheet = loadWriter("template.xlsx")

    for file in io_files:
        my_sheet = loadExcelSheet(file)
        template_sheet = concatenateSheets(my_sheet, template_sheet)

    writeToExcelSheet(template_sheet, out_file)


if __name__ == '__main__':
    start_time = time.time()
    main()
    print("--- %s seconds ---" % (time.time() - start_time))
