import os
import shutil
from openpyxl.worksheet.datavalidation import DataValidation
from datetime import datetime

class Utils:

    @staticmethod
    def duplicateTemplateLTGC(tempLTGC_Path, out, outputFilename):
        companyDir = out + "/"

        shutil.copy(tempLTGC_Path,
                    companyDir + outputFilename + ".xlsx")

        return companyDir + outputFilename + ".xlsx"

    @staticmethod
    def addingDataValidation(currentSheet, numrows):
        Category_data_val = DataValidation(type="list", formula1="=LOVCategories")
        currentSheet.add_data_validation(Category_data_val)

        CategoryID_data_val = DataValidation(type="list", formula1="=LOVCategoryID")
        currentSheet.add_data_validation(CategoryID_data_val)

        Suffix_data_val = DataValidation(type="list", formula1="=LOVSuffix")
        currentSheet.add_data_validation(Suffix_data_val)

        C_residence_region_data_val = DataValidation(type="list", formula1="=Region")
        currentSheet.add_data_validation(C_residence_region_data_val)

        C_residence_province_data_val = DataValidation(type="list", formula1="=INDIRECT(L3)")
        currentSheet.add_data_validation(C_residence_province_data_val)

        C_residence_municipality_data_val = DataValidation(type="list", formula1="=INDIRECT(M3)")
        currentSheet.add_data_validation(C_residence_municipality_data_val)

        C_residence_Barangay_data_val = DataValidation(type="list", formula1="=INDIRECT(N3)")
        currentSheet.add_data_validation(C_residence_Barangay_data_val)

        sex_data_val = DataValidation(type="list", formula1="=LOVSex")
        currentSheet.add_data_validation(sex_data_val)

        civilStatus_data_val = DataValidation(type="list", formula1="=LOVCivilStatus")
        currentSheet.add_data_validation(civilStatus_data_val)

        employmentStatus_data_val = DataValidation(type="list", formula1="=LOVEmploymentStatus")
        currentSheet.add_data_validation(employmentStatus_data_val)

        Directly_in_interaction_with_COVID_patient_data_val = DataValidation(type="list", formula1="=LOVDirectCovid")
        currentSheet.add_data_validation(Directly_in_interaction_with_COVID_patient_data_val)

        Profession_data_val = DataValidation(type="list", formula1="=LOVProfession")
        currentSheet.add_data_validation(Profession_data_val)

        ICC_of_Employer_data_val = DataValidation(type="list", formula1="=LOVProvinceHUCICCofEmployer")
        currentSheet.add_data_validation(ICC_of_Employer_data_val)

        Pregnancy_status_data_val = DataValidation(type="list", formula1="=LOVPregnancyStatus")
        currentSheet.add_data_validation(Pregnancy_status_data_val)

        YesNo_data_val = DataValidation(type="list", formula1="=LOVYesNo")
        currentSheet.add_data_validation(YesNo_data_val)

        With_Comorbidity_data_val = DataValidation(type="list", formula1="=LOVYesNone")
        currentSheet.add_data_validation(With_Comorbidity_data_val)

        Classification_of_COVID_19_data_val = DataValidation(type="list", formula1="=LOVCovidClass")
        currentSheet.add_data_validation(Classification_of_COVID_19_data_val)

        Willing_to_be_Vaccinated_data_val = DataValidation(type="list", formula1="=LOVConsent")
        currentSheet.add_data_validation(Willing_to_be_Vaccinated_data_val)

        Signup_coompletion_Time_data_val = DataValidation(type="list", formula1="=LOVWFH")
        currentSheet.add_data_validation(Signup_coompletion_Time_data_val)

        A2_Senior_data_val = DataValidation(type="list", formula1="=A2LOV")
        currentSheet.add_data_validation(A2_Senior_data_val)

        A3_With_Co_morbidity_data_val = DataValidation(type="list", formula1="=A3LOV")
        currentSheet.add_data_validation(A3_With_Co_morbidity_data_val)

        AgeRiskFactor_data_val = DataValidation(type="list", formula1="=AgeRiskFactor")  # 55-59_y/o
        currentSheet.add_data_validation(AgeRiskFactor_data_val)

        Confirmed_Vaccination_Site_data_val = DataValidation(type="list", formula1="=VaccinationSites")
        currentSheet.add_data_validation(Confirmed_Vaccination_Site_data_val)

        row = numrows + 3
        Category_data_val.add("A3:A" + str(row))
        CategoryID_data_val.add("B3:B" + str(row))
        Suffix_data_val.add("I3:I" + str(row))
        C_residence_region_data_val.add("L3:L" + str(row))
        C_residence_province_data_val.add("M3:M" + str(row))
        C_residence_municipality_data_val.add("N3:N" + str(row))
        C_residence_Barangay_data_val.add("O3:O" + str(row))
        sex_data_val.add("P3:P" + str(row))
        civilStatus_data_val.add("R3:R" + str(row))
        employmentStatus_data_val.add("S3:S" + str(row))
        Directly_in_interaction_with_COVID_patient_data_val.add("T3:T" + str(row))
        Profession_data_val.add("U3:U" + str(row))
        ICC_of_Employer_data_val.add("W3:W" + str(row))
        Pregnancy_status_data_val.add("Z3:Z" + str(row))
        YesNo_data_val.add("AA3:AA" + str(row))
        YesNo_data_val.add("AB3:AB" + str(row))
        YesNo_data_val.add("AC3:AC" + str(row))
        YesNo_data_val.add("AD3:AD" + str(row))
        YesNo_data_val.add("AE3:AE" + str(row))
        YesNo_data_val.add("AF3:AF" + str(row))
        YesNo_data_val.add("AG3:AG" + str(row))
        With_Comorbidity_data_val.add("AH3:AH" + str(row))
        YesNo_data_val.add("AI3:AI" + str(row))
        YesNo_data_val.add("AJ3:AJ" + str(row))
        YesNo_data_val.add("AK3:AK" + str(row))
        YesNo_data_val.add("AL3:AL" + str(row))
        YesNo_data_val.add("AM3:AM" + str(row))
        YesNo_data_val.add("AN3:AN" + str(row))
        YesNo_data_val.add("AO3:AO" + str(row))
        YesNo_data_val.add("AP3:AP" + str(row))
        YesNo_data_val.add("AQ3:AQ" + str(row))
        Classification_of_COVID_19_data_val.add("AS3:AS" + str(row))
        Willing_to_be_Vaccinated_data_val.add("AT3:AT" + str(row))
        A2_Senior_data_val.add("BD3:BD" + str(row))
        A3_With_Co_morbidity_data_val.add("BE3:BE" + str(row))
        # AgeRiskFactor_data_val.add("BF3:BF" + str(row))
        Confirmed_Vaccination_Site_data_val.add("BL3:BL" + str(row))

        pass

    @staticmethod
    def createLogFile(logfile, errtmpList):
        today = datetime.today()
        dateTime = today.strftime("%m%d%y%H%M%S")

        logpath = os.path.join(logfile, "log-hh_" + dateTime + ".txt")

        mode = 'a' if os.path.exists(logpath) else 'w'
        with open(logpath, mode) as f:
            for err in errtmpList:
                f.writelines(err + "\n")

        pass
