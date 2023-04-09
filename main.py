import os
import re
import pandas as pd

fileName = "Technical Test.xls"
gender_mapping = {
    "F": 1,
    "M": 2
}
age_group_mapping = {
    0: "G1",
    1: "G2",
    2: "G3",
    3: "G4",
    4: "G5",
    5: "G6",
    6: "G7",
    7: "G8",
    8: "G9",
    9: "G0"
}
age_group_counts = {
    value: 1 for value in age_group_mapping.values()
}

def reformat_nric(nric_input: str) -> str:
    nric_input = nric_input.upper()
    match = re.fullmatch("^[A-Z]\d{7}[A-Z]$", nric_input)
    if match == None:
        digits = re.findall("\d{7}", nric_input)[0]
        first, last = re.findall("[A-Z]", nric_input)
        return first + digits + last
    else:
        return nric_input
    
def gender_code(gender_input: str) -> int:
    return gender_mapping[gender_input]

def age_group(age_input: int) -> str:
    return age_group_mapping[min(age_input // 10, 9)]

def study_number(age_group_input: str) -> str:
    result = "{} - {}".format(age_group_input, age_group_counts[age_group_input])
    age_group_counts[age_group_input] += 1
    return result

def populate_study_data(demographics: pd.DataFrame, extra_info: pd.DataFrame, study_data: pd.DataFrame) -> pd.DataFrame:
    new_columns = {
        "Study Number": [],
        "New NRIC": [],
        "Gender": [],
        "Age": [],
        "Marital Status": [],

        "Ethnic Group": [],
        "Address 1": [],
        "Address 2": [],
        "Contact Number": []
    }

    for nric in study_data["Old NRIC"]:
        for col_name in ["Ethnic Group", "Address 1", "Address 2", "Contact Number"]:
            try:
                feature = extra_info[extra_info["NRIC"] == nric.upper()][col_name].values[0]
            except:
                feature = None
            new_columns[col_name].append(feature)

        for col_name in ["Study Number", "New NRIC", "Gender", "Age", "Marital Status"]:
            try:
                feature = demographics[demographics["NRIC"] == nric.upper()][col_name].values[0]
            except:
                feature = None
            new_columns[col_name].append(feature)
        
    for feature in new_columns.keys():
        study_data[feature] = new_columns[feature]
    return study_data


if __name__ == "__main__":
    demographics = pd.read_excel(fileName, sheet_name="Demographics")

    demographics["New NRIC"] = demographics["NRIC"].apply(reformat_nric)
    demographics["Coding - Gender"] = demographics["Gender"].apply(gender_code)
    demographics["Age Group"] = demographics["Age"].apply(age_group)
    demographics["Study Number"] = demographics["Age Group"].apply(study_number)

    extra_info = pd.read_excel(fileName, sheet_name="Extra information")

    study_data = pd.read_excel(fileName, sheet_name="Study Data")
    study_data = populate_study_data(demographics, extra_info, study_data)

    exception_list = pd.read_excel(fileName, sheet_name="Exception List")
    unique_nric = pd.DataFrame(demographics["New NRIC"].unique(), columns=["No. of unique NRIC"])
    missing_nric = pd.DataFrame(set(demographics["NRIC"].unique()).difference(set(extra_info["NRIC"].unique())), columns = ["No. of NRIC not found in Extra Information"])
    exceptions = unique_nric.join(missing_nric)

    pivot_table = pd.read_excel(fileName, sheet_name="Pivot Table")
    pivot_age_group = pd.DataFrame(demographics["Age Group"].value_counts())
    pivot_gender = pd.DataFrame(demographics["Gender"].value_counts())
    pivot_marital_status = pd.DataFrame(demographics["Marital Status"].value_counts())
    pivot_result = pd.concat(
        [
            pd.DataFrame({"count": [""]}, index=["Age Group"]), 
            pivot_age_group, 
            pd.DataFrame({"count": [""]}, index=["Gender"]), 
            pivot_gender, 
            pd.DataFrame({"count": [""]}, index=["Marital Status"]),
            pivot_marital_status
        ], 
        axis = 0
    )

    with pd.ExcelWriter('technical_test_output.xlsx') as writer:  
        demographics.to_excel(writer, sheet_name='Demographics')
        extra_info.to_excel(writer, sheet_name='Extra information')
        study_data.to_excel(writer, sheet_name="Study Data")
        exceptions.to_excel(writer, sheet_name="Exception List")
        pivot_result.to_excel(writer, sheet_name="Pivot Table")