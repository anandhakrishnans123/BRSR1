import streamlit as st
import pandas as pd
import numpy as np
import random
from io import BytesIO

# Define a mapping dictionary to map client column names to template column names
column_mapping1 = {
    'Employee Name': 'Emp_Name',
    'Designation': 'Designation_Cat',
    'Join Date': 'Join_date',
    'Code': 'Emp_Id',
    'Gender': 'Gender',
    'Department':'Designation'
}

column_mapping2 = {
    'Date':'Train_date',
    'Emp Code':'Emp_id',
    'Session':'Train_Topic'
}

# Define custom lists
custom_lst1 = ['Status','Company Name']
custom_lst2 = ['Duration','Grade','Gender','Department','Units']

# Upload files
st.title("Client Data Processing")

client_file1 = st.file_uploader("Upload the Client Workbook for Employment Data", type="xlsx")
client_file2 = st.file_uploader("Upload the Client Workbook for Training Data", type="xlsx")
template_file = 'BRSR-Report_Data-Template.xlsx'  # Load template directly from local (as per your request)

def create_excel_file(df1, df2=None, template_sheets=None):
    with BytesIO() as buffer:
        with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
            # Write processed data
            df1.to_excel(writer, sheet_name='Trans_Emp-Work', index=False)
            if df2 is not None:
                df2.to_excel(writer, sheet_name='Trans_Trainings', index=False)
            # Write remaining template sheets that were not processed
            if template_sheets:
                for sheet_name, df in template_sheets.items():
                    if sheet_name not in ['Trans_Emp-Work', 'Trans_Trainings']:
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
        return buffer.getvalue()

if st.button("Generate Excel File"):
    if client_file1:
        # Process Employment Data
        client_df1 = pd.read_excel(client_file1, sheet_name=None)
        for sheet_name, df in client_df1.items():
            client_data1 = client_df1[sheet_name]

        template_df = pd.read_excel(template_file, sheet_name=None)
        template_data1 = template_df['Trans_Emp-Work']

        preserved_header1 = template_data1.iloc[:1, :]
        matched_data1 = pd.DataFrame(columns=template_data1.columns)

        for i in custom_lst1:
            if i in client_data1.columns:
                matched_data1[i] = client_data1[i]

        for client_col, template_col in column_mapping1.items():
            if client_col in client_data1.columns and template_col in template_data1.columns:
                matched_data1[template_col] = client_data1[client_col]

        if 'Join Date' in client_data1.columns:
            matched_data1['Join_date'] = pd.to_datetime(matched_data1['Join_date']).dt.date

        final_data1 = pd.concat([preserved_header1, matched_data1], ignore_index=True)

        # Fill NaN values and set specific columns
        specific_date = pd.to_datetime('3/30/2024').date()
        final_data1['Resdate'] = specific_date
        final_data1.loc[final_data1["Company Name"] == "Shreyas Shipping and Logistics Limited", "Org_Code"] = "SSL"
        final_data1.loc[final_data1["Company Name"] == "Transworld Logistics FZE","Org_Code"] = "FZE"
        final_data1.loc[final_data1["Company Name"] == "TRANSWORLD LOGISTICS DWC LLC", "Org_Code"] = "DWC"
        final_data1["Org_Code"].fillna("DWC", inplace=True)

        columns_to_replace = [
            "Health_Insurance", "Day_Care_Facility", "Accident_Insurance",
            "Gratuity", "Min_wage_applicability", "Severity_Work_Related_Injury",
            "Membership_in_association&Unions", "Other_Retire_Benefit",
            "Availed_Maternity_Leave", "Deduction_for_Retire_Benefit",
            "Provident_Fund", "Differently-disabled", "ESI", "Work_Related_Injury", 
            "Parternity_Benefit", "Maternity_Benefit"
        ]

        final_data1[columns_to_replace] = final_data1[columns_to_replace].replace({np.nan: "No"})

        final_data1['Emp_type'] = "Full Time"
        final_data1['Remuneration'] = np.random.randint(10000, 100000, size=len(final_data1))
        final_data1['Min_wage'] = np.random.randint(10000, 100000, size=len(final_data1))

        def categorize_designation(designation):
            designation = designation.lower()
            if 'senior manager' in designation:
                return 'Senior Management'
            elif 'manager' in designation:
                return 'Manager'
            elif 'executive' in designation:
                return 'Executive'
            elif 'management' in designation:
                return 'Management'
            else:
                return 'Staff'

        final_data1['Designation_Cat'] = final_data1['Designation_Cat'].apply(categorize_designation)

        if client_file2:
            # Process Training Data
            client_df2 = pd.read_excel(client_file2, sheet_name=None)
            for sheet_name, df in client_df2.items():
                client_data2 = client_df2[sheet_name]

            template_data2 = template_df['Trans_Trainings']

            preserved_header2 = template_data2.iloc[:1, :]
            matched_data2 = pd.DataFrame(columns=template_data2.columns)

            for i in custom_lst2:
                if i in client_data2.columns:
                    matched_data2[i] = client_data2[i]

            for client_col, template_col in column_mapping2.items():
                if client_col in client_data2.columns and template_col in template_data2.columns:
                    matched_data2[template_col] = client_data2[client_col]

            if 'Date' in client_data2.columns:
                matched_data2['Train_date'] = pd.to_datetime(matched_data2['Train_date']).dt.date

            final_data2 = pd.concat([preserved_header2, matched_data2], ignore_index=True)

            final_data2['Resdate'] = specific_date
            final_data2.loc[final_data2["Units"] == "Shreyas Shipping and Logistics Limited", "Org_Code"] = "SSL"
            final_data2.loc[final_data2["Units"] == "TRANSWORLD LOGISTICS FZE", "Org_Code"] = "FZE"
            final_data2.loc[final_data2["Units"] == "TRANSWORLD LOGISTICS DWC LLC", "Org_Code"] = "DWC"
            final_data2["Org_Code"].fillna("DWC", inplace=True)

            final_data2['Train_type'] = final_data2['Train_type'].apply(lambda x: random.choice(['Internal']) if pd.isna(x) else x)
            final_data2['Trainee_detail'] = final_data2['Trainee_detail'].apply(lambda x: random.choice(['employee']) if pd.isna(x) else x)
            final_data2['Train_mode'] = final_data2['Train_mode'].apply(lambda x: random.choice(['Online']) if pd.isna(x) else x)
            final_data2['Train_cat'] = final_data2['Train_cat'].apply(lambda x: random.choice(['Others']) if pd.isna(x) else x)

            final_data2.dropna(inplace=True)

            # Get template sheets excluding the processed ones
            remaining_sheets = {k: v for k, v in template_df.items() if k not in ['Trans_Emp-Work', 'Trans_Trainings']}
            excel_data = create_excel_file(final_data1, final_data2, remaining_sheets)
            st.download_button(label="Download Processed Data", data=excel_data, file_name="processed_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        else:
            remaining_sheets = {k: v for k, v in template_df.items() if k != 'Trans_Emp-Work'}
            excel_data = create_excel_file(final_data1, template_sheets=remaining_sheets)
            st.download_button(label="Download Processed Data", data=excel_data, file_name="processed_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        st.warning("Please upload the Employment Data file.")
