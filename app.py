import streamlit as st
import pandas as pd
import math
import io
import zipfile
from datetime import datetime


# set page configurations
st.set_page_config(
    page_title="Excel Cleaner",
    page_icon="🤖",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# the custom CSS lives here:
hide_default_format = """
        <style>
            .reportview-container .main footer {visibility: hidden;}    
            #MainMenu, footer {visibility: hidden;}
            section.main > div:has(~ footer ) {
                padding-bottom: 1px;
                padding-left: 20px;
                padding-right: 40px;
                padding-top: 25px;
            }
            [data-testid="collapsedControl"] {
                display: none;
                } 
            [data-testid="stDecoration"] {
                background-image: linear-gradient(90deg, rgb(25, 130, 196), rgb(25, 130, 196));
                height: 20%;
                }
            div.stActionButton{visibility: hidden;}
        </style>
       """

# inject the CSS
st.markdown(hide_default_format, unsafe_allow_html=True)

# main title
st.markdown(
    "<p style='color:#000000; font-weight: 900; font-size: 46px'>ABI Excel Cleaning Robot</p>", unsafe_allow_html=True)

st.write("")
st.write("")


# Function to clean each dataframe that's uploaded
def clean_dataframe(df):
    # Insert the new columns at the beginning of the DataFrame
    df.insert(0, 'Prime Contractor/Vendor:',
              '', allow_duplicates=True)
    df.insert(0, 'Current Invoice Amount:',
              '', allow_duplicates=True)
    df.insert(0, 'Total Contract or Work Order Amount:',
              '', allow_duplicates=True)
    df.insert(0, 'Work Order #', '', allow_duplicates=True)
    df.insert(0, 'Project #', '', allow_duplicates=True)
    df.insert(0, 'Project or Work Order Name:',
              '', allow_duplicates=True)
    df.insert(0, 'Contract #:', '', allow_duplicates=True)
    df.insert(0, 'Contract Name:', '', allow_duplicates=True)
    df.insert(0, 'Invoice #', '', allow_duplicates=True)
    df.insert(0, 'Invoice Date:', '', allow_duplicates=True)

    # set renaming column schema to be used in spreadsheet
    mapping = {
        df.columns[10]: 'Prime/Sub',
        df.columns[11]: 'Vendor/Subcontractor',
        df.columns[12]: 'Certifying Agency',
        df.columns[13]: 'Race/Ethnicity',
        df.columns[14]: 'Additional DBE Types',
        df.columns[15]: 'Current Invoice Amount ($)',
        df.columns[16]: 'Total Contracted Amount ($)',
        df.columns[17]: 'Total Invoiced to Date ($)',
        df.columns[18]: 'temp'
    }

    df.rename(columns=mapping, inplace=True)

    # Now grab the values that have been inputted into the form and use to populate other columns
    # first, invoice date
    invoice_date = pd.to_datetime(df['Vendor/Subcontractor'].iloc[0]).strftime(
        '%m/%d/%Y')
    df['Invoice Date:'] = invoice_date

    # second, invoice #
    invoice_number = df['Vendor/Subcontractor'].iloc[1]
    df['Invoice #'] = invoice_number

    # third, contract name
    contract_name = df['Vendor/Subcontractor'].iloc[2]
    df['Contract Name:'] = contract_name

    # fourth, contract #
    contract_number = df['Vendor/Subcontractor'].iloc[3]
    df['Contract #:'] = contract_number

    # fifth, project or work order name
    project_wo_name = df['Vendor/Subcontractor'].iloc[4]
    df['Project or Work Order Name:'] = project_wo_name

    # sixth, project #
    project_number = df['Vendor/Subcontractor'].iloc[5]
    df['Project #'] = project_number

    # seventh, work order # (may or may not be blank, so have to check)
    wo_number = df['Vendor/Subcontractor'].iloc[6]
    if math.isnan(wo_number):
        df['Work Order #'] = 'N/A'
    else:
        df['Work Order #'] = wo_number

    # eigth, Total Contract or Work Order Amount:
    total_contract_wo_amount = df['Vendor/Subcontractor'].iloc[7]
    df['Total Contract or Work Order Amount:'] = total_contract_wo_amount

    # ninth, Current Invoice Amount:
    current_invoice_amt = df['Vendor/Subcontractor'].iloc[8]
    df['Current Invoice Amount:'] = current_invoice_amt

    # tenth, prime contractor/vendor
    prime_contractor_vendor = df['Vendor/Subcontractor'].iloc[10]
    df['Prime Contractor/Vendor:'] = prime_contractor_vendor

    # remaining variables from the Excel columns
    vendor_subcontractors = df['Prime/Sub'].iloc[13:36].dropna()
    row_length = vendor_subcontractors.shape[0]

    # we'll come back to this column, but for now, clear it out since we already assigned the values needed above
    df['Prime/Sub'] = ''

    # clear out and then assign these remaining variables to their respective "true" dataframe columns
    with pd.option_context('mode.chained_assignment', None):
        # do the vendor/subcontractor column
        certifying_agency = df['Vendor/Subcontractor'].iloc[13:13+row_length]
        df['Vendor/Subcontractor'] = ''
        df['Vendor/Subcontractor'].iloc[0:row_length] = vendor_subcontractors

        # do the certifying agency column
        race_ethnicity = df['Certifying Agency'].iloc[13:13+row_length]
        df['Certifying Agency'] = ''
        df['Certifying Agency'].iloc[0:row_length] = certifying_agency

        # do the race/ethnicity column
        df['Race/Ethnicity'] = df['Race/Ethnicity'].astype(object)
        df['Race/Ethnicity'].iloc[0:row_length] = race_ethnicity

        # do the additional DBE types column
        additional_DBE_types = df['Additional DBE Types'].iloc[13:13+row_length]
        df['Additional DBE Types'] = ''
        df['Additional DBE Types'].iloc[0:row_length] = additional_DBE_types

        # do the current invoice amount column
        current_invoice_amt_2 = df['Total Contracted Amount ($)'].iloc[13:13 +
                                                                       row_length]
        df['Current Invoice Amount ($)'] = df['Current Invoice Amount ($)'].astype(
            object)
        df['Current Invoice Amount ($)'].iloc[0:row_length] = current_invoice_amt_2

        # do the total contracted amount
        total_contracted_amt = df['Total Invoiced to Date ($)'].iloc[13:13 +
                                                                     row_length]
        df['Total Contracted Amount ($)'] = ''
        df['Total Contracted Amount ($)'].iloc[0:row_length] = total_contracted_amt

        # do the total invoiced to date
        total_invoiced_to_date = df['temp'].iloc[13:13+row_length]
        df['Total Invoiced to Date ($)'] = ''
        df['Total Invoiced to Date ($)'].iloc[0:row_length] = total_invoiced_to_date

    # drop the temp column
    df.drop(columns='temp', inplace=True)

    # create the Prime/Sub column by comparing 2 other columns
    def compare_values(row):
        if row['Prime Contractor/Vendor:'] == row['Vendor/Subcontractor']:
            return 'Prime'
        else:
            return 'Sub'
    df['Prime/Sub'] = df.apply(compare_values, axis=1)

    # now, just take the amount of rows corresponding to the number of Vendor/subcontractors
    df = df.head(row_length)

    # Fill in NaN values with the string 'N/A'
    df['Race/Ethnicity'].fillna('N/A', inplace=True)
    df['Additional DBE Types'].fillna('N/A', inplace=True)

    return df


# Function to handle file uploading and cleaning
def handle_upload():

    uploaded_files = st.file_uploader(
        label="Upload excel file(s) to be cleaned",
        accept_multiple_files=True
    )

    number_of_files = len(uploaded_files)

    cleaned_dataframes = {}

    if uploaded_files:
        for file in uploaded_files:
            try:
                # Read Excel file into a Pandas DataFrame
                df = pd.read_excel(file)

                # Clean the DataFrame (modify this based on your cleaning requirements)
                cleaned_df = clean_dataframe(df)

                # Get the filename without extension
                filename = file.name.split('.')[0]

                # Save cleaned dataframe with the original filename
                cleaned_dataframes[filename] = cleaned_df

            except Exception as e:
                st.write(
                    f'Yo, Layla! Error. Check {file.name}')
                continue

        # Create a ZipFile object to store individual Excel files
        timestamp = datetime.now().strftime("%m-%d-%Y_%I.%M%p")
        zip_file_name = f"cleaned_files_{timestamp}.zip"

        buffer_zip = io.BytesIO()

        # Create a ZipFile object to store individual Excel files
        with zipfile.ZipFile(buffer_zip, 'w') as zip_file:
            # Save each dataframe as a separate Excel file in the zip archive
            for filename, cleaned_df in cleaned_dataframes.items():
                excel_file = io.BytesIO()
                with pd.ExcelWriter(excel_file, engine='xlsxwriter') as writer:
                    cleaned_df.to_excel(
                        writer,
                        sheet_name='Table 1',
                        index=False
                    )

                    # Get the xlsxwriter workbook and worksheet objects
                    workbook = writer.book
                    worksheet = writer.sheets["Table 1"]

                    # Define a currency format with dollar signs and thousands separators
                    dollar_format = workbook.add_format(
                        {'num_format': '$#,##0.00'})

                    # Set the number format for the specified columns
                    # Adjust column range as needed
                    worksheet.set_column("H:I", None, dollar_format)
                    worksheet.set_column("P:R", None, dollar_format)

                    # autofit columns
                    worksheet.autofit()

                    # Close the Pandas Excel writer and save the Excel file
                    writer.close()

                excel_file.seek(0)
                zip_file.writestr(
                    f'{filename}.xlsx', excel_file.read())

        # Close the ZipFile and output the zip file to the buffer
        buffer_zip.seek(0)

        # show user how many files were uploaded
        st.markdown(
            f"<p style='color:#000000; font-weight: 600; font-size: 18px'><em>Total files uploaded: {number_of_files}</em></p>", unsafe_allow_html=True)

        st.download_button(
            label="Clean & download to zip",
            data=buffer_zip,
            file_name=zip_file_name,
            mime="application/zip"
        )


handle_upload()
