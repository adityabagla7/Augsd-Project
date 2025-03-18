import streamlit as st
import pandas as pd
import zipfile
import os
from io import BytesIO
from test4app import process_allocation
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Set the title and header with color
st.markdown("<h1 style='text-align: center; color: blue;'>BITS Pilani KK Birla Goa Campus</h1>", unsafe_allow_html=True)
st.markdown("<h2 style='text-align: center; color: green;'>Invigilator Assignment</h2>", unsafe_allow_html=True)

# Create columns for file uploaders
col1, col2 = st.columns(2)

with col1:
    file1 = st.file_uploader("Upload the first Excel file (ICS.xlsx)", type=["xlsx", "xls"], help="Add data for IC")
    file2 = st.file_uploader("Upload the second Excel file (masterFile.xlsx)", type=["xlsx", "xls"], help="Add master File data")

with col2: 
    guidelines_pdf = st.file_uploader("Upload the Guidelines PDF", type=["pdf"], help="Add Guidelines PDF")

# Display the uploaded files
if file1 is not None and file2 is not None and guidelines_pdf is not None:
    df1_dict = pd.read_excel(file1, sheet_name=None)
    df2_dict = pd.read_excel(file2, sheet_name=None)
    
    st.markdown("<h3 style='color: purple;'>First Excel File (ICS.xlsx):</h3>", unsafe_allow_html=True)
    for sheet_name, df in df1_dict.items():
        st.write(f"Sheet: {sheet_name}")
        df = df.convert_dtypes()
        st.dataframe(df)
    
    st.markdown("<h3 style='color: purple;'>Second Excel File (masterFile.xlsx):</h3>", unsafe_allow_html=True)
    for sheet_name, df in df2_dict.items():
        st.write(f"Sheet: {sheet_name}")
        st.dataframe(df)
    
    # Add a run button with color
    if st.button("Run", key="run_button", help="Click to process the uploaded files"):
        st.write("Processing the uploaded files and wait for Download Button...")
        # Save the uploaded files to disk
        file1_path = 'ICS.xlsx'
        file2_path = 'masterFile.xlsx'
        guidelines_pdf_path = 'guidelines.pdf'
        
        with pd.ExcelWriter(file1_path) as writer:
            for sheet_name, df in df1_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        with pd.ExcelWriter(file2_path) as writer:
            for sheet_name, df in df2_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        with open(guidelines_pdf_path, "wb") as f:
            f.write(guidelines_pdf.getbuffer())
        
        # Run the backend code
        process_allocation(file1_path, file2_path)
        st.write("Processing complete")
        st.write("Processing complete\n")
        st.write("Click the button below to download the output files as a ZIP archive.")
        st.write("You can also send the output files via mail.")
        # Create a zip file of the output directory
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for root, _, files in os.walk('output'):
                for file in files:
                    zip_file.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), 'output'))
        
        zip_buffer.seek(0)
        
        # Add a download button for the zip file with color
        st.download_button(
            label="Download ZIP",
            data=zip_buffer,
            file_name="output_files.zip",
            mime="application/zip",
            key="download_button",
            help="Click to download the output files as a ZIP archive"
        )
        
        # Add a button to send mail with color
        if st.button("Send Mail", key="send_mail_button", help="Click to send the output files via mail"):
            st.write("Sending mail with the attached files...")

            # Mail sending logic
            def attach_file(msg, file_data, filename):
                try:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(file_data)
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {filename}',
                    )
                    msg.attach(part)
                except Exception as e:
                    st.write(f'Failed to attach file {filename}: {e}')

            def send_email(subject, body, to_email, from_email, smtp_server, smtp_port, smtp_user, smtp_password, attachments=None):
                msg = MIMEMultipart()
                msg['From'] = from_email
                msg['To'] = to_email
                msg['Subject'] = subject

                msg.attach(MIMEText(body, 'plain'))

                if attachments:
                    for attachment in attachments:
                        attach_file(msg, attachment['data'], attachment['filename'])

                try:
                    with smtplib.SMTP_SSL(smtp_server, smtp_port) as server:
                        server.login(smtp_user, smtp_password)
                        server.send_message(msg)
                        st.write(f'Email sent successfully to {to_email}!')
                except Exception as e:
                    st.write(f'Failed to send email to {to_email}: {e}')

            SUBJECT = 'Test Email'
            BODY = '''Dear Sir/ Madam,
            PFA file with instructions and excel file containing
            1) Invigilator
            2) Rooms
            3) Seating Arrangement'''
            FROM_EMAIL = 'adityabagla0044@gmail.com'
            SMTP_SERVER = 'smtp.gmail.com'
            SMTP_PORT = 465
            SMTP_USER = 'adityabagla0044@gmail.com'
            SMTP_PASSWORD = 'your_email_app_password'  # Replace with your email app password

            # Read attachments from the zip buffer
            zip_buffer.seek(0)
            with zipfile.ZipFile(zip_buffer, 'r') as zip_file:
                attachment_files = []
                for file_info in zip_file.infolist():
                    if file_info.filename.endswith('.xlsx'):
                        with zip_file.open(file_info.filename) as file:
                            file_data = file.read()
                            attachment_files.append({'filename': os.path.basename(file_info.filename), 'data': file_data})
                # Add the PDF attachment
                with open(guidelines_pdf_path, 'rb') as pdf_file:
                    pdf_data = pdf_file.read()
                    attachment_files.append({'filename': os.path.basename(guidelines_pdf_path), 'data': pdf_data})

            # Send email
            to_email = 'f20220497@goa.bits-pilani.ac.in'  # Replace with recipient's email
            send_email(SUBJECT, BODY, to_email, FROM_EMAIL, SMTP_SERVER, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, attachments=attachment_files)

            st.write("All emails sent successfully!")

else:
    st.write("Please upload both Excel files and the Guidelines PDF to proceed.")
