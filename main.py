import streamlit as st
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account, id_token
from google.auth.transport import requests
from google_auth_oauthlib.flow import Flow
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import io  # Import io for PDF handling
import base64


# Your client_id and client_secret from Streamlit secrets
CLIENT_ID = st.secrets["GOOGLE_AUTH_CLIENT_ID"]
CLIENT_SECRET = st.secrets["GOOGLE_AUTH_CLIENT_SECRET"]
REDIRECT_URI = st.secrets["REDIRECT_URI"]
GOOGLE_PROJECT_ID = st.secrets["GOOGLE_PROJECT_ID"]
GOOGLE_PRIVATE_KEY_ID = st.secrets["GOOGLE_PRIVATE_KEY_ID"]
GOOGLE_PRIVATE_KEY = st.secrets["GOOGLE_PRIVATE_KEY"].replace('\\n', '\n')
GOOGLE_CLIENT_EMAIL = st.secrets["GOOGLE_CLIENT_EMAIL"]
GOOGLE_CLIENT_ID = st.secrets["GOOGLE_CLIENT_ID"]
EMAIL_ADDRESS = st.secrets["EMAIL_ADDRESS"]
EMAIL_PASSWORD = st.secrets["EMAIL_PASSWORD"]

# Function to check if the user is in the allowed domain
def is_user_in_domain(email):
    return email.endswith("@geekroom.in")

# OAuth2 flow
if "credentials" not in st.session_state:
    flow = Flow.from_client_config(
        client_config={
            "web": {
                "client_id": CLIENT_ID,
                "client_secret": CLIENT_SECRET,
                "redirect_uris": [REDIRECT_URI],
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
            }
        },
        scopes=["openid", "https://www.googleapis.com/auth/userinfo.email"],
        redirect_uri=REDIRECT_URI
    )
    authorization_url, state = flow.authorization_url(prompt='select_account')
    st.session_state["state"] = state

    # Check for the authorization code from the redirect
    if "code" in st.query_params:
        code = st.query_params["code"]
        flow.fetch_token(code=code)
        st.session_state["credentials"] = flow.credentials  # Store credentials in session state

        # Clear the code from the query parameters
        del st.query_params["code"]
        st.rerun()  # Refresh the app after successful login
    
    # Redirect to the authorization URL using JavaScript
    st.markdown(f'<meta http-equiv="refresh" content="0; url={authorization_url}">', unsafe_allow_html=True)
else:
    credentials = st.session_state["credentials"]

    # If the credentials are expired, refresh them
    if credentials and credentials.expired and credentials.refresh_token:
        credentials.refresh(requests.Request())
        st.session_state["credentials"] = credentials  # Update the session state with new credentials

    # User is logged in, verify the token
    try:
        idinfo = id_token.verify_oauth2_token(
            credentials.id_token,
            requests.Request(),
            CLIENT_ID
        )
        email = idinfo['email']
        if is_user_in_domain(email):
            st.success(f'Welcome, {email}!')
            
            # # Global variables for Drive and Slides services
            # drive_service = None
            # slides_service = None

            # Get credentials from environment variables
            service_account_info = {
                "type": "service_account",
                "project_id": GOOGLE_PROJECT_ID,
                "private_key_id": GOOGLE_PRIVATE_KEY_ID,
                "private_key": GOOGLE_PRIVATE_KEY,
                "client_email": GOOGLE_CLIENT_EMAIL,
                "client_id": GOOGLE_CLIENT_ID,
                "auth_uri": "https://accounts.google.com/o/oauth2/auth",
                "token_uri": "https://oauth2.googleapis.com/token",
                "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
                "client_x509_cert_url": f"https://www.googleapis.com/robot/v1/metadata/x509/{GOOGLE_CLIENT_EMAIL}",
            }

            # Create credentials
            global drive_service, slides_service  # Declare global before use
            credentials = service_account.Credentials.from_service_account_info(service_account_info)

            # Create the Drive and Slides services
            drive_service = build('drive', 'v3', credentials=credentials)
            slides_service = build('slides', 'v1', credentials=credentials)


            # Define functions to send emails and process certificates
            def send_email(full_name, email, pdf_blob, subject, body):
                msg = MIMEMultipart()
                msg['From'] = EMAIL_ADDRESS
                msg['To'] = email
                msg['Subject'] = subject
                
                msg.attach(MIMEText(body, 'html'))

                pdf_attachment = MIMEApplication(pdf_blob, Name=f"{full_name}_Certificate.pdf")
                pdf_attachment['Content-Disposition'] = f'attachment; filename="{full_name}_Certificate.pdf"'
                msg.attach(pdf_attachment)

                try:
                    with smtplib.SMTP('smtp.gmail.com', 587) as server:
                        server.starttls()
                        server.login(EMAIL_ADDRESS, EMAIL_PASSWORD)
                        server.send_message(msg)
                except Exception as e:
                    st.error(f'Error sending email: {e}')

            def process_certificates(sheet_data, presentation_id, subject, body):
                global drive_service, slides_service  # Declare global before use
                num_rows = len(sheet_data)
                progress = st.progress(0,text=f"Mail Sent to [0/{num_rows}] rows")
                for index, row in sheet_data.iterrows():
                    full_name = row['Full Name']
                    email = row['Email']
                    body = body.replace("{Full_Name}", full_name)
                    if pd.notna(full_name) and pd.notna(email):
                        # Create a copy of the presentation
                        presentation_copy = drive_service.files().copy(
                            fileId=presentation_id,
                            body={'name': f'{full_name} Presentation'}
                        ).execute()
                        copied_presentation_id = presentation_copy['id']

                        # Replace placeholders in the copied presentation
                        requests = [{
                            'replaceAllText': {
                                'containsText': {
                                    'text': '{{Full_Name}}',
                                    'matchCase': True
                                },
                                'replaceText': full_name
                            }
                        }]
                        slides_service.presentations().batchUpdate(
                            presentationId=copied_presentation_id,
                            body={'requests': requests}
                        ).execute()

                        # Export as PDF
                        pdf_blob = drive_service.files().export(
                            fileId=copied_presentation_id,
                            mimeType='application/pdf'
                        ).execute()

                        # Send email with the attached PDF
                        send_email(full_name, email, pdf_blob, subject, body)

                        # Delete the copied presentation
                        drive_service.files().delete(fileId=copied_presentation_id).execute()
                    progress.progress((index+1)/num_rows, text=f"Mail Sent to {full_name} [{index+1}/{num_rows}]")
                progress.progress(1.0, text=f"Mail Sent to {full_name} [{num_rows}/{num_rows}] rows")

            # Streamlit app layout
            st.title('Certificate Generator')

            # Upload CSV file
            uploaded_file = st.file_uploader("Upload your CSV file", type=["csv"])

            # Instructions
            with st.expander("View Instructions"):
                with open("./README.md","r") as f:
                    st.markdown(f.read())

            # Google Slides presentation link input
            presentation_link = st.text_input("Enter the Google Slides presentation link")

            # Email template input
            with st.expander("Edit Template"):
                subject = st.text_input("Enter the Email Subject", value=f"🎉 Your Code Cubicle 3.0 Certificate Awaits! 🎉")
                email_template = st.text_area("Enter your email template (use {Full_Name} for the recipient's name):", 
                                            value=f"""<html>
    <body>
        <p>Hey {{Full_Name}}! 😊</p>
        <p>We hope you're still riding the high from <strong>Code Cubicle 3.0</strong>!</p>
        <p>Your certificate of participation is attached!</p>
        <p>If you have any questions, feel free to reach out.</p>
        <p>Cheers,<br>Arnav Kohli<br>Geek Room</p>
    </body>
    </html>""", 
                                            height=200)

                # Button to preview email and certificate
                if st.button("Preview Email & Certificate"):
                    if uploaded_file:
                        # Read the first entry in the uploaded CSV
                        data = pd.read_csv(uploaded_file)
                        if not data.empty:
                            first_entry = data.iloc[0]
                            full_name = first_entry['Full Name']
                            email = first_entry['Email']
                            
                            # Preview the email
                            preview_email = email_template.replace("{Full_Name}", full_name)
                            st.markdown("### Email Preview:")
                            st.markdown(preview_email, unsafe_allow_html=True)

                            # Create a copy of the presentation
                            if presentation_link:
                                presentation_id = presentation_link.split('/')[-2]
                                presentation_copy = drive_service.files().copy(
                                    fileId=presentation_id,
                                    body={'name': f'{full_name} Presentation'}
                                ).execute()
                                copied_presentation_id = presentation_copy['id']

                                # Replace placeholders in the copied presentation
                                requests = [{
                                    'replaceAllText': {
                                        'containsText': {
                                            'text': '{{Full_Name}}',
                                            'matchCase': True
                                        },
                                        'replaceText': full_name
                                    }
                                }]
                                slides_service.presentations().batchUpdate(
                                    presentationId=copied_presentation_id,
                                    body={'requests': requests}
                                ).execute()

                                # Export as PDF
                                pdf_blob = drive_service.files().export(
                                    fileId=copied_presentation_id,
                                    mimeType='application/pdf'
                                ).execute()

                                # Convert pdf_blob to a bytes buffer
                                pdf_buffer = io.BytesIO(pdf_blob).read()

                                # Display PDF preview
                                st.subheader(f"Certificate Preview for {full_name}")
                                base64_pdf = base64.b64encode(pdf_buffer).decode('utf-8')
                                # Embedding PDF in HTML
                                pdf_display = F'<iframe src="data:application/pdf;base64,{base64_pdf}" width="400" height="250" type="application/pdf"></iframe>'
                                # Displaying File
                                st.markdown(pdf_display, unsafe_allow_html=True)

                                # Clean up: Delete the copied presentation
                                drive_service.files().delete(fileId=copied_presentation_id).execute()
                            else:
                                st.warning("Please provide a valid Google Slides presentation link.")
                        else:
                            st.warning("The uploaded CSV file is empty.")
                    else:
                        st.warning("Please upload a CSV file to preview the email and certificate.")

            # Button to process all certificates
            if st.button("Generate Certificates"):
                if uploaded_file and presentation_link:
                    # Extract presentation ID from the link
                    presentation_id = presentation_link.split('/')[-2]
                    
                    # Read CSV file
                    data = pd.read_csv(uploaded_file)

                    

                    # Generate PDFs and send emails
                    process_certificates(data, presentation_id, subject, body=email_template)
                else:
                    st.warning("Please upload a CSV file and provide a valid Google Slides presentation link.")
        else:
            st.error("Access denied: This application is restricted to geekroom.in users.")
    except ValueError as e:
        st.error(f"Invalid token. Please log in again. Error: {e}")
        del st.session_state["credentials"]
