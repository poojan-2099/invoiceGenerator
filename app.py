import os
import json
from flask import Flask, request, jsonify
from flask_cors import CORS
from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import inch
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
import urllib.request
import io
import requests

# Google Auth and API Libraries
import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- Configuration ---
# This configuration is now set up for Brevo (formerly Sendinblue).

# === PRODUCTION CONFIG (Brevo) ===

EMAIL_HOST = os.environ.get('EMAIL_HOST', 'smtp-relay.brevo.com')
EMAIL_PORT = int(os.environ.get('EMAIL_PORT', 587))
EMAIL_HOST_USER = os.environ.get('EMAIL_HOST_USER')
EMAIL_HOST_PASSWORD = os.environ.get('EMAIL_HOST_PASSWORD')
SENDER_EMAIL = os.environ.get('SENDER_EMAIL')    

# === GOOGLE APIS CONFIG (SHEETS & DRIVE) ===
GOOGLE_CREDENTIALS_FILE = 'credentials.json' # Used as a fallback for local dev
GOOGLE_SHEET_NAME = os.environ.get('GOOGLE_SHEET_NAME')
# The name of the specific sheet (tab) for vendors
GOOGLE_VENDORS_SHEET_NAME = 'Vendors'
GOOGLE_DRIVE_FOLDER_NAME = os.environ.get('GOOGLE_DRIVE_FOLDER_NAME')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']


# --- Flask App Initialization ---
app = Flask(__name__)
CORS(app)

## --- Helper function for Google Credentials ---
def get_google_creds():
    """
    Gets Google credentials from an environment variable or a local file.
    In production (e.g., Render), use the environment variable.
    """
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json_str:
        print("Loading Google credentials from environment variable.")
        creds_info = json.loads(creds_json_str)
        creds = Credentials.from_service_account_info(creds_info, scopes=SCOPES)
    elif os.path.exists(GOOGLE_CREDENTIALS_FILE):
        print(f"Loading Google credentials from '{GOOGLE_CREDENTIALS_FILE}'.")
        creds = Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
    else:
        raise FileNotFoundError("Google credentials not found. Set GOOGLE_CREDENTIALS_JSON or provide a 'credentials.json' file.")
    return creds

# --- Google Drive Integration ---
def upload_to_google_drive(file_path, file_name):
    """Uploads a file to a specific folder in Google Drive."""
    try:
        creds = get_google_creds()
        service = build('drive', 'v3', credentials=creds)

        folder_id = None
        query = f"mimeType='application/vnd.google-apps.folder' and name='{GOOGLE_DRIVE_FOLDER_NAME}' and trashed=false"
        response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        
        if not response.get('files'):
            print(f"Folder '{GOOGLE_DRIVE_FOLDER_NAME}' not found. Creating it...")
            folder_metadata = {'name': GOOGLE_DRIVE_FOLDER_NAME, 'mimeType': 'application/vnd.google-apps.folder'}
            folder = service.files().create(body=folder_metadata, fields='id').execute()
            folder_id = folder.get('id')
        else:
            folder_id = response.get('files')[0].get('id')

        file_metadata = {'name': file_name, 'parents': [folder_id]}
        media = MediaFileUpload(file_path, mimetype='application/pdf')
        
        print(f"Uploading '{file_name}' to Google Drive...")
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print("Successfully uploaded file to Google Drive.")

    except Exception as e:
        print(f"An error occurred with Google Drive integration: {e}")

# --- Google Sheets Integration ---
def add_invoice_to_sheet(data):
    """Adds a new row with invoice data to the specified Google Sheet."""
    try:
        creds = get_google_creds()
        client = gspread.authorize(creds)
        
        sh = client.open(GOOGLE_SHEET_NAME)
        worksheet = sh.worksheet('Invoices') # Assuming the invoice sheet is named 'Invoices'
        
        total = float(data['quantity']) * float(data['price'])
        new_row = [datetime.now().strftime('%Y-%m-%d %H:%M:%S'), data['date'], data['vendor_name'], data['vendor_email'], data['item'], data['quantity'], data['price'], total, data.get('notes', '')]

        if len(worksheet.get_all_records()) == 0:
            worksheet.append_row(["Timestamp", "Invoice Date", "Vendor Name", "Vendor Email", "Item", "Quantity", "Unit Price", "Total", "Notes"])

        worksheet.append_row(new_row)
        print("Successfully added new row to Google Sheet.")
    except Exception as e:
        print(f"An error occurred with Google Sheets integration: {e}")

# --- PDF Generation Logic ---
def create_invoice_pdf(data):
    file_name = f"invoice_{data['vendor_name'].replace(' ', '_')}_{data['date']}.pdf"
    temp_dir = '/tmp'
    if not os.path.exists(temp_dir):
        os.makedirs(temp_dir)
    pdf_path = os.path.join(temp_dir, file_name)

    doc = SimpleDocTemplate(pdf_path, pagesize=letter, rightMargin=72, leftMargin=72, topMargin=72, bottomMargin=18)
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name='CenterAlign', alignment=TA_CENTER, fontSize=24, spaceAfter=10, leading=30))
    styles.add(ParagraphStyle(name='CompanyInfo', alignment=TA_RIGHT, leading=14))
    elements = []
    
    logo_url = 'https://malkitsweetsandcatering.com/img/logo.png'
    try:
        response = requests.get(logo_url, stream=True)
        response.raise_for_status()
        logo = Image(io.BytesIO(response.content), width=1.5*inch, height=1*inch, kind='bound')
    except Exception as e:
        print(f"Warning: Could not fetch logo from URL. Error: {e}")
        logo = Paragraph("Your Company", styles['h2'])

    header_data = [[logo, Paragraph("INVOICE", styles['CenterAlign'])]]
    header_table = Table(header_data, colWidths=[2*inch, 4.5*inch])
    header_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (1, 0), (1, 0), 'RIGHT')]))
    elements.append(header_table)
    elements.append(Spacer(1, 0.2*inch))
    
    company_info_data = [[Paragraph("<b>Your Company Name</b><br/>123 Sweet Lane<br/>Pastryville, PV 54321", styles['Normal']), Paragraph(f"<b>Invoice #:</b> INV-001<br/><b>Date:</b> {datetime.strptime(data['date'], '%Y-%m-%d').strftime('%B %d, %Y')}", styles['CompanyInfo'])]]
    company_info_table = Table(company_info_data, colWidths=[3.5*inch, 3*inch])
    company_info_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'TOP')]))
    elements.append(company_info_table)
    elements.append(Spacer(1, 0.4*inch))
    elements.append(Paragraph(f"<b>BILL TO:</b><br/>{data['vendor_name']}<br/>{data['vendor_email']}", styles['Normal']))
    elements.append(Spacer(1, 0.4*inch))

    table_data = [['Item Description', 'Quantity', 'Unit Price', 'Total'], [data['item'], f"{data['quantity']}", f"${float(data['price']):.2f}", f"${float(data['quantity']) * float(data['price']):.2f}"]]
    invoice_table = Table(table_data, colWidths=[3.5*inch, 1*inch, 1*inch, 1*inch])
    invoice_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4F46E5')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('BOTTOMPADDING', (0, 0), (-1, 0), 12), ('BACKGROUND', (0, 1), (-1, -1), colors.beige), ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    elements.append(invoice_table)
    elements.append(Spacer(1, 0.2 * inch))

    total_price = float(data['quantity']) * float(data['price'])
    total_data = [['', '', 'Total:', f'${total_price:.2f}']]
    total_table = Table(total_data, colWidths=[3.5*inch, 1*inch, 1*inch, 1*inch])
    total_table.setStyle(TableStyle([('ALIGN', (2, 0), (2, 0), 'RIGHT'), ('FONTNAME', (2, 0), (3, 0), 'Helvetica-Bold'), ('GRID', (2, 0), (3, 0), 1, colors.black), ('BACKGROUND', (2,0), (3,0), colors.lightgrey)]))
    elements.append(total_table)
    elements.append(Spacer(1, 0.5 * inch))

    if data.get('notes'):
        elements.append(Paragraph("<b>Notes:</b>", styles['Normal']))
        elements.append(Paragraph(data['notes'], styles['Normal']))

    doc.build(elements)
    return pdf_path, file_name

# --- Email Sending Logic ---
def send_email_with_attachment(recipient_email, subject, body, file_path):
    msg = MIMEMultipart()
    msg['From'] = SENDER_EMAIL
    msg['To'] = recipient_email
    msg['Bcc'] = SENDER_EMAIL
    msg['Subject'] = subject
    msg.attach(MIMEText(body, 'plain'))
    with open(file_path, "rb") as attachment:
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(file_path)}")
    msg.attach(part)
    try:
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        if EMAIL_HOST_PASSWORD:
            server.starttls()
            server.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
        all_recipients = [recipient_email, SENDER_EMAIL]
        server.sendmail(SENDER_EMAIL, all_recipients, msg.as_string())
        server.quit()
        print(f"Email sent successfully to {recipient_email} and BCC'd to {SENDER_EMAIL}")
    except Exception as e:
        print(f"Failed to send email: {e}")

# --- API Endpoints ---
@app.route('/get-vendors', methods=['GET'])
def get_vendors():
    """Fetches all vendors from the Google Sheet and standardizes keys to lowercase."""
    try:
        creds = get_google_creds()
        client = gspread.authorize(creds)
        sh = client.open(GOOGLE_SHEET_NAME)
        worksheet = sh.worksheet(GOOGLE_VENDORS_SHEET_NAME)
        records = worksheet.get_all_records()
        
        # Standardize keys to lowercase to ensure consistency for the frontend
        standardized_vendors = []
        for record in records:
            # Create a new dictionary with lowercase keys
            standardized_record = {key.lower(): value for key, value in record.items()}
            standardized_vendors.append(standardized_record)
            
        return jsonify(standardized_vendors), 200
    except Exception as e:
        print(f"Error fetching vendors: {e}")
        return jsonify({"error": "Could not fetch vendors from the source."}), 500

@app.route('/add-vendor', methods=['POST'])
def add_vendor():
    """Adds a new vendor to the Google Sheet."""
    try:
        data = request.get_json()
        if not data or 'name' not in data or 'email' not in data:
            return jsonify({"error": "Missing vendor name or email."}), 400
        
        creds = get_google_creds()
        client = gspread.authorize(creds)
        sh = client.open(GOOGLE_SHEET_NAME)
        worksheet = sh.worksheet(GOOGLE_VENDORS_SHEET_NAME)
        
        new_row = [data['name'], data['email']]
        worksheet.append_row(new_row)
        
        return jsonify({"message": "Vendor added successfully."}), 201
    except Exception as e:
        print(f"Error adding vendor: {e}")
        return jsonify({"error": "Could not add vendor."}), 500

@app.route('/generate-invoice', methods=['POST'])
def generate_invoice():
    """API endpoint for the complete invoice generation and storage process."""
    try:
        data = request.get_json()
        required_fields = ['vendor_name', 'vendor_email', 'date', 'item', 'quantity', 'price']
        if not all(field in data for field in required_fields):
            return jsonify({"error": "Missing required fields"}), 400
        
        pdf_path, pdf_filename = create_invoice_pdf(data)
        subject = f"Invoice from Your Company for {data['item']}"
        body = f"Hello {data['vendor_name']},\n\nPlease find attached the invoice for your recent order.\n\nThank you,\nYour Company Name"
        send_email_with_attachment(data['vendor_email'], subject, body, pdf_path)
        add_invoice_to_sheet(data)
        upload_to_google_drive(pdf_path, pdf_filename)
        
        if os.path.exists(pdf_path):
            os.remove(pdf_path)
        return jsonify({"message": f"Process complete for invoice to {data['vendor_email']}."}), 200

    except Exception as e:
        print(f"An unexpected error occurred in the main process: {e}")
        return jsonify({"error": "An internal server error occurred."}), 500

# --- Startup Check ---
def check_env_vars():
    """Checks for required environment variables and credentials before starting."""
    print("Checking for required environment variables...")
    required_vars = ['EMAIL_HOST_USER', 'EMAIL_HOST_PASSWORD', 'SENDER_EMAIL', 'GOOGLE_SHEET_NAME', 'GOOGLE_DRIVE_FOLDER_NAME']
    missing_vars = [var for var in required_vars if not os.getenv(var)]
    
    if missing_vars:
        raise EnvironmentError(f"Startup failed: Missing required environment variables: {', '.join(missing_vars)}")
    
    try:
        get_google_creds()
        print("Google credentials loaded successfully.")
    except Exception as e:
        raise EnvironmentError(f"Startup failed: Could not load Google credentials. Error: {e}")
    
    print("Environment variables and credentials check passed.")

# --- Run the App ---
if __name__ == '__main__':
    # For local development, you can use a .env file.
    from dotenv import load_dotenv
    load_dotenv()

    check_env_vars()
    app.run(debug=True)

