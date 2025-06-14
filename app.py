import os
import json
import re
import io
import requests
import smtplib
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_RIGHT, TA_CENTER
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload

# --- Configuration ---
EMAIL_HOST = os.environ.get('EMAIL_HOST', 'smtp-relay.brevo.com')  # Brevo SMTP server
EMAIL_PORT = int(os.environ.get('EMAIL_PORT', 587))  # Brevo SMTP port
EMAIL_HOST_USER = os.environ.get('EMAIL_HOST_USER')  # Your Brevo SMTP username
EMAIL_HOST_PASSWORD = os.environ.get('EMAIL_HOST_PASSWORD')  # Your Brevo SMTP key
SENDER_EMAIL = os.environ.get('SENDER_EMAIL')  # Your verified sender email

GOOGLE_CREDENTIALS_FILE = 'credentials.json'
GOOGLE_SHEET_NAME = os.environ.get('GOOGLE_SHEET_NAME')
GOOGLE_VENDORS_SHEET_NAME = 'Vendors'
GOOGLE_INVOICES_SHEET_NAME = 'Invoices'
GOOGLE_DRAFTS_SHEET_NAME = 'Drafts'
GOOGLE_SWEETS_SHEET_NAME = 'Sweets'
GOOGLE_DRIVE_FOLDER_NAME = os.environ.get('GOOGLE_DRIVE_FOLDER_NAME')
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

# --- Flask App Initialization ---
app = Flask(__name__)
CORS(app, resources={
    r"/*": {
        "origins": [
            "https://poojan-2099.github.io",
            "http://localhost:5000",
            "http://127.0.0.1:5000"
        ],
        "methods": ["GET", "POST", "OPTIONS"],
        "allow_headers": ["Content-Type"]
    }
})

# --- Helper Functions ---
def get_google_creds():
    creds_json_str = os.environ.get('GOOGLE_CREDENTIALS_JSON')
    if creds_json_str:
        return Credentials.from_service_account_info(json.loads(creds_json_str), scopes=SCOPES)
    elif os.path.exists(GOOGLE_CREDENTIALS_FILE):
        return Credentials.from_service_account_file(GOOGLE_CREDENTIALS_FILE, scopes=SCOPES)
    raise FileNotFoundError("Google credentials not found.")

def is_valid_email(email):
    return re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email)

def get_sheet_and_records(sheet_name):
    creds = get_google_creds()
    client = gspread.authorize(creds)
    sh = client.open(GOOGLE_SHEET_NAME)
    worksheet = sh.worksheet(sheet_name)
    records = worksheet.get_all_records()
    return worksheet, records

# --- PDF Generation ---
def create_invoice_pdf(data):
    file_name = f"{data['invoice_num']}_{data['vendor_name'].replace(' ', '_')}.pdf"
    pdf_path = os.path.join('/tmp', file_name)
    
    # Create a custom page template with footer
    class FooterCanvas(canvas.Canvas):
        def __init__(self, *args, **kwargs):
            canvas.Canvas.__init__(self, *args, **kwargs)
            self.pages = []

        def showPage(self):
            self.pages.append(dict(self.__dict__))
            self._startPage()

        def save(self):
            page_count = len(self.pages)
            for page in self.pages:
                self.__dict__.update(page)
                self.draw_footer(page_count)
                canvas.Canvas.showPage(self)
            canvas.Canvas.save(self)

        def draw_footer(self, page_count):
            self.setFont("Helvetica", 8)
            self.setFillColor(colors.gray)
            footer_text = f"© {datetime.now().year} Malkit Sweets. All rights reserved."
            self.drawString(72, 30, footer_text)

    doc = SimpleDocTemplate(
        pdf_path,
        pagesize=letter,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name='RightAlign', alignment=TA_RIGHT))
    styles.add(ParagraphStyle(name='CenterAlign', alignment=TA_CENTER, fontSize=24, spaceAfter=10, leading=30))
    styles.add(ParagraphStyle(name='CompanyInfo', alignment=TA_RIGHT, leading=14))
    elements = []
    
    # Header
    header_data = [[Paragraph("MALKIT SWEETS AND CATERING", styles['h2']), Paragraph("INVOICE", styles['CenterAlign'])]]
    header_table = Table(header_data, colWidths=[2*inch, 4.5*inch])
    header_table.setStyle(TableStyle([('VALIGN', (0, 0), (-1, -1), 'MIDDLE'), ('ALIGN', (1, 0), (1, 0), 'RIGHT')]))
    elements.append(header_table)
    
    # Company Info
    company_info_data = [[Paragraph("<b>MALKIT SWEETS AND CATERING</b><br/>18110 Parthenia St<br/> Northridge, CA 91324", styles['Normal']), Paragraph(f"<b>Invoice #:</b> {data['invoice_num']}<br/><b>Date:</b> {datetime.strptime(data['date'], '%m/%d/%Y').strftime('%B %d, %Y')}", styles['CompanyInfo'])]]
    company_info_table = Table(company_info_data, colWidths=[3.5*inch, 3*inch])
    elements.append(company_info_table)
    elements.append(Spacer(1, 0.4*inch))
    
    # Bill To
    bill_to_address = f"{data.get('vendor_address', '')}<br/>{data.get('vendor_city', '')}"
    elements.append(Paragraph(f"<b>BILL TO:</b><br/>{data['vendor_name']}<br/>{data['vendor_email']}<br/>{data.get('vendor_phone', '')}<br/>{bill_to_address}", styles['Normal']))
    elements.append(Spacer(1, 0.4*inch))

    # Items Table
    table_data = [['Item Description', 'Quantity', 'Unit Price', 'Total']]
    grand_total = 0
    for item in data.get('items', []):
        try:
            quantity = float(item.get('quantity', 1))
            price = float(item.get('price', 0))
            total = quantity * price
            grand_total += total
            table_data.append([item.get('item', 'N/A'), f"{quantity}", f"${price:.2f}", f"${total:.2f}"])
        except (ValueError, TypeError):
            continue
    
    invoice_table = Table(table_data, colWidths=[3.5*inch, 1*inch, 1*inch, 1*inch])
    invoice_table.setStyle(TableStyle([('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#4F46E5')), ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke), ('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'), ('BOTTOMPADDING', (0, 0), (-1, 0), 12), ('BACKGROUND', (0, 1), (-1, -1), colors.beige), ('GRID', (0, 0), (-1, -1), 1, colors.black)]))
    elements.append(invoice_table)
    
    # Total
    total_data = [['', '', 'Grand Total:', f'${grand_total:.2f}']]
    total_table = Table(total_data, colWidths=[3.5*inch, 1*inch, 1*inch, 1*inch])
    total_table.setStyle(TableStyle([('ALIGN', (2, 0), (2, 0), 'RIGHT'), ('FONTNAME', (2, 0), (3, 0), 'Helvetica-Bold'), ('GRID', (2, 0), (3, 0), 1, colors.black), ('BACKGROUND', (2,0), (3,0), colors.lightgrey)]))
    elements.append(total_table)
    elements.append(Spacer(1, 0.5 * inch))

    # Notes
    if data.get('notes'):
        elements.append(Paragraph(f"<b>Notes:</b><br/>{data['notes']}", styles['Normal']))
    
    # Build PDF with footer
    doc.build(elements, canvasmaker=FooterCanvas)
    return pdf_path, file_name, grand_total

def upload_to_google_drive(file_path, file_name):
    try:
        creds = get_google_creds()
        service = build('drive', 'v3', credentials=creds)
        
        query = f"mimeType='application/vnd.google-apps.folder' and name='{GOOGLE_DRIVE_FOLDER_NAME}' and trashed=false"
        response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
        folder_id = response.get('files')[0].get('id') if response.get('files') else None
        
        if not folder_id:
            folder_metadata = {'name': GOOGLE_DRIVE_FOLDER_NAME, 'mimeType': 'application/vnd.google-apps.folder'}
            folder_id = service.files().create(body=folder_metadata, fields='id').execute().get('id')

        file_metadata = {'name': file_name, 'parents': [folder_id]}
        media = MediaFileUpload(file_path, mimetype='application/pdf')
        service.files().create(body=file_metadata, media_body=media, fields='id').execute()
    except Exception as e:
        print(f"Google Drive upload error: {e}")

def get_next_invoice_number():
    try:
        _, records = get_sheet_and_records(GOOGLE_INVOICES_SHEET_NAME)
        return f"INV-{(len(records) + 1):04d}"
    except Exception:
        return f"INV-TS-{int(datetime.now().timestamp())}"

def add_invoice_to_sheet(data, grand_total):
    try:
        worksheet, records = get_sheet_and_records(GOOGLE_INVOICES_SHEET_NAME)
        item_summary = ", ".join([f"{item.get('item', 'N/A')} (x{item.get('quantity', 0)})" for item in data.get('items', [])])
        new_row = [datetime.now().strftime('%Y-%m-%d %H:%M:%S'), data['invoice_num'], data['date'], data['vendor_name'], data['vendor_email'], f"${grand_total:.2f}", data.get('notes', ''), 'Due', item_summary]
        if not records:
            worksheet.append_row(["Timestamp", "Invoice #", "Invoice Date", "Vendor Name", "Vendor Email", "Total", "Notes", "Status", "Items"])
        worksheet.append_row(new_row)
    except Exception as e:
        print(f"Error adding invoice to sheet: {e}")

def send_email_with_attachment(recipient_email, subject, body, file_path):
    if not all([EMAIL_HOST_USER, EMAIL_HOST_PASSWORD, SENDER_EMAIL]):
        print("Error: SMTP settings not fully configured.")
        return False

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
    part.add_header('Content-Disposition', f"attachment; filename={os.path.basename(file_path)}")
    msg.attach(part)
    
    try:
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.starttls()
        server.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
        server.sendmail(SENDER_EMAIL, [recipient_email, SENDER_EMAIL], msg.as_string())
        return True
    except Exception as e:
        print(f"Error sending email: {e}")
        return False
    finally:
        try:
            server.quit()
        except:
            pass

# --- API Endpoints ---
@app.route('/get-vendors', methods=['GET'])
def get_vendors():
    try:
        _, records = get_sheet_and_records(GOOGLE_VENDORS_SHEET_NAME)
        vendors = []
        for i, record in enumerate(records):
            std_rec = {key.strip().lower().replace(' ', '_'): value for key, value in record.items()}
            std_rec['name'] = std_rec.get('name', '')
            std_rec['email'] = std_rec.get('email', '')
            std_rec['address'] = std_rec.get('address', '')
            std_rec['city'] = std_rec.get('city', '')
            std_rec['phone'] = std_rec.get('phone', std_rec.get('phone_number', ''))
            std_rec['row_number'] = i + 2
            vendors.append(std_rec)
        return jsonify(vendors), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/add-vendor', methods=['POST'])
def add_vendor():
    try:
        data = request.get_json()
        required = ['name', 'email', 'address', 'city', 'phone']
        if not all(data.get(k) for k in required) or not is_valid_email(data['email']):
            return jsonify({"error": "Invalid or missing vendor fields."}), 400
        worksheet, _ = get_sheet_and_records(GOOGLE_VENDORS_SHEET_NAME)
        new_row = [data['name'], data['email'], data['address'], data['city'], data['phone']]
        worksheet.append_row(new_row)
        return jsonify({"message": "Vendor added."}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/edit-vendor', methods=['POST'])
def edit_vendor():
    try:
        data = request.get_json()
        required = ['row_number', 'name', 'email', 'address', 'city', 'phone']
        if not all(data.get(k) for k in required) or not is_valid_email(data['email']):
            return jsonify({"error": "Invalid or missing fields."}), 400
        worksheet, _ = get_sheet_and_records(GOOGLE_VENDORS_SHEET_NAME)
        cell_range = f'A{data["row_number"]}:E{data["row_number"]}'
        cell_values = [[data['name'], data['email'], data['address'], data['city'], data['phone']]]
        worksheet.update(cell_range, cell_values)
        return jsonify({"message": "Vendor updated."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/delete-vendor', methods=['POST'])
def delete_vendor():
    try:
        worksheet, _ = get_sheet_and_records(GOOGLE_VENDORS_SHEET_NAME)
        worksheet.delete_rows(int(request.json['row_number']))
        return jsonify({"message": "Vendor deleted."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/get-invoices', methods=['GET'])
def get_invoices():
    try:
        _, records = get_sheet_and_records(GOOGLE_INVOICES_SHEET_NAME)
        invoices = []
        for i, record in enumerate(records):
            std_rec = {key.strip().lower().replace(' ', '_'): value for key, value in record.items()}
            
            # Handle items JSON
            if 'items' in std_rec and isinstance(std_rec['items'], str):
                try: 
                    std_rec['items'] = json.loads(std_rec['items'])
                except json.JSONDecodeError: 
                    std_rec['items'] = []
            
            # Format timestamp if it exists
            if 'timestamp' in std_rec and std_rec['timestamp']:
                try:
                    # Convert timestamp to datetime object
                    if isinstance(std_rec['timestamp'], str):
                        # Remove any timezone information if present
                        timestamp_str = std_rec['timestamp'].split('+')[0].strip()
                        # Parse the timestamp
                        timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                        # Format the time
                        std_rec['formatted_time'] = timestamp.strftime('%I:%M %p')
                    else:
                        std_rec['formatted_time'] = 'N/A'
                except Exception as e:
                    print(f"Error formatting timestamp: {e}")
                    std_rec['formatted_time'] = 'N/A'
            else:
                std_rec['formatted_time'] = 'N/A'
            
            # Map invoice number from "Invoice #" column
            if 'invoice_#' in std_rec:
                std_rec['invoice_num'] = std_rec['invoice_#']
            elif 'invoice_num' in std_rec:
                std_rec['invoice_num'] = std_rec['invoice_num']
            else:
                std_rec['invoice_num'] = 'N/A'
                
            std_rec['row_number'] = i + 2
            invoices.append(std_rec)
        return jsonify(invoices), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/update-status', methods=['POST'])
def update_status():
    try:
        data = request.json
        row_number = data.get('row_number')
        new_status = data.get('status', 'Due')  # Default to 'Due' if not provided
        
        if not row_number:
            return jsonify({"error": "Row number is required"}), 400
            
        sheet, records = get_sheet_and_records(GOOGLE_INVOICES_SHEET_NAME)
        
        # Find the status column
        status_col = None
        for col in range(1, sheet.col_count + 1):
            if sheet.cell(1, col).value == 'Status':
                status_col = col
                break
                
        if not status_col:
            return jsonify({"error": "Status column not found"}), 500
            
        # Update the status
        sheet.update_cell(row_number, status_col, new_status)
        
        return jsonify({"message": f"Status updated to {new_status}"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/get-drafts', methods=['GET'])
def get_drafts():
    try:
        _, records = get_sheet_and_records(GOOGLE_DRAFTS_SHEET_NAME)
        drafts = []
        for i, record in enumerate(records):
            std_rec = {key.strip().lower().replace(' ', '_'): value for key, value in record.items()}
            if 'items' in std_rec and isinstance(std_rec['items'], str):
                try: 
                    std_rec['items'] = json.loads(std_rec['items'])
                except json.JSONDecodeError: 
                    std_rec['items'] = []
            
            # Format timestamp if it exists
            if 'timestamp' in std_rec and std_rec['timestamp']:
                try:
                    # Convert timestamp to datetime object
                    if isinstance(std_rec['timestamp'], str):
                        # Remove any timezone information if present
                        timestamp_str = std_rec['timestamp'].split('+')[0].strip()
                        # Parse the timestamp
                        timestamp = datetime.strptime(timestamp_str, '%Y-%m-%d %H:%M:%S')
                        # Format the time
                        std_rec['formatted_time'] = timestamp.strftime('%I:%M %p')
                    else:
                        std_rec['formatted_time'] = 'N/A'
                except Exception as e:
                    print(f"Error formatting timestamp: {e}")
                    std_rec['formatted_time'] = 'N/A'
            else:
                std_rec['formatted_time'] = 'N/A'
                
            std_rec['row_number'] = i + 2
            drafts.append(std_rec)
        return jsonify(drafts), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/get-draft/<int:row_number>', methods=['GET'])
def get_draft(row_number):
    try:
        worksheet, _ = get_sheet_and_records(GOOGLE_DRAFTS_SHEET_NAME)
        draft_data = worksheet.row_values(row_number)
        if not draft_data:
            return jsonify({"error": "Draft not found"}), 404
            
        headers = worksheet.row_values(1)
        draft = {}
        for i, header in enumerate(headers):
            key = header.strip().lower().replace(' ', '_')
            value = draft_data[i] if i < len(draft_data) else ''
            if key == 'items' and value:
                try:
                    draft[key] = json.loads(value)
                except json.JSONDecodeError:
                    draft[key] = []
            else:
                draft[key] = value
                
        draft['row_number'] = row_number
        return jsonify(draft), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/save-draft', methods=['POST'])
def save_draft():
    try:
        data = request.json
        worksheet, records = get_sheet_and_records(GOOGLE_DRAFTS_SHEET_NAME)
        
        # Add timestamp to the data
        data['timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Convert items to JSON string
        items_json = json.dumps(data.get('items', []))
        
        # Prepare data in the correct column order
        draft_data = [
            data.get('vendor_name', ''),
            data.get('vendor_email', ''),
            data.get('date', ''),
            data.get('notes', ''),
            items_json,
            data.get('vendor_address', ''),
            data.get('vendor_city', ''),
            data.get('vendor_phone', ''),
            data.get('timestamp', '')
        ]
        
        # If row_number exists, update existing draft
        if 'row_number' in data and data['row_number']:
            worksheet.update(f'A{data["row_number"]}:I{data["row_number"]}', [draft_data])
            return jsonify({"message": "Draft updated successfully"}), 200
        else:
            # Add new draft
            if not records:
                # Add headers if sheet is empty
                worksheet.append_row(["Vendor Name", "Vendor Email", "Date", "Notes", "Items", "Address", "City", "Phone", "Timestamp"])
            worksheet.append_row(draft_data)
            return jsonify({"message": "Draft saved successfully"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/delete-draft', methods=['POST'])
def delete_draft():
    try:
        worksheet, _ = get_sheet_and_records(GOOGLE_DRAFTS_SHEET_NAME)
        worksheet.delete_rows(int(request.json['row_number']))
        return jsonify({"message": "Draft deleted."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/generate-invoice', methods=['POST'])
def generate_invoice():
    try:
        data = request.get_json()
        if not data:
            return jsonify({"error": "No data provided"}), 400
            
        required_fields = ['vendor_name', 'vendor_email', 'vendor_address', 'vendor_city', 'vendor_phone', 'date', 'items']
        missing_fields = [field for field in required_fields if not data.get(field)]
        if missing_fields:
            return jsonify({"error": f"Missing required fields: {', '.join(missing_fields)}"}), 400
        
        if not data.get('items'):
            return jsonify({"error": "No items provided"}), 400
            
        invoice_num = get_next_invoice_number()
        data['invoice_num'] = invoice_num
        
        try:
            pdf_path, pdf_filename, grand_total = create_invoice_pdf(data)
        except Exception as e:
            print(f"Error creating PDF: {str(e)}")
            return jsonify({"error": f"Failed to create PDF: {str(e)}"}), 500
        
        subject = f"Invoice {invoice_num} from MALKIT SWEETS AND CATERING"
        body = f"Hello {data['vendor_name']},\n\nPlease find attached Invoice {invoice_num}.\n\nThank you,\nMALKIT SWEETS AND CATERING"
        
        if not send_email_with_attachment(data['vendor_email'], subject, body, pdf_path):
            return jsonify({"error": "Failed to send email"}), 500
        
        try:
            add_invoice_to_sheet(data, grand_total)
            upload_to_google_drive(pdf_path, pdf_filename)
        except Exception as e:
            print(f"Error updating records: {str(e)}")
            return jsonify({"error": f"Failed to update records: {str(e)}"}), 500
        
        if 'row_number' in data and data['row_number']:
            try:
                worksheet, _ = get_sheet_and_records(GOOGLE_DRAFTS_SHEET_NAME)
                worksheet.delete_rows(int(data['row_number']))
            except Exception as e:
                print(f"Error deleting draft: {str(e)}")
        
        if os.path.exists(pdf_path):
            os.remove(pdf_path)

        return jsonify({"message": f"Invoice {invoice_num} generated and sent."}), 200
    except Exception as e:
        print(f"Unexpected error in generate_invoice: {str(e)}")
        return jsonify({"error": str(e)}), 500

@app.route('/download-draft-preview', methods=['POST'])
def download_draft_preview():
    try:
        data = request.get_json()
        if not data: return jsonify({"error": "No data provided"}), 400
        data['invoice_num'] = f"DRAFT-{data.get('row_number', 'PREVIEW')}"
        pdf_path, pdf_filename, _ = create_invoice_pdf(data)
        return send_file(pdf_path, as_attachment=True, download_name=pdf_filename)
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/get-sweets', methods=['GET'])
def get_sweets():
    try:
        _, records = get_sheet_and_records(GOOGLE_SWEETS_SHEET_NAME)
        sweets = []
        for i, record in enumerate(records):
            std_rec = {key.strip().lower().replace(' ', '_'): value for key, value in record.items()}
            # Convert price to float
            try:
                std_rec['price'] = float(std_rec.get('price', 0))
            except (ValueError, TypeError):
                std_rec['price'] = 0.0
            std_rec['row_number'] = i + 2
            sweets.append(std_rec)
        return jsonify(sweets), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/add-sweet', methods=['POST'])
def add_sweet():
    try:
        data = request.get_json()
        required = ['name', 'price']
        if not all(data.get(k) for k in required):
            return jsonify({"error": "Invalid or missing fields."}), 400
        
        # Validate price is a number
        try:
            price = float(data['price'])
        except (ValueError, TypeError):
            return jsonify({"error": "Price must be a valid number."}), 400
        
        worksheet, _ = get_sheet_and_records(GOOGLE_SWEETS_SHEET_NAME)
        new_row = [
            data['name'],
            str(price)
        ]
        worksheet.append_row(new_row)
        return jsonify({"message": "Sweet added."}), 201
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/edit-sweet', methods=['POST'])
def edit_sweet():
    try:
        data = request.get_json()
        required = ['row_number', 'name', 'price']
        if not all(data.get(k) for k in required):
            return jsonify({"error": "Invalid or missing fields."}), 400
        
        # Validate price is a number
        try:
            price = float(data['price'])
        except (ValueError, TypeError):
            return jsonify({"error": "Price must be a valid number."}), 400
        
        worksheet, _ = get_sheet_and_records(GOOGLE_SWEETS_SHEET_NAME)
        cell_range = f'A{data["row_number"]}:B{data["row_number"]}'
        cell_values = [[
            data['name'],
            str(price)
        ]]
        worksheet.update(cell_range, cell_values)
        return jsonify({"message": "Sweet updated."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/delete-sweet', methods=['POST'])
def delete_sweet():
    try:
        worksheet, _ = get_sheet_and_records(GOOGLE_SWEETS_SHEET_NAME)
        worksheet.delete_rows(int(request.json['row_number']))
        return jsonify({"message": "Sweet deleted."}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/test-email-config', methods=['GET'])
def test_email_config():
    try:
        # Create a test message
        msg = MIMEMultipart()
        msg['From'] = SENDER_EMAIL
        msg['To'] = SENDER_EMAIL  # Send to self for testing
        msg['Subject'] = "Test Email from Invoice Generator"
        msg.attach(MIMEText("This is a test email to verify the configuration.", 'plain'))
        
        # Try to send the email
        server = smtplib.SMTP(EMAIL_HOST, EMAIL_PORT)
        server.set_debuglevel(1)
        server.ehlo()
        server.starttls()
        server.ehlo()
        server.login(EMAIL_HOST_USER, EMAIL_HOST_PASSWORD)
        server.sendmail(SENDER_EMAIL, [SENDER_EMAIL], msg.as_string())
        server.quit()
        
        return jsonify({"message": "Test email sent successfully!"}), 200
    except Exception as e:
        return jsonify({"error": str(e)}), 500

# --- Startup Check & Run ---
def check_env_vars():
    required_vars = ['EMAIL_HOST_USER', 'EMAIL_HOST_PASSWORD', 'SENDER_EMAIL', 'GOOGLE_SHEET_NAME', 'GOOGLE_DRIVE_FOLDER_NAME']
    missing_vars = [v for v in required_vars if not os.getenv(v)]
    if missing_vars:
        raise EnvironmentError(f"Startup failed: Missing env vars: {', '.join(missing_vars)}")
    
    get_google_creds()

if __name__ == '__main__':
    from dotenv import load_dotenv
    load_dotenv()
    try:
        check_env_vars()
        app.run(debug=True, port=5001)
    except EnvironmentError as e:
        print(e)
