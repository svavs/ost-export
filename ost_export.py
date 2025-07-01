import os
import sys
import logging
import mimetypes
import re
import email
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from email.utils import formatdate
from datetime import datetime
from bs4 import BeautifulSoup
import pypff
import traceback
from typing import Optional

# Define pypff constants if they don't exist
if not hasattr(pypff, 'file'):
    class PffFile:
        ACCESS_READ = 0x01
    pypff.file = PffFile()

if not hasattr(pypff, 'error'):
    class PffError(Exception):
        pass
    pypff.error = PffError

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

def create_message(msg_obj, folder_name, output_format):
    """Create an email message from a pypff message object with proper MIME structure."""
    try:
        # Create a multipart message that will contain the email
        msg = MIMEMultipart('mixed')
        
        # Create a related part for the main content (can contain both text and HTML)
        msg_related = MIMEMultipart('related')
        
        # Add basic headers
        if hasattr(msg_obj, 'subject') and msg_obj.subject:
            msg['Subject'] = str(msg_obj.subject)
        if hasattr(msg_obj, 'sender_name') and msg_obj.sender_name:
            msg['From'] = str(msg_obj.sender_name)
        if hasattr(msg_obj, 'display_to') and msg_obj.display_to:
            msg['To'] = str(msg_obj.display_to)
        if hasattr(msg_obj, 'display_cc') and msg_obj.display_cc:
            msg['Cc'] = str(msg_obj.display_cc)
        if hasattr(msg_obj, 'delivery_time') and msg_obj.delivery_time:
            msg['Date'] = msg_obj.delivery_time.strftime('%a, %d %b %Y %H:%M:%S %z')
        else:
            msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S %z')
        
        # Ensure proper MIME version
        msg['MIME-Version'] = '1.0'
        
        # Add message body
        try:
            # Try HTML body first
            if hasattr(msg_obj, 'html_body') and msg_obj.html_body:
                try:
                    body = msg_obj.html_body
                    if isinstance(body, bytes):
                        body = body.decode('utf-8', errors='replace')
                    body_type = 'html'
                except Exception as e:
                    logging.warning(f"Error processing HTML body: {e}")
                    # Fall through to plain text
            
            # Try plain text body
            if not hasattr(msg_obj, 'html_body') or not msg_obj.html_body:
                if hasattr(msg_obj, 'plain_text_body') and msg_obj.plain_text_body:
                    try:
                        body = msg_obj.plain_text_body
                        if isinstance(body, bytes):
                            body = body.decode('utf-8', errors='replace')
                        body_type = 'plain'
                    except Exception as e:
                        logging.warning(f"Error processing plain text body: {e}")
            
            # If no body was added, add a placeholder
            if not hasattr(msg_obj, 'html_body') and not hasattr(msg_obj, 'plain_text_body'):
                body = 'No message body found.'
                body_type = 'plain'
            
            # Add the body to the message as an alternative part
            alternative = MIMEMultipart('alternative')
            
            # Add plain text part
            if body_type == 'html':
                # Try to extract plain text from HTML
                try:
                    soup = BeautifulSoup(body, 'html.parser')
                    plain_text = soup.get_text('\n')
                    alternative.attach(MIMEText(plain_text, 'plain', 'utf-8'))
                except:
                    # Fallback to HTML only if conversion fails
                    pass
                alternative.attach(MIMEText(body, 'html', 'utf-8'))
            else:
                alternative.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # Add the alternative part to the related part
            msg_related.attach(alternative)
            
            # Add the related part to the main message
            msg.attach(msg_related)
        
        except Exception as e:
            logging.error(f"Error processing message body: {e}")
            msg.attach(MIMEText('Error processing message body.', 'plain', 'utf-8'))

        # Process attachments
        try:
            if hasattr(msg_obj, 'attachments') and msg_obj.attachments:
                for attachment in msg_obj.attachments:
                    try:
                        # Skip if no data
                        if not hasattr(attachment, 'read_buffer') or not hasattr(attachment, 'size'):
                            continue
                            
                        # Read attachment data
                        attach_data = attachment.read_buffer(attachment.size)
                        if not attach_data:
                            continue
                            
                        # Create MIME part
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attach_data)
                        encoders.encode_base64(part)
                        
                        # Get filename and clean it
                        filename = None
                        if hasattr(attachment, 'name') and attachment.name:
                            filename = attachment.name
                        elif hasattr(attachment, 'get_name'):
                            filename = attachment.get_name()
                        
                        # If no filename, create a generic one
                        if not filename:
                            filename = f"attachment_{getattr(attachment, 'identifier', 'unknown')}"
                        
                        # Clean filename and ensure it's a string
                        filename = str(filename).strip()
                        filename = "".join(c for c in filename if c.isprintable() and c not in '\\/*?:"<>|')
                        
                        # Common MIME types mapping
                        mime_types = {
                            # Documents
                            'pdf': 'application/pdf',
                            'doc': 'application/msword',
                            'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                            'xls': 'application/vnd.ms-excel',
                            'xlsx': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            'ppt': 'application/vnd.ms-powerpoint',
                            'pptx': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
                            'txt': 'text/plain',
                            'rtf': 'application/rtf',
                            'csv': 'text/csv',
                            # Images
                            'png': 'image/png',
                            'jpg': 'image/jpeg',
                            'jpeg': 'image/jpeg',
                            'gif': 'image/gif',
                            'bmp': 'image/bmp',
                            'tiff': 'image/tiff',
                            'svg': 'image/svg+xml',
                            # Archives
                            'zip': 'application/zip',
                            'rar': 'application/x-rar-compressed',
                            '7z': 'application/x-7z-compressed',
                            'tar': 'application/x-tar',
                            'gz': 'application/gzip',
                            # Audio/Video
                            'mp3': 'audio/mpeg',
                            'wav': 'audio/wav',
                            'mp4': 'video/mp4',
                            'avi': 'video/x-msvideo',
                            'mov': 'video/quicktime',
                            'wmv': 'video/x-ms-wmv'
                        }
                        
                        # Get file extension
                        file_ext = ''
                        if '.' in filename:
                            file_ext = filename.rsplit('.', 1)[1].lower()
                        
                        # Determine content type
                        content_type = mime_types.get(file_ext, 'application/octet-stream')
                        
                        # Special handling for known binary formats
                        if file_ext in ['pdf', 'doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx', 'zip', 'rar', '7z']:
                            if not filename.lower().endswith(f'.{file_ext}'):
                                filename = f"{filename}.{file_ext}"
                        
                        # Debug: Print attachment info
                        logging.info(f"Processing attachment: {filename}, Type: {content_type}, Size: {len(attach_data)} bytes")
                        
                        # Special handling for PDFs - check magic number
                        if filename.lower().endswith('.pdf') or (len(attach_data) > 4 and attach_data.startswith(b'%PDF-')):
                            content_type = 'application/pdf'
                            if not filename.lower().endswith('.pdf'):
                                filename += '.pdf'
                            logging.info(f"Detected PDF file by content or extension: {filename}")
                        
                        # Create the MIME part with proper headers
                        maintype, subtype = content_type.split('/', 1) if '/' in content_type else ('application', 'octet-stream')
                        
                        # Additional debug for PDFs
                        if content_type == 'application/pdf':
                            logging.info(f"PDF content starts with: {attach_data[:100]}")
                            # Try to extract PDF version number
                            pdf_header = attach_data[:8].decode('ascii', errors='ignore')
                            logging.info(f"PDF header: {pdf_header}")
                        
                        if maintype == 'text':
                            # For text files, use MIMEText to handle encoding
                            try:
                                if isinstance(attach_data, bytes):
                                    attach_data = attach_data.decode('utf-8', errors='replace')
                                part = MIMEText(attach_data, _subtype=subtype, _charset='utf-8')
                            except Exception as e:
                                logging.warning(f"Error creating text part: {e}")
                                part = MIMEBase(maintype, subtype)
                                part.set_payload(attach_data)
                        else:
                            # For binary files, use MIMEBase
                            part = MIMEBase(maintype, subtype)
                            part.set_payload(attach_data)
                            encoders.encode_base64(part)
                        
                        # Add headers with explicit content type
                        if content_type == 'application/pdf':
                            # For PDFs, be very explicit with headers
                            part.set_type('application/pdf')
                            part.add_header('Content-Type', 'application/pdf', name=filename)
                            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
                            part.add_header('Content-Transfer-Encoding', 'base64')
                            part.add_header('Content-Description', 'PDF Document')
                            logging.info(f"Set explicit PDF headers for: {filename}")
                        else:
                            # For other file types
                            part.add_header('Content-Type', f'{maintype}/{subtype}', name=filename)
                            part.add_header('Content-Disposition', 'attachment', filename=filename)
                            part.add_header('Content-Transfer-Encoding', 'base64' if maintype != 'text' else '8bit')
                        
                            # For known binary formats, ensure proper content type
                            if file_ext in ['doc', 'docx', 'xls', 'xlsx', 'ppt', 'pptx']:
                                part.set_type(content_type)
                                part.add_header('Content-Type', content_type, name=filename)
                                logging.info(f"Set explicit content type for {file_ext}: {content_type}")
                        msg.attach(part)
                        
                    except Exception as e:
                        logging.error(f"Error processing attachment: {e}")
                        continue
                        
        except Exception as e:
            logging.error(f"Error processing attachments: {e}")

        return msg

    except Exception as e:
        logging.error(f"Error creating message: {e}")
        logging.debug(traceback.format_exc())
        # Return a minimal valid message even if there was an error
        msg = MIMEMultipart()
        msg['From'] = "error@local"
        msg['To'] = "unknown@local"
        msg['Subject'] = "Error processing message"
        msg.attach(MIMEText(f'Error processing message: {str(e)}', 'plain', 'utf-8'))
        return msg

def export_to_mbox(messages, output_file):
    """Export messages to an MBOX file with proper formatting."""
    # Ensure the output directory exists
    os.makedirs(os.path.dirname(os.path.abspath(output_file)), exist_ok=True)
    
    with open(output_file, 'ab') as mbox_file:
        for msg in messages:
            try:
                # Ensure proper MIME structure
                if not msg.is_multipart():
                    # Convert to multipart if not already
                    new_msg = MIMEMultipart()
                    for key, value in msg.items():
                        new_msg[key] = value
                    new_msg.attach(MIMEText(msg.get_payload(), 'plain' if msg.get_content_type() == 'text/plain' else 'html'))
                    msg = new_msg
                
                # Ensure proper headers for MBOX format
                if 'From' not in msg:
                    msg['From'] = 'unknown@example.com'
                if 'Date' not in msg:
                    msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S %z')
                if 'Message-ID' not in msg:
                    msg['Message-ID'] = f"<{datetime.now().timestamp()}@{os.uname().nodename}>"
                
                # Add From_ line required by mbox format
                from_line = f"From {msg['From']} {datetime.now().strftime('%a %b %d %H:%M:%S %Y')}\n"
                mbox_file.write(from_line.encode('utf-8', errors='replace'))
                
                # Write the message with proper line endings
                msg_str = msg.as_string()
                # Ensure proper line endings for MBOX format
                msg_str = msg_str.replace('\r\n', '\n').replace('\r', '\n').replace('\n', '\r\n')
                mbox_file.write(msg_str.encode('utf-8', errors='replace'))
                mbox_file.write(b'\r\n\r\n')  # Add separator between messages with proper line endings
                
            except Exception as e:
                logging.error(f"Error exporting message to MBOX: {e}")
                continue

def process_folder(folder, output_dir: str, format: str = 'mbox'):
    """Process a folder and its subfolders."""
    # Initialize folder_name with a default value
    folder_name = "UnknownFolder"
    safe_folder_name = folder_name
    
    try:
        # Get folder name if available
        if hasattr(folder, 'get_name'):
            folder_name = folder.get_name() or folder_name
        elif hasattr(folder, 'name'):
            folder_name = getattr(folder, 'name', folder_name)
        
        # Create safe folder name
        safe_folder_name = str(folder_name).replace('/', '_').replace('\\', '_')
        logging.info(f"Processing folder: {folder_name}")
        
        # Process subfolders
        try:
            if hasattr(folder, 'sub_folders'):
                for subfolder in folder.sub_folders:
                    process_folder(subfolder, output_dir, format)
            elif hasattr(folder, 'get_sub_folders'):
                for subfolder in folder.get_sub_folders():
                    process_folder(subfolder, output_dir, format)
        except Exception as e:
            logging.warning(f"Could not process subfolders in {folder_name}: {e}")

        # Process messages
        if format == 'mbox':
            mbox_path = os.path.join(output_dir, f"{safe_folder_name}.mbox")
            try:
                messages = []
                # Different ways to get messages based on pypff version
                if hasattr(folder, 'sub_messages'):
                    messages = folder.sub_messages
                elif hasattr(folder, 'get_number_of_sub_messages'):
                    messages = [folder.get_sub_message(i) for i in range(folder.get_number_of_sub_messages())]
                else:
                    messages = []
                
                for message in messages:
                    try:
                        if not message:
                            continue
                            
                        msg = create_message(message, folder_name, format)
                        export_to_mbox([msg], mbox_path)
                    except Exception as e:
                        logging.error(f"Error processing message in folder '{folder_name}': {e}")
                        continue
            except Exception as e:
                logging.error(f"Error writing to mbox file {mbox_path}: {e}")
        
        elif format == 'eml':
            folder_path = os.path.join(output_dir, safe_folder_name)
            os.makedirs(folder_path, exist_ok=True)
            
            try:
                # Different ways to get messages based on pypff version
                if hasattr(folder, 'sub_messages'):
                    messages = folder.sub_messages
                elif hasattr(folder, 'get_number_of_sub_messages'):
                    messages = [folder.get_sub_message(i) for i in range(folder.get_number_of_sub_messages())]
                else:
                    messages = []
                
                for message in messages:
                    try:
                        if not message:
                            continue
                            
                        msg = create_message(message, folder_name, format)
                        
                        # Create safe filename
                        subject = getattr(message, 'subject', f"message_{getattr(message, 'entry_id', 'unknown')}")
                        safe_subject = "".join(c for c in str(subject) if c.isalnum() or c in (' ', '.', '_')).replace(' ', '_')
                        
                        eml_path = os.path.join(folder_path, f"{safe_subject}.eml")
                        with open(eml_path, "w", encoding='utf-8') as eml_file:
                            eml_file.write(msg.as_string())
                            
                    except Exception as e:
                        logging.error(f"Error processing message in folder '{folder_name}': {e}")
                        continue
            except Exception as e:
                logging.error(f"Error processing messages in folder '{folder_name}': {e}")

    except Exception as e:
        logging.error(f"Error processing folder '{folder_name}': {e}")
        logging.debug(traceback.format_exc())
        raise

def export_ost(ost_path: str, output_dir: str, format: str = 'mbox'):
    """Export OST file to either MBOX or EML format."""
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    try:
        # Open OST file
        ost_file = pypff.file()
        try:
            ost_file.open(ost_path)
            
            # Get root folders
            root_folders = []
            try:
                root_folders = ost_file.get_root_folder().sub_folders
            except AttributeError:
                # Try alternative method for getting root folders
                root_folders = ost_file.get_root_folders()
            
            if not root_folders:
                logging.error("No root folders found in the OST file")
                return
                
            # Process each root folder
            for root_folder in root_folders:
                process_folder(root_folder, output_dir, format)
                
        except Exception as e:
            logging.error(f"Error processing OST file: {e}")
            logging.debug(traceback.format_exc())
            raise
            
        finally:
            if hasattr(ost_file, 'close'):
                ost_file.close()
        logging.info(f"Conversion completed successfully! Files saved in: {output_dir}")
        
    except Exception as e:
        if hasattr(e, '__module__') and e.__module__.startswith('pypff'):
            logging.error(f"Error processing OST file (pypff error): {e}")
        else:
            logging.error(f"Unexpected error: {e}")
        logging.debug(traceback.format_exc())
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python ost_export.py <ost_file> <output_directory> <format>")
        print("Format can be either 'mbox' or 'eml'")
        sys.exit(1)

    ost_path = sys.argv[1]
    output_dir = sys.argv[2]
    format = sys.argv[3].lower()
    
    if format not in ['mbox', 'eml']:
        print("Format must be either 'mbox' or 'eml'")
        sys.exit(1)
    
    export_ost(ost_path, output_dir, format)
