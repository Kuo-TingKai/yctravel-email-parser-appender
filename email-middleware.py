import imaplib
import email
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
from bs4 import BeautifulSoup
import re
from datetime import datetime, timedelta

class EmailProcessor:
    def __init__(self):
        # Gmail IMAP settings, 這邊原來是由woocommerce跟billionconnect api發送
        # 我是為了測試所以改成自己發送
        self.gmail_imap = "imap.gmail.com"
        self.gmail_smtp = "smtp.gmail.com"
        self.gmail_user = "nehsm30126@gmail.com"
        self.gmail_password = "oymwnycrwfmpgklu"  # Replace with your 16-digit app password
        
        # YC Travel SMTP settings (中繼信箱)
        self.yc_smtp = "smtp.hostinger.com"
        self.yc_user = "service@yctravel.shop"
        self.yc_password = "Cloudservice2024!"
        
        # Target email （客戶信箱）
        self.target_email = "kaicocat@proton.me"
        self.search_string = "89812003919115380009"

        # Add QR code parser functions
        self.qr_parser = QRCodeParser()

    def fetch_matching_emails(self):
        # Connect to Gmail
        imap = imaplib.IMAP4_SSL(self.gmail_imap)
        imap.login(self.gmail_user, self.gmail_password)
        imap.select('INBOX')

        # Search for emails within the last 24 hours
        date = (datetime.now() - timedelta(days=1)).strftime("%d-%b-%Y")
        search_criteria = f'(TO "{self.target_email}" SINCE "{date}")'
        _, message_numbers = imap.search(None, search_criteria)
        
        matching_emails = []
        
        for num in message_numbers[0].split():
            try:
                _, msg_data = imap.fetch(num, '(RFC822)')
                email_body = msg_data[0][1]
                message = email.message_from_bytes(email_body)
                
                # Check if email is HTML and contains the search string
                if self._is_html_with_string(message, self.search_string):
                    matching_emails.append(message)
            except Exception as e:
                print(f"Error processing email {num}: {str(e)}")
                continue
        
        imap.logout()
        return matching_emails

    def _is_html_with_string(self, message, search_string):
        try:
            if message.is_multipart():
                for part in message.walk():
                    if part.get_content_type() == "text/html":
                        content = part.get_payload(decode=True).decode()
                        # Remove HTML tags for string search
                        text_content = BeautifulSoup(content, 'html.parser').get_text()
                        if search_string in text_content:
                            return True
            return False
        except Exception as e:
            print(f"Error checking HTML content: {str(e)}")
            return False

    def forward_to_yc(self, emails):
        # Combine all emails into one
        combined_html = self._combine_emails(emails)
        
        # Create message
        msg = MIMEMultipart()
        msg['From'] = self.yc_user
        msg['To'] = self.yc_user
        msg['Subject'] = f'Combined Forwarded Emails - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg.attach(MIMEText(combined_html, 'html'))

        # Send to YC
        with smtplib.SMTP(self.yc_smtp, 587) as server:
            server.starttls()
            server.login(self.yc_user, self.yc_password)
            server.send_message(msg)

    def forward_to_final(self, combined_content):
        # Create message
        msg = MIMEMultipart()
        msg['From'] = self.gmail_user
        msg['To'] = self.target_email
        msg['Subject'] = f'Processed Combined Emails - {datetime.now().strftime("%Y-%m-%d %H:%M")}'
        msg.attach(MIMEText(combined_content, 'html'))

        # Send to final recipient
        with smtplib.SMTP(self.gmail_smtp, 587) as server:
            server.starttls()
            server.login(self.gmail_user, self.gmail_password)
            server.send_message(msg)

    def _combine_emails(self, emails):
        combined_html = "<html><body>"
        
        # Process each email through QR code parser first
        for i, email_msg in enumerate(emails, 1):
            combined_html += f"<h2>Email #{i}</h2>"
            for part in email_msg.walk():
                if part.get_content_type() == "text/html":
                    html_content = part.get_payload(decode=True).decode()
                    
                    # Parse QR code and other info
                    customer_email, iccids, qr_code_url = self.qr_parser.parse_qr_code(html_content)
                    
                    # If this is a BillionConnect email (contains QR code)
                    if qr_code_url:
                        # Add the parsed eSIM info section
                        esim_section = self.qr_parser.append_to_order_processing(iccids, qr_code_url)
                        # Insert the eSIM section after the first margin-bottom div
                        first_margin = html_content.find('<div style="margin-bottom:20px">')
                        if first_margin != -1:
                            second_margin = html_content.find('<div style="margin-bottom:20px">', first_margin + 1)
                            if second_margin != -1:
                                html_content = (
                                    html_content[:second_margin] + 
                                    '\n<!-- Added by Email Processor -->\n' + 
                                    esim_section + 
                                    html_content[second_margin:]
                                )
                    
                    combined_html += html_content
                    
        combined_html += "</body></html>"
        return combined_html

    def process(self):
        try:
            # 1. Fetch matching emails
            matching_emails = self.fetch_matching_emails()
            if not matching_emails:
                print("No matching emails found")
                return

            # 2. Forward to YC
            self.forward_to_yc(matching_emails)

            # 3. Forward combined content to final recipient
            combined_content = self._combine_emails(matching_emails)
            self.forward_to_final(combined_content)

            print("Email processing completed successfully")
        except Exception as e:
            print(f"Error occurred: {str(e)}")

    def test_with_sample_files(self):
        """Test the email processing with sample HTML files"""
        try:
            # Create sample email messages
            sample_emails = []
            
            # Read and create email from billionconnect.html
            with open('billionconnect.html', 'r', encoding='utf-8') as file:
                billion_html = file.read()
                msg1 = MIMEMultipart()
                msg1['From'] = self.gmail_user
                msg1['To'] = self.target_email
                msg1['Subject'] = 'BillionConnect eSIM Activation'
                msg1.attach(MIMEText(billion_html, 'html'))
                sample_emails.append(msg1)
            
            # Read and create email from order_processing.html
            with open('order_processing.html', 'r', encoding='utf-8') as file:
                order_html = file.read()
                msg2 = MIMEMultipart()
                msg2['From'] = self.gmail_user
                msg2['To'] = self.target_email
                msg2['Subject'] = 'Order Processing Details'
                msg2.attach(MIMEText(order_html, 'html'))
                sample_emails.append(msg2)
            
            # Process the sample emails
            print("開始處理測試郵件...")
            
            # Forward to YC
            print("轉寄至 YC Travel...")
            self.forward_to_yc(sample_emails)
            
            # Forward combined content to final recipient
            print("合併後轉寄至最終收件者...")
            combined_content = self._combine_emails(sample_emails)
            self.forward_to_final(combined_content)
            
            print("測試郵件處理完成！")
            
        except Exception as e:
            print(f"測試過程中發生錯誤: {str(e)}")

class QRCodeParser:
    def parse_qr_code(self, input_html):
        """Parse QR code image, customer email and ICCIDs from email HTML content"""
        soup = BeautifulSoup(input_html, 'html.parser')
        
        # Find QR code image
        qr_code_img = soup.find('img', {'src': re.compile(r'.*op-flow-public.*qrcode.*\.png')})
        qr_code_url = ''
        
        if qr_code_img and 'src' in qr_code_img.attrs:
            src = qr_code_img['src']
            match = re.search(r'#(https://.*\.png)', src)
            if match:
                qr_code_url = match.group(1)
        
        # Get customer email
        customer_email = ''
        dear_customer = soup.find('p', string=re.compile(r'Dear Customer'))
        if dear_customer:
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', str(dear_customer.parent))
            if email_match:
                customer_email = email_match.group(0)
        
        # Get ICCIDs
        iccids = []
        iccid_text = soup.find('span', {'class': 'm_-6454673849054811020rColor'})
        if iccid_text:
            iccid = re.search(r'\d{19,20}', iccid_text.text)
            if iccid:
                iccids.append(iccid.group(0))
        
        # Try alternative format
        if not customer_email:
            customer_email_div = soup.find('div', {'id': 'customer_email'})
            if customer_email_div:
                customer_email = customer_email_div.text
        
        if not iccids:
            iccid_divs = soup.find_all('div', {'class': 'iccid'})
            for div in iccid_divs:
                iccids.append(div.text)
        
        return customer_email, iccids, qr_code_url

    def append_to_order_processing(self, iccids, qr_code_url):
        """Create order processing HTML content with the parsed information"""
        html_template = f"""
<div style="margin-bottom:20px">
    <table cellspacing="0" cellpadding="6" border="1" style="color:#333;border:1px solid #e5e5e5;vertical-align:middle;width:100%;font-family:'Helvetica Neue',Helvetica,Roboto,Arial,sans-serif" width="100%">
        <thead>
            <tr style="color:white;background-color:black" bgcolor="black">
                <th colspan="2" style="border:.5px solid #000;padding:12px">eSIM 資訊</th>
            </tr>
        </thead>
        <tbody>
            <tr>
                <td width="50%" style="color:#333;border:.5px solid #000;padding:12px;text-align:left;vertical-align:middle">
                    <p><strong>ICCIDs:</strong></p>
                    <ul style="margin:0;padding-left:20px">
                        {''.join([f'<li>{iccid}</li>' for iccid in iccids])}
                    </ul>
                </td>
                <td width="50%" style="color:#333;border:.5px solid #000;padding:12px;text-align:center;vertical-align:middle">
                    <p><strong>QR Code:</strong></p>
                    <img src="{qr_code_url}" alt="eSIM QR Code" style="max-width:200px;height:auto">
                </td>
            </tr>
        </tbody>
    </table>
</div>
"""
        return html_template

# Usage
if __name__ == "__main__":
    processor = EmailProcessor()
    
    # 執行測試
    print("執行測試範例...")
    processor.test_with_sample_files()
    
    # 如果要執行實際的郵件處理，取消下面的註解
    # print("執行實際郵件處理...")
    # processor.process()