from bs4 import BeautifulSoup
import re

def parse_qr_code(input_html):
    """
    Parse QR code image, customer email and ICCIDs from billionconnect.html
    
    Args:
        input_html (str): HTML content from billionconnect.html
    
    Returns:
        tuple: (customer_email, iccids, qr_code_url)
    """
    # Parse HTML
    soup = BeautifulSoup(input_html, 'html.parser')
    
    # Find QR code image
    qr_code_img = soup.find('img', {'src': re.compile(r'.*op-flow-public.*qrcode.*\.png')})
    qr_code_url = ''
    
    if qr_code_img and 'src' in qr_code_img.attrs:
        # Extract the actual URL from the Google proxy URL
        src = qr_code_img['src']
        match = re.search(r'#(https://.*\.png)', src)
        if match:
            qr_code_url = match.group(1)
    
    # Get customer email from billionconnect.html
    customer_email = ''
    # Look for text that contains "Dear Customer" to find the email
    dear_customer = soup.find('p', string=re.compile(r'Dear Customer'))
    if dear_customer:
        # Try to find customer email in nearby text
        email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', str(dear_customer.parent))
        if email_match:
            customer_email = email_match.group(0)
    
    # Get ICCIDs from billionconnect.html
    iccids = []
    # Look for text containing "Your ICCID is:"
    iccid_text = soup.find('span', {'class': 'm_-6454673849054811020rColor'})
    if iccid_text:
        iccid = re.search(r'\d{19,20}', iccid_text.text)
        if iccid:
            iccids.append(iccid.group(0))
    
    # If no email or ICCIDs found in billionconnect.html format, try order_processing.html format
    if not customer_email:
        customer_email_div = soup.find('div', {'id': 'customer_email'})
        if customer_email_div:
            customer_email = customer_email_div.text
    
    if not iccids:
        iccid_divs = soup.find_all('div', {'class': 'iccid'})
        for div in iccid_divs:
            iccids.append(div.text)
    
    return customer_email, iccids, qr_code_url

def append_to_order_processing(iccids, qr_code_url):
    """
    Create order processing HTML content with the parsed information
    
    Args:
        iccids (list): List of ICCIDs
        qr_code_url (str): URL of the QR code image
    
    Returns:
        str: Generated HTML content
    """
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

def main():
    try:
        with open('billionconnect.html', 'r', encoding='utf-8') as file:
            input_html = file.read()
            
        # Parse the content
        customer_email, iccids, qr_code_url = parse_qr_code(input_html)
        
        # Generate order processing content
        new_content = append_to_order_processing(iccids, qr_code_url)
        
        # Read the existing order_processing.html
        with open('order_processing.html', 'r', encoding='utf-8') as file:
            content = file.read()
        
        # Find the position to insert the new content
        first_margin = content.find('<div style="margin-bottom:20px">')
        if first_margin != -1:
            second_margin = content.find('<div style="margin-bottom:20px">', first_margin + 1)
            if second_margin != -1:
                # Insert the new content between the two margin-bottom divs
                new_full_content = (
                    content[:second_margin] + 
                    '\n<!-- Added by QR Code Parser -->\n' + 
                    new_content + 
                    content[second_margin:]
                )
                
                # Write the modified content back to the file
                with open('order_processing.html', 'w', encoding='utf-8') as file:
                    file.write(new_full_content)
                
                print("成功處理完成！")
                print(f"ICCID 數量: {len(iccids)}")
                print(f"QR Code URL: {qr_code_url}")
            else:
                raise Exception("找不到第二個 margin-bottom:20px div")
        else:
            raise Exception("找不到第一個 margin-bottom:20px div")
        
    except Exception as e:
        print(f"處理過程中發生錯誤: {str(e)}")

if __name__ == "__main__":
    main()
