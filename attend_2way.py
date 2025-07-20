import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin
import time
import openpyxl
from openpyxl.styles import Alignment, Font
import io
import sys
sys.path.append('.')
from attendance_downloader import get_attendance_data, download_photo
from openpyxl.drawing.image import Image as XLImage

url = "https://mnregaweb4.nic.in/nregaarch/View_NMMS_atten_date_new.aspx?fin_year=2024-2025&Digest=HNrisV4bhHnb7Gve3mAKYQ"
session = requests.Session()
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/58.0.3029.110 Safari/537.3'
}
resp = session.get(url, headers=headers)
soup = BeautifulSoup(resp.content, 'html.parser')

viewstate = soup.find('input', {'id': '__VIEWSTATE'})['value']
eventvalidation = soup.find('input', {'id': '__EVENTVALIDATION'})['value']
viewstategen = soup.find('input', {'id': '__VIEWSTATEGENERATOR'})['value']

state_value = '15'  # Karnataka
attendance_select = soup.find('select', {'name': 'ctl00$ContentPlaceHolder1$ddl_attendance'})
date_options = [opt['value'] for opt in attendance_select.find_all('option')]
print("Available dates:", date_options)
attendance_date = input("Enter attendance date from above options (e.g., 18/07/2025): ").strip()

data = {
    '__VIEWSTATE': viewstate,
    '__VIEWSTATEGENERATOR': viewstategen,
    '__EVENTVALIDATION': eventvalidation,
    'ctl00$ContentPlaceHolder1$ddlstate': state_value,
    'ctl00$ContentPlaceHolder1$ddl_attendance': attendance_date,
    'ctl00$ContentPlaceHolder1$btn_showreport': 'Show Attendance',
}
headers_post = headers.copy()
headers_post['Referer'] = url
resp2 = session.post(url, data=data, headers=headers_post)
time.sleep(3)
soup2 = BeautifulSoup(resp2.content, 'html.parser')

# State table
state_table = soup2.find('table', {'id': 'grdTable'})
if not state_table:
    state_div = soup2.find('div', {'id': 'RepPr1'})
    if state_div:
        state_table = state_div.find('table')
if not state_table:
    print("Could not find state table by id or inside <div id='RepPr1'>.")
    exit(1)
karnataka_link = None
for row in state_table.find_all('tr'):
    cols = row.find_all('td')
    if len(cols) > 1 and cols[1].get_text(strip=True).upper() == 'KARNATAKA':
        a = cols[1].find('a', href=True)
        if a:
            karnataka_link = a['href']
        break
if not karnataka_link:
    print("Could not find Karnataka link in state table.")
    exit(1)
karnataka_url = urljoin(url, karnataka_link)

# Districts table
resp3 = session.get(karnataka_url, headers=headers)
soup3 = BeautifulSoup(resp3.content, 'html.parser')
dist_table = soup3.find('table', {'id': 'grdTable'})
if not dist_table:
    dist_div = soup3.find('div', {'id': 'RepPr1'})
    if dist_div:
        dist_table = dist_div.find('table')
if not dist_table:
    print("Could not find districts table by id or inside <div id='RepPr1'>.")
    exit(1)
ballari_link = None
for row in dist_table.find_all('tr'):
    cols = row.find_all('td')
    if len(cols) >= 4:
        s_no = cols[0].get_text(strip=True)
        if not s_no.isdigit():
            continue
        district_name = cols[1].get_text(strip=True)
        a = cols[1].find('a', href=True)
        href = a['href'] if a else None
        if district_name.upper() == 'BALLARI' and href:
            ballari_link = href
            break
if not ballari_link:
    print("Could not find Ballari link in districts table.")
    exit(1)
ballari_url = urljoin(karnataka_url, ballari_link)

# Block/Taluk table
resp4 = session.get(ballari_url, headers=headers)
soup4 = BeautifulSoup(resp4.content, 'html.parser')
block_table = soup4.find('table', {'id': 'grdTable'})
if not block_table:
    block_div = soup4.find('div', {'id': 'RepPr1'})
    if block_div:
        block_table = block_div.find('table')
if not block_table:
    print("Could not find block/taluk table by id or inside <div id='RepPr1'>.")
    exit(1)
siruguppa_link = None
for row in block_table.find_all('tr'):
    cols = row.find_all('td')
    if len(cols) >= 2:
        s_no = cols[0].get_text(strip=True)
        if not s_no.isdigit():
            continue
        block_name = cols[1].get_text(strip=True)
        a = cols[1].find('a', href=True)
        href = a['href'] if a else None
        if block_name.strip().upper() == 'SIRUGUPPA' and href:
            siruguppa_link = href
            break
if not siruguppa_link:
    print("Could not find Siruguppa link in block/taluk table.")
    exit(1)
siruguppa_url = urljoin(ballari_url, siruguppa_link)

# Panchayath table
resp5 = session.get(siruguppa_url, headers=headers)
soup5 = BeautifulSoup(resp5.content, 'html.parser')
panch_div = soup5.find('div', {'id': 'RepPr1'})
if not panch_div:
    print("Could not find panchayath table container <div id='RepPr1'>.")
    exit(1)
panch_table = panch_div.find('table')
if not panch_table:
    print("Could not find panchayath table inside <div id='RepPr1'>.")
    exit(1)
panchayath_name = input("Enter Panchayath name: ").strip().upper()
panchayath_link = None
for row in panch_table.find_all('tr'):
    cols = row.find_all('td')
    if len(cols) >= 4:
        s_no = cols[0].get_text(strip=True)
        if not s_no.isdigit():
            continue
        panch_name = cols[1].get_text(strip=True).upper()
        muster_rolls_a = cols[3].find('a', href=True)
        href = muster_rolls_a['href'] if muster_rolls_a else None
        if panch_name == panchayath_name and href:
            panchayath_link = href
            break
if not panchayath_link:
    print("No NMR generated by the Panchayath")
    exit(0)
panchayath_url = urljoin(siruguppa_url, panchayath_link)

# Muster Roll table (always extract from <div id='RepPr1'>)
resp6 = session.get(panchayath_url, headers=headers)
soup6 = BeautifulSoup(resp6.content, 'html.parser')
muster_div = soup6.find('div', {'id': 'RepPr1'})
if not muster_div:
    print("Could not find muster roll table container <div id='RepPr1'>.")
    exit(1)
muster_table = muster_div.find('table')
if not muster_table:
    print("Could not find muster roll table inside <div id='RepPr1'>.")
    exit(1)
# Find the column indices for Work Code and Muster Roll No. (case-insensitive, strip spaces)
header_row = muster_table.find('tr')
header_cols = [th.get_text(strip=True).replace('\u00a0', ' ').strip().lower() for th in header_row.find_all(['th', 'td'])]
try:
    workcode_idx = next(i for i, h in enumerate(header_cols) if 'work code' in h)
    muster_no_idx = next(i for i, h in enumerate(header_cols) if 'mustroll no' in h)
except StopIteration:
    print("Could not find required columns in muster roll table header.")
    print("Header columns found:", header_cols)
    exit(1)

choice = input("Type 'all' for all muster rolls or 'work' for specific work: ").strip().lower()
rows_to_save = []
if choice == 'all':
    for row in muster_table.find_all('tr')[1:]:
        cols = row.find_all('td')
        if len(cols) > muster_no_idx:
            muster_a = cols[muster_no_idx].find('a', href=True)
            if muster_a:
                muster_href = muster_a['href']
                rows_to_save.append((cols, muster_href))
elif choice == 'work':
    workcode = input("Enter workcode: ").strip()
    for row in muster_table.find_all('tr')[1:]:
        cols = row.find_all('td')
        if len(cols) > muster_no_idx and len(cols) > workcode_idx:
            if workcode in cols[workcode_idx].get_text(strip=True):
                muster_a = cols[muster_no_idx].find('a', href=True)
                if muster_a:
                    muster_href = muster_a['href']
                    rows_to_save.append((cols, muster_href))
else:
    print("Invalid choice.")
    exit(1)

if not rows_to_save:
    print("No muster roll data found for the selection.")
    exit(0)

# For each muster roll, request the page, extract attendance table and image, and save to Excel
wb = openpyxl.Workbook()
ws = wb.active
row_cursor = 1
# Add custom header section at the top
ws.cell(row=row_cursor, column=1, value="District:").font = Font(bold=True)
ws.cell(row=row_cursor, column=2, value="Ballari")
ws.cell(row=row_cursor, column=3, value="Taluk/Block:").font = Font(bold=True)
ws.cell(row=row_cursor, column=4, value="Siruguppa")
row_cursor += 1
ws.cell(row=row_cursor, column=1, value="Panchayath:").font = Font(bold=True)
ws.cell(row=row_cursor, column=2, value=panchayath_name)
row_cursor += 1
# We'll fill Work code and Work Name after fetching the first muster roll
first_header_cells = None
first_attendance_data = None
first_img_bytes = None
first_muster_processed = False
first_work_code = ''
first_work_name = ''
# For image-only Excel
img_wb = openpyxl.Workbook()
img_ws = img_wb.active
img_row_cursor = 1
img_bytes_refs = []  # Keep references to BytesIO objects for images
# Add header section to image-only Excel
img_ws.cell(row=img_row_cursor, column=1, value="District:").font = Font(bold=True)
img_ws.cell(row=img_row_cursor, column=2, value="Ballari")
img_ws.cell(row=img_row_cursor, column=3, value="Taluk/Block:").font = Font(bold=True)
img_ws.cell(row=img_row_cursor, column=4, value="Siruguppa")
img_row_cursor += 1
img_ws.cell(row=img_row_cursor, column=1, value="Panchayath:").font = Font(bold=True)
img_ws.cell(row=img_row_cursor, column=2, value=panchayath_name)
img_row_cursor += 1
# We'll fill Work code and Work Name after fetching the first muster roll
img_first_work_code = ''
img_first_work_name = ''
img_first_muster_processed = False
# Pre-reserve a row for Work code/Work Name (to be filled after first muster roll)
workcode_row_idx = img_row_cursor
img_ws.cell(row=workcode_row_idx, column=1, value="Work code:").font = Font(bold=True)
img_ws.cell(row=workcode_row_idx, column=3, value="Work Name:").font = Font(bold=True)
img_row_cursor += 1
# Add header row to image-only Excel (after custom header)
header_row_idx = img_row_cursor
img_ws.cell(row=header_row_idx, column=1, value='Muster Roll No').font = Font(bold=True)
img_ws.cell(row=header_row_idx, column=2, value='Image').font = Font(bold=True)
img_row_cursor += 1
for cols, muster_href in rows_to_save:
    muster_url = urljoin(panchayath_url, muster_href)
    attendance_data, photo_url, work_name, header_cells = get_attendance_data(muster_url)
    img_bytes = download_photo(photo_url) if photo_url else None
    muster_roll_no = cols[muster_no_idx].get_text(strip=True)
    # Fill work code and work name in header if not done
    if not img_first_muster_processed:
        img_first_work_code = cols[workcode_idx].get_text(strip=True)
        img_first_work_name = work_name or ''
        img_ws.cell(row=workcode_row_idx, column=2, value=img_first_work_code)
        img_ws.cell(row=workcode_row_idx, column=4, value=img_first_work_name)
        img_first_muster_processed = True
    # Attendance data rows, prepend muster_roll_no
    if attendance_data:
        for att_row in attendance_data:
            ws.cell(row=row_cursor, column=1, value=muster_roll_no)
            for col_idx, val in enumerate(att_row, 2):
                ws.cell(row=row_cursor, column=col_idx, value=val)
            row_cursor += 1
    # Insert image (if available)
    if img_bytes:
        img_bytes.seek(0)
        img = XLImage(img_bytes)
        # Place image in column H
        img_cell = f"H{row_cursor-len(attendance_data) if attendance_data else row_cursor}"
        ws.add_image(img, img_cell)
        row_cursor += 3  # leave space for image
    else:
        row_cursor += 2
    row_cursor += 2  # space between muster rolls
    # --- Image-only Excel ---
    start_img_row = img_row_cursor
    img_ws.cell(row=img_row_cursor, column=1, value=muster_roll_no).font = Font(bold=True, size=18)
    if img_bytes:
        img_bytes.seek(0)
        img_bytes_for_imgwb = io.BytesIO(img_bytes.getbuffer())
        img2 = XLImage(img_bytes_for_imgwb)
        img_ws.add_image(img2, f"B{img_row_cursor}")
        img_bytes_refs.append(img_bytes_for_imgwb)  # Keep reference alive
        # Estimate image height in rows (20 is a rough guess, adjust if needed)
        img_height_rows = 20
        end_img_row = img_row_cursor + img_height_rows - 1
        img_ws.merge_cells(start_row=start_img_row, start_column=1, end_row=end_img_row, end_column=1)
        img_ws.cell(row=start_img_row, column=1).alignment = Alignment(vertical='center', horizontal='center')
        img_row_cursor += img_height_rows
    else:
        end_img_row = img_row_cursor
        img_row_cursor += 3
        img_ws.merge_cells(start_row=start_img_row, start_column=1, end_row=end_img_row, end_column=1)
        img_ws.cell(row=start_img_row, column=1).alignment = Alignment(vertical='center', horizontal='center')
    img_row_cursor += 2  # space between entries
# If no muster rolls, still fill work code/name as blank
if not img_first_muster_processed:
    img_ws.cell(row=workcode_row_idx, column=2, value='')
    img_ws.cell(row=workcode_row_idx, column=4, value='')
wb.save(f"muster_rolls_{panchayath_name}_{attendance_date.replace('/', '_')}.xlsx")
img_wb.save(f"muster_roll_images_{panchayath_name}_{attendance_date.replace('/', '_')}.xlsx")
print(f"Saved muster_rolls_{panchayath_name}_{attendance_date.replace('/', '_')}.xlsx")
print(f"Saved muster_roll_images_{panchayath_name}_{attendance_date.replace('/', '_')}.xlsx")
