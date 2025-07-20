import io
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from bs4 import BeautifulSoup
from datetime import datetime
from openpyxl.styles import Alignment, Font

STARTING_URL = "https://mnregaweb4.nic.in/nregaarch/View_NMMS_atten_date_dtl_rpt.aspx?page=&short_name=KN&state_name=KARNATAKA&state_code=15&district_name=BALLARI&district_code=1505&block_name=SIRUGUPPA&block_code=1505007&"

DEFAULT_DISTRICT = "Ballari"
DEFAULT_TALUK = "Siruguppa"

def get_attendance_data(url):
    print(f"Fetching attendance data from: {url}")
    try:
        response = requests.get(url)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        print(f"Error fetching attendance data: {e}")
        return None, None, None, None
    soup = BeautifulSoup(response.content, 'html.parser')
    # Work Name (extract text after <b>Work Name</b>)
    work_name = None
    for b in soup.find_all('b'):
        if b.text.strip().startswith('Work Name'):
            # Get the next_sibling text after the <b>Work Name ...</b>
            next_text = b.next_sibling
            if next_text:
                work_name = str(next_text).strip(' :\u00a0-')
            break
    # Fallback to previous method if not found
    if not work_name:
        work_name_elem = soup.find(id="ContentPlaceHolder1_lbl_dtl")
        if work_name_elem:
            work_name = work_name_elem.text.strip()
    # Attendance Table
    attendance_data = []
    tables = soup.find_all('table')
    if not tables:
        print("No tables found on the page.")
        return None, None, work_name, None
    attendance_table = tables[-1]
    rows = attendance_table.find_all('tr')
    header_cells = [th.text.strip() for th in rows[0].find_all(['th', 'td'])]
    col_map = {name: idx for idx, name in enumerate(header_cells)}
    wanted_cols = ['S.No', 'Job Card No', 'Worker Name (Gender)', 'Attendance Date', 'Present/Absent']
    for row in rows[1:]:
        cols = row.find_all('td')
        if cols and any(c.get_text(strip=True) for c in cols):
            name_td = ''
            for td in cols:
                span = td.find('span', id=lambda x: x and 'lbl_workerName_' in x)
                if span:
                    name_td = span.get_text(strip=True)
                    break
            extracted = [
                cols[col_map.get('S.No', -1)].get_text(strip=True) if 'S.No' in col_map else '',
                cols[col_map.get('Job Card No', -1)].get_text(strip=True) if 'Job Card No' in col_map else '',
                name_td,
                cols[col_map.get('Attendance Date', -1)].get_text(strip=True) if 'Attendance Date' in col_map else '',
                cols[col_map.get('Present/Absent', -1)].get_text(strip=True) if 'Present/Absent' in col_map else ''
            ]
            attendance_data.append(extracted)
    photo_url = None
    img_link = soup.find('a', text='Click here for large image')
    if img_link and img_link.has_attr('href'):
        photo_url = img_link['href']
    return attendance_data, photo_url, work_name, header_cells

def download_photo(url):
    if not url:
        return None
    try:
        response = requests.get(url, stream=True)
        response.raise_for_status()
        img_bytes = io.BytesIO(response.content)
        return img_bytes
    except requests.exceptions.RequestException as e:
        print(f"Error downloading photo: {e}")
        return None

def run_attendance_downloader(panchayat_name, panchayat_code, fin_year, work_code, msr_start, msr_end, attendance_date, digest, progress_callback=None):
    attendance_records = []
    image_records = []
    option_c_records = []
    work_name = None
    table_headers = None
    for msr_no in range(msr_start, msr_end + 1):
        # URL with work code
        url_with_workcode = (
            f"{STARTING_URL}"
            f"panchayat_name={panchayat_name}&panchayat_code={panchayat_code}"
            f"&fin_year={fin_year}"
            f"&source=&work_code={work_code}"
            f"&msr_no={msr_no}"
            f"&AttendanceDate={attendance_date}"
            f"&Digest={digest}"
        )
        # URL without work code
        url_without_workcode = (
            f"{STARTING_URL}"
            f"panchayat_name={panchayat_name}&panchayat_code={panchayat_code}"
            f"&fin_year={fin_year}"
            f"&source="
            f"&msr_no={msr_no}"
            f"&AttendanceDate={attendance_date}"
            f"&Digest={digest}"
        )
        # Try with work code first
        att_data, photo_url, wname, headers = get_attendance_data(url_with_workcode)
        used_url = 'with work code'
        if not att_data:
            att_data, photo_url, wname, headers = get_attendance_data(url_without_workcode)
            used_url = 'without work code'
        if wname and not work_name:
            work_name = wname
        if headers and not table_headers:
            table_headers = headers
        if att_data:
            for row in att_data:
                attendance_records.append({'muster_roll_no': msr_no, 'row': row})
        if photo_url:
            img_bytes = download_photo(photo_url)
            image_records.append({'muster_roll_no': msr_no, 'image': img_bytes})
        else:
            image_records.append({'muster_roll_no': msr_no, 'image': None})
        # For option C
        option_c_records.append({'muster_roll_no': msr_no, 'attendance': att_data, 'image': img_bytes if photo_url else None})
        if progress_callback:
            progress_callback(f"Muster Roll {msr_no} parsed ,")
    file_base = f"{work_code}_{attendance_date}".replace('/', '_')

    # Write attendance data to Excel
    att_wb = Workbook()
    att_ws = att_wb.active
    att_ws.title = 'Attendance Data'
    # Add header section
    att_ws.append([f'Work Code: {work_code}'])
    att_ws.append([f'Work Name: {work_name if work_name else ""}'])
    att_ws.append([f'District: {DEFAULT_DISTRICT}'])
    att_ws.append([f'Taluk/Block: {DEFAULT_TALUK}'])
    att_ws.append([f'Panchayath Name: {panchayat_name}'])
    att_ws.append([])
    
    for record in attendance_records:
        att_ws.append([record['muster_roll_no']] + record['row'])
    # att_wb.save(f'attendance_data_{file_base}.xlsx') # Commented out
    print(f'Saved attendance_data_{file_base}.xlsx')

    # Write images to Excel
    img_wb = Workbook()
    img_ws = img_wb.active
    img_ws.title = 'Images'
    # Add header section
    img_ws.append([f'Work Code: {work_code}'])
    img_ws.append([f'Work Name: {work_name if work_name else ""}'])
    img_ws.append([f'District: {DEFAULT_DISTRICT}'])
    img_ws.append([f'Taluk/Block: {DEFAULT_TALUK}'])
    img_ws.append([f'Panchayath Name: {panchayat_name}'])
    img_ws.append([])
    img_ws.append(['Muster Roll No.', 'Image'])
    img_row = img_ws.max_row + 1  # Start after header
    img_refs = []  # Keep references to BytesIO objects
    for entry in image_records:
        muster_no = entry['muster_roll_no']
        img_bytes = entry['image']
        if img_bytes:
            img_bytes.seek(0)
            img_data = img_bytes.read()
            img_bytes_for_excel = io.BytesIO(img_data)
            xl_img = XLImage(img_bytes_for_excel)
            img_refs.append(img_bytes_for_excel)  # Keep reference alive
            cell = img_ws.cell(row=img_row, column=1, value=muster_no)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(size=16, bold=True)
            cell_ref = f'B{img_row}'
            img_ws.add_image(xl_img, cell_ref)
            # Dynamically set row height based on image height
            from PIL import Image as PILImage
            img_bytes_for_excel.seek(0)
            pil_img = PILImage.open(img_bytes_for_excel)
            img_height_px = pil_img.height
            img_height_pt = img_height_px * 0.75  # 1 px ≈ 0.75 pt
            img_ws.row_dimensions[img_row].height = img_height_pt
            img_ws.column_dimensions['B'].width = 20
        else:
            cell = img_ws.cell(row=img_row, column=1, value=muster_no)
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.font = Font(size=16, bold=True)
            img_ws.cell(row=img_row, column=2, value='No Image')
        img_row += 1
    # img_wb.save(f'attendance_images_{file_base}.xlsx') # Commented out
    print(f'Saved attendance_images_{file_base}.xlsx')

    # Write Option C Excel (attendance + image in one file)
    optc_wb = Workbook()
    optc_ws = optc_wb.active
    optc_ws.title = 'Attendance+Images'
    # Header rows
    optc_ws.append([f'Work Code: {work_code}'])
    optc_ws.append([f'Work Name: {work_name if work_name else ""}'])
    optc_ws.append([f'District: {DEFAULT_DISTRICT}'])
    optc_ws.append([f'Taluk/Block: {DEFAULT_TALUK}'])
    optc_ws.append([f'Panchayath Name: {panchayat_name}'])
    optc_ws.append([])
    # Table header (match screenshot order)
    table_header = ['Muster Roll No.', 'S.No', 'Job Card No', 'Worker Name(Gender)', 'Attendance Date', 'Present/Absent', 'Image']
    optc_ws.append(table_header)
    img_refs_c = []  # Keep references to BytesIO objects for Option C
    for entry in option_c_records:
        muster_no = entry['muster_roll_no']
        att_rows = entry['attendance']
        img_bytes = entry['image']
        first_row = True
        if att_rows:
            for row in att_rows:
                # row: [S.No, Job Card No, Worker Name(Gender), Attendance Date, Present/Absent]
                # Fix date to keep 'DD Mon YYYY'
                if len(row) >= 4 and row[3]:
                    row[3] = ' '.join(row[3].split()[:3])
                excel_row = [muster_no if first_row else '',
                             row[0] if len(row) > 0 else '',  # S.No from table
                             row[1] if len(row) > 1 else '',  # Job Card No
                             row[2] if len(row) > 2 else '',  # Worker Name(Gender)
                             row[3] if len(row) > 3 else '',  # Attendance Date
                             row[4] if len(row) > 4 else '',  # Present/Absent
                             '']
                if first_row:
                    if img_bytes:
                        img_bytes.seek(0)
                        img_data = img_bytes.read()
                        img_bytes_for_excel = io.BytesIO(img_data)
                        xl_img = XLImage(img_bytes_for_excel)
                        img_refs_c.append(img_bytes_for_excel)  # Keep reference alive
                        row_idx = optc_ws.max_row + 1
                        optc_ws.append(excel_row)
                        cell_ref = f'G{row_idx}'
                        optc_ws.add_image(xl_img, cell_ref)
                        optc_ws.row_dimensions[row_idx].height = 100
                        optc_ws.column_dimensions['G'].width = 20
                    else:
                        optc_ws.append(excel_row)
                    first_row = False
                else:
                    optc_ws.append(excel_row)
        else:
            excel_row = [muster_no, '', '', '', '', '', '']
            if img_bytes:
                img_bytes.seek(0)
                img_data = img_bytes.read()
                img_bytes_for_excel = io.BytesIO(img_data)
                xl_img = XLImage(img_bytes_for_excel)
                img_refs_c.append(img_bytes_for_excel)  # Keep reference alive
                row_idx = optc_ws.max_row + 1
                optc_ws.append(excel_row)
                cell_ref = f'G{row_idx}'
                optc_ws.add_image(xl_img, cell_ref)
                optc_ws.row_dimensions[row_idx].height = 100
                optc_ws.column_dimensions['G'].width = 20
    else:
                optc_ws.append(excel_row)
    # optc_wb.save(f'attendance_with_images_{file_base}.xlsx') # Commented out
    print(f'Saved attendance_with_images_{file_base}.xlsx')


    # Save Excel files to memory instead of disk
    att_xlsx = io.BytesIO()
    att_wb.save(att_xlsx)
    att_xlsx.seek(0)
    img_xlsx = io.BytesIO()
    img_wb.save(img_xlsx)
    img_xlsx.seek(0)
    optc_xlsx = io.BytesIO()
    optc_wb.save(optc_xlsx)
    optc_xlsx.seek(0)
    # Return in-memory files for frontend
    return att_xlsx, img_xlsx, optc_xlsx, None
