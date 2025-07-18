import os
import io
import requests
from openpyxl import Workbook
from openpyxl.drawing.image import Image as XLImage
from PIL import Image as PILImage
from bs4 import BeautifulSoup
from datetime import datetime
import argparse

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
    # Work Name
    work_name = None
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
    # Only extract these columns
    wanted_cols = ['S.No', 'Job Card No', 'Worker Name (Gender)', 'Attendance Date', 'Present/Absent']
    for row in rows[1:]:
        cols = row.find_all('td')
        if cols and any(c.text.strip() for c in cols):
            extracted = [cols[col_map.get(col, -1)].text.strip() if col in col_map else '' for col in wanted_cols]
            attendance_data.append(extracted)
    print(f"Parsed {len(attendance_data)} attendance rows for muster roll.")
    # Extract Photo URL
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
            progress_callback(f"Muster Roll {msr_no} parsed")
    file_base = f"{work_code}_{attendance_date}".replace('/', '_')

    # Write attendance data to Excel
    att_wb = Workbook()
    att_ws = att_wb.active
    att_ws.title = 'Attendance Data'
    att_ws.append(['Muster Roll No.'] + (table_headers if table_headers else []))
    for record in attendance_records:
        att_ws.append([record['muster_roll_no']] + record['row'])
    att_wb.save(f'attendance_data_{file_base}.xlsx')
    print(f'Saved attendance_data_{file_base}.xlsx')

    # Write images to Excel
    img_wb = Workbook()
    img_ws = img_wb.active
    img_ws.title = 'Images'
    img_ws.append(['Muster Roll No.', 'Image'])
    temp_img_paths = []
    img_col = 2  # Column B for images
    img_row = 2  # Start from row 2 (row 1 is header)
    for entry in image_records:
        muster_no = entry['muster_roll_no']
        img_bytes = entry['image']
        if img_bytes:
            img_bytes.seek(0)
            pil_img = PILImage.open(img_bytes)
            img_path = f'{muster_no}.png'
            pil_img.save(img_path)
            xl_img = XLImage(img_path)
            img_ws.cell(row=img_row, column=1, value=muster_no)
            cell_ref = f'B{img_row}'
            img_ws.add_image(xl_img, cell_ref)
            img_ws.row_dimensions[img_row].height = pil_img.height * 0.75
            img_ws.column_dimensions['B'].width = pil_img.width * 0.14
            temp_img_paths.append(img_path)
        else:
            img_ws.cell(row=img_row, column=1, value=muster_no)
            img_ws.cell(row=img_row, column=2, value='No Image')
        img_row += 1
    img_wb.save(f'attendance_images_{file_base}.xlsx')
    print(f'Saved attendance_images_{file_base}.xlsx')
    for img_path in temp_img_paths:
        if os.path.exists(img_path):
            os.remove(img_path)

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
    temp_img_paths_c = []
    for entry in option_c_records:
        muster_no = entry['muster_roll_no']
        att_rows = entry['attendance']
        img_bytes = entry['image']
        first_row = True
        s_no = 1
        if att_rows:
            for row in att_rows:
                # Remove time from Attendance Date (keep only date)
                if len(row) >= 5:
                    row[4] = row[4].split()[0] if row[4] else ''
                # Compose row as per header
                excel_row = [muster_no if first_row else '', str(s_no), row[1] if len(row) > 1 else '', row[2] if len(row) > 2 else '', row[4] if len(row) > 4 else '', row[5] if len(row) > 5 else '', '']
                if first_row:
                    if img_bytes:
                        img_bytes.seek(0)
                        pil_img = PILImage.open(img_bytes)
                        img_path = f'{muster_no}_optc.png'
                        pil_img.save(img_path)
                        xl_img = XLImage(img_path)
                        row_idx = optc_ws.max_row + 1
                        optc_ws.append(excel_row)
                        cell_ref = f'G{row_idx}'
                        optc_ws.add_image(xl_img, cell_ref)
                        optc_ws.row_dimensions[row_idx].height = pil_img.height * 0.75
                        optc_ws.column_dimensions['G'].width = pil_img.width * 0.14
                        temp_img_paths_c.append(img_path)
                    else:
                        optc_ws.append(excel_row)
                    first_row = False
                else:
                    optc_ws.append(excel_row)
                s_no += 1
        else:
            # No attendance rows, just show muster no and image
            excel_row = [muster_no, '', '', '', '', '', '']
            if img_bytes:
                img_bytes.seek(0)
                pil_img = PILImage.open(img_bytes)
                img_path = f'{muster_no}_optc.png'
                pil_img.save(img_path)
                xl_img = XLImage(img_path)
                row_idx = optc_ws.max_row + 1
                optc_ws.append(excel_row)
                cell_ref = f'G{row_idx}'
                optc_ws.add_image(xl_img, cell_ref)
                optc_ws.row_dimensions[row_idx].height = pil_img.height * 0.75
                optc_ws.column_dimensions['G'].width = pil_img.width * 0.14
                temp_img_paths_c.append(img_path)
            else:
                optc_ws.append(excel_row)
    optc_wb.save(f'attendance_with_images_{file_base}.xlsx')
    print(f'Saved attendance_with_images_{file_base}.xlsx')
    for img_path in temp_img_paths_c:
        if os.path.exists(img_path):
            os.remove(img_path)

    # Generate PDF report for attendance with images
    from reportlab.lib.pagesizes import landscape, A4
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
    from reportlab.lib.styles import getSampleStyleSheet

    pdf_filename = f'attendance_with_images_{file_base}.pdf'
    doc = SimpleDocTemplate(pdf_filename, pagesize=landscape(A4))
    elements = []
    styles = getSampleStyleSheet()

    # Add headers
    elements.append(Paragraph(f'Work Code: {work_code}', styles['Heading2']))
    elements.append(Paragraph(f'Work Name: {work_name if work_name else ""}', styles['Normal']))
    elements.append(Paragraph(f'District: {DEFAULT_DISTRICT}', styles['Normal']))
    elements.append(Paragraph(f'Taluk/Block: {DEFAULT_TALUK}', styles['Normal']))
    elements.append(Paragraph(f'Panchayath Name: {panchayat_name}', styles['Normal']))
    elements.append(Spacer(1, 12))

    for entry in option_c_records:
        muster_no = entry['muster_roll_no']
        att_rows = entry['attendance']
        img_bytes = entry['image']
        # Table header
        table_header = ['S.No', 'Job Card No', 'Worker Name(Gender)', 'Attendance Date', 'Present/Absent']
        data = [table_header]
        s_no = 1
        if att_rows:
            for row in att_rows:
                # Remove time from Attendance Date (keep only date)
                if len(row) >= 5:
                    row[4] = row[4].split()[0] if row[4] else ''
                data.append([
                    str(s_no),
                    row[1] if len(row) > 1 else '',
                    row[2] if len(row) > 2 else '',
                    row[4] if len(row) > 4 else '',
                    row[5] if len(row) > 5 else ''
                ])
                s_no += 1
        # Add muster roll number as a title
        elements.append(Spacer(1, 12))
        elements.append(Paragraph(f'<b>Muster Roll No.: {muster_no}</b>', styles['Heading3']))
        # Add table and image side by side
        table_width = 400
        img_width = 250
        t = Table(data, repeatRows=1, colWidths=[40, 120, 120, 100, 80])
        t.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
            ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        max_img_height = 200  # points
        if img_bytes:
            img_bytes.seek(0)
            rl_img = RLImage(img_bytes, width=img_width, mask='auto')
            # Scale image if too tall
            if rl_img.imageHeight > max_img_height:
                rl_img.drawHeight = max_img_height
                rl_img.drawWidth = rl_img.imageWidth * (max_img_height / rl_img.imageHeight)
            # If table is short, place side by side; else, stack
            if len(data) <= 10:
                side_by_side = Table([[t, rl_img]], colWidths=[table_width, img_width])
                side_by_side.setStyle(TableStyle([
                    ('VALIGN', (0,0), (-1,-1), 'TOP'),
                    ('ALIGN', (0,0), (-1,-1), 'CENTER'),
                ]))
                elements.append(side_by_side)
            else:
                elements.append(t)
                elements.append(Spacer(1, 8))
                elements.append(rl_img)
        else:
            elements.append(t)
        elements.append(Spacer(1, 24))
    doc.build(elements)
    print(f'Saved {pdf_filename}')
    # Return file names for frontend to use
    return file_base
