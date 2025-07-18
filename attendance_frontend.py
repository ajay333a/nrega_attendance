import streamlit as st
import os
from datetime import date
from attendance_downloader import run_attendance_downloader

st.title('Attendance Downloader')

# Session state for reset and file tracking
if 'submitted' not in st.session_state:
    st.session_state['submitted'] = False
if 'file_base' not in st.session_state:
    st.session_state['file_base'] = None
if 'progress_msgs' not in st.session_state:
    st.session_state['progress_msgs'] = []

# User input fields
panchayat_name = st.text_input('Panchayath Name (e.g., BALAKUNDHI)', key='panchayat_name')
panchayat_code = st.text_input('Panchayath Code (last 3 digits, e.g., 016)', key='panchayat_code')
fin_year = st.text_input('Financial Year (e.g., 2024-2025)', value='2024-2025', key='fin_year')
work_code = st.text_input('Work Code', key='work_code')
msr_start = st.number_input('Muster Roll Start Number', min_value=1, step=1, key='msr_start')
msr_end = st.number_input('Muster Roll End Number', min_value=1, step=1, key='msr_end')
attendance_date = st.date_input('Attendance Date', value=date.today(), key='attendance_date')
digest = st.text_input('Digest', key='digest')

# Progress area
progress_area = st.empty()

# Progress callback for backend
progress_msgs = st.session_state.get('progress_msgs', [])
def progress_callback(msg):
    progress_msgs.append(msg)
    st.session_state['progress_msgs'] = progress_msgs
    progress_area.write('\n'.join(progress_msgs))

# Download button and status
download_btn_col, status_col = st.columns([2, 1])
with download_btn_col:
    download_clicked = st.button('Download Attendance Data')

file_base = st.session_state.get('file_base')
if download_clicked:
    st.session_state['submitted'] = True
    st.session_state['progress_msgs'] = []
    att_date_str = attendance_date.strftime('%d/%m/%Y')
    if not panchayat_code.startswith('1505007'):
        panchayat_code_full = '1505007' + panchayat_code
    else:
        panchayat_code_full = panchayat_code
    st.info('Running backend process...')
    file_base = run_attendance_downloader(
        panchayat_name, panchayat_code_full, fin_year, work_code, int(msr_start), int(msr_end), att_date_str, digest,
        progress_callback=progress_callback
    )
    st.session_state['file_base'] = file_base

# Show progress messages
if st.session_state.get('progress_msgs'):
    progress_area.write('\n'.join(st.session_state['progress_msgs']))

# Show download buttons if files exist
file_base = st.session_state.get('file_base')
if file_base:
    att_date_str = attendance_date.strftime('%d/%m/%Y')
    excel1 = f'attendance_data_{file_base}.xlsx'
    excel2 = f'attendance_images_{file_base}.xlsx'
    excel3 = f'attendance_with_images_{file_base}.xlsx'
    pdf = f'attendance_with_images_{file_base}.pdf'
    for fname, label in zip([excel1, excel2, excel3, pdf],
                            ['Attendance Data Excel', 'Images Excel', 'Attendance+Images Excel', 'PDF Report']):
        if os.path.exists(fname):
            with open(fname, 'rb') as f:
                st.download_button(f'Download {label}', f, file_name=fname, key=fname)
    with status_col:
        st.success('✔️ Parsing complete! Files are ready for download.')
    # Reset button below download buttons
    if st.button('Reset App'):
        for key in list(st.session_state.keys()):
            del st.session_state[key]
        st.rerun() 