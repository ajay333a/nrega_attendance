# Attendance Download

## Let's first create the scrit for downloading data after that lets add Streamlit UI for user interface

### URL for direct Data
As there was a long list pages and clicks between data and process, I will provide you with direct link, but you have edit the link at various locations and join it with the variables I give you to go to the direct Muster roll page.

link: https://mnregaweb4.nic.in/nregaarch/View_NMMS_atten_date_dtl_rpt.aspx?page=&short_name=KN&state_name=KARNATAKA&state_code=15&district_name=BALLARI&district_code=1505&block_name=SIRUGUPPA&block_code=1505007&panchayat_name=BALAKUNDHI&panchayat_code=1505007016&fin_year=2024-2025&source=&work_code=1505007016/IC/93393042892330899&msr_no=16950&AttendanceDate=03/07/2025&Digest=ttm2SylWYUsKMSsmiDjONQ

parts of the link:
1. starting = https://mnregaweb4.nic.in/nregaarch/View_NMMS_atten_date_dtl_rpt.aspx?page=&short_name=KN&state_name=KARNATAKA&state_code=15&district_name=BALLARI&district_code=1505&block_name=SIRUGUPPA&block_code=1505007&
This part of link is fine as we are setting Karnataka, Ballari and Siruguppa as default

Ask for the rest of elements in the CLI or APP

2. panchayth name = panchayat_name=BALAKUNDHI&panchayat_code=1505007016
    - two elements are needed to be provided here
    1. Panchayath Name
    2. Panchayath Code

3. financial year = 
    fin_year=2024-2025

4. source will add work code of the work
    source=&work_code?

5. msr_no is Muster Roll number for which I will give a range and loop through them to collect data for all

6. AttendendanceDate is the date which you will ask input for

7. Digest , i cannot see any pattern for that so, ask input for that firld to


you should open a link

link = starting+panchayath_name+panchayath_code+fin_year+source=&+work_code+msr_no+Date+digest

you should open that link and search for data and download it into two excel files




Now you are under acheiving

url --[enter details in form] - State table(appears on same page after filling the fields in the input boxes) --> url --> Districts table(Find Ballari district and No of Muster rolls)



## state table 
state table element is in the element
<div id="RepPr1" class="table-responsive">

try to find table by id

## District table
district table element is in
<div id="RepPr1" class="table-responsive">
try to find the table by id

I want to go deeper, just how you found href links for district, find href for taluk and Panchayaths too
## Block/Taluk table
block/taluk table is in element
<div id="RepPr1" class="table-responsive">

