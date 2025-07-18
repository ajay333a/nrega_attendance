# Attendance download process

You have to go to the attendance portal, then select state and date, after which the website will give list of Districts in which you have to click on given district which will take you to page that gives you list of taluks, click on taluk again, which will take you to panchayaths, now you click on Muster rools of Panchayath which will display a table for muster rolls, when clicked on a muster roll number it will take you to muster roll data which I want to download for all muster rolls for a given workcode.

Attendance Portal URL: https://mnregaweb4.nic.in/nregaarch/View_NMMS_atten_date_new.aspx?fin_year=2024-2025&Digest=HNrisV4bhHnb7Gve3mAKYQ

For first page State element name = 'ctl00$ContentPlaceHolder1$ddlstate'
For first page date element name = 'ctl00$ContentPlaceHolder1$ddl_attendance'

There is no CAPTCHA requirements.

## Implementation

1. Setup Selenium and Streamlit

2. Outline of Selenium Automation
The script will:
1. Open the portal URL.
2. Set Karnataka state, Ballari distict, Siruguppa Block/Taluk as defaults
3. For the selected Panchayath and Date find all Muster Rolls for the given Work Code.
For each Muster Roll, scrape data using the element names which I will provide you again when you reach Muster roll data page.
4. Elements to be scrapped from final page 1. Attendance table 2. Attendance Image

## Streamlit UI
1. Set Karnataka state, Ballari distict, Siruguppa Block/Taluk as defaults
2. Ask for Panchayath, Date, Work Code
3. On submit, run the Selenium script and provide the results for download

## Downloading Data

Combine all data into a 2 excel files one for attendance data other for images in a table format:

1. Attendance data
--------------------------------------------------------------------------------
| Muster Roll No.   | Attendance table                                          |
--------------------------------------------------------------------------------

2. Images and muster roll

--------------------------------------------
| Muster Roll No. | Image                   |
--------------------------------------------
