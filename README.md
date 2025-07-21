# Attendance Download Automation

This project automates the process of downloading muster roll attendance data (including images) from the MGNREGA web portal for Karnataka (Ballari district, Siruguppa block). It provides both a command-line interface (CLI) and a Streamlit web frontend for user interaction, and outputs the data into Excel files.

---

## Project Structure

```
attendence_download/
├── attend_2way.py              # CLI script for interactive attendance download
├── attendance_downloader.py    # Core backend logic for scraping and Excel output
├── attendance_frontend.py      # Streamlit web frontend
├── PRD.md                     # Product requirements and workflow
├── requirements.txt            # Python dependencies
├── README.md                   # Project documentation (this file)
└── venv/                       # (Optional) Python virtual environment
```

---

## File Descriptions

- **attend_2way.py**
  - CLI tool for interactively downloading attendance data and images for a selected Panchayath and date.
  - Navigates the MGNREGA portal, prompts for user input, fetches data, and saves:
    - A detailed Excel file (attendance + images)
    - An image-only Excel file
    - A raw data Excel file (no headers/images, just attendance info)
  - Now supports parallel fetching for faster execution.

- **attendance_downloader.py**
  - Core backend logic for fetching, parsing, and saving attendance data and images.
  - Provides functions for scraping muster roll data and generating Excel files.
  - Used by both the CLI and the Streamlit frontend.

- **attendance_frontend.py**
  - Streamlit web app for user-friendly attendance data download.
  - Collects user input, calls backend logic, and provides download buttons for generated Excel files.

- **PRD.md**
  - Product Requirements Document describing the workflow, required user inputs, and expected outputs.

- **requirements.txt**
  - Lists all Python dependencies required for the project.

---

## Installation

1. **Clone the repository** (if not already):
   ```bash
   git clone <repo-url>
   cd attendence_download
   ```

2. **(Optional) Create a virtual environment:**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```

3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

---

## Usage

### 1. **Command-Line Interface (CLI) - `attend_2way.py`**

Run the script:
```bash
python attend_2way.py
```

**You will be prompted for:**
- Attendance date (choose from available options)
- Panchayath name
- Whether to download all muster rolls or for a specific work code
- (If 'work') The work code

**Outputs:**
- `muster_rolls_<PANCHAYATH>_<DATE>.xlsx` — Attendance data with images
- `muster_roll_images_<PANCHAYATH>_<DATE>.xlsx` — Image-only Excel
- `muster_rolls_raw_<PANCHAYATH>_<DATE>.xlsx` — Raw attendance data (no headers/images)

**Features:**
- Fast parallel fetching of muster roll data
- Robust error handling and user prompts

### 2. **Web Frontend (Streamlit) - `attendance_frontend.py`**

Run the Streamlit app:
```bash
streamlit run attendance_frontend.py
```

**In the web UI, provide:**
- Panchayath Name and Code
- Financial Year
- Work Code
- Muster Roll Start/End Number
- Attendance Date
- Digest

Click **Download Attendance Data** to generate and download the Excel files.

---

## Process Flow

1. **User Input:**
   - User provides Panchayath details, date, work code, muster roll range, and digest (via CLI or Streamlit UI).
2. **Web Scraping:**
   - The system constructs the appropriate URLs and scrapes the MGNREGA portal for attendance data and images.
3. **Data Extraction:**
   - Attendance data and images are parsed and organized.
4. **Excel Generation:**
   - Data is saved into one or more Excel files, with images embedded where available.
5. **Download:**
   - User can download the generated files (via Streamlit) or find them saved locally (via CLI).

---

## Notes
- The scripts are tailored for Karnataka, Ballari district, Siruguppa block, but can be adapted for other regions with minor changes.
- Ensure you have a stable internet connection, as the scripts rely on live web scraping.
- For large numbers of muster rolls, parallel fetching is used to speed up the process.

---

## Troubleshooting
- If you encounter errors about missing columns or no muster rolls found, check the Panchayath name and date inputs.
- If execution is slow, ensure your internet connection is stable. The code is optimized for parallel fetching, but network speed is still a factor.
- For any issues, review the terminal output for error messages.

---

## License
This project is for educational and automation purposes. Please respect the terms of use of the MGNREGA portal.