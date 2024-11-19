# CLAI Setup Instructions

## Overview

CLAI is a Python-based tool that extracts job details from a given URL, processes the information, and generates professional documents tailored to the job posting. It streamlines the process of creating personalized content using automation and AI.

## Prerequisites

- Python 3.6 or higher
- LibreOffice (for PDF conversion)
- Chrome WebDriver (for Selenium)
- Required Python packages

## Setup Instructions

### Step 1: Install Python

Make sure you have Python 3.6 or higher installed. You can download it from the official [Python website](https://www.python.org/downloads/).

### Step 2: Install LibreOffice

Download and install LibreOffice from the official [LibreOffice website](https://www.libreoffice.org/download/download/). Ensure it is installed in the default location to avoid path issues.

- I opted for LibreOffice because it doesn't require validation for each DOCX to PDF conversion on macOS.
- If you're on Windows, you can use the “docx2pdf” package as an alternative to installing LibreOffice.

### Step 3: Install Chrome and ChromeDriver

1. ChromeDriver should install automatically when the script first runs. If it doesn't, follow the steps below:
2. Download Google Chrome from the [Chrome website](https://www.google.com/chrome/).
3. Download the appropriate version of ChromeDriver from [ChromeDriver downloads](https://chromedriver.chromium.org/downloads) that matches your Chrome version.
4. Extract the downloaded file and place it in a known directory, preferably in your system's PATH.

### Step 4: Install Required Python Packages

Use pip to install the necessary packages:

- `docxtpl`
- `tqdm`
- `selenium`
- `pyperclip`
- `google-generativeai`
- `urllib3 (v1.26.16)`

```bash
pip install docxtpl tqdm selenium pyperclip google-generativeai urllib3==1.26.16
```

### Step 5: Download the Template

Ensure you have a Word template named `Template.docx` in the same directory as the script. This template will serve as the foundation for generating customized documents.

- This document must include the following variables (you can change their name in the code):
  - `{{ today_date }}`
  - `{{ company_name }}`
  - `{{ a }}`
  - `{{ position_name }}`
  - `{{ generate }}`
  - `{{ glazing }}`

These placeholders will be dynamically replaced during the generation process.

### Step 6: Get a Google Gemini API Key

To use the Google Gemini API for content generation, you will need an API key. Follow these steps

1. Visit the [Google Gemini API documentation](https://ai.google.dev/gemini-api/docs/api-key).
2. Follow the instructions to sign up for access and generate an API key.
3. Copy your API key and add it to the script.

### Step 7: Run CLAI

1. Navigate to the directory where your script is located.
2. Run the script using the following command or in your preferred IDE.

```bash
python main.py
```

### Step 8: Provide Input

- When prompted, enter a valid job posting URL.
- Follow the prompts to manually input any information if extraction fails.

### Step 9: Output

CLAI will generate a customized document saved as both a DOCX and PDF file in the same directory. The content will also be copied to your clipboard for quick reference.

- The MultiGenerator feature allows batch processing by using a text file containing a list of URLs. CLAI will iterate through the list to generate all required documents.

## Troubleshooting

- If you encounter issues with WebDriver, ensure your Chrome and ChromeDriver versions match.
- For LibreOffice conversion errors, verify that the `LIBREOFFICE_BINARY` path is correct in the script (update it as needed for your OS).

## Conclusion

CLAI is now ready to use for generating personalized documents based on job descriptions. If you have any questions or run into issues, refer to the documentation for the specific tools used or reach out for assistance.
