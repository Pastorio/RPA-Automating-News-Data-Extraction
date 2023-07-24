# RPA-Automating-News-Data-Extraction
## Overview
This repository contains the RPA (Robotic Process Automation) of the process of extracting data from the New York Times website. The extracted data is then stored in an Excel file, including the title, date, description, picture filename, count of search phrases in the title and description, and whether the title or description contains any amount of money.

## Details

- **Configured Variables:**
  1. Search phrase
  2. News category or section
  3. Number of months for which news is to be received

## Solution Details

The solution is implemented in Python using the RPA Framework's Selenium library. The main steps of the solution are as follows:

1. Open the New York Times website using Selenium.
2. Accept the cookies.
3. Perform a search based on the given search phrase (if provided).
4. Apply the filter based on the specified news category (if provided).
5. Extract the news data, including title, date, description, and picture filename.
6. Count the occurrences of the search phrase in the title and description.
7. Check if the title or description contains any amount of money in the specified formats.
8. Store the extracted data in an Excel file.
9. Close the browser.

## Requirements

To run this solution locally, you need the following:

- Python 3.x
- RPA Framework

## How to Use

1. Clone the repository to your local machine.
2. Install the required dependencies using pip, Robocorp Vscode Extension or Robocorp Automation Studio following the [link](https://robocorp.com/developers) 

```bash
pip install rpaframework openpyxl
```

3. Modify the `work-items.json` file with the desired values for the search phrase, news category, and number of months.
4. Run the Python script to execute the RPA bot or the [Control Room](https://robocorp.com/docs/control-room):

```bash
python main.py
```

5. After the script finishes running, you will find the extracted news data in the `news_data.xlsx` file.
## Disclaimer

This solution is a demonstration of how to automate a specific task using RPA. Always make sure to comply with the terms of use of any website you are automating and respect the website's policies on data scraping.

---