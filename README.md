
# Excel Merger App

## Overview
The Excel Merger App is a Streamlit-based web application designed to merge data from multiple Excel files into a single main Excel file. The app maintains the original structure of the main Excel file, merging specified columns from other Excel files based on a common ID column.

## Features
- Upload a main Excel file with multiple worksheets.
- Upload additional Excel files to merge data from.
- Specify the common ID column for merging.
- Specify the columns to merge from the additional files.
- Retain the structure of the main Excel file while merging data.
- Download the merged Excel file.

## Installation

### Prerequisites
- Python 3.7 or higher
- `pip` package manager

### Step-by-Step Instructions

1. **Clone the Repository**
   ```bash
   git clone https://github.com/yourusername/excel-merger-app.git
   cd excel-merger-app
   ```

2. **Install Dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the App**
   ```bash
   streamlit run app.py
   ```

## Usage

1. **Upload Main Excel File**
   - Click on the "Choose the main Excel file" button.
   - Select your main Excel file (.xlsx) which contains multiple worksheets.

2. **Upload Other Excel Files**
   - Click on the "Choose other Excel files" button.
   - Select one or more additional Excel files (.xlsx) to merge data from.

3. **Specify the ID Column**
   - Enter the name of the common ID column used for merging data across the files.

4. **Specify Columns to Merge**
   - Enter the names of the columns to merge from the additional files, separated by commas.

5. **Merge and Download**
   - Click the "Merge Excel Files" button.
   - The app will display the merged data for each worksheet.
   - A "Download Merged Excel File" button will appear to download the merged Excel file.

## Example

1. **Main Excel File (main.xlsx)**:
   - Worksheet: `Sheet1`
     | ID  | Name    | Age |
     | --- | ------- | --- |
     | 1   | Alice   | 30  |
     | 2   | Bob     | 25  |
   - Worksheet: `Sheet2`
     | ID  | Product | Price |
     | --- | ------- | ----- |
     | 1   | Laptop  | 1000  |
     | 2   | Phone   | 500   |

2. **Additional Excel File (additional.xlsx)**:
   - Worksheet: `Sheet1`
     | ID  | Address    | Salary |
     | --- | ---------- | ------ |
     | 1   | Address 1  | 70000  |
     | 2   | Address 2  | 60000  |

3. **Merged Result**:
   - Worksheet: `Sheet1`
     | ID  | Name    | Age | Address    | Salary |
     | --- | ------- | --- | ---------- | ------ |
     | 1   | Alice   | 30  | Address 1  | 70000  |
     | 2   | Bob     | 25  | Address 2  | 60000  |
   - Worksheet: `Sheet2`
     | ID  | Product | Price |
     | --- | ------- | ----- |
     | 1   | Laptop  | 1000  |
     | 2   | Phone   | 500   |

## Contributing
Contributions are welcome! Please feel free to submit a Pull Request.

