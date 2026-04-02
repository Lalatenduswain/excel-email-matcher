# Excel Email Matcher

A Python script that matches Employee IDs across two Excel sheets and fills in missing Email addresses.

## Problem
You have an Excel file with:
- **Sheet1** (source): Contains Employee IDs and their Email IDs
- **Employee Template** (target): Contains employee records but the Email Address column (Column H) is empty

The script matches Employee IDs from both sheets and copies the Email IDs into the template.

## How to Use

1. Install the dependency:
   ```
   pip install -r requirements.txt
   ```

2. Run the script:
   ```
   python email_matcher.py
   ```

The script uses `sample_employee_data.xlsx` (included) as a demo. To use your own file, update the filename in `email_matcher.py`.

## Files
| File | Description |
|------|-------------|
| `email_matcher.py` | Main Python script |
| `sample_employee_data.xlsx` | Sample Excel file with dummy data (25 employees, 20 emails) |
| `email_matching_steps.md` | Detailed steps and verification script |
| `requirements.txt` | Python dependencies |

## How It Works
1. Reads Sheet1 and builds a lookup dictionary (Employee ID -> Email)
2. Loops through each row in the Employee Template
3. If the Employee ID matches, writes the email into Column H
4. Reports matched, unmatched, and saves the file
