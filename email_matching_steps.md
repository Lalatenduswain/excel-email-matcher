# Email Matching - Steps and Script

## Task
Match Employee IDs from Sheet1 (source) to the Employee Template (target) and copy Email IDs into Column H of the Employee Template.

---

## Steps Performed

### Step 1: Analyzed the Excel File Structure
- Opened the Excel file and found 3 sheets: **Employee Template**, **Sheet1**, **Lists**
- **Sheet1**: Contains Employee IDs (Column A) and Email IDs (Column B), no header row
- **Employee Template**: Contains employee data with 14 columns. Column B = Employee ID, Column H = Email Address (empty)

### Step 2: Checked ID Formats
- Verified that Employee IDs in Sheet1 match the format used in Employee Template
- Confirmed which IDs from Sheet1 exist in the Employee Template

### Step 3: Matched and Filled Emails
- Built a lookup dictionary from Sheet1 (Employee ID -> Email)
- Looped through all rows in Employee Template
- For each row, checked if the Employee ID exists in the lookup
- If found, wrote the email into Column H
- Saved the file

### Step 4: Cross-Verified Results
- Re-read the saved file and verified all filled emails match Sheet1 exactly
- 0 wrong emails, 0 missing emails
- Ran random spot-checks - all correct
- Any IDs from Sheet1 not found in the Employee Template are skipped

---

## Sample Result (using `sample_employee_data.xlsx`)
| Metric                          | Count |
|---------------------------------|-------|
| Total records in Sheet1         | 20    |
| Total records in Employee Template | 25 |
| Successfully matched and filled | 20    |
| Not found in template (skipped) | 0     |

---

## Python Script Used

```python
import openpyxl

# Load the workbook
wb = openpyxl.load_workbook('sample_employee_data.xlsx')
sheet1 = wb['Sheet1']
template = wb['Employee Template']

# Step 1: Build a lookup dictionary from Sheet1 (Employee ID -> Email)
email_lookup = {}
for row in sheet1.iter_rows(min_row=1, max_row=sheet1.max_row, values_only=True):
    emp_id = str(row[0]).strip() if row[0] else ''
    email = row[1]
    if emp_id and email:
        email_lookup[emp_id] = email

print(f'Email lookup entries: {len(email_lookup)}')

# Step 2: Match Employee IDs and fill Column H in Employee Template
matched = 0
for row_num in range(2, template.max_row + 1):
    emp_id_cell = template.cell(row=row_num, column=2)  # Column B = Employee ID
    emp_id = str(emp_id_cell.value).strip() if emp_id_cell.value else ''
    if emp_id in email_lookup:
        template.cell(row=row_num, column=8).value = email_lookup[emp_id]  # Column H = Email
        matched += 1

# Step 3: Check for unmatched IDs from Sheet1
template_ids = set()
for row in template.iter_rows(min_row=2, max_row=template.max_row, values_only=True):
    template_ids.add(str(row[1]).strip() if row[1] else '')

unmatched = [eid for eid in email_lookup if eid not in template_ids]

print(f'Matched and filled: {matched}')
print(f'Sheet1 IDs not found in template: {len(unmatched)}')
if unmatched:
    print('Unmatched IDs:', unmatched)

# Step 4: Save the file
wb.save('sample_employee_data.xlsx')
print('File saved successfully!')
```

---

## Verification Script (Optional)

```python
import openpyxl
import random

wb = openpyxl.load_workbook('sample_employee_data.xlsx')
sheet1 = wb['Sheet1']
template = wb['Employee Template']

# Build lookup from Sheet1
email_lookup = {}
for row in sheet1.iter_rows(min_row=1, max_row=sheet1.max_row, values_only=True):
    emp_id = str(row[0]).strip() if row[0] else ''
    email = row[1]
    if emp_id and email:
        email_lookup[emp_id] = email

# Verify all filled emails
filled_correct = 0
filled_wrong = 0
should_have_email_but_empty = 0

for row_num in range(2, template.max_row + 1):
    emp_id = str(template.cell(row=row_num, column=2).value).strip() if template.cell(row=row_num, column=2).value else ''
    email_in_template = template.cell(row=row_num, column=8).value

    if emp_id in email_lookup:
        expected_email = email_lookup[emp_id]
        if email_in_template == expected_email:
            filled_correct += 1
        elif email_in_template is None:
            should_have_email_but_empty += 1
        else:
            filled_wrong += 1

print(f'Correctly matched: {filled_correct}')
print(f'Wrong emails: {filled_wrong}')
print(f'Missing emails: {should_have_email_but_empty}')

# Spot-check 10 random entries
matched_rows = []
for row_num in range(2, template.max_row + 1):
    emp_id = str(template.cell(row=row_num, column=2).value).strip() if template.cell(row=row_num, column=2).value else ''
    if emp_id in email_lookup:
        matched_rows.append(row_num)

samples = random.sample(matched_rows, min(10, len(matched_rows)))
print('\nSpot-check:')
for row_num in sorted(samples):
    emp_id = str(template.cell(row=row_num, column=2).value).strip()
    name = f'{template.cell(row=row_num, column=3).value} {template.cell(row=row_num, column=4).value}'
    email = template.cell(row=row_num, column=8).value
    status = 'OK' if email == email_lookup[emp_id] else 'MISMATCH'
    print(f'  Row {row_num}: ID={emp_id}, Name={name}, Email={email} [{status}]')
```
