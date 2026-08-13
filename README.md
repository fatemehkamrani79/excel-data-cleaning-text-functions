# Excel Data Cleaning with Text Functions

This project is based on Excel Exercise 02 from the Data Analyst training program at Tose'e Institute, taught by Dr. Majid Eyvazian.

The exercise focuses on cleaning and transforming names, phone numbers, addresses, zip codes, and dates using Excel formulas. After completing the required tasks, I reorganized the workbook, added validation checks, and tested a few alternative formula methods.

---

## Assignment Information

- **Institute:** Tose'e Institute
- **Instructor:** Dr. Majid Eyvazian
- **Course:** Data Analyst Training Program
- **Exercise:** Excel Exercise 02
- **Topic:** Data Cleaning with Excel Formulas

The original assignment is included in the `docs` folder.

---

## Original Assignment

The exercise required the following tasks:

1. Combine first name and last name into one column with a single space between them.
2. Extract the three-digit area code from the `Phone` field.
3. Extract the seven-digit phone number without the area code.
4. Remove parentheses and spaces from the `Phone` field.
5. Extract the street name from `Street Address` without the house number.
6. Combine the zipcode and street name in this format:

   `Zipcode - Street Name`

7. Create an ID using:
   - The first three letters of the first name
   - The first three letters of the last name
   - The phone area code

8. Replace `St.` with `Street` in the street address.
9. Extract the house number from the street address.
10. Extract the year, month, and day from the request date.
11. Extract the year from the request date.

---

## My Approach

I kept the original data separate from the cleaned output.

The `Raw_Data` sheet contains only the original fields, while the `Cleaned_Data` sheet contains the source data together with the columns created for the exercise.

I used Excel Tables and structured references for the main formulas. I also added a validation sheet to check the results across all 200 records.

For a few tasks, I tested a second formula approach using newer Excel text functions.

---

## Source Data

The dataset contains 200 records and six original fields:

| Field | Type | Description |
|---|---|---|
| First Name | Text | First name |
| Last Name | Text | Last name |
| Phone | Text | Phone number with area code |
| Street Address | Text | House number and street name |
| Zipcode | Text | Postal code |
| Req Date | Date | Request date |

Phone numbers and zip codes are stored as text. This keeps their original format and preserves leading zeros in zip codes.

---

## Created Fields

| Field | Purpose |
|---|---|
| Full Name | Combines first and last name |
| Area Code | Extracts the three-digit area code |
| Phone Number | Extracts the seven-digit phone number |
| Clean Phone | Removes parentheses and spaces |
| Street Name | Extracts the street name |
| Zipcode - Street Name | Combines zipcode and street name |
| Unique ID | Combines name characters and area code |
| Standardized Address | Replaces `St.` with `Street` |
| House Number | Extracts the house number |
| Year | Extracts the year from the request date |
| Month | Extracts the month |
| Day | Extracts the day |

The same `Year` column is used for Questions 10 and 11 instead of creating the same result twice.

---

## Example

Here is one record after cleaning:

| Field | Example |
|---|---|
| First Name | Naomi |
| Last Name | Rev |
| Full Name | Naomi Rev |
| Phone | (383) 835-9398 |
| Area Code | 383 |
| Phone Number | 8359398 |
| Clean Phone | 383835-9398 |
| Street Address | 41839 Franti Drive |
| Street Name | Franti Drive |
| Zipcode | 03249 |
| Zipcode - Street Name | 03249 - Franti Drive |
| Unique ID | NaoRev383 |
| House Number | 41839 |
| Year | 2014 |
| Month | 7 |
| Day | 19 |

---

## Excel Functions Used

The main solution uses:

- `TRIM`
- `LEFT`
- `RIGHT`
- `MID`
- `FIND`
- `SUBSTITUTE`
- `YEAR`
- `MONTH`
- `DAY`

For selected alternative methods, I also used:

- `TEXTJOIN`
- `TEXTAFTER`
- `TEXTBEFORE`

---

## Alternative Methods

I compared a second formula approach for a few selected tasks:

- Full Name
- Area Code
- Phone Number
- Street Name
- House Number

For example, the main Street Name formula uses `MID` and `FIND`, while the alternative version uses `TEXTAFTER`.

This section is mainly there to compare older and newer Excel text functions on the same cleaning tasks.

---

## Validation

The `Validation` sheet checks:

- Raw and cleaned record counts
- Missing source values
- Full Name
- Area Code
- Phone Number
- Clean Phone
- Street Name
- Zipcode - Street Name
- Unique ID formula
- Duplicate Unique IDs
- Standardized Address
- House Number
- Year, Month, and Day
- Formula errors

Each check returns either `PASS` or `CHECK`.

All checks in the final workbook currently return `PASS`.

---

## Data Quality Notes

- No missing values were found in the required source fields.
- Zip codes were kept as text to preserve leading zeros.
- Phone numbers were kept as text because they contain formatting characters.
- House numbers were treated as text because address identifiers are not always purely numeric.
- The original source fields were kept unchanged.
- Generated IDs were checked for duplicates.

---

## Workbook Structure

The workbook contains five sheets:

### Raw_Data
Original source fields only.

### Cleaned_Data
Main solution with the cleaned and extracted fields.

### Alternative_Methods
Alternative formulas for selected tasks.

### Validation
Checks the cleaning results and flags mismatches.

### Documentation
Project notes, cleaning rules, methods, and workbook structure.

---

## Project Files

- [Excel Workbook](./Fatemeh_Kamrani_Excel_Data_Cleaning.xlsx)
- [Original Assignment](./docs/Excel_Exercise_02_Assignment.pdf)

```text
excel-data-cleaning-text-functions/
│
├── README.md
├── Fatemeh_Kamrani_Excel_Data_Cleaning.xlsx
└── docs/
    └── Excel_Exercise_02_Assignment.pdf
