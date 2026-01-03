# Data Mappings - BIS to Schedule Field Translations

## Overview

This document catalogs all data transformations from BIS (Benefits Information System) delimited files to QC schedule workbooks. Each program has specific field mappings and code translations.

## BIS File Structure

### Case Worksheet Columns

Key columns in the Case worksheet:

| Column | Field Name | Format | Example |
|--------|------------|--------|---------|
| A | State Code | Text | "42" (Pennsylvania) |
| B | County Code | Text | "10" |
| C | Review Number | Numeric | 50123 |
| J | Case Number | Text | "A12345678" |
| M | Application Date | YYYYMMDD | 20231015 |
| W | Number of Individuals | Numeric | 3 |
| X-AC | Certification Dates | YYYY/MM/DD split | 2023, 10, 01 |
| AH | State of Residence | Text | "PA" |
| AJ-AB | Contact Information | Text | Phone, address |
| AK | Benefit Amount | Currency | 250.00 |
| AU | Rent Amount | Currency | 800.00 |
| AW | LIHEAP Received | Y/N | "Y" |
| AX | Utility Type | Code | "H" (heating) |

### Individual Worksheet Columns

Key columns in the Individual worksheet:

| Column | Field Name | Format | Example |
|--------|------------|--------|---------|
| C | Review Number | Numeric | 50123 |
| J | Individual Category | Code | "01" |
| L | Line Number | Numeric | 1-99 |
| N | First Name | Text | "John" |
| O | Middle Initial | Text | "M" |
| P | Last Name | Text | "Smith" |
| Q | Suffix | Text | "Jr" |
| R | Birth Date | YYYYMMDD | 19850315 |
| S | Age | Numeric | 38 |
| T | Relationship Code | Text | "X  " (head) |
| U | Gender | M/F | "M" |
| V | Race Code | Numeric | 1-8 |
| W | Hispanic Indicator | Numeric | 1, 2 |
| X | Relationship to Payee | Text | "X  " |
| Y | Eligibility Status | Text | "EM", "EB", "NM" |
| Z | SSN | Text | "123456789" |
| AA | Education Level | Numeric | 12 |
| AB | School Code | Numeric | 16, 17 |
| AC | Student Status | Numeric | 21-24 |
| AD | TANF Eligibility | Text | "ES", "EC" |
| AE | ETP Code | Numeric | 1-40 |
| AF | Citizenship | Numeric | 1-6 |
| AN | Something Status | Numeric | 1-4 |
| AO | Another Code | Numeric | Various |
| AP | Citizenship (alt) | Numeric | 1-6 |

**Income Columns (farther right):**
- BQ-BR: Earned income frequency/amount
- BS: Self-employment income
- BW-BX: RSDI frequency/amount
- BY-BZ: VA benefits frequency/amount
- CA-CB: SSI frequency/amount
- CC-CD: Unemployment frequency/amount
- CE-CF: Workers Comp frequency/amount
- CG-CH: Deemed income frequency/amount
- CI-CJ: Public assistance frequency/amount
- CK-CL: Educational grants frequency/amount

## Code Translation Tables

### Relationship to Head of Household

#### TANF/MA Format

```vb
Select Case UCase(wb_bis.Range("X" & i))  ' BIS Column
    Case "X  "  ' Head of Household
        If age <= 19 Then
            schedule_code = "02"  ' Minor head
        Else
            schedule_code = "01"  ' Adult head
        End If
    Case "W  ", "H  ", "CLH", "CLW"  ' Spouse variations
        If age <= 19 Then
            schedule_code = "04"  ' Minor spouse
        Else
            schedule_code = "03"  ' Adult spouse
        End If
    Case "F  ", "M  ", "SF ", "SM "  ' Father, Mother, Stepfather, Stepmother
        schedule_code = "05"  ' Parent
    Case "D  ", "S  "  ' Daughter, Son
        schedule_code = "06"  ' Child
    Case "SS ", "SD "  ' Stepson, Stepdaughter
        schedule_code = "07"  ' Stepchild
    Case "NR "  ' Non-Relative
        schedule_code = "13"  ' TANF uses 13
        ' schedule_code = "20"  ' MA uses 20!
    Case "GD ", "GS ", "GGS", "GGD"  ' Grandchild, Great-grandchild
        schedule_code = "10"  ' Grandchild
    Case "A  ", "B  ", "C  ", "GN ", "GNC", "N  ", "NC ", "SR "  ' Other relatives
        schedule_code = "11"  ' TANF other
        ' schedule_code = "14"  ' MA uses 14!
End Select
```

**Key Differences:**
- **TANF:** Unrelated = 13, Other = 11
- **MA:** Unrelated = 20, Other = 14

#### SNAP Format (Simpler)

```vb
Select Case UCase(wb_bis.Range("T" & i))  ' Uses column T, not X
    Case "X  "
        schedule_code = 1  ' Head
    Case "W  ", "H  ", "CLH", "CLW"
        schedule_code = 2  ' Spouse
    Case "F  ", "M  ", "SF ", "SM "
        schedule_code = 3  ' Parent
    Case "D  ", "SD ", "S  ", "SS "  ' Child or Stepchild
        schedule_code = 4  ' Child
    Case "NR "
        schedule_code = 6  ' Unrelated
    Case Else
        schedule_code = 5  ' Other related
End Select
```

**Note:** SNAP uses single-digit codes

### Race Codes

```vb
Select Case wb_bis.Range("V" & i)  ' BIS Race Code
    Case 1  ' White
        tanf_code = 2
        ma_code = 2
        snap_code = "05"
    Case 3  ' Black/African American
        tanf_code = 5
        ma_code = 5
        snap_code = "03"
    Case 4  ' Asian
        tanf_code = 4
        ma_code = 4
        snap_code = "04"
    Case 5  ' Native American
        tanf_code = 1
        ma_code = 1
        snap_code = "07"
    Case 6  ' Pacific Islander
        tanf_code = 12  ' TANF-specific code
        ma_code = 9
        snap_code = "12"
    Case 7  ' Other
        tanf_code = 6  ' TANF-specific code
        ma_code = 4
        snap_code = "06"
    Case 8  ' Hispanic or Unknown
        If wb_bis.Range("W" & i) = 2 Then  ' Hispanic indicator
            tanf_code = 3
            ma_code = 3
            snap_code = "99"  ' Hispanic
        Else
            tanf_code = 9
            ma_code = 9
            snap_code = "99"  ' Unknown
        End If
End Select
```

**Format Differences:**
- **TANF/MA:** Single-digit numeric (1-9, 12)
- **SNAP:** Zero-padded string ("01"-"99")

### Citizenship Codes

```vb
Select Case wb_bis.Range("AF" & i)  ' or Range("AO" & i) for some programs
    Case 1  ' US Born/Naturalized Citizen
        tanf_code = 1
        ma_code = "01"
        snap_code = "01"
    Case 2  ' Permanent Resident Alien
        tanf_code = 2
        ma_code = 2  ' MA doesn't populate most alien codes
        snap_code = 2   ' SNAP doesn't populate most alien codes
    Case 3  ' Temporary Alien
        tanf_code = 3
        ' Not typically populated in MA/SNAP
    Case 4  ' Refugee
        tanf_code = 4
        ma_code = "05"
        snap_code = "05"
    Case 5  ' Illegal Alien
        tanf_code = 6
        ' Not populated in MA/SNAP
    Case 6  ' Unaccompanied Refugee Minor
        tanf_code = 5
        ' Not typically used
End Select
```

**Policy Note:** Only certain citizenship types are commonly populated in schedules due to eligibility requirements.

### Gender Codes

```vb
' BIS uses "M" or "F"
If wb_bis.Range("U" & i) = "F" Or wb_bis.Range("CO" & i) = "F" Then
    tanf_code = 2
    ma_code = "02"
    snap_code = "02"
ElseIf wb_bis.Range("U" & i) = "M" Or wb_bis.Range("CO" & i) = "M" Then
    tanf_code = 1
    ma_code = "01"
    snap_code = "01"
Else
    ' Leave blank
End If
```

**Note:** Some BIS files use column U, others use column CO for gender

### Marital Status Codes (TANF)

```vb
Select Case wb_bis.Range("Y" & i)  ' BIS Marital Status
    Case 1  ' Single
        schedule_code = 1
    Case 2  ' Married & Living with spouse
        schedule_code = 2
    Case 4  ' Separated-Married-Not living with spouse
        schedule_code = 3
    Case 5  ' Divorced
        schedule_code = 5
    Case 6  ' Widow/Widower
        schedule_code = 4
    Case 3, 7  ' Common Law/Divorced with spousal support
        schedule_code = 9
End Select
```

### Education Level Codes

**Special Transformations:**

```vb
' Format to 2 digits
education_code = Format(wb_bis.Range("AA" & i), "00")

' Special code translations
If Val(education_code) = 98 Then
    education_code = "00"  ' Unknown
End If

If Val(education_code) = 16 Then
    education_code = "14"  ' Graduate degree (consolidated)
End If
```

**Valid Ranges:**
- 00 = Unknown/No formal education
- 01-12 = Grades 1-12
- 13-16 = College years
- 17-19 = Graduate education
- 98 = Not in school (converted to 00)
- 99 = Other

### Eligibility Status Codes

Used to determine household composition:

| BIS Code | Meaning | SNAP Participant? | TANF Participant? |
|----------|---------|-------------------|-------------------|
| **EM** | Eligible Member | Yes | Maybe |
| **EB** | Eligible ABAWD | Yes | No |
| **EW** | Eligible Waiver | Yes | No |
| **ES** | Eligible TANF Standard | No | Yes |
| **EC** | Eligible TANF Child Only | No | Yes |
| **NM** | Not Member | No | No |
| **DS** | Disqualified Sanction | No | No |
| **DF** | Disqualified Fraud | No | No |

**Usage in Code:**

```vb
' SNAP participation check
If Left(wb_bis.Range("Y" & i), 1) = "E" Then
    ' Eligible for SNAP
    wb_sch.Range("E" & j) = "01"  ' Participation indicator
End If

' TANF participation check
If wb_bis.Range("AD" & i) = "ES" Or wb_bis.Range("AD" & i) = "EC" Then
    wb_sch.Range("AI" & write_row) = "Yes"
Else
    wb_sch.Range("AI" & write_row) = "No"
End If
```

### ETP (Employment and Training) Codes

Used for SNAP work requirement exemptions:

| Code Range | Category | Schedule Location |
|------------|----------|-------------------|
| 30, 40 | Mandatory | Element 160 - "mandatory" |
| 1, 2 | Age (under 16, over 59) | Element 160 - "age" |
| 3 | Health/Disability | Element 160 - "health" |
| 13, 20 | Student | Element 160 - "student" |
| 17 | Employment (30+ hours) | Element 160 - "employment" |
| 4, 6, 10, 14-21 | Exempt ABAWD | Element 161 |
| Other | Other exemptions | Element 160 - "other" |

**Collection Logic:**

```vb
Select Case wb_bis.Range("AE" & i)
    Case 30, 40
        If Len(mandatory) > 0 Then
            mandatory = mandatory & ", " & LineNumberFormatted
        Else
            mandatory = LineNumberFormatted
        End If
    Case 1, 2
        Age = Age & ", " & LineNumberFormatted
    ' ... etc
End Select

' Write to schedule
wb_sch.Range("G267") = mandatory
wb_sch.Range("L268") = "mandatory"
wb_sch.Range("G269") = Age
wb_sch.Range("L270") = "age"
' ... etc
```

## Income Frequency Conversion

BIS stores income frequencies as codes; conversion needed for monthly amounts:

```vb
income_freq = Array(0, 1, 4, 2, 2, 1, 0.5, 0.333333, 0.166667, 0.083333)

' Index = BIS frequency code
' Value = Multiplier to convert to monthly
```

| Index | BIS Frequency | Multiplier | Calculation |
|-------|---------------|------------|-------------|
| 0 | None/Not applicable | 0 | amount × 0 = 0 |
| 1 | Monthly | 1 | amount × 1 |
| 2 | Semi-monthly (2x/month) | 2 | amount × 2 |
| 3 | Bi-weekly (26/year) | 2 | amount × 2 (approx) |
| 4 | Weekly | 4 | amount × 4 (approx) |
| 5 | Hourly | 1 | amount × 1 (pre-calculated?) |
| 6 | Quarterly | 0.333333 | amount / 3 |
| 7 | Semi-annually | 0.166667 | amount / 6 |
| 8 | Annually | 0.083333 | amount / 12 |

**Usage Example:**

```vb
' Read frequency code and amount from BIS
freq_code = wb_bis.Range("BQ" & i)  ' e.g., 4 (weekly)
amount = wb_bis.Range("BR" & i)      ' e.g., 300

' Calculate monthly amount
monthly_income = Round(income_freq(freq_code) * amount, 0)
' Result: Round(4 * 300, 0) = 1200
```

### Income Type Codes

Common income types populated in computation sheets:

| Code | Type | BIS Column | Used In |
|------|------|------------|---------|
| **11** | Earned Income - Regular | BR | All programs |
| **12** | Earned Income - Self Employment | BS | All programs |
| **14** | Earned Income - Other | BT | All programs |
| **16** | Earned Income - Training | BV | All programs |
| **31** | RSDI (Social Security Disability) | BX | All programs |
| **32** | VA (Veterans Benefits) | BZ | All programs |
| **33** | SSI (Supplemental Security Income) | CB | All programs |
| **34** | Unemployment Compensation | CD | All programs |
| **35** | Workers Compensation | CF | All programs |
| **43** | Deemed Income | CH | All programs |
| **44** | Public Assistance | CJ | All programs |
| **45** | Educational Grants | CL | All programs |
| **47** | TANF Cash Payment | Special | TANF only |
| **50** | Child Support (consolidated) | CV+CX+CZ | All programs |

**Child Support Special Logic:**

```vb
' Combines three types of child support into one income entry
monthly_income = Round(income_freq(Range("CU" & i)) * Range("CV" & i), 0) + _
                 Round(income_freq(Range("CW" & i)) * Range("CX" & i), 0) + _
                 Round(income_freq(Range("CY" & i)) * Range("CZ" & i), 0)
```

## Element Code System

"Elements" are data validation checkpoints tracked in schedules. They are numbered sequentially in QCMIS documentation.

### Common Elements (All Programs)

| Element | Description | Validation |
|---------|-------------|------------|
| **110** | Birth dates present for all | All individuals have non-zero birth date |
| **111** | School attendance (18+) | Track individuals 18+ in school |
| **130** | Citizenship status | Categorize as citizen/qualified/unqualified |
| **140** | State of residence | Must be PA for eligibility |
| **150** | Household membership | List EM, EB, EW, NM statuses |

### SNAP-Specific Elements

| Element | Description | Validation |
|---------|-------------|------------|
| **151** | No disqualified members | No DS or DF statuses |
| **160** | ETP exemptions | Categorize work requirement exemptions |
| **161** | ABAWD in household | Track able-bodied adults |
| **163** | ABAWD listing | List all ABAWD members |
| **165** | Student status tracking | Various student situations |
| **166** | Additional exemptions | Other exemption scenarios |
| **170** | Specific exemption | Code 24 situations |
| **311** | Certification period | Date range tracking |
| **331** | RSDI recipients | List line numbers |
| **332** | VA recipients | List line numbers |
| **333** | SSI recipients | List line numbers |
| **334** | UC recipients | List line numbers |
| **335** | WC recipients | List line numbers |
| **342** | Child support received | Yes/No indicator |
| **343** | Deemed income | List line numbers |

**Checkbox Population Example:**

```vb
' Element 151 - Initialize to "Yes" (no disqualified)
wb_sch.Worksheets("FS Workbook").Shapes("CB 104").OLEFormat.Object.Value = 1

' Check for disqualified members
If wb_bis.Range("Y" & i) = "DS" Or wb_bis.Range("Y" & i) = "DF" Then
    wb_sch.Worksheets("FS Workbook").Shapes("CB 104").OLEFormat.Object.Value = 0
End If
```

## Utility Type Codes

```vb
Select Case wb_bis.Range("AX" & case_row.Row)
    Case "N", "U"
        utility_text = "Non Heating"
    Case "H"
        utility_text = "Heating"
    Case "L"
        utility_text = "Limited"
    Case "P"
        utility_text = "Telephone"
End Select

wb_sch.Range("O1489") = utility_text
```

**Impact on Benefits:** Heating utilities receive higher standard deduction

## Date Field Conversions

### BIS Date Format: YYYYMMDD

**Conversion Function:**

```vb
temp = wb_bis.Range("R" & i)  ' e.g., "20231015"
birth_date = DateSerial(Val(Left(temp, 4)), _      ' Year: 2023
                        Val(Mid(temp, 5, 2)), _     ' Month: 10
                        Val(Right(temp, 2)))        ' Day: 15
```

### BIS Date Format: Split Columns (Y/M/D)

Some dates are split across multiple columns:

```vb
' Recertification From Date
recert_from_date = DateSerial( _
    2000 + Val(wb_bis.Range("X" & case_row.Row)), _  ' Year (2-digit + 2000)
    Val(wb_bis.Range("Y" & case_row.Row)), _         ' Month
    1)                                                ' First of month

' Recertification Thru Date
recert_thru_date = DateSerial( _
    2000 + Val(wb_bis.Range("AA" & case_row.Row)), _ ' Year
    Val(wb_bis.Range("AB" & case_row.Row)), _        ' Month
    Val(wb_bis.Range("AC" & case_row.Row)))          ' Day
```

### Zero Date Handling

```vb
time_zero = DateSerial(0, 0, 1)  ' Represents missing date

If temp = 0 Then
    appl_date = time_zero
Else
    appl_date = DateSerial(...)
End If

' Only populate if date is valid
If appl_date > time_zero Then
    wb_sch.Range("G20") = appl_date
End If
```

## Line Number Formatting

All person line numbers formatted to 2 digits:

```vb
Function lnf(linenum As Integer)
    lnf = WorksheetFunction.Text(linenum, "00")
End Function

' Usage
wb_sch.Range("K11") = lnf(1)  ' Returns "01"
wb_sch.Range("K12") = lnf(15) ' Returns "15"
```

**Consistency:** Ensures all line number references use same format

## Program-Specific Field Locations

### SNAP Positive

**Schedule Sheet (Review Number):**
- Section 1: Review header (rows 18-25)
- Section 2: Household list (rows 29-43, step 2)
- Section 3: Absent relatives (rows varies)
- Section 4: Person details (rows 89-122, step 3)
- Section 5: Income (rows 131-143, step 3)
- Section 7: Supplemental (rows 155-165)

**FS Workbook Sheet:**
- Review info (rows 16-25)
- Elements 110-170 (rows 100-300)
- Elements 311-343 (rows 900-1500)

**FS Computation Sheet:**
- Number of individuals (B11)
- Earned income (A8:B10)
- Unearned income (A18:B22)
- Deductions (various)

### TANF

**Schedule Sheet (Review Number):**
- Section I: Case header (row 10)
- Section II: Case details (rows 16-24)
- Section III: Household (rows 30-44, step 2)
- Section IV: Income by person (rows 50-56, step 2)
- Section V: Resources (rows 61-67, step 2)
- Section VI: Deductions (rows 72-82, step 2)
- Section VII: Vendor payments (row 85)

**TANF Workbook Sheet:**
- Additional case information
- Supervisor approval (G38, G41, G44)

**TANF Computation Sheet:**
- Similar to SNAP but TANF-specific

### MA Positive

**Schedule Sheet (Review Number):**
- Section I: Header (row 10)
- Section II: Case details (rows 16-41)
- Section III: Person details (rows 51-73, step 2)
- Section IV: Income (rows 78-84, step 2)
- Section V: Error findings (rows 96-112, step 2)

**MA Workbook Sheet:**
- Case narrative
- Supporting details
- Supervisor approval (F41, F43, F45)

### MA Negative

**Schedule Sheet (Review Number):**
- Section I: Header (row 15)
- Section II: Action details (rows 19-56)
- Simplified structure vs Positive

### SNAP Negative

**Schedule Sheet (Review Number):**
- Header (rows 16-24)
- Action information
- Minimal data entry

## Special Field Handling

### SSN Formatting

```vb
' BIS stores SSN as 9-digit string
ssn = wb_bis.Range("Z" & i)  ' "123456789"
wb_sch.Range("AE" & write_row) = ssn  ' Write as-is

' Some schedules format as ###-##-####
' But current code writes unformatted
```

### Phone Number

```vb
' BIS stores as 10-digit string
phone = wb_bis.Range("AJ" & case_row.Row)  ' "2155551234"
wb_sch.Range("D16") = phone  ' Write as-is

' Schedule uses custom formatting to display (215) 555-1234
```

### Case Number Handling (MA Negative)

```vb
' MA Negative may have leading "A"
case_num = wb_bis.Range("S" & 15)  ' e.g., "A12345678"

' Strip leading "A" if present
If Left(case_num, 1) = "A" Then
    case_num = Right(case_num, Len(case_num) - 1)
End If
```

### Review Number Formatting

```vb
' Remove leading zeros
reviewtxt = WorksheetFunction.Text(thissht.Range("E" & i), "#")

If Left(reviewtxt, 1) = "0" Then
    reviewtxt = Right(reviewtxt, Len(reviewtxt) - 1)
End If

' 050123 becomes 50123
```

## Validation Code Lookups

### Income Type Validation (SNAP)

Schedule contains valid income type list in range BB126:BB147:

```vb
' Check if income type is valid
flag = 0
For n = 126 To 147
    If thisws.Range("BB" & n) = inctype Then
        flag = 1
        Exit For
    End If
Next n

If flag = 0 Then
    MsgBox "Income type on row " & i & " is invalid. " & _
           "Please enter a valid income type"
End If
```

**Valid Types:** 11-50 (various income categories)

## Data Quality Checks During Population

### Birthdate Validation

```vb
birthdate = True  ' Assume all present

For i = bis_ind_start_row To bis_ind_stop_row
    If wb_bis.Range("R" & i) = "" Or Val(wb_bis.Range("R" & i)) = 0 Then
        birthdate = False  ' Found missing birthdate
    End If
Next i

' Set Element 110 checkbox
If birthdate = True Then
    wb_sch.Shapes("CB 25").OLEFormat.Object.Value = 1
End If
```

### LIHEAP (Low Income Home Energy Assistance)

```vb
If wb_bis.Range("AW" & case_row.Row) = "N" Then
    wb_sch.Range("L1499") = "No"
ElseIf wb_bis.Range("AW" & case_row.Row) = "Y" Then
    wb_sch.Range("L1499") = "Yes"
End If
```

## Mapping Summary by Program

### TANF Full Mapping

| BIS Source | Schedule Destination | Transformation |
|------------|---------------------|----------------|
| Case.AB (Phone) | TANF Workbook.D20 | Direct copy |
| Individual.L (Line #) | Schedule.J{row} | Format to 00 |
| Individual.N-Q (Name) | Schedule.L{row} | Concatenate with spaces |
| Individual.J (Category) | Schedule.AC{row} | Direct copy |
| Individual.R (DOB) | Schedule.V{row} | YYYYMMDD → Date |
| Individual.T (Age) | Schedule.Y{row} | Direct copy |
| Individual.X (Relationship) | Schedule.AA{row} | Code translation |
| Individual.Z (SSN) | Schedule.AE{row} | Direct copy |
| Individual.AD (TANF Status) | Schedule.AI{row} | Yes/No translation |
| Individual.V (Race) | Schedule.R{row} | Code translation |
| Individual.U (Gender) | Schedule.P{row} | M/F → 1/2 |
| Individual.AP (Citizenship) | Schedule.T{row} | Code translation |
| Individual.Y (Marital) | Schedule.AK{row} | Code translation |
| Individual.AA (Education) | Schedule.V{row} | Format 00, special rules |

### SNAP Positive Full Mapping

| BIS Source | Schedule Destination | Transformation |
|------------|---------------------|----------------|
| Case.AJ (Phone) | FS Workbook.D16 | Direct copy |
| Case.M (Application Date) | FS Workbook.G20 | YYYYMMDD → Date |
| Case.X-AC (Cert Period) | FS Workbook.H24-H25 | Split columns → Date |
| Case.AK (Benefit Amount) | FS Workbook.R42 | Direct copy |
| Case.AH (State) | Element 140 checkbox | "PA" → checked |
| Case.AU (Rent) | FS Workbook.N1462 | Direct copy |
| Case.AX (Utility Type) | FS Workbook.O1489 | Code → text |
| Case.AW (LIHEAP) | FS Workbook.L1499 | Y/N → Yes/No |
| Individual.L (Line #) | Schedule.B{row} | Format to 00 |
| Individual.N-Q (Name) | FS Workbook.M{row} | Concatenate |
| Individual.J (Category) | FS Workbook.W{row} | Direct copy |
| Individual.R (DOB) | FS Workbook.Y{row} | YYYYMMDD → Date |
| Individual.S (Age) | FS Workbook.AB{row} | Direct copy |
| Individual.T (Relationship) | Schedule.H{row} | Code translation |
| Individual.U (SSN) | FS Workbook.AF{row} | Direct copy |
| Individual.Y (Status) | FS Workbook.AJ{row} | Yes/No |
| Individual.CO (Gender) | Schedule.M{row} | M/F → 01/02 |
| Individual.CP (Race) | Schedule.P{row} | Code translation |
| Individual.AF (Citizenship) | Schedule.S{row} | Code translation |
| Individual.V (Education) | Schedule.V{row} | Format, special rules |
| Individual.AE (ETP) | Element 160 | Categorize by type |

### MA Positive Full Mapping

| BIS Source | Schedule Destination | Transformation |
|------------|---------------------|----------------|
| Case.AB (Phone) | MA Workbook.D20 | Direct copy |
| Case.AC-AE (Open Date) | MA Workbook.F25 | Split → Date |
| Case.AF-AH (Action Date) | MA Workbook.F27 | Split → Date |
| Individual.L (Line #) | Schedule.B{row} | Format to 00 |
| Individual.N-Q (Name) | MA Workbook.L{row} | Concatenate |
| Individual.J (Category) | MA Workbook.AC{row} | Direct copy |
| Individual.R (DOB) | MA Workbook.V{row} | YYYYMMDD → Date |
| Individual.T (Age) | MA Workbook.Y{row} | Direct copy |
| Individual.T (Relationship) | Schedule.N{row} | Code translation |
| Individual.Z (SSN) | MA Workbook.AE{row} | Direct copy |
| Individual.U (Gender) | Schedule.V{row} | M/F → 01/02 |
| Individual.V (Race) | Schedule.Y{row} | Code translation |
| Individual.AO (Citizenship) | Schedule.AB{row} | Code translation |

**Note:** MA uses different BIS columns than SNAP/TANF for some fields

## Absent Relatives (SNAP Only)

SNAP tracks absent parents/relatives for child support purposes:

```vb
writerowabs = 25  ' Starting row
For j = 74 To 134 Step 30  ' BIS columns for up to 3 absent relatives
    If Trim(wb_bis.Cells(case_row.Row, j)) <> "" Then
        writerowabs = writerowabs + 2
        
        ' Name
        wb_sch.Range("K" & writerowabs) = _
            wb_bis.Cells(case_row.Row, j + 1) & " " & _
            wb_bis.Cells(case_row.Row, j)
        
        ' Legal Responsible Relative info
        wb_sch.Range("S" & writerowabs) = _
            "LRR to " & wb_bis.Cells(case_row.Row, j + 9)
        
        ' SSN
        wb_sch.Range("V" & writerowabs) = _
            wb_bis.Cells(case_row.Row, j + 2)
        
        ' Address
        wb_sch.Range("Z" & writerowabs) = _
            wb_bis.Cells(case_row.Row, j + 3)
        wb_sch.Range("Z" & writerowabs + 1) = _
            wb_bis.Cells(case_row.Row, j + 5) & ", " & _
            wb_bis.Cells(case_row.Row, j + 6)
    Else
        Exit For  ' No more absent relatives
    End If
Next j
```

**Columns:**
- j+0: Last name
- j+1: First name
- j+2: SSN
- j+3: Street address
- j+5: City
- j+6: State
- j+9: LRR info

## Common Mapping Patterns

### Pattern 1: Direct Copy
```vb
schedule_cell = bis_cell
```

### Pattern 2: Date Conversion
```vb
schedule_cell = DateSerial(Year, Month, Day)
```

### Pattern 3: Code Translation
```vb
Select Case bis_cell
    Case value1: schedule_cell = mapped1
    Case value2: schedule_cell = mapped2
End Select
```

### Pattern 4: Concatenation
```vb
schedule_cell = bis_cell1 & " " & bis_cell2 & " " & bis_cell3
```

### Pattern 5: Format and Clean
```vb
schedule_cell = Format(bis_cell, "00")
If schedule_cell = "98" Then schedule_cell = "00"
```

### Pattern 6: Conditional Aggregation
```vb
result_string = ""
For each row
    If condition Then
        If Len(result_string) > 0 Then
            result_string = result_string & ", "
        End If
        result_string = result_string & value
    End If
Next
schedule_cell = result_string
```

---

**Next:** [Optimizations & Recommendations](08-optimizations-recommendations.md) - Performance improvements and refactoring suggestions

