
---

# HighlightAgreements Macro Documentation

## Table of Contents
- [Purpose](#purpose)
- [1. Setup and Initialization](#1-setup-and-initialization)
  - [Workbook and Worksheet Initialization](#workbook-and-worksheet-initialization)
  - [View Settings](#view-settings)
  - [Range Determination](#range-determination)
- [2. Data Comparison](#2-data-comparison)
  - [Data Processing for Each Column](#data-processing-for-each-column)
  - [Comparison Mechanics](#comparison-mechanics)
  - [Highlighting & Categorization](#highlighting--categorization)
  - [Data Storage](#data-storage)
- [3. Summary Report](#3-summary-report)
  - [Color-based Tally](#color-based-tally)
  - [Data Presentation](#data-presentation)
- [Step-by-Step Breakdown of `HighlightAgreements` Macro with Code Fragments](#step-by-step-breakdown-of-highlightagreements-macro-with-code-fragments)

---
# HighlightAgreements Macro Documentation

## Purpose:
The macro is designed to compare data in a specified range of cells between two Excel workbooks. Based on the comparison, it highlights disagreements, partial, or agreements and provides a summary.

---

### 1. Setup and Initialization

#### Workbook and Worksheet Initialization:
- The active workbook's "Sheet1" is referenced as `ws1` (worksheet with the data).
- A secondary workbook (specified in cell A2 of the active workbook) is opened, and its first sheet is denoted as `ws` (highlight button worksheet).

#### View Settings:
- Columns A to J are hidden, ensuring a more transparent view of the data of interest.

#### Range Determination:
- The range for analysis is determined from cells B2 (start column) and C2 (end column) of `ws1`.

---

### 2. Data Comparison

#### Data Processing for Each Column:
- For every column in the defined range:
  - Data from rows 3 and 4 is divided based on the comma `,`.
  - The divided data is consolidated and stored in `ArrMerged`.

#### Comparison Mechanics:
- Every element in `Arr1` and `Arr2` is compared to tally matched and unmatched items.

#### Highlighting & Categorization:
- Depending on comparison results:
  - **Partial Match:** Cells in rows 3 and 4 of the current column are highlighted in yellow, and a value of 2 is assigned to row 5.
  - **No Match:** Cells in rows 3 and 4 of the current column are highlighted in red, and a value of 3 is assigned to row 5.
  - **Full Match:** Cells in rows 3 and 4 of the current column are highlighted in green, and a value of 1 is assigned to row 5.

#### Data Storage:
- The results are captured in a dictionary (`dataDict`) for subsequent utilization.

---

### 3. Summary Report

#### Color-based Tally:
- The macro enumerates the overall instances of yellow, red, and green cells. These totals are displayed in cells L6, L7, and L8 of the `ws` sheet, respectively.

#### Data Presentation:
- A report is generated in columns K, L, and M of the `ws` sheet starting at row 6:
  - **Column K:** Contains the unique answers from `ArrMerged`.
  - **Column L:** Displays the total occurrences of each unique answer.
  - **Column M:** Illustrates the percentage representation for each category (Agree, Disagree, Partial) concerning each answer.

---

### Step-by-Step Breakdown of `HighlightAgreements` Macro with Code Fragments

---

1. **Initialization:**
    ```vba
    Dim wb As Workbook
    Dim ws, ws1 As Worksheet
    Dim totalYellow, totalRed, totalGreen As Long
    Dim dataDict As New Scripting.Dictionary

    Set ws1 = ActiveWorkbook.Worksheets("Sheet1")
    Set wb = Workbooks.Open(ActiveWorkbook.Path & "/" & Range("A2").Value)
    Set ws = wb.Sheets(1)
    Columns("A:J").Hidden = True
    ```
    - Variables are initialized, including workbooks, worksheets, dictionaries, and counters.
    - "Sheet1" of the active workbook is set as `ws1`.
    - A new workbook (from the path in cell A2) is opened, and its first worksheet is set as `ws`.
    - Columns A to J in the active workbook are hidden for clarity.

2. **Setting the Range for Analysis:**
    ```vba
    FirstCol = Columns(ws1.Range("B2").Value).Column
    LastCol = Columns(ws1.Range("C2").Value).Column
    ```
    - The starting and ending columns for analysis are determined using cells B2 and C2 of `ws1`.

3. **Looping Through Specified Columns and Processing Data:**
    ```vba
    For k = FirstCol To LastCol
        FirstValue = ws.Cells(3, k).Value
        SecondValue = ws.Cells(4, k).Value

        Arr1 = Split(FirstValue, ",")
        Arr2 = Split(SecondValue, ",")
    ```
    - For each column within the defined range, the values in the third and fourth rows are split based on commas and stored in arrays `Arr1` and `Arr2`.

4. **Comparing Data in the Arrays:**
    ```vba
    For i = LBound(Arr1) To UBound(Arr1)
        For j = LBound(Arr2) To UBound(Arr2)
            If Arr1(i) = Arr2(j) Then
                Matched = Matched + 1
            Else
                NotMatched = NotMatched + 1
            End If
        Next j
    Next i
    ```
    - Each element from `Arr1` is compared to each element from `Arr2`. If there's a match, the `Matched` counter is incremented. Otherwise, the `NotMatched` counter is incremented.

5. **Determining Match Category and Highlighting:**
    ```vba
    If Matched >= 1 And NotMatched >= 1 Then
        ws.Cells(3, k).Interior.Color = vbYellow
        ws.Cells(4, k).Interior.Color = vbYellow
    ElseIf Matched = 0 And NotMatched >= 1 Then
        ws.Cells(3, k).Interior.Color = vbRed
        ws.Cells(4, k).Interior.Color = vbRed
    Else
        ws.Cells(3, k).Interior.Color = vbGreen
        ws.Cells(4, k).Interior.Color = vbGreen
    End If
    ```
    - Depending on the comparison results, cells are highlighted:
      - **Partial Match (Yellow):** At least one match and one non-match.
      - **No Match (Red):** No matches at all.
      - **Full Match (Green):** All elements match.

6. **Populating Data Dictionary:**
    ```vba
    For Each arrKey In ArrMerged
        If Not dataDict.Exists(arrKey) Then
            Set dataDict(arrKey) = CreateObject("Scripting.Dictionary")
            ...
        End If
        ...
    Next arrKey
    ```
    - Each unique value in `ArrMerged` is checked against the dictionary. If it doesn't exist, it's added to `dataDict` along with sub-keys for counting occurrences.

7. **Generating the Summary Report:**
    ```vba
    i = 10
    For Each key In dataDict.Keys
        ws.Cells(i, "K").Value = key
        ...
        i = i + 5
    Next
    ```
    - The report is written in columns K, L, and M of the `ws` sheet, presenting unique values, their counts, and the percentage representation for each match category.

8. **Counting Color-based Results:**
    ```vba
    For Each cell In ws.Range("5:5")
        Select Case cell.Value
            Case 1
                totalGreen = totalGreen + 1
            ...
        End Select
    Next cell
    ```
    - Each cell in row 5 is checked for its value to count the number of each color (yellow, red, green). The results are then written in the `ws` sheet.

---