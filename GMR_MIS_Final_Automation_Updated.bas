Attribute VB_Name = "GMR_MIS_Automation1"
'===============================================================================
' MODULE: RawToExpected_Converter
' PURPOSE: Converts the "Raw_data" MIS export (one row per Job#+Charge Code)
'          into the pivoted "Expected_output" layout (one row per Job#, with
'          every Charge Code split into an Income column and an Expense column).
'
' HOW TO USE:
'   1. Open Raw_data.xlsx (or a workbook that contains a sheet with this raw
'      layout) in Excel.
'   2. Press ALT+F11 to open the VBA editor.
'   3. File > Import File... and select this .bas file (or paste its contents
'      into a new Module).
'   4. Update the constants RAW_SHEET_NAME / OUTPUT_SHEET_NAME below if your
'      sheet names differ.
'   5. Run Sub ConvertRawToExpectedOutput (F5, or Developer > Macros).
'   6. A new sheet named per OUTPUT_SHEET_NAME is created (or replaced) with
'      the pivoted result.
'
' LOGIC SUMMARY:
'   - Raw layout:  Row 5 = headers, Row 6+ = data. Columns A:AE (31 columns).
'       Col 2  = Job #
'       Col 7  = "Custom Document Details" -> split into SB/BE # + SB/BE Date
'       Col 29 = Charge Code
'       Col 30 = Income (INR)
'       Col 31 = Expense (INR)
'   - Output layout:
'       Row 1  = TOTALS row - =SUBTOTAL(109,...) formulas for every charge
'                code column, Total Income, Total Expenses and GP (SUBTOTAL
'                109 ignores manually hidden / filtered-out rows, so it always
'                reflects only the rows currently visible). GP% in row 1 is
'                =IFERROR(GP/TotalIncome,0%).
'       Row 2  = column headers
'       Row 3+ = one row per unique Job #, in first-seen order, with:
'         Cols 1-29   : base job fields (# is re-numbered sequentially), which
'                       now also include:
'                         Col 3 - Shipment Branch: looked up from the 4th-6th
'                                 characters of Job # (e.g. "MAA" -> "CHENNAI")
'                         Col 4 - Shipment Mode: looked up from the 8th-9th
'                                 characters of Job # (e.g. "AE" -> "AIR EXPORT")
'                         Col 9 - Month: "mmm yy" (e.g. "Apr 26"), derived from
'                                 the Job Execution/MIS Date, placed right
'                                 after the Job Date / MIS Date / Invoice Date
'                                 columns. Stored as literal text (the column
'                                 is pre-formatted "@") so Excel doesn't
'                                 re-interpret it as a date.
'       All date columns are formatted "dd-mm-yyyy" and the whole sheet uses
'       the Tahoma font.
'         Cols 30-74  : one column per unique Charge Code (Income), sorted A-Z
'                       (GROUPED via Data > Group; Total Income is not in the group)
'         Col  75     : Total Income
'         Col  76     : CGST 9%      (= Total Income x 9%)            - HIDDEN by default
'         Col  77     : SGST 9%      (= Total Income x 9%)            - HIDDEN by default
'         Col  78     : Gross Income (= Total Income + CGST + SGST)   - HIDDEN by default
'         Cols 79-123 : same Charge Codes again (Expense), sorted A-Z
'                       (GROUPED; Total Expenses is not in the group)
'         Col  124    : Total Expenses
'         Col  125    : GP  (Total Income - Total Expenses)
'         Col  126    : GP% (GP / Total Income, or 0 if Total Income = 0)
'         Col  127    : (blank spacer)
'         Col  128    : Charge Code Income  - live lookup for the charge code chosen
'                                in the dropdown (col 131), one value per job
'         Col  129    : Charge Code Expense - same lookup, expense side
'         Col  130    : (blank spacer, column width 0.15) - narrow separator
'                                column immediately before the dropdown
'         Col  131    : Charge-code dropdown cell (data validation list, built
'                                from the charge codes found in this run),
'                                placed in row 1, with the label "MIS Remarks"
'                                in row 2 underneath it. Change the selection
'                                and columns 128/129 recalculate instantly -
'                                no macro re-run needed. Row 1 also shows bold
'                                SUBTOTAL totals for the two lookup columns,
'                                same as every other money column.
'       (Column numbers shift automatically if the charge-code count changes.
'        CGST/SGST/Gross always sit directly after Total Income.)
'
'       NOTE: charge codes are discovered dynamically from the data (case-
'       sensitive exact match, blank values and "(N/A)" are ignored), so the
'       macro keeps working even if new charge codes are added later - the
'       number of output columns will simply grow/shrink to match.
'
' FORMATTING (GMR MIS Automation):
'   - Money columns (Income/Expense charge codes, Total Income, Total
'     Expenses, GP) use Accounting format, 2 decimals, Indian Rupee symbol.
'   - Income headers (charge codes + Total Income)   -> fill #139D20, white bold text
'   - Expense headers (charge codes + Total Expenses) -> fill #FA0A0A, white bold text
'   - GP / GP% headers                                -> fill #201799, white bold text
'   - Conditional formatting: font colour turns red (#FA0A0A) wherever a value
'     is negative (< Rs.0 or < 0%) - applies to every charge-code Income/
'     Expense column, Total Income, Total Expenses, GP, GP%, and the
'     charge-code lookup columns (Sales/Cost), across the totals row and all
'     data rows.
'   - A narrow (width 0.15) blank spacer column separates the charge-code
'     lookup columns from the Charge-code dropdown column.
'   - All borders applied across the header + data range.
'   - Workbook "Title" document property set to "GMR MIS Automation".
'===============================================================================

Option Explicit

Public Sub ConvertRawToExpectedOutput()

    Const RAW_SHEET_NAME As String = "MIS Final"       ' name of the raw data sheet
    Const OUTPUT_SHEET_NAME As String = "MIS Final (2)" ' name to give the pivoted output sheet
    Const HEADER_ROW As Long = 5                        ' row containing raw column headers
    Const RAW_COL_COUNT As Long = 31                    ' columns A:AE in Raw_data

    ' Raw column positions (1-based) we need
    Const RC_JOBNO As Long = 2
    Const RC_JOBTYPE As Long = 3
    Const RC_JOBDATE As Long = 4
    Const RC_MISDATE As Long = 5
    Const RC_INVOICEDATE As Long = 6
    Const RC_CUSTOMDOC As Long = 7
    Const RC_MAWB As Long = 8
    Const RC_HAWB As Long = 9
    Const RC_CUSTOMER As Long = 10
    Const RC_SHIPPER As Long = 11
    Const RC_CONSIGNEE As Long = 12
    Const RC_COMMODITY As Long = 13
    Const RC_CHWT As Long = 14
    Const RC_PACKAGES As Long = 15
    Const RC_VOUCHERPARTY As Long = 16
    Const RC_VOUCHERNO As Long = 17
    Const RC_VOUCHERDATE As Long = 18
    Const RC_POL As Long = 19
    Const RC_POD As Long = 20
    Const RC_CNTNO As Long = 21
    Const RC_CNTTYPE As Long = 22
    Const RC_NUMCNTRS As Long = 23
    Const RC_STATUS As Long = 24
    Const RC_ACCSTATUS As Long = 25
    Const RC_CHARGECODE As Long = 29
    Const RC_INCOME As Long = 30
    Const RC_EXPENSE As Long = 31

    Const BASE_COL_COUNT As Long = 29   ' output columns 1-29 (see header layout above;
                                         ' includes Shipment Branch, Shipment Mode and Month)

    ' output column positions of the derived/lookup columns, used for formatting
    Const OC_SHIPMENT_BRANCH As Long = 3
    Const OC_SHIPMENT_MODE As Long = 4
    Const OC_MONTH As Long = 9

    Dim wbBook As Workbook
    Dim wsRaw As Worksheet
    Dim wsOut As Worksheet
    Dim rawData As Variant
    Dim lastRow As Long
    Dim i As Long, j As Long
    Dim t0 As Double
    t0 = Timer

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail

    Set wbBook = ThisWorkbook

    ' --- locate raw sheet -----------------------------------------------------
    On Error Resume Next
    Set wsRaw = wbBook.Sheets(RAW_SHEET_NAME)
    On Error GoTo 0
    If wsRaw Is Nothing Then Set wsRaw = wbBook.Sheets(1)

    lastRow = wsRaw.Cells(wsRaw.Rows.Count, RC_JOBNO).End(xlUp).Row
    If lastRow < HEADER_ROW + 1 Then
        MsgBox "No data found below the header row on sheet '" & wsRaw.Name & "'.", vbExclamation
        GoTo CleanExit
    End If

    ' pull the whole data block into memory in one shot (fast)
    rawData = wsRaw.Range(wsRaw.Cells(HEADER_ROW + 1, 1), wsRaw.Cells(lastRow, RAW_COL_COUNT)).Value

    Dim rowCount As Long
    rowCount = UBound(rawData, 1)

    '======================================================================
    ' PASS 1: discover unique charge codes (ignore blanks and "(N/A)")
    '======================================================================
    Dim dictCharges As Object
    Set dictCharges = CreateObject("Scripting.Dictionary")
    dictCharges.CompareMode = 1 ' vbTextCompare - treat codes case-insensitively

    Dim ccRaw As String
    For i = 1 To rowCount
        ccRaw = Trim$(CStr(rawData(i, RC_CHARGECODE) & ""))
        If ccRaw <> "" And StrComp(ccRaw, "(N/A)", vbTextCompare) <> 0 Then
            If Not dictCharges.Exists(ccRaw) Then dictCharges.Add ccRaw, ccRaw
        End If
    Next i

    Dim chargeCount As Long
    chargeCount = dictCharges.Count

    Dim chargeCodes() As String
    ReDim chargeCodes(1 To chargeCount)
    j = 1
    Dim keyItem As Variant
    For Each keyItem In dictCharges.Keys
        chargeCodes(j) = CStr(keyItem)
        j = j + 1
    Next keyItem

    Call SortStringArrayAsc(chargeCodes)

    ' map charge code -> its 1-based position within the charge block
    Dim dictChargeIdx As Object
    Set dictChargeIdx = CreateObject("Scripting.Dictionary")
    dictChargeIdx.CompareMode = 1
    For j = 1 To chargeCount
        dictChargeIdx.Add chargeCodes(j), j
    Next j

    '======================================================================
    ' PASS 2: aggregate one record per Job #, in first-seen order
    '======================================================================
    Dim dictJobs As Object
    Set dictJobs = CreateObject("Scripting.Dictionary")
    dictJobs.CompareMode = 1

    Dim baseData() As Variant
    Dim incomeData() As Double
    Dim expenseData() As Double

    ReDim baseData(1 To rowCount, 1 To BASE_COL_COUNT)   ' upper bound = rowCount (never exceeded)
    ReDim incomeData(1 To rowCount, 1 To chargeCount)
    ReDim expenseData(1 To rowCount, 1 To chargeCount)

    Dim jobCount As Long: jobCount = 0
    Dim jobNo As String, jobIdx As Long
    Dim customDoc As String, sbNo As Variant, sbDateVal As Variant
    Dim spacePos As Long, sPart As String, dPart As String

    For i = 1 To rowCount

        jobNo = Trim$(CStr(rawData(i, RC_JOBNO) & ""))
        If jobNo = "" Then GoTo NextRow

        If dictJobs.Exists(jobNo) Then
            jobIdx = dictJobs(jobNo)
        Else
            jobCount = jobCount + 1
            jobIdx = jobCount
            dictJobs.Add jobNo, jobIdx

            ' --- split "Custom Document Details" into SB/BE # and SB/BE Date ---
            customDoc = Trim$(CStr(rawData(i, RC_CUSTOMDOC) & ""))
            sbNo = Empty
            sbDateVal = Empty
            If customDoc <> "" Then
                spacePos = InStr(customDoc, " ")
                If spacePos > 0 Then
                    sPart = Left$(customDoc, spacePos - 1)
                    dPart = Mid$(customDoc, spacePos + 1)
                    If IsNumeric(sPart) Then
                        sbNo = CDbl(sPart)
                    Else
                        sbNo = sPart
                    End If
                    sbDateVal = ParseFlexibleDate(dPart)
                Else
                    If IsNumeric(customDoc) Then
                        sbNo = CDbl(customDoc)
                    Else
                        sbNo = customDoc
                    End If
                End If
            End If

            ' --- base (non-charge) fields -> output columns 1-29 ---
            Dim misDateVal As Variant
            misDateVal = ParseFlexibleDate(rawData(i, RC_MISDATE))

            baseData(jobIdx, 1) = jobCount                                    ' # (re-sequenced)
            baseData(jobIdx, 2) = jobNo                                       ' Job #
            baseData(jobIdx, 3) = GetShipmentBranch(jobNo)                    ' Shipment Branch (from Job # chars 4-6)
            baseData(jobIdx, 4) = GetShipmentMode(jobNo)                      ' Shipment Mode (from Job # chars 8-9)
            baseData(jobIdx, 5) = rawData(i, RC_JOBTYPE)                      ' Job Type
            baseData(jobIdx, 6) = ParseFlexibleDate(rawData(i, RC_JOBDATE))   ' Job Date
            baseData(jobIdx, 7) = misDateVal                                  ' Job Execution/MIS Date
            baseData(jobIdx, 8) = ParseFlexibleDate(rawData(i, RC_INVOICEDATE)) ' Invoice Date
            baseData(jobIdx, 9) = MonthYearLabel(misDateVal)                  ' Month ("mmm yy", from Job Execution date)
            baseData(jobIdx, 10) = sbNo                                       ' SB/BE #
            baseData(jobIdx, 11) = sbDateVal                                  ' SB/BE Date
            baseData(jobIdx, 12) = rawData(i, RC_MAWB)                        ' MAWB / MBL # / CN #
            baseData(jobIdx, 13) = rawData(i, RC_HAWB)                        ' HAWB / HBL #
            baseData(jobIdx, 14) = rawData(i, RC_CUSTOMER)                    ' Customer
            baseData(jobIdx, 15) = rawData(i, RC_SHIPPER)                     ' Shipper
            baseData(jobIdx, 16) = rawData(i, RC_CONSIGNEE)                   ' Consignee
            baseData(jobIdx, 17) = rawData(i, RC_COMMODITY)                   ' Commodity Description
            baseData(jobIdx, 18) = ToNumberOrBlank(rawData(i, RC_CHWT))       ' Ch. Wt.
            baseData(jobIdx, 19) = ToNumberOrBlank(rawData(i, RC_PACKAGES))   ' Total Number of Packages
            baseData(jobIdx, 20) = rawData(i, RC_VOUCHERPARTY)                ' Voucher Party
            baseData(jobIdx, 21) = rawData(i, RC_VOUCHERNO)                   ' Voucher #
            baseData(jobIdx, 22) = ParseFlexibleDate(rawData(i, RC_VOUCHERDATE)) ' Voucher Date
            baseData(jobIdx, 23) = rawData(i, RC_POL)                         ' POL Name
            baseData(jobIdx, 24) = rawData(i, RC_POD)                         ' POD Name
            baseData(jobIdx, 25) = rawData(i, RC_CNTNO)                       ' Cnt. #
            baseData(jobIdx, 26) = rawData(i, RC_CNTTYPE)                     ' Cnt. Type
            baseData(jobIdx, 27) = ToNumberOrBlank(rawData(i, RC_NUMCNTRS))   ' # of Cntrs
            baseData(jobIdx, 28) = rawData(i, RC_STATUS)                      ' Status
            baseData(jobIdx, 29) = rawData(i, RC_ACCSTATUS)                   ' Acc. Status
        End If

        ' --- pivot the charge code amount into the right income/expense column ---
        ccRaw = Trim$(CStr(rawData(i, RC_CHARGECODE) & ""))
        If ccRaw <> "" And StrComp(ccRaw, "(N/A)", vbTextCompare) <> 0 Then
            If dictChargeIdx.Exists(ccRaw) Then
                Dim cIdx As Long
                cIdx = dictChargeIdx(ccRaw)
                incomeData(jobIdx, cIdx) = incomeData(jobIdx, cIdx) + ToNumberOrBlank(rawData(i, RC_INCOME))
                expenseData(jobIdx, cIdx) = expenseData(jobIdx, cIdx) + ToNumberOrBlank(rawData(i, RC_EXPENSE))
            End If
        End If

NextRow:
    Next i

    '======================================================================
    ' BUILD THE OUTPUT ARRAY (row 1 = grand totals, row 2 = headers,
    ' rows 3.. = one row per job)
    '======================================================================
    Dim colTotalIncome As Long, colCGST As Long, colSGST As Long, colGross As Long
    Dim colTotalExpense As Long, colGP As Long, colGPpct As Long
    Dim totalCols As Long

    colTotalIncome = BASE_COL_COUNT + chargeCount + 1
    colCGST = colTotalIncome + 1        ' CGST 9%       - right next to Total Income
    colSGST = colCGST + 1               ' SGST 9%
    colGross = colSGST + 1              ' Gross Income
    colTotalExpense = colGross + chargeCount + 1   ' Expense block starts after Gross Income
    colGP = colTotalExpense + 1
    colGPpct = colGP + 1
    totalCols = colGPpct

    Dim outArr() As Variant
    ReDim outArr(1 To jobCount + 1, 1 To totalCols)

    ' ---- headers (row 1) ----
    Dim baseHeaders As Variant
    baseHeaders = Array("#", "Job #", "Shipment Branch", "Shipment Mode", "Job Type", _
                         "Job Date", "Job Execution/MIS Date", "Invoice Date", "Month", _
                         "SB/BE #", "SB/BE Date", "MAWB / MBL # / CN #", _
                         "HAWB / HBL #", "Customer", "Shipper", "Consignee", _
                         "Commodity Description", "Ch. Wt.", "Total Number of Packages", _
                         "Voucher Party", "Voucher #", "Voucher Date", "POL Name", "POD Name", _
                         "Cnt. #", "Cnt. Type", "# of Cntrs", "Status", "Acc. Status")

    For j = 1 To BASE_COL_COUNT
        outArr(1, j) = baseHeaders(j - 1)
    Next j
    For j = 1 To chargeCount
        outArr(1, BASE_COL_COUNT + j) = chargeCodes(j)
    Next j
    outArr(1, colTotalIncome) = "Total Income"
    outArr(1, colCGST) = "CGST 9%"
    outArr(1, colSGST) = "SGST 9%"
    outArr(1, colGross) = "Gross Income"
    For j = 1 To chargeCount
        outArr(1, colGross + j) = chargeCodes(j)
    Next j
    outArr(1, colTotalExpense) = "Total Expenses"
    outArr(1, colGP) = "GP"
    outArr(1, colGPpct) = "GP%"

    ' ---- data rows (row 2 onward) ----
    Dim outRow As Long, rowIncome As Double, rowExpense As Double
    For i = 1 To jobCount
        outRow = i + 1
        For j = 1 To BASE_COL_COUNT
            outArr(outRow, j) = baseData(i, j)
        Next j

        rowIncome = 0
        For j = 1 To chargeCount
            outArr(outRow, BASE_COL_COUNT + j) = incomeData(i, j)
            rowIncome = rowIncome + incomeData(i, j)
        Next j
        outArr(outRow, colTotalIncome) = rowIncome
        ' CGST / SGST / Gross Income cells are overwritten with live formulas
        ' after the array is written to the sheet (see GST COLUMNS below)

        rowExpense = 0
        For j = 1 To chargeCount
            outArr(outRow, colGross + j) = expenseData(i, j)
            rowExpense = rowExpense + expenseData(i, j)
        Next j
        outArr(outRow, colTotalExpense) = rowExpense

        outArr(outRow, colGP) = rowIncome - rowExpense
        If rowIncome <> 0 Then
            outArr(outRow, colGPpct) = (rowIncome - rowExpense) / rowIncome
        Else
            outArr(outRow, colGPpct) = 0
        End If
    Next i

    '======================================================================
    ' WRITE TO THE OUTPUT SHEET
    '======================================================================
    On Error Resume Next
    Application.DisplayAlerts = False
    wbBook.Sheets(OUTPUT_SHEET_NAME).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOut = wbBook.Sheets.Add(After:=wsRaw)
    wsOut.Name = OUTPUT_SHEET_NAME

    ' Pre-format the Month column as Text BEFORE writing values into it, so
    ' Excel stores "Apr 26" literally instead of recognising it as a date and
    ' re-displaying it as e.g. "26-Apr" under a different, inherited format.
    wsOut.Columns(OC_MONTH).NumberFormat = "@"

    wsOut.Range(wsOut.Cells(2, 1), wsOut.Cells(jobCount + 2, totalCols)).Value = outArr

    '======================================================================
    ' ROW 1: SUBTOTAL(109, ...) formulas for every charge-code column,
    ' Total Income, Total Expenses and GP (SUBTOTAL 109 = SUM that ignores
    ' manually hidden / filtered-out rows, so this total always reflects
    ' only the rows currently visible under the autofilter).
    ' GP% is a formula referencing this same row: =IFERROR(GP/Income,0%)
    '======================================================================
    Dim dataStartRow As Long, dataEndRow As Long
    dataStartRow = 3
    dataEndRow = jobCount + 2

    Dim totalsCol As Long, colLtr As String
    For totalsCol = BASE_COL_COUNT + 1 To colGP   ' covers all income charge codes,
                                                    ' Total Income, all expense charge
                                                    ' codes, Total Expenses, and GP
        colLtr = ColLetter(wsOut, totalsCol)
        wsOut.Cells(1, totalsCol).Formula = "=SUBTOTAL(109," & colLtr & dataStartRow & ":" & colLtr & dataEndRow & ")"
    Next totalsCol

    Dim incomeColLtr As String, gpColLtr As String
    incomeColLtr = ColLetter(wsOut, colTotalIncome)
    gpColLtr = ColLetter(wsOut, colGP)
    wsOut.Cells(1, colGPpct).Formula = "=IFERROR(" & gpColLtr & "1/" & incomeColLtr & "1,0%)"

    '======================================================================
    ' CHARGE-CODE LOOKUP: a dropdown of charge codes, with that code's
    ' Income and Expenses pulled out per job right next to it - so instead
    ' of hunting across 45+ charge-code columns, you pick one code and see
    ' its numbers for every job immediately.
    '   - One blank spacer column after GP%, then "Income", "Expenses"
    '     headers, then the dropdown cell itself.
    '   - The dropdown's list source IS the Income charge-code header row
    '     (already the sorted, de-duplicated list), so it always matches
    '     whatever codes this run's data produced - nothing to maintain.
    '   - The two columns are INDEX/MATCH formulas (not static values), so
    '     changing the dropdown updates them instantly with no macro needed.
    '======================================================================
    Dim colIncomeSel As Long, colExpenseSel As Long, colSpacerBeforeDropdown As Long, colDropdown As Long
    colIncomeSel = colGPpct + 2      ' colGPpct + 1 is left as a blank spacer column
    colExpenseSel = colIncomeSel + 1
    colSpacerBeforeDropdown = colExpenseSel + 1   ' narrow blank spacer column, width 0.15,
                                                    ' right before the Charge-code dropdown
    colDropdown = colSpacerBeforeDropdown + 1

    wsOut.Cells(2, colIncomeSel).Value = "Charge Code Income"
    wsOut.Cells(2, colExpenseSel).Value = "Charge Code Expense"

    If chargeCount > 0 Then
        Dim incFirstLtr As String, incLastLtr As String
        Dim expFirstLtr As String, expLastLtr As String
        Dim dropLtr As String, selColLtr As String
        incFirstLtr = ColLetter(wsOut, BASE_COL_COUNT + 1)
        incLastLtr = ColLetter(wsOut, BASE_COL_COUNT + chargeCount)
        expFirstLtr = ColLetter(wsOut, colGross + 1)
        expLastLtr = ColLetter(wsOut, colGross + chargeCount)
        dropLtr = ColLetter(wsOut, colDropdown)

        ' dropdown cell: data-validation list, sourced from the Income
        ' charge-code header row itself. Placed in row 1, with its label
        ' ("MIS Remarks") in row 2 underneath it.
        With wsOut.Cells(1, colDropdown).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:="=$" & incFirstLtr & "$2:$" & incLastLtr & "$2"
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
        wsOut.Cells(1, colDropdown).Value = chargeCodes(1)   ' default to the first code
        wsOut.Cells(2, colDropdown).Value = "MIS Remarks"    ' label shown directly below the dropdown

        ' row 1: SUBTOTAL totals for the two lookup columns (same
        ' hidden/filtered-row-aware pattern as every other money column)
        selColLtr = ColLetter(wsOut, colIncomeSel)
        wsOut.Cells(1, colIncomeSel).Formula = "=SUBTOTAL(109," & selColLtr & dataStartRow & ":" & selColLtr & dataEndRow & ")"
        selColLtr = ColLetter(wsOut, colExpenseSel)
        wsOut.Cells(1, colExpenseSel).Formula = "=SUBTOTAL(109," & selColLtr & dataStartRow & ":" & selColLtr & dataEndRow & ")"

        ' data rows: INDEX/MATCH the selected charge code for that job's row.
        ' Column refs are $-anchored, row refs relative, so setting the
        ' formula once on the whole range fills it down correctly.
        Dim incomeLookupFormula As String, expenseLookupFormula As String
        incomeLookupFormula = "=IFERROR(INDEX($" & incFirstLtr & dataStartRow & ":$" & incLastLtr & dataStartRow & _
            ",MATCH($" & dropLtr & "$1,$" & incFirstLtr & "$2:$" & incLastLtr & "$2,0)),0)"
        expenseLookupFormula = "=IFERROR(INDEX($" & expFirstLtr & dataStartRow & ":$" & expLastLtr & dataStartRow & _
            ",MATCH($" & dropLtr & "$1,$" & expFirstLtr & "$2:$" & expLastLtr & "$2,0)),0)"

        wsOut.Range(wsOut.Cells(dataStartRow, colIncomeSel), wsOut.Cells(dataEndRow, colIncomeSel)).Formula = incomeLookupFormula
        wsOut.Range(wsOut.Cells(dataStartRow, colExpenseSel), wsOut.Cells(dataEndRow, colExpenseSel)).Formula = expenseLookupFormula
    End If

    '======================================================================
    ' GST COLUMNS (sit immediately after Total Income):
    '   CGST 9%      = Total Income x 9%
    '   SGST 9%      = Total Income x 9%
    '   Gross Income = Total Income + CGST + SGST
    ' Live formulas; row-1 SUBTOTALs come from the totals loop above. The
    ' three columns are HIDDEN by default (unhide them when needed).
    '======================================================================
    Dim tiLtr As String, cgstLtr As String, sgstLtr As String
    tiLtr = ColLetter(wsOut, colTotalIncome)
    cgstLtr = ColLetter(wsOut, colCGST)
    sgstLtr = ColLetter(wsOut, colSGST)

    wsOut.Range(wsOut.Cells(dataStartRow, colCGST), wsOut.Cells(dataEndRow, colCGST)).Formula = "=" & tiLtr & dataStartRow & "*9%"
    wsOut.Range(wsOut.Cells(dataStartRow, colSGST), wsOut.Cells(dataEndRow, colSGST)).Formula = "=" & tiLtr & dataStartRow & "*9%"
    wsOut.Range(wsOut.Cells(dataStartRow, colGross), wsOut.Cells(dataEndRow, colGross)).Formula = _
        "=" & tiLtr & dataStartRow & "+" & cgstLtr & dataStartRow & "+" & sgstLtr & dataStartRow

    Dim extendedLastCol As Long
    extendedLastCol = colDropdown

    '---- formatting ----
    Dim rupeeSymbol As String, acctFmt As String
    rupeeSymbol = ChrW$(8377)  ' Indian Rupee sign (U+20B9) - ChrW$ is required for
                                ' Unicode code points above 255; Chr$ only covers ANSI
                                ' and raises a runtime error for this character.
    acctFmt = "_-" & rupeeSymbol & "* #,##0.00_-;-" & rupeeSymbol & "* #,##0.00_-;" & _
              "_-" & rupeeSymbol & "* ""-""??_-;_-@_-"

    With wsOut
        .Cells.Font.Name = "Tahoma"

        ' ---- Row 1: totals row (SUBTOTAL formulas) ----
        With .Range(.Cells(1, 1), .Cells(1, totalCols))
            .Font.Bold = True
            .Interior.Color = RGB(242, 242, 242)   ' light grey, to set it apart from data
        End With

        ' ---- Row 1: Total Income/Total Expenses subtotal cells + the charge-code dropdown - bold too ----
        With .Range(.Cells(1, colIncomeSel), .Cells(1, colDropdown))
            .Font.Bold = True
            .Interior.Color = RGB(242, 242, 242)
        End With

        ' ---- Row 2: column headers ----
        .Rows(2).Font.Bold = True
        .Range(.Cells(2, 1), .Cells(2, totalCols)).AutoFilter

        ' date columns: Job Date, MIS Date, Invoice Date, SB/BE Date, Voucher Date
        Dim dateCols As Variant, dCol As Variant
        dateCols = Array(6, 7, 8, 11, 22)
        For Each dCol In dateCols
            .Range(.Cells(3, CLng(dCol)), .Cells(jobCount + 2, CLng(dCol))).NumberFormat = "dd-mm-yyyy"
        Next dCol

        ' Month column ("mmm yy") was already pre-formatted as Text ("@")
        ' before the values were written, so no further NumberFormat needed.

        ' money columns: totals row + every charge-code column + Total Income + Total Expenses + GP
        ' -> Accounting format, 2 decimals, Indian Rupee symbol
        .Range(.Cells(1, BASE_COL_COUNT + 1), .Cells(jobCount + 2, colGP)).NumberFormat = acctFmt
        ' GP% column (totals row + data rows) stays a percentage
        .Range(.Cells(1, colGPpct), .Cells(jobCount + 2, colGPpct)).NumberFormat = "0.00%"
        ' charge-code lookup columns (Income / Expenses) - same accounting format
        .Range(.Cells(1, colIncomeSel), .Cells(jobCount + 2, colExpenseSel)).NumberFormat = acctFmt

        ' ---- header colour coding (row 2): Income = green, Expense = red, GP/GP% = blue ----
        ' Income charge-code headers + "Total Income" (contiguous block)
        With .Range(.Cells(2, BASE_COL_COUNT + 1), .Cells(2, colTotalIncome))
            .Interior.Color = RGB(19, 157, 32)     ' #139D20
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' Expense charge-code headers + "Total Expenses" (contiguous block)
        With .Range(.Cells(2, colGross + 1), .Cells(2, colTotalExpense))
            .Interior.Color = RGB(250, 10, 10)     ' #FA0A0A
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' GP and GP% headers
        With .Range(.Cells(2, colGP), .Cells(2, colGPpct))
            .Interior.Color = RGB(32, 23, 153)     ' #201799
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' Sales header - same colour as Total Income
        With .Cells(2, colIncomeSel)
            .Interior.Color = RGB(19, 157, 32)     ' #139D20
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' Charge Code Expense header - same colour as Total Expenses
        With .Cells(2, colExpenseSel)
            .Interior.Color = RGB(250, 10, 10)     ' #FA0A0A
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' CGST 9% / SGST 9% / Gross Income headers - income side, so same green
        With .Range(.Cells(2, colCGST), .Cells(2, colGross))
            .Interior.Color = RGB(19, 157, 32)     ' #139D20
            .Font.Color = RGB(255, 255, 255)
            .Font.Bold = True
        End With

        ' ---- Font turns red wherever a value is negative (< Rs.0 or < 0%) ----
        ' Covers every charge-code Income/Expense column, Total Income, CGST/SGST/Gross, Total
        ' Expenses and GP (one contiguous money block), the GP% column, and
        ' the charge-code lookup columns (Sales/Cost) - across the totals
        ' row (row 1) and every data row.
        Dim rngMoney As Range, rngGPpct As Range, rngLookup As Range
        Set rngMoney = .Range(.Cells(1, BASE_COL_COUNT + 1), .Cells(jobCount + 2, colGP))
        Set rngGPpct = .Range(.Cells(1, colGPpct), .Cells(jobCount + 2, colGPpct))

        rngMoney.FormatConditions.Delete
        rngMoney.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        rngMoney.FormatConditions(1).Font.Color = RGB(250, 10, 10)   ' #FA0A0A

        rngGPpct.FormatConditions.Delete
        rngGPpct.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
        rngGPpct.FormatConditions(1).Font.Color = RGB(250, 10, 10)   ' #FA0A0A

        If chargeCount > 0 Then
            Set rngLookup = .Range(.Cells(1, colIncomeSel), .Cells(jobCount + 2, colExpenseSel))
            rngLookup.FormatConditions.Delete
            rngLookup.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="0"
            rngLookup.FormatConditions(1).Font.Color = RGB(250, 10, 10)   ' #FA0A0A
        End If

        ' ---- borders: all borders across the totals row + header + data range ----
        With .Range(.Cells(1, 1), .Cells(jobCount + 2, extendedLastCol)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        .Application.ActiveWindow.FreezePanes = False
        .Cells(3, 1).Select
        .Application.ActiveWindow.FreezePanes = True

        .Cells.EntireColumn.AutoFit
        .Cells.EntireRow.AutoFit

        ' narrow blank spacer column right before the Charge-code dropdown -
        ' set AFTER AutoFit so AutoFit doesn't widen it back out
        .Columns(colSpacerBeforeDropdown).ColumnWidth = 0.15

        ' GST columns are hidden by default - set AFTER AutoFit (unhide via
        ' right-click > Unhide when needed)
        .Range(.Columns(colCGST), .Columns(colGross)).EntireColumn.Hidden = True

        ' ---- column grouping (outline) ----
        ' Income charge-code columns are grouped WITHOUT Total Income, and
        ' Expense charge-code columns are grouped WITHOUT Total Expenses.
        ' Summary column sits on the right, so the +/- button appears above
        ' Total Income / Total Expenses. Groups start expanded; click the "1"
        ' outline button (top-left) or the "-" button to collapse them.
        If chargeCount > 0 Then
            .Outline.SummaryColumn = xlSummaryOnRight
            .Range(.Columns(BASE_COL_COUNT + 1), .Columns(BASE_COL_COUNT + chargeCount)).Group
            .Range(.Columns(colGross + 1), .Columns(colGross + chargeCount)).Group
        End If
    End With

    ' file / workbook title
    On Error Resume Next
    wbBook.BuiltinDocumentProperties("Title") = "GMR MIS Automation"
    On Error GoTo 0

    wsOut.Cells(1, 1).Select

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.DisplayAlerts = True

    MsgBox "Conversion complete." & vbCrLf & _
           "Jobs processed: " & jobCount & vbCrLf & _
           "Charge codes found: " & chargeCount & vbCrLf & _
           "Output sheet: " & OUTPUT_SHEET_NAME & " (totals in row 1, headers in row 2, data from row 3)" & vbCrLf & _
           "Charge-code lookup dropdown added after GP% - pick a code to see its Sales/Cost per job." & vbCrLf & _
           "Time taken: " & Format(Timer - t0, "0.00") & " sec", vbInformation
    Exit Sub

CleanFail:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical
    Resume CleanExit

End Sub

'===============================================================================
' HELPER: derive the Shipment Mode from a Job #, using the 8th-9th characters
' of the Job # (e.g. Job # "ABCMAA01AE12345" -> code "AE" -> "AIR EXPORT"). If
' the 2-character code isn't in the reference table, the raw code is returned
' as-is (rather than blank) so unmapped modes are still visible for follow-up.
' Returns "" if the Job # is too short to contain an 8th-9th character.
'===============================================================================
Private Function GetShipmentMode(jobNo As String) As String
    Static dictMode As Object
    Dim code As String

    If dictMode Is Nothing Then
        Set dictMode = CreateObject("Scripting.Dictionary")
        dictMode.CompareMode = 1 ' vbTextCompare
        dictMode.Add "AE", "AIR EXPORT"
        dictMode.Add "AI", "AIR IMPORT"
        dictMode.Add "SE", "SEA EXPORT"
        dictMode.Add "SI", "SEA IMPORT"
        dictMode.Add "RD", "ROAD DOMESTIC"
        dictMode.Add "RI", "ROAD IMPORT"
        dictMode.Add "RE", "ROAD EXPORT"
    End If

    If Len(jobNo) >= 9 Then
        code = UCase$(Mid$(jobNo, 8, 2))
        If dictMode.Exists(code) Then
            GetShipmentMode = dictMode(code)
        Else
            GetShipmentMode = code   ' unrecognised mode code - surface it as-is
        End If
    Else
        GetShipmentMode = ""
    End If
End Function

'===============================================================================
' HELPER: derive the Shipment Branch from a Job #, using the 4th-6th characters
' of the Job # (e.g. Job # "ABCMAA12345" -> code "MAA" -> "CHENNAI"). If the
' 3-character code isn't in the reference table, the raw code is returned as-is
' (rather than blank) so unmapped branches are still visible for follow-up.
' Returns "" if the Job # is too short to contain a 4th-6th character.
'===============================================================================
Private Function GetShipmentBranch(jobNo As String) As String
    Static dictBranch As Object
    Dim code As String

    If dictBranch Is Nothing Then
        Set dictBranch = CreateObject("Scripting.Dictionary")
        dictBranch.CompareMode = 1 ' vbTextCompare
        dictBranch.Add "MAA", "CHENNAI"
        dictBranch.Add "BLR", "BANGALORE"
        dictBranch.Add "CJB", "COIMBATORE"
        dictBranch.Add "HYD", "HYDERABAD"
        dictBranch.Add "COK", "COCHIN"
        dictBranch.Add "CCJ", "CALICUT"
        dictBranch.Add "TRV", "TRIVANDRUM"
        dictBranch.Add "TRZ", "TRICHY"
        dictBranch.Add "BOM", "MUMBAI"
        dictBranch.Add "IXM", "MADURAI"
        dictBranch.Add "KNN", "KANNUR"
    End If

    If Len(jobNo) >= 6 Then
        code = UCase$(Mid$(jobNo, 4, 3))
        If dictBranch.Exists(code) Then
            GetShipmentBranch = dictBranch(code)
        Else
            GetShipmentBranch = code   ' unrecognised branch code - surface it as-is
        End If
    Else
        GetShipmentBranch = ""
    End If
End Function

'===============================================================================
' HELPER: build a "mmm yy" label (e.g. "Aug 26") from a date value, for the
' Month column. Uses a fixed English month-abbreviation array (rather than
' VBA's Format "mmm") so the result doesn't depend on the system's regional
' language settings. Returns "" if the input isn't a valid date.
'===============================================================================
Private Function MonthYearLabel(v As Variant) As String
    Dim months As Variant
    months = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    If Not IsDate(v) Then
        MonthYearLabel = ""
        Exit Function
    End If

    Dim d As Date
    d = CDate(v)
    MonthYearLabel = months(Month(d) - 1) & " " & Format(d, "yy")
End Function

'===============================================================================
' HELPER: convert a value to a Double if possible, otherwise return "" (blank).
' Used for numeric fields (Ch. Wt., package counts, income/expense amounts)
' that may arrive from the raw sheet as text.
'===============================================================================
Private Function ToNumberOrBlank(v As Variant) As Variant
    Dim s As String
    s = Trim$(CStr(v & ""))
    If s = "" Then
        ToNumberOrBlank = ""
    ElseIf IsNumeric(s) Then
        ToNumberOrBlank = CDbl(s)
    Else
        ToNumberOrBlank = 0
    End If
End Function

'===============================================================================
' HELPER: parse a date value that may already be a real Date, an ISO string
' ("2026-08-23"), or a "dd-Mmm-yyyy" string ("23-Aug-2026"). Locale-safe
' (does not rely on Windows regional month names). Returns "" for blank input.
'===============================================================================
Private Function ParseFlexibleDate(v As Variant) As Variant
    Dim s As String

    If VarType(v) = vbDate Then
        ParseFlexibleDate = CDate(v)
        Exit Function
    End If

    s = Trim$(CStr(v & ""))
    If s = "" Then
        ParseFlexibleDate = ""
        Exit Function
    End If

    ' ISO format: yyyy-mm-dd
    If Len(s) = 10 And Mid$(s, 5, 1) = "-" And Mid$(s, 8, 1) = "-" Then
        If IsNumeric(Left$(s, 4)) And IsNumeric(Mid$(s, 6, 2)) And IsNumeric(Mid$(s, 9, 2)) Then
            ParseFlexibleDate = DateSerial(CInt(Left$(s, 4)), CInt(Mid$(s, 6, 2)), CInt(Mid$(s, 9, 2)))
            Exit Function
        End If
    End If

    ' dd-Mmm-yyyy format, e.g. 23-Aug-2026
    Dim parts() As String
    parts = Split(s, "-")
    If UBound(parts) = 2 Then
        Dim dd As Integer, mm As Integer, yyyy As Integer
        If IsNumeric(parts(0)) And IsNumeric(parts(2)) Then
            dd = CInt(parts(0))
            yyyy = CInt(parts(2))
            mm = MonthAbbrevToNumber(parts(1))
            If mm > 0 Then
                ParseFlexibleDate = DateSerial(yyyy, mm, dd)
                Exit Function
            End If
        End If
    End If

    ' last resort: let VBA try its default parsing
    On Error Resume Next
    ParseFlexibleDate = CDate(s)
    If Err.Number <> 0 Then
        ParseFlexibleDate = s   ' give up gracefully - keep the original text
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function MonthAbbrevToNumber(m As String) As Integer
    Dim months As Variant
    months = Array("jan", "feb", "mar", "apr", "may", "jun", _
                    "jul", "aug", "sep", "oct", "nov", "dec")
    Dim i As Integer
    Dim mLower As String
    mLower = LCase$(Left$(Trim$(m), 3))
    For i = 0 To 11
        If mLower = months(i) Then
            MonthAbbrevToNumber = i + 1
            Exit Function
        End If
    Next i
    MonthAbbrevToNumber = 0
End Function

'===============================================================================
' HELPER: return the column letter(s) for a given 1-based column number on a
' given worksheet (e.g. 27 -> "AA"). Used to build SUBTOTAL/IFERROR formula
' text such as "AA3:AA3809".
'===============================================================================
Private Function ColLetter(ws As Worksheet, colNum As Long) As String
    Dim addr As String
    addr = ws.Columns(colNum).Address(False, False)   ' e.g. "AA:AA"
    ColLetter = Left$(addr, InStr(addr, ":") - 1)
End Function

'===============================================================================
' HELPER: simple in-place ascending sort (case-insensitive) for a String array.
' Charge-code lists are small (tens of items), so an O(n^2) sort is plenty fast.
'===============================================================================
Private Sub SortStringArrayAsc(arr() As String)
    Dim i As Long, j As Long
    Dim temp As String
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If UCase$(arr(i)) > UCase$(arr(j)) Then
                temp = arr(i)
                arr(i) = arr(j)
                arr(j) = temp
            End If
        Next j
    Next i
End Sub
