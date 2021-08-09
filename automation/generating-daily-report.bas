Attribute VB_Name = "Module1"
Option Explicit
'Option Explicit forces the developer to declare the type of each variable

Sub create_button000()

'creating button
ActiveSheet.Buttons.Add(190, 227, 180, 20).Select
Selection.Name = "button000"
Selection.OnAction = "create_temp_file"
Selection.Characters.Text = "START"
With Selection.Characters(Start:=1, length:=12).Font
    .Name = "Calibri"
    .FontStyle = "Regular"
    .Size = 11
    .Strikethrough = False
    .Superscript = False
    .Subscript = False
    .OutlineFont = False
    .Shadow = False
    .Underline = xlUnderlineStyleNone
    .ColorIndex = 1
End With

'displaying message
Range("D14") = "Ready for next NAV"

'Activating Protection
Range("B2").Select
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True

    
End Sub
Sub create_temp_file()

'unprotecting workbook
ActiveSheet.Unprotect

'Declaring variables
Dim folder_picker As FileDialog
Dim monitoring_folder As String
Dim temp_fullname As String

'Showing message to user
MsgBox "Please select the folder where you want to store the report"

'Use pickup folder to store the file / initializing variables
Set folder_picker = Application.FileDialog(msoFileDialogFolderPicker)
folder_picker.AllowMultiSelect = False
folder_picker.Show
monitoring_folder = folder_picker.SelectedItems(1)
temp_fullname = monitoring_folder + "\" + "TEMP_Technical_Alerts_Report.xlsm"

'Changing file name to temp'
ActiveWorkbook.SaveAs Filename:=temp_fullname, _
FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
CreateBackup:=False

'updating button
ActiveSheet.Shapes.Range(Array("button000")).Select
Selection.Characters.Text = "Search Breachtype file"
Selection.OnAction = "search_transactions_file"

'updating macro sheet messages
Range("D14") = "temp file has been created"
Range("D8") = monitoring_folder

'Activating Protection
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True


End Sub

Sub search_transactions_file()

'unprotecting workbook
ActiveSheet.Unprotect

'Declaring & initializing variables
Dim transactions_file As String

'Showing message to user
MsgBox "Please select the Brech type file you want to analyse"

'show pop up window
transactions_file = Application.GetOpenFilename(, , "Browse for the Bloomberg original report")

'updating button
ActiveSheet.Shapes.Range(Array("button000")).Select
Selection.Characters.Text = "Get Max Nav"
Selection.OnAction = "get_max_nav"

'updating Macro sheet messages
Range("D10") = transactions_file
Range("D14") = "Breach type file has been selected"


'protecting sheet
Range("B2").Select
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True


End Sub

Sub get_max_nav()

'unprotecting workbook
ActiveSheet.Unprotect

'Declaring variables
Dim temp_fullname As String
Dim temp_name As String
Dim transactions_file As String
Dim length As Integer
Dim yyyy As String
Dim mm As String
Dim dd As String

'Initializing variables
transactions_file = Range("D10")
temp_name = ActiveWorkbook.Name
temp_fullname = ActiveWorkbook.FullName

'displaying messages for user
Range("D14") = "Searching Nav"

''Opening breachtype file
Workbooks.Open transactions_file
Sheets("Sheet_Data").Select

'checking if the file has been edited
If ActiveWorkbook.Sheets.Count = 1 Then
    'Creating Edit Sheet
    Sheets.Add
    ActiveSheet.Name = "Edit"
    Sheets("Edit").Move After:=Sheets(2)
    'Creating NAV Sheet
    Sheets.Add
    ActiveSheet.Name = "NAV"
    Sheets("NAV").Move After:=Sheets(3)
    Sheets("Sheet_Data").Select

Else:
    'Deleting old files
    Application.DisplayAlerts = False
    Sheets("Edit").Delete
    Sheets("NAV").Delete
    Application.DisplayAlerts = True
    'Creating Edit Sheet
    Sheets.Add
    ActiveSheet.Name = "Edit"
    Sheets("Edit").Move After:=Sheets(2)
    'Create Nav Sheet
    Sheets.Add
    ActiveSheet.Name = "NAV"
    Sheets("NAV").Move After:=Sheets(3)
    Sheets("Sheet_Data").Select
End If

'''Deleting blank cells
Sheets("Sheet_Data").Select
Cells.Select
Selection.Copy
Sheets("Edit").Select
Range("A1").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("A:A").Select
Selection.SpecialCells(xlCellTypeBlanks).Select
Selection.EntireRow.Delete
       
    
'''Changing format of dates to get max nav
Range("AA2").Select
ActiveCell.FormulaR1C1 = "=MID(RC[-22], 4, 2)"
Range("AB2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[-23], 2)"
Range("AC2").Select
ActiveCell.FormulaR1C1 = "=RIGHT(RC[-24], 4)"
Range("AD2").Select
ActiveCell.FormulaR1C1 = "=DATE(RC[-1],RC[-3],RC[-2])"

'getting amount of lines
length = WorksheetFunction.CountA(Range("A:A"))

'Autofill dates
Range(Cells(2, 27), Cells(2, 30)).Select
Selection.AutoFill Destination:=Range(Cells(2, 27), Cells(length, 30)), Type:=xlFillDefault

'save so that values update
ActiveWorkbook.Save

'getting max nav
Range("AE2").Select
Selection.NumberFormat = "mm/dd/yyyy"
ActiveCell.FormulaR1C1 = "=MAX(C[-1])"
Range("AF2").Select
ActiveCell.FormulaR1C1 = "=TEXT(RC[-1], ""yyyymmdd"")"
Range("AG2").Select
ActiveCell.FormulaR1C1 = "=TEXT(RC[-2], ""dd"")"
Range("AG2").Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("AH2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=TEXT(RC[-3], ""mm"")"
Range("AH2").Select
Selection.Copy
Range("AH2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("AI2").Select
Application.CutCopyMode = False
ActiveCell.FormulaR1C1 = "=TEXT(RC[-4], ""yyyy"")"
Range("AI2").Select
Selection.Copy
Range("AI2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
        
'storing max nav
dd = Range("AG2")
mm = Range("AH2")
yyyy = Range("AI2")

'Converting breach duration from string to int
Range("Z2").Select
ActiveCell.FormulaR1C1 = "=RC[-15]*1"
Range("Z2").Select
Selection.AutoFill Destination:=Range(Cells(2, 26), Cells(length, 26)), Type:=xlFillDefault
Range(Cells(2, 26), Cells(length, 26)).Select
Selection.Copy
Range("K2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range(Cells(2, 26), Cells(length, 26)).Select
Selection.ClearContents

'Closing breachtype file
ActiveWorkbook.Save
ActiveWorkbook.Close

'Month and Year in white
Range("D4:D6").Select
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorDark1
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

'Updating Messages
Windows(temp_name).Activate
MsgBox "Please confirm the Max Nave Date"
MsgBox ("This Macro suggests: Year = " + yyyy + " ,Month = " + mm + " ,Day = " + dd)
Range("D14") = "Waiting NAV confirmation, nav suggested: " + yyyy + "-" + mm + "-" + dd

'updating button
ActiveSheet.Shapes.Range(Array("Button000")).Select
Selection.Characters.Text = "Confirm Nav"
Selection.OnAction = "run_macro_upload_breachtype"

'protecting sheet
Range("B2").Select
Range("D4:D6").Select
Selection.Locked = False
Selection.FormulaHidden = False
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
ActiveWorkbook.Save


End Sub

Sub run_macro_upload_breachtype()

'unprotecting workbook
ActiveSheet.Unprotect

'check if user entered all mm dd yyyy
If Range("D4") = "--Select--" Or Range("D5") = "--Select--" Or Range("D6") = "--Select--" Then
    MsgBox "Please select year, month and day"
    ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
    Exit Sub
    
Else:

End If

'declaring variables
Dim transactions_file_fullname As String
Dim transactions_file_name As String
Dim temp_name As String
Dim temp_fullname As String
Dim date_technical_alerts_report As String
Dim length As Long
Dim yyyy As String
Dim mm As String
Dim dd As String

'initializing variables
yyyy = Range("D4")
mm = Range("D5")
dd = Range("D6")
transactions_file_fullname = Range("D10")
temp_name = ActiveWorkbook.Name
temp_fullname = ActiveWorkbook.FullName
date_technical_alerts_report = Range("D8") + "\" + yyyy + mm + dd + "_Technical_Alerts_Report.xlsm"

'displaying messages
Range("D12").Select
ActiveCell.FormulaR1C1 = "=NOW()"
Selection.Copy
Range("D12").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
Range("D14").Select
Range("D14") = "Running Macro"

'Opening breachtype file
Sheets("Macro").Select
Workbooks.Open transactions_file_fullname
Sheets("Edit").Select

'initializing variable
transactions_file_name = ActiveWorkbook.Name

'getting lenght
length = WorksheetFunction.CountA(Range("A:A"))

'Creating additional columns
Range("A1").Select
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
Selection.EntireColumn.Insert , CopyOrigin:=xlFormatFromLeftOrAbove

'Creating group formula
Range("A1") = "Group"
Range("A2").Select
ActiveCell.FormulaR1C1 = "=IF(RC[18]<>"""",RC[18],""Unclassified"")"


'Creating closed formula
Range("B1") = "Closed"
Range("B2").Select
ActiveCell.FormulaR1C1 = "=IF(RC[10]<>"""",""Yes"",""No"")"

'Creating ISS formula
Range("C1") = "ISS"
Range("C2").Select
ActiveCell.FormulaR1C1 = _
    "=IFERROR(RIGHT(MID(RC[11], FIND(""ISS"", RC[11], 1), 20), 5)*1, """")"

'Creating unique formula
Range("D1") = "Unique"
Range("D2").Select
ActiveCell.FormulaR1C1 = "=IF(COUNTIF(R2C3:RC[-1], RC[-1]) > 1, 0, 1)"

'Creating country formula
Range("E1") = "Country"
Range("E2").Select
ActiveCell.FormulaR1C1 = "=LEFT(RC[2], 2)"

'Using autofill to extend the cells
Range("A2:E2").Select
Selection.AutoFill Destination:=Range(Cells(2, 1), Cells(length, 5))

'DELETING CLOSED = YES
Range(Cells(1, 1), Cells(1, 50)).Select
Selection.AutoFilter
ActiveSheet.Range(Cells(2, 1), Cells(length, 50)).AutoFilter Field:=2, Criteria1:="No"

'copying results where closed = NO
Range(Cells(1, 1), Cells(length, 50)).Select
'Selection.ClearContents
Selection.Copy
Sheets("NAV").Select
Range("A1").Select
ActiveSheet.Paste

'!!!!!!!!!COPY NAV TO MACRO FILE
Sheets("NAV").Select
Sheets("NAV").Copy After:=Workbooks(temp_name). _
    Sheets(4)

Windows(transactions_file_name).Activate
Application.DisplayAlerts = False
ActiveWorkbook.Save
Application.DisplayAlerts = True
ActiveWorkbook.Close
Windows(temp_name).Activate

''GETTING MAX INDEX
'Windows(ThisFileName).Activate
Sheets("Data").Select
Dim max_index As Long
max_index = WorksheetFunction.Max(Rows("1:1"))

'Adding max nav date
Cells(2, max_index + 1) = dd + "/" + mm
Cells(63, max_index + 1) = dd + "/" + mm
Cells(55, max_index + 1) = "NAV"

'updating so that new column updates correctly
ActiveWorkbook.Save

'''''USING AutoFill
Range(Cells(2, max_index + 1), Cells(509, max_index + 1)).Select
Selection.AutoFill Destination:=Range(Cells(2, max_index + 1), Cells(509, max_index + 2)), Type:=xlFillDefault

'CONTERTING (TODAY - 1) TO VALUE
Range(Cells(2, max_index + 1), Cells(509, max_index + 1)).Select
Selection.Copy
Cells(2, max_index + 1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
'updating last column
Cells(2, max_index + 2) = "next"
Cells(55, max_index + 2) = "next"
Cells(67, max_index + 2) = "next"
Cells(55, max_index + 1).Select
Selection.ClearContents

'updating index
Cells(1, max_index + 1) = max_index + 1

'updating data in overview sheet
Sheets("Overview").Select
Range(Cells(129, max_index), Cells(149, max_index)).Select
Selection.AutoFill Destination:=Range(Cells(129, max_index), Cells(149, max_index + 1)), Type:=xlFillDefault

''''DELETING nav sheet because the data has been updated Files
Application.DisplayAlerts = False
Sheets("NAV").Delete
Application.DisplayAlerts = True


'updating overview
Sheets("Data").Select
Range(Cells(2, max_index - 13), Cells(53, max_index + 1)).Select
Selection.Copy
Sheets("Overview").Select
Range("B57").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False

'Updating chart 1
Sheets("Overview").Select
ActiveSheet.ChartObjects("Chart 9").Activate
ActiveChart.SetSourceData Source:=Range(Cells(129, 1), Cells(133, max_index + 1))

'updating chart2
ActiveSheet.ChartObjects("Chart 10").Activate
ActiveChart.SetSourceData Source:=Range(Cells(135, 1), Cells(137, max_index + 1))

'updating chart3
ActiveSheet.ChartObjects("Chart 11").Activate
ActiveChart.SetSourceData Source:=Range(Cells(139, 1), Cells(141, max_index + 1))

'updating chart4
ActiveSheet.ChartObjects("Chart 12").Activate
ActiveChart.SetSourceData Source:=Range(Cells(143, 1), Cells(145, max_index + 1))

'updating chart5
ActiveSheet.ChartObjects("Chart 13").Activate
ActiveChart.SetSourceData Source:=Range(Cells(147, 1), Cells(149, max_index + 1))

'updating button
Sheets("Macro").Select
ActiveSheet.Shapes.Range(Array("Button000")).Select
Selection.Characters.Text = "START"
Selection.OnAction = "create_temp_file"

'updating messages
Range("D14") = "Ready for Next Nav"

'displaying last nav info below
Range("C4:D14").Select
Selection.Copy
Range("C21").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
Range("D31") = "Success"
Range("C4:D12").Select
Selection.ClearContents
Range("D4") = "--Select--"
Range("D5") = "--Select--"
Range("D6") = "--Select--"
Range("C4") = "Year"
Range("C5") = "Month"
Range("C6") = "Day"

'locking day, month and year
Range("D4:D6").Select
Selection.Locked = True
Selection.FormulaHidden = False

'changing color from white to gray for day,  month and year
Range("D4:D6").Select
Selection.Interior.ThemeColor = xlThemeColorDark1
Selection.Interior.TintAndShade = -0.149998474074526

'protecting sheet
ActiveSheet.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True
ActiveWorkbook.Save

'Change name from TEMP
ActiveWorkbook.SaveAs Filename:=date_technical_alerts_report, _
FileFormat:=xlOpenXMLWorkbookMacroEnabled, _
CreateBackup:=False

'Deleting temo file
Kill temp_fullname

'displaying last message
MsgBox "The report has been updated successfully"

End Sub





