Attribute VB_Name = "Module1"
Option Explicit

Sub SEND_ACTIVE_REMINDER()

''' ACTIVE BREACH REMINDER


''GET ISSUE ID
    
    ' Get Issue ID and remove duplicates

    Dim lastrow As Integer
    lastrow = Sheets("Archer Search Report").Cells(1, 1).End(xlDown).Row

    Sheets("Archer Search Report").Activate
    Range(Cells(2, 1), Cells(lastrow, 1)).Select
    Selection.Copy
    Cells(2, 17).Select
    Selection.PasteSpecial xlPasteValues
    Range(Cells(2, 17), Cells(lastrow, 17)).RemoveDuplicates Columns:=1, Header:=xlNo


    ' Get 5 digit Issue ID

    Dim lastidcolumn As Integer
    lastidcolumn = Sheets("Archer Search Report").Cells(2, 17).End(xlDown).Row

    Cells(2, 18).Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-1],16,5)"
    Cells(2, 19).FormulaR1C1 = "=VLOOKUP(RC[-2],C[-18]:C[-7],9,FALSE)"
    Cells(2, 20).FormulaR1C1 = "=VLOOKUP(RC[-3],C[-19]:C[-6],13,FALSE)"
    Cells(2, 21).FormulaR1C1 = "=VLOOKUP(RC[-4],C[-20]:C[-6],7,FALSE)"
    Range(Cells(2, 18), Cells(2, 21)).Select
    Selection.Copy
    Range(Cells(2, 18), Cells(lastidcolumn, 21)).Select
    Selection.PasteSpecial xlPasteFormulas
    Selection.PasteSpecial xlPasteFormats
    
    Range(Cells(2, 18), Cells(lastidcolumn, 21)).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
    
    
    Range(Columns(1), Columns(17)).Select
    Selection.Delete
    
    Rows(1).Select
    Selection.Delete

'' DELETE COMMENTED BREACHES

    ' Define unnecesarry data
    
    Dim lastcolumn3 As Integer
    Dim lastrow4 As Integer


    lastcolumn3 = Cells(1, 1).End(xlToRight).Column
    lastrow4 = Cells(1, 1).End(xlDown).Row


    Cells(1, lastcolumn3 + 1).Formula = "=NOW()"
    Cells(1, lastcolumn3 + 2).Formula = "=MONTH(C1)"
    Cells(1, lastcolumn3 + 3).Formula = "=DAY(C1)"
    Cells(1, lastcolumn3 + 4).Formula = "=YEAR(C1)"
    Cells(1, lastcolumn3 + 5).Formula = "=MONTH(E1)"
    Cells(1, lastcolumn3 + 6).Formula = "=DAY(E1)"
    Cells(1, lastcolumn3 + 7).Formula = "=YEAR(E1)"
    Cells(1, lastcolumn3 + 8).Formula = "=MONTH(D1)"
    Cells(1, lastcolumn3 + 9).Formula = "=DAY(D1)"
    Cells(1, lastcolumn3 + 10).Formula = "=YEAR(D1)"

    
       
    Range(Cells(1, lastcolumn3 + 1), Cells(1, lastcolumn3 + 10)).Select
    Selection.Copy
    Range(Cells(1, lastcolumn3 + 1), Cells(lastrow4, lastcolumn3 + 10)).Select
    Selection.PasteSpecial xlPasteFormulas
    Selection.PasteSpecial xlPasteFormats
    Range(Cells(1, lastcolumn3 + 1), Cells(lastrow4, lastcolumn3 + 10)).Select
    Selection.Copy
    Selection.PasteSpecial xlPasteValues
   
    
    Dim e As Integer

    For e = 1 To lastrow4
        
    If Cells(e, lastcolumn3 + 2).Value = Cells(e, lastcolumn3 + 5).Value And Cells(e, lastcolumn3 + 3).Value = Cells(e, lastcolumn3 + 6).Value And Cells(e, lastcolumn3 + 4).Value = Cells(e, lastcolumn3 + 7).Value Then

    Cells(e, lastcolumn3 + 11).Value = "YES"

    End If
    
    If Cells(e, lastcolumn3 - 1).Value = "0" And Cells(e, lastcolumn3 + 8).Value = Cells(e, lastcolumn3 + 5).Value And Cells(e, lastcolumn3 + 9).Value = Cells(e, lastcolumn3 + 6).Value And Cells(e, lastcolumn3 + 10).Value = Cells(e, lastcolumn3 + 7).Value Then
    
    Cells(e, lastcolumn3 + 11).Value = "YES"
    
    End If
    
    Next e


    ' Add headlines

    Rows("1:1").Select
    Selection.Insert Shift:=xlDown

    Cells(1, 1).Value = "Heading 1"
    Cells(1, 2).Value = "Heading 2"
    Cells(1, 3).Value = "Heading 3"
    Cells(1, 4).Value = "Heading 4"
    Cells(1, 5).Value = "Heading 5"
    Cells(1, 6).Value = "Heading 6"
    Cells(1, 7).Value = "Heading 7"
    Cells(1, 8).Value = "Heading 8"
    Cells(1, 9).Value = "Heading 9"
    Cells(1, 10).Value = "Heading 10"
    Cells(1, 11).Value = "Heading 11"
    Cells(1, 12).Value = "Heading 12"
    Cells(1, 13).Value = "Heading 13"
    Cells(1, 14).Value = "Heading 14"
    Cells(1, 15).Value = "Heading 15"
    
    ' Delete unnecessary data
    
    Cells(1, 1).Select
    Range(Cells(1, 1), Cells(lastrow4 + 1, lastcolumn3 + 11)).AutoFilter
    ActiveSheet.Range(Cells(1, 1), Cells(lastrow4, lastcolumn3 + 11)).AutoFilter Field:=15, Criteria1:="YES"
    Cells.Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Delete
        
    
    ' Delete unnecessary columns
    
    Range(Columns(lastcolumn3 - 1), Columns(lastcolumn3 + 11)).Select
    Selection.Delete
    
    
'' ADD EMAILS OF COMPLIANCE OFFICERS

    Dim f As Integer
    Dim lastrow5 As Integer
    
    lastrow5 = Cells(1, 1).End(xlDown).Row
    
    For f = 1 To lastrow5
   
    If Cells(f, 2).Value = "Last 1, First 1" Then
    Cells(f, 3).Value = "last1.first1@dummy.com"
    Else
    If Cells(f, 2).Value = "Last 1, First 1" Then
    Cells(f, 3).Value = "last1.first1@dummy.com"
    
    End if
    End if
    Next f
        
    ' Format the final table
    
    
    Columns(1).Select
    Selection.ColumnWidth = 13
    Columns(2).Select
    Selection.ColumnWidth = 25
    Columns(3).Select
    Selection.ColumnWidth = 50
    Cells(1, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.InsertIndent 1
    
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 1
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
        
    
    Columns("A:C").Sort key1:=Range("B2"), _
      order1:=xlAscending, Header:=xlNo
        
    
    ' Add final headings
    
    Rows("1:1").Select
    Selection.Insert Shift:=xlDown

    Cells(1, 1).Value = "Issue ID"
    Cells(1, 2).Value = "Compliance Officer"
    Cells(1, 3).Value = "Emails"
    
    
'' FORMAT THE FINAL TABLE FOR THE EMAIL
    
    Cells(1, 1).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.InsertIndent 1
    Selection.Style = "Heading 2"
    Cells(2, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:= _
        "=MOD(ROW(),2)=1"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorLight2
        .TintAndShade = 0.799981688894314
    End With
    Selection.FormatConditions(1).StopIfTrue = True
    
    Cells(1, 1).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ThemeColor = 1
        .TintAndShade = 0
        .Weight = xlThick
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
       
       
'' ADD TABLE TO EMAIL AND DISPLAY !!!!!!!!

    ' Create an distribution list

    Dim lastrowfinal As Integer
    lastrowfinal = Cells(1, 1).End(xlDown).Row

    Range(Cells(2, 3), Cells(lastrowfinal, 3)).Select
    Selection.Copy
    Cells(2, 4).Select
    Selection.PasteSpecial xlPasteValues
    Range(Cells(2, 4), Cells(lastrowfinal, 4)).RemoveDuplicates Columns:=1, Header:=xlNo
    
    
    Range(Cells(2, 4), Cells(lastrowfinal, 4)).Select
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Cells(2, 5).Select
    
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    

       
    ' Define variables
    
    Dim wb As Workbook, rng As Range
    Set wb = ActiveWorkbook

    Set rng = Range(Cells(1, 1), Cells(lastrowfinal, 2))


    Dim fso As Scripting.FileSystemObject
    Set fso = New FileSystemObject

    ' Create a html text file

    fso.CreateTextFile ("O:\GMR\Processes\Downloads\Vügar\Working Student\Active Breach Reminder\test.html")

    wb.PublishObjects.Add(xlSourceRange, "O:\GMR\Processes\Downloads\Vügar\Working Student\Active Breach Reminder\test.html", wb.Sheets(1).name, wb.Sheets(1).Range(Cells(1, 1), Cells(lastrowfinal, 2)).Address, xlHtmlStatic).Publish (True)
    
    
    ' Convert table into variable
    
    Dim table As Variant

    Dim MyFile As Scripting.TextStream
    Set MyFile = fso.OpenTextFile("O:\GMR\Processes\Downloads\Vügar\Working Student\Active Breach Reminder\test.html")

    table = MyFile.ReadAll

    ' Open outlook and paste table into email and display

    Dim outapp As Outlook.Application
    Set outapp = Outlook.Application

    Dim omail As Outlook.MailItem
    Set omail = outapp.CreateItem(olMailItem)


    Dim distb As String
    Dim counter As Integer

    With omail

    distb = ""
 
       For counter = 1 To 25
           If distb = "" Then
               distb = Cells(counter, 4).Value
           Else
               distb = distb & ";" & Cells(counter, 4).Value
           End If
           
       Next counter
       

        .To = distb
        .CC = "Guideline_Monitoring_Reporting@dummy.com"
        .Subject = "Active Breach Reminder - " & Date
        .HTMLBody = "Dear colleagues,<p>Please find below table <b>ACTIVE breaches</b> which are to be reviewed / commented by <b>today</b>:<p>" & "<table align = left>" & table & "<p>Regards, Victoria Nghiem"
        .Importance = olImportanceHigh
        .Display

    End With

    
    ActiveWorkbook.Close SaveChanges:=False


End Sub

Sub ARCHER_REPORTS_OPEN()
'
' Open_Arche_Reports Macro
'

Dim IE As New SHDocVw.InternetExplorer

IE.Visible = True

IE.Navigate "https://archer.blank.com/apps/ArcherApp/Home.aspx"


'
End Sub






