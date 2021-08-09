Attribute VB_Name = "Module6"
Sub AGENCY_LENDING_BNP()

'
' Agency_Lending_BNP Macro
'

'DEFINE THE KAG COLLATERAL EXCEL FILE

Dim MyPath As String
Dim MyFile As String
Dim deutschebank_file As String
MyPath = "O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\Files for BNP\"
If Right(MyPath, 1) <> "\" Then MyPath = MyPath & " \ "
MyFile = Dir(MyPath & "*.xls", vbNormal)
If Len(MyFile) = 0 Then
MsgBox "Kag Collateral Excel file was not found…", vbExclamation
Exit Sub
End If
Do While Len(MyFile) > 0
If Len(MyFile) = 29 Then deutschebank_file = MyFile
MyFile = Dir
Loop



' OPEN THE EXCEL FILE FROM HBSC AND HBSC TEMPLATE

Workbooks.Open ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\BNP_template.xls")
Workbooks.Open ("O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\POSI_ALL003_AGI_LENDING_COLLATERAL_POSITION_XLS.xls")


' COPY THE NECESSARY INFORMATION FROM HBSC INTO HBSC TEMPLATE


Workbooks("POSI_ALL003_AGI_LENDING_COLLATERAL_POSITION_XLS.xls").Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(2, 5).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("BNP_template.xls").Activate
ActiveWorkbook.Sheets("Depotbestande").Activate
ActiveSheet.Cells(4, 1).Select
Selection.PasteSpecial xlPasteValues

Workbooks("POSI_ALL003_AGI_LENDING_COLLATERAL_POSITION_XLS.xls").Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(2, 6).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("BNP_template.xls").Activate
ActiveWorkbook.Sheets("Depotbestande").Activate
ActiveSheet.Cells(4, 2).Activate
Selection.PasteSpecial xlPasteValues
Cells(4, 1).Select



' SORT AND CONSOLIDATE THE FIGURES FROM HBSC

Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select

With Selection

    .Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo

End With

Range("C4").Select
Selection.Consolidate Sources:= _
        "'O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\[BNP_template.xls]Depotbestande'!R4C1:R632C2" _
        , Function:=xlSum, TopRow:=False, LeftColumn:=True, CreateLinks:=False



' OPEN THE DEUTSCHE BANK FILE AND COPY THE INFO INTO HBSC TEMPLATE


Workbooks.Open MyPath & deutschebank_file

ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(9, 2).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("BNP_template.xls").Activate
ActiveWorkbook.Sheets(2).Activate
ActiveSheet.Cells(4, 1).Activate
Selection.PasteSpecial xlPasteValues

Workbooks(deutschebank_file).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(9, 10).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("BNP_template.xls").Activate
ActiveWorkbook.Sheets("KAG Collateral").Activate
ActiveSheet.Cells(4, 2).Activate
Selection.PasteSpecial xlPasteValues


' SORT AND CONSOLIDATE THE FIGURES FROM HBSC

Cells(4, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select

With Selection

    .Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo

End With

Range("C4").Select
Selection.Consolidate Sources:= _
        "'O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\[BNP_template.xls]KAG Collateral'!R4C1:R632C2" _
        , Function:=xlSum, TopRow:=False, LeftColumn:=True, CreateLinks:=False



' COPY THE FORMULA ON THE CHECK SHEET


Dim lastrow1 As Integer
Dim lastrow2 As Integer

lastrow1 = Sheets("KAG Collateral").Cells(4, 3).End(xlDown).Row
lastrow2 = Sheets("Depotbestande").Cells(4, 3).End(xlDown).Row


If lastrow1 > lastrow2 Then

Sheets(3).Activate
Range(Cells(4, 1), Cells(4, 11)).Select
Selection.Copy

Range(Cells(4, 1), Cells(lastrow1, 11)).Select
Selection.PasteSpecial xlPasteFormulas

End If


If lastrow1 < lastrow2 Then

Sheets(3).Activate
Range(Cells(4, 1), Cells(4, 11)).Select
Selection.Copy

Range(Cells(4, 1), Cells(lastrow2, 11)).Select
Selection.PasteSpecial xlPasteFormulas

End If


If lastrow1 = lastrow2 Then

Sheets(3).Activate
Range(Cells(4, 1), Cells(4, 11)).Select
Selection.Copy

Range(Cells(4, 1), Cells(lastrow1, 11)).Select
Selection.PasteSpecial xlPasteFormulas

End If


Cells(1, 1).Activate


' CLOSE THE WORKBOOKS

Workbooks("POSI_ALL003_AGI_LENDING_COLLATERAL_POSITION_XLS.xls").Close
Workbooks(deutschebank_file).Close


' SAVE THE TEMPLATE

Workbooks("BNP_template.xls").Activate
ActiveWorkbook.SaveAs ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\BNP\BNP_" & Format(Now(), "YYYYMMDD") & ".xls")


'

End Sub

Sub AGENCY_LENDING_HSBC()

'
' Agency_Lending_HSBC Macro


'DEFINE THE KAG COLLATERAL EXCEL FILE

Dim MyPath As String
Dim MyFile As String
Dim deutschebank_file As String
MyPath = "O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\Files for HSBC\"
If Right(MyPath, 1) <> "\" Then MyPath = MyPath & " \ "
MyFile = Dir(MyPath & "*.xls", vbNormal)
If Len(MyFile) = 0 Then
MsgBox "Kag Collateral Excel file was not found…", vbExclamation
Exit Sub
End If
Do While Len(MyFile) > 0
If Len(MyFile) = 35 Then deutschebank_file = MyFile
MyFile = Dir
Loop


' OPEN THE EXCEL FILE FROM HSBC AND HSBC TEMPLATE

Workbooks.Open ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\HSBC_template.xls")
Workbooks.Open ("O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\Depotbestände WP Leihe.xls")


' COPY THE NECESSARY INFORMATION FROM HSBC FILE INTO HSBC TEMPLATE


Workbooks("Depotbestände WP Leihe.xls").Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(2, 4).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("HSBC_template.xls").Activate
ActiveWorkbook.Sheets("Depotbestande").Activate
ActiveSheet.Cells(4, 1).Activate
Selection.PasteSpecial xlPasteValues

Workbooks("Depotbestände WP Leihe.xls").Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(2, 13).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("HSBC_template.xls").Activate
ActiveWorkbook.Sheets("Depotbestande").Activate
ActiveSheet.Cells(4, 2).Activate
Selection.PasteSpecial xlPasteValues
Cells(4, 1).Select



' SORT AND CONSOLIDATE THE FIGURES FROM HSBC

Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select

With Selection

    .Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo

End With

Range("C4").Select
Selection.Consolidate Sources:= _
        "'O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\[HSBC_template.xls]Depotbestande'!R4C1:R632C2" _
        , Function:=xlSum, TopRow:=False, LeftColumn:=True, CreateLinks:=False



' OPEN THE DEUTSCHE BANK FILE AND COPY THE INFO INTO HSBC TEMPLATE


Workbooks.Open MyPath & deutschebank_file

ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(9, 2).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("HSBC_template.xls").Activate
ActiveWorkbook.Sheets("KAG Collateral").Activate
ActiveSheet.Cells(4, 1).Activate
Selection.PasteSpecial xlPasteValues


Workbooks(deutschebank_file).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(9, 10).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("HSBC_template.xls").Activate
ActiveWorkbook.Sheets("KAG Collateral").Activate
ActiveSheet.Cells(4, 2).Activate
Selection.PasteSpecial xlPasteValues


' SORT AND CONSOLIDATE THE FIGURES FROM DEUTSCHEBANK

Cells(4, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select

With Selection

    .Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo

End With

Range("C4").Select
Selection.Consolidate Sources:= _
        "'O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\[HSBC_template.xls]KAG Collateral'!R4C1:R632C2" _
        , Function:=xlSum, TopRow:=False, LeftColumn:=True, CreateLinks:=False



' EXTEND THE THE FORMULA ON THE CHECK SHEET


Dim lastrow1 As Integer
Dim lastrow2 As Integer

lastrow1 = Sheets("KAG Collateral").Cells(4, 3).End(xlDown).Row
lastrow2 = Sheets("Depotbestande").Cells(4, 3).End(xlDown).Row


If lastrow1 > lastrow2 Then

Sheets(3).Activate
Range(Cells(4, 1), Cells(4, 11)).Select
Selection.Copy

Range(Cells(4, 1), Cells(lastrow1, 11)).Select
Selection.PasteSpecial xlPasteFormulas

End If


If lastrow1 < lastrow2 Then

Sheets(3).Activate
Range(Cells(4, 1), Cells(4, 11)).Select
Selection.Copy

Range(Cells(4, 1), Cells(lastrow2, 11)).Select
Selection.PasteSpecial xlPasteFormulas

End If


If lastrow1 = lastrow2 Then

Sheets(3).Activate
Range(Cells(4, 1), Cells(4, 11)).Select
Selection.Copy

Range(Cells(4, 1), Cells(lastrow1, 11)).Select
Selection.PasteSpecial xlPasteFormulas

End If


Cells(1, 1).Activate


' CLOSE THE WORKBOOKS

Workbooks("Depotbestände WP Leihe.xls").Close
Workbooks(deutschebank_file).Close

'SAVE THE TEMPLATE

Workbooks("HSBC_template.xls").Activate
ActiveWorkbook.SaveAs ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\HSBC\HSBC_" & Format(Now(), "YYYYMMDD") & ".xls")



End Sub

Sub AGENCY_LENDING_SPARKASSE()

'
' Sparkasse_Agency_Lending Macro
'


'DEFINE THE KAG COLLATERAL EXCEL FILE

Dim MyPath As String
Dim MyFile As String
Dim deutschebank_file As String
MyPath = "O:\GMR\Processes\Downloads\Vügar\Working Student\Sparkasse\Files for Sparkasse\"
If Right(MyPath, 1) <> "\" Then MyPath = MyPath & " \ "
MyFile = Dir(MyPath & "*.xls", vbNormal)
If Len(MyFile) = 0 Then
MsgBox "Kag Collateral Excel file was not found…", vbExclamation
Exit Sub
End If
Do While Len(MyFile) > 0
If Len(MyFile) = 38 Then deutschebank_file = MyFile
MyFile = Dir
Loop


' Open Sparkasse Template

Workbooks.Open ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\SpK KölnB_template.xls")


' OPEN THE DEUTSCHE BANK FILE AND COPY THE INFO INTO SpK TEMPLATE


Workbooks.Open MyPath & deutschebank_file

ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(9, 2).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("SpK KölnB_template.xls").Activate
ActiveWorkbook.Sheets("KAG Collateral").Activate
ActiveSheet.Cells(4, 1).Activate
Selection.PasteSpecial xlPasteValues

Workbooks(deutschebank_file).Activate
ActiveWorkbook.Sheets(1).Activate
ActiveSheet.Cells(9, 10).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Workbooks("SpK KölnB_template.xls").Activate
ActiveWorkbook.Sheets("KAG Collateral").Activate
ActiveSheet.Cells(4, 2).Activate
Selection.PasteSpecial xlPasteValues


' SORT AND CONSOLIDATE THE FIGURES FROM SPK
Cells(4, 1).Select
Range(Selection, Selection.End(xlDown)).Select
Range(Selection, Selection.End(xlToRight)).Select

With Selection

    .Sort key1:=Range("A1"), _
      order1:=xlAscending, Header:=xlNo

End With

Range("C4").Select
Selection.Consolidate Sources:= _
        "'O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\[SpK KölnB_template.xls]KAG Collateral'!R4C1:R632C2" _
        , Function:=xlSum, TopRow:=False, LeftColumn:=True, CreateLinks:=False



MsgBox "Please check the mailbox for the attachment. The Password is 'Bofferding123'"

Dim IE As New SHDocVw.InternetExplorer

IE.Visible = True
IE.Navigate "https://securemail.sparkasse.de/sparkasse-koelnbonn//login.jsp?username=surveillance@allianzgi.com"




' CLOSE THE WORKBOOKS
Workbooks(deutschebank_file).Close


' SAVE THE TEMPLATE

Workbooks("SpK KölnB_template.xls").Activate
ActiveWorkbook.SaveAs ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\SpK KölnB\SpK KölnB_" & Format(Now(), "YYYYMMDD"))

'

End Sub
