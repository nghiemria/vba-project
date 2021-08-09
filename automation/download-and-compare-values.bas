Attribute VB_Name = "Module5"
Sub ATTACHMENTS_DOWNLOAD()

'
' Download_Attachements Macro
'
' Variables

Dim ol As Outlook.Application
Dim ns As Outlook.Namespace
Dim fol As Outlook.Folder
Dim fol1 As Outlook.Folder
Dim i As Object
Dim mi As Outlook.MailItem
Dim n As Long
Dim at As Outlook.Attachment
Dim ShellApp As Object

Set ShellApp = CreateObject("Shell.Application")
Set ol = New Outlook.Application
Set ns = ol.GetNamespace("MAPI")
Set fol = ns.GetDefaultFolder(olPublicFoldersAllPublicFolders).Folders("KAG").Folders("Units").Folders("Guideline Monitoring").Folders("Surveillance@dummy.de")
Set fol1 = ns.GetDefaultFolder(olPublicFoldersAllPublicFolders).Folders("KAG").Folders("Gruppenfaxe").Folders("IC").Folders("Ueberwachung (Fax 14822)")
Dim d As Date
Dim day As Integer
d = Date
day = Weekday(Date, vbMonday)


'Status

oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "Please wait while macro saves all attachments from mailbox..."


'Download attachements for HSBC, BNP, Sparkasse, Thresholds and Money Market

'Ueberwachung Fax

For Each i In fol1.Items
    If i.Class = olMail Then
        Set mi = i
        
        If mi.ReceivedTime > d And mi.Subject = "A fax has arrived from remote ID ''." And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\Thresholds\" & at.Filename
                            
                Next at
                
        End If
End If
    
Next i


'HSBC, BNP, Sparkasse,and Money Market


For Each i In fol.Items
    If i.Class = olMail Then
        Set mi = i
        
      'Money Market___________
        
        If mi.ReceivedTime > d And mi.Subject = "OPICS-Confirmation" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\Money Market\" & at.Filename
                            
                Next at
                
        End If
        
        
        
      'HSBC_____________
         
        If mi.ReceivedTime > d And mi.Subject = "Depotbestände AGI WP Leihe Coll Depots" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\" & at.Filename
                            
                Next at
                
        End If
            
        If mi.ReceivedTime > d - 4 And mi.Subject = "Deutsche Bank Reports - AGI at Trinkaus Compliance reports" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\" & at.Filename
                            
                Next at
                
        End If
        
    'BNP_____________
    
        If mi.ReceivedTime > d And mi.Subject = "Customised - POSI_ALL003 - AGI LENDING COLLATERAL POSITION XLS" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\" & at.Filename
                            
                Next at
                
        End If
            
        If mi.ReceivedTime > d - 4 And mi.Subject = "Deutsche Bank Reports - AGI at BNP compliance reports" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\" & at.Filename
                            
                Next at
                
        End If
        
            
     'Sparkasse__________
    
        If mi.ReceivedTime > d - 4 And mi.Subject = "Deutsche Bank Reports - KE Sparkasse Reporting" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\Sparkasse\" & at.Filename
                            
                Next at
                
        End If
    
    End If
    
    
    'Breach type technical


If day = 1 Then



        If i.Class = olMail Then
            Set mi = i
        
            If mi.ReceivedTime > d And mi.Subject = "[PROD] BreachType Technical" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\Technical Alerts\" & at.Filename
                            
                Next at
                
            End If
        End If
   
Else

   
        If i.Class = olMail Then
            Set mi = i
        
            If mi.ReceivedTime = d - 2 And mi.Subject = "[PROD] BreachType Technical" And mi.Attachments.Count > 0 Then
                
                For Each at In mi.Attachments
                
                     at.SaveAsFile "O:\GMR\Processes\Downloads\Vügar\Working Student\Technical Alerts\" & at.Filename
                            
                Next at
                
            End If
        End If
    
   
End If
    
Next i





' UNZIP HSBC ZIP FILES. THIS IS NECESSARY FOR THE AGENCY LENGING RECONCILIATION FOR THE HSBC

ShellApp.Namespace("O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\Files for HSBC").CopyHere ShellApp.Namespace("O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\AgencyLendingReports14726.zip").Items


' UNZIP BNP ZIP FILES. THIS IS NECESSARY FOR THE AGENCY LENGING RECONCILIATION FOR THE BNP

ShellApp.Namespace("O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\Files for BNP").CopyHere ShellApp.Namespace("O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\AgencyLendingReports14486.zip").Items


' UNZIP SpK ZIP FILES. THIS IS NECESSARY FOR THE AGENCY LENGING RECONCILIATION FOR THE SpK

ShellApp.Namespace("O:\GMR\Processes\Downloads\Vügar\Working Student\Sparkasse\Files for Sparkasse").CopyHere ShellApp.Namespace("O:\GMR\Processes\Downloads\Vügar\Working Student\Sparkasse\AgencyLendingReports11101.zip").Items



Application.StatusBar = False
Application.DisplayStatusBar = oldStatusBar

oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "Ready"
Application.StatusBar = False

End Sub

Sub FOLDER_CLEAN()

'Status

oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "Please wait while macro deletes all attachments from the folder..."


On Error Resume Next
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\HSBC\Files for HSBC\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\BNP\Files for BNP\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Sparkasse\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Sparkasse\Files for Sparkasse\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Money Market\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Thresholds\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Technical Alerts\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Long-Lasting Technicals\*.*"
Kill "O:\GMR\Processes\Downloads\Vügar\Working Student\Active Breach Reminder\*.*"
On Error GoTo 0


oldStatusBar = Application.DisplayStatusBar
Application.DisplayStatusBar = True
Application.StatusBar = "Ready"
Application.StatusBar = False


End Sub



Sub REPORT_DISCREPANCY()

Dim outapp As Outlook.Application
Dim outemail As Outlook.MailItem

Set outapp = New Outlook.Application
Set outemail = outapp.CreateItem(olMailItem)

Dim name As String

name = ActiveWorkbook.name

If Len(name) = 16 Then


ActiveWorkbook.SaveAs ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\BNP\BNP_" & Format(Now(), "YYYYMMDD") & "v1" & ".xls")

With outemail

.To = "dummy@bnpparibas.com"
.CC = "AG-Verletzungen@dummy.de"
.Subject = "Depotbestände AGI WP Leihe Coll Depots"
.HTMLBody = "Dear sir/madam!<p>We found discrepancy while reconcile internally your figures with Deutsche Bank figures.<p>The ISIN number below and the amount, which is shown in Deutsche Bank does not exists  in BNP records. <p> Account ID: 41329975721030029033X<p>Could you please take a look at the file and come back to us as soon as possible?<p>Thanks in advance!"
.Attachments.Add ActiveWorkbook.FullName
.Display

End With


End If


If Len(name) = 17 Then

ActiveWorkbook.SaveAs ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\HSBC\HSBC_" & Format(Now(), "YYYYMMDD") & "v1" & ".xls")

With outemail

.To = "dummy@hsbc.de"
.CC = "AG-Verletzungen@dummy.de"
.Subject = "Depotbestände AGI WP Leihe Coll Depots"
.HTMLBody = "Dear sir/madam!<p>We found discrepancy while reconcile internally your figures with Deutsche Bank figures.<p>The ISIN number below and the amount, which is shown in Deutsche Bank does not exists  in HSBC records.<p>Could you please take a look at the issue and come bakc to us as soon as possible?<p>Thanks in advance!"
.Attachments.Add ActiveWorkbook.FullName
.Display

End With

End If


If Len(name) = 22 Then

ActiveWorkbook.SaveAs ("O:\GMR\Processes\Germany\M_Agency_Lending_Deutsche_Bank\Agency Lending\SpK KölnB\SpK KölnB_" & Format(Now(), "YYYYMMDD") & "v1" & ".xls")


With outemail

.To = "pk-dummy@sparkasse-koelnbonn.de"
.CC = "AG-Verletzungen@dummy.de"
.Subject = "Depotbestände AGI WP Leihe Coll Depots"
.HTMLBody = "Dear sir/madam!<p>We found discrepancy while reconcile internally your figures with Deutsche Bank figures.<p>The ISIN number below and the amount, which is shown in Deutsche Bank does not exists  in Sparkasse records.<p>Could you please take a look at the issue and come bakc to us as soon as possible?<p>Thanks in advance!"
.Attachments.Add ActiveWorkbook.FullName
.Display

End With

End If


End Sub


