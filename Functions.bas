Attribute VB_Name = "Functions"
Dim File As String
Dim intFileNum As Integer
Dim strTextLine As String
Dim b() As Byte


Function ClearAllFields()
' Clears All fields
frmMain.txtURL.Text = ""
frmMain.txtTitle.Text = ""
frmMain.txtDes.Text = ""
frmMain.k1.Text = ""
frmMain.k2.Text = ""
frmMain.k3.Text = ""
frmMain.k4.Text = ""
frmMain.k5.Text = ""
frmMain.k6.Text = ""
frmMain.k7.Text = ""
frmMain.k8.Text = ""
frmMain.k9.Text = ""
frmMain.k10.Text = ""
frmMain.k11.Text = ""
frmMain.k12.Text = ""
frmMain.k13.Text = ""
frmMain.k14.Text = ""
frmMain.k15.Text = ""
frmMain.k16.Text = ""
frmMain.k17.Text = ""
frmMain.k18.Text = ""
frmMain.txtName.Text = ""
frmMain.txtCompany.Text = ""
frmMain.txtCity.Text = ""
frmMain.txtCountry.Text = ""
frmMain.txtMail.Text = ""
frmMain.txtAddress.Text = ""
frmMain.txtProvince.Text = ""
frmMain.txtPostal.Text = ""
frmMain.txtPhone.Text = ""
End Function

Function ExitApp()
'Exit application
Unload frmMain
End Function


Function LoadCategories()
' Loads the list of categories
Dim intFileNum As Integer
Dim strTextLine As String
intFileNum = FreeFile
Open App.Path & "\data\categ.txt" For Input As #intFileNum
Do While Not EOF(intFileNum)
Line Input #intFileNum, strTextLine
frmMain.CoCategory.AddItem strTextLine
Loop
Close #intFileNum
End Function

Function RtfLoadCall()
' Loads some info to the RTF
Dim intFileNum As Integer
Dim strTextLine As String
intFileNum = FreeFile
Open App.Path & "\data\rtfload.txt" For Input As #intFileNum
Do While Not EOF(intFileNum)
Line Input #intFileNum, strTextLine
frmMain.Rtf.TextRTF = frmMain.Rtf.TextRTF & strTextLine
Loop
Close #intFileNum
End Function




Function LoadAvailableEngines()
' Speaks for it self
frmMain.AvailableEngines.AddItem "Altavista.com"
frmMain.AvailableEngines.AddItem "AllAmericasBest.com"
frmMain.AvailableEngines.AddItem "DirectHit.com"
frmMain.AvailableEngines.AddItem "EuroSeek.com"
frmMain.AvailableEngines.AddItem "Excite.com"
frmMain.AvailableEngines.AddItem "HotBot.com"
frmMain.AvailableEngines.AddItem "InfoSeek.com"
frmMain.AvailableEngines.AddItem "Lycos.com"
frmMain.AvailableEngines.AddItem "MSN.com"
frmMain.AvailableEngines.AddItem "WebCrawler.com"
End Function

' -----------------------------------------------------------------
' We don't need this for now.
' -----------------------------------------------------------------
'Function CheckInfoFields()
' This function only checks wether there is enough information _
  to start submitions.
'ExitThisFunction = 0
'If frmMain.txtURL.Text = "" Then GoTo Err:
'If frmMain.txtTitle.Text = "" Then GoTo Err:
'If frmMain.txtDes.Text = "" Then GoTo Err:
'If frmMain.k1.Text = "" Then GoTo Err:
'If frmMain.k2.Text = "" Then GoTo Err:
'If frmMain.txtName.Text = "" Then GoTo Err:
'If frmMain.txtCompany.Text = "" Then GoTo Err:
'If frmMain.txtCity.Text = "" Then GoTo Err:
'If frmMain.txtCountry.Text = "" Then GoTo Err:
'If frmMain.txtMail.Text = "" Then GoTo Err:
'If frmMain.txtAddress.Text = "" Then GoTo Err:
'If frmMain.txtProvince.Text = "" Then GoTo Err:
'If frmMain.txtPostal.Text = "" Then GoTo Err:
'If frmMain.txtPhone.Text = "" Then GoTo Err:
' If not the goto Tab0 (WebSite Data) and fill it up
'Err:
'  frmMain.lblEngine.Caption = "None"
'  Scrap.lblCheck.Caption = 0
'  frmMain.SSTab1.Tab = 0
'  Exit Function
'End Function
' -----------------------------------------------------------------

Function CheckForNewVersionOnTheNet()


    'Open Free File
    File = FreeFile
    'set protocol to HTTP
    Scrap.Inet1.Protocol = icHTTP
    'set URL for initialization File
    Scrap.Inet1.URL = "http://www.BohemiaTrading.com/update.txt"
    ' Retrieve the HTML data into a byte array.
    b() = Scrap.Inet1.OpenURL(Scrap.Inet1.URL, icByteArray)
    ' Create a local file from the retrieved data.
    Open "c:\update.txt" For Binary Access Write As #File  'App.Path & "update.txt"
    Put #File, , b()
    Close #File

    intFileNum = FreeFile
    Open "c:\update.txt" For Input As #intFileNum ' App.Path & "update.txt"
    Do While Not EOF(intFileNum)
    Line Input #intFileNum, strTextLine
    Loop
    Close #intFileNum

    'Kill= app.Path & "\update.txt"   (Theoreticly)
    
    ' strTextLine is a 1st and only line in file which contains either "1" or "0" _
    1=Yes Go for Update, 0=No don't update
    

    If strTextLine = 1 Then
    ' Update
    frmMain.mnuUpdate.Enabled = True
    MsgBox "Update Available." & vbCrLf & "Please Select Live Update from Menu."
    Else
    'Dont's update
MsgBox "No Update Available at this time."
frmMain.mnuUpdate.Enabled = False
    End If
End Function
