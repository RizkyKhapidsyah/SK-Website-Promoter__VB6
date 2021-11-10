Attribute VB_Name = "Banner"
Option Explicit
Public Const URL = "http://www.bohemiatrading.com/download/tools/Tool.exe"
Public Const email = "friend's@email.com?subject=Check Out This Application"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Const SW_SHOWNORMAL = 1

Public Sub GotoWebAndDownload()
Dim Success As Long
Success = ShellExecute(0&, vbNullString, URL, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

Public Sub sendemail()
Dim Success As Long
Success = ShellExecute(0&, vbNullString, "mailto:" & email, vbNullString, "C:\", SW_SHOWNORMAL)
End Sub

  

