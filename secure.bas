Attribute VB_Name = "Api"
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
    Public Const LB_FINDSTRINGEXACT = &H1A2
   
' This ensures that no more than once is the same selection made
Function LBDupe(lpBox As ListBox) As Integer
    Dim nCount As Integer, nPos1 As Integer, nPos2 As Integer, nDelete As Integer
    Dim sText As String
    If lpBox.ListCount < 3 Then
        LBDupe = 0
        Exit Function
    End If
    For nCount = 0 To lpBox.ListCount - 1
       Do
                sText = lpBox.List(nCount)
                nPos1 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nCount, sText)
                nPos2 = SendMessageByString(lpBox.hwnd, LB_FINDSTRINGEXACT, nPos1 + 1, sText)
                If nPos2 = -1 Or nPos2 = nPos1 Then Exit Do
                lpBox.RemoveItem nPos2
                nDelete = nDelete + 1
            Loop
        Next nCount
        LBDupe = nDelete
End Function

