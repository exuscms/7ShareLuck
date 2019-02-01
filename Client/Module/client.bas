Attribute VB_Name = "client"
Public DirectDownload As Boolean
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, Optional ByVal lpParameters As String = vbNullString, Optional ByVal lpDirectory As String = vbNullString, Optional ByVal nShowCmd As Long = 5&) As Long

Public Function dly_Time(ByVal tmpTime As Single) As Boolean
    
    On Error Resume Next
    Dim dmyRespond
    Dim resFirst
    Dim resLast
    Dim Respond
    Dim tmpCount
    resFirst = Timer
    
    Do
        Respond = DoEvents()
        resLast = Timer
    Loop Until Abs((resLast - resFirst)) > tmpTime
    
    dly_Time = True

End Function

Public Function SendFile(strfilename As String)
    
    'On Error Resume Next
    Dim d() As Byte
    Dim i As Long
    ReDim d(1023)
    
    Open strfilename For Binary Access Read As #1
    
    For i = 1 To LOF(1) \ 1024
        Get #1, , d
        frmclient.wsdownload.SendData d
    Next
    
    ReDim d(0 To ((LOF(1) Mod 1024) - 1))
    Get #1, , d
    frmclient.wsdownload.SendData d
    
    Close #1
    Exit Function

End Function
