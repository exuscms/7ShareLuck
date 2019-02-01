Attribute VB_Name = "server"
Public Strtext As String '클라이언트 퇴장 코멘트
Public Cntclient, Cntclient2, Cntclient3, pw As Integer '클라이언트 총 갯수
Public File_Size, k
Public Sendfilechk As Boolean
Public SizeLimit As Long

'==============================================================='
'==============================================================='

Public Function Log(Info, str) 'Show Log
frmserver.lvlog.ListItems.Add , , Info
frmserver.lvlog.ListItems.Item(frmserver.lvlog.ListItems.Count).ListSubItems.Add , , str
frmserver.lvlog.ListItems(frmserver.lvlog.ListItems.Count).Selected = True
frmserver.lvlog.ListItems(frmserver.lvlog.ListItems.Count).EnsureVisible
End Function

'==============================================================='
'==============================================================='

Public Function Remote(Info, str)
frmserver.lvremote.ListItems.Add , , Info
frmserver.lvremote.ListItems.Item(frmserver.lvremote.ListItems.Count).ListSubItems.Add , , str
frmserver.lvremote.ListItems(frmserver.lvremote.ListItems.Count).Selected = True
frmserver.lvremote.ListItems(frmserver.lvremote.ListItems.Count).EnsureVisible
End Function

'==============================================================='
'==============================================================='

Public Function Share(Info, str)
frmserver.lvshare.ListItems.Add , , Info
frmserver.lvshare.ListItems.Item(frmserver.lvshare.ListItems.Count).ListSubItems.Add , , str
frmserver.lvshare.ListItems(frmserver.lvshare.ListItems.Count).Selected = True
frmserver.lvshare.ListItems(frmserver.lvshare.ListItems.Count).EnsureVisible
End Function

'==============================================================='
'==============================================================='

Public Function Chat(Info, str)
frmserver.lvchat.ListItems.Add , , Info
frmserver.lvchat.ListItems.Item(frmserver.lvchat.ListItems.Count).ListSubItems.Add , , str
frmserver.lvchat.ListItems(frmserver.lvchat.ListItems.Count).Selected = True
frmserver.lvchat.ListItems(frmserver.lvchat.ListItems.Count).EnsureVisible
End Function


'==============================================================='
'==============================================================='

'지연
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

'==============================================================='
'==============================================================='

'파일 전송
Public Function SendFile(strfilename As String, ByVal Index As Integer)

    On Error Resume Next
    Dim d() As Byte
    Dim i As Long
    ReDim d(1023) '1024Kb
    
    Open strfilename For Binary Access Read As #1
    
    For i = 1 To LOF(1) \ 1024
        Get #1, , d
        
        'If frmserver.wsupload(Index).State = sckConnected Then
            frmserver.wsupload(Index).SendData d
        'End If
        
    Next
    
    ReDim d(0 To ((LOF(1) Mod 1024) - 1))
    Get #1, , d
    
    'If frmserver.wsupload(Index).State = sckConnected Then
        frmserver.wsupload(Index).SendData d
    'End If
    
    Close #1
    
    Sendfilechk = False
    
    'Exit Function
    
'SendNextFile:
    'SendFile strfilename, Index + 1
    
End Function
