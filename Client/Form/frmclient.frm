VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmclient 
   BorderStyle     =   1  '단일 고정
   Caption         =   "7 Share Luck"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   4095
   BeginProperty Font 
      Name            =   "돋움"
      Size            =   8.25
      Charset         =   129
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmclient.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4095
   StartUpPosition =   2  '화면 가운데
   Begin VB.TextBox txtnick 
      Alignment       =   2  '가운데 맞춤
      Appearance      =   0  '평면
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '없음
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   960
      TabIndex        =   8
      Text            =   "이름없음"
      Top             =   1440
      Width           =   2895
   End
   Begin VB.CommandButton cmdremote 
      Caption         =   "접속(&C)"
      Height          =   375
      Left            =   2880
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Frame fmremote 
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   3855
      Begin MSWinsockLib.Winsock wsdownload 
         Left            =   -240
         Top             =   -120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin MSWinsockLib.Winsock wsclient 
         Left            =   -240
         Top             =   -240
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer Timer 
         Interval        =   5000
         Left            =   -240
         Top             =   -240
      End
      Begin VB.TextBox txtip 
         Alignment       =   2  '가운데 맞춤
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   480
         TabIndex        =   2
         Text            =   "116.127.5.252"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtport 
         Alignment       =   2  '가운데 맞춤
         Appearance      =   0  '평면
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  '없음
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2400
         TabIndex        =   1
         Text            =   "12345"
         Top             =   240
         Width           =   1335
      End
      Begin MSWinsockLib.Winsock wsinfo 
         Left            =   -240
         Top             =   -120
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Label labnick 
         AutoSize        =   -1  'True
         Caption         =   "Nick :"
         Height          =   165
         Left            =   240
         TabIndex        =   7
         Top             =   780
         Width           =   450
      End
      Begin VB.Label lbip 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "IP :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   225
      End
      Begin VB.Label lbport 
         Alignment       =   2  '가운데 맞춤
         AutoSize        =   -1  'True
         BackStyle       =   0  '투명
         Caption         =   "Port :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   390
      End
   End
   Begin VB.Label lbtop 
      AutoSize        =   -1  'True
      BackStyle       =   0  '투명
      Caption         =   "서버 접속"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   12
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   120
      TabIndex        =   9
      Top             =   240
      Width           =   1035
   End
   Begin VB.Image imgtop 
      Height          =   675
      Left            =   0
      Picture         =   "frmclient.frx":0ECA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblocal 
      AutoSize        =   -1  'True
      Caption         =   "ip"
      Height          =   165
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   150
   End
End
Attribute VB_Name = "frmclient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public DirSend, NameOpen, DirOpen As String

Dim file_size '파일 크기

'==============================================================='
'==============================================================='

Private Sub cmdremote_Click() '접속

    If txtnick <> "" Then
        wsclient.Close
        wsclient.RemoteHost = txtip
        wsclient.RemotePort = txtport
        wsclient.Connect '채팅 클라이언트 접속
        wsinfo.Close
        wsinfo.RemoteHost = txtip
        wsinfo.RemotePort = 12347
        wsinfo.Connect '파일 목록 클라이언트 접속
        wsdownload.Close
        wsdownload.LocalPort = 12346
        wsdownload.Listen '다운로드 서버 열기
    End If
    
End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Load()

    DirectDownload = False
    lblocal = wsclient.LocalIP 'PC IP

End Sub

Private Sub Form_Unload(Cancel As Integer)

midiOutClose hMidi

End Sub

'==============================================================='
'==============================================================='

Private Sub lblocal_Click() 'IP 더블클릭

    txtip.Text = lblocal 'IP 삽입

End Sub

'==============================================================='
'==============================================================='

Private Sub wsclient_close() '서버와 접속 끎힘

    MsgBox "서버와의 접속이 끎혔습니다."
    frmclient.wsclient.Close
    frmclient.wsdownload.Close
    End
    
End Sub

'==============================================================='
'==============================================================='

Private Sub wsclient_Connect() '접속 완료

    frmchat.Show '채팅창 열기
    Me.Hide
    
End Sub

'==============================================================='
'==============================================================='

Private Sub wsclient_DataArrival(ByVal bytesTotal As Long) '채팅 클라이언트 파일 전송 완료
    
    On Error Resume Next
    Dim i As Long
    Dim strdata As String
    Dim strsplit() As String
    
    wsclient.GetData strdata '클라이언트 데이터
    
    strsplit() = Split(strdata, "/") '데이터 분할
    
    '#바운드 사이즈 (1)
    If UBound(strsplit) = 1 Then
    
        Select Case strsplit(0)
        
            Case "ban" '강제 퇴장
            
                wsclient.Close
                MsgBox "서버에서 강제 퇴장당하셨습니다." & vbCrLf & vbCrLf & "강제 퇴장 사유 : " & vbCrLf & vbCrLf & strsplit(1), vbInformation, "퇴장"
                End
                
            Case "sends" '중복 파일 전송 방지
            
                frmchat.tbtool.Buttons(2).Enabled = False
                
            Case "name" '파일 이름 불러오기
            
                frmchat.names = strsplit(1)
                
            Case "logcnt" '접속자 불러오기
            
                frmchat.Chat "총 접속자 : " & strsplit(1), 2
                frmclient.wsclient.SendData "memlst"
                
            Case "admin" '관리자 메시지
            
                frmchat.Chat "관리자>>" & strsplit(1), 4
                
            Case "lens" '파일 크기 불러오기
            
                file_size = Val(strsplit(1))
                
            Case "len" '파일 크기 & 파일 전송 시작
            
                file_size = Val(strsplit(1))
                wsclient.SendData "start/" & DirSend

            Case "setdir" '파일 목록 불러오기 전 준비
            
                frmdir.lvfile.ListItems.Clear
                wsclient.SendData "dir"
                
            Case "dir" '파일 목록 불러오기
            
                strsplit = Split(Mid(strdata, 4, (Len(strdata) - 3)), "?") '/로 분할
                
                For i = 0 To UBound(strsplit)
                
                    If strsplit(i) <> "" Then '내용이 존재한다면
                        frmdir.lvfile.ListItems.Add , , strsplit(i) '목록에 추과
                    End If
                    
                Next i
                
            Case "msg" '채팅 메시지
            
                frmchat.Chat strsplit(1), 3 '채팅 목록 추과
                frmchat.SetFocus
    End Select
        
    End If

    If UBound(strsplit) = 2 Then
        If strsplit(0) = "pplay" Then
            InitializInstrument Val(strsplit(2))
            PlayNote strsplit(1)
        ElseIf strsplit(0) = "splay" Then
            StopNote strsplit(1)
        End If
    End If
End Sub

'==============================================================='
'==============================================================='

Private Sub wsdownload_ConnectionRequest(ByVal requestID As Long) '접속 요청

    'On Error GoTo ErrOverlab

        wsdownload.Close

    wsdownload.Accept requestID

    Exit Sub
    
ErrOverlab:
    
    MsgBox ("클라이언트 중복 사용!")
    
    End
    
End Sub

'==============================================================='
'==============================================================='

Private Sub wsdownload_DataArrival(ByVal bytesTotal As Long) '서버에서 파일 다운로드
    
    On Error GoTo ErrOp
    
    Dim d() As Byte
    Dim k, a, b
    
    wsdownload.GetData d '데이터 얻기
    
    Open DirOpen & "\" & NameOpen For Binary Access Write As #1 '파일 저장
        Put #1, LOF(1) + 1, d
        k = LOF(1)
    Close #1
    
    If k >= file_size Then '다운로드 완료
    
        MsgBox frmchat.names & " 다운로드 완료!"
        
        frmchat.tbtool.Buttons(2).Enabled = True
        frmdir.lvfile.Enabled = True
    
        a = frmchat.names '파일이름에서 확장명 구하기
    
        Do While Not InStr(a, ".") = "0"
            DoEvents
            
            If InStr(1, a, ".") Then
                b = InStr(1, a, ".")
                If b >= 2 Then
                    a = Right(v, Len(a) - (b - 1))
                ElseIf w = 1 Then
                    a = Right(v, Len(a) - (b))
                End If
            End If
            
        Loop
        
        frmchat.Chat frmchat.names, 1

        '======================================================='
        '======================================================='
        
        If Not Dir(App.Path & "\down\" & frmchat.names) <> "" And DirectDownload = False Then '파일 이름 복구
            Name DirOpen & "\" & NameOpen As App.Path & "\down\" & frmchat.names
        ElseIf Dir(App.Path & "\down\" & frmchat.names) <> "" And DirectDownload = False Then
            Kill App.Path & "\down\" & frmchat.names
            Name DirOpen & "\" & NameOpen As App.Path & "\down\" & frmchat.names
        End If
        
        DirectDownload = False
        
        '========================================================='
        
        End If

    Exit Sub
    
ErrOp:

End Sub

'==============================================================='
'==============================================================='

Private Sub wsinfo_DataArrival(ByVal bytesTotal As Long) '파일 목록 불러오기
'On Error Resume Next
Dim c, v
Dim strdata As String
Dim strsplit() As String

wsinfo.GetData strdata

If strdata = "senddingg" Then '파일 중복 전송 방지

    If ready = False Then
        frmchat.tbtool.Buttons(2).Enabled = False
    End If
    
End If

If InStr(strdata, "<") > 0 Then '#사용자 목록일 경우

    frmchat.lstmember.ListItems.Clear
    
    strsplit = Split(strdata, "<")
    
    For i = 0 To UBound(strsplit)
    
        frmchat.lstmember.ListItems.Add , , strsplit(i)
        
    Next i
    
End If

'체크 (?) ======================================================'

If InStr(strdata, "?") > 0 Then '#파일 목록일 경우

    strsplit = Split(strdata, "?")
    
    For i = 0 To UBound(strsplit)
    c = strsplit(i)
    v = c
    
    '==========================================================='
    '체크 (\) =================================================='
    
    Do While Not InStr(c, "\") = "0"
        DoEvents
    
        If InStr(1, c, "\") Then
            w = InStr(1, c, "\")
            If w >= 2 Then
                c = Right(v, Len(c) - (w - 1))
            ElseIf w = 1 Then
                c = Right(v, Len(c) - (w))
            End If
        End If
            
    Loop
    
    a = c
    
    '체크 (.) =================================================='
    
    Do While Not InStr(a, ".") = "0"
        DoEvents
        
        If InStr(1, a, ".") Then
            b = InStr(1, a, ".")
            If b >= 2 Then
                a = Right(v, Len(a) - (b - 1))
            ElseIf w = 1 Then
                a = Right(v, Len(a) - (b))
            End If
        End If
        
    Loop
    
    frmdir.lvfile.ListItems.Add , , c
    frmdir.lvfile.ListItems(frmdir.lvfile.ListItems.Count).ListSubItems.Add , , strsplit(i)
    
    Next i
    
    '==========================================================='
    '==========================================================='
    
Else '파일 다운로드 전 준비

    strsplit = Split(strdata, "/")

    If UBound(strsplit) = 2 Then
    
        '======================================================='
        '(View) ================================================'
        
        If strsplit(0) = "view" Then
        
            frmchat.Chat "파일 다운로드중...", 5

            If Dir(App.Path & "\down\" & strsplit(2)) <> "" Then  '#파일 중복 다운로드 방지
                Kill App.Path & "\down\" & strsplit(2)
            End If
            
            If Dir(App.Path & "\down\tmp.dat") <> "" Then '#파일 중복 다운로드 방지
                Kill App.Path & "\down\tmp.dat"
            End If
            
            If Not Dir(App.Path & "\down", vbDirectory) <> "" Then '#다운로드 폴더 생성
                MkDir App.Path & "\down"
            End If
            
            If frmchat.ready = False Then '#중복 전송 방지
                frmchat.tbtool.Buttons(2).Enabled = False
            End If
            
            frmchat.names = strsplit(2) '#원 파일 이름 불러오기
            
            frmclient.DirOpen = App.Path
            frmclient.NameOpen = "\down" & "\tmp.dat"
            
            file_size = Val(strsplit(1)) '파일 사이즈 불러오기
            
        End If
        
        '======================================================='
        '======================================================='
        
    End If
    
End If
End Sub
