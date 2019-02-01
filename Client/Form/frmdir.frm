VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmdir 
   Caption         =   "파일 다운로드"
   ClientHeight    =   5535
   ClientLeft      =   60
   ClientTop       =   525
   ClientWidth     =   8655
   Icon            =   "frmdir.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5535
   ScaleWidth      =   8655
   StartUpPosition =   2  '화면 가운데
   Begin MSComctlLib.ImageList imglst 
      Left            =   240
      Top             =   4200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog cdfile 
      Left            =   120
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdload 
      Caption         =   "목록 다시읽기(&R)"
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6960
      TabIndex        =   1
      Top             =   5040
      Width           =   1575
   End
   Begin MSComctlLib.ListView lvfile 
      Height          =   4815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   8493
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imglst"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "파일 이름"
         Object.Width           =   3000
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "파일 경로"
         Object.Width           =   12347
      EndProperty
   End
End
Attribute VB_Name = "frmdir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'==============================================================='
'==============================================================='

Private Sub cmdload_Click() '파일 목록 다시읽기

    lvfile.ListItems.Clear
    frmclient.wsinfo.SendData "dir" '파일 목록 불러오기
    
End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Load()

    lvfile.ListItems.Clear
    frmclient.wsinfo.SendData "dir" '파일 목록 불러오기
    
End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Resize() '컨트롤 크기 조정

    If Not Me.WindowState = 1 Then
        lvfile.Width = Me.ScaleWidth - 200
        lvfile.Height = Me.ScaleHeight - (cmdload.Height + 350)
        cmdload.Left = (Me.ScaleWidth - cmdload.Width) - 100
        cmdload.Top = Me.ScaleHeight - (cmdload.Height + 100)
        pb.Width = lvfile.Width
    End If
    
End Sub

'==============================================================='
'==============================================================='

Private Sub lvfile_DblClick() '더블 클릭
    
    Dim i As Long
    Dim c, v As String
    
    c = lvfile.SelectedItem.ListSubItems(1).Text
    v = c
    
    '==========================================================='
    
    Do While Not InStr(c, "\") = "0" '\ 제거
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
    
    '==========================================================='
    
    i = MsgBox(c & "파일을 다운로드하시겠습니까?", vbYesNo, "다운로드")
    
    If i = 6 Then 'i = 예
        DirectDownload = True
        lvfile.Enabled = False
        cdfile.FileName = c
        cdfile.DialogTitle = "저장"
        cdfile.ShowSave
        
        '======================================================='
        
        If cdfile.FileName <> "" Then
            strFullFileName = Left(cdfile.FileName, Len(cdfile.FileName) - Len(cdfile.FileTitle))
            frmclient.DirOpen = strFullFileName
            frmclient.NameOpen = c
            frmclient.DirSend = lvfile.SelectedItem.ListSubItems(1).Text
            '#다운로드 시작
            frmclient.wsclient.SendData "download/" & lvfile.SelectedItem.ListSubItems(1).Text
            cdfile.FileName = "" '파일 이름 초기화
        End If
        
        '======================================================='
        
    End If

End Sub

