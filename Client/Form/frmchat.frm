VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmchat 
   Caption         =   "7 Share Luck - Chat"
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   540
   ClientWidth     =   7530
   Icon            =   "frmchat.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7530
   StartUpPosition =   2  'ȭ�� ���
   Begin MSComctlLib.ImageList k 
      Left            =   0
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":058A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":0924
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":0CBE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbtool 
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   4320
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "k"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstmember 
      Height          =   3615
      Left            =   0
      TabIndex        =   3
      Top             =   675
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   6376
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "il"
      ForeColor       =   -2147483640
      BackColor       =   8421504
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "���"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstmsg 
      Height          =   3615
      Left            =   1680
      TabIndex        =   2
      Top             =   675
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   6376
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ä��"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList c 
      Left            =   0
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":1258
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":15F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":198C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":1D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmchat.frx":20C0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsmusic 
      Left            =   120
      Top             =   840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtchat 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      TabIndex        =   0
      Top             =   4680
      Width           =   6615
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   135
      Left            =   1800
      TabIndex        =   1
      Top             =   5040
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image imgDP 
      Appearance      =   0  '���
      BorderStyle     =   1  '���� ����
      Height          =   480
      Index           =   0
      Left            =   120
      Picture         =   "frmchat.frx":23DA
      Stretch         =   -1  'True
      ToolTipText     =   "Contact's Display Picture"
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lbip 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   720
      TabIndex        =   5
      ToolTipText     =   "Your IP"
      Top             =   360
      Width           =   60
   End
   Begin VB.Label lbnick 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   720
      TabIndex        =   4
      ToolTipText     =   "Your Nickname"
      Top             =   120
      Width           =   60
   End
   Begin VB.Image imgDP 
      Appearance      =   0  '���
      BorderStyle     =   1  '���� ����
      Height          =   480
      Index           =   1
      Left            =   120
      Stretch         =   -1  'True
      ToolTipText     =   "Your Display Picture"
      Top             =   120
      Width           =   480
   End
   Begin VB.Image imgtop 
      Height          =   675
      Left            =   0
      Picture         =   "frmchat.frx":2EE4
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6735
   End
End
Attribute VB_Name = "frmchat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public names
Public ready As Boolean

'==============================================================='
'==============================================================='

Public Function Chat(txtstr As String, icoint As Integer)

    frmchat.lstmsg.ListItems.Add , , txtstr ', , icoint
    frmchat.lstmsg.ListItems(frmchat.lstmsg.ListItems.Count).Selected = True
    frmchat.lstmsg.SelectedItem.EnsureVisible
    If icoint = 2 Then
        frmchat.lstmsg.SelectedItem.Bold = True
        frmchat.lstmsg.SelectedItem.ForeColor = vbGreen
    ElseIf icoint = 4 Then
        frmchat.lstmsg.SelectedItem.Bold = True
        frmchat.lstmsg.SelectedItem.ForeColor = vbBlue
    ElseIf icoint = 5 Then
        frmchat.lstmsg.SelectedItem.Bold = True
        frmchat.lstmsg.SelectedItem.ForeColor = vbYellow
    End If
End Function

Private Sub Form_Load()

    frmclient.wsclient.SendData "request/" & frmclient.txtnick '���� ��û(���� �ߺ� �ٿ�ε� üũ)
    lbip.Caption = frmclient.wsclient.LocalIP
    lbnick.Caption = frmclient.txtnick
    
End Sub


'==============================================================='
'==============================================================='

Private Sub lstmsg_DblClick()

    On Error Resume Next

    Dim strName As String, strFolder As String, strOption As String
    
    strName = lstmsg.SelectedItem.Text
    strFolder = App.Path & "\down\"
    strOption = "-window"

    ShellExecute Me.hwnd, vbNullString, strFolder & "\" & strName, strOption, strFolder, 1

End Sub

Private Sub mnuclear_Click()
    lstmsg.ListItems.Clear
End Sub

'==============================================================='
'==============================================================='

Private Sub mnudownload_Click() '���� �ٿ�ε� â
    frmdir.Show
End Sub

Private Sub tbtool_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        frmdir.Show
    Case 2
            
    Dim SizeLimit As Long
    Dim v, c
    
    SizeLimit = 10200000 '���� ũ�� ���� �뷮
    
    If ready = False Then '#���� �غ� ������ Ȯ��
        
        frmdir.cdfile.DialogTitle = "����"
        frmdir.cdfile.Filter = "��� ����(*.*)|*.*|Mp3 ����(*.mp3)|*.mp3|���̺� ����(*.wav)|*.wav|�ؽ�Ʈ ����(*.txt)|*.txt|�׸� ����(*.jpg)|*.jpg|���� ����(*.zip)|*.zip|��Ʈ�� ����(*.bmp)|�̵� ����(*.mid)|���� ����(*.zip)|"
        frmdir.cdfile.ShowOpen
        
        If frmdir.cdfile.FileName <> "" Then '#���� �̸��� �����ϴ��� Ȯ��

            '==================================================='
            
            If FileLen(frmdir.cdfile.FileName) < SizeLimit Then '#���� ũ�� ���ѿ��� ������� Ȯ��
            
                c = frmdir.cdfile.FileName
                v = c
                
                '(\) ����======================================='
                
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
                
                '==============================================='
                
                frmclient.wsclient.SendData "len/" & FileLen(frmdir.cdfile.FileName) & "/" & c
                
                ready = True
                
                Exit Sub
                
            End If
            
            '==================================================='
            
        ElseIf frmdir.cdfile.FileName <> "" Then
            If FileLen(frmdir.cdfile.FileName) > SizeLimit Then '���� ũ�� ���ѿ��� ���
            
                MsgBox "10MB �̻��� ������ ������ �� �����ϴ�."
                Exit Sub
            End If
        End If
        
    ElseIf ready = True Then '#���� ����
    
        frmclient.wsclient.SendData "sending" '������ "���� ������..." Ȯ��
        SendFile frmdir.cdfile.FileName '���� ���� ����
        frmdir.cdfile.FileName = ""
        ready = False
        
        tbtool.Buttons(2).Enabled = False
        
    End If
    Case 3
        frmPiano.Show
End Select
End Sub

'==============================================================='
'==============================================================='

Private Sub txtchat_KeyPress(KeyAscii As Integer) 'ä��

    If KeyAscii = 13 And txtchat <> "" Then
        frmclient.wsclient.SendData "msg/" & frmclient.txtnick & ">> " & txtchat
        txtchat = ""
    End If
    
End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Resize()

On Error Resume Next

If Not Me.WindowState = 1 Then
    imgtop.Width = Me.ScaleWidth
    lstmsg.Width = Me.ScaleWidth - (lstmember.Width)
    lstmsg.ColumnHeaders(1).Width = lstmsg.Width - 270
    lstmsg.Height = Me.ScaleHeight - (txtchat.Height + tbtool.Height + imgtop.Height)
    lstmember.Height = lstmsg.Height
    tbtool.Top = imgtop.Height + lstmsg.Height
    txtchat.Top = imgtop.Height + lstmsg.Height + tbtool.Height
    txtchat.Width = lstmsg.Width + lstmember.Width
    pb.Width = lstmsg.Width
End If

End Sub

'==============================================================='
'==============================================================='

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmclient.wsclient.Close
frmclient.wsdownload.Close
midiOutClose hMidi
End
End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Terminate()
frmclient.wsclient.Close
frmclient.wsdownload.Close
End
End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Unload(Cancel As Integer)
frmclient.wsclient.Close
frmclient.wsdownload.Close
End
End Sub

'==============================================================='
'==============================================================='

Private Sub wsmusic_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

End Sub
