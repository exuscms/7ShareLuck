VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmserver 
   BorderStyle     =   1  '���� ����
   Caption         =   "7 Share Luck"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   510
   ClientWidth     =   10215
   Icon            =   "frmserver.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   10215
   StartUpPosition =   2  'ȭ�� ���
   Begin MSWinsockLib.Winsock wsupload 
      Index           =   0
      Left            =   240
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock wslogin 
      Index           =   0
      Left            =   240
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton cmdListen 
      Caption         =   "���� ����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   6360
      Width           =   1455
   End
   Begin VB.TextBox txtport 
      Appearance      =   0  '���
      BackColor       =   &H8000000F&
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
      Height          =   165
      Left            =   1080
      TabIndex        =   0
      Text            =   "12345"
      Top             =   6460
      Width           =   495
   End
   Begin MSWinsockLib.Winsock wsinfo 
      Index           =   0
      Left            =   240
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin prjserver.ucXTab ucTab 
      Height          =   5775
      Left            =   120
      TabIndex        =   6
      Top             =   480
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10186
      TabCaption(0)   =   "Tab 0"
      TabCaption(1)   =   "Tab 1"
      TabCaption(2)   =   "Tab 2"
      ActiveTabBackEndColor=   16514555
      ActiveTabBackStartColor=   16514555
      BeginProperty ActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   16514555
      DisabledTabBackColor=   -2147483633
      DisabledTabForeColor=   10526880
      ForeColor       =   -2147483630
      InActiveTabBackEndColor=   15397104
      InActiveTabBackStartColor=   16777215
      BeginProperty InActiveTabFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
         Size            =   9
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OuterBorderColor=   10198161
      PictureMaskColor=   16711935
      TabTheme        =   1
      TabOffset       =   11305
      Begin VB.CommandButton cmdadd 
         Caption         =   "���� �߰�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -13970
         TabIndex        =   17
         Top             =   5160
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Appearance      =   0  '���
         Caption         =   "��� �����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -15170
         TabIndex        =   16
         Top             =   5160
         Width           =   1095
      End
      Begin VB.Frame fmshare 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5175
         Left            =   -22490
         TabIndex        =   14
         Top             =   480
         Width           =   9735
         Begin MSComDlg.CommonDialog cdfile 
            Left            =   0
            Top             =   2400
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
         Begin MSComctlLib.ListView lvshare 
            Height          =   4335
            Left            =   10
            TabIndex        =   15
            Top             =   240
            Width           =   9700
            _ExtentX        =   17092
            _ExtentY        =   7646
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            OLEDragMode     =   1
            OLEDropMode     =   1
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            OLEDragMode     =   1
            OLEDropMode     =   1
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "���� �̸�"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "���"
               Object.Width           =   14552
            EndProperty
         End
      End
      Begin VB.Frame fmchat 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "ä�� �α�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   5160
         Left            =   -11185
         TabIndex        =   11
         Top             =   480
         Width           =   9735
         Begin VB.TextBox txtadmin 
            Appearance      =   0  '���
            BackColor       =   &H8000000E&
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
            TabIndex        =   12
            Top             =   4890
            Width           =   9725
         End
         Begin MSComctlLib.ListView lvchat 
            Height          =   4695
            Left            =   15
            TabIndex        =   13
            Top             =   240
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   8281
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "����"
               Object.Width           =   14588
            EndProperty
         End
      End
      Begin VB.Frame fmlog 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "�α�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2640
         Left            =   120
         TabIndex        =   9
         Top             =   3000
         Width           =   9735
         Begin MSComctlLib.ImageList il 
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
                  Picture         =   "frmserver.frx":08CA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmserver.frx":0C64
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmserver.frx":11FE
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComctlLib.ListView lvlog 
            Height          =   2405
            Left            =   15
            TabIndex        =   10
            Top             =   210
            Width           =   9705
            _ExtentX        =   17119
            _ExtentY        =   4233
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "����"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "����"
               Object.Width           =   14588
            EndProperty
         End
      End
      Begin VB.Frame fmremote 
         Appearance      =   0  '���
         BackColor       =   &H80000005&
         Caption         =   "������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   8.25
            Charset         =   129
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2415
         Left            =   120
         TabIndex        =   7
         Top             =   480
         Width           =   9715
         Begin MSComctlLib.ListView lvremote 
            Height          =   2180
            Left            =   15
            TabIndex        =   8
            Top             =   220
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   3836
            View            =   3
            Arrange         =   1
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            HideColumnHeaders=   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            HotTracking     =   -1  'True
            HoverSelection  =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   8.25
               Charset         =   129
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            NumItems        =   1
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "IP"
               Object.Width           =   2540
            EndProperty
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  '�� ����
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "il"
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
      BorderStyle     =   1
   End
   Begin VB.Label lbips 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ������ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   1680
      TabIndex        =   5
      Top             =   6460
      Width           =   1110
   End
   Begin VB.Label lbip 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   2880
      TabIndex        =   4
      Top             =   6460
      Width           =   60
   End
   Begin VB.Label lbcnt 
      Alignment       =   1  '������ ����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�� ������ : 0"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   7440
      TabIndex        =   3
      Top             =   6480
      Width           =   975
   End
   Begin VB.Label lbport 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���� ��Ʈ :"
      BeginProperty Font 
         Name            =   "����"
         Size            =   8.25
         Charset         =   129
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   120
      TabIndex        =   1
      Top             =   6460
      Width           =   930
   End
   Begin VB.Menu mnuadmin 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuban 
         Caption         =   "����"
      End
   End
End
Attribute VB_Name = "frmserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type INITCOMMONCONTROLSEX_TYPE
dwSize As Long
dwICC As Long
End Type
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (lpInitCtrls As INITCOMMONCONTROLSEX_TYPE) As Long
Private Const ICC_INTERNET_CLASSES = &H800

'==============================================================='
'==============================================================='

Private Sub cmdadd_Click() '���� ���� �߰�(����)
    
    On Error Resume Next
    
    '���� ����
    cdfile.ShowOpen
    
    If cdfile.FileName <> "" Then
    
        '======================================================='
        
        '���� (Limit)
        If FileLen(cdfile.FileName) < SizeLimit Then
        
            '���� ���� �߰�
            Share cdfile.FileTitle, cdfile.FileName
            
            '---------------------------------------------------'
            For i = 1 To Cntclient
            
            '���� Ȯ��
            If wslogin(i).State = sckConnected Then
            
                'Ŭ���̾�Ʈ ���� ���� ���� (���)
                wslogin(i).SendData "setdir/"
            End If
            
            Next i
            '---------------------------------------------------'
            
        End If
        
        '======================================================='
        
    End If

End Sub

'==============================================================='
'==============================================================='

Public Sub cmdListen_Click() '���� ����
    
    On Error Resume Next
    
    Sendfilechk = False
    Me.Caption = "7 Share Luck - " & wslogin(0).LocalIP
    
    'Listen ä�� ����
    wslogin(0).Close
    wslogin(0).LocalPort = txtport
    wslogin(0).Listen
    
    'Listen ���� ��� ����
    wsinfo(0).Close
    wsinfo(0).LocalPort = 12347
    wsinfo(0).Listen
    
    Log "����", "���� ���� �Ϸ�"
    
End Sub

'==============================================================='
'==============================================================='

Private Sub Command1_Click() '���� ���� �ʱ�ȭ
    
    '���� ���� ����
    lvshare.ListItems.Clear

End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Load()

    On Error Resume Next
    Dim comctls As INITCOMMONCONTROLSEX_TYPE ' identifies the control to register
    Dim retval As Long ' generic return value
    With comctls
    .dwSize = Len(comctls)
    .dwICC = ICC_INTERNET_CLASSES
    End With
    
    retval = InitCommonControlsEx(comctls)
    ucTab.TabCaption(0) = "����"
    ucTab.TabCaption(1) = "ä��"
    ucTab.TabCaption(2) = "����"
    SizeLimit = 10100000
    pw = 0
    
    Sendfilechk = False '(���� üũ) Boolean = False
    lbip = wslogin(0).LocalIP
    
    '�ӽ� ������ �����ϴ��� ����
    If Dir(App.Path & "\tmp" & "\tmp.dat") <> "" Then
    
        '�ӽ� ���� ����
        Kill App.Path & "\tmp" & "\tmp.dat"
        
    End If

End Sub

'==============================================================='
'==============================================================='

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer) '��������
    
    wslogin(0).Close

End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Terminate() '��������
    
    wslogin(0).Close

End Sub

'==============================================================='
'==============================================================='

Private Sub Form_Unload(Cancel As Integer) '����
    
    wslogin(0).Close

End Sub

'==============================================================='
'==============================================================='

'���콺 Ŭ�� (���� ���)
Private Sub lvremote_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error GoTo adminnull '>>SelectedItem�� ������ ����
    
    If Button = 2 Then '2=(������)
    
        '======================================================='
        
        If lvremote.SelectedItem.Text <> "" Then
            PopupMenu mnuadmin '�˾�
        End If
        
        '======================================================='
    End If
    
    Exit Sub
    
adminnull:     '���� ����

End Sub

'==============================================================='
'==============================================================='

Private Sub lvshare_DblClick() '����Ŭ��
    
    '�������� ����
    lvshare.ListItems.Remove (lvshare.SelectedItem.Index)

End Sub

'==============================================================='
'==============================================================='

'Drag-Drop ���� �߰�
Private Sub lvshare_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    On Error Resume Next
    
    Dim i As Integer
    Dim dragtext
    Dim W, process1, a, b As String
    
    If Not FileLen(dragtext) < SizeLimit Then '���� ũ�� Ȯ��(Limit)
    
        For Each v In data.Files
        
        '======================================================='
        
        c = dragtext
        
        '======================================================='
        
        Do While Not InStr(process1, "\") = "0" '\ ����
            DoEvents
            If InStr(1, process1, "\") Then
                W = InStr(1, process1, "\")
                If W >= 2 Then
                    process1 = Right(dragtext, Len(process1) - (W - 1))
                ElseIf W = 1 Then
                    process1 = Right(dragtext, Len(process1) - (W))
                End If
            End If
        Loop
        
        '======================================================='
        
        lvshare.ListItems.Add , , process1
        lvshare.ListItems.Item(lvshare.ListItems.Count).ListSubItems.Add , , dragtext
        
        Next
        
        For i = 1 To wslogin.Count - 1
        
        If wslogin(i).State = 7 Then
            wslogin(i).SendData "setdir/"
        End If
        
        Next i
        
    End If

End Sub

'==============================================================='
'==============================================================='

Private Sub mnuban_Click() '��������
    
    On Error Resume Next
    
    Strtext = InputBox("���� ���� ����", "���� ����")
    
    If Strtext <> "" Then '���� ���翩�� Ȯ��
    
        Log "����", wslogin(lvremote.SelectedItem.Index).RemoteHostIP
        wslogin(lvremote.SelectedItem.Index).SendData "ban," & Strtext 'Ŭ���̾�Ʈ ����
        lvremote.ListItems.Remove lvremote.SelectedItem.Index
        
    End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
    Case 1
        cmdListen_Click
End Select
End Sub

'==============================================================='
'==============================================================='

Private Sub txtadmin_KeyPress(KeyAscii As Integer) '������ ä��
    On Error Resume Next
    If KeyAscii = 13 Then
    
        '======================================================='
        If txtadmin <> "" Then
            For i = 1 To wslogin.Count - 1
                If wslogin(i).State = 7 Then
                    Chat "������", txtadmin
                    wslogin(i).SendData "admin/" & txtadmin
                End If
            Next i
            txtadmin = ""
        End If
        '======================================================='
        
    End If
    
End Sub


'==============================================================='
'==============================================================='

Private Sub wsinfo_Close(Index As Integer) '���� ��� Ŭ���̾�Ʈ �ݱ�
    
    On Error Resume Next
    Cntclient2 = Cntclient2 - 1 '���� ��ȸ -1
    Unload wsinfo(Index) 'Ŭ���̾�Ʈ �ݱ�

End Sub

'==============================================================='
'==============================================================='

'���� ��� Ŭ���̾�Ʈ ���� �õ�
Private Sub wsinfo_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    On Error Resume Next
    Cntclient2 = Cntclient2 + 1 '���� ���� +1
    Load wsinfo(Cntclient2) 'Ŭ���̾�Ʈ �ҷ�����
    wsinfo(Cntclient2).Accept requestID '���� ���
    
End Sub

'==============================================================='
'==============================================================='

'���� ��� ������
Private Sub wsinfo_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    
    On Error Resume Next
    
    Dim strdata As String

    wsinfo(Index).GetData strdata
    
    If strdata = "dir" Then '���� ��� �Ѱ��ֱ�
        
        '======================================================='
        
        If lvshare.ListItems.Count > 0 Then
            For i = 1 To lvshare.ListItems.Count
                Log "���� �������", lvshare.ListItems(i).ListSubItems(1).Text & i
                wsinfo(Index).SendData lvshare.ListItems(i).ListSubItems(1).Text & "?"
            Next i
        End If
        
        '======================================================='
        
    End If

End Sub

'==============================================================='
'==============================================================='

Private Sub wslogin_Close(Index As Integer) '�α��� ���� ����
    Dim i As Long
    On Error Resume Next
    Cntclient = Cntclient - 1
    'piclog(Index).BackColor = vbRed
    Log "���� ����", wslogin(Cntclient).RemoteHostIP & "(" & Index & ")" & ",(" & Cntclient & ")"
    lbcnt = "�� ������ : " & Cntclient
    Unload wslogin(Index) '�α��� ���� ����
    Unload wsupload(Index) '���ε� ���� ����
        For i = 0 To lvremote.ListItems.Count
        If lvremote.ListItems(i).ListSubItems(1).Text = Index Then
            lvremote.ListItems.Remove i
        End If
    Next i
End Sub

'==============================================================='
'==============================================================='

'�α��� ���� ����
Private Sub wslogin_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    On Error Resume Next
    Cntclient = Cntclient + 1 '������ ��ȸ +1
    Cntclient3 = Cntclient3 + 1
    lbcnt = "�� ������ : " & Cntclient & "(" & Index & ")"
    'piclog(Index).BackColor = vbBlue
    
    '#�α��� ����
    Load wslogin(Cntclient)
    wslogin(Cntclient).Close
    wslogin(Cntclient).Accept requestID 'ä�� Ŭ���̾�Ʈ
    
    '#���ε� ����
    Load wsupload(Cntclient)
    wsupload(Cntclient).Close
    wsupload(Cntclient).RemoteHost = wslogin(Cntclient).RemoteHostIP '���ε� Ŭ���̾�Ʈ
    wsupload(Cntclient).RemotePort = 12346
    wsupload(Cntclient).Connect
    
End Sub

'==============================================================='
'==============================================================='

'�α��� ���� ������
Private Sub wslogin_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    On Error Resume Next
    Dim i, z As Long
    Dim strdata As String
    Dim strlist As String
    Dim ipsplit() As String
    Dim strsplit() As String
    
    wslogin(Index).GetData strdata
    
    strsplit() = Split(strdata, "/") '/�� �ɰ���
    
    '#�ٿ�� ������ (0)
    If UBound(strsplit) = 0 Then
        
        '======================================================='
        
        '���۰� ���õ� ����� ���� ����(Sends)
        If strsplit(0) = "sending" Then
            For i = 1 To Cntclient3
                If wslogin(i).State = sckConnected Then
                    wslogin(i).SendData "sends" '���� ���� ����
                End If
            Next i
        End If
        
        '======================================================='
        
        If strsplit(0) = "memlst" Then
            If lvremote.ListItems.Count > 0 Then
                For z = 1 To lvremote.ListItems.Count
                    If lvremote.ListItems(z).Text <> "" Then
                        Log "����� �������", lvremote.ListItems(z).Text
                        wsinfo(Index).SendData lvremote.ListItems(z).Text & "<"
                    End If
                Next z
            End If
        End If
        
    End If
    
    '#�ٿ�� ������ (1)
    If UBound(strsplit) = 1 Then
    
        '======================================================='
        
        '����� ������ �� ���� & ���ۿ� ���þ��� ����� ���� ����
        'Logcnt / Senddingg
        If strsplit(0) = "request" Then
            Log "���� ��û", wslogin(Index).RemoteHostIP & "(" & Index & ")" & ",(" & Cntclient & ") - " & strsplit(1)
            
            If lvremote.ListItems.Count > 0 Then
                
                For z = 1 To lvremote.ListItems.Count
                    If strsplit(1) = lvremote.ListItems(z).Text Then
                        If wslogin(Index).State = sckConnected Then
                            wslogin(Index).SendData "ban/�ߺ� �г���"
                        End If
                        Exit Sub
                    End If
                Next z
            
            End If
            
            Remote strsplit(1), Index
            
            '---------------------------------------------------'
            
            For i = 1 To wslogin.Count - 1
            
                If wslogin(i).State = sckConnected Then
                
                    '==========================================='
                    wslogin(i).SendData "logcnt/" & Val(Cntclient)
                    
                    If Sendfilechk = True Then
                        wsinfo(i).SendData "senddingg" '���þ��� ����� ���� ����
                    End If
                    
                    '==========================================='
                    
                End If
                
            Next i
            '---------------------------------------------------'
            
        End If
        
        '======================================================='
        '======================================================='
        
        '���� ���� �� �غ�(Len)
        If strsplit(0) = "download" Then
        
            '==================================================='
            
            If wslogin(Index).State = sckConnected Then
                If Sendfilechk = False Then
                    wslogin(Index).SendData "len/" & FileLen(strsplit(1))
                End If
            End If
            
            '==================================================='
            
        End If
        
        '======================================================='
        '======================================================='
        
        '���� ���� ����
        If strsplit(0) = "start" Then
            Log "��������", i & "/" & strsplit(1)
            SendFile strsplit(1), Index
        End If
        
        '======================================================='
        '======================================================='
        
        '�� Ŭ���̾�Ʈ �޽��� ����(Msg)
        
        If strsplit(0) = "msg" Then
        
            Chat wslogin(Index).RemoteHostIP, Mid(strdata, 5, Len(strdata) - 2)
            
            '---------------------------------------------------'
            
            '������ ����
            For i = 1 To Cntclient3
                If wslogin(i).State = sckConnected Then
                    wslogin(i).SendData "msg/" & Mid(strdata, 5, Len(strdata) - 2)
                End If
            Next i
            
            '---------------------------------------------------'
            
        End If
        
        '======================================================='
    
    End If
    
    '�ٿ�� ������ (2)
    If UBound(strsplit) = 2 Then
    
        '==================================================='
        
        '���� �뷮 ����(View)
        
        '#Len -> View
        If strsplit(0) = "len" Then
            File_Size = Val(strsplit(1))
            For i = 1 To Cntclient3
                If wsinfo(i).State = sckConnected Then
                    Sendfilechk = True
                    pw = wsupload.Count - 1
                    Log "���� ���� �غ�", strsplit(2)
                    wsinfo(i).SendData "view/" & File_Size & "/" & strsplit(2)
                End If
            Next i
        End If
        
        '==================================================='
        
        If strsplit(0) = "pplay" Then
            For i = 1 To Cntclient3
                If wslogin(i).State = sckConnected Then
                    wslogin(i).SendData "pplay/" & strsplit(1) & "/" & strsplit(2)
                End If
            Next i
        End If
        
        If strsplit(0) = "splay" Then
            For i = 1 To Cntclient3
                If wslogin(i).State = sckConnected Then
                    wslogin(i).SendData "splay/" & strsplit(1) & "/" & strsplit(2)
                End If
            Next i
        End If
    End If

End Sub

'==============================================================='
'==============================================================='

'(���� >> ����) �ٿ�ε�
Private Sub wsupload_DataArrival(Index As Integer, ByVal bytesTotal As Long)

    On Error Resume Next
    Dim d() As Byte
    Dim k
    Dim i As Long
    Dim Sendcount As Integer
    
    wsupload(Index).GetData d '���� ������ ���
    
    '��ΰ� ������� ���� �����
    If Not Dir(App.Path & "\tmp", vbDirectory) <> "" Then
        MkDir App.Path & "\tmp"
    End If
    
    '���� �ۼ�
    Open App.Path & "\tmp" & "\tmp.dat" For Binary Access Write As #1
        Put #1, LOF(1) + 1, d
        k = LOF(1)
    Close #1
    
    '�ٿ�ε� �Ϸ�
    If k >= File_Size Then
    
        '��� Ŭ���̾�Ʈ�� ���� ����
        For i = 1 To Cntclient3

        If wslogin(i).State = sckConnected Then
            Log "�ӽ� ���� ����", i
            SendFile App.Path & "\tmp" & "\tmp.dat", i
        ElseIf wslogin(i).State = sckClosed Then
            Log "���� �����ڰ� ��������", i
        End If
        Next i

        Sendfilechk = False
        
        '������ �ӽ����� ����
        If Dir(App.Path & "\tmp" & "\tmp.dat") <> "" Then
            Kill App.Path & "\tmp" & "\tmp.dat"
        End If
        
        pw = 0
        
    End If

End Sub
