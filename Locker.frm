VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Ether 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Hroton OpenEther"
   ClientHeight    =   9405
   ClientLeft      =   120
   ClientTop       =   1065
   ClientWidth     =   18840
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "΢���ź�"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Locker.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9405
   ScaleWidth      =   18840
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text2 
      ForeColor       =   &H00808080&
      Height          =   375
      Left            =   13680
      TabIndex        =   8
      Text            =   "Microsoft Bing ����"
      Top             =   240
      Width           =   3495
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   1455
      Left            =   0
      TabIndex        =   7
      Top             =   840
      Width           =   5535
      ExtentX         =   9763
      ExtentY         =   2566
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.CommandButton Commandfuck 
      Caption         =   "Command3"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1320
      Width           =   495
   End
   Begin VB.CommandButton REFRESH 
      Caption         =   "ˢ��"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton FORWARD 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton BACK 
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Height          =   420
      Left            =   3120
      TabIndex        =   2
      Top             =   240
      Width           =   10455
   End
   Begin VB.CommandButton HOME 
      Caption         =   "��ҳ"
      Height          =   375
      Left            =   17400
      TabIndex        =   1
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      Caption         =   "sdf"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   975
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000000&
      X1              =   0
      X2              =   5520
      Y1              =   810
      Y2              =   810
   End
   Begin VB.Menu Web 
      Caption         =   "��ҳ"
      Begin VB.Menu NEW 
         Caption         =   "��"
      End
      Begin VB.Menu Do 
         Caption         =   "����"
         Begin VB.Menu Bac 
            Caption         =   "����"
         End
         Begin VB.Menu FORWAR 
            Caption         =   "ǰ��"
         End
         Begin VB.Menu FRESH 
            Caption         =   "ˢ��"
         End
         Begin VB.Menu STOP 
            Caption         =   "ֹͣ"
         End
         Begin VB.Menu GoHome 
            Caption         =   "��ҳ"
         End
      End
      Begin VB.Menu Save 
         Caption         =   "���Ϊ"
      End
      Begin VB.Menu Print 
         Caption         =   "��ӡ"
      End
      Begin VB.Menu T 
         Caption         =   "����"
      End
      Begin VB.Menu M 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "�˳�"
      End
   End
   Begin VB.Menu View 
      Caption         =   "�鿴"
      Begin VB.Menu big150 
         Caption         =   "����150%"
      End
      Begin VB.Menu Big125 
         Caption         =   "����125%"
      End
      Begin VB.Menu Big100 
         Caption         =   "����100%"
      End
      Begin VB.Menu little75 
         Caption         =   "����75%"
      End
      Begin VB.Menu Little50 
         Caption         =   "����50%"
      End
      Begin VB.Menu Little25 
         Caption         =   "����25%"
      End
   End
   Begin VB.Menu About 
      Caption         =   "����"
   End
End
Attribute VB_Name = "Ether"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Bac_Click() '�˵�����
On Error Resume Next
WebBrowser1.GoBack
End Sub
Private Sub BACK_Click() '����
    On Error Resume Next
    WebBrowser1.GoBack
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub Big100_Click() '����
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "100%"
End Sub
Private Sub Big125_Click() '����
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "125%"
End Sub
Private Sub Big150_Click() '����
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "150%"
End Sub
Private Sub Command1_Click()
    On Error Resume Next
    WebBrowser1.Navigate Trim(Text1.Text) '����ҳ
If Text1.Text = "about:" Then
Form1.Show '��ʾ���ڴ���
End If
End Sub
Private Sub About_Click() '����
Form1.Show
End Sub
Private Sub FORWAR_Click() 'ǰ��
On Error Resume Next
WebBrowser1.GoForward
End Sub
Private Sub FRESH_Click() 'ˢ��
On Error Resume Next
 WebBrowser1.REFRESH
End Sub
Private Sub GoHome_Click() '��ҳ
On Error Resume Next
WebBrowser1.GoHome
End Sub
Private Sub EXIT_Click() '�Ƴ�
End
End Sub
Private Sub Little25_Click() '����
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "25%"
End Sub
Private Sub Little50_Click() '����
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "50%"
End Sub
Private Sub little75_Click() '����
On Error Resume Next
WebBrowser1.Document.body.Style.Zoom = "75%"
End Sub
Private Sub New_Click() '���ļ�
WebBrowser1.ExecWB OLECMDID_OPEN, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub Print_Click() '��ӡ
WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub Save_Click() '������ҳ
WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub T_Click() '����
WebBrowser1.ExecWB OLECMDID_FIND, OLECMDEXECOPT_DODEFAULT
End Sub
Private Sub Form_Load()
Call Command1_Click
  HomeAddress = " https://cn.bing.com" '��д��ҳ��ַ
    WebBrowser1.Navigate HomeAddress
Text1.Width = 11175
Set W = CreateObject("wscript.shell")
W.regwrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", "11000", "REG_DWORD"
Set W = Nothing
End Sub
Private Sub FORWARD_Click() 'ǰ��
On Error Resume Next
    WebBrowser1.GoForward
    GoAdDress.Text = WebBrowser1.LocationURL
End Sub
Private Sub Home_Click()
On Error Resume Next
 WebBrowser1.GoHome '������ҳ
End Sub
Private Sub REFRESH_Click() 'ˢ��
On Error Resume Next
    WebBrowser1.REFRESH
End Sub
Private Sub STOP_Click() 'ֹͣ
On Error Resume Next
WebBrowser1.STOP
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    WebBrowser1.Height = Ether.Height - 700
    WebBrowser1.Width = Ether.Width
    Text1.Width = Me.Width - 8800
    Text2.Left = Text1.Left + Text1.Width + 100
    HOME.Left = Text2.Left + Text2.Width + 100
    Line1.X2 = Me.Width
    Text1.Height = Text2.Height
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer) '�س���
    On Error Resume Next
    If KeyAscii = vbKeyReturn Then
        Command1_Click
    End If
End Sub
Private Sub Text2_KeyPress(KeyAscii As Integer)
'������»س������������ҳ
If KeyAscii = vbKeyReturn Then
    WebBrowser1.Navigate "https://cn.bing.com/search?q=" + (UTF8EncodeURI(Text2.Text))
End If
End Sub
Private Sub Text2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
'�������������������������ʧ
If Text2.Text = Text2.Text Then
Text2.ForeColor = &H0&
Text2.Text = ""
End If
End Sub
Function UTF8EncodeURI(szInput) 'תUTF8������
Dim wch, uch, szRet
Dim x
Dim nAsc, nAsc2, nAsc3
If szInput = "" Then
UTF8EncodeURI = szInput
Exit Function
End If
For x = 1 To Len(szInput)
wch = Mid(szInput, x, 1)
nAsc = AscW(wch)
If nAsc < 0 Then nAsc = nAsc + 65536
If (nAsc And &HFF80) = 0 Then
szRet = szRet & wch
Else
If (nAsc And &HF000) = 0 Then
uch = "%" & Hex(((nAsc \ 2 ^ 6)) Or &HC0) & Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
Else
uch = "%" & Hex((nAsc \ 2 ^ 12) Or &HE0) & "%" & _
Hex((nAsc \ 2 ^ 6) And &H3F Or &H80) & "%" & _
Hex(nAsc And &H3F Or &H80)
szRet = szRet & uch
End If
End If
Next
UTF8EncodeURI = szRet
End Function
Function GBKEncodeURI(szInput)          '����GBK15����
Dim i As Long
Dim x() As Byte
Dim szRet As String
szRet = ""
x = StrConv(szInput, vbFromUnicode)
For i = LBound(x) To UBound(x)
szRet = szRet & "%" & Hex(x(i))
Next
GBKEncodeURI = szRet
End Function
Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
Text1.Text = WebBrowser1.LocationURL '��������ַ�ı�ʱ���ı��������ݸı�
End Sub
