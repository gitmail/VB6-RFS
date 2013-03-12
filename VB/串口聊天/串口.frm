VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   7935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7935
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.VScrollBar VScroll1 
      Height          =   3615
      Left            =   11280
      TabIndex        =   18
      Top             =   4080
      Width           =   375
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   840
      Top             =   7200
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   3240
      MultiLine       =   -1  'True
      TabIndex        =   14
      Top             =   4080
      Width           =   8415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "发送"
      Height          =   615
      Left            =   10080
      TabIndex        =   13
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   3255
      Left            =   3240
      TabIndex        =   12
      Top             =   480
      Width           =   8175
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "串口.frx":0000
      Left            =   960
      List            =   "串口.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "串口设置"
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   2655
      Begin VB.ComboBox Combo5 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2280
         Width           =   1215
      End
      Begin VB.ComboBox Combo4 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.ComboBox Combo3 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1320
         Width           =   1215
      End
      Begin VB.ComboBox Combo2 
         Height          =   300
         Left            =   720
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "打开串口"
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H000000FF&
         BorderColor     =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   375
         Left            =   2040
         Shape           =   3  'Circle
         Top             =   3000
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "校验位"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "停止位"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "数据位"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "波特率"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "端口号"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   240
      Top             =   7080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   16
      DTREnable       =   -1  'True
      InBufferSize    =   10240
      OutBufferSize   =   5120
   End
   Begin VB.Label Label10 
      Caption         =   "已接收"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   20
      Top             =   5400
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "0"
      Height          =   255
      Left            =   1200
      TabIndex        =   19
      Top             =   5400
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "接收框"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   17
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "接收框"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   15
      Left            =   960
      TabIndex        =   16
      Top             =   5280
      Width           =   1455
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      Height          =   135
      Left            =   3240
      TabIndex        =   15
      Top             =   4080
      Width           =   15
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim s_tem$
If Combo5.ListIndex = 0 Then
    s_tem$ = "E"
ElseIf Combo5.ListIndex = 1 Then
    s_tem$ = "M"
ElseIf Combo5.ListIndex = 2 Then
    s_tem$ = "N"
ElseIf Combo5.ListIndex = 3 Then
    s_tem$ = "O"
Else: s_tem$ = "S"
End If
MSComm1.Settings = Combo2.Text & "," & s_tem$ & "," & Combo3.Text & "," & Combo4.Text
If (MSComm1.PortOpen = False) Then
    MSComm1.CommPort = Combo1.ListIndex + 1
    MSComm1.PortOpen = True
    Shape1.FillColor = &HFF000
    Command1.Caption = "关闭串口(&S)"
    'MsgBox ("com=" & MSComm1.CommPort & "setting=" & MSComm1.Settings)
ElseIf MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
    Shape1.FillColor = &H0
    Command1.Caption = "打开串口(&S)"
 'MsgBox ("com=" & MSComm1.CommPort & "setting=" & MSComm1.Settings)
End If

End Sub


Private Sub Command2_Click()
SendMsg (Text1.Text)
End Sub

Private Sub Form_Load()
For i = 0 To 19 Step 1
Combo1.AddItem "COM" & (i + 1)
Next i
Combo1.ListIndex = 0
'波特率
Combo2.List(0) = 2400
Combo2.List(1) = 4800
Combo2.List(2) = 9600
Combo2.List(3) = 14400
Combo2.List(4) = 19200
Combo2.List(5) = 38400
Combo2.List(6) = 19200
Combo2.List(7) = 57600
Combo2.List(8) = 115200
Combo2.ListIndex = 8

'数据位
Combo3.List(0) = 5
Combo3.List(1) = 6
Combo3.List(2) = 7
Combo3.List(3) = 8
Combo3.ListIndex = 3
'停止位
Combo4.List(0) = 1
Combo4.List(1) = 2
Combo4.ListIndex = 0
'校验位
Combo5.AddItem "Even"
Combo5.AddItem "Mark"
Combo5.AddItem "Default"
Combo5.AddItem "Odd"
Combo5.AddItem "Space"
Combo5.ListIndex = 2
'按Command2键的快捷键设定
Command1.Caption = "打开串口(&S)"
Command2.Caption = "发送(&A)"
'链接本地数据库
'On Error GoTo Err100
'Err100:
'MsgBox ("oh fuck")
End Sub

'按F5键发送数据
Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 116 Then
Command2_Click
End If
End Sub
'*************串口接收子程序****************
Public Sub ReciveMsg()
Dim StrRec$
If MSComm1.PortOpen = True Then
StrRec$ = MSComm1.Input

 If Left(StrRec$, 8) = "shutdown" Then
      Shell "E:\program files\Warcraft III\Frozen Throne.exe"
 End If
 SendMsg (StrRec$)
Text2.Text = Text2.Text + StrRec$
End If
End Sub
'*************定时刷新接收****************
Private Sub Timer1_Timer()
ReciveMsg
Label9.Caption = Len(Text2.Text)
Text2.SelStart = Len(Text2.Text)
End Sub
'*************串口发送子程序****************
Private Sub SendMsg(send$)
If send$ <> "" And MSComm1.PortOpen = True Then
MSComm1.Output = send$
End If

End Sub

