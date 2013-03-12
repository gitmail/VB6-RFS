VERSION 5.00
Begin VB.Form frmmain 
   Caption         =   "TAB1"
   ClientHeight    =   6390
   ClientLeft      =   1110
   ClientTop       =   450
   ClientWidth     =   8355
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8355
   Visible         =   0   'False
   Begin VB.TextBox txtFields 
      DataField       =   "ID"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   0
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   18
      Top             =   1005
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "NAME"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   1
      Left            =   2280
      MaxLength       =   8
      TabIndex        =   17
      Top             =   1320
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Index"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   2
      Left            =   2280
      TabIndex        =   16
      Top             =   1635
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Date"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   3
      Left            =   2280
      TabIndex        =   15
      Top             =   1965
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Time"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   4
      Left            =   2280
      TabIndex        =   14
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Temperature"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   5
      Left            =   2280
      TabIndex        =   13
      Top             =   2595
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "WindSpeed"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   6
      Left            =   2280
      TabIndex        =   12
      Top             =   2925
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "WCI"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   7
      Left            =   2280
      TabIndex        =   11
      Top             =   3240
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "ECT"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   8
      Left            =   2280
      TabIndex        =   10
      Top             =   3555
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "Teq"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   9
      Left            =   2280
      TabIndex        =   9
      Top             =   3885
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "WeiHai"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   10
      Left            =   2280
      TabIndex        =   8
      Top             =   4200
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "LowLabor"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   11
      Left            =   2280
      TabIndex        =   7
      Top             =   4515
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "MidLabor"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   12
      Left            =   2280
      TabIndex        =   6
      Top             =   4845
      Width           =   1935
   End
   Begin VB.TextBox txtFields 
      DataField       =   "HighLabor"
      DataSource      =   "Data1"
      Height          =   285
      Index           =   13
      Left            =   2280
      TabIndex        =   5
      Top             =   5160
      Width           =   1935
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "添加(&A)"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   5505
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "删除(&D)"
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   5505
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "刷新(&R)"
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   5505
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "保存更新(&U)"
      Height          =   1620
      Left            =   4680
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "关闭返回(&C)"
      Height          =   1620
      Left            =   6480
      TabIndex        =   0
      Top             =   4185
      Width           =   1575
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Administrator\桌面\New_sd\上位机\db1.mdb"
      DefaultCursorType=   0  '缺省游标
      DefaultType     =   2  '使用 ODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "TAB1"
      Top             =   6045
      Width           =   8355
   End
   Begin VB.Label lblLabels 
      Caption         =   "如： 2012/8/1"
      Height          =   255
      Index           =   18
      Left            =   4440
      TabIndex        =   35
      Top             =   2080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "如： 19:30"
      Height          =   255
      Index           =   17
      Left            =   4440
      TabIndex        =   34
      Top             =   2350
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID(必填):"
      Height          =   255
      Index           =   0
      Left            =   360
      TabIndex        =   33
      Top             =   1020
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "NAME:"
      Height          =   255
      Index           =   1
      Left            =   360
      TabIndex        =   32
      Top             =   1335
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Index:"
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   31
      Top             =   1665
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Date:"
      Height          =   255
      Index           =   3
      Left            =   360
      TabIndex        =   30
      Top             =   1980
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Time:"
      Height          =   255
      Index           =   4
      Left            =   360
      TabIndex        =   29
      Top             =   2295
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Temperature:"
      Height          =   255
      Index           =   5
      Left            =   360
      TabIndex        =   28
      Top             =   2625
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "WindSpeed:"
      Height          =   255
      Index           =   6
      Left            =   360
      TabIndex        =   27
      Top             =   2940
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "WCI:"
      Height          =   255
      Index           =   7
      Left            =   360
      TabIndex        =   26
      Top             =   3255
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ECT:"
      Height          =   255
      Index           =   8
      Left            =   360
      TabIndex        =   25
      Top             =   3585
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "Teq:"
      Height          =   255
      Index           =   9
      Left            =   360
      TabIndex        =   24
      Top             =   3900
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "WeiHai:"
      Height          =   255
      Index           =   10
      Left            =   360
      TabIndex        =   23
      Top             =   4215
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "LowLabor:"
      Height          =   255
      Index           =   11
      Left            =   360
      TabIndex        =   22
      Top             =   4545
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "MidLabor:"
      Height          =   255
      Index           =   12
      Left            =   360
      TabIndex        =   21
      Top             =   4860
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "HighLabor:"
      Height          =   255
      Index           =   13
      Left            =   360
      TabIndex        =   20
      Top             =   5175
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "手动添加数据"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   19
      Top             =   240
      Width           =   3135
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
  Data1.Recordset.AddNew
End Sub

Private Sub cmdDelete_Click()
  '如果删除记录集的最后一条记录
  '记录或记录集中唯一的记录
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  '这仅对多用户应用程序才是需要的
  Data1.Refresh
End Sub

Private Sub cmdUpdate_Click()
  Data1.UpdateRecord
  Data1.Recordset.Bookmark = Data1.Recordset.LastModified
End Sub

Private Sub cmdClose_Click()
frmmain.Visible = False
main.Visible = True
 
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  '这就是放置错误处理代码的地方
  '如果想忽略错误，注释掉下一行代码
  '如果想捕捉错误，在这里添加错误处理代码
  MsgBox "数据错误事件命中错误：" & Error$(DataErr)
  Response = 0  '忽略错误
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  '这将显示当前记录位置
  '为动态集和快照
  Data1.Caption = "记录：" & (Data1.Recordset.AbsolutePosition + 1)
  '对于 Table 对象，当记录集创建后并使用下面的行时，
  '必须设置 Index 属性
  'Data1.Caption = "记录：" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  '这是放置验证代码的地方
  '当下面的动作发生时，调用这个事件
  Select Case Action
    Case vbDataActionMoveFirst
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
  Screen.MousePointer = vbHourglass
End Sub

Private Sub Form_Load()

End Sub
