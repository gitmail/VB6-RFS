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
      Caption         =   "���(&A)"
      Height          =   300
      Left            =   360
      TabIndex        =   4
      Top             =   5505
      Width           =   975
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   300
      Left            =   1440
      TabIndex        =   3
      Top             =   5505
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "ˢ��(&R)"
      Height          =   300
      Left            =   2520
      TabIndex        =   2
      Top             =   5505
      Width           =   975
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "�������(&U)"
      Height          =   1620
      Left            =   4680
      TabIndex        =   1
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "�رշ���(&C)"
      Height          =   1620
      Left            =   6480
      TabIndex        =   0
      Top             =   4185
      Width           =   1575
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\Documents and Settings\Administrator\����\New_sd\��λ��\db1.mdb"
      DefaultCursorType=   0  'ȱʡ�α�
      DefaultType     =   2  'ʹ�� ODBC
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
      Caption         =   "�磺 2012/8/1"
      Height          =   255
      Index           =   18
      Left            =   4440
      TabIndex        =   35
      Top             =   2080
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "�磺 19:30"
      Height          =   255
      Index           =   17
      Left            =   4440
      TabIndex        =   34
      Top             =   2350
      Width           =   1815
   End
   Begin VB.Label lblLabels 
      Caption         =   "ID(����):"
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
      Caption         =   "�ֶ��������"
      BeginProperty Font 
         Name            =   "����"
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
  '���ɾ����¼�������һ����¼
  '��¼���¼����Ψһ�ļ�¼
  Data1.Recordset.Delete
  Data1.Recordset.MoveNext
End Sub

Private Sub cmdRefresh_Click()
  '����Զ��û�Ӧ�ó��������Ҫ��
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
  '����Ƿ��ô��������ĵط�
  '�������Դ���ע�͵���һ�д���
  '����벶׽������������Ӵ��������
  MsgBox "���ݴ����¼����д���" & Error$(DataErr)
  Response = 0  '���Դ���
End Sub

Private Sub Data1_Reposition()
  Screen.MousePointer = vbDefault
  On Error Resume Next
  '�⽫��ʾ��ǰ��¼λ��
  'Ϊ��̬���Ϳ���
  Data1.Caption = "��¼��" & (Data1.Recordset.AbsolutePosition + 1)
  '���� Table ���󣬵���¼��������ʹ���������ʱ��
  '�������� Index ����
  'Data1.Caption = "��¼��" & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  '���Ƿ�����֤����ĵط�
  '������Ķ�������ʱ����������¼�
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
