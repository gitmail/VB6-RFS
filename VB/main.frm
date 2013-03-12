VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form main 
   Appearance      =   0  'Flat
   Caption         =   "冷强度检测仪"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   456
   ClientWidth     =   14268
   ForeColor       =   &H80000002&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9420
   ScaleWidth      =   14268
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer2 
      Interval        =   10
      Left            =   12960
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   13320
      Top             =   360
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   13680
      Top             =   840
      _ExtentX        =   995
      _ExtentY        =   995
      _Version        =   393216
      CommPort        =   16
      DTREnable       =   -1  'True
      InBufferSize    =   10240
      OutBufferSize   =   5120
      RThreshold      =   1
      InputMode       =   1
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   8412
      Left            =   120
      TabIndex        =   16
      Top             =   1080
      Width           =   13932
      _ExtentX        =   24575
      _ExtentY        =   14838
      _Version        =   393216
      TabHeight       =   520
      ForeColor       =   -2147483641
      TabCaption(0)   =   "通讯设置(F2)"
      TabPicture(0)   =   "main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label21"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "TextSend"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Command4"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Frame2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame3"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Frame4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).ControlCount=   7
      TabCaption(1)   =   " 实时显示（&F3)"
      TabPicture(1)   =   "main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Command11"
      Tab(1).Control(1)=   "Command10"
      Tab(1).Control(2)=   "Channel(5)"
      Tab(1).Control(3)=   "Channel(4)"
      Tab(1).Control(4)=   "Channel(3)"
      Tab(1).Control(5)=   "Channel(2)"
      Tab(1).Control(6)=   "Channel(0)"
      Tab(1).Control(7)=   "Channel(1)"
      Tab(1).Control(8)=   "LabelUnconnect"
      Tab(1).Control(9)=   "Line1"
      Tab(1).Control(10)=   "Line7"
      Tab(1).Control(11)=   "Label4"
      Tab(1).Control(12)=   "Label13"
      Tab(1).Control(13)=   "Label12"
      Tab(1).Control(14)=   "Label11"
      Tab(1).Control(15)=   "Label10"
      Tab(1).Control(16)=   "Label9"
      Tab(1).Control(17)=   "Label8"
      Tab(1).Control(18)=   "Label7"
      Tab(1).Control(19)=   "Label6"
      Tab(1).Control(20)=   "Label5"
      Tab(1).Control(21)=   "Label3"
      Tab(1).Control(22)=   "Label14"
      Tab(1).Control(23)=   "Line2"
      Tab(1).Control(24)=   "Line3"
      Tab(1).Control(25)=   "Line4"
      Tab(1).ControlCount=   26
      TabCaption(2)   =   "数据查询(F4)"
      TabPicture(2)   =   "main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command12"
      Tab(2).Control(1)=   "MSFlexGrid1"
      Tab(2).Control(2)=   "Command3"
      Tab(2).ControlCount=   3
      Begin VB.CommandButton Command12 
         Caption         =   "显示全部"
         Height          =   375
         Left            =   -74640
         TabIndex        =   145
         Top             =   480
         Width           =   1095
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         DragMode        =   1  'Automatic
         Height          =   6855
         Left            =   -74880
         TabIndex        =   144
         Top             =   960
         Width           =   13695
         _ExtentX        =   24151
         _ExtentY        =   12086
         _Version        =   393216
         Rows            =   10
         Cols            =   11
         FixedCols       =   0
         ScrollBars      =   0
         AllowUserResizing=   3
      End
      Begin VB.CommandButton Command11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000009&
         Caption         =   "开启自动保存"
         CausesValidation=   0   'False
         Height          =   495
         Left            =   -62640
         TabIndex        =   14
         Top             =   360
         Width           =   1455
      End
      Begin VB.CommandButton Command10 
         Caption         =   "保存结果"
         Height          =   495
         Left            =   -63840
         TabIndex        =   13
         Top             =   360
         Width           =   975
      End
      Begin VB.Frame Channel 
         Caption         =   "Frame5"
         Height          =   7452
         Index           =   5
         Left            =   -63800
         TabIndex        =   84
         Top             =   840
         Visible         =   0   'False
         Width           =   1800
         Begin VB.TextBox TextFLow 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   5
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   143
            Text            =   "main.frx":0054
            Top             =   3480
            Width           =   1600
         End
         Begin VB.TextBox TextFMid 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   5
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   142
            Text            =   "main.frx":0058
            Top             =   4800
            Width           =   1600
         End
         Begin VB.TextBox TextFHigh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   5
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   141
            Text            =   "main.frx":005A
            Top             =   6120
            Width           =   1600
         End
         Begin VB.Label Labelweihai 
            Caption         =   "较小"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   153
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label LabelFWSC 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   360
            TabIndex        =   147
            Top             =   2040
            Width           =   972
         End
         Begin VB.Label LabelFTeq 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   126
            Top             =   2760
            Width           =   852
         End
         Begin VB.Label LabelFECT 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   125
            Top             =   2400
            Width           =   1092
         End
         Begin VB.Label LabelFWindSpeed 
            Caption         =   "2.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   360
            TabIndex        =   124
            Top             =   1680
            Width           =   492
         End
         Begin VB.Label LabelFTemp 
            Caption         =   "-20.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   123
            Top             =   1320
            Width           =   1092
         End
         Begin VB.Label LabelFTime 
            Caption         =   "04:08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   240
            TabIndex        =   122
            Top             =   840
            Width           =   1212
         End
         Begin VB.Label LabelDate 
            Caption         =   "20120909"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   121
            Top             =   480
            Width           =   1092
         End
      End
      Begin VB.Frame Channel 
         Caption         =   "Frame5"
         Height          =   7452
         Index           =   4
         Left            =   -65600
         TabIndex        =   82
         Top             =   840
         Visible         =   0   'False
         Width           =   1800
         Begin VB.TextBox TextFLow 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   4
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   140
            Text            =   "main.frx":005C
            Top             =   3480
            Width           =   1600
         End
         Begin VB.TextBox TextFMid 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   4
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   139
            Text            =   "main.frx":0060
            Top             =   4800
            Width           =   1600
         End
         Begin VB.TextBox TextFHigh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   4
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   138
            Text            =   "main.frx":0062
            Top             =   6120
            Width           =   1600
         End
         Begin VB.Label Labelweihai 
            Caption         =   "较小"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   152
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label LabelFTeq 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   120
            Top             =   2640
            Width           =   852
         End
         Begin VB.Label LabelFECT 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   119
            Top             =   2280
            Width           =   1092
         End
         Begin VB.Label LabelFWSC 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   118
            Top             =   1920
            Width           =   972
         End
         Begin VB.Label LabelFWindSpeed 
            Caption         =   "2.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   117
            Top             =   1560
            Width           =   492
         End
         Begin VB.Label LabelFTemp 
            Caption         =   "-20.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   240
            TabIndex        =   116
            Top             =   1200
            Width           =   1092
         End
         Begin VB.Label LabelFTime 
            Caption         =   "04:08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   480
            TabIndex        =   115
            Top             =   840
            Width           =   1212
         End
         Begin VB.Label LabelDate 
            Caption         =   "20120909"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   4
            Left            =   360
            TabIndex        =   114
            Top             =   480
            Width           =   1092
         End
      End
      Begin VB.Frame Channel 
         Caption         =   "Frame5"
         Height          =   7452
         Index           =   3
         Left            =   -67400
         TabIndex        =   81
         Top             =   840
         Visible         =   0   'False
         Width           =   1800
         Begin VB.TextBox TextFLow 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   3
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   137
            Text            =   "main.frx":0064
            Top             =   3480
            Width           =   1600
         End
         Begin VB.TextBox TextFMid 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   3
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   136
            Text            =   "main.frx":0068
            Top             =   4800
            Width           =   1600
         End
         Begin VB.TextBox TextFHigh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   3
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   135
            Text            =   "main.frx":006A
            Top             =   6120
            Width           =   1600
         End
         Begin VB.Label Labelweihai 
            Caption         =   "较小"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   151
            Top             =   3240
            Width           =   1452
         End
         Begin VB.Label LabelFTeq 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   113
            Top             =   2760
            Width           =   852
         End
         Begin VB.Label LabelFECT 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   112
            Top             =   2400
            Width           =   1092
         End
         Begin VB.Label LabelFWSC 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   360
            TabIndex        =   111
            Top             =   1920
            Width           =   972
         End
         Begin VB.Label LabelFWindSpeed 
            Caption         =   "2.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   360
            TabIndex        =   110
            Top             =   1560
            Width           =   492
         End
         Begin VB.Label LabelFTemp 
            Caption         =   "-20.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   109
            Top             =   1200
            Visible         =   0   'False
            Width           =   1092
         End
         Begin VB.Label LabelFTime 
            Caption         =   "04:08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   240
            TabIndex        =   108
            Top             =   840
            Width           =   1332
         End
         Begin VB.Label LabelDate 
            Caption         =   "20120909"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   3
            Left            =   120
            TabIndex        =   107
            Top             =   480
            Width           =   1092
         End
      End
      Begin VB.Frame Channel 
         Caption         =   "Frame5"
         Height          =   7452
         Index           =   2
         Left            =   -69200
         TabIndex        =   80
         Top             =   840
         Visible         =   0   'False
         Width           =   1800
         Begin VB.TextBox TextFLow 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   134
            Text            =   "main.frx":006C
            Top             =   3480
            Width           =   1600
         End
         Begin VB.TextBox TextFMid 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   133
            Text            =   "main.frx":0070
            Top             =   4800
            Width           =   1600
         End
         Begin VB.TextBox TextFHigh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   2
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   132
            Text            =   "main.frx":0072
            Top             =   6120
            Width           =   1600
         End
         Begin VB.Label Labelweihai 
            Caption         =   "较小"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   120
            TabIndex        =   150
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label LabelFTeq 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   106
            Top             =   2760
            Width           =   852
         End
         Begin VB.Label LabelFECT 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   105
            Top             =   2400
            Width           =   1092
         End
         Begin VB.Label LabelFWSC 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   104
            Top             =   1920
            Width           =   972
         End
         Begin VB.Label LabelFWindSpeed 
            Caption         =   "2.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   103
            Top             =   1560
            Width           =   492
         End
         Begin VB.Label LabelFTemp 
            Caption         =   "-20.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   102
            Top             =   1200
            Width           =   1092
         End
         Begin VB.Label LabelFTime 
            Caption         =   "04:08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   101
            Top             =   840
            Width           =   1452
         End
         Begin VB.Label LabelDate 
            Caption         =   "20120909"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   2
            Left            =   240
            TabIndex        =   100
            Top             =   480
            Width           =   1092
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "串口监视："
         Height          =   6135
         Left            =   8280
         TabIndex        =   70
         Top             =   720
         Width           =   5295
         Begin VB.CommandButton Command7 
            Caption         =   "清空显示"
            Height          =   372
            Left            =   4200
            TabIndex        =   83
            Top             =   5760
            Width           =   1095
         End
         Begin VB.ListBox List1 
            Height          =   5628
            Left            =   0
            OLEDragMode     =   1  'Automatic
            TabIndex        =   71
            Top             =   0
            Width           =   5655
         End
         Begin VB.Label Label_buffer 
            BackColor       =   &H80000017&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.4
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   2040
            TabIndex        =   73
            Top             =   5760
            Width           =   2172
         End
         Begin VB.Label Label_state 
            BackColor       =   &H80000017&
            Caption         =   " "
            BeginProperty Font 
               Name            =   "微软雅黑"
               Size            =   14.4
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   372
            Left            =   0
            TabIndex        =   72
            Top             =   5760
            Width           =   2052
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "数据采集设定"
         Height          =   6135
         Left            =   2520
         TabIndex        =   54
         Top             =   600
         Width           =   5535
         Begin VB.TextBox TextCY 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.6
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   600
            TabIndex        =   77
            Text            =   "5"
            Top             =   5520
            Width           =   525
         End
         Begin VB.CommandButton Command9 
            Caption         =   "每隔       分钟采样"
            Height          =   615
            Left            =   120
            TabIndex        =   11
            Top             =   5400
            Width           =   1935
         End
         Begin VB.CommandButton Command1 
            Caption         =   "同步时钟"
            Height          =   615
            Left            =   2640
            TabIndex        =   12
            Top             =   5400
            Width           =   1095
         End
         Begin VB.ComboBox Combo8 
            Height          =   300
            Left            =   1440
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   3000
            Width           =   1215
         End
         Begin VB.CommandButton Command8 
            Caption         =   "3 发送命令(&G)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3480
            TabIndex        =   3
            Top             =   3840
            Width           =   1815
         End
         Begin VB.TextBox TextSEC 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3480
            TabIndex        =   68
            Text            =   "00"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox TextMIN 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2640
            TabIndex        =   66
            Text            =   "05"
            Top             =   2280
            Width           =   495
         End
         Begin VB.TextBox TextHOUR 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   64
            Text            =   "00"
            Top             =   2280
            Width           =   375
         End
         Begin VB.TextBox TextDD 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3960
            TabIndex        =   62
            Text            =   "15"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox TextMM 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   3120
            TabIndex        =   60
            Text            =   "09"
            Top             =   1560
            Width           =   495
         End
         Begin VB.TextBox TextYY 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.4
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   2040
            TabIndex        =   58
            Text            =   "2012"
            Top             =   1560
            Width           =   735
         End
         Begin VB.ComboBox Combo7 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   840
            Width           =   1095
         End
         Begin VB.ComboBox Combo6 
            Height          =   300
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label32 
            Caption         =   "存储位置："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   75
            Top             =   3000
            Width           =   1215
         End
         Begin VB.Label Label20 
            Caption         =   "快捷命令："
            Height          =   255
            Left            =   120
            TabIndex        =   74
            Top             =   4920
            Width           =   1455
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000B&
            X1              =   0
            X2              =   5520
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Label LabelSEC 
            Caption         =   "秒"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4080
            TabIndex        =   69
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Labelmin 
            Caption         =   "分"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   67
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label Labelhour 
            Caption         =   "时"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   65
            Top             =   2400
            Width           =   255
         End
         Begin VB.Label LabelDD 
            Caption         =   "日"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4560
            TabIndex        =   63
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LabelMM 
            Caption         =   "月"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3720
            TabIndex        =   61
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label LabelYear 
            Caption         =   "年"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   59
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label31 
            Caption         =   "时间设定："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   57
            Top             =   1320
            Width           =   1215
         End
         Begin VB.Label Label30 
            Caption         =   "命    令："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   56
            Top             =   840
            Width           =   1215
         End
         Begin VB.Label Label24 
            Caption         =   "目标设备："
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   55
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "串口设置"
         Height          =   4455
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   2055
         Begin VB.ComboBox Combo5 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1680
            Width           =   1215
         End
         Begin VB.ComboBox Combo4 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1320
            Width           =   1215
         End
         Begin VB.ComboBox Combo3 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   960
            Width           =   1215
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   600
            Width           =   1215
         End
         Begin VB.CommandButton Command2 
            Caption         =   "1 打开串口(&O)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   1
            Top             =   2040
            Width           =   1455
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton Command6 
            Caption         =   "清空"
            Height          =   495
            Left            =   120
            TabIndex        =   8
            Top             =   3720
            Width           =   1335
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H000000FF&
            BorderColor     =   &H00FFFFFF&
            FillStyle       =   0  'Solid
            Height          =   375
            Left            =   1560
            Shape           =   3  'Circle
            Top             =   2160
            Width           =   375
         End
         Begin VB.Label Label15 
            Caption         =   "校验位"
            Height          =   375
            Left            =   120
            TabIndex        =   34
            Top             =   1720
            Width           =   615
         End
         Begin VB.Label Label16 
            Caption         =   "停止位"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   1380
            Width           =   615
         End
         Begin VB.Label Label17 
            Caption         =   "数据位"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   1020
            Width           =   615
         End
         Begin VB.Label Label18 
            Caption         =   "波特率"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   660
            Width           =   615
         End
         Begin VB.Label Label19 
            Caption         =   "端口号"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   330
            Width           =   735
         End
         Begin VB.Label Label29 
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   29
            Top             =   3240
            Width           =   1095
         End
         Begin VB.Label Label28 
            BackColor       =   &H00C0C0C0&
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   840
            TabIndex        =   28
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label Label27 
            Caption         =   "接收"
            Height          =   375
            Left            =   240
            TabIndex        =   27
            Top             =   3360
            Width           =   375
         End
         Begin VB.Label Label26 
            Caption         =   "发送"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   2880
            Width           =   375
         End
      End
      Begin VB.CommandButton Command3 
         Caption         =   "手动添加数据"
         Height          =   615
         Left            =   -62760
         TabIndex        =   24
         Top             =   9480
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "从机状态"
         Height          =   2415
         Left            =   120
         TabIndex        =   21
         Top             =   5160
         Width           =   2055
         Begin VB.CommandButton Command5 
            Caption         =   "2 搜索从机(&L)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   120
            TabIndex        =   2
            Top             =   1680
            Width           =   1695
         End
         Begin VB.Label Label22 
            Caption         =   "DEV5"
            Enabled         =   0   'False
            ForeColor       =   &H8000000A&
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   53
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "DEV4"
            Enabled         =   0   'False
            ForeColor       =   &H8000000A&
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   52
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "DEV3"
            Enabled         =   0   'False
            ForeColor       =   &H8000000A&
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   51
            Top             =   960
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "DEV2"
            Enabled         =   0   'False
            ForeColor       =   &H8000000A&
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   50
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "DEV1"
            Enabled         =   0   'False
            ForeColor       =   &H8000000A&
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   49
            Top             =   480
            Width           =   1455
         End
         Begin VB.Label Label22 
            Caption         =   "DEV0"
            Enabled         =   0   'False
            ForeColor       =   &H8000000A&
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   48
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label Label23 
            Height          =   495
            Left            =   5880
            TabIndex        =   23
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label25 
            Height          =   495
            Left            =   11160
            TabIndex        =   22
            Top             =   600
            Width           =   1335
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "发送(&s)"
         Height          =   372
         Left            =   12480
         TabIndex        =   20
         Top             =   7200
         Width           =   1095
      End
      Begin VB.TextBox TextSend 
         Height          =   372
         Left            =   3720
         TabIndex        =   19
         Top             =   7200
         Width           =   8655
      End
      Begin VB.Frame Channel 
         Caption         =   "Frame5"
         Height          =   7452
         Index           =   0
         Left            =   -72800
         TabIndex        =   78
         Top             =   840
         Width           =   1800
         Begin VB.TextBox TextFMid 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   146
            Text            =   "main.frx":0074
            Top             =   4800
            Width           =   1600
         End
         Begin VB.TextBox TextFHigh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   128
            Top             =   6120
            Width           =   1600
         End
         Begin VB.TextBox TextFLow 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   0
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   127
            Text            =   "main.frx":0076
            Top             =   3480
            Width           =   1600
         End
         Begin VB.Label Labelweihai 
            Caption         =   "冻伤危害性较小"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   10.8
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   148
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label LabelFTeq 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   92
            Top             =   2760
            Width           =   1332
         End
         Begin VB.Label LabelFECT 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   91
            Top             =   2400
            Width           =   1332
         End
         Begin VB.Label LabelFWSC 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   90
            Top             =   2040
            Width           =   1332
         End
         Begin VB.Label LabelFWindSpeed 
            Caption         =   "2.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   89
            Top             =   1680
            Width           =   1332
         End
         Begin VB.Label LabelFTemp 
            Caption         =   "-20.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   88
            Top             =   1320
            Width           =   1092
         End
         Begin VB.Label LabelFTime 
            Caption         =   "04:08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   87
            Top             =   840
            Width           =   1332
         End
         Begin VB.Label LabelDate 
            Caption         =   "20120909"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   0
            Left            =   240
            TabIndex        =   86
            Top             =   480
            Width           =   1092
         End
      End
      Begin VB.Frame Channel 
         Caption         =   "Frame5"
         ForeColor       =   &H00000000&
         Height          =   7452
         Index           =   1
         Left            =   -71000
         TabIndex        =   79
         Top             =   840
         Visible         =   0   'False
         Width           =   1800
         Begin VB.TextBox TextFLow 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   131
            Text            =   "main.frx":007A
            Top             =   3480
            Width           =   1600
         End
         Begin VB.TextBox TextFMid 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   130
            Text            =   "main.frx":007E
            Top             =   4800
            Width           =   1600
         End
         Begin VB.TextBox TextFHigh 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            Height          =   1200
            Index           =   1
            Left            =   120
            MultiLine       =   -1  'True
            TabIndex        =   129
            Text            =   "main.frx":0080
            Top             =   6120
            Width           =   1600
         End
         Begin VB.Label Labelweihai 
            Caption         =   "较小"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   149
            Top             =   3240
            Width           =   1572
         End
         Begin VB.Label LabelFTeq 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   99
            Top             =   2760
            Width           =   1332
         End
         Begin VB.Label LabelFECT 
            Caption         =   "-25.4"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   98
            Top             =   2400
            Width           =   1452
         End
         Begin VB.Label LabelFWSC 
            Caption         =   "30"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   97
            Top             =   2040
            Width           =   1332
         End
         Begin VB.Label LabelFWindSpeed 
            Caption         =   "2.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   96
            Top             =   1680
            Width           =   1452
         End
         Begin VB.Label LabelFTemp 
            Caption         =   "-20.1"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   95
            Top             =   1200
            Width           =   1092
         End
         Begin VB.Label LabelFTime 
            Caption         =   "04:08"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   240
            TabIndex        =   94
            Top             =   840
            Width           =   1452
         End
         Begin VB.Label LabelDate 
            Caption         =   "20120909"
            BeginProperty Font 
               Name            =   "隶书"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   1
            Left            =   120
            TabIndex        =   93
            Top             =   480
            Width           =   1092
         End
      End
      Begin VB.Label LabelUnconnect 
         Caption         =   "检测仪全部断开连接，请重连。"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   36
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Left            =   -71040
         TabIndex        =   85
         Top             =   3600
         Visible         =   0   'False
         Width           =   9375
      End
      Begin VB.Line Line1 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   -74760
         X2              =   -72600
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Line Line7 
         X1              =   -75000
         X2              =   -75000
         Y1              =   0
         Y2              =   6960
      End
      Begin VB.Label Label4 
         Caption         =   "检测时间:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   47
         Top             =   1680
         Width           =   1692
      End
      Begin VB.Label Label13 
         Caption         =   "高强度作业:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   46
         Top             =   7080
         Width           =   1692
      End
      Begin VB.Label Label12 
         Caption         =   "中等强度作业:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74880
         TabIndex        =   45
         Top             =   5640
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "安静作业:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   44
         Top             =   4440
         Width           =   1692
      End
      Begin VB.Label Label10 
         Caption         =   "冻伤危害性:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   372
         Left            =   -74760
         TabIndex        =   43
         Top             =   4080
         Width           =   1932
      End
      Begin VB.Label Label9 
         Caption         =   "相当温度:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   42
         Top             =   3600
         Width           =   1692
      End
      Begin VB.Label Label8 
         Caption         =   "等价制冷温度:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   41
         Top             =   3240
         Width           =   2292
      End
      Begin VB.Label Label7 
         Caption         =   "风冷指数:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   40
         Top             =   2880
         Width           =   1692
      End
      Begin VB.Label Label6 
         Caption         =   "风    速:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   39
         Top             =   2520
         Width           =   1692
      End
      Begin VB.Label Label5 
         Caption         =   "环境温度:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   38
         Top             =   2160
         Width           =   1692
      End
      Begin VB.Label Label3 
         Caption         =   "检测仪的部署："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74880
         TabIndex        =   37
         Top             =   840
         Width           =   2172
      End
      Begin VB.Label Label14 
         Caption         =   "检测日期:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   -74760
         TabIndex        =   36
         Top             =   1320
         Width           =   1692
      End
      Begin VB.Line Line2 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   -75000
         X2              =   -72600
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   -74760
         X2              =   -72720
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line4 
         BorderStyle     =   5  'Dash-Dot-Dot
         X1              =   -74760
         X2              =   -72720
         Y1              =   3960
         Y2              =   3960
      End
      Begin VB.Label Label21 
         Caption         =   "手动发送："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.4
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   35
         Top             =   7200
         Width           =   1455
      End
   End
   Begin VB.Line Line11 
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   6960
   End
   Begin VB.Label LabelTime 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.6
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   9240
      TabIndex        =   18
      Top             =   720
      Width           =   4452
   End
   Begin VB.Label Label2 
      Caption         =   "当前时间："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.6
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   7800
      TabIndex        =   17
      Top             =   720
      Width           =   1692
   End
   Begin VB.Label Label1 
      Caption         =   "冷强度检测结果实时显示"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   21.6
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   2640
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************通信相关寄存器定义************
'常量定义
Const FRAME_MAX_LEN = 80   '最大帧长
Const FRAME_MIN_LEN = 5   '最小帧长
Const FRAME_HEAD = &HFA   '帧头
Const FRAME_END = &HFB    '帧尾

'id
Const DEV1ID = &HB1
Const DEV2ID = &HB2
Const DEV3ID = &HB3
Const DEV4ID = &HB4
Const DEV5ID = &HB5
Const DEV6ID = &HB6

'串口接收状态机 状态定义
Const RecIdle = 0   '空状态
Const RecRead = 1   '读缓存
Const RecCheck = 2  '校验
Const RecDeal = 3   '数据处理
Const RecRetry = 4  '重试
Const RecFind = 5  '查找


'************数据库定义************
Dim Conn As New ADODB.Connection
Dim rs As New ADODB.Recordset


'************设备信息相关结构体************
Private Type DEVICEDRIVER
id As Integer
name As String
Date As String
Time As String
Temperature As Single
WindSpeed As Single
WCI As Single
ECT As Single
TEQ As Single
WeiHai As Byte
LowLabor As Byte
MidLabor As Byte
HighLabor As Byte
End Type

Private Type DeviceState
DR(6) As DEVICEDRIVER '设备信息
DeviceCount As Integer '设备个数

End Type
'定义设备信息结构体
Dim DS1 As DeviceState
'ReDim DS1.DR(DS1.DeviceCount) '重定义设备个数，不保留之前的数据
'ReDim Preserve DS1.DR(DS1.DeviceCount) '重定义设备个数，保留之前的数据


'************实时数据保存开关************
Dim SaveToDb As Boolean
Dim AlwaysSaveToDb As Boolean
Dim RecData(40) As Byte
Dim SndData(12) As Byte
Dim SndCount As Integer
Dim CheckSum As Byte
Dim RecSum As Byte
Dim RecState As Integer  '指示当前接收状态
'用于发送的帧数据
Dim MainState As Byte '上位机状态存储
Const MainIlde = 0
Const MainWait = 1
Const MainDeal = 2
'定义一个缓冲池 用于存储串口数据
Dim RecBuf As Collection
 
Public Function adodbjet(Optional DBfile As String, Optional pwd As String) As ADODB.Connection
On Error Resume Next
Dim defDB As String
defDB = App.Path & "\db.mdb"
If DBfile = "" Then DBfile = defDB
Set adodbjet = New ADODB.Connection
adodbjet.Open "Provider=Microsoft.Jet.OLEDB4.0;" & "Persist Security Info=False;User ID=Admin;" & _
"Jet OLEDB:Database Password=" & pwd & ";"
'Data Source = " & DBfile"
If adodbjet.State <> 1 Then Set adodbjet = Nothing
End Function
 
 
'Sub data()
'On Error GoTo err
'Dbpath = App.Path & "\db.mdb"
'Constr = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & Dbpath
'Conn.Open Constr
'err:
'If err.Number Then
'MsgBox "数据库出错"
'End
'End If
'End Sub


Private Sub Command12_Click()
'Row = "天安门" & vbTab & "故宫" & vbTab & iijj & vbTab & "abc"
'iijj = iijj + 1

MSFlexGrid1.AddItem Row
'MSFlexGrid1.AddItem Row, 2

End Sub

Private Sub DataGrid1_Click()

End Sub

'*******************************界面初始化************************************
'
'初始化程序
Private Sub Form_Load()
'当前时间显示
LabelTime.Caption = Now

'计时器1 初始化
Timer1.Enabled = True
Timer1.Interval = 200   '20ms触发一次

'----------------初始化串口控件------------------
For i = 0 To 19 Step 1
Combo1.AddItem "COM" & (i + 1)
Next i
Combo1.ListIndex = 2
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
Combo2.ListIndex = 5

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
'''''''''''''''''''''''''''''''''''''
'数据采集设定
'目标设备
Combo6.List(0) = "全部从机"
Combo6.ListIndex = 0
'命令
Combo7.List(0) = "连接从机"
Combo7.List(1) = "单次检测"
Combo7.List(2) = "循环检测"
Combo7.List(3) = "查询数据"
Combo7.List(4) = "同步时钟"
Combo7.List(5) = "终止检测"
Combo7.ListIndex = 1
'数据存储方式选择
Combo8.List(0) = "计算机+检测仪"
Combo8.List(1) = "计算机"
Combo8.List(2) = "检测仪"
Combo8.ListIndex = 1
'''''''''''''''''''''''''''''''''''''
'按Command2键的快捷键设定
Command2.Caption = "打开串口(&O)"
'设置所有设备状态为关闭
DeviceCount = 0

'设置从机初始名字和颜色状态
Call DevStateRefresh
'初始化选项卡
SSTab1.Tab = 0
'list 初始化
List1.Clear    '清除list内容
'初始化数据缓冲池
 Set RecBuf = New Collection
'初始化接收状态机状态到IDLE
RecState = RecIdle
'数据保存开关初始化
SaveToDb = False
AlwaysSaveToDb = False
'数据库加载

End Sub


' *************************************TAB1控件处理*************************************
'
'“打开串口” 按钮控制
Private Sub Command2_Click()
    '根据combo5的值确定校验位设置
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
'串口设定        ：  波特率       ， 校验位      ，     数据位       ，    停止位
MSComm1.Settings = Combo2.Text & "," & s_tem$ & "," & Combo3.Text & "," & Combo4.Text
'错误处理
On Error GoTo ErrMSCOMM
'串口开关指示
If (MSComm1.PortOpen = False) Then
    MSComm1.CommPort = Combo1.ListIndex + 1
    MSComm1.PortOpen = True
    Shape1.FillColor = &HFF000
    Combo1.Enabled = False
    Combo2.Enabled = False
    Combo3.Enabled = False
    Combo4.Enabled = False
    Combo5.Enabled = False
    Command2.Caption = "关闭串口(&O)"
    'MsgBox ("com=" & MSComm1.CommPort & "setting=" & MSComm1.Settings)
ElseIf MSComm1.PortOpen = True Then
    MSComm1.PortOpen = False
    Shape1.FillColor = &H0
    Combo1.Enabled = True
    Combo2.Enabled = True
    Combo3.Enabled = True
    Combo4.Enabled = True
    Combo5.Enabled = True
   
    Command2.Caption = "打开串口(&O)"
 'MsgBox ("com=" & MSComm1.CommPort & "setting=" & MSComm1.Settings)
End If
Exit Sub
ErrMSCOMM: MsgBox ("COM" & MSComm1.CommPort & "通讯错误，请核对串口号是否正确，并确认其未被占用！")
End Sub

'
'发送命令 下拉列表处理
Private Sub Combo7_Click()
        TextYY.Enabled = False
        TextMM.Enabled = False
        TextDD.Enabled = False
        TextHOUR.Enabled = False
        TextMIN.Enabled = False
        TextSEC.Enabled = False
        Combo8.Enabled = False
    If Combo7.ListIndex = 0 Then
    List1.AddItem ("[连接从机]")

    ElseIf Combo7.ListIndex = 1 Then
    List1.AddItem ("[单次检测]")
    Combo8.Enabled = True
    ElseIf Combo7.ListIndex = 2 Then
    List1.AddItem ("[循环检测]")
        TextHOUR.Enabled = True
        TextMIN.Enabled = True
        TextSEC.Enabled = True
        Combo8.Enabled = True
    ElseIf Combo7.ListIndex = 3 Then
    List1.AddItem ("[查询数据]")
        TextYY.Enabled = True
        TextMM.Enabled = True
        TextDD.Enabled = True
    ElseIf Combo7.ListIndex = 4 Then
    List1.AddItem ("[同步时钟]")
        TextYY.Text = Year(Now)
        TextMM.Text = Month(Now)
        TextDD.Text = Day(Now)
        TextHOUR.Text = Hour(Now)
        TextMIN.Text = Minute(Now)
        TextSEC.Text = Second(Now)
        List1.AddItem ("将当前时钟发送到所有相连的从机上")
        TextYY.Enabled = True
        TextMM.Enabled = True
        TextDD.Enabled = True
        TextHOUR.Enabled = True
        TextMIN.Enabled = True
        TextSEC.Enabled = True
        
    ElseIf Combo7.ListIndex = 5 Then
        List1.AddItem ("[终止检测]")
    End If
End Sub

'
'“同步时钟”快捷键
Private Sub Command1_Click()
 Combo6.ListIndex = 0
 Combo7.ListIndex = 1
 Combo6.ListIndex = 0
 Combo7.ListIndex = 4
 Call Command8_Click
End Sub

'
'快捷键响应
Private Sub Command2_KeyDown(KeyCode As Integer, Shift As Integer)
'调用快捷键处理函数 处理快捷键
Call masterkey(KeyCode, Shift)
End Sub

'
'[发送]按钮
Private Sub Command4_Click()

 If (MSComm1.PortOpen = False) Then '判断串口有没有开启
        If MsgBox("串口未开启，是否开启？", vbYesNo) = vbYes Then
                Call Command2_Click
                List1.AddItem ("yes")
        Else
            Exit Sub
        End If
  Else:
        SendMsg (TextSend.Text)
 End If
End Sub

'
'[查找从机]按钮
Private Sub Command5_Click()
If MSComm1.PortOpen = False Then   '判断串口有没有开启
        If MsgBox("串口未开启，是否开启？", vbYesNo) = vbYes Then
                Call Command2_Click
            
        Else
            Exit Sub
        End If
'发送查找命令
Else
            
            List1.AddItem ("搜索全部从机...")
            SndData(0) = &HFA
            SndData(1) = &HD
            SndData(2) = &HFF '搜索从机命令 发向所有从机
            SndData(3) = &HD0
            For i = 4 To 9
            SndData(i) = 0
            Next i
            SndData(10) = &H0
            SndData(11) = &H0
            SndData(12) = &HFB
            SndCount = 13
            SendByte
            DS1.DeviceCount = 0 '关闭所有从机
            '删除从机选项
            Combo6.Clear
            '添加默认从机选项
            Combo6.AddItem ("全部从机")
            '设置显示默认选项
            Combo6.ListIndex = 0
            Call Tab1Refresh  '刷新tab2显示
            SndCount = 0

End If
End Sub

'
'清空收发计数器
Private Sub Command6_Click()
'如果清空按下 清空串口收发计数
Label28.Caption = 0
Label29.Caption = 0
Call Tab1Refresh
End Sub


'
'清空接收区对话框
Private Sub Command7_Click()
'清空串口接收区的显示
List1.Clear
End Sub

'
'发送命令 按钮
Private Sub Command8_Click()
    Dim StorageMode As Byte
    
    If (MSComm1.PortOpen = False) Then
        If MsgBox("串口未开启，是否开启？", vbYesNo) = vbYes Then
                Call Command2_Click
        Else
                Exit Sub
        End If
    Else
        '清空发送寄存器
        For i = 0 To 12
            SndData(i) = 0
        Next i
        '设备选择
        If Combo6.ListIndex = 0 Then
            SndData(2) = &HFF
        ElseIf Combo6.ListCount >= 2 Then
        
        End If
        StorageMode = Combo8.ListIndex
        '帧头
        SndData(0) = &HFA
        '帧尾
        SndData(12) = &HFB
        '帧长  上位机发送数据帧长固定
        SndData(1) = &HD
        '
        '操作选择
        SndData(3) = &HD0 + Combo7.ListIndex
        '附加信息设定
         Select Case SndData(3)
            Case &HD0   '连接从机
                 Call DevStateRefresh
                 
            Case &HD1   '单次检测
                 SndData(4) = StorageMode ' 第4位存储检测模式
            Case &HD2   '循环检测
                SndData(4) = StorageMode ' 第4位存储检测模式
                SndData(5) = Val(TextHOUR.Text)
                SndData(6) = Val(TextMIN.Text)
                SndData(7) = Val(TextSEC.Text)
            Case &HD3   '查询数据
                SndData(4) = Val(TextYY.Text - 2000)
                SndData(5) = Val(TextMM.Text)
                SndData(6) = Val(TextDD.Text)
            Case &HD4   '同步时钟
                SndData(4) = Val(TextYY.Text - 2000)
                SndData(5) = Val(TextMM.Text)
                SndData(6) = Val(TextDD.Text)
                SndData(7) = Val(TextHOUR.Text)
                SndData(8) = Val(TextMIN.Text)
                SndData(9) = Val(TextSEC.Text)
            Case &HD5  '终止检测
                 
        End Select
    
    SndCount = 13
    SendByte   '发送数据
    SndCount = 0
    End If
    
End Sub

'
'每N分钟采样 快捷键
Private Sub Command9_Click()
    TextMIN.Text = TextCY.Text
    TextHOUR.Text = 0
    TextSEC.Text = 0
    Combo6.ListIndex = 0
    Combo7.ListIndex = 2
    Call Command8_Click
    
End Sub

'
'快捷键处理
Private Function masterkey(KeyCode As Integer, Shift As Integer)
'快捷键处理函数 ，通过F2 F3 F4 切换显示窗口
Const f2 = 113   'f2 键值
Const F3 = 114
Const F4 = 115
If KeyCode = f2 Then
    SSTab1.Tab = 0
ElseIf KeyCode = F3 Then
    SSTab1.Tab = 1
ElseIf KeyCode = F4 Then
    SSTab1.Tab = 2
 
End If
End Function



'
'串口接收 事件
Private Sub MSComm1_OnComm()
'串口数据接收并保存
Dim yy As Long
Timer1.Enabled = False
Timer2.Enabled = False
Select Case MSComm1.CommEvent
Case comEvReceive
    Dim xx() As Byte
 
    xx = MSComm1.Input  '接收数据
    For yy = 0 To UBound(xx)    '把串口接收到的所有数据都先保存到数据缓冲池暂时不去处理，因为无法保证数据已经接收完整
        RecBuf.Add xx(yy)
       'List1.AddItem (Hex(xx(yy))) '实时显示接收到的字节，调试用
    Next yy
    Label29.Caption = Label29.Caption + yy
End Select
Timer1.Enabled = True
Timer2.Enabled = True
End Sub


'
'快捷键响应
Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
Call masterkey(KeyCode, Shift)
End Sub

'
'crc校验
Private Function CheckFrame(crc1 As Byte, crc2 As Byte) As Boolean
'   取出数据 计算校验  从缓冲池中删除本条数据 返回校验结果
    If crc1 = &HCC And crc2 = &HCC Then
        CheckFrame = True
    Else
        CheckFrame = False
    End If
    End Function
     
 
'
'时间框调整后 时间输入合法性检查
Private Sub TextCY_LostFocus()
    If TextCY.Text > 59 Or TextCY < 0 Then
                                           TextCY = 5
    End If
End Sub

Private Sub TextDD_LostFocus()
    If TextMM.Text = 1 Or TextMM.Text = 3 Or TextMM.Text = 5 Or TextMM.Text = 7 Or TextMM.Text = 8 Or TextMM.Text = 10 Or TextMM.Text = 12 Then
        If TextDD.Text > 31 Or TextDD.Text < 1 Then
            TextDD.Text = 1
        End If
   Else
        If TextDD.Text > 30 Or TextDD.Text < 1 Then
            TextDD.Text = 1
        End If
   End If
End Sub

Private Sub TextHOUR_LostFocus()
    If TextHOUR.Text > 23 Or TextHOUR < 0 Then
                                            TextHOUR = 0
    End If
End Sub
Private Sub Textsec_LostFocus()
    If TextSEC.Text > 59 Or TextSEC < 0 Then
                                            TextSEC = 0
    End If
End Sub
Private Sub TextMIN_LostFocus()
    If TextMIN.Text > 59 Or TextMIN < 0 Then
                                            TextMIN = 0
    End If
End Sub

Private Sub TextMM_LostFocus()
    If Val(TextMM.Text) > 12 Or Val(TextMM.Text < 1) Then
                TextMM.Text = 1
    End If
End Sub

Private Sub TextYY_LostFocus()
    If TextYY.Text > 2099 Or TextYY.Text < 2000 Then
                    TextYY.Text = 2012
    End If
End Sub

'
'定时器1 中断处理
Private Sub Timer1_Timer()
'task1 更新主界面时间显示
LabelTime.Caption = Now
End Sub
 
'
'串口发送子程序
Private Sub SendMsg(send$)
If send$ <> "" And MSComm1.PortOpen = True Then
'计数器自增
Label28.Caption = Val(Label28.Caption) + Len(send$)
'数据送出
MSComm1.Output = send$
End If
End Sub

'
'串口发送子程序
'说明 ： 从SndData()中取出13字节数据 按二进制格式发送
Private Sub SendByte()
If SndCount > 0 And MSComm1.PortOpen = True Then
'计数器自增
Label28.Caption = Val(Label28.Caption) + SndCount
For i = 0 To 12
    If SndData(i) <= 15 Then
        add_0$ = "0"
    Else: add_0$ = ""
    End If
    temp$ = temp$ & " " & add_0$ & Hex(SndData(i))
Next i
List1.AddItem ("HOST：" & temp$)
temp$ = ""
'数据送出
MSComm1.Output = SndData
End If
End Sub
  
'
'在定时器2中处理接收数据的状态转移 以及数据处理
Private Sub Timer2_Timer()
List1.ListIndex = List1.ListCount - 1
'On Error GoTo commandreset
Label_buffer.Caption = "BUF=" & RecBuf.Count
Label_state.Caption = "STATE =" & RecState

        Dim FrameEndTmp As Integer
        Dim str As String
        Dim blank As String
        Select Case RecState
            Case RecIdle
DATAPROCESS:
                 If RecBuf.Count > FRAME_MIN_LEN Then
                       RecState = RecRead
                 Else
                       RecState = RecIdle
                       Exit Sub
                 End If
            Case RecRead
                 '删除接收字符串前未成帧数据
                 For xx = 1 To RecBuf.Count
                    If Not RecBuf(1) = FRAME_HEAD Then
                         RecBuf.Remove (1)
                    Else
                   ' List1.AddItem ("H=" & xx)
                        Exit For
                    End If
                  Next xx

                 '删除操作完成后，若帧数据小于最小帧，回到空闲状态 结束子程序
                 If RecBuf.Count < FRAME_MIN_LEN Then
                    RecState = RecIdle
                    Exit Sub
                 End If
            '    List1.AddItem ("BUF COUNT=" & RecBuf.Count)
                  '找帧尾
                  For xx = 1 To RecBuf.Count
                        If RecBuf(xx) = FRAME_END Then
                       '如果找到帧尾
                         ' List1.AddItem ("E=" & xx)
                          RecState = RecCheck
                          FrameEndTmp = xx
                           ' List1.AddItem ("FrameEndTmp=" & FrameEndTmp)
                         'RecBuf.Remove (1)   '从缓冲区移除帧尾
                         GoTo RecCheckProcess
                    Else
                        RecState = RecIdle
                        
                    End If
                   
                  Next xx
  
  

            Case RecCheck   'now head = recbuf(1) and end = recbuf(xx)

RecCheckProcess:
                If (FrameEndTmp > FRAME_MAX_LEN Or FrameEndTmp <> RecBuf(2)) Then
                
                   'Or CheckFrame(RecBuf(FrameEndTmp - 1), RecBuf(FrameEndTmp - 2)
                   ' List1.AddItem ("帧校验失败")    '如果帧长大于最大帧长 或者 帧长与帧内长度标识不符 或者crc校验失败  :移除帧头 返回空状态
                    RecBuf.Remove (1)
                    RecState = RecIdle
                    GoTo DATAPROCESS
                Else
                 
                str = "DEV: "
                For i = 1 To FrameEndTmp
                    RecData(i) = RecBuf(1)
                    RecBuf.Remove (1)
                    '接收框显示命令
                    If (RecData(i) < 16) Then '此判断为显示补零
                    blank = "0"
                    Else
                    blank = ""
                    End If
                    str = str & blank & Hex(RecData(i)) & " "
                Next i
                List1.AddItem (str)
                RecState = RecDeal
                GoTo RECDEALPROCESSS
               End If
            Case RecDeal
RECDEALPROCESSS:
                '数据格式
                '帧头    长度    目的地址   命令  内容     校验    帧尾
                'FA       AA      B1        D1    01 03    00 00      FB
                Dim RecLen As Byte
                Dim RecCmd As Byte
                Dim RecDest As Byte
                Dim RecSrc As Byte
                RecLen = RecData(2)
                RecDest = RecData(3)
                RecSrc = RecData(4)
                RecCmd = RecData(5)
                
        If RecDest = 0 Then
                Select Case RecCmd
                    Case &HD0   '从机返回设备信息命令
                            Dim DeviceName As String
                            DeviceName = "DEVICE"  '默认设备名
                            If RecLen >= 8 Then  '查看命令中包含的设备名
                                DeviceName = ""
                                For xx = 6 To RecLen - 3
                                    DeviceName = DeviceName + Chr(RecData(xx))
                                Next xx
                                
                            End If
                        If DS1.DeviceCount > 0 Then
                            For devcount_tmp = 0 To DS1.DeviceCount - 1 '遍历所有现有设备
                                If DS1.DR(devcount_tmp).id = RecSrc Then '若存在重名 ID
                                '此设备已经注册 不进行注册
                                GoTo CMD1_EXIT_LABEL '跳出命令处理函数
                                End If
                            Next devcount_tmp
                        Else '添加新设备
                            Call AddDevice(RecSrc, DeviceName)
                            Call Tab0Refresh 'tab1实时显示界面刷新
                            Call Tab1Refresh 'tab2实时显示界面刷新'添加combo6控件选项 电量设备名称
                            Combo6.AddItem (DeviceName)
                        End If
                        List1.AddItem ("命令 = D0 搜索到从机")
                        List1.AddItem (" 从机编号:" & Hex(RecSrc) & " 从机名:" & DeviceName)
                       
CMD1_EXIT_LABEL:
                        RecState = RecIdle '清除命令状态 退出循环
                        Exit Sub
                    
                        
                    Case &HD1   '将数据存入各自管道 如果收到包结束标识 开启处理流程
                                'D1命令帧格式.
                                'tmp = DeviceIDtoIndex(RecSrc)
                               ' If (tmp = 255) Then
                               '     Call AddDevice(RecSrc, "dev" & RecSrc)
                               ' End If
                               ' tmp = DeviceIDtoIndex(RecSrc)
                                DS1.DR(0).Date = getString(RecData(), 6, 0)
                                List1.AddItem ("命令 D1 收到日期")
                                
                                RecState = RecIdle
                                Exit Sub
                    Case &HD2
                                DS1.DR(0).Time = getString(RecData(), 6, 0)
                                List1.AddItem ("命令 D2 收到时间")
                                RecState = RecIdle
                                Exit Sub
                    Case &HD3
                                DS1.DR(0).Temperature = getString(RecData(), 6, 0)
                                List1.AddItem ("命令 D3 收到温度")
                                RecState = RecIdle
                                Exit Sub
                    Case &HD4
                                DS1.DR(0).WindSpeed = getString(RecData(), 6, 0)
                                RecState = RecIdle
                                Exit Sub
                    Case &HD5
                                DS1.DR(0).WCI = getString(RecData(), 6, 0)
                                List1.AddItem ("命令 D5 WCI")
                                RecState = RecIdle
                                Exit Sub
                    Case &HD6
                                DS1.DR(0).ECT = getString(RecData(), 6, 0)
                                List1.AddItem ("命令 D6 ETC")
                                RecState = RecIdle
                                Exit Sub
                    Case &HD7
                                DS1.DR(0).TEQ = getString(RecData(), 6, 0)
                                List1.AddItem ("命令 D7 TEQ")
                                RecState = RecIdle
                    Case &HD8
                                 DS1.DR(0).WeiHai = RecData(6)
                                 DS1.DR(0).LowLabor = RecData(7)
                                 DS1.DR(0).MidLabor = RecData(8)
                                 DS1.DR(0).HighLabor = RecData(9)
                                 
                              
                                Call Tab1Refresh
                                RecState = RecIdle
                                Exit Sub
                     Case Else

                                 List1.AddItem ("无效命令")
                                 RecState = RecIdle
                                 Exit Sub
                End Select
        End If
        
           
            Case Else
                    '忽视无效命令
                    RecState = RecIdle
                    
        End Select
 
     Exit Sub
commandreset:
             
              List1.AddItem ("数据接收有误 ")
              List1.Clear

End Sub

Private Sub AddDevice(id As Byte, name As String)
    DS1.DeviceCount = 1
    DS1.DR(0).name = name
End Sub

Private Function DeviceIDtoIndex(id As Byte) As Byte
    For tmp = 0 To DS1.DeviceCount - 1
        If DS1.DR(tmp).id = id Then
            DeviceIDtoIndex = tmp
        Else
            DeviceIDtoIndex = 255
        End If
    Next
End Function
Private Function getString(buf() As Byte, buf_start As Byte, buf_end As Byte) As String
    getString1 = ""
    If buf_end = 0 Then
        For xx = buf_start To 255
            If buf(xx) <> 0 And buf(xx) <> &HCC Then
                getString = getString + Chr(buf(xx))
            Else
                Exit Function
            End If
        Next xx
    Else
        For xx = buf_start To buf_end
            getString = getString + Chr(buf(xx))
        Next xx
        Exit Function
    End If
End Function

'
'设置从机状态


'
'刷新从机状态
Private Sub DevStateRefresh()
            Combo6.Clear
            '添加默认从机选项
            Combo6.AddItem ("全部从机")
            '设置显示默认选项
            Combo6.ListIndex = 0
  For i = 0 To 5
    'If IsDevOn(i) = True Then
            Label22(i).ForeColor = &HFF000
            'If Not DevName(i) = "" Then
            '    Label22(i).Caption = DevName(i)
           ' End If
           ' Combo6.AddItem (DevName(i))
   ' Else
           ' Label22(i).ForeColor = &H8000000A
           ' Label22(i).Caption = "Device" & i
           ' DevName(i) = Label22(i).Caption
   ' End If
  Next
    
End Sub
'TAB0 重绘
Private Sub Tab0Refresh()
    '刷新从机状态显示
    For tmp = 0 To 5
        Label22(tmp).Visible = False
    Next
    For tmp = 0 To DS1.DeviceCount - 1
        Label22(tmp).Visible = True
        Label22(tmp).Caption = DS1.DR(tmp).name
        Label22(tmp).ForeColor = &HFF000
    Next
    
End Sub




'********************************界面2控件函数********************************
'
'保存数据 按钮
Private Sub Command10_Click()
'数据库添加数据语句

End Sub
'
'自动保存数据 按钮
Private Sub Command11_Click()
'自动保存 数据结果有更新就自动保存到数据库
If (AlwaysSaveToDb = False) Then
    AlwaysSaveToDb = True
    Command11.Caption = "关闭自动保存"
     Command11.BackColor = &HFF000
Else
    AlwaysSaveToDb = False
    Command11.Caption = "开启自动保存"
    Command11.BackColor = &H8000000F
End If

End Sub
 

'tab2 显示相关函数
'说明： 根据当前活跃的从机数，安排TAB2中显示布局
Private Sub Tab1Refresh()
    tmp_tab = SSTab1.Tab
    SSTab1.Tab = 1
            Channel(0).Visible = True
            Channel(0).Caption = DS1.DR(0).name
            LabelDate(0).Caption = DS1.DR(0).Date
            LabelFTime(0).Caption = DS1.DR(0).Time
            LabelFTemp(0).Caption = Format(DS1.DR(0).Temperature, "0.0") & " " & "℃"
            LabelFWindSpeed(0).Caption = Format(DS1.DR(0).WindSpeed, "0.0") & " " & "m/s"
            LabelFWSC(0).Caption = Format(DS1.DR(0).WCI, "0.0")
            LabelFECT(0).Caption = Format(DS1.DR(0).ECT, "0.0") & " " & "℃"
            LabelFTeq(0).Caption = Format(DS1.DR(0).TEQ, "0.0") & " " & "℃"
            
            If DS1.DR(0).WeiHai = &H30 Then
                Labelweihai(0).Caption = "冻伤危害性小"
            ElseIf DS1.DR(0).WeiHai = &H31 Then
                Labelweihai(0).Caption = "冻伤危害性较大"
            ElseIf DS1.DR(0).WeiHai = &H32 Then
                Labelweihai(0).Caption = "冻伤危害性很大"
            End If
             '低强度劳动
            If DS1.DR(0).LowLabor = &H30 Then
                 TextFLow(0).Text = "缩短作业时间。"
            ElseIf DS1.DR(0).LowLabor = &H31 Then
                 TextFLow(0).Text = "戴面罩；禁涂油彩;缩短作业时间."
            ElseIf DS1.DR(0).LowLabor = &H32 Then
                 TextFLow(0).Text = "取消非必需作业；必需劳动时间<15 min；禁止单独作业；保持皮肤干燥，禁止裸露。"
            Else
                 TextFLow(0).Text = "取消户外作业。"
            End If
                         
             '中强度劳动
            If DS1.DR(0).MidLabor = &H30 Then
                TextFMid(0).Text = "加强劳动监督；保持皮肤干燥，禁止裸露。"
            ElseIf DS1.DR(0).MidLabor = &H31 Then
                TextFMid(0).Text = "加强劳动监督；保持皮肤干燥，禁止裸露；禁涂油彩。"
            ElseIf DS1.DR(0).MidLabor = &H32 Then
                TextFMid(0).Text = "加强劳动监督；保持皮肤干燥，禁止裸露；戴面罩；禁涂油彩。"
            ElseIf DS1.DR(0).MidLabor = &H33 Then
                TextFMid(0).Text = "加强劳动监督；减少非必需的作业；必需劳动时间<30 min。。"
            Else
                TextFMid(0).Text = "取消户外作业。"
            End If
                                
            '高强度劳动
            If DS1.DR(0).HighLabor = &H30 Then
                TextFHigh(0).Text = "加强劳动监督；增加饮水；保持皮肤干燥，禁止裸露。"
            ElseIf DS1.DR(0).HighLabor = &H31 Then
                TextFHigh(0).Text = "禁涂油彩;保持皮肤干燥，禁止裸露；休息时注意保暖。"
            Else
                TextFHigh(0).Text = "取消非必需的作业；必需劳动时间<15 min；禁止单独作业；保持皮肤干燥，禁止裸露。"
            End If
    SSTab1.Tab = tmp_tab
End Sub

'**************************************************************************************************

