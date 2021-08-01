VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00808000&
   Caption         =   "浏览学生基本信息数据 表"
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7800
   LinkTopic       =   "Form1"
   ScaleHeight     =   6570
   ScaleWidth      =   7800
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6855
      Begin VB.TextBox Text4 
         DataField       =   "出生日期"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   7
         Top             =   3240
         Width           =   1935
      End
      Begin VB.TextBox Text3 
         DataField       =   "姓别"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   5
         Top             =   2400
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         DataField       =   "姓名"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   3
         Top             =   1560
         Width           =   1935
      End
      Begin VB.TextBox Text1 
         DataField       =   "学号"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "生日"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   4
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "学号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   855
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   735
      Left            =   1200
      Top             =   5400
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1296
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":0000
      OLEDBString     =   $"Form1.frx":00B4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "student"
      Caption         =   "学生信息"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xh, xm, xb, sr





Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
Text_reFresh
End Sub


Private Sub Command1_Click()
Adodc1.Recordset.Update
End Sub

Private Sub Form_Load()
Adodc1.Refresh
Text_reFresh


End Sub


Sub Text_reFresh()
If (Adodc1.Recordset.EOF Or Adodc1.Recordset.BOF) Then
Exit Sub
End If


xh = Adodc1.Recordset.Fields("学号")
xm = Adodc1.Recordset.Fields("姓名")
xb = Adodc1.Recordset.Fields("姓别")
If (xb = True) Then
xb = "男"
Else
xb = "女"
End If
sr = Adodc1.Recordset.Fields("出生日期")
Text1.Text = xh
Text2.Text = xm
Text3.Text = xb
Text4.Text = sr

End Sub


