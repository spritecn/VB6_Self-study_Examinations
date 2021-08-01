VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form Form1 
   Caption         =   "学生信息管理"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7590
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   7590
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "按姓名搜索"
      Height          =   255
      Left            =   5520
      TabIndex        =   5
      Top             =   120
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "删除"
      Height          =   495
      Left            =   5880
      TabIndex        =   3
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "修改"
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "新增"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   5520
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   4335
      Left            =   360
      TabIndex        =   0
      Top             =   600
      Width           =   7095
      _ExtentX        =   12515
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   6
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public db As New db
Public Type Student
    xh As String
    xm As String
    xb As Boolean
    zy As String
    jxj As Currency
End Type
Public optType As Integer '操作类型:0默认，1  新增，2修改

Private Sub Command1_Click()
optType = 1
Form2.Show
End Sub

Private Sub Command2_Click()
optType = 2
End Sub

Private Sub Form_Load()
db.init
Set MSHFlexGrid1.DataSource = db.recordSet

End Sub

Private Sub MSHFlexGrid1_Click()

End Sub
