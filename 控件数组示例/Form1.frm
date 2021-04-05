VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4605
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6600
   ForeColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   4605
   ScaleWidth      =   6600
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame2 
      Caption         =   "颜色"
      Height          =   2415
      Left            =   3480
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
      Begin VB.OptionButton Option2 
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   4
         Top             =   120
         Visible         =   0   'False
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "大小"
      Height          =   2415
      Left            =   360
      TabIndex        =   1
      Top             =   1680
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "10#"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      Caption         =   "字体测试"
      Height          =   855
      Left            =   600
      TabIndex        =   0
      Top             =   240
      Width           =   5415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1

Const firstFontSize = 10
Dim fontColors()

Private Sub Form_Load()
    fontColors = Array("Black", "Red", "Yollow", "Blue")
    Dim i As Integer
    For i = 1 To 5
        Load Option1(i)
        Option1(i).Caption = CStr(firstFontSize + i) + "#"
        Option1(i).Top = Option1(i - 1).Top + 300
        Option1(i).Visible = True
    Next i
    
    For i = LBound(fontColors) To UBound(fontColors)
        Load Option2(i)
        Option2(i).Caption = fontColors(i)
        Option2(i).Top = Option2(i - 1).Top + 350
        Option2(i).Enabled = True
        Option2(i).Visible = True
    Next i
    
    
End Sub

Private Sub Option1_Click(Index As Integer)
    Label1.FontSize = 10 + Index
End Sub

Private Sub Option2_Click(Index As Integer)
    Select Case Index
    Case 1
        Label1.ForeColor = vbBlack
    Case 2
        Label1.ForeColor = vbRed
    Case 3
        Label1.ForeColor = vbYellow
    Case 4
        Label1.ForeColor = vbBlue
    End Select
End Sub
