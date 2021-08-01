VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "数制转换"
   ClientHeight    =   2625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4005
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   3735
      Begin VB.OptionButton Option3 
         Caption         =   "16进制"
         Height          =   180
         Left            =   2520
         TabIndex        =   7
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton Option2 
         Caption         =   "8进制"
         Height          =   180
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton Option1 
         Caption         =   "2进制"
         Height          =   180
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1080
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Text            =   "请输入数字"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "二进制数"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "十进制数"
      Height          =   255
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Const binLabelStr = "二进制数"
Const octLabelStr = "八进制数"
Const hexLabelStr = "16进制数"



Private Sub Option1_Click()
Label2.Caption = binLabelStr
Dim inNum As Integer
If (Text1.Text = "" Or Text1.Text = "请输入数字") Then
    Text2.Text = ""
    Exit Sub
End If
inNum = Val(Text1.Text)
Text2.Text = numToBin(inNum)
End Sub

Private Sub Option2_Click()
Label2.Caption = octLabelStr
Dim inNum As Integer
If (Text1.Text = "") Then
    Text2.Text = ""
    Exit Sub
End If
inNum = Val(Text1.Text)
Text2.Text = Oct(inNum)
End Sub

Private Sub Option3_Click()
Label2.Caption = hexLabelStr
Dim inNum As Integer
If (Text1.Text = "") Then
    Text2.Text = ""
    Exit Sub
End If
inNum = Val(Text1.Text)
Text2.Text = Hex(inNum)
End Sub

Private Sub Text1_Change()
Dim inNum As Integer
'初始或空值直接写空
If (Text1.Text = "") Then
    Text2.Text = ""
    Exit Sub
End If
inNum = Val(Text1.Text)
If (Option1.Value = True) Then
    Label2.Caption = binLabelStr
    Text2.Text = numToBin(inNum)
End If

If (Option2.Value = True) Then
    Label2.Caption = octLabelStr
    Text2.Text = Oct(inNum)
End If

If (Option3.Value = True) Then
    Label2.Caption = hexLabelStr
    Text2.Text = Hex(inNum)
     
End If



End Sub


Private Sub Text1_GotFocus()
'清理初始内容
If (Text1.Text = "请输入数字") Then
Text1.Text = ""
End If

End Sub


Private Function numToBin(inNum As Integer) As String
  Dim result As String
  Dim divisionResult As Integer
  Dim modResult As Integer
  divisionResult = inNum
  result = ""
  If (inNum = 0 Or inNum = 1) Then
    numToBin = CStr(inNum)
    Exit Function
  End If
  
  
  While divisionResult >= 2
    modResult = divisionResult Mod 2
    divisionResult = divisionResult / 2
    result = modResult & result
  Wend
  If (divisionResult > 0) Then
  result = divisionResult & result
  End If
  numToBin = result
End Function
