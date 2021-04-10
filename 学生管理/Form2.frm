VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4035
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4035
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "保存"
      Height          =   495
      Left            =   2280
      TabIndex        =   3
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Text            =   "年龄"
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Text            =   "姓名"
      Top             =   240
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Text            =   "学号"
      Top             =   240
      Width           =   1575
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    Dim s As student
    s.id = Val(Text1.Text)
    s.name = Text2.Text
    s.age = Val(Text3.Text)
    
    addStudent2File s
    Form1.loadDb
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Form1.Enabled = True
End Sub
