VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3645
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2460
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   2460
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command1 
      Caption         =   "ȷ��"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   3000
      Width           =   1815
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "Form2.frx":0000
      Left            =   240
      List            =   "Form2.frx":000A
      TabIndex        =   4
      Text            =   "�Ա�"
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Text            =   "��ѧ��"
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Text            =   "רҵ"
      Top             =   1680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "����"
      Top             =   600
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Text            =   "ѧ��"
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form2.Visible = False


End Sub

Private Sub Form_Load()
If (Form1.optType = 1) Then
    Form2.Caption = "����"
ElseIf (Form1.optType = 2) Then
    Form2.Caption = "�޸�"
End If






End Sub

