VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3360
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   5655
   StartUpPosition =   3  '����ȱʡ
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.Label Label3 
         Caption         =   "����ʽ�˵�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "����ʽ�˵�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "����ʽ�˵�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Menu font 
      Caption         =   "����"
      Index           =   1
      Begin VB.Menu fontName 
         Caption         =   "��������"
         Begin VB.Menu songTi 
            Caption         =   "����"
         End
         Begin VB.Menu kaiTi 
            Caption         =   "����"
         End
      End
      Begin VB.Menu style 
         Caption         =   "�ı����"
         Begin VB.Menu cuTi 
            Caption         =   "����(&B)"
         End
         Begin VB.Menu xieTi 
            Caption         =   "б��(&I)"
         End
         Begin VB.Menu xihuaxian 
            Caption         =   "�»���(&U)"
         End
      End
   End
   Begin VB.Menu color 
      Caption         =   "��ɫ"
      Index           =   2
      Begin VB.Menu blueColor 
         Caption         =   "��ɫ"
      End
      Begin VB.Menu redColor 
         Caption         =   "��ɫ"
      End
      Begin VB.Menu greenColor 
         Caption         =   "��ɫ"
      End
   End
   Begin VB.Menu end 
      Caption         =   "����"
      Index           =   3
      WindowList      =   -1  'True
      Begin VB.Menu exit 
         Caption         =   "�˳�"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub blueColor_Click()
Form1.BackColor = vbBlue
End Sub



Private Sub cuTi_Click()
Label1.font.Bold = True
Label2.font.Bold = True
Label3.font.Bold = True
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
    Form1.BackColor = RGB(55, 127, 127)
End Sub

Private Sub greenColor_Click()
Form1.BackColor = vbGreen
End Sub



Private Sub kaiTi_Click()
Label1.font.Name = "����"
Label2.font.Name = "����"
Label3.font.Name = "����"


End Sub

Private Sub redColor_Click()
Form1.BackColor = vbRed
End Sub


Private Sub songTi_Click()
Label1.font.Name = "����"
Label2.font.Name = "����"
Label3.font.Name = "����"

End Sub

Private Sub xieTi_Click()
Label1.font.Italic = True
Label2.font.Italic = True
Label3.font.Italic = True
End Sub

Private Sub xihuaxian_Click()
Label1.font.Underline = True
Label2.font.Underline = True
Label3.font.Underline = True
End Sub
