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
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   4935
      Begin VB.Label Label3 
         Caption         =   "下拉式菜单"
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
         Left            =   840
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label2 
         Caption         =   "下拉式菜单"
         BeginProperty Font 
            Name            =   "宋体"
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
         Caption         =   "下拉式菜单"
         BeginProperty Font 
            Name            =   "宋体"
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
      Caption         =   "字体"
      Index           =   1
      Begin VB.Menu fontName 
         Caption         =   "字体名称"
         Begin VB.Menu songTi 
            Caption         =   "宋体"
         End
         Begin VB.Menu kaiTi 
            Caption         =   "楷体"
         End
      End
      Begin VB.Menu style 
         Caption         =   "文本风格"
         Begin VB.Menu cuTi 
            Caption         =   "粗体(&B)"
         End
         Begin VB.Menu xieTi 
            Caption         =   "斜体(&I)"
         End
         Begin VB.Menu xihuaxian 
            Caption         =   "下划线(&U)"
         End
      End
   End
   Begin VB.Menu color 
      Caption         =   "颜色"
      Index           =   2
      Begin VB.Menu blueColor 
         Caption         =   "蓝色"
      End
      Begin VB.Menu redColor 
         Caption         =   "红色"
      End
      Begin VB.Menu greenColor 
         Caption         =   "绿色"
      End
   End
   Begin VB.Menu end 
      Caption         =   "结束"
      Index           =   3
      WindowList      =   -1  'True
      Begin VB.Menu exit 
         Caption         =   "退出"
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
Label1.font.Name = "楷体"
Label2.font.Name = "楷体"
Label3.font.Name = "楷体"


End Sub

Private Sub redColor_Click()
Form1.BackColor = vbRed
End Sub


Private Sub songTi_Click()
Label1.font.Name = "宋体"
Label2.font.Name = "宋体"
Label3.font.Name = "宋体"

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
