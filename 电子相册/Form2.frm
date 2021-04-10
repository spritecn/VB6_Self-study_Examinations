VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "选择文件夹"
   ClientHeight    =   3675
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4770
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3675
   ScaleWidth      =   4770
   StartUpPosition =   3  '窗口缺省
   Visible         =   0   'False
   Begin VB.CommandButton Command2 
      Caption         =   "确定"
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "取消"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   3000
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   2400
      Left            =   600
      TabIndex        =   1
      Top             =   480
      Width           =   3615
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   600
      TabIndex        =   0
      Top             =   120
      Width           =   3615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Command1_Click()
    Form2.Visible = False
End Sub

Private Sub Command2_Click()
    Dim showFolder, showFileListLength As Integer
    showFolder = Dir1.Path
    Form2.Visible = False
    '筛选目录下图片文件,赋值给showFileList
   Dim fp, i As Integer
   fp = Dir(showFolder & "\*.jpg")
   Dim tmpArr(1000) As String '最大支持一千个
   Do While fp <> ""
        tmpArr(i) = showFolder & "\" & fp
        i = i + 1
        fp = Dir
   Loop
   showFileListLength = 0
   For i = 0 To 1000
     If tmpArr(i) <> "" Then
        showFileListLength = showFileListLength + 1
    Else
        Exit For
    End If
   Next i
   
   If showFileListLength > 0 Then
        Form1.resetData tmpArr, showFileListLength
   End If
   
End Sub

Private Sub Drive1_Change()
    Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
    Dir1.Path = App.Path & "\jpg"
End Sub
