VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6930
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   9960
   LinkTopic       =   "Form1"
   ScaleHeight     =   6930
   ScaleWidth      =   9960
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   5000
      Left            =   1320
      Top             =   3840
   End
   Begin VB.Image Image1 
      Height          =   6495
      Left            =   240
      Top             =   240
      Width           =   9495
   End
   Begin VB.Menu File 
      Caption         =   "文件"
      Begin VB.Menu Open 
         Caption         =   "打开"
      End
      Begin VB.Menu Save 
         Caption         =   "另存"
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu view 
      Caption         =   "查看"
      Begin VB.Menu ToLeft 
         Caption         =   "向上"
      End
      Begin VB.Menu ToRight 
         Caption         =   "向下"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Rem 定义目录
Dim showFileList() As String
Dim curIndexForShow As Integer
Dim showFileListLength As Integer







Private Sub Open_Click()
    Form2.Visible = True
    
End Sub

Private Sub Timer1_Timer()
    If showFileListLength > 0 Then
        curIndexForShow = IIf(curIndexForShow < UBound(showFileList), curIndexForShow + 1, 0)
        Image1.Picture = LoadPicture(showFileList(curIndexForShow))
    End If
End Sub

Public Sub resetData(showFileList_data() As String, showFileListLength_data As Integer)
    Dim i As Integer
    showFileListLength = showFileListLength_data
    If showFileListLength_data > 0 Then
        ReDim showFileList(showFileListLength_data - 1)
        For i = 0 To showFileListLength_data - 1
            showFileList(i) = showFileList_data(i)
        Next i
        curIndexForShow = 0
    End If
End Sub

