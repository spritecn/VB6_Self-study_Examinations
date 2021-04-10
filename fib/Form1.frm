VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Fibonacci"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1680
      TabIndex        =   0
      Top             =   2040
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   2
      Text            =   "请输入项数"
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Height          =   180
      Left            =   480
      TabIndex        =   1
      Top             =   960
      Width           =   90
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Text1_Change()
    Dim count As Integer, result()
    count = Val(Text1.text)
    result = fib(count)
    Label1 = arr2Text(result)
    
    
End Sub

Function arr2Text(arr()) As String
    Dim i, lineSize, text
    lineSize = 20
    On Error GoTo z
    i = LBound(arr)
    For i = LBound(arr) To UBound(arr)
        If Len(text) > 20 And Len(text) Mod 20 < 5 Then '换行
            text = text & vbCrLf
        End If
        
        If i <> UBound(arr) Then
            text = text & arr(i) & " "
        Else
            text = text & arr(i)
        End If
    Next i
    arr2Text = text
z:
    arr2Text = text
End Function

Function fib(num As Integer) As Variant
    Dim result(), i
    If num > 0 Then
   
        ReDim result(num - 1)
        For i = 0 To num - 1
            If i = 0 Or i = 1 Then
                result(i) = 1
            Else
                result(i) = result(i - 2) + result(i - 1)
            End If
            
        Next i
    End If
    fib = result
End Function
