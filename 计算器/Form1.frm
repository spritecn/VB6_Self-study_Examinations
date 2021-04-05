VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "简单计算器"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4005
   StartUpPosition =   3  '窗口缺省
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   2280
      TabIndex        =   0
      Text            =   "结果"
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   240
      TabIndex        =   2
      Text            =   "输入"
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "支持整数的+ - * /"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Text1_LostFocus()
    Text2.Text = computer(Text1.Text)
End Sub

Public Function computer(inputText As String) As Integer
    '清掉空格
    inputText = Replace(inputText, " ", "")
    If Len(inputText) < 1 Then Exit Function
    '计算器函数,计算 + - * /
    Dim allowCharArr(), test(4) As Integer
    Dim result As Integer
    allowCharArr = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "+", "-", "*", "/")
    '判断是不是全在允许的字符中
    Dim char As String
    Dim isAllowed As Boolean
    isAllowed = True
    Dim i
    For i = 1 To Len(inputText)
        char = Mid(inputText, i, 1)
        If Not isInArr(allowCharArr, char) Then
            isAllowed = False
            Exit For
        End If
    Next i
    If Not isAllowed Then
        MsgBox ("输入错误,仅支持整数加减乘除")
        Text1.Text = ""
        computer = result
        Exit Function
    End If
    
    '正式计算
    '使用一个数组保存操作符和操作数,操作数直接放入数组,操作符如果是乘除,先和下一操作符计算完再放入
    '然后使用这个数组计算加减
    '使用3个操作数和2个操作符来判断哪个先算
    '创建一个数组用于保存第一轮解析后的值
    Dim parseedArr() As String
    '解析数组的最大长度为input的长度
    ReDim parseedArr(Len(inputText))
    Dim parseedArrItem As Integer
    Dim num1 As String, optStr As String, num2 As String
    For i = 1 To Len(inputText)
        char = Mid(inputText, i, 1)
        If IsNumeric(char) Then
            If num2 = "" Then
                num1 = num1 + char
            Else
                num2 = num2 + char
            End If
            '如果是最后一个并且num1有值,根据num2是否有值判断是否需要计算,计算完将num1放入
            If i = Len(inputText) And num1 <> "" Then
                If num2 <> "" And num2 <> " " Then num1 = opt(Val(num1), optStr, Val(num2))
                parseedArr(parseedArrItem) = num1
            End If
            
        Else
            If i = 1 Then
                num1 = num1 + char '仅在第一个字符是操作符时对操作数做为操作数的符号处理
            Else
                            '如果是最后一个并且num1有值,根据num2是否有值判断是否需要计算,计算完将num1放入
                If i = Len(inputText) And num1 <> "" Then
                    If num2 <> "" And num2 <> " " Then num1 = opt(Val(num1), optStr, Val(num2))
                    parseedArr(parseedArrItem) = num1
                End If
                
                '不能同时出现两个符号,比如 1++ 2报错
                If Not IsNumeric(Mid(inputText, i - 1, 1)) Then
                    MsgBox ("不能同时出现两个字符")
                    Text1.Text = ""
                    Exit Function
                End If
                
                If char = "+" Or char = "-" Then
                    '判断上一轮值是不是需要计算
                    If num2 <> "" Then
                        num1 = opt(Val(num1), optStr, Val(num2))
                        num2 = ""
                        optStr = ""
                    Else
                        'num1 和 char放入数组
                        parseedArr(parseedArrItem) = num1
                        num1 = ""
                        
                        parseedArrItem = parseedArrItem + 1
                        parseedArr(parseedArrItem) = char
                        optStr = ""
                        
                        parseedArrItem = parseedArrItem + 1
                    End If
                Else
                     If num2 <> "" Then  '判断上一个次是否需要计算
                        num1 = opt(Val(num1), optStr, Val(num2))
                    End If
                    optStr = char '
                    num2 = " "
                End If
            End If
        End If
    Next i
    
    optStr = ""
    result = 0
    '循环parseedArr算出结果
    For i = 0 To parseedArrItem
        '如果是元素是空的表示是运算结束了
        If parseedArr(i) = "" Then Exit For
        If i = 0 Then
            result = parseedArr(i)
        Else
            If Not IsNumeric(parseedArr(i)) Then
                optStr = parseedArr(i)
            Else
                result = opt(result, optStr, Val(parseedArr(i)))
            End If
        End If
    Next i
    
    '结果返回
    computer = result
End Function

Function opt(num1 As Integer, optStr As String, num2 As Integer) As Integer
    Dim result
    Select Case optStr
        Case "+"
            result = num1 + num2
        Case "-"
            result = num1 - num2
        Case "*"
            result = num1 * num2
        Case "/"
            result = num1 / num2
        Case Else
            'pass
    End Select
    opt = result
End Function

Function isInArr(arr(), var As Variant) As Boolean
    '判断变量在数组中
    Dim result As Boolean
    result = False
    Dim i
    For i = LBound(arr) To UBound(arr)
        If arr(i) = var Then
            result = True
            Exit For
        End If
    Next i
    isInArr = result
End Function


Private Sub Text1_GotFocus()
    If Text1.Text = "输入" Then Text1.Text = ""
End Sub
