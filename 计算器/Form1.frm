VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "�򵥼�����"
   ClientHeight    =   1530
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1530
   ScaleWidth      =   4005
   StartUpPosition =   3  '����ȱʡ
   Begin VB.TextBox Text2 
      Height          =   510
      Left            =   2280
      TabIndex        =   0
      Text            =   "���"
      Top             =   585
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   510
      Left            =   240
      TabIndex        =   2
      Text            =   "����"
      Top             =   585
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "֧��������+ - * /"
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
    '����ո�
    inputText = Replace(inputText, " ", "")
    If Len(inputText) < 1 Then Exit Function
    '����������,���� + - * /
    Dim allowCharArr(), test(4) As Integer
    Dim result As Integer
    allowCharArr = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0", "+", "-", "*", "/")
    '�ж��ǲ���ȫ��������ַ���
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
        MsgBox ("�������,��֧�������Ӽ��˳�")
        Text1.Text = ""
        computer = result
        Exit Function
    End If
    
    '��ʽ����
    'ʹ��һ�����鱣��������Ͳ�����,������ֱ�ӷ�������,����������ǳ˳�,�Ⱥ���һ�������������ٷ���
    'Ȼ��ʹ������������Ӽ�
    'ʹ��3����������2�����������ж��ĸ�����
    '����һ���������ڱ����һ�ֽ������ֵ
    Dim parseedArr() As String
    '�����������󳤶�Ϊinput�ĳ���
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
            '��������һ������num1��ֵ,����num2�Ƿ���ֵ�ж��Ƿ���Ҫ����,�����꽫num1����
            If i = Len(inputText) And num1 <> "" Then
                If num2 <> "" And num2 <> " " Then num1 = opt(Val(num1), optStr, Val(num2))
                parseedArr(parseedArrItem) = num1
            End If
            
        Else
            If i = 1 Then
                num1 = num1 + char '���ڵ�һ���ַ��ǲ�����ʱ�Բ�������Ϊ�������ķ��Ŵ���
            Else
                            '��������һ������num1��ֵ,����num2�Ƿ���ֵ�ж��Ƿ���Ҫ����,�����꽫num1����
                If i = Len(inputText) And num1 <> "" Then
                    If num2 <> "" And num2 <> " " Then num1 = opt(Val(num1), optStr, Val(num2))
                    parseedArr(parseedArrItem) = num1
                End If
                
                '����ͬʱ������������,���� 1++ 2����
                If Not IsNumeric(Mid(inputText, i - 1, 1)) Then
                    MsgBox ("����ͬʱ���������ַ�")
                    Text1.Text = ""
                    Exit Function
                End If
                
                If char = "+" Or char = "-" Then
                    '�ж���һ��ֵ�ǲ�����Ҫ����
                    If num2 <> "" Then
                        num1 = opt(Val(num1), optStr, Val(num2))
                        num2 = ""
                        optStr = ""
                    Else
                        'num1 �� char��������
                        parseedArr(parseedArrItem) = num1
                        num1 = ""
                        
                        parseedArrItem = parseedArrItem + 1
                        parseedArr(parseedArrItem) = char
                        optStr = ""
                        
                        parseedArrItem = parseedArrItem + 1
                    End If
                Else
                     If num2 <> "" Then  '�ж���һ�����Ƿ���Ҫ����
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
    'ѭ��parseedArr������
    For i = 0 To parseedArrItem
        '�����Ԫ���ǿյı�ʾ�����������
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
    
    '�������
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
    '�жϱ�����������
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
    If Text1.Text = "����" Then Text1.Text = ""
End Sub
