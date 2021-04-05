VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6315
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   15
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   6315
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Caption         =   "日历"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   240
      TabIndex        =   4
      Top             =   840
      Width           =   5775
      Begin VB.Label Label2 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   5415
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0000
      Left            =   4080
      List            =   "Form1.frx":002B
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   240
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "2021"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim data(6, 7) As String '定义一个6*7的数组来存日历参数
Dim header As String
Dim currentYear, currentMonth
Dim days As Integer   '此月有几天



Private Sub Command1_Click()
    currentYear = currentYear - 1
    Label1.Caption = currentYear
    computerData CStr(currentYear), CStr(currentMonth)
    showData
End Sub

Private Sub Command2_Click()
    currentYear = currentYear + 1
    Label1.Caption = currentYear
    computerData CStr(currentYear), CStr(currentMonth)
    showData
End Sub

Private Sub Form_Load()
    header = "一二三四五六日" '表头字符
    currentYear = year(Now)
    currentMonth = month(Now)
    Label1.Caption = currentYear
    Combo1.ListIndex = currentMonth - 1  '设置
    computerData CStr(currentYear), CStr(currentMonth)
    showData
End Sub

Private Sub Combo1_click()
    currentMonth = Combo1.ListIndex + 1  'index从0开始的,所以要加1
    computerData CStr(currentYear), CStr(currentMonth)
    showData

End Sub

Private Sub showData()
    Dim i, j
    Dim dataStr As String, showStr
    dataStr = vbCrLf   '空一行
    For i = 1 To 6
        For j = 1 To 7
            showStr = data(i, j)
            If i <> 1 Then
                showStr = IIf(Len(showStr) = 4, showStr, Space(4 - Len(showStr)) + showStr)
            Else
                showStr = "  " + showStr
            End If
            
            dataStr = dataStr + showStr
            If Not j = 7 Then
                dataStr = dataStr + "  "
            End If
        Next j
        dataStr = dataStr + vbCrLf + vbCrLf
    Next i
    Label2 = dataStr
End Sub

'根据年月计算日历显示数组
Sub computerData(computerYear As String, computerMonth As String)
     Dim firstDayWeek As Integer, i, j, dayAdded As Integer
     days = getDaysByMonth(currentYear & "/" & currentMonth)
     firstDayWeek = Weekday(currentYear & "/" & currentMonth & "/" & 1, vbMonday)
     For i = 1 To 7
        data(1, i) = Mid(header, i, 1)  '第一行放表头
     Next i
     
     Rem 双循环操作数组
     dayAdded = 1
     For i = 2 To 6
        For j = 1 To 7
            If dayAdded = 1 And i = 2 And j < firstDayWeek Then
                data(i, j) = ""
            Else
                If dayAdded <= days Then
                    data(i, j) = "" & dayAdded
                    dayAdded = dayAdded + 1
                Else
                    data(i, j) = ""
                End If
            End If
        Next j
    Next i

End Sub


'计算这个月有几天,输入格式是
Function getDaysByMonth(yeearAndMonth As String) As Integer
    '算法,此月第一天加一个月,然后减去此月第一天的天数
    Dim firstDay
    Dim nextMonthFirstDay
    firstDay = yeearAndMonth + "/1"
    nextMonthFirstDay = DateAdd("m", 1, firstDay)
    getDaysByMonth = DateDiff("d", firstDay, nextMonthFirstDay)
End Function

