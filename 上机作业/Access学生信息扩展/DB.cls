VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "db"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public s As Student
Private conn As New ADODB.Connection
Public recordSet As New ADODB.recordSet

Const connStr = "Provider=Microsoft.ace.OLEDB.12.0;Data Source=Z:\github\VB6_Self-study_Examinations\上机作业\Access学生信息扩展\student1.accdb;Mode=ReadWrite;Persist Security Info=False"




Sub init()
conn.CursorLocation = adUseClient
If conn.State = 1 Then conn.Close
conn.Open connStr
Dim sqlStr
sqlStr = "select * from student"
Set recordSet = conn.Execute(sqlStr)
End Sub

Sub save()

Dim sqlStr
sqlStr = "insert into student values(% s.xh %, %s.xm %,% s.xb%, %s.zy%,%s.jxj%)"
conn.Execute (sqlStr)
recordSet.Close
sqlStr = "select * from student"
Set recordSet = conn.Execute(sqlStr)
End Sub


Sub update()
Dim sqlStr
sqlStr = "update  student set  = %s.xh %, %s.xm %,% s.xb%, %s.zy%,%s.jxj%)"
conn.Execute (sqlStr)
recordSet.Close
sqlStr = "select * from student"
Set recordSet = conn.Execute(sqlStr)
End Sub








