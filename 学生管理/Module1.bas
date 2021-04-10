Attribute VB_Name = "Module1"
Public dbFile As String

Public Type student
    name As String * 8
    id As Integer
    age As Integer
End Type



Public Function student2str(stu As student) As String
    student2str = Trim(stu.name) & " - " & stu.id & " - " & stu.age
End Function

Public Sub addStudent2File(s As student)
    Dim f
    f = FreeFile
    Open dbFile For Random As #f Len = Len(s)
    Put #f, , s
    Close (f)
    
End Sub


