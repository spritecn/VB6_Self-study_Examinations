Public  Sub Macro()
    dim a,b,c  as Int
    a = 1, b = 1
    c = add(a,b)
    MsgBox "c", "msgTitle"

End Sub


Public Function add(a as Int,b as Int ) As Int
    add = a + b
    '// add declarations
    On Error GoTo catchError
exitFunction:
    Exit Function
catchError:
    '// add error handling
    GoTo exitFunction
End Function