Sub DeleteNotBuiltinStyles()
'
' Deletes styles, which are not built-in
' from the currently open Excel file
'

    Dim style As Style
    Dim intRet As Integer

    For Each style In ActiveWorkbook.Styles
        On Error Resume Next

        If Not style.BuiltIn Then
            style.Delete
        End If
    Next style
End Sub
