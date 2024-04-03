Sub unhideAllNames()
'
' Unhide all cell and range names in the currently open Excel file
'
' Intention:
' Excel files with several years' history tend to cumulate range and cell names.
' Such files could contain several thousands of such names/ranges,
' which leads to high processor consumption. 
  
    For Each rangeName In ActiveWorkbook.Names
        rangeName.Visible = True
    Next
End Sub
