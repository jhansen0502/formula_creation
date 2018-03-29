Attribute VB_Name = "Module2"
Sub reset_all()
Dim this_sheet As String
Dim new_sheet As String
Dim source_sheet As String
this_sheet = Sheets(1).Name
source_sheet = "Source Data"

    Sheets(source_sheet).Delete
    
End Sub
