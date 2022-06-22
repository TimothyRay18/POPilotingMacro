Attribute VB_Name = "BrowseButton"
Function GetFilenameFromPath(ByVal strPath As String) As String
    If Right$(strPath, 1) <> "\" And Len(strPath) > 0 Then
        GetFilenameFromPath = GetFilenameFromPath(Left$(strPath, Len(strPath) - 1)) + Right$(strPath, 1)
    End If
End Function
Sub SelectPO()
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
    dialogBox.InitialFileName = Range("A14").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("B2").Value = dialogBox.SelectedItems(1)
    End If
End Sub

Sub SelectFile()
    Dim dialogBox As FileDialog
    Set dialogBox = Application.FileDialog(msoFileDialogOpen)
    
    dialogBox.AllowMultiSelect = False
    
    dialogBox.Title = "Select a file"
    
    dialogBox.InitialFileName = Range("A14").Value
    
    dialogBox.Filters.Clear
    
    dialogBox.Filters.Add "Excel workbooks", "*.xlsx;*.xlsm;*.xls"
    
    If dialogBox.Show = -1 Then
        ActiveSheet.Range("B6").Value = dialogBox.SelectedItems(1)
        If LCase(Left(GetFilenameFromPath(ActiveSheet.Range("B6").Value), 3)) = "pop" Then
            ActiveSheet.OptionButtons("Option Button 2").Value = True
        ElseIf LCase(Left(GetFilenameFromPath(ActiveSheet.Range("B6").Value), 3)) = "pou" Then
            ActiveSheet.OptionButtons("Option Button 3").Value = True
        End If
    End If
End Sub
