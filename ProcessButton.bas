Attribute VB_Name = "ProcessButton"
Sub Process()
    Dim file As String
    file = Range("B6").Value
    
    Dim fn As String
    fn = Controller.GetFilenameFromPath(file)
    
    Dim ob As Integer
    ob = Range("A10").Value
    Controller.PO_Template
    
    Controller.PO_Pivot fn, ob
End Sub
