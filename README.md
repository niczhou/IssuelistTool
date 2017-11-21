# IssuelistTool

    On Error GoTo ErrorHandler

    Dim xlapp As Excel.Application
    
    Set xlapp = GetObject(, "Excel.Application")

'    xlapp.Visible = True

    Debug.Print xlapp.Workbooks.Count
    For i = 1 To xlapp.Workbooks.Count
        Debug.Print xlapp.Workbooks(i).Name
    Next
    
    Set xlapp = Nothing
ErrorHandler:
    Set xlapp = CreateObject("Excel.Application")
    Exit Sub
'    xlappWorkbooks.Open FileName:=Application.GetOpenFilename("Excel Files (*.*), Excel Files(*.*)", , "打开新的工作簿")
