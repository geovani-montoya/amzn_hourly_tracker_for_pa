Attribute VB_Name = "deleteConnections"
Option Explicit

'by Geovani Montoya (DA at KRB1)
Sub delConnect()
    Dim xConnect As Object
    For Each xConnect In ActiveWorkbook.Connections
    If xConnect.Name <> "ThisWorkbookDataModel" Then xConnect.Delete
    Next xConnect
End Sub


