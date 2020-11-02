Attribute VB_Name = "delete_data"
'by Geovani Montoya (DA at KRB1)

Sub reset_report()
Attribute reset_report.VB_ProcData.VB_Invoke_Func = " \n14"
'This clears the data to recycle the report
    Dim itemm As Worksheet
    Dim arrWs
    
    Application.ScreenUpdating = False
    Range("B14:P24").Select
    Selection.ClearContents
    Range("B9:P10").Select
    Selection.ClearContents
    Application.ScreenUpdating = True
    Sheets("Report Generator").Range("D2").Select
    
    Set arrWs = Sheets(Array("ppr1", "pid1", "frr1", "ur1", _
                             "ppr2", "pid2", "frr2", "ur2", _
                             "ppr3", "pid3", "frr3", "ur3", _
                             "ppr4", "pid4", "frr4", "ur4", _
                             "ppr5", "pid5", "frr5", "ur5", _
                             "ppr6", "pid6", "frr6", "ur6", _
                             "ppr7", "pid7", "frr7", "ur7", _
                             "ppr8", "pid8", "frr8", "ur8", _
                             "ppr9", "pid9", "frr9", "ur9", _
                             "ppr10", "pid10", "frr10", "ur10", _
                             "ppr11", "pid11", "frr11", "ur11", _
                             "ppr12", "pid12", "frr12", "ur12"))

    For Each itemm In arrWs
        Sheets(itemm.Name).UsedRange.ClearContents
    Next itemm
    
    Application.ScreenUpdating = True
    
End Sub


   

