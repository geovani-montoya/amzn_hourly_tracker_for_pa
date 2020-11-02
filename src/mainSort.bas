Attribute VB_Name = "mainSort"
Option Explicit
'by Geovani Montoya (DA at KRB1)

Sub sort()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False
    
    ''' THIS SUB TRANSFORMS ARRAYS TO COLUMN/CELL FORMAT AND MAPS DATA ONTO REPORT'''
    '!!!!!!!!!!!!!! FIX: horrible way of doing this (e.g. loop through the rows)
    Dim itemm As Worksheet
    Dim arrWs
    
    
    On Error Resume Next
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
        Sheets(itemm.Name).Visible = True
        itemm.Select
        Columns("A:A").Select
        
        Call transform
        itemm.Columns.AutoFit
        
        'Sort for NOW
        If itemm.Name = "ppr1" Then
            Call mapPPRNow(itemm, 10)
        ElseIf itemm.Name = "pid1" Then
            Call mapPIDNow(itemm, 10)
        ElseIf itemm.Name = "frr1" Then
            Call mapFRR(itemm, 10)
        ElseIf itemm.Name = "ur1" Then
            Call mapUR(itemm, Sheets("ppr1"), 10)
            
        'Sort for Hour1
        ElseIf itemm.Name = "ppr2" Then
            Call mapPPR(itemm, 14)
        ElseIf itemm.Name = "pid2" Then
            Call mapPID(itemm, 14, "B8")
        ElseIf itemm.Name = "frr2" Then
            Call mapFRR(itemm, 14)
        ElseIf itemm.Name = "ur2" Then
            Call mapUR(itemm, Sheets("ppr2"), 14)
            
        'Sort for Hour2
        ElseIf itemm.Name = "ppr3" Then
            Call mapPPR(itemm, 15)
        ElseIf itemm.Name = "pid3" Then
            Call mapPID(itemm, 15, "B9")
        ElseIf itemm.Name = "frr3" Then
            Call mapFRR(itemm, 15)
        ElseIf itemm.Name = "ur3" Then
            Call mapUR(itemm, Sheets("ppr3"), 15)
            
        'Sort for Hour3
        ElseIf itemm.Name = "ppr4" Then
            Call mapPPR(itemm, 16)
        ElseIf itemm.Name = "pid4" Then
            Call mapPID(itemm, 16, "B10")
        ElseIf itemm.Name = "frr4" Then
            Call mapFRR(itemm, 16)
        ElseIf itemm.Name = "ur4" Then
            Call mapUR(itemm, Sheets("ppr3"), 16)
            
        'Sort for Hour4
        ElseIf itemm.Name = "ppr5" Then
            Call mapPPR(itemm, 17)
        ElseIf itemm.Name = "pid5" Then
            Call mapPID(itemm, 17, "B11")
        ElseIf itemm.Name = "frr5" Then
            Call mapFRR(itemm, 17)
        ElseIf itemm.Name = "ur5" Then
            Call mapUR(itemm, Sheets("ppr5"), 17)
            
        'Sort for Hour5
        ElseIf itemm.Name = "ppr6" Then
            Call mapPPR(itemm, 18)
        ElseIf itemm.Name = "pid6" Then
            Call mapPID(itemm, 18, "B12")
        ElseIf itemm.Name = "frr6" Then
            Call mapFRR(itemm, 18)
        ElseIf itemm.Name = "ur6" Then
            Call mapUR(itemm, Sheets("ppr6"), 18)
            
        'Sort for Hour6
        ElseIf itemm.Name = "ppr7" Then
            Call mapPPR(itemm, 19)
        ElseIf itemm.Name = "pid7" Then
            Call mapPID(itemm, 19, "B13")
        ElseIf itemm.Name = "frr7" Then
            Call mapFRR(itemm, 19)
        ElseIf itemm.Name = "ur7" Then
            Call mapUR(itemm, Sheets("ppr7"), 19)
            
        'Sort for Hour7
        ElseIf itemm.Name = "ppr8" Then
            Call mapPPR(itemm, 20)
        ElseIf itemm.Name = "pid8" Then
            Call mapPID(itemm, 20, "B14")
        ElseIf itemm.Name = "frr8" Then
            Call mapFRR(itemm, 20)
        ElseIf itemm.Name = "ur8" Then
            Call mapUR(itemm, Sheets("ppr8"), 20)
            
        'Sort for Hour8
        ElseIf itemm.Name = "ppr9" Then
            Call mapPPR(itemm, 21)
        ElseIf itemm.Name = "pid9" Then
            Call mapPID(itemm, 21, "B15")
        ElseIf itemm.Name = "frr9" Then
            Call mapFRR(itemm, 21)
        ElseIf itemm.Name = "ur9" Then
            Call mapUR(itemm, Sheets("ppr9"), 21)
            
        'Sort for Hour9
        ElseIf itemm.Name = "ppr10" Then
            Call mapPPR(itemm, 22)
        ElseIf itemm.Name = "pid10" Then
            Call mapPID(itemm, 22, "B16")
        ElseIf itemm.Name = "frr10" Then
            Call mapFRR(itemm, 22)
        ElseIf itemm.Name = "ur10" Then
            Call mapUR(itemm, Sheets("ppr10"), 22)
            
        'Sort for Hour10
        ElseIf itemm.Name = "ppr11" Then
            Call mapPPR(itemm, 23)
        ElseIf itemm.Name = "pid11" Then
            Call mapPID(itemm, 23, "B17")
        ElseIf itemm.Name = "frr11" Then
            Call mapFRR(itemm, 23)
        ElseIf itemm.Name = "ur11" Then
            Call mapUR(itemm, Sheets("ppr11"), 23)
            
        'Sort for Hour11
        ElseIf itemm.Name = "ppr12" Then
            Call mapPPR(itemm, 24)
        ElseIf itemm.Name = "pid12" Then
            Call mapPID(itemm, 24, "B18")
        ElseIf itemm.Name = "frr12" Then
            Call mapFRR(itemm, 24)
        ElseIf itemm.Name = "ur12" Then
            Call mapUR(itemm, Sheets("ppr12"), 24)
            
        Else
            Debug.Print itemm, "worksheet does not exist"
        End If
    Sheets(itemm.Name).Visible = False
        
    Next itemm
    
    Call clearzero
    Call delConnect
    Application.ScreenUpdating = True
  
Sheets("Report Generator").Select

End Sub


Sub mapPPRNow(ws As Worksheet, j As Integer)
    
    '''''map data onto report
        'Get reveive dock values
        Sheets("Report Generator").Cells(j, 2).value = WorksheetFunction.IfError(Round(ws.Cells(2, 10), 1), " ")
        'Get stow
        Sheets("Report Generator").Cells(j, 4).value = Round(ws.Cells(46, 10), 1)
        'Get IB Total
        Sheets("Report Generator").Cells(j, 5).value = Round(ws.Cells(54, 10), 1)
        'Get Receive Volume
        Sheets("Report Generator").Cells(j, 6).value = Round(ws.Cells(54, 8), 1)
        'Get inbound UPC
        Sheets("Report Generator").Cells(j, 8).value = WorksheetFunction.IfError(Round(ws.Cells(54, 8) / ws.Cells(14, 8), 1), "not found")
        'Get Pick Volume
        Sheets("Report Generator").Cells(j, 11).value = Round(ws.Cells(69, 8), 1)
        'Get TO Dock
        Sheets("Report Generator").Cells(j, 14).value = Round(ws.Cells(71, 10), 1)
        'TO total
        Sheets("Report Generator").Cells(j, 15).value = Round(ws.Cells(74, 10), 1)
        'Get IB CPLH
        Sheets("Report Generator").Cells(j, 7).value = Round(ws.Cells(46, 8) / ws.Cells(180, 9), 1)

        '++++++ planned data +++++++++
        
        'Get planned stow rate
        Sheets("Report Generator").Cells(j - 1, 4).value = Round(ws.Range("K46"), 1)
        
        'Get planned IB Total
        Sheets("Report Generator").Cells(j - 1, 5).value = Round(ws.Range("K54"), 1)
        
        'gets planned pick rate
        'Sheets("Report Generator").Cells(j - 3, 10).value = Round(ws.Range("L69"), 1)
        
        'gets planned TO total
        Sheets("Report Generator").Cells(j - 1, 15).value = Round(ws.Range("K74"), 1)
        
End Sub
        
Sub mapPPR(ws As Worksheet, j As Integer)
    
    '''''map data onto report
        'Get reveive dock values
        Sheets("Report Generator").Cells(j, 2).value = WorksheetFunction.IfError(Round(ws.Cells(2, 10), 1), " ")
        'Get stow
        Sheets("Report Generator").Cells(j, 4).value = Round(ws.Cells(46, 10), 1)
        'Get IB Total
        Sheets("Report Generator").Cells(j, 5).value = Round(ws.Cells(54, 10), 1)
        'Get Receive Volume
        Sheets("Report Generator").Cells(j, 6).value = Round(ws.Cells(54, 8), 1)
        'Get inbound UPC
        Sheets("Report Generator").Cells(j, 8).value = WorksheetFunction.IfError(Round(ws.Cells(54, 8) / ws.Cells(14, 8), 1), "not found")
        'Get Pick Volume
        Sheets("Report Generator").Cells(j, 11).value = Round(ws.Cells(69, 8), 1)
        'Get TO Dock
        Sheets("Report Generator").Cells(j, 14).value = Round(ws.Cells(71, 10), 1)
        'TO total
        Sheets("Report Generator").Cells(j, 15).value = Round(ws.Cells(74, 10), 1)
        'Get IB CPLH
        Sheets("Report Generator").Cells(j, 7).value = Round(ws.Cells(46, 8) / ws.Cells(180, 9), 1)
        
End Sub

Sub mapPIDNow(ws As Worksheet, j As Integer)
    '''' map PID data report '''
    'LP receive
    Sheets("Report Generator").Cells(j, 3).value = Round(ws.Cells(5, 2), 1)
    
    'planned LP receive
    Sheets("Report Generator").Cells(9, 3).value = Round(Sheets("ppr1").Range("K14"), 1)
    
End Sub

Sub mapPID(ws As Worksheet, j As Integer, refCell As String)
    '''' map PID data report '''
    'LP receive
    Sheets("Report Generator").Cells(j, 3).value = Round(ws.Range(refCell), 1)
    
End Sub

Sub mapFRR(ws As Worksheet, j As Integer)
    
    'gets pick rate
    Sheets("Report Generator").Cells(j, 10).value = Round(Application.SumIfs(ws.Columns(17), ws.Columns(16), "Total", ws.Columns(15), "Case") / _
    Application.SumIfs(ws.Columns(11), ws.Columns(16), "Total", ws.Columns(15), "Case"), 1)
    
    'Outbound UPC
    Sheets("Report Generator").Cells(j, 13).value = Round(Application.SumIfs(ws.Columns(17), ws.Columns(15), "EACH", ws.Columns(16), "Total") / _
    Application.SumIfs(ws.Columns(13), ws.Columns(15), "EACH", ws.Columns(16), "Total"), 1)

End Sub


Sub mapUR(ws As Worksheet, ws2 As Worksheet, j As Integer)
    
    'gets OB CLPH from PPR and UR calculation
    Sheets("Report Generator").Cells(j, 12).value = Round(Application.SumIfs(ws.Columns(9), ws.Columns(8), "Total", ws.Columns(7), "Case") / ws2.Range("I181"), 1)

End Sub

Sub transform()

    Selection.TextToColumns Destination:=Range("A1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=True, Space:=False, Other:=False, OtherChar _
        :="#", FieldInfo:=Array(Array(1, 1), Array(2, 1), Array(3, 1), Array(4, 1), Array(5, _
        1), Array(6, 1), Array(7, 1), Array(8, 1), Array(9, 1), Array(10, 1), Array(11, 1), Array(12 _
        , 1), Array(13, 1), Array(14, 1), Array(15, 1), Array(16, 1), Array(17, 1), Array(18, 1), _
        Array(19, 1)), TrailingMinusNumbers:=True

End Sub

Sub clearzero()
    Dim rng As Range
    For Each rng In Sheets("Report Generator").Range("B14:L24") ' substitute your range here
        If rng.value = 0 Then
            rng.value = ""
        End If
    Next
    
    For Each rng In Sheets("Report Generator").Range("B9:F10") ' substitute your range here
        If rng.value = 0 Then
            rng.value = ""
        End If
    Next
End Sub


Sub delayedSort1()
'''THIS SUB HELPS DELAY SUB '''
    Application.OnTime Now() + TimeValue("0:00:20"), "sort"
    sort
    Debug.Print "sorting..."

End Sub

Sub delayedSort2()
'''THIS SUB HELPS DELAY SUB '''
    Application.OnTime Now() + TimeValue("0:00:20"), "sort"
    sort
    Debug.Print "sorting..."

End Sub


Sub mainReset()
'This clears the data to recycle the report
    'Application.ScreenUpdating = False
    Range("B14:P24").Select
    Selection.ClearContents
    'Application.ScreenUpdating = True
    Sheets("Report Generator").Range("D2").Select
    
End Sub

