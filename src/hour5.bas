Attribute VB_Name = "hour5"
Option Explicit
'by Geovani Montoya (DA at KRB1)
Sub hour5()

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = False

    Dim dtDate As Date, dtStartDate As Date, dtEndDate As Date
    Dim iter As Integer
    Dim stIter As String
    Dim building As String

    dtStartDate = Range("B2").value
    dtEndDate = dtStartDate
    
    
    building = Range("B3").value
    

    iter = 5

    For dtDate = dtStartDate To dtEndDate
        iter = iter + 1
        stIter = CStr(iter)
        

        '!!!!!!!!!!!!!!! ToDo: array and loop for database names !!!!!!!!!
        
        Application.Wait (Now + TimeValue("0:00:01"))
        Call import5("ppr", stIter, dtDate, building)
        Application.Wait (Now + TimeValue("0:00:01"))
        Call import5("pid", stIter, dtDate, building)
        Application.Wait (Now + TimeValue("0:00:01"))
        Call import5("frr", stIter, dtDate, building)
        Application.Wait (Now + TimeValue("0:00:01"))
        Call import5("ur", stIter, dtDate, building)
        Application.Wait (Now + TimeValue("0:00:01"))
   
        
    Next dtDate
    
    
    Call delayedSort2
    
    Range("B18:P18").Select
    Selection.ClearContents
    Sheets("Report Generator").Range("D2").Select
    Application.ScreenUpdating = True

End Sub


Sub import5(dataBase As String, refIter As String, dtDate, building)
'''' THIS SUB MAKES SURE THE RIGHT WORKSHEETS ARE PRESENT OR CREATES THEM'''
    Dim Flag
    Dim Count
    Dim i
    Dim wsName
    Dim itemm As Worksheet
    Dim arrWs
    
    Flag = 0
    Count = ActiveWorkbook.Worksheets.Count
    
        For i = 1 To Count
        
            wsName = ActiveWorkbook.Worksheets(i).Name
            If wsName = dataBase + refIter Then Flag = 1
        
            
        Next i
        
            If Flag = 1 Then
                Debug.Print dataBase & refIter & " worksheet exist."
            Else
                Debug.Print dataBase & refIter & " worksheet was created"
                Sheets.Add(After:=Sheets(Sheets.Count)).Name = dataBase + refIter
            End If
            
    Set arrWs = Sheets(Array("ppr6", "pid6", "frr6", "ur6"))

    For Each itemm In arrWs
        Sheets(itemm.Name).UsedRange.ClearContents
    Next itemm
            
    '++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Call websiteDictionaryIntraday(dataBase, refIter, dtDate, building, "11", "12")
    
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    dataBase = vbNullString
    
    Sheets("Report Generator").Select

    Debug.Print "Connecting to import data for " & dtDate & " ..."
    
End Sub



