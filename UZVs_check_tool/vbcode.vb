Public Sub FindLastRowNumber()
    Dim maxrow As Integer
    Dim rowrange As Range
    Set rowrange = Worksheets("allevent").Range("Q:Q")
    maxrow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Rept("z", 255), rowrange)
    Worksheets("allevent").Cells(2, 11) = maxrow
    For i = 2 To maxrow
     Worksheets("allevent").Cells(i, 14) = i - 1
    Next
End Sub

Sub ExtractUniqueTag()
    Dim arr As New Collection, a
    Dim i As Long
    Dim rowrange As Range
    Set rowrange = Worksheets("allevent").Range("B:B")
    
    On Error Resume Next
    For Each a In rowrange
       arr.Add a, a
    Next
    On Error GoTo 0 ' added to original example by PEH
    
    For i = 1 To arr.Count
       Worksheets("allevent").Cells(i, 17) = arr(i)
    Next
End Sub
Sub CountDeEnergy()
    Dim maxrow As Integer
    Dim numberUniqueTagname As Integer
    Dim arr As New Collection, a
    Dim i As Integer
    Dim j As Integer
    Dim ready2countDO() As Boolean
    Dim readycount As Boolean
    Dim DOcount() As Integer
    Dim rowrange As Range
    Dim UniqueRange As Range
    Set rowrange = ThisWorkbook.Worksheets("alleventinweek").Range("B:B")

    maxrow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Rept("z", 255), rowrange)
    ThisWorkbook.Worksheets("alleventinweek").Cells(1, 10) = maxrow
    ThisWorkbook.Worksheets("alleventinweek").Sort.SortFields.Clear
    ThisWorkbook.Worksheets("alleventinweek").UsedRange.Sort Key1:=ThisWorkbook.Worksheets("alleventinweek").Range("B1"), Key2:=ThisWorkbook.Worksheets("alleventinweek").Range("A1"), Header:=xlYes, _
    Order1:=xlAscending, Order2:=xlAscending
    ' Extract Unique Tagname
    On Error Resume Next
    For Each a In rowrange
        If Mid(a, 3, 4) = "HZV9" Then
        Else
           arr.Add a, a
        End If
           
    Next
    On Error GoTo 0 ' added to original example by PEH
    
    For i = 2 To arr.Count
        Dim celltag As String
        Dim celldescription As String
        celltag = "=LEFT(D" + CStr(i + 1) + ",3)"
        celldescription = "=VLOOKUP(D" + CStr(i + 1) + ",alleventinweek!$B:$E,4,FALSE)"
        
        ThisWorkbook.Worksheets(1).Cells(i + 1, 4) = arr(i)
        ThisWorkbook.Worksheets(1).Cells(i + 1, 3) = celltag
        ThisWorkbook.Worksheets(1).Cells(i + 1, 5) = celldescription
    Next
    
    Set UniqueRange = ThisWorkbook.Worksheets(1).Range("C:C")
    numberUniqueRange = Application.WorksheetFunction.Match(Application.WorksheetFunction.Rept("z", 255), UniqueRange)
    numberUniqueRange = numberUniqueRange - 2
    numberUniqueRange = arr.Count - 1
    
    ' Remove duplicate DIS line
    For i = 2 To maxrow
       If ((ThisWorkbook.Worksheets("alleventinweek").Cells(i, 1) = ThisWorkbook.Worksheets("alleventinweek").Cells(i - 1, 1)) And (Worksheets("alleventinweek").Cells(i, 9) = Worksheets("alleventinweek").Cells(i - 1, 9))) Then
           If (ThisWorkbook.Worksheets("alleventinweek").Cells(i, 7) = "DIS") Then
            ThisWorkbook.Worksheets("alleventinweek").Cells(i - 1, 7) = "DIS_"
           End If
       End If
    Next

    ' Count DeEnergy
    For i = 2 To numberUniqueRange + 1
        ' Write So thu tu
        ThisWorkbook.Worksheets(1).Cells(i + 1, 1) = i - 1
        'ready2countDO(i - 1) = True
        Dim readycountDO As Boolean
        readycountDO = True
        Dim readycountDIS As Boolean
        readycountDIS = True
        Dim DISunknowposibility As Boolean
        Dim DISunknown As Boolean
        DISunknown = False
        DISunknowposibility = True
        Dim countDO As Integer
        Dim countDIS As Integer
        countDIS = 0
        countDO = 0
        For j = 2 To maxrow
            
            If (ThisWorkbook.Worksheets("alleventinweek").Cells(j, 2) = ThisWorkbook.Worksheets(1).Cells(i + 1, 4)) Then ' Check with each unique tagname
                ' Count De enegize
                If (Left(ThisWorkbook.Worksheets("alleventinweek").Cells(j, 7), 2) = "DO") Then ' Check Block column contain DO
                    If (ThisWorkbook.Worksheets("alleventinweek").Cells(j, 8) = "0") And readycountDO Then 'ready2countDO(i - 1) Then ' If DO = 0 and ready to count
                        'DOcount(i - 1) = DOcount(i - 1) + 1
                        countDO = countDO + 1
                        'ready2countDO(i - 1) = False
                        readycountDO = False
                        readycountDIS = True
                        DISunknowposibility = False
                    End If
                    If (ThisWorkbook.Worksheets("alleventinweek").Cells(j, 8) = "1") Then 'ready2countDO(i - 1) Then ' If DO = 1 so ready to count
                        'ready2countDO(i - 1) = True
                        readycountDO = True
                        readycountDIS = False
                        DISunknowposibility = False
                    End If
                End If
                
                ' Count DIS
                If (ThisWorkbook.Worksheets("alleventinweek").Cells(j, 7) = "DIS") Then
                    If (ThisWorkbook.Worksheets("alleventinweek").Cells(j, 9) = "") Then
                        ' Detemine if Unknown discrepancy alarm
                        If DISunknowposibility Then
                            DISunknown = True
                            DISunknowposibility = False
                        End If
                        If (readycountDIS) Then
                            countDIS = countDIS + 1
                        End If
                        readycountDIS = False
                    End If
                    If (ThisWorkbook.Worksheets("alleventinweek").Cells(j, 9) = "OK") Then
                        readycountDIS = True
                    End If
                End If
                
            End If
        Next
        
        If DISunknown Then
            ThisWorkbook.Worksheets(1).Cells(i + 1, 8) = "Limit switch status change without command --> Limit switch error or valve moving due to lost Instrument Air --> need confirm Operation"
        End If
        
        ThisWorkbook.Worksheets(1).Cells(i + 1, 6) = countDO
        ThisWorkbook.Worksheets(1).Cells(i + 1, 7) = countDIS
    Next
    
End Sub
Sub MultiLevelSort()
 
    Worksheets("allevent").Sort.SortFields.Clear
    Worksheets("allevent").UsedRange.Sort Key1:=Range("B1"), Key2:=Range("A1"), Header:=xlYes, _
    Order1:=xlAscending, Order2:=xlAscending
 
End Sub
Sub SortingTagnameDOTime()
 
    Worksheets("allevent").Sort.SortFields.Clear
    Worksheets("allevent").UsedRange.Sort Key1:=Range("B1"), Key2:=Range("A1"), Header:=xlYes, _
    Order1:=xlAscending, Order2:=xlAscending
 
End Sub
Sub DeleteUnnecessaryEvent()
    Dim maxrow As Integer
    Dim i As Integer
    Dim j As Integer
    Dim rowrange As Range
    Set rowrange = Worksheets("allevent").Range("B:B")
    maxrow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Rept("z", 255), rowrange)
    For i = 2 To maxrow
        If ((Worksheets("allevent").Cells(i, 7) = "DIS") Or (Left(Worksheets("allevent").Cells(i, 7), 2) = "DO")) Then
        Else
            For j = 1 To 9
                ' Worksheets("allevent").Cells(i, j).ClearContents
                Worksheets("allevent").Cells(i, j) = Empty
            Next
        End If
    Next
End Sub

Sub GetFilePath()

    
    Set myFile = Application.FileDialog(msoFileDialogOpen)
    With myFile
        .Title = "Choose File"
        .AllowMultiSelect = False
        If .Show <> -1 Then
            Exit Sub
        End If
    FileSelected = .SelectedItems(1)
    End With
    
    ActiveSheet.Range("C1") = FileSelected
    
    Dim maxrow As Integer
    Dim arr As New Collection, a
    Dim i As Integer
    Dim j As Integer
    Dim rowrange As Range
    Set rowrange = ThisWorkbook.Worksheets("UZVDODIScount").Range("A:A")
    maxrow = Application.WorksheetFunction.Match(Application.WorksheetFunction.Rept("z", 255), rowrange)
    ActiveSheet.Range("H1") = maxrow
    For i = 3 To maxrow
        ThisWorkbook.Worksheets("UZVDODIScount").Row(i).ClearContents
    Next
    
    CopySheetFromClosedWB
End Sub
Sub CopySheetMultipleTimes()
Dim n As Integer
Dim i As Integer
On Error Resume Next
 
    n = InputBox("How many copies do you want to make?")
 
    If n > 0 Then
        For i = 1 To n
            ActiveSheet.Copy After:=ActiveWorkbook.Sheets(Worksheets.Count)
        Next
    End If
 
End Sub
Public Sub CopySheetFromClosedWB()
    Dim clswkb As Excel.Workbook
    Dim opdwkb As Excel.Workbook: Set opdwkb = ThisWorkbook
    Dim wks As Excel.Worksheet
    Dim filename As String
    filename = Worksheets("UZVDODIScount").Cells(1, 3).Value
    'filename = """" + filename + """"
    Worksheets("UZVDODIScount").Cells(1, 6) = filename
    
    'Set clswkb = Excel.Workbooks.Open(filename)
    'Set wks = clswkb.Worksheets(2)

    Application.ScreenUpdating = False
    If WorksheetExists("alleventinweek", opdwkb) Then
        Application.DisplayAlerts = False
        opdwkb.Sheets("alleventinweek").Delete
        Application.DisplayAlerts = True
    End If
    'Set closedBook = Workbooks.Open("C:\Users\vinh.kq\Desktop\UZV EVENT_2021-Aug-08-8-30_736.xlsx")
    Set clswkb = Excel.Workbooks.Open(filename)
    clswkb.Sheets(2).Copy After:=opdwkb.Sheets(1)  ' ThisWorkbook.Sheets(1)
    'clswkb.Sheets(1).Copy After:=opdwkb.Sheets(1) '"UZVDODIScount")  ' ThisWorkbook.Sheets(1)
    clswkb.Close SaveChanges:=False
    opdwkb.ActiveSheet.Name = "alleventinweek"
    opdwkb.Sheets(1).Select
    
    'another way
    'clswkb.Sheets(2).Columns("A:I").Copy _
    'opdwkb.Sheets("alleventinweek").Columns("A:I")
    'clswkb.Close SaveChanges:=False
    
    Application.ScreenUpdating = True
End Sub
Function WorksheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

    If wb Is Nothing Then Set wb = ThisWorkbook
    On Error Resume Next
    Set sht = wb.Sheets(shtName)
    On Error GoTo 0
    WorksheetExists = Not sht Is Nothing
End Function
