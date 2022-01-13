Attribute VB_Name = "Module2"
Option Explicit

Sub Runall()
Call filterbyCentral
Call filterbyEast
Call filterbyWest
Call filterbySouth
Call EditTxtCopy
Call PrintRegion
Call clearfilter

End Sub


Sub filterbyCentral()
'
' filterbyregion Macro
'

'
    'ActiveWindow.SmallScroll ToRight:=6
    Worksheets("Orders").Select
    Rows("1").Select
    Selection.AutoFilter
    'ActiveWindow.SmallScroll ToRight:=2
    
    
    ActiveSheet.Range("$A$1:$U$9995").AutoFilter Field:=13, Criteria1:= _
        "Central"
    ActiveWindow.SmallScroll ToRight:=7
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("R1:R9995"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll ToRight:=-6
    
    
    'Rows("" & Rows.Count).End(xlUp).Select
    'Selection.CurrentRegion.Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "Central"
    
    Dim i As Long
    For i = 1 To Rows.Count
    Sheet1.Select
    
    
    If Range("M" & i).Value = "Central" Then
    Rows(i).Select
    Selection.Copy

    Worksheets("Central").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste
    Columns("C:C").EntireColumn.AutoFit
    Columns("C:C").Select
    Columns("D:D").EntireColumn.AutoFit
    Application.CutCopyMode = False
    
    End If
    
    
    Next i

    
    Worksheets("Orders").Select
    Application.CutCopyMode = False
    
    Worksheets("Central").Select
    Worksheets("Central").Range("A12:U" & Rows.Count).Select
    Selection.Clear
    Columns("A:U").Select
    Selection.Columns.AutoFit
End Sub
Sub filterbyEast()
'
' filterbyregion Macro
'

'
    'ActiveWindow.SmallScroll ToRight:=6
    Worksheets("Orders").Select
    Rows("1").Select
    Selection.AutoFilter
    'ActiveWindow.SmallScroll ToRight:=2
    
    
    ActiveSheet.Range("$A$1:$U$9995").AutoFilter Field:=13, Criteria1:= _
        "East"
    ActiveWindow.SmallScroll ToRight:=7
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("R1:R9995"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll ToRight:=-6
    
    
    'Rows("" & Rows.Count).End(xlUp).Select
    'Selection.CurrentRegion.Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "East"
    
    Dim i As Long
    For i = 1 To Rows.Count
    Sheet1.Select
    
    
    If Range("M" & i).Value = "East" Then
    Rows(i).Select
    Selection.Copy

    Worksheets("East").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste
    Columns("C:C").EntireColumn.AutoFit
    Columns("C:C").Select
    Columns("D:D").EntireColumn.AutoFit
    Application.CutCopyMode = False
    
    End If
    
    
    Next i

    
    Worksheets("Orders").Select
    Application.CutCopyMode = False
    
    Worksheets("East").Select
    Worksheets("East").Range("A12:U" & Rows.Count).Select
    Selection.Clear
    Columns("A:U").Select
    Selection.Columns.AutoFit
End Sub

Sub filterbyWest()
'
' filterbyregion Macro
'

'
    'ActiveWindow.SmallScroll ToRight:=6
    Worksheets("Orders").Select
    Rows("1").Select
    Selection.AutoFilter
    'ActiveWindow.SmallScroll ToRight:=2
    
    
    ActiveSheet.Range("$A$1:$U$9995").AutoFilter Field:=13, Criteria1:= _
        "West"
    ActiveWindow.SmallScroll ToRight:=7
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("R1:R9995"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll ToRight:=-6
    
    
    'Rows("" & Rows.Count).End(xlUp).Select
    'Selection.CurrentRegion.Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "West"
    
    Dim i As Long
    For i = 1 To Rows.Count
    Sheet1.Select
    
    
    If Range("M" & i).Value = "West" Then
    Rows(i).Select
    Selection.Copy

    Worksheets("West").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste
    Columns("C:C").EntireColumn.AutoFit
    Columns("C:C").Select
    Columns("D:D").EntireColumn.AutoFit
    Application.CutCopyMode = False
    
    End If
    
    
    Next i

    
    Worksheets("Orders").Select
    Application.CutCopyMode = False
    
    Worksheets("West").Select
    Worksheets("West").Range("A12:U" & Rows.Count).Select
    Selection.Clear
    Columns("A:U").Select
    Selection.Columns.AutoFit
End Sub
Sub filterbySouth()
'
' filterbyregion Macro
'

'
    'ActiveWindow.SmallScroll ToRight:=6
    Worksheets("Orders").Select
    Rows("1").Select
    Selection.AutoFilter
    'ActiveWindow.SmallScroll ToRight:=2
    
    
    ActiveSheet.Range("$A$1:$U$9995").AutoFilter Field:=13, Criteria1:= _
        "South"
    ActiveWindow.SmallScroll ToRight:=7
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort.SortFields.Add Key:=Range _
        ("R1:R9995"), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Orders").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll ToRight:=-6
    
    
    'Rows("" & Rows.Count).End(xlUp).Select
    'Selection.CurrentRegion.Select
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "South"
    
    Dim i As Long
    For i = 1 To Rows.Count
    Sheet1.Select
    
    
    If Range("M" & i).Value = "South" Then
    Rows(i).Select
    Selection.Copy

    Worksheets("South").Select
    Range("A" & Rows.Count).End(xlUp).Select
    Selection.Offset(1, 0).Select
    ActiveSheet.Paste
    Columns("C:C").EntireColumn.AutoFit
    Columns("C:C").Select
    Columns("D:D").EntireColumn.AutoFit
    Application.CutCopyMode = False
    
    End If
    
    
    Next i

    
    Worksheets("Orders").Select
    Application.CutCopyMode = False
    
    Worksheets("South").Select
    Worksheets("South").Range("A12:U" & Rows.Count).Select
    Selection.Clear
    Columns("A:U").Select
    Selection.Columns.AutoFit
    
End Sub

Sub EditTxtCopy()

'Top 10 sales in each region

Worksheets.Add After:=Worksheets(5)
Worksheets(6).Name = "Top10 salesperregion"
Range("A1").Value = "Top 10 sales per region"
   With ActiveCell.Font
         .Bold = "Yes"
         .Size = 20
         .Name = "Arial"
   End With
   Selection.Interior.ThemeColor = xlThemeColorDark2
   
Dim x As Integer
x = 2
Do While x < 6

Worksheets("Orders").Select
Rows(1).Copy

Worksheets(x).Select
Range("A1").Select
ActiveSheet.Paste
Application.CutCopyMode = False

Range("A1:U1").Select
With Selection.Font
     .Bold = "Yes"
     .Size = 14
     .Name = "Arial"
End With
Selection.Interior.ThemeColor = xlThemeColorDark2

Range("A1").CurrentRegion.Select
Selection.Copy

Worksheets(6).Select
Range("B2000").End(xlUp).Select
ActiveCell.Offset(2, 0).Select
ActiveCell.Value = Worksheets(x).Name
ActiveCell.Offset(1, 0).Select
ActiveSheet.Paste
Application.CutCopyMode = False

Cells.Find(What:="" & Worksheets(x).Name, After:=Range("A1"), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False).Activate
        
        
        With Selection.Font
            .Name = "Arial"
            .Size = 14
            .Bold = True
        
        End With


x = x + 1
Loop

Columns("C:V").Select
Selection.Columns.AutoFit



      





End Sub
Sub PrintRegion()

Dim message As String
Dim inputanswer, region As Integer

message = "Enter the region of your choice to print" & vbCrLf & _
      "1 - Central " & vbCrLf & _
      "2 - East " & vbCrLf & _
      "3 - West " & vbCrLf & _
      "4 - South" & vbCrLf & _
      "5 - Top10sales "

region = InputBox(message, "Region", "Enter 1, 2, 3, 4 or 5")

Select Case region
    Case 1
        Worksheets("Central").PrintOut
    Case 2
        Worksheets("East").PrintOut
    Case 3
        Worksheets("West").PrintOut
    Case 4
        Worksheets("South").PrintOut
    Case 5
        Worksheets("Top10 salesperregion").PrintOut
    
    Case Else
        inputanswer = MsgBox("You didn't type a number between 1 and 4. Try Again?", vbYesNo)
        If inputanswer = vbYes Then
        Call PrintRegion
        End If
    End Select
      



End Sub

Sub clearfilter()
Attribute clearfilter.VB_ProcData.VB_Invoke_Func = " \n14"
'
' clearfilter Macro
'

'   Worksheets("Orders").Select
    ActiveWindow.SmallScroll ToRight:=4
    ActiveSheet.Range("$A$1:$U$9995").AutoFilter Field:=13
    Selection.AutoFilter
End Sub

