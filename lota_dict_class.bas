Attribute VB_Name = "Samplesuperstore"
Option Explicit
Sub Runall2()
Dim dict As Dictionary


Set dict = Readdata()

Call Writedata(dict)




End Sub

Function Readdata() As Dictionary

Dim dict As New Dictionary
Dim rg As Range, x As Long
Dim Category As String, Sales As Currency, Profit As Currency


For x = 2 To Sheet1.Rows.Count

Category = Sheet1.Cells(x, 15).Value
Sales = Sheet1.Cells(x, 18).Value
Profit = Sheet1.Cells(x, 21).Value

Dim o As New Superstoreclass

If dict.exists(Category) Then
   Set o = dict(Category)
   
Else
   Set o = New Superstoreclass
   dict.Add (Category), o
End If
   
o.Sales = o.Sales + Sales
o.Profit = o.Profit + Profit
   
Next x


Set Readdata = dict

End Function


Sub Writedata(dict As Dictionary)
Dim owrite As Superstoreclass, rg As Range, row As Long
Dim item As Variant


Worksheets.Add after:=Worksheets("Orders")
Worksheets(2).Name = "Category&Sales&Profit"

Worksheets("Category&Sales&Profit").Select
Set rg = Range("A1").CurrentRegion
row = 1
For Each item In dict

Set owrite = dict(item)

rg.Cells(row, 1).Value = item
rg.Cells(row, 2).Value = owrite.Sales
rg.Cells(row, 3).Value = owrite.Profit

Debug.Print item, owrite.Sales, owrite.Profit


row = row + 1

Next item

Worksheets("Category&Sales&Profit").Select
Rows(1).Insert shift:=xlDown


Range("A1").Value = "Category"
Range("B1").Value = "Sales"
Range("C1").Value = "Profit"

With Range("A1:C1").Font
     .Bold = True
     .Size = 16
     .Color = vbBlue
     .Name = "Verdana"
End With

Range("A1:C1").Interior.ThemeColor = xlThemeColorDark2
Columns("A:C").EntireColumn.AutoFit



End Sub
