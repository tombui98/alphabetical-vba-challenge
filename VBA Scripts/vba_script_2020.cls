VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Y2020()

Dim counter As Long
Dim LastRow As Long
Dim DataYear As Integer
Dim Vol As Double
Dim C As Range

'Not to write in the first row
counter = 2
DataYear = 2020
LastRow = Worksheets("Summary").Cells(Rows.Count, 1).End(xlUp).Row + 1
Worksheets("Summary").Range("A2:G" & LastRow).Clear
Worksheets("Summary").Range("A2:G" & LastRow).FormatConditions.Delete
Worksheets("Summary").Range("L9:M11").Clear

'Initial inputs to the table before going to the loop
Sheets("Summary").Cells(counter, 1).Value = Cells(counter, 1).Value
Sheets("Summary").Cells(counter, 2).Value = DataYear
Sheets("Summary").Cells(counter, 3).Value = Cells(counter, 3).Value
Vol = Cells(counter, 7).Value

For i = 3 To 770000

  'to sum the volume
  Vol = Vol + Cells(i, 7).Value

'The below "if" to filter out not necessary data
  If Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i, 1).Value <> "" Then
      Sheets("Summary").Cells(counter, 4).Value = Cells(i, 6).Value
      Sheets("Summary").Cells(counter, 5).Value = Sheets("Summary").Cells(counter, 4).Value - Sheets("Summary").Cells(counter, 3).Value
      Sheets("Summary").Cells(counter, 6).Value = Sheets("Summary").Cells(counter, 5).Value / Sheets("Summary").Cells(counter, 3).Value
      Sheets("Summary").Cells(counter, 7).Value = Vol
      Sheets("Summary").Cells(counter + 1, 1).Value = Cells(i + 1, 1).Value
      Sheets("Summary").Cells(counter + 1, 2).Value = DataYear
      Sheets("Summary").Cells(counter + 1, 3).Value = Cells(i + 1, 3).Value
      Vol = 0
      counter = counter + 1
  End If
Next i

LastRow = Worksheets("Summary").Cells(Rows.Count, 1).End(xlUp).Row + 1

'to change the format
Worksheets("Summary").Range("G2:G" & LastRow).NumberFormat = "#,##0"

With Worksheets("Summary").Range("F2:F" & LastRow)
                                                .NumberFormat = "0.00%"
                                                .FormatConditions.Add(xlCellValue, xlGreater, "=0").Interior.Color = vbGreen
                                                .FormatConditions.Add(xlCellValue, xlLess, "=0").Interior.Color = vbRed
End With

With Worksheets("Summary").Range("E2:E" & LastRow)
                                                .FormatConditions.Add(xlCellValue, xlGreater, "=0").Interior.Color = vbGreen
                                                .FormatConditions.Add(xlCellValue, xlLess, "=0").Interior.Color = vbRed
End With

'The biggest numbers
With Worksheets("Summary")
                            .Range("M9").Value = Application.WorksheetFunction.Max(.Range("F2:F" & LastRow))
                            .Range("M10").Value = Application.WorksheetFunction.Min(.Range("F2:F" & LastRow))
                            .Range("M11").Value = Application.WorksheetFunction.Max(.Range("G2:G" & LastRow))
                            .Range("M9").NumberFormat = "0.00%"
                            .Range("M10").NumberFormat = "0.00%"
                            .Range("M11").NumberFormat = "#,##0"

'The Ticker names for the biggest numbers

            For Each C In .Range("F2:F" & LastRow)
                            If C.Value = .Range("M9").Value Then .Cells(9, 12).Value = .Cells(C.Row, 1).Value
            Next

            For Each C In .Range("F2:F" & LastRow)
                            If C.Value = .Range("M10").Value Then .Cells(10, 12).Value = .Cells(C.Row, 1).Value
            Next

            For Each C In .Range("G2:G" & LastRow)
                            If C.Value = .Range("M11").Value Then .Cells(11, 12).Value = .Cells(C.Row, 1).Value
            Next

End With

End Sub

