# Unity3D I2Localization Merge
Prefix columns and merge Google Spreadsheet for local use. Requires Excel (tested on version 2016)

Usage
-----

1. Save Google Spreadsheet as Excel
2. Open exported file on Excel
3. Press `Alt`+ `F11`. Goto `Insert > Module` and paste the code bellow and run the `Export` function
4. Export the created sheet as `tab delimited` or `csv`

```vbnet
Sub Export()

    Prefix
    Combine

End Sub

Sub Prefix()
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Range(Range("A2"), Range("A2").End(xlDown)).Select
        RowCount = 1
        For Each rw In Selection.Rows
            Selection.Rows.Cells(RowCount, 1).Value = ws.Name + "/" + Selection.Rows.Cells(RowCount, 1).Value
            RowCount = RowCount + 1
        Next
    Next ws
End Sub

Sub Combine()
    Dim J As Integer
    On Error Resume Next
    Sheets(1).Select
    Worksheets.Add
    Sheets(1).Name = "ExportSheet"
    Sheets(2).Activate
    Range("A1").EntireRow.Select
    Selection.Copy Destination:=Sheets(1).Range("A1")
    For J = 2 To Sheets.Count
        Sheets(J).Activate
        Range("A1").Select
        Selection.CurrentRegion.Select
        Selection.Offset(1, 0).Resize(Selection.Rows.Count - 1).Select
        Selection.Copy Destination:=Sheets(1).Range("A65536").End(xlUp)(2)
    Next
End Sub
```