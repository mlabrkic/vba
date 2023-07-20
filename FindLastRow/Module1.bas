Attribute VB_Name = "Module1"

' Tools, Options: Require Variable Declaration:
Option Explicit


Sub No_00_Select_Cell_below_LastRow_Continuous_Data()
'  https://support.microsoft.com/hr-hr/help/291308/how-to-select-cells-ranges-by-using-visual-basic-procedures-in-excel
'  18: How to Select the Blank Cell at Bottom of a Column of Continuous Data

    Dim sMessage As String, sTitle As String, sDefault As String, sMyValue As String

    sMessage = "Enter a value between 1 and last row"    ' Set prompt.
    sTitle = "InputBox - enter the line number!"    ' Set title.
    sDefault = "1"    ' Set default (==> first row)

    ' Display message, title, and default value.
    sMyValue = InputBox(sMessage, sTitle, sDefault)

    ' To select the cell below a range of contiguous cells, use the following example:
    ' ActiveSheet.Range("A1").End(xlDown).Offset(1, 0).Select
    ActiveSheet.Range("A" & sMyValue).End(xlDown).Offset(1, 0).Select

End Sub

'============================================================
'https://github.com/AllenMattson/VBA/blob/master/All%20BAS%20Files/FindLastRow.bas
'============================================================

Function FindLastRow(ByVal Col As Long) As Long
    ' Function Copy_only_uredjaj(i)
    ' https://stackoverflow.com/questions/43631926/lastrow-and-excel-table

    ' Gives you the last cell with data in the specified row
    ' Will not work correctly if the last row is hidden

    ' FindLastRow = Worksheets(1).Cells(Worksheets(1).Rows.Count, Col).End(xlUp).Row

    ' ActiveSheet:
    ' FindLastRow = ActiveSheet.Cells(ActiveSheet.Rows.Count, Col).End(xlUp).Row
    ' FindLastRow = ActiveSheet.Cells(Rows.Count, Col).End(xlUp).Row
    ' FindLastRow = Cells(Rows.Count, Col).End(xlUp).Row

    ' With Worksheets(1)
    '    FindLastRow = .Cells(.Rows.Count, Col).End(xlUp).Row
    ' End With

    With ActiveSheet
        FindLastRow = .Cells(.Rows.Count, Col).End(xlUp).Row
    End With

End Function


Sub No_01_FindLastRow()
    'Sample usage for FindLastRow()

    Dim LastRow As Long
    Dim ColNum As Long

    ColNum = 1
    LastRow = FindLastRow(ColNum)

    MsgBox "The last row in column number " & ColNum & " is " & LastRow

End Sub


Sub No_02_Select_Cell_below_LastRow()
'  https://support.microsoft.com/hr-hr/help/291308/how-to-select-cells-ranges-by-using-visual-basic-procedures-in-excel
'  How to Select the Blank Cell at Bottom of a Column

    ' Sheets("MP_REPORT_UR").Cells(Rows.Count, 2).End(xlUp).Offset(1, 0).Select
    ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Offset(1, 0).Select

End Sub




Function FindLastCol( _
    ByVal Row As Long) As Long

    'Gives you the last cell with data in the specified row
    '  Will not work correctly if the last row is hidden

    With ActiveSheet
        FindLastCol = .Cells(Row, .Columns.Count).End(xlToLeft).Column
    End With

End Function


Sub No_09_FindLastCol()
    'Sample usage for FindLastCol()

    Dim LastCol As Long
    Dim RowNum As Long

    RowNum = 3
    LastCol = FindLastCol(RowNum)

    MsgBox "The last column in row number " & RowNum & " is " & LastCol

End Sub

