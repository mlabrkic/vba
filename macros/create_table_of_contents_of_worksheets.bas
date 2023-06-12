
Option Explicit

Sub Create_table_of_contents_of_worksheets()
' https://gist.github.com/walkergv/5623571
' It creates a linked table of contents of worksheets in the excel workbook.

    Dim wbMacro As Workbook
    Dim wbTOC As Workbook   ' TOC -Table Of Contents

    Dim Sheet As Worksheet
    Dim ShTOC As Worksheet   ' TOC -Table Of Contents

    Dim i As Integer

    Dim rngLinkCell         As Range
    Dim strSubAddress       As String
    Dim strDisplayText      As String

    Set wbMacro = Workbooks("create_table_of_contents_of_worksheets.xlsm")
    Set wbTOC = Workbooks("our business Excel file with many sheets.xlsx")

    ' If you want to see changes happen on screen as the code runs
    ' Screen Updating = True, Leaving it false speeds things up
    ' as the screen doens't have to render.
    ' Application.ScreenUpdating = True

    wbTOC.Activate

    ' Add A Worksheet to the Workbook called Table of Contents
    Set ShTOC = Worksheets.Add
    ShTOC.Name = "Table of Contents"

    ' For Each worksheet in the Entire Workbook Loop through create a
    ' Hyperlink to the Worksheet in the Table of Contents
    For Each Sheet In ActiveWorkbook.Worksheets
        Set rngLinkCell = ShTOC.Range("A" & Rows.Count).End(xlUp)
        If ShTOC.Range("A1") = "" Then
            Set rngLinkCell = Worksheets("Table of Contents").Range("A1")
            End If
            If rngLinkCell <> "" Then Set rngLinkCell = rngLinkCell.Offset(1, 0)
            strSubAddress = "'" & Sheet.Name & "'!A1"
            strDisplayText = Sheet.Name

            Worksheets("Table of Contents").Hyperlinks.Add _
            Anchor:=rngLinkCell, _
            Address:="", _
            SubAddress:=strSubAddress, _
            TextToDisplay:=strDisplayText
    Next Sheet

    '
    ' Move the Table of Content to make it the First Worksheet in the Workbook
    With ShTOC   ' TOC -Table Of Contents
         .Move Before:=Worksheets(1)
        For i = 1 To 3
            .Rows(1).EntireRow.Insert
        Next i
 
    ' Let's add some fields above the Table of Content to add some nice titles and style them
        .Columns("A").EntireColumn.Insert
        .Range("B2").Value = "Table of Contents"
        .Range("B2").Font.Size = 14
        .Range("B2").Font.Bold = True
        .Activate
    End With

    ' Turn Off those Damn Gridlines
    ' ActiveWindow.DisplayGridlines = False

    Set wbMacro = Nothing
    Set wbTOC = Nothing

End Sub

