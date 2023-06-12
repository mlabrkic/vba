
Option Explicit

' https://github.com/AllenMattson/VBA/tree/master/CREATE%20FILE%20LIST
' Module1.bas

Sub ListFiles()
    Dim Directory As String
    Dim r As Long
    Dim f As String
    Dim FileSize As Double

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select a location containing the files you want to list."
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            Directory = .SelectedItems(1) & "\"
        End If
    End With
    r = 1

'   Insert headers
    Cells.ClearContents
    Cells(r, 1) = "Files in " & Directory
    Cells(r, 2) = "Size"
    Cells(r, 3) = "Date/Time"
    Range("A1:C1").Font.Bold = True

'   Get first file
    f = Dir(Directory, vbReadOnly + vbHidden + vbSystem)
    Do While f <> ""
        r = r + 1
        Cells(r, 1) = f
        'adjust for filesize > 2 gigabytes
        FileSize = FileLen(Directory & f)
        If FileSize < 0 Then FileSize = FileSize + 4294967296#
        Cells(r, 2) = FileSize
        Cells(r, 3) = FileDateTime(Directory & f)
    '   Get next file
        f = Dir
    Loop
End Sub


Sub ListDirs()
    Dim Directory As String
    Dim r As Long
    Dim f As String
    Dim FileSize As Double

    With Application.FileDialog(msoFileDialogFolderPicker)
        .InitialFileName = Application.DefaultFilePath & "\"
        .Title = "Select a location containing the files you want to list."
        .Show
        If .SelectedItems.Count = 0 Then
            Exit Sub
        Else
            Directory = .SelectedItems(1) & "\"
        End If
    End With
    r = 1

'   Insert headers
    Cells.ClearContents
    Cells(r, 1) = "Files in " & Directory
    Cells(r, 2) = "Size"
    Cells(r, 3) = "Date/Time"
    Range("A1:C1").Font.Bold = True

'   Get first file
'    f = Dir(Directory, vbReadOnly + vbHidden + vbSystem)
    f = Dir(Directory, vbDirectory)
    Do While f <> ""
        r = r + 1
        Cells(r, 1) = f
        'adjust for filesize > 2 gigabytes
'        FileSize = FileLen(Directory & f)
'        If FileSize < 0 Then FileSize = FileSize + 4294967296#
'        Cells(r, 2) = FileSize
        Cells(r, 3) = FileDateTime(Directory & f)
    '   Get next file
        f = Dir
    Loop
End Sub

