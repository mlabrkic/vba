Attribute VB_Name = "Module1"

Option Explicit

' Early Binding:
' Scripting.Dictionary ==> Needs a reference to the "Microsoft Scripting Runtime" library.

' ----------------------------------------------------------
' Late Binding vs Early Binding:
' https://learn.microsoft.com/en-us/office/vba/api/project.application

' Early Binding
' has better performance because it loads the type library at design time.

' For better performance in VBA and other compiled languages,
' you should use early binding by setting a reference to the "Scripting.Dictionary" type library.

' In the Visual Basic Editor (VBE) for a Excel document, click References on the Tools menu,
' scroll through the Available References list,
' and then choose the "Microsoft Scripting Runtime" library checkbox.

  ' Dim dict1 As Scripting.Dictionary
  ' Set dict1 = New Scripting.Dictionary

' ----------------------------------------------------------
'https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/dictionary-object

' ----------------------------------------------------------
' https://stackoverflow.com/questions/74098950/how-to-watch-a-scripting-dictionary-of-more-than-256-elements-in-excel-vba-whe
Function PrintDictionary( _
        ByVal dict As Scripting.Dictionary, _
        ByVal StartIndex As Long, _
        ByVal EndIndex As Long)

  ' Note that the dictionary indexes are zero-based!
  ' Note that you can loop backwards!

  Dim dStep As Long

  If StartIndex <= EndIndex Then ' ascending
      dStep = 1
      ' Account for out of bounds.
      If StartIndex < 0 Then StartIndex = 0
      If EndIndex > dict.Count - 1 Then EndIndex = dict.Count - 1
  Else ' descending
      dStep = -1
      ' Account for out of bounds.
      If EndIndex < 0 Then EndIndex = 0
      If StartIndex > dict.Count - 1 Then StartIndex = dict.Count - 1
  End If

  Dim Key As Variant
  Dim n As Long

  For n = StartIndex To EndIndex Step dStep
      Debug.Print n, dict.keys(n), dict.items(n) ' if simple datatype
  Next n

End Function


Sub No_01_PrintDictionary()
' https://stackoverflow.com/questions/74098950/how-to-watch-a-scripting-dictionary-of-more-than-256-elements-in-excel-vba-whe

' Because "Debug.Print" ==>
' Choose "Immediate Window": View, Immediate Window

'  Early Binding:
'  Index         Key           Item
'   0             1             1
'   1             2             4
'   2             3             9

  Dim dict As Scripting.Dictionary
  Set dict = New Scripting.Dictionary

  Dim i As Long

  ' Putting Data Into A Dictionary
  For i = 1 To 20
      dict.Add i, i ^ 2
  Next i

  ' This works only if a reference has been created.
  Debug.Print "Early Binding:"
  Debug.Print "Index", "Key", "Item"

  On Error Resume Next ' prevent error if no reference created
'      PrintDictionary dict, 10, 5
'      PrintDictionary dict, 10, 15
      PrintDictionary dict, 0, 6
  On Error GoTo 0

  Set dict = Nothing

End Sub


Sub No_02_dictionary_from_RangeOfCells()
' Create a dictionary from a specified range of cells ( refList1 )
' https://learn.microsoft.com/en-us/office/vba/project/concepts/ole-programmatic-identifiers-late-binding-and-early-binding-project

' mlabrkic

  Dim dict1 As Scripting.Dictionary
  ' Note that the dictionary indexes are zero-based!

  Dim refList1 As Range, refElem1 As Range
  Dim arKeys1 As Variant, arItems1 As Variant

  Dim i As Long, lDict1 As Long

'    "RADNA", "C9:C16" --> Dictionary
  Set dict1 = New Scripting.Dictionary

' ----------------------------------------------------------
  Set refList1 = Sheets("RADNA").Range("C9:C16") 'Range of your strings in the database

  ' Putting Data Into A Dictionary
  With dict1
    For Each refElem1 In refList1
        If Not .Exists(refElem1) And Not IsEmpty(refElem1) Then
            .Add refElem1.Value, refElem1.Offset(0, 1).Value
        End If
    Next refElem1
  End With

' ----------------------------------------------------------
  arKeys1 = dict1.keys       ' keys method ==> Get the keys
  arItems1 = dict1.items     ' items method ==> Get the items
  lDict1 = dict1.Count - 1  ' Count property

  For i = 0 To lDict1
     Debug.Print i, arKeys1(i), arItems1(i)  'Print i, key, item

    ' https://learn.microsoft.com/en-us/office/vba/language/reference/constants-visual-basic-for-applications
    ' use vbNewLine, vbCrLf or vbCR to insert a line break / new paragraph
'    Debug.Print i, arKeys1(i), vbNewLine, vbTab; arItems1(i) 'Print i, key, item
  Next i


  Set refList1 = Nothing
  Set dict1 = Nothing

End Sub

