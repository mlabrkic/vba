# vba

#### Open Visual Basic Editor

* [devtut - Getting started with excel-vba](https://devtut.github.io/excelvba/getting-started-with-excel-vba.html)

Step 1: Open a Workbook

Step 2 Option **A**: Press Alt + F11 <br>
This is the standard shortcut to open the VBE.

Step 2 Option **B**: Developer Tab --> View Code <br>
First, the Developer Tab must be added to the ribbon. <br>
Go to File -> Options -> Customize Ribbon, then check the box for developer. <br>
Then, go to the developer tab and click "View Code" or "Visual Basic"

Step 2 Option **C**: View tab > Macros > Click Edit button to open an Existing Macro


#### How to get help

``` vba
'Require Variable Declaration
Option Explicit

Sub temp1()

' 1) Write:
inp
  ' press Ctrl+Space together for autocompletion
  ' ==>
  ' InputBox

' 2)
  InputBox(

' 3) Get HELP:
  InputBox()
  ' move the cursor to the InputBox (select that keyword in the code)
  ' hit F1

End Sub
```

In Visual Basic Editor, select Tools, Options<br>
choose: Require Variable Declaration<br>
==>
Option Explicit


#### Debug.Print ( Immediate Window )

* In Excel, Open the Visual Basic Editor (VBE)<br>
* Click View ==> Immediate Window to open the Immediate Window (or ctrl + G).

You should see the Immediate Window at the bottom on VBE.<br>
This window allow you to directly test some VBA code.

So let's start, type in this console :<br>
?Worksheets.<br>
==><br>
?Worksheets.Count

If you have "Debug.Print" in your code, then "Immediate Window" should be enabled, and then the macro should be run from VBE.

How to run Excel macro?<br>
Click on the green "play" arrow (or press F5) in the VBE toolbar to run the program,<br>
while the cursor is inside the Sub procedure.<br>
( Or Click Run ==> "Run Sub/UserForm F5". )


#### How to export source code

In Visual Basic Editor ...

1) Open Module1<br>
(Choose Project Explorer: Select View, Project Explorer)<br>
In Project Explorer, double click on Module1 to open it.

2) Export source code<br>
Select File, Export File<br>
==> Module1.bas


#### HOW TO COPY SOME BAS FILE TO THE NEW EXCEL FILE

1. Open Excel and save new Excel xlsm file

Win-s (search), excel, Enter<br>
Choose "Blank Workbook"

Excel, File, Save As<br>
to Documents

Enter File Name:  test<br>
Excel Workbook xlsm (m- with macros)<br>
Save

2. Download some bas file:

On the github portal, open:
[create_table_of_contents_of_worksheets.bas](https://github.com/mlabrkic/vba/blob/main/macros/create_table_of_contents_of_worksheets.bas)<br>
Click on the "Download raw file". ==> C:\Users\username\Downloads\

3. Copy the contents of bas file to Excel

Open "Visual Basic Editor":  Alt+F11<br>
Click on the Insert,  Module  ==>  Module1<br>
Paste the contents of bas file to Module1, and Save.<br>
Click on the Debug,  Compile  VBAproject.




#### References 1 ([devtut - Excel VBA](https://devtut.github.io/excelvba/))

* [devtut - **Getting started** with excel-vba](https://devtut.github.io/excelvba/getting-started-with-excel-vba.html)
* [Last Used Row or Column in a Worksheet](https://devtut.github.io/excelvba/methods-for-finding-the-last-used-row-or-column-in-a-worksheet.html)
* [Creating a drop-down menu in the Active Worksheet with a Combo Box](https://devtut.github.io/excelvba/creating-a-drop-down-menu-in-the-active-worksheet-with-a-combo-box.html)
* [Early Binding vs Late Binding](https://devtut.github.io/excelvba/binding.html)
* [SQL in Excel VBA - Best Practices](https://devtut.github.io/excelvba/sql-in-excel-vba-best-practices.html)
* [Excel-VBA Optimization](https://devtut.github.io/excelvba/excel-vba-optimization.html)
* [Debugging and Troubleshooting](https://devtut.github.io/excelvba/debugging-and-troubleshooting.html)
* [VBA Best Practices](https://devtut.github.io/excelvba/vba-best-practices.html)
* [Excel VBA Tips and Tricks](https://devtut.github.io/excelvba/excel-vba-tips-and-tricks.html)
* [**Common Mistakes**](https://devtut.github.io/excelvba/common-mistakes.html)


#### References 2 ([devtut - VBA](https://devtut.github.io/vba/))

* [Scripting.Dictionary object](https://devtut.github.io/vba/scripting-dictionary-object.html)
* [CreateObject vs. GetObject](https://devtut.github.io/vba/createobject-vs-getobject.html)
* [Non-Latin Characters](https://devtut.github.io/vba/non-latin-characters.html)
* [VBA Run-Time Errors](https://devtut.github.io/vba/vba-run-time-errors.html)
* [Error Handling](https://devtut.github.io/vba/error-handling.html)


#### References 3

* [ScottSchaen - How To Get Started](https://github.com/ScottSchaen/excel-vba-macros#how-to-get-started)
* [AllenMattson/VBA](https://github.com/AllenMattson/VBA)
* [thesmallman - EXCEL VBA SCRIPTING DICTIONARY](https://www.thesmallman.com/blog/2020/4/24/excel-vba-scripting-dictionary)


#### References 4

* [Microsoft vba](https://learn.microsoft.com/en-us/office/vba/api/overview/) <br>
&nbsp;&nbsp;1. [Microsoft vba - Language reference for VBA](https://learn.microsoft.com/en-us/office/vba/api/overview/language-reference) <br>
&nbsp;&nbsp;&nbsp;&nbsp;1.1. [Microsoft vba - Language reference for VBA - Visual Basic conceptual topics](https://learn.microsoft.com/en-us/office/vba/language/reference/user-interface-help/visual-basic-conceptual-topics) <br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;* [DECLARING VARIABLES](https://learn.microsoft.com/en-us/office/vba/language/concepts/getting-started/declaring-variables) <br>
&nbsp;&nbsp;2. [Excel VBA reference](https://learn.microsoft.com/en-us/office/vba/api/overview/excel) <br>

