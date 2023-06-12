# vba

#### Open Visual Basic Editor

In Excel, select Tools, Macro, Visual Basic Editor, or use the keystroke
Alt+F11.


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

* [devtut - Getting started with excel-vba](https://devtut.github.io/excelvba/getting-started-with-excel-vba.html)


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




#### References

* [devtut - Getting started with excel-vba](https://devtut.github.io/excelvba/getting-started-with-excel-vba.html)
* [devtut - Common Mistakes](https://devtut.github.io/excelvba/common-mistakes.html)

* [ScottSchaen - How To Get Started](https://github.com/ScottSchaen/excel-vba-macros#how-to-get-started)
* [AllenMattson/VBA](https://github.com/AllenMattson/VBA)
* [thesmallman - EXCEL VBA SCRIPTING DICTIONARY](https://www.thesmallman.com/blog/2020/4/24/excel-vba-scripting-dictionary)



