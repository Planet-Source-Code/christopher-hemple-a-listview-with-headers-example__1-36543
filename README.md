<div align="center">

## A ListView With Headers Example


</div>

### Description

This Code will Show you how to use list Headers
 
### More Info
 
U Will need 3 command buttons with the default names, and 1 listview with the default name.

Put these captions for the Command Buttons

Command1 = Add

Command2 = Delete

Command3 = Exit


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Christopher Hemple](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/christopher-hemple.md)
**Level**          |Beginner
**User Rating**    |4.5 (99 globes from 22 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/christopher-hemple-a-listview-with-headers-example__1-36543/archive/master.zip)





### Source Code

```
Private Sub Command1_Click()
'Inputbox Code
Dim name 'Remember name
Dim age 'Remember age
Dim dob 'remember dob
name = InputBox("What is the persons name ?", "Name ?") 'Show a inputbox for the persons name
age = InputBox("What is the personsage ?", "Age ?") 'Show a inputbox for the persons age
dob = InputBox("What is the persons Date Of Birth ?", "Date Of Birth ?") 'Show a inputbox for the persons Date Of Birth
'End Of Inputbox Code
'Adding To List Code
Dim ListObj As ListItem 'Set listObj as a listitem
Set ListObj = ListView1.ListItems.Add(, , name) 'this allways adds to the 1st column , this lines adds the name to the 1st Column
ListObj.SubItems(1) = age 'this allways adds to the second column , this lines adds the age to the 2nd Column
ListObj.SubItems(2) = dob 'this allways adds to the third column , this lines adds the dob to theb 3rd Column
'End Of Adding To List Code
End Sub
Private Sub Command2_Click()
On Error Resume Next 'If Theres a error resume the next line ( the error here would be nothing in the listview or no selected item )
ListView1.ListItems.Remove ListView1.SelectedItem.Index 'Delete The Selected Item
End Sub
Private Sub Command3_Click()
Unload Me 'Exits The Program
End Sub
Private Sub Form_Load()
'This Code is needed
ListView1.View = lvwReport 'Set The Listview1 View So we Can See Our Columns/Headers
ListView1.ColumnHeaders.Add , , "Name" 'Add a column Called Name
ListView1.ColumnHeaders.Add , , "Age" 'Add a column Called Age
ListView1.ColumnHeaders.Add , , "Date Of Birth" 'Add a column Called Age
'End Of Needed Code
End Sub
Private Sub Form_Unload(Cancel As Integer)
'My Code
If MsgBox("If This Code Helped You Please Come Back and Either Vote Or Comment,:). Would You Like To Vote Or Comment Now?", vbYesNo, "Thankx For Using My Code") = vbYes Then GoTo open_url
Exit Sub
open_url:
MsgBox "I have Copyed The Url To The clipboard", vbInformation, "Url"
Clipboard.SetText "http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=36543&lngWId=1"
End Sub
Private Sub ListView1_DblClick()
On Error Resume Next ' resume The Next line On a Error
MsgBox "Name : " + ListView1.SelectedItem.Text + vbCrLf + "Age : " + ListView1.SelectedItem.SubItems(1) + vbCrLf + "Dob : " + ListView1.SelectedItem.SubItems(2) 'Make The Msgbox
End Sub
```

