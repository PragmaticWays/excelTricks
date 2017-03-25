# How to add a DatePicker (Calendar) to Excel Rows

### Step 1. Enable Developer Mode in Excel
- File
- Options
- Customize Ribbon
- Ensure Developer checkbox is enabled

![Enable Developer Mode](https://github.com/Amallard/excelTricks/blob/master/images/developer-enabled.png)


### Step 2. Setup DateTime Control box
- Developer Tab
- Insert
- More Controls

![More Controls](https://github.com/Amallard/excelTricks/blob/master/images/more-controls.png)

- Microsoft Date and Time Picker Control
- Click on random cell to place control box
- Edit or remember the name box of the DateTime Picker, in this case, we will leave it as DatePicker1

![Name Date Picker](https://github.com/Amallard/excelTricks/blob/master/images/date-pick-name.png)

- Right click on DateTime Control box
- DTPicker Object
- Properties

![DTPicker Object > Properties](https://github.com/Amallard/excelTricks/blob/master/images/properties.png)

- Enable CheckBox

![Enable Checkbox](https://github.com/Amallard/excelTricks/blob/master/images/checkbox-enabled.png)

### Step 3. Copy the Visual Basic code

To assign column A as a DatePicker, copy the following code:

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
  
  With Sheet1.DTPicker1
    .Height = 20         
    .Width = 20         
    If Not Intersect(Target, Range("A:A")) Is Nothing Then             
      .Visible = True             
      .Top = Target.Top             
      .Left = Target.Offset(0, 1).Left             
      .LinkedCell = Target.Address         Else             
      .Visible = False         
    End If     
  End With               
      
End Sub
```
- Right click the DateTime Conrol box
- View Code

![View Code](https://github.com/Amallard/excelTricks/blob/master/images/view-code.png)

- Replace all of the existing code with the code from above
- Close the Visual Basic editor

![Paste Code](https://github.com/Amallard/excelTricks/blob/master/images/paste-code.png)

- Deselect Design Mode

![Deselect Design Mode](https://github.com/Amallard/excelTricks/blob/master/images/deselect-design.png)

- Click on any cell to remove DateTime Control box
- Click on any cell in Column A to add a Date

![Pick Date](https://github.com/Amallard/excelTricks/blob/master/images/pick-date.png)

### Congratulations!

#### How to have multiple columns with DatePickers

The above steps only work for single columns, or columns that are right next to each other. If you wanted the DatePicker column to be in column B instead, then you would change the line 

```
If Not Intersect(Target, Range("A:A")) Is Nothing Then
```
to

```
If Not Intersect(Target, Range("B:B")) Is Nothing Then
```

Or if you wanted it from Column E to Column G, then you would change that line to 

```
If Not Intersect(Target, Range("E:G")) Is Nothing Then
```

However, if you need a DatePicker in two or more non-adjacent columns, then you will need a separate DatePicker (each with a separate name in the Name Box) for each non-adjacenet group. For example, let's say we need column A, B, E, F, and H to all be DatePickers. We would need to perform the above steps 3 separate times for 3 separate DatePickers.
- 1 for columns A, B
- 1 for columns E, F
- 1 for column H

The steps will be very similar, but the code will be:

```
Private Sub Worksheet_SelectionChange(ByVal Target As Range)    

  With Sheet1.DTPicker1    
    .Height = 20             
    .Width = 20             
    If Not Intersect(Target, Range("A:B")) Is Nothing Then
      .Visible = True
      .Top = Target.Top
      .Left = Target.Offset(0, 1).Left
      .LinkedCell = Target.Address
    Else
      .Visible = False
    End If
  End With

  With Sheet1.DTPicker2    
    .Height = 20             
    .Width = 20             
    If Not Intersect(Target, Range("E:F")) Is Nothing Then
      .Visible = True
      .Top = Target.Top
      .Left = Target.Offset(0, 1).Left
      .LinkedCell = Target.Address
    Else
      .Visible = False
    End If
  End With

  With Sheet1.DTPicker3    
    .Height = 20             
    .Width = 20             
    If Not Intersect(Target, Range("H:H")) Is Nothing Then
      .Visible = True
      .Top = Target.Top
      .Left = Target.Offset(0, 1).Left
      .LinkedCell = Target.Address
    Else
      .Visible = False
    End If
  End With
  
End Sub
```
