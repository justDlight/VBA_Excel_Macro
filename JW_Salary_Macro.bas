Attribute VB_Name = "JW_Salary_Macro"
Sub JW_Salary_Pivot_Macro()

' JW_Macro Macro
'
'Prevent Computer Screen from running
  Application.ScreenUpdating = False


'Clearing the worksheet
Sheets("Salary Pivot Output").Cells.Clear
Sheets("Salary Pivot Output").Cells.ClearFormats
Sheets("Salary Pivot Output").Cells.ClearContents

'Pasting all the values dynamically
Range("L1:S" & Rows.Count).Select
Selection.Copy
Sheets("Salary Pivot Output").Select
Range("A1").Select
'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

'Formatting date column
Range("B:B").NumberFormat = "d/mm/yyyy;@"

' Header_Highlight Macro
' Highlight the header of the Output2 file
Range("A1:H1").Select
Selection.Font.Bold = True
With Selection.Interior
    .Pattern = xlSolid
    .PatternColorIndex = xlAutomatic
    .ThemeColor = xlThemeColorAccent4
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With

' Matching Macro
'Looking for (blank) in columns D and F, and highlighting row red if found
Dim lastRow As Long
lastRow = Cells(Rows.Count, "F").End(xlUp).row ' Get last row with data in column F

For i = lastRow To 2 Step -1 ' Start from the last row and move up to the second row
    If Cells(i, "F").Value = "(blank)" Or Cells(i, "D").Value = "(blank)" Then 'If the cell value matches (blank) in column F or D, highlight that row red up until column H
        With Range("A" & i & ":H" & i).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
        Rows(i).Delete 'Delete the entire row if it is highlighted red
    End If
Next i



'Part two Kilometer macro Copied data new way with shortcut

'Swap back to Salary Pivot worksheet
Sheets("Salary Pivot").Select

'Select the next range only two columns for kilometers and copy them
Range("T2:U2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    'Back to Worksheet Output 3
    Sheets("Salary Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-21
    'Select cell C1
    Range("C1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
'HighLighting all pasted values
Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With


'Grabbing the Cost code for Kms
    Sheets("Salary Pivot").Select
    Range("V2:V26").Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

'HighLighting all pasted values
Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    'Going back to Salary Pivot worksheet to copy the Emp Code and date columns
    Sheets("Salary Pivot").Select
    
    'Selecting Emp Code and Date column
    'Same as previous code just copying data to Worksheet Output 3
    Range("L2:M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
    'HighLighting all pasted values
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    
    
    'Part Three OA1
    'same as part 2 just with OA1
Sheets("Salary Pivot").Select

Range("W2:X2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 63535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    'Grabbing the Cost code for Kms
    Sheets("Salary Pivot").Select
    Range("V2:V26").Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

'HighLighting all pasted values
Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Sheets("Salary Pivot").Select
    
    Range("L2:M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 63535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    'Part four OA2
    'same as part 3 just with OA2
Sheets("Salary Pivot").Select

Range("Y2:Z2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-21
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 25535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
       'Grabbing the Cost code for Kms
    Sheets("Salary Pivot").Select
    Range("V2:V26").Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

'HighLighting all pasted values
Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
        
    Sheets("Salary Pivot").Select
    
    Range("L2:M2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Salary Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 25535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
  'ADD here
'ADD here
Dim lastRow3 As Long
Dim rng As Range
Dim cell As Range
Dim closestNumber As Double
Dim diff As Double
Dim closestDiff As Double

' Set the range to Column F
lastRow3 = Sheets("Salary Pivot Output").Cells(Rows.Count, "F").End(xlUp).row
Set rng = Sheets("Salary Pivot Output").Range("F1:F" & lastRow3)

' Loop through each cell in the range
For Each cell In rng
    ' Check if the cell is blank or "(blank)"
    If IsNumeric(cell.Value) Or cell.Value = "(blank)" Then
        closestNumber = 0
        closestDiff = 0
        diff = 0
        
        ' Find the closest non-blank number above
        For Each c In rng.Resize(cell.row - 1, 1)
            If IsNumeric(c.Value) Then
                diff = cell.row - c.row
                If closestNumber = 0 Or diff < closestDiff Then
                    closestNumber = c.Value
                    closestDiff = diff
                End If
            End If
        Next c
        
        ' Find the closest non-blank number below
        For Each c In rng.Resize(lastRow3 - cell.row + 1, 1).Offset(cell.row - 1)
            If IsNumeric(c.Value) Then
                diff = c.row - cell.row
                If closestNumber = 0 Or diff < closestDiff Then
                    closestNumber = c.Value
                    closestDiff = diff
                End If
            End If
        Next c
        
        ' Assign the closest number to the blank cell
        cell.Value = closestNumber
    End If
Next cell




    
' Specify the worksheet where you want to delete rows
Set ws = Sheets("Salary Pivot Output")

' Define the last row in the worksheet
lastRow2 = ws.Cells(ws.Rows.Count, "D").End(xlUp).row

' Loop through each row from the last row to the second row
For i2 = lastRow2 To 2 Step -1
    ' Check if the value in column D is 0
    If ws.Cells(i2, "D").Value = 0 Then
        ' Delete the entire row if the condition is met
        ws.Rows(i2).Delete
    End If
Next i2



    'Allow Computer Screen to refresh screen flickering elemination code
Application.ScreenUpdating = True
End Sub
