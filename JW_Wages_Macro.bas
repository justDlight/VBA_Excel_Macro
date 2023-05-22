Attribute VB_Name = "JW_Wages_Macro"
Sub JW_Wages_Pivot_Macro()
'
' JW_Macro Macro
'
'Prevent Computer Screen from running
  Application.ScreenUpdating = False


'Clearing the worksheet
Sheets("Wages Pivot Output").Cells.Clear
Sheets("Wages Pivot Output").Cells.ClearFormats
Sheets("Wages Pivot Output").Cells.ClearContents

'Pasting all the values dynamically
Range("X1:AE" & Rows.Count).Select
Selection.Copy
Sheets("Wages Pivot Output").Select
Range("A1").Select
'Paste special so we just paste the values to overcome #REF!
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

'Formatting date column
Range("B:B").NumberFormat = "d/mm/yyyy;@"

' Header_Highlight Macro
' Highlight the header of the Wages Pivot Output file
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
'Looking for (blank) and highlighting row red if found
Range("F2").Select ' Select cell F2, *first line of data*.
Do Until IsEmpty(ActiveCell) ' Set Do loop to stop when an empty cell is reached.
    If ActiveCell.Value = "(blank)" Then 'If the cell value matches (blank) we found a match and highlight that row red up until column H
        With Range("A" & ActiveCell.row & ":H" & ActiveCell.row).Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .Color = 255
            .TintAndShade = 0
            .PatternTintAndShade = 0
        End With
    End If
    ' Step down 1 row from present location.
    ActiveCell.Offset(1, 0).Select
Loop

' Delete all red highlighted rows where column D is 0
Dim i As Long
For i = Sheets("Wages Pivot Output").UsedRange.Rows.Count To 2 Step -1 'loop through all rows in reverse order
    If Sheets("Wages Pivot Output").Range("H" & i).Interior.Color = 255 And Sheets("Wages Pivot Output").Range("D" & i).Value = 0 Then 'check if row is highlighted in red and value in column D is 0
        Sheets("Wages Pivot Output").Rows(i).Delete 'delete the row
    End If
Next i




'Part two Travel Time Macro
'Comments can be found on Salary Pivot Macro same code as that just different columns
Sheets("Wages Pivot updated").Select
Range("AF2:AG2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    
    ' fixingCostCode Macro
'


    Sheets("Wages Pivot updated").Select
    Range("AH2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        
    Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
                'Part three Kilometers of Macro
Sheets("Wages Pivot updated").Select
Range("AI2:AK2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
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
    
    
    
    ' fixingCostCode Macro
'

'
    Sheets("Wages Pivot updated").Select
    Range("AC2").Select
    Range(Selection, Selection.End(xlDown)).Select
    ActiveWindow.SmallScroll Down:=-22
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
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

        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
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
        
        'Part four Travel Allowance Macro
Sheets("Wages Pivot updated").Select
Range("AL2:AM2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
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
    
    
    '
' TA Cost code
'

'
    Sheets("Wages Pivot updated").Select
    Range("AN2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
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
        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
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
        
        'Part five OA1 Macro
Sheets("Wages Pivot updated").Select
Range("AO2:AP2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 97535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    '
' Cost Code macro OA
'

'
    Sheets("Wages Pivot updated").Select
    Range("AQ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 97535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 97535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
        'Part six OA2 Macro
Sheets("Wages Pivot updated").Select
Range("AR2:AS2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 69030
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    '
' Cost code macro OA2
'

'
    Sheets("Wages Pivot updated").Select
    Range("AT2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 69030
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 69030
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
        'Part seven Site Allowance Macro
Sheets("Wages Pivot updated").Select
Range("AU2:AV2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 94353
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        '
' Cost code macro OA2
'

'
    Sheets("Wages Pivot updated").Select
    Range("AT2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 94353
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With '
' Cost code macro
'

'
    Sheets("Wages Pivot updated").Select
    Range("AT2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 93432
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 94353
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
        'Part eight Crib Time Macro
Sheets("Wages Pivot updated").Select
Range("AW2:AX2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 93432
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
        '
' Cost code macro
'

'
    Sheets("Wages Pivot updated").Select
    Range("AT2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    ActiveWindow.SmallScroll Down:=-123
    Range("F1").Select
    'Going to bottom via excel shortcut ctrl+downArrow
    'Offsetting by one to go to next empty cell
    Selection.End(xlDown).Offset(1).Select
    'Paste special so we just paste the values to overcome #REF!
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
    :=False, Transpose:=False
    
    Application.CutCopyMode = False
    
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 784921
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 93432
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
        
        'Part nine Asbestos Macro
Sheets("Wages Pivot updated").Select
Range("AY2:AZ2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("C1").Select
    Selection.End(xlDown).Offset(1).Select
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 784921
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    

        
    Sheets("Wages Pivot updated").Select
        
    Range("X2:Y2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Wages Pivot Output").Select
    Range("A1").Select
    Selection.End(xlDown).Offset(1).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 784921
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    
    

    
    
    
    
    
    
    
'Filling in the cost code.
Dim lastRow3 As Long
Dim rng As Range
Dim cell As Range
Dim closestNumber As Double
Dim diff As Double
Dim closestDiff As Double

' Set the range to Column F
lastRow3 = Sheets("Wages Pivot Output").Cells(Rows.Count, "F").End(xlUp).row
Set rng = Sheets("Wages Pivot Output").Range("F1:F" & lastRow3)

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
Set ws = Sheets("Wages Pivot Output")

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

