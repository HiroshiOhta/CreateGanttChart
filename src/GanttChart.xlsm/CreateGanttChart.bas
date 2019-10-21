Attribute VB_Name = "CreateGanttChart"
Function Plot_Line(Start_X_Coordinate, Y_Coordinate, End_X_Coordinate, Plot_mode)

ActiveSheet.Shapes.AddConnector(msoConnectorStraight, Start_X_Coordinate, Y_Coordinate, End_X_Coordinate, Y_Coordinate).Select
With Selection.ShapeRange.Line
    .Visible = msoTrue
    .BeginArrowheadStyle = msoArrowheadOval
    .BeginArrowheadLength = msoArrowheadShort
    .BeginArrowheadWidth = msoArrowheadNarrow
    .EndArrowheadStyle = msoArrowheadOval
    .EndArrowheadLength = msoArrowheadShort
    .EndArrowheadWidth = msoArrowheadNarrow
End With
    
If Plot_mode = 1 Then
    '# Plot Schedule
    With Selection.ShapeRange.Line
        .DashStyle = msoLineSysDash
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .ForeColor.Brightness = 0.400000006
        .Weight = 1.5
    End With
ElseIf Plot_mode = 2 Then
    '# Plot Schedule (behind schedule)
    With Selection.ShapeRange.Line
        .DashStyle = msoLineSysDash
        .ForeColor.RGB = RGB(204, 0, 0)
        .ForeColor.Brightness = 0.400000006
        .Weight = 1.5
    End With
    '# Plot finish Task
ElseIf Plot_mode = 3 Then
    With Selection.ShapeRange.Line
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .Weight = 2
    End With
    '# Plot Going Task
ElseIf Plot_mode = 4 Then
    With Selection.ShapeRange.Line
        .ForeColor.ObjectThemeColor = msoThemeColorText2
        .Weight = 2
        .EndArrowheadStyle = msoArrowheadNone
    End With
    '# Plot Going Task (behind schedule)
ElseIf Plot_mode = 5 Then
    With Selection.ShapeRange.Line
        .ForeColor.RGB = RGB(204, 0, 0)
        .Weight = 2
        .EndArrowheadStyle = msoArrowheadNone
    End With
End If
    
End Function


Sub Plot_WBS_Line()
'#
'# Plot_WBS_Line
'#
'#
Application.ScreenUpdating = False
ActiveSheet.Outline.ShowLevels RowLevels:=3

Dim Start_WBS_Plan As Byte
Dim End_WBS_Plan As Byte
Dim Check_Flg As Byte
Dim Plot_mode As Byte
Dim Excel_Max_Row As Long
Dim Excel_Max_Column As Long
Dim Check_Data_Row As Long
Dim Check_Data_Column As Long
Dim End_WBS_Row As Long
Dim WBS_Data_Row As Long
Dim WBS_Data_Column As Long
Dim Start_Calendar_Column As Long
Dim Start_Calendar_row As Long
Dim Start_Data_Row As Long
Dim End_Data_Row As Long
Dim End_Calendar_Column As Long
Dim Exec_Data_Column As Long

Dim Reference_Date As Date
Dim Exec_Data As Date
Dim WBS_DATA As Variant

'# delete all shapes
For Each All_Shapes In ActiveSheet.Shapes
   All_Shapes.Delete
Next


'# define WBS Plan and results date area
Start_WBS_Row = 6
Start_WBS_Column = 12
End_WBS_Column = Start_WBS_Column + 4
'Start_Calendar_Column = Start_WBS_Column + 9
Start_Calendar_Column = Start_WBS_Column + 10                               '# TóÒí«â¡Ç…ïπÇπÇƒèCê≥
Reference_Date = Cells(Start_WBS_Row - 1, Start_Calendar_Column)
Excel_Max_Row = 1048576
Excel_Max_Column = 16384
Exec_Data = Date

'#
For Check_Data_Row = Start_WBS_Row To Excel_Max_Row Step 1
   If Cells(Check_Data_Row, 2) <> "" Then
      End_WBS_Row = Check_Data_Row
   Else
      Check_Data_Row = Excel_Max_Row
   End If
Next

For Check_Data_Column = Start_Calendar_Column To Excel_Max_Column Step 1
   If Cells(Start_WBS_Row - 3, Check_Data_Column) <> "" Then
      End_Calendar_Column = Check_Data_Column
      If Cells(Start_WBS_Row - 3, Check_Data_Column) = Exec_Data Then
        Exec_Data_Column = Check_Data_Column
      End If
   Else
      Check_Data_Column = Excel_Max_Column
   End If
Next

'# Get WBS Plan and results date
WBS_DATA = Range(Cells(Start_WBS_Row, Start_WBS_Column), Cells(End_WBS_Row, End_WBS_Column))
Range(Cells(Start_WBS_Row - 3, Start_Calendar_Column), Cells(End_WBS_Row, End_Calendar_Column)).Select
With Selection.Interior
    .Pattern = xlNone
    .TintAndShade = 0
    .PatternTintAndShade = 0
End With


If Reference_Date <= Exec_Data Then
    Range(Cells(Start_WBS_Row - 3, Exec_Data_Column), Cells(End_WBS_Row, Exec_Data_Column)).Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorAccent6
        .TintAndShade = 0.399945066682943
        .PatternTintAndShade = 0
    End With
    With ActiveWindow
        .ScrollColumn = Exec_Data_Column - Weekday(Cells(Start_WBS_Row - 3, Exec_Data_Column), vbMonday) + 1 - 7
        .ScrollRow = 6
    End With
End If


'# Plot Line
For WBS_Data_Row = 1 To UBound(WBS_DATA, 1) Step 1

   For WBS_Data_Column = 1 To UBound(WBS_DATA, 2) Step 3

      Start_data = WBS_DATA(WBS_Data_Row, WBS_Data_Column)
      End_Data = WBS_DATA(WBS_Data_Row, WBS_Data_Column + 1)
      
      If Start_data = "" Or Start_data = "äJén" Or Start_data = "-" Then
        Exit For
      End If
      
      If End_Data = "" Then
         End_Data = Exec_Data
         Plot_mode = 4
         If End_Data > WBS_DATA(WBS_Data_Row, WBS_Data_Column - 2) Then
            Plot_mode = 5
         End If
      End If
      
      If WBS_Data_Column = 1 Then
         If Start_data <= Exec_Data And WBS_DATA(WBS_Data_Row, WBS_Data_Column + 3) = "" Then
            Plot_mode = 2
         Else
            Plot_mode = 1
         End If
      End If
      
      
      Start_Data_Column = DateDiff("d", Reference_Date, Start_data)
      End_Data_Column = DateDiff("d", Reference_Date, End_Data)
      
      Start_X_Coordinate = Cells(WBS_Data_Row + Start_WBS_Row - 1, Start_Calendar_Column + Start_Data_Column).Left
      End_X_Coordinate = Cells(WBS_Data_Row + Start_WBS_Row - 1, Start_Calendar_Column + End_Data_Column + 1).Left
      
      If WBS_Data_Column > 2 Then
         If Plot_mode = 0 Then
           Plot_mode = 3
         End If
         Y_Coordinate = Cells(WBS_Data_Row + Start_WBS_Row, Start_Calendar_Column + Start_Data_Column).Top + (((Cells(WBS_Data_Row + Start_WBS_Row, Start_Calendar_Column + Start_Data_Column).Top - Cells(WBS_Data_Row + Start_WBS_Row + 1, Start_Calendar_Column + Start_Data_Column).Top)) / 3)
         Plot_Line Start_X_Coordinate, Y_Coordinate, End_X_Coordinate, Plot_mode
      Else
         Y_Coordinate = Cells(WBS_Data_Row + Start_WBS_Row, Start_Calendar_Column + Start_Data_Column).Top + (((Cells(WBS_Data_Row + Start_WBS_Row, Start_Calendar_Column + Start_Data_Column).Top - Cells(WBS_Data_Row + Start_WBS_Row + 1, Start_Calendar_Column + Start_Data_Column).Top)) / 3) * 2
         Plot_Line Start_X_Coordinate, Y_Coordinate, End_X_Coordinate, Plot_mode
      End If
      Plot_mode = 0
   Next
Next

Erase WBS_DATA

Cells(1, 1).Select
Application.ScreenUpdating = True

End Sub
