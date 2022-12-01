'==================
'2020_10_05
'==================
Sub Yes_to_No_sig()
'===========================
' Writen by Michael Ryckin
' Replace all Yes to no+fails signature
' Keyboard Shortcut: Ctrl+t
'===========================
    Columns("Q:Q").Select
    Selection.Replace What:="yes", Replacement:="no + fail signature", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub


Sub No_sig_to_Yes()
'===========================
' Writen by Michael Ryckin
' Replace all no+fails signature to Yes for debug
' Keyboard Shortcut: Ctrl+y
'===========================
    Columns("Q:Q").Select
    Selection.Replace What:="no + fail signature", Replacement:="yes", LookAt _
        :=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
End Sub


Sub Insert_Selection()
    Selection.Insert Shift:=xlDown
End Sub


Sub Select_row()
'===========================
' Writen by Michael Ryckin
' Selects whole row based on user cell selected
'===========================
    Dim row As Integer
    row = ActiveCell.row
    Range(Cells(row, "B"), Cells(row, "P")).Select
    
End Sub


Sub New_Main_Excel_Colors_Preconditions()
'===========================
' Writen by Michael Ryckin
' New Main Excel Colors Precondition
' Ver 3
'===========================
Dim row As Long    'Row variable
'===========================
'Rows Heigh
'===========================
Range("A:A").RowHeight = 12
'===========================
'Columns Width
'===========================
Columns("A").ColumnWidth = 3    'Execute
Columns("B").ColumnWidth = 2    'Loop 2
Columns("C").ColumnWidth = 2    'Loop 1
Columns("D").ColumnWidth = 2    'Device
Columns("E").ColumnWidth = 6    'Sub Device
Columns("F").ColumnWidth = 12   'Address 1
Columns("G").ColumnWidth = 1    'Address 2 for IP10 Use
Columns("H").ColumnWidth = 6    'Slot
Columns("I").ColumnWidth = 3    'State
Columns("J").ColumnWidth = 4    'Command Set,Get,Walk...
Columns("K").ColumnWidth = 40   'Topic
Columns("L").ColumnWidth = 40   'SubTopic
Columns("M").ColumnWidth = 5    'Operator
Columns("N").ColumnWidth = 70   'Value
Columns("O").ColumnWidth = 3    'Protocol
Columns("P").ColumnWidth = 8    'Delay
Columns("Q").ColumnWidth = 23    'Stop on Error
Columns("R").ColumnWidth = 30   'Description
'Columns("S").ColumnWidth = 5
'Columns("T").ColumnWidth = 4
'Columns("U").ColumnWidth = 4
'Columns("V").ColumnWidth = 5
'Columns("W").ColumnWidth = 2
'Columns("X").ColumnWidth = 2
'Columns("Y").ColumnWidth = 2
'Columns("Z").ColumnWidth = 6
'===========================
'Columns Alignment Properties
'===========================
Columns("D").HorizontalAlignment = xlLeft
Columns("E").HorizontalAlignment = xlLeft
Columns("H").HorizontalAlignment = xlLeft
Columns("K").HorizontalAlignment = xlLeft
Columns("Q").HorizontalAlignment = xlLeft
Columns("R").HorizontalAlignment = xlLeft

'===========================
'Attach Colors
'===========================
row = 2
While IsEmpty(Cells(row, "A")) = False  'Continiue until end
        'Color Text to report
        If Cells(row, "K") = "Text to report" Then
            Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(0, 128, 0) 'Green
            Cells(row, "N").Font.Bold = True
            Cells(row, "N").Font.ColorIndex = 2 'White
        'Label Start,End,Start Numeric loop,Save and Reload,Dump to file - Multi labels - Set Row color
        ElseIf Cells(row, "D") = "File_Loop" Then
            Range(Cells(row, "B"), Cells(row, "P")).Interior.ColorIndex = 19
        'Comparison - Set Row color
        ElseIf Cells(row, "K") = "Comparison" Then
            Range(Cells(row, "B"), Cells(row, "P")).Interior.ColorIndex = 44
        'Reference Row, Jump to row - Set Row color
        ElseIf Cells(row, "K") = "Reference line" Then
            Range(Cells(row, "B"), Cells(row, "P")).Interior.ColorIndex = 40
        'N2X or testers - Set Row color
        ElseIf Cells(row, "D") = "TnM" Then
            Range(Cells(row, "B"), Cells(row, "P")).Interior.ColorIndex = 37
        'Run test format Path
        'ElseIf Cells(row, "K") = "Run Test" Then    'Relevant only for Runt Test Rows
        '    Range(Cells(row, "B"), Cells(row, "P")).Interior.Color = RGB(255, 204, 255) 'Range(Cells(row, "N").Address).HorizontalAlignment = xlLeft 'Align test name to Left - but actually it done for whole column so no need this
        'Set - Set Row color
        'ElseIf (Cells(row, "J") = "set" Or Cells(row, "K") = "Set Values" Or Cells(row, "K") = "NG_Set_Values" Or Cells(row, "J") = "add" Or Cells(row, "J") = "edit" Or Cells(row, "J") = "delete") Then
        '    Range(Cells(row, "J"), Cells(row, "K")).Interior.Color = RGB(255, 128, 128) 'Light Red
        'Get - Set Row color
        'ElseIf (Cells(row, "J") = "get" Or Cells(row, "K") = "Get Values") Then
        '    Range(Cells(row, "J"), Cells(row, "K")).Interior.Color = RGB(190, 215, 240) 'Light Blue
        Else
        End If
    'Increment Row counter and check again
    row = row + 1
Wend
'===========================
'Conditional formating
'===========================
    'Clear all conditional formating from current sheet
    Sheets(1).Cells.FormatConditions.Delete
'Attach new rules
    'Green color for cell if Execute = yes
    Columns("A:A").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""yes"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'Gray color for whole row if Execute = No
    'Columns("A:S").Select
    Range(Cells(2, "A"), Cells(row, "S")).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=$A2=""no"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.14996795556505
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
    'Gray color for all rows below "start from" index
    Range(Cells(2, "B"), Cells(row, "P")).Select
    Selection.FormatConditions.Add Type:=xlExpression, Formula1:="=ROW()<=$S$2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249946592608417
    End With
    Selection.FormatConditions(1).StopIfTrue = False
End Sub


Sub Report_Arrangement12()
'===========================
' Writen by Michael Rykin
' Automation Report Arrangement Macro
' Version 14
'===========================

If ActiveWorkbook.Sheets(1).Name = "Result" Then
    'Excel file is appropriate for this macro - Run
    'Start Timer to measure run time
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer

    'Variables
    Dim hyperlinkSheetName As String
    Dim row As Long
    Dim maxRow As Integer
    Dim ws As Worksheet
    Dim btn As Button

    Application.ScreenUpdating = False

    maxRow = Cells(Rows.Count, "A").End(xlUp).row   'Determine Max row

    'Remove unessasary rows from original sheet to reduce final file size (based on automation open case)
    Worksheets("Result").Rows(maxRow + 5 & ":" & Worksheets("Result").Rows.Count).Delete
    
    'Copy Current report sheet for backup
    Worksheets(1).Copy After:=Worksheets(1) 'Backup original Report from Testshell
    ActiveWorkbook.Sheets(1).Activate 'Go back to First sheet

    'Rows Heigh
    Range("A:A").RowHeight = 12

    'Columns Width
    Columns("A").ColumnWidth = 3 'Execute
    Columns("B").ColumnWidth = 3 'Loop 2
    Columns("C").ColumnWidth = 3 'Loop 1
    Columns("D").ColumnWidth = 6 'Device
    Columns("E").ColumnWidth = 8 'Sub Device
    Columns("F").ColumnWidth = 12 'Address 1
    Columns("G").ColumnWidth = 1 'Address 2 for IP10 Use
    Columns("H").ColumnWidth = 6 'Slot
    Columns("I").ColumnWidth = 4 'State
    Columns("J").ColumnWidth = 4 'Command Set,Get,Walk...
    Columns("K").ColumnWidth = 30 'Topic
    Columns("L").ColumnWidth = 30 'SubTopic
    Columns("M").ColumnWidth = 5 'Operator
    Columns("N").ColumnWidth = 35 'Value
    Columns("O").ColumnWidth = 3 'Measured
    Columns("P").ColumnWidth = 8 'Protocol
    Columns("Q").ColumnWidth = 5 'Delay
    Columns("R").ColumnWidth = 23 'Stop on Error
    Columns("S").ColumnWidth = 5 'Status
    Columns("T").ColumnWidth = 4 'Error
    Columns("U").ColumnWidth = 4 'System Log
    Columns("V").AutoFit 'Time Stamp
    Columns("W").ColumnWidth = 10 'Description
    Columns("X").AutoFit 'Duration
    'Columns("Y").ColumnWidth = 2
    'Columns("Z").ColumnWidth = 6 'Duration

    'Columns Alignment Properties
    Columns("D").HorizontalAlignment = xlLeft
    Columns("E").HorizontalAlignment = xlLeft
    Columns("H").HorizontalAlignment = xlLeft
    Columns("K").HorizontalAlignment = xlLeft
    Columns("Q").HorizontalAlignment = xlCenter
    Columns("R").HorizontalAlignment = xlLeft

    'Go through Rows and apply colors
    For row = 2 To maxRow

      'Set another colors
      If Cells(row, "D").value = "TnM" Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 37 'Blue
      ElseIf InStr(1, Cells(row, "K").value, "Run Test") > 0 Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 4 'Green bright
      ElseIf Cells(row, "D").value = "File_Loop" Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 27 'yellow
      ElseIf Cells(row, "K") = "Text to report" Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 10 'Green
      Range(Cells(row, "A"), Cells(row, "R")).Font.Color = vbWhite
      Range(Cells(row, "A"), Cells(row, "R")).Font.Bold = True
      ElseIf Cells(row, "J").value = "set" Then
      Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
      ElseIf Cells(row, "J").value = "edit" Then
      Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
      ElseIf Cells(row, "J").value = "get" Then
      Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 37 'Light Blue
      ElseIf Cells(row, "K") = "Comparison" Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 45 'Orange
      ElseIf Cells(row, "K") = "Reference line" Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 12 'Brown
      ElseIf Cells(row, "K") = "NG_DynamicDelay" Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 39 'Purple
      End If

      'Set row color Red if Fail
      If LCase(Cells(row, "S").value) = "fail" Then
      'Fail
      'Debug.Print ("row failed")
      If InStr(1, Cells(row, "R").value, "no + fail") > 0 Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 3 'Red
      ElseIf InStr(1, Cells(row, "R").value, "yes") > 0 Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 3 'Red
      ElseIf InStr(1, Cells(row, "R").value, "if not") > 0 Then
      Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 3 'Red
      End If
      End If

        'Create links to sheets for all "See walk results in sheet x" Cells
        'Testshell
        If InStr(1, Cells(row, "O").value, "See Walk results") > 0 Then
            hyperlinkSheetName = Mid(Cells(row, "O"), InStr(1, Cells(row, "O"), "WalkResult", 1), 10) & "s" & Right(Cells(row, "O"), (Len(Cells(row, "O")) - (InStr(1, Cells(row, "O"), "WalkResult", 1) + 9)))
            'Debug.Print ("<" & hyperlinkSheetName & ">")
            ActiveCell.Hyperlinks.Add Anchor:=Cells(row, "O"), Address:="", SubAddress:="'" & hyperlinkSheetName & "'" & "!A1"
        'CeraRun
        ElseIf InStr(1, Cells(row, "O").value, "See the measured results") > 0 Then
            hyperlinkSheetName = Mid(Cells(row, "O"), InStr(1, Cells(row, "O"), "'", 1) + 1, InStrRev(Cells(row, "O"), "'") - InStr(1, Cells(row, "O"), "'", 1) - 1)
            'Debug.Print ("<" & hyperlinkSheetName & ">")
            ActiveCell.Hyperlinks.Add Anchor:=Cells(row, "O"), Address:="", SubAddress:="'" & hyperlinkSheetName & "'" & "!A1"
        End If

    Next row

    'Apply Format for Delay column
    Columns("Q").Font.Bold = True 'Bold
    Columns("Q").Font.ColorIndex = 9 'Color = Red

    'With Columns("A:Z").Borders(xlEdgeLeft)
    '.LineStyle = xlContinuous
    '.ColorIndex = 15
    'End With
    'With Columns("A:Z").Borders(xlEdgeTop)
    '.LineStyle = xlContinuous
    '.ColorIndex = 15
    'End With
    'With Columns("A:Z").Borders(xlEdgeBottom)
    '.LineStyle = xlContinuous
    '.ColorIndex = 15
    'End With
    'With Columns("A:Z").Borders(xlEdgeRight)
    '.LineStyle = xlContinuous
    '.ColorIndex = 15
    'End With
    'With Columns("A:Z").Borders(xlInsideVertical)
    '.LineStyle = xlContinuous
    '.ColorIndex = 15
    'End With
    With Columns("A:Z").Borders(xlInsideHorizontal)
    .LineStyle = xlContinuous
    .ColorIndex = 48
    End With

    'Create links from all sheets to Results sheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Index > 2 Then
            'Debug.Print (ws.Name)
            With ws.Buttons.Add(1, 1, 45, 15)
            .OnAction = "ReturnToFirstSheet"
            .Text = "Results"
            End With
        End If
    Next
    ActiveWindow.ScrollColumn = 1   'Scroll to the left
    ActiveWorkbook.Save
    Application.ScreenUpdating = True

    'Stop Timer
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print ("Time took to run: " & SecondsElapsed)
Else
    MsgBox "This file is not appropriate for Report arrangement macro - Abort run", vbCritical
End If
End Sub
Sub ReturnToFirstSheet()
 Sheets("Result").Select
End Sub


Sub FileExist()
'===========================
' Writen by Michael Rykin
' Checks if Test exist on its location - Indicate Exist/Missing near test
'===========================
Dim i As Integer
Dim maxRows As Integer
Dim testPath As String

maxRows = Worksheets(1).Cells(Rows.Count, "A").End(xlUp).row

For i = 2 To maxRows
    'Set path to test based on Strikt or relative access
    If Worksheets(1).Cells(i, "K").value = "Run Test" Then
        testPath = "c:\Program Files\qualisystems\TestShell\TS files\MainExcel\" & _
                    Worksheets(1).Cells(i, "N").value
        Call Result(testPath, i)
    ElseIf Worksheets(1).Cells(i, "K").value = "Run Test from relative path" Then
        testPath = ActiveWorkbook.Path & "\" & Worksheets(1).Cells(i, "N").value
        Call Result(testPath, i)
    End If
Next i
End Sub
Function Result(testPath As String, row As Integer)
'Check file existance and put answer in cell
    If Dir(testPath) <> vbNullString Then
        Worksheets(1).Cells(row, "U").value = "Exist"
        Worksheets(1).Cells(row, "U").Interior.ColorIndex = 50 'Green
    Else
        Worksheets(1).Cells(row, "U").value = "Missing"
        Worksheets(1).Cells(row, "U").Interior.ColorIndex = 3 'Red
    End If
End Function


Sub Clear_Styles()
'===========================
' Writen by Michael Rykin
' Clear all junk styles and leave only default one - help to clean excel file
'===========================
Dim mpStyle As Style
For Each mpStyle In ActiveWorkbook.Styles
If Not mpStyle.BuiltIn Then
mpStyle.Delete
End If
Next mpStyle
End Sub


Sub BreakLinksDataValidation()
'===========================
' Writen by Michael Rykin
' Print all location where corrupted links located to let user fix them
'===========================
Dim row As Integer
Dim col As Integer
maxRow = Cells(Rows.Count, "A").End(xlUp).row
Dim value As String
Dim textPrint As String
textPrint = ""

For row = 2 To maxRow
    For col = 1 To 19
        On Error GoTo skip
        value = Cells(row, col).Validation.Formula1
        If InStr(value, "\") <> 0 Then
            Debug.Print ("Address " & Cells(row, col).Address & " Value: " & value)
            textPrint = textPrint & Cells(row, col).Address & vbCrLf
        End If
skip:
    'Cell have no Data validation (it is by defaul "Any Value")
    'Continiue to next cell
    Resume skip2
skip2:
    Next col
Next row
MsgBox ("Found corrupted links in data validation in following cells : " & vbCrLf & textPrint)
End Sub


Sub Find_and_Replace()
'===========================
' Writen by Michael Rykin
' Use for quick replacement - you can copy segment bellow to perform multiple replace operations at once
'===========================
    Cells.Replace What:="8856", Replacement:="2000", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False



Sub pass_fail_colors_cond_formating()
'===========================
' Writen by Michael Rykin
' Put colors on Selection: Pass = Green, Fail = Red
'===========================
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Pass"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Fail"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub

Sub True_False_colors_cond_formating()
'===========================
' Writen by Michael Rykin
' Put colors on Selection: True = Green, False = Red
'===========================
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""TRUE"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""FALSE"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    
End Sub

Sub CheckXMLSheetToDelete()
'=============================================================================================
' Writen by Michael Rykin
' Use to check which XML sheets possible to delete based on searching for rows which use them.
'=============================================================================================
Dim IP20NCounter As Integer
Dim IP20GCounter As Integer
Dim IP20GXCounter As Integer
Dim IP20CCounter As Integer
Dim IP20Eounter As Integer
Dim IP20Founter As Integer

IP20NCounter = Application.CountIf(Range("N:N"), "*@Genesis_*")
IP20CCounter = Application.CountIf(Range("N:N"), "*@Nexus_*")
IP20GCounter = Application.CountIf(Range("N:N"), "*@IP20G_*")
IP20GXCounter = Application.CountIf(Range("N:N"), "*@IP20GX_*")
IP20Eounter = Application.CountIf(Range("N:N"), "*@E-Band_*")
IP20Founter = Application.CountIf(Range("N:N"), "*@IP20F_*")

MsgBox ("IP20N used " & IP20NCounter & " times" & vbCrLf & _
        "IP20G used " & IP20GCounter & " times" & vbCrLf & _
        "IP20GX used " & IP20GXCounter & " times" & vbCrLf & _
        "IP20C used " & IP20CCounter & " times" & vbCrLf & _
        "IP20E used " & IP20Eounter & " times" & vbCrLf & _
        "IP20F used " & IP20Founter & " times" & vbCrLf)
End Sub

Sub AddAboutSheet()
'=============================================================================================
' Writen by Michael Rykin
' Check if About sheet exist and Add it if not
' Ver 1
'=============================================================================================
Dim SheetExists As Boolean
Dim ws As Worksheet
Dim rng As Range

SheetExists = False
For Each Sheet In Worksheets
    If Sheet.Name = "About" Then
        SheetExists = True
        MsgBox ("Sheet already exists")
        Exit Sub
    End If
Next Sheet
Worksheets.Add(After:=Worksheets(Sheets.Count)).Name = "About"
'Design About Worksheet
Range("A1").value = "Version History"
Range("A2").value = "Version Number"
Range("B2").value = "What changed compared to old version"
Range("A1:B1").Merge
Set rng = Range("A1:B10")
With rng.Borders
    .LineStyle = xlContinuous
    .Weight = xlThin
End With
Range("A1:B2").Interior.Color = 14136213
Set rng = Range("A2:B2")
With rng.Font
    .Size = 14
    .FontStyle = "Calibri"
    .Bold = True
    .ColorIndex = 30
End With
Columns("A").HorizontalAlignment = xlCenter
Columns("B").HorizontalAlignment = xlLeft
Columns("A").AutoFit
Columns("B").ColumnWidth = 120
Range("A1").HorizontalAlignment = xlCenter
Range("A1").Font.Size = 18
Range("A1").Font.Bold = True


End Sub