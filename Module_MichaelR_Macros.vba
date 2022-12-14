'==================
Public Const moduleVersion As String = "V13.6"
Public Const whatIsNew As String = "Add notification about new macro version, improve logging, optimize access to network, clear filter only if active"
'==================

Sub Yes_to_No_sig()
'===========================
' Writen by Michael Ryckin
' Replace all Yes to no+fails signature
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

Sub Transpose_Table()
'===========================
' Writen by Michael Ryckin
' Function transpose current excel area
'===========================
Range("A1").CurrentRegion.Copy
Sheets.Add.Name = "Transpose"
Sheets("Transpose").Range("A1").PasteSpecial Transpose:=True

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
Columns("B").ColumnWidth = 1    'Loop 2
Columns("C").ColumnWidth = 1    'Loop 1
Columns("D").ColumnWidth = 2    'Device
Columns("E").ColumnWidth = 6    'Sub Device
Columns("F").ColumnWidth = 12   'Address 1
Columns("G").ColumnWidth = 1    'Address 2 for IP10 Use
Columns("H").ColumnWidth = 1    'Slot
Columns("I").ColumnWidth = 1    'State
Columns("J").ColumnWidth = 4    'Command Set,Get,Walk...
Columns("K").ColumnWidth = 40   'Topic
Columns("L").ColumnWidth = 2    'SubTopic
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
' Shortcut: ctrl+r
'===========================

If ActiveWorkbook.Sheets(1).Name = "Result" Then
    'Excel file is appropriate for this macro - Run
    'Start Timer to measure run time
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer

    'Create Sheet for macro logs - Must happen before timer print
    Sheets.Add(After:=Sheets("Result")).Name = "Macro Logs"
    ActiveWorkbook.Sheets("Result").Activate 'Go back to First sheet
    
    printDebug StartTime, Timer, "Timer started and added Macro logs sheet"
    
    'Inform user for update
    CheckForLatestMacroVersion
    printDebug StartTime, Timer, "Verified if macro upgrade available"
    
    'Indicate Macro version and what is new
    Cells(2, "Z") = "Macro Version: " & moduleVersion
    Cells(3, "Z") = "What is new? " & whatIsNew
    printDebug StartTime, Timer, "Added current runing version, whats new"
    
    'Variables
    Dim hyperlinkSheetName As String
    Dim row As Long
    Dim maxRow As Long
    Dim ws As Worksheet
    Dim btn As Button
    printDebug StartTime, Timer, "Defined variables"
    
    Application.ScreenUpdating = False

    maxRow = Cells(Rows.Count, "A").End(xlUp).row   'Determine Max row
    printDebug StartTime, Timer, "Calculated max row with content"
    
    'Remove unessasary rows from original sheet to reduce final file size (based on automation open case)
    Worksheets("Result").Rows(maxRow + 5 & ":" & Worksheets("Result").Rows.Count).Delete
    printDebug StartTime, Timer, "Removed unnessasary rows"
    
    'Copy Current report sheet for backup
    Worksheets(1).Copy After:=Worksheets(1) 'Backup original Report from Testshell
    ActiveWorkbook.Sheets("Result").Activate 'Go back to First sheet
    printDebug StartTime, Timer, "Original sheet copied for backup purpose"
    
    'Rows Heigh
    Range("A:A").RowHeight = 12
    Range("1:1").RowHeight = 20

    'Columns Width
    Columns("A").ColumnWidth = 3    'Execute
    Columns("B").ColumnWidth = 0.5  'Loop 2
    Columns("C").ColumnWidth = 0.5  'Loop 1
    Columns("D").ColumnWidth = 6    'Device
    Columns("E").ColumnWidth = 8    'Sub Device
    Columns("F").ColumnWidth = 12   'Address 1
    Columns("G").ColumnWidth = 0.5  'Address 2 for IP10 Use
    Columns("H").ColumnWidth = 0.5  'Slot
    Columns("I").ColumnWidth = 0.5  'State
    Columns("J").ColumnWidth = 2    'Command Set,Get,Walk...
    Columns("K").ColumnWidth = 25   'Topic
    Columns("L").ColumnWidth = 1    'SubTopic
    Columns("M").ColumnWidth = 1    'Operator
    Columns("N").ColumnWidth = 12   'Value
    Columns("O").ColumnWidth = 75   'Measured
    Columns("P").ColumnWidth = 6    'Protocol
    Columns("Q").ColumnWidth = 8    'Delay
    Columns("R").ColumnWidth = 4    'Stop on Error
    Columns("S").ColumnWidth = 5    'Status
    Columns("T").ColumnWidth = 4    'Error
    Columns("U").ColumnWidth = 4    'System Log
    Columns("V").AutoFit            'Time Stamp
    Columns("W").ColumnWidth = 35   'Description
    Columns("X").AutoFit            'Duration

    'Columns Alignment Properties
    Columns("D").HorizontalAlignment = xlLeft
    Columns("E").HorizontalAlignment = xlLeft
    Columns("H").HorizontalAlignment = xlLeft
    Columns("K").HorizontalAlignment = xlLeft
    Columns("Q").HorizontalAlignment = xlCenter
    Columns("R").HorizontalAlignment = xlLeft
    
    printDebug StartTime, Timer, "Formatted rows and columns"
    
    printDebug StartTime, Timer, "Start For Loop and cycle through rows"
    'Cycle through all Rows which hava data in A column and apply colors
    For row = 2 To maxRow
    
        Select Case Cells(row, "K").value
            Case "Run Suite Project"
                Rows(row).RowHeight = 26
                Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(191, 191, 191) 'light grey
            Case "Run Test"
                Rows(row).RowHeight = 26
                Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(191, 191, 191) 'light grey
            Case "Set Variables"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(250, 250, 170) 'yellow
            Case "Text to report"
                Cells(row, "O").Font.Color = vbWhite
                Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 10 'Green
                If Left(Cells(row, "O"), 1) = "#" Then
                    Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(83, 141, 213) 'Light blue Internal loop color
                ElseIf Left(Cells(row, "O"), 3) = ":::" Then
                    Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(191, 191, 191) 'light grey
                ElseIf Left(Cells(row, "O"), 3) = "===" Then
                    Cells(row, "O").wrapText = True
                    Cells(row, "O").EntireRow.AutoFit
                ElseIf Left(Cells(row, "O"), 3) = "---" Then
                    Cells(row, "O").wrapText = True
                    Cells(row, "O").EntireRow.AutoFit
                ElseIf Left(Cells(row, "O"), 3) = "***" Then
                    Cells(row, "O").wrapText = True
                    Cells(row, "O").EntireRow.AutoFit
                Else
                    Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 10 'Green
                    'Cells(row, "O").Font.Bold = True   'Starting 23-5-22 this row make macro stuck for 60 sec
                End If
            Case "SET"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
            Case "ADD"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
            Case "EDIT"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
            Case "GET"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 37 'Light Blue
            Case "Comparison"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 45 'Orange
            Case "Reference line"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 12 'Brown
            Case "NG_DynamicDelay"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(204, 192, 218) 'light purple
            Case "Ping"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(204, 192, 218) 'light purple
        End Select
        
        Select Case Cells(row, "J").value
            Case "set"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
            Case "add"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
            Case "edit"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 22 'Light Red
            Case "get"
                Range(Cells(row, "J"), Cells(row, "K")).Interior.ColorIndex = 37 'Light Blue
        End Select

        Select Case Cells(row, "D").value
            Case "TnM"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 37 'Blue
            Case "File_Loop"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.Color = RGB(250, 250, 170) 'yellow
        End Select

        Select Case Cells(row, "S").value
            Case "FAIL"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 3 'Red
            Case "ERROR"
                Range(Cells(row, "A"), Cells(row, "R")).Interior.ColorIndex = 3 'Red
        End Select

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

    printDebug StartTime, Timer, "For loop end, start color set for fonts"
    'Apply Format for Delay column
    Columns("Q").Font.Bold = True 'Bold 'Starting 23-5-22 this row make macro stuck for 60 sec
    Columns("Q").Font.ColorIndex = 9 'Color = Red
    Columns("N").Font.ColorIndex = 16 'Color = Gray
    Columns("P").Font.ColorIndex = 16 'Color = Gray
    Columns("R").Font.ColorIndex = 16 'Color = Gray
    Columns("D").Font.ColorIndex = 16 'Color = Gray
    Columns("E").Font.ColorIndex = 16 'Color = Gray
    Columns("V").Font.ColorIndex = 16 'Color = Gray
    Columns("W").Font.Color = RGB(79, 129, 189)        'Color = Gray
    Columns("W").Font.Bold = True       'Bold

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
    
    printDebug StartTime, Timer, "Colors and fonts applied"

    'Create links from all sheets to Results sheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.index > 2 Then
            'Debug.Print (ws.Name)
            With ws.Buttons.Add(1, 1, 45, 15)
            .OnAction = "ReturnToFirstSheet"
            .text = "Results"
            End With
        End If
    Next
    ActiveWindow.ScrollColumn = 1   'Scroll to the left
    printDebug StartTime, Timer, "Created Links to results sheets"
    
    'Add button to filter
    Set filterBtn = ActiveSheet.Buttons.Add(Range("O1").Left + 1, 1, 45, Range("O1").Height - 1)
    With filterBtn
      .OnAction = "ReportAutofilterFilterItems"
      .Caption = "Filter"
      .Name = "Filter"
      .Font.Size = 14
      .Font.Bold = True
    End With
    printDebug StartTime, Timer, "Created Filter button"
    
    'Add button to clear filter
    Set clearBtn = ActiveSheet.Buttons.Add(Range("O1").Left + 1 + 45, 1, 45, Range("O1").Height - 1)
    With clearBtn
      .OnAction = "ReportAutofilterClear"
      .Caption = "Clear"
      .Name = "Clear"
      .Font.Size = 14
      .Font.Bold = True
    End With
    printDebug StartTime, Timer, "Created Clear filter button"
    
    'Freeze top row
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    printDebug StartTime, Timer, "Top row freezed"
    
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    printDebug StartTime, Timer, "Workbook saved"
    
    'Stop Timer
    printDebug StartTime, Timer, "Macro finished !!!"
    SecondsElapsed = Round(Timer - StartTime, 2)
    Debug.Print ("Time took to run: " & SecondsElapsed)
    'Indicate Runtime in result
    Cells(4, "Z") = "Macro duration: " & SecondsElapsed
    
Else
    MsgBox "This file is not appropriate for Report arrangement macro - Abort run", vbCritical
End If
End Sub
Sub ReturnToFirstSheet()
 Sheets("Result").Select
End Sub
Function printDebug(start, current, inputText)
    lastEmptyMacroSheetRow = Worksheets("Macro Logs").Cells(Rows.Count, "A").End(xlUp).row + 1
    Debug.Print (Round(current - start, 2) & " : " & inputText)
    Worksheets("Macro Logs").Cells(lastEmptyMacroSheetRow, "A") = Round(current - start, 2)
    Worksheets("Macro Logs").Cells(lastEmptyMacroSheetRow, "B") = inputText
End Function
Sub ReportAutofilterFilterItems()
'===========================
' Writen by Michael Rykin
' Used in CeraRun Result file to filter only relevant rows
'===========================

    If ActiveSheet.AutoFilterMode = True Then
        With Range("$A:$X")
            .AutoFilter Field:=4, Criteria1:=Array("IDU", "Test", "TnM", "System", "Communication", "="), Operator:=xlFilterValues
            .AutoFilter Field:=11, Criteria1:=Array("Text to report", "Run Suite Project", "NG_DynamicDelay", "Free text command", "="), Operator:=xlFilterValues
        End With
    Else
        MsgBox "Auto Filter is turned off - TBD: implement autofilter set if it missing"
    End If

End Sub
Sub ReportAutofilterClear()
'===========================
' Writen by Michael Rykin
' Used in CeraRun Result file to clear autofilter criterias
'===========================
    On Error Resume Next
    ActiveSheet.ShowAllData
End Sub

Sub CheckForLatestMacroVersion()

    Dim notifyUserToUpdate As Boolean
    Dim updateRequired As Boolean
    
    notifyUserToUpdate = CheckIfShowUpdateNotification()
    Debug.Print ("notifyUserToUpdate: " & notifyUserToUpdate)
    
    'First verify if need notify user at all and only then access network to check which version relevant -
    If notifyUserToUpdate Then
        updateRequired = CheckIfRequiredUpdate()    'Access network folder - be aware of connections issues here
        Debug.Print ("Update Required: " & updateRequired)
        
        If updateRequired Then
            Debug.Print ("Update package")
            MsgBox "Your macro version is outdated - consider to update by running updater macro", vbExclamation
        End If
        
    End If
        
End Sub

Function CheckIfRequiredUpdate() As Boolean

    Dim statusFilePath As String
    statusFilePath = "\\emcsrv\R&D\r&d_work_space\Teams\Validation & Verification\Hadar&Meira\Alex_H_team\VBA-Script-Report-Analyzing\CurrentMacroVersion.txt"
    Dim textString As String
    
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    Dim textLine As String
    Dim text As String

    Open statusFilePath For Input As #1
    Do Until EOF(1)
        Line Input #1, textLine
        text = text & textLine
    Loop
    Close #1
    Set fso = Nothing
    
    Debug.Print ("Installed macro: " & moduleVersion)
    Debug.Print ("New macro version ready to install: " & text)
    
    'Check Installed version = last saved version from file
    If text <> moduleVersion Then
        Debug.Print ("Running macro version != Released macro version")
        CheckIfRequiredUpdate = True
    Else
        Debug.Print ("Running macro version= Released macro version")
        CheckIfRequiredUpdate = False
    End If
    
End Function

Function CheckIfShowUpdateNotification() As Boolean
    Dim checkTimeFilePath As String
    checkTimeFilePath = "C:\tmp\reportArrangementMacroLastNotification.txt"
    Dim todayDate As Date
    todayDate = Date
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim oFile As Object
    
    If Dir(checkTimeFilePath) = "" Then
        'Debug.Print ("Last notification file didn't found - Create it, put curent date - Notify user to update")
        CheckIfShowUpdateNotification = True
        Set oFile = fso.CreateTextFile(checkTimeFilePath)
        oFile.WriteLine todayDate
        oFile.Close
        Set oFile = Nothing
    Else
        'Debug.Print ("Last notification file found - Check if today = value from file")
        Dim textLine As String
        Dim text As String
        Open checkTimeFilePath For Input As #1
        Do Until EOF(1)
            Line Input #1, textLine
            text = text & textLine
        Loop
        Close #1
        
        'Check today = saved value
        If text = todayDate Then
            Debug.Print ("Text = Today - User already informed today - Don't bother user to update again today")
            CheckIfShowUpdateNotification = False
        Else
            Debug.Print ("Text != Today - Notify user to update")
            CheckIfShowUpdateNotification = True
            Set oFile = fso.CreateTextFile(checkTimeFilePath)
            oFile.WriteLine todayDate
            oFile.Close
            Set oFile = Nothing
        End If
    End If

    Set fso = Nothing
    
End Function

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
' Shortcut ctrl+e
'===========================
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Pass"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(198, 239, 206)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.Color = RGB(0, 97, 0)
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Fail"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 199, 206)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.Color = RGB(156, 0, 6)
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Error"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(217, 217, 217)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.Color = RGB(166, 166, 166)
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Warning"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = RGB(255, 235, 156)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.Color = RGB(156, 101, 0)
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

Sub TransposeArray()
'=============================================================================================
' Writen by Michael Rykin
' Transpose Selected array
'=============================================================================================
    
Range("A1").Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy

Sheets.Add.Name = "TransposetArray"

Sheets("TransposetArray").Range("A2").PasteSpecial Transpose:=True

End Sub

Sub cellColorYellowLight()
    Selection.Interior.Color = RGB(255, 242, 204)
End Sub
Sub cellColorYellowDark()
    Selection.Interior.Color = RGB(255, 230, 153)
End Sub
Sub cellColorGreenLight()
    Selection.Interior.Color = RGB(226, 239, 218)
End Sub
Sub cellColorGreenDark()
    Selection.Interior.Color = RGB(198, 224, 180)
End Sub
Sub cellColorBlueLight()
    Selection.Interior.Color = RGB(221, 235, 247)
End Sub
Sub cellColorBlueDark()
    Selection.Interior.Color = RGB(189, 215, 238)
End Sub
Sub cellColorRedLight()
    Selection.Interior.Color = RGB(255, 204, 204)
End Sub
Sub cellColorRedDark()
    Selection.Interior.Color = RGB(255, 153, 153)
End Sub