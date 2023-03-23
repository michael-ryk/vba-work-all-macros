'==================
Public Const moduleVersion  As String = "V15.5"
Public Const whatIsNew      As String = "Add color code for Get Set operations"
'==================


Sub Transpose_Table()
    '===========================
    ' Writen by Michael Ryckin
    ' Function transpose current excel area
    '===========================
    Range("A1").CurrentRegion.Copy
    Sheets.Add.Name = "Transpose"
    Sheets("Transpose").Range("A1").PasteSpecial Transpose:=True

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
        Dim StartTime           As Double
        Dim SecondsElapsed      As Double
        StartTime = Timer
        
        'Constants
        Const heightHighRow = 26
        Const colorLightGrey = "&Hbfbfbf"
        Const colorDarkGrey = "&H808080"
        Const colorYellow = "&Haafafa"
        Const colorLightPurple = "&Headae1"
        Const colorLightBlue = "&Hffcc99"
        Const colorBlue = "&Hd58d53"
        Const colorGreen = "&H008000"
        Const colorLightRed = "&H8080ff"
        Const colorRed = "&H0000ff"
        Const colorDarkRed = "&H000080"
        Const colorBrown = "&H008080"
        Const colorOrange = "&H0099ff"
        Const colorCommentBlue = "&Hbd814f"
        Const colorGetBlue = "&Hfce3cf"
        Const colorGetRed = "&Hddddff"

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
        Dim hyperlinkSheetName  As String
        Dim row                 As Long
        Dim maxRow              As Long
        Dim ws                  As Worksheet
        Dim btn                 As Button
        Dim nColumnData         As String
        Dim currentRange        As Range
        
        printDebug StartTime, Timer, "Defined variables"
        
        Application.ScreenUpdating = False

        maxRow = Cells(Rows.Count, "A").End(xlUp).row   'Determine Max row
        printDebug StartTime, Timer, "Calculated max row with content"
        
        'Remove unessasary rows from original sheet to reduce final file size (based on automation open case)
        Worksheets("Result").Rows(maxRow + 5 & ":" & Worksheets("Result").Rows.Count).Delete
        printDebug StartTime, Timer, "Removed unnecessary rows"
        
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
        Columns("K").ColumnWidth = 16   'Topic
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
            
            Set currentRange = Range("A" & row & ":R" & row)
            nColumnData = Range("N" & row).value
            
            ' Column K Test
            Select Case Cells(row, "K").value
                Case "Run Suite Project"
                    Rows(row).RowHeight = heightHighRow
                    currentRange.Interior.color = colorLightGrey
                Case "Run Test"
                    Rows(row).RowHeight = heightHighRow
                    currentRange.Interior.color = colorLightGrey
                Case "Set Variables"
                    currentRange.Interior.color = colorYellow
                Case "Text to report"
                    Cells(row, "O").Font.color = vbWhite
                    currentRange.Interior.color = colorGreen
                    If Left(Cells(row, "O"), 1) = "#" Then
                        currentRange.Interior.color = colorBlue
                    ElseIf Left(Cells(row, "O"), 3) = ":::" Then
                        currentRange.Interior.color = colorLightGrey
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
                        currentRange.Interior.color = colorGreen
                        'Cells(row, "O").Font.Bold = True   'Starting 23-5-22 this row make macro stuck for 60 sec
                    End If
                Case "Comparison"
                    currentRange.Interior.color = colorOrange
                Case "Reference line"
                    currentRange.Interior.color = colorBrown
                Case "NG_DynamicDelay"
                    currentRange.Interior.color = colorLightPurple
                Case "Ping"
                    currentRange.Interior.color = colorLightPurple
            End Select
            
            ' Column D test
            Select Case Cells(row, "D").value
                Case "TnM"
                    currentRange.Interior.color = colorLightBlue
                Case "File_Loop"
                    currentRange.Interior.color = colorYellow
            End Select
            
            ' Column S test - Failure red color
            Select Case Cells(row, "S").value
                Case "FAIL"
                    currentRange.Interior.color = colorRed
                Case "ERROR"
                    currentRange.Interior.color = colorRed
            End Select
            
            ' Set Get colors for NG REST SNMP commands
            If Cells(row, "E").value = "NG_Rest_SNMP" Then
                If (InStr(nColumnData, "ADD") > 0 Or InStr(nColumnData, "EDIT") > 0 Or InStr(nColumnData, "SET") > 0) Then
                    Range("O" & row).Interior.color = colorGetRed
                ElseIf (InStr(nColumnData, "GET") > 0 Or InStr(nColumnData, "WALK") > 0 Or InStr(nColumnData, "FIND") > 0) Then
                    Range("O" & row).Interior.color = colorGetBlue
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

        printDebug StartTime, Timer, "For loop end, start color set for fonts"
        'Apply Format for Delay column
        Columns("Q").Font.Bold = True 'Bold 'Starting 23-5-22 this row make macro stuck for 60 sec
        Columns("Q").Font.color = colorDarkRed
        Columns("N").Font.color = colorDarkGrey
        Columns("P").Font.color = colorDarkGrey
        Columns("R").Font.color = colorDarkGrey
        Columns("D").Font.color = colorDarkGrey
        Columns("E").Font.color = colorDarkGrey
        Columns("V").Font.color = colorDarkGrey
        Columns("W").Font.color = colorCommentBlue
        Columns("W").Font.Bold = True

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
        
        'Create Filter buttons
        addFilterButton 0, "IDU", "ReportAutoFilterIDU"
        addFilterButton 1, "Filter", "ReportAutofilterFilterItems"
        addFilterButton 2, "Clear", "ReportAutofilterClear"
        addFilterButton 3, "NextFail", "GotoNextFail"
        printDebug StartTime, Timer, "Created Filter buttons"
                
        'Freeze top row
        ActiveWindow.ScrollRow = 1  'Must freeze only when first row seen in screen
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


Function addFilterButton(buttonIndex, buttonName, onClickMacroName)
    '===========================
    ' Writen by Michael Rykin
    ' Create all buttons used for filtering results
    '===========================
    Const ButtonWidth = 70
    Set filterBtn = ActiveSheet.Buttons.Add(Range("O1").Left + 1 + buttonIndex * ButtonWidth, 1, ButtonWidth, Range("O1").Height - 1)
    With filterBtn
    .OnAction = onClickMacroName
    .Caption = buttonName
    .Name = buttonName
    .Font.Size = 14
    .Font.Bold = True
    End With
    
End Function


Function printDebug(start, current, inputText)
    lastEmptyMacroSheetRow = Worksheets("Macro Logs").Cells(Rows.Count, "A").End(xlUp).row + 1
    Debug.Print (Round(current - start, 2) & " : " & inputText)
    Worksheets("Macro Logs").Cells(lastEmptyMacroSheetRow, "A") = Round(current - start, 2)
    Worksheets("Macro Logs").Cells(lastEmptyMacroSheetRow, "B") = inputText
End Function


Sub ReportAutofilterIDU()
    '===========================
    ' Writen by Michael Rykin
    ' Used in CeraRun Result file to filter only IDU rows
    '===========================

    If ActiveSheet.AutoFilterMode = True Then
        Range("$A:$X").AutoFilter Field:=4, Criteria1:=Array("IDU", "System", "Communication"), Operator:=xlFilterValues
    Else
        MsgBox "Auto Filter is turned off - TBD: implement autofilter set if it missing"
    End If

End Sub


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


Sub GotoNextFail()
    '===========================
    ' Writen by Michael Rykin
    ' Used in report arrangement button go to next fail
    '===========================
    Dim FindString As String
    Dim Rng As Range
    Dim ActiveRow As Long
    FindString = "FAIL"
    ActiveRow = ActiveCell.row + 1
    Debug.Print (ActiveRow)
    
    Set Rng = Range("S:S").Find(What:=FindString, _
                    After:=Range("S" & ActiveRow), _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    SearchOrder:=xlByRows, _
                    SearchDirection:=xlNext, _
                    MatchCase:=True)
    If Not Rng Is Nothing Then
        Application.GoTo Sheets("Result").Range("A" & Rng.row - 1), True
        ActiveRow = ActiveCell.row
    Else
        MsgBox "No Failures found in this result file"
    End If

End Sub


Sub CheckForLatestMacroVersion()

    Dim notifyUserToUpdate As Boolean
    Dim updateRequired As Boolean
    
    notifyUserToUpdate = CheckIfShowUpdateNotification()
    
    Debug.Print ("notifyUserToUpdate = " & notifyUserToUpdate)
    
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
    '----------------------------------------------------------------
    ' Return true if new version availabe for user
    '----------------------------------------------------------------

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
    '------------------------------------------------------------------
    ' Return true if user wasn't notified today about potential upgrade
    '------------------------------------------------------------------

    Dim checkTimeFilePath As String
    Dim macroFilesFolder As String
    Dim checkTimeFileName As String
    Dim todayDate As Date
    Dim fso As Object
    Dim oFile As Object
    
    todayDate = Date
    macroFilesFolder = "C:\tmp"
    alternativeMacroFilesFolder = "c:\Users\testing\Documents"
    checkTimeFileName = "reportArrangementMacroLastNotification.txt"
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    If Not fso.FolderExists(macroFilesFolder) Then
        'C:\tmp note exist - use alternative path - currently testing todo: make some common path for all users
        macroFilesFolder = alternativeMacroFilesFolder
    End If
    
    checkTimeFilePath = macroFilesFolder & "\" & checkTimeFileName
    
    If Dir(checkTimeFilePath) = "" Then
        'Debug.Print ("Last notification file didn't found - Create it, put curent date - Notify user to update")
        Debug.Print (checkTimeFilePath)
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
        .color = RGB(198, 239, 206)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.color = RGB(0, 97, 0)
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Fail"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = RGB(255, 199, 206)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.color = RGB(156, 0, 6)
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Error"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = RGB(217, 217, 217)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.color = RGB(166, 166, 166)
    Selection.FormatConditions(1).StopIfTrue = False
    
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""Warning"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = RGB(255, 235, 156)
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).Font.color = RGB(156, 101, 0)
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
        .color = 5287936
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=""FALSE"""
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .color = 255
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
    Selection.Interior.color = RGB(255, 242, 204)
End Sub


Sub cellColorYellowDark()
    Selection.Interior.color = RGB(255, 230, 153)
End Sub


Sub cellColorGreenLight()
    Selection.Interior.color = RGB(226, 239, 218)
End Sub


Sub cellColorGreenDark()
    Selection.Interior.color = RGB(198, 224, 180)
End Sub


Sub cellColorBlueLight()
    Selection.Interior.color = RGB(221, 235, 247)
End Sub


Sub cellColorBlueDark()
    Selection.Interior.color = RGB(189, 215, 238)
End Sub


Sub cellColorRedLight()
    Selection.Interior.color = RGB(255, 204, 204)
End Sub


Sub cellColorRedDark()
    Selection.Interior.color = RGB(255, 153, 153)
End Sub