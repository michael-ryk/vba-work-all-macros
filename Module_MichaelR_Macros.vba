Option Explicit

'==================
Public Const moduleVersion  As String = "V16.7"
Public Const whatIsNew      As String = "Detect columns by name"
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

    ' Test if abort macro execution
    If ActiveSheet.Name <> "Result" Then
        MsgBox "Macro is not applicable for current sheet - Abort", vbCritical
        Exit Sub
    End If
    
    If macroAlreadyApplied() Then
        MsgBox "Macro already applied - Abort", vbCritical
        Exit Sub
    End If

    ' Start Timer to measure run time
    Dim dStartTime           As Double
    Dim dSecondsElapsed      As Double
    dStartTime = Timer
    
    ' Constants
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
    Const colorBlack = "&H0d0d0d"

    ' Create Sheet for macro logs - Must happen before timer print
    Sheets.Add(After:=Sheets("Result")).Name = "Macro Logs"
    ActiveWorkbook.Sheets("Result").Activate 'Go back to First sheet
    
    printDebug dStartTime, Timer, "Timer started and added Macro logs sheet"
    
    ' Inform user for update
    CheckForLatestMacroVersion
    printDebug dStartTime, Timer, "Verified if macro upgrade available"
    
    ' Indicate Macro version and what is new
    Cells(2, "Z") = "Macro Version: " & moduleVersion
    Cells(3, "Z") = "What is new? " & whatIsNew
    printDebug dStartTime, Timer, "Added current runing version, whats new"
    
    ' Variables
    Dim sHyperlinkSheetName  As String
    Dim lRow                 As Long
    Dim lMaxRow              As Long
    
    ' Declare All Columns variables
    Dim colStatus            As Integer
    Dim colSystemlog         As Integer
    Dim colError             As Integer
    Dim colTimestamp         As Integer
    Dim colDescription       As Integer
    Dim colDuration          As Integer
    
    printDebug dStartTime, Timer, "Defined variables"
    
    Application.ScreenUpdating = False

    lMaxRow = Cells(Rows.Count, "A").End(xlUp).row   'Determine Max row
    printDebug dStartTime, Timer, "Calculated max row with content"
    
    ' Remove unessasary rows from original sheet to reduce final file size (based on automation open case)
    Worksheets("Result").Rows(lMaxRow + 5 & ":" & Worksheets("Result").Rows.Count).Delete
    printDebug dStartTime, Timer, "Removed unnecessary rows"
    
    ' Copy Current report sheet for backup
    Worksheets(1).Copy After:=Worksheets(1) 'Backup original Report from Testshell
    ActiveWorkbook.Sheets("Result").Activate 'Go back to First sheet
    printDebug dStartTime, Timer, "Original sheet copied for backup purpose"
    
    ' Rows Heigh
    Range("A:A").RowHeight = 12
    Range("1:1").RowHeight = 20

    ' Columns Width
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
    Columns("S").ColumnWidth = 5    'Run Condition
    
    ' Detect Columns indexes by header name
    colStatus = WorksheetFunction.Match("Status", Range("1:1"), 0)
    colError = WorksheetFunction.Match("Error", Range("1:1"), 0)
    colSystemlog = WorksheetFunction.Match("SystemLog", Range("1:1"), 0)
    colTimestamp = WorksheetFunction.Match("TimeStamp", Range("1:1"), 0)
    colDescription = WorksheetFunction.Match("UserDescription", Range("1:1"), 0)
    colDuration = WorksheetFunction.Match("Duration", Range("1:1"), 0)
    
    ' Set Column width
    Columns(colStatus).ColumnWidth = 5
    Columns(colError).ColumnWidth = 4
    Columns(colSystemlog).ColumnWidth = 4
    Columns(colTimestamp).AutoFit
    Columns(colDescription).ColumnWidth = 47
    Columns(colDuration).ColumnWidth = 12
    
    ' Columns Alignment Properties
    Columns("D").HorizontalAlignment = xlLeft
    Columns("E").HorizontalAlignment = xlLeft
    Columns("H").HorizontalAlignment = xlLeft
    Columns("K").HorizontalAlignment = xlLeft
    Columns("Q").HorizontalAlignment = xlCenter
    Columns("R").HorizontalAlignment = xlLeft
    
    printDebug dStartTime, Timer, "Formatted rows and columns"
    
    printDebug dStartTime, Timer, "Start For Loop and cycle through rows"
    
    '====================================================================
    'Cycle through all Rows which hava data in A column and apply colors
    '====================================================================
    
    Dim rngFullRowColorApply        As Range
    Dim rngMeasuredColO             As Range
    Dim sDeviceColD                 As String
    Dim sSubDeviceColE              As String
    Dim sTopicColK                  As String
    Dim sStatusColS                 As String
    Dim sMeasuredColO               As String
    Dim sColNValue                  As String
    
    For lRow = 2 To lMaxRow
        
        Set rngFullRowColorApply = Range(Cells(lRow, 1), Cells(lRow, colStatus))
        Set rngMeasuredColO = Cells(lRow, "O")
        sDeviceColD = Cells(lRow, "D").value
        sSubDeviceColE = Cells(lRow, "E").value
        sTopicColK = Cells(lRow, "K").value
        sStatusColS = Cells(lRow, colStatus).value
        sMeasuredColO = rngMeasuredColO.value
        
        sColNValue = Range("N" & lRow).value
        
        Select Case sStatusColS
            Case "FAIL"
                rngFullRowColorApply.Interior.Color = colorRed
                Select Case sTopicColK
                    Case "Run Suite Project"
                        Rows(lRow).RowHeight = heightHighRow
                    Case "Run Test"
                        Rows(lRow).RowHeight = heightHighRow
                End Select
            Case "ERROR"
                rngFullRowColorApply.Interior.Color = colorRed
                Select Case sTopicColK
                    Case "Run Suite Project"
                        Rows(lRow).RowHeight = heightHighRow
                    Case "Run Test"
                        Rows(lRow).RowHeight = heightHighRow
                End Select
            Case Else
                Select Case sDeviceColD
                    Case "TnM"
                        rngFullRowColorApply.Interior.Color = colorLightBlue
                    Case "File_Loop"
                        rngFullRowColorApply.Interior.Color = colorYellow
                    Case "Test"
                        Select Case sSubDeviceColE
                            Case "Running"
                                Select Case sTopicColK
                                    Case "Run Suite Project"
                                        Rows(lRow).RowHeight = heightHighRow
                                        rngFullRowColorApply.Interior.Color = colorLightGrey
                                    Case "Run Test"
                                        Rows(lRow).RowHeight = heightHighRow
                                        rngFullRowColorApply.Interior.Color = colorLightGrey
                                    Case "Set Variables"
                                        rngFullRowColorApply.Interior.Color = colorYellow
                                    Case "Comparison"
                                        rngFullRowColorApply.Interior.Color = colorOrange
                                    Case "Reference line"
                                        rngFullRowColorApply.Interior.Color = colorBrown
                                End Select
                            Case "Report"
                                Select Case sTopicColK
                                    Case "Text to report"
                                        rngMeasuredColO.Font.Color = vbWhite
                                        rngFullRowColorApply.Interior.Color = colorGreen
                                        If Left(sMeasuredColO, 1) = "#" Then
                                            rngFullRowColorApply.Interior.Color = colorBlue
                                        ElseIf Left(sMeasuredColO, 3) = ":::" Then
                                            rngFullRowColorApply.Interior.Color = colorBlack
                                        ElseIf Left(sMeasuredColO, 3) = "===" Then
                                            rngMeasuredColO.WrapText = True
                                            rngMeasuredColO.EntireRow.AutoFit
                                        ElseIf Left(sMeasuredColO, 3) = "---" Then
                                            rngMeasuredColO.WrapText = True
                                            rngMeasuredColO.EntireRow.AutoFit
                                        ElseIf Left(sMeasuredColO, 3) = "***" Then
                                            rngMeasuredColO.WrapText = True
                                            rngMeasuredColO.EntireRow.AutoFit
                                        Else
                                            rngFullRowColorApply.Interior.Color = colorGreen
                                            'Cells(row, "O").Font.Bold = True   'Starting 23-5-22 this row make macro stuck for 60 sec
                                        End If
                                End Select
                        End Select
                    Case Else
                        Select Case sSubDeviceColE
                            Case "NG_Rest_SNMP"
                                If (InStr(sColNValue, "ADD") > 0 Or InStr(sColNValue, "EDIT") > 0 Or InStr(sColNValue, "SET") > 0 Or InStr(sColNValue, "DELETE") > 0) Then
                                    rngMeasuredColO.Interior.Color = colorGetRed
                                ElseIf (InStr(sColNValue, "GET") > 0 Or InStr(sColNValue, "WALK") > 0 Or InStr(sColNValue, "FIND") > 0) Then
                                    rngMeasuredColO.Interior.Color = colorGetBlue
                                End If
                            Case "NG_SpecialCommands"
                                Select Case sTopicColK
                                    Case "NG_DynamicDelay"
                                        rngFullRowColorApply.Interior.Color = colorLightPurple
                                End Select
                            Case "Communication"
                                Select Case sTopicColK
                                    Case "Ping"
                                        rngFullRowColorApply.Interior.Color = colorLightPurple
                                End Select
                        End Select
                End Select
        End Select
        
        ' Create links to sheets for all "See walk results in sheet x" Cells
        If InStr(1, sMeasuredColO, "See the measured results") > 0 Then
            sHyperlinkSheetName = Mid(sMeasuredColO, InStr(1, sMeasuredColO, "'", 1) + 1, InStrRev(sMeasuredColO, "'") - InStr(1, sMeasuredColO, "'", 1) - 1)
            'Debug.Print ("<" & sHyperlinkSheetName & ">")
            ActiveCell.Hyperlinks.Add Anchor:=rngMeasuredColO, Address:="", SubAddress:="'" & sHyperlinkSheetName & "'" & "!A1"
        End If

    Next lRow

    printDebug dStartTime, Timer, "For loop end, start color set for fonts"
    
    ' Apply Format for Delay column
    Columns("Q").Font.Bold = True
    Columns("Q").Font.Color = colorDarkRed
    Columns("N").Font.Color = colorDarkGrey
    Columns("P").Font.Color = colorDarkGrey
    Columns("R").Font.Color = colorDarkGrey
    Columns("D").Font.Color = colorDarkGrey
    Columns("E").Font.Color = colorDarkGrey
    Columns("V").Font.Color = colorDarkGrey
    Columns("X").Font.Color = colorDarkGrey
    Columns("W").Font.Color = colorCommentBlue
    Columns("W").Font.Bold = True

    ' Define borders
    With Columns("A:Z").Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 48
    End With
    
    printDebug dStartTime, Timer, "Colors and fonts applied"
    
    ' Create links from all sheets to Results sheet
    Dim ws                  As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        If ws.Index > 2 Then
            'Debug.Print (ws.Name)
            With ws.Buttons.Add(1, 1, 45, 15)
            .OnAction = "ReturnToFirstSheet"
            .text = "Results"
            End With
        End If
    Next
    
    ActiveWindow.ScrollColumn = 1   'Scroll to the left
    printDebug dStartTime, Timer, "Created Links to results sheets"
    
    ' Create Filter buttons
    addFilterButton 0, "IDU", "ReportAutoFilterIDU"
    addFilterButton 1, "Filter", "ReportAutofilterFilterItems"
    addFilterButton 2, "Clear", "ReportAutofilterClear"
    addFilterButton 3, "NextFail", "GotoNextFail"
    printDebug dStartTime, Timer, "Created Filter buttons"
            
    ' Freeze top row
    ' First scroll to the top so first row seen in sight
    ActiveWindow.ScrollRow = 1
    With ActiveWindow
        If .FreezePanes Then .FreezePanes = False
        .SplitColumn = 0
        .SplitRow = 1
        .FreezePanes = True
    End With
    printDebug dStartTime, Timer, "Top row freezed"
    
    ActiveWorkbook.Save
    Application.ScreenUpdating = True
    printDebug dStartTime, Timer, "Workbook saved"
    
    ' Stop Timer
    printDebug dStartTime, Timer, "Macro finished !!!"
    dSecondsElapsed = Round(Timer - dStartTime, 2)
    Debug.Print ("Time took to run: " & dSecondsElapsed)
    
    ' Indicate Runtime in result
    Cells(4, "Z") = "Macro duration: " & dSecondsElapsed
    
End Sub


Sub ReturnToFirstSheet()
    Sheets("Result").Select
End Sub

Function macroAlreadyApplied() As Boolean
    Dim wks As Worksheet
    On Error Resume Next
    Set wks = Worksheets("Macro Logs")
    
    If wks Is Nothing Then
        macroAlreadyApplied = False
    Else
        macroAlreadyApplied = True
    End If
End Function

Function addFilterButton(buttonIndex, buttonName, onClickMacroName)
    '===========================
    ' Writen by Michael Rykin
    ' Create all buttons used for filtering results
    '===========================
    Const ButtonWidth = 70
    Dim filterBtn       As Button
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
    Dim lastEmptyMacroSheetRow      As Long
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
    Dim FindString      As String
    Dim Rng             As Range
    Dim ActiveRow       As Long
    Dim colStatus       As Integer
    
    FindString = "FAIL"
    ActiveRow = ActiveCell.row + 1
    colStatus = WorksheetFunction.Match("Status", Range("1:1"), 0)
    
    Debug.Print (ActiveRow)
    
    Set Rng = Columns(colStatus).Find(What:=FindString, _
                    After:=Cells(ActiveRow, colStatus), _
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

    Dim checkTimeFilePath               As String
    Dim macroFilesFolder                As String
    Dim alternativeMacroFilesFolder     As String
    Dim checkTimeFileName               As String
    Dim todayDate                       As Date
    Dim fso                             As Object
    Dim oFile                           As Object
    
    todayDate = Date
    macroFilesFolder = "C:\tmp"
    alternativeMacroFilesFolder = Environ("USERPROFILE") & "\Documents"
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
    Dim i           As Integer
    Dim iMaxRows    As Integer
    Dim testPath    As String

    iMaxRows = Worksheets(1).Cells(Rows.Count, "A").End(xlUp).row

    For i = 2 To iMaxRows
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
    Dim iRow             As Integer
    Dim col             As Integer
    Dim lMaxRow         As Long
    Dim value           As String
    Dim textPrint       As String
    lMaxRow = Cells(Rows.Count, "A").End(xlUp).row
    textPrint = ""

    For iRow = 2 To lMaxRow
        For col = 1 To 19
            On Error GoTo skip
            value = Cells(iRow, col).Validation.Formula1
            If InStr(value, "\") <> 0 Then
                Debug.Print ("Address " & Cells(iRow, col).Address & " Value: " & value)
                textPrint = textPrint & Cells(iRow, col).Address & vbCrLf
            End If
skip:
        'Cell have no Data validation (it is by defaul "Any Value")
        'Continiue to next cell
        Resume skip2
skip2:
        Next col
    Next iRow
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