Attribute VB_Name = "modLearnerWebsiteSync"
Option Explicit

Private Const CONFIG_SHEET As String = "LEARNER_WEBSITE_SYNC"
Private Const CFG_FOLDER_CELL As String = "B3"
Private Const CFG_FILE_CELL As String = "B4"
Private Const CFG_ACTIVE_ONLY_CELL As String = "B5"
Private Const CFG_VERIFICATION_METHOD_CELL As String = "B6"

Public Sub SetupLearnerWebsiteSync()
    Dim ws As Worksheet

    Set ws = EnsureConfigSheet()
    If ws Is Nothing Then Exit Sub

    ws.Range("A1").Value = "Learner Website Sync"
    ws.Range("A3").Value = "JSON Output Folder"
    ws.Range(CFG_FOLDER_CELL).Value = IIf(Len(ThisWorkbook.Path) > 0, ThisWorkbook.Path, "")
    ws.Range("A4").Value = "JSON File Name"
    ws.Range(CFG_FILE_CELL).Value = "learners.json"
    ws.Range("A5").Value = "Include Only Active Learners?"
    ws.Range(CFG_ACTIVE_ONLY_CELL).Value = "Yes"
    ws.Range("A6").Value = "Verification Method"
    ws.Range(CFG_VERIFICATION_METHOD_CELL).Value = "Student Number + ID Number"
    ws.Range("A8").Value = "Privacy note: a public static HTML + JSON site is not true private verification."
    ws.Range("A9").Value = "This JSON should ideally be used behind a backend/API for real learner privacy."

    ws.Columns("A:B").AutoFit

    MsgBox "Learner website sync has been set up on the '" & CONFIG_SHEET & "' sheet." & vbCrLf & _
           "Review the settings, then run ExportLearnersJson.", vbInformation, "MSH Learner Website Sync"
End Sub

Public Sub ChooseLearnerJsonFolder()
    Dim ws As Worksheet
    Dim fd As FileDialog

    Set ws = EnsureConfigSheet()
    If ws Is Nothing Then Exit Sub

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select Website Folder for learners.json"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        ws.Range(CFG_FOLDER_CELL).Value = .SelectedItems(1)
    End With

    MsgBox "Website folder saved to " & CONFIG_SHEET & "!" & CFG_FOLDER_CELL, vbInformation, "MSH Learner Website Sync"
End Sub

Public Sub ExportLearnersJson()
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long, r As Long
    Dim folderPath As String, fileName As String, fullPath As String
    Dim studentNumber As String, idNumber As String, fullName As String, gradeText As String, schoolName As String, statusText As String
    Dim activeOnly As Boolean
    Dim exportedCount As Long, skippedCount As Long
    Dim json As String, itemJson As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets("MASTER_REGISTER")
    Set headerMap = GetHeaderMap(ws)

    RequireHeader headerMap, "Student Number"
    RequireHeader headerMap, "Full Name"
    RequireHeader headerMap, "Grade"
    RequireHeader headerMap, "Status"
    RequireHeader headerMap, "ID Number"
    RequireHeader headerMap, "School Name"

    folderPath = GetConfigValue(CFG_FOLDER_CELL, ThisWorkbook.Path)
    fileName = GetConfigValue(CFG_FILE_CELL, "learners.json")
    activeOnly = IsYesValue(GetConfigValue(CFG_ACTIVE_ONLY_CELL, "Yes"))

    If Len(folderPath) = 0 Then
        MsgBox "Please choose or enter the website output folder on " & CONFIG_SHEET & "!" & CFG_FOLDER_CELL & ".", vbExclamation, "MSH Learner Website Sync"
        GoTo CleanExit
    End If

    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        fullPath = folderPath & fileName
    Else
        fullPath = folderPath & Application.PathSeparator & fileName
    End If

    lastRow = ws.Cells(ws.Rows.Count, GetColIndex(headerMap, "Student Number")).End(xlUp).Row

    json = "[" & vbCrLf
    exportedCount = 0
    skippedCount = 0

    For r = 2 To lastRow
        studentNumber = NzText(GetCellValue(ws, r, headerMap, "Student Number"))
        If Len(studentNumber) > 0 Then
            idNumber = NzText(GetCellValue(ws, r, headerMap, "ID Number"))
            fullName = NzText(GetCellValue(ws, r, headerMap, "Full Name"))
            gradeText = NzText(GetCellValue(ws, r, headerMap, "Grade"))
            schoolName = NzText(GetCellValue(ws, r, headerMap, "School Name"))
            statusText = NzText(GetCellValue(ws, r, headerMap, "Status"))

            If activeOnly And LCase$(statusText) <> "active" Then
                skippedCount = skippedCount + 1
            ElseIf Len(idNumber) = 0 Then
                skippedCount = skippedCount + 1
            Else
                itemJson = "  {" & vbCrLf & _
                           "    ""studentNumber"": """ & EscapeJson(studentNumber) & """," & vbCrLf & _
                           "    ""idNumber"": """ & EscapeJson(idNumber) & """," & vbCrLf & _
                           "    ""fullName"": """ & EscapeJson(fullName) & """," & vbCrLf & _
                           "    ""grade"": """ & EscapeJson(gradeText) & """," & vbCrLf & _
                           "    ""schoolName"": """ & EscapeJson(schoolName) & """," & vbCrLf & _
                           "    ""status"": """ & EscapeJson(statusText) & """" & vbCrLf & _
                           "  }"

                If exportedCount > 0 Then
                    json = json & "," & vbCrLf
                End If
                json = json & itemJson
                exportedCount = exportedCount + 1
            End If
        End If
    Next r

    json = json & vbCrLf & "]"
    SaveUtf8Text fullPath, json

    MsgBox exportedCount & " learner record(s) exported to:" & vbCrLf & fullPath & vbCrLf & vbCrLf & _
           skippedCount & " record(s) were skipped (inactive or missing ID Number).", vbInformation, "MSH Learner Website Sync"

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Learner JSON export failed: " & Err.Description, vbCritical, "MSH Learner Website Sync"
End Sub

Public Sub ExportLearnersJsonAndOpenFolder()
    ExportLearnersJson
    OpenLearnerWebsiteFolder
End Sub

Public Sub OpenLearnerWebsiteFolder()
    Dim folderPath As String
    folderPath = GetConfigValue(CFG_FOLDER_CELL, ThisWorkbook.Path)

    If Len(folderPath) = 0 Then
        MsgBox "No learner website folder has been set yet.", vbExclamation, "MSH Learner Website Sync"
        Exit Sub
    End If

    Shell "explorer.exe " & Chr$(34) & folderPath & Chr$(34), vbNormalFocus
End Sub

Private Function EnsureConfigSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CONFIG_SHEET
    End If

    Set EnsureConfigSheet = ws
End Function

Private Function GetHeaderMap(ws As Worksheet) As Object
    Dim map As Object
    Dim lastCol As Long, c As Long
    Dim headerText As String

    Set map = CreateObject("Scripting.Dictionary")
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        headerText = Trim$(CStr(ws.Cells(1, c).Value))
        If Len(headerText) > 0 Then
            map(UCase$(headerText)) = c
        End If
    Next c

    Set GetHeaderMap = map
End Function

Private Sub RequireHeader(headerMap As Object, headerName As String)
    If Not headerMap.Exists(UCase$(headerName)) Then
        Err.Raise vbObjectError + 713, , "Required column not found in MASTER_REGISTER: " & headerName
    End If
End Sub

Private Function GetColIndex(headerMap As Object, headerName As String) As Long
    GetColIndex = CLng(headerMap(UCase$(headerName)))
End Function

Private Function GetCellValue(ws As Worksheet, rowNum As Long, headerMap As Object, headerName As String) As Variant
    GetCellValue = ws.Cells(rowNum, GetColIndex(headerMap, headerName)).Value
End Function

Private Function GetConfigValue(configCell As String, defaultValue As String) As String
    Dim ws As Worksheet
    Dim rawValue As String

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONFIG_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        GetConfigValue = defaultValue
        Exit Function
    End If

    rawValue = Trim$(CStr(ws.Range(configCell).Value))
    If Len(rawValue) = 0 Then
        GetConfigValue = defaultValue
    Else
        GetConfigValue = rawValue
    End If
End Function

Private Function IsYesValue(valueText As String) As Boolean
    valueText = LCase$(Trim$(valueText))
    IsYesValue = (valueText = "yes" Or valueText = "y" Or valueText = "true" Or valueText = "1")
End Function

Private Function NzText(v As Variant) As String
    If IsError(v) Then
        NzText = ""
    ElseIf IsNull(v) Then
        NzText = ""
    Else
        NzText = Trim$(CStr(v))
    End If
End Function

Private Function EscapeJson(ByVal valueText As String) As String
    valueText = Replace(valueText, "\", "\\")
    valueText = Replace(valueText, """", "\""")
    valueText = Replace(valueText, "/", "\/")
    valueText = Replace(valueText, vbCrLf, "\n")
    valueText = Replace(valueText, vbCr, "\n")
    valueText = Replace(valueText, vbLf, "\n")
    EscapeJson = valueText
End Function

Private Sub SaveUtf8Text(ByVal fullPath As String, ByVal fileText As String)
    Dim stm As Object

    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "utf-8"
        .Open
        .WriteText fileText
        .SaveToFile fullPath, 2
        .Close
    End With
End Sub
