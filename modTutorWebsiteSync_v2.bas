Attribute VB_Name = "modTutorWebsiteSync"
Option Explicit

Private Const CONFIG_SHEET As String = "TUTOR_WEBSITE_SYNC"
Private Const CFG_FOLDER_CELL As String = "B3"
Private Const CFG_FILE_CELL As String = "B4"
Private Const CFG_YEAR_CELL As String = "B5"
Private Const CFG_WRITEBACK_CELL As String = "B6"
Private Const CFG_INCLUDE_PENDING_CELL As String = "B7"

Public Sub SetupTutorWebsiteSync()
    Dim ws As Worksheet

    Set ws = EnsureConfigSheet()
    If ws Is Nothing Then Exit Sub

    ws.Range("A1").Value = "Tutor Website Sync"
    ws.Range("A3").Value = "JSON Output Folder"
    ws.Range(CFG_FOLDER_CELL).Value = IIf(Len(ThisWorkbook.Path) > 0, ThisWorkbook.Path, "")
    ws.Range("A4").Value = "JSON File Name"
    ws.Range(CFG_FILE_CELL).Value = "tutors.json"
    ws.Range("A5").Value = "Export Year"
    ws.Range(CFG_YEAR_CELL).Value = Year(Date)
    ws.Range("A6").Value = "Write Back Missing Codes?"
    ws.Range(CFG_WRITEBACK_CELL).Value = "Yes"
    ws.Range("A7").Value = "Include Pending Tutors?"
    ws.Range(CFG_INCLUDE_PENDING_CELL).Value = "Yes"
    ws.Range("A9").Value = "Use this sheet for tutor website JSON export settings only."
    ws.Range("A10").Value = "Your learner export data on the EXPORTS sheet remains unchanged."

    ws.Columns("A:B").AutoFit

    MsgBox "Tutor website sync has been set up on the '" & CONFIG_SHEET & "' sheet." & vbCrLf & _
           "1. Review the folder path in " & CFG_FOLDER_CELL & vbCrLf & _
           "2. Save this workbook as .xlsm" & vbCrLf & _
           "3. Run ExportTutorsJson", vbInformation, "MSH Tutor Website Sync"
End Sub

Public Sub ChooseTutorJsonFolder()
    Dim ws As Worksheet
    Dim fd As FileDialog

    Set ws = EnsureConfigSheet()
    If ws Is Nothing Then Exit Sub

    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    With fd
        .Title = "Select Website Folder for tutors.json"
        .AllowMultiSelect = False
        If .Show <> -1 Then Exit Sub
        ws.Range(CFG_FOLDER_CELL).Value = .SelectedItems(1)
    End With

    MsgBox "Website folder saved to " & CONFIG_SHEET & "!" & CFG_FOLDER_CELL, vbInformation, "MSH Tutor Website Sync"
End Sub

Public Sub ExportTutorsJson()
    Dim ws As Worksheet
    Dim headerMap As Object
    Dim lastRow As Long
    Dim r As Long
    Dim tutorName As String, code As String, verStatus As String
    Dim activeStatus As String, statusOut As String
    Dim subjectSpec As String, qualification As String, institution As String, roleText As String
    Dim folderPath As String, fileName As String, fullPath As String
    Dim exportYear As Long, writeBack As Boolean, includePending As Boolean
    Dim seq As Long, exportedCount As Long
    Dim json As String, itemJson As String

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets("TUTOR_REGISTER")
    Set headerMap = GetHeaderMap(ws)

    RequireHeader headerMap, "Tutor Name"
    RequireHeader headerMap, "Subject Specialty"
    RequireHeader headerMap, "Highest Qualification"
    RequireHeader headerMap, "University/Institution"
    RequireHeader headerMap, "Status (Active/Inactive)"
    RequireHeader headerMap, "Verification Code"
    RequireHeader headerMap, "Verification Status"
    RequireHeader headerMap, "Role / Position"

    folderPath = GetConfigValue(CFG_FOLDER_CELL, ThisWorkbook.Path)
    fileName = GetConfigValue(CFG_FILE_CELL, "tutors.json")
    exportYear = CLng(Val(GetConfigValue(CFG_YEAR_CELL, CStr(Year(Date)))))
    If exportYear = 0 Then exportYear = Year(Date)
    writeBack = IsYesValue(GetConfigValue(CFG_WRITEBACK_CELL, "Yes"))
    includePending = IsYesValue(GetConfigValue(CFG_INCLUDE_PENDING_CELL, "Yes"))

    If Len(folderPath) = 0 Then
        MsgBox "Please choose or enter the website output folder on " & CONFIG_SHEET & "!" & CFG_FOLDER_CELL & ".", vbExclamation, "MSH Tutor Website Sync"
        GoTo CleanExit
    End If

    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        fullPath = folderPath & fileName
    Else
        fullPath = folderPath & Application.PathSeparator & fileName
    End If

    lastRow = ws.Cells(ws.Rows.Count, GetColIndex(headerMap, "Tutor Name")).End(xlUp).Row

    json = "[" & vbCrLf
    seq = 1
    exportedCount = 0

    For r = 2 To lastRow
        tutorName = NzText(GetCellValue(ws, r, headerMap, "Tutor Name"))
        If Len(tutorName) > 0 Then
            subjectSpec = NzText(GetCellValue(ws, r, headerMap, "Subject Specialty"))
            qualification = NzText(GetCellValue(ws, r, headerMap, "Highest Qualification"))
            institution = NzText(GetCellValue(ws, r, headerMap, "University/Institution"))
            activeStatus = LCase$(NzText(GetCellValue(ws, r, headerMap, "Status (Active/Inactive)")))
            verStatus = NzText(GetCellValue(ws, r, headerMap, "Verification Status"))
            code = NzText(GetCellValue(ws, r, headerMap, "Verification Code"))
            roleText = NzText(GetCellValue(ws, r, headerMap, "Role / Position"))

            If Len(code) = 0 Then
                code = "MSH-TUT-" & exportYear & "-" & Format$(seq, "000")
                If writeBack Then
                    ws.Cells(r, GetColIndex(headerMap, "Verification Code")).Value = code
                    If Len(verStatus) = 0 And activeStatus = "active" Then
                        ws.Cells(r, GetColIndex(headerMap, "Verification Status")).Value = "Verified"
                        verStatus = "Verified"
                    End If
                End If
            End If

            If Len(verStatus) > 0 Then
                If LCase$(verStatus) = "verified" Then
                    statusOut = "Verified"
                Else
                    statusOut = "Pending"
                End If
            ElseIf activeStatus = "active" Then
                statusOut = "Verified"
            Else
                statusOut = "Pending"
            End If

            If includePending Or statusOut = "Verified" Then
                If Len(roleText) = 0 Then
                    If Len(subjectSpec) > 0 Then
                        roleText = subjectSpec & " Tutor"
                    Else
                        roleText = "Tutor"
                    End If
                End If

                itemJson = "  {" & vbCrLf & _
                           "    ""code"": """ & EscapeJson(code) & """," & vbCrLf & _
                           "    ""displayName"": """ & EscapeJson(MakeDisplayName(tutorName)) & """," & vbCrLf & _
                           "    ""name"": """ & EscapeJson(tutorName) & """," & vbCrLf & _
                           "    ""role"": """ & EscapeJson(roleText) & """," & vbCrLf & _
                           "    ""qualification"": """ & EscapeJson(qualification) & """," & vbCrLf & _
                           "    ""institution"": """ & EscapeJson(institution) & """," & vbCrLf & _
                           "    ""subjects"": " & JsonArrayFromSubjects(subjectSpec) & "," & vbCrLf & _
                           "    ""status"": """ & EscapeJson(statusOut) & """" & vbCrLf & _
                           "  }"

                If exportedCount > 0 Then
                    json = json & "," & vbCrLf
                End If
                json = json & itemJson
                exportedCount = exportedCount + 1
            End If

            seq = seq + 1
        End If
    Next r

    json = json & vbCrLf & "]"
    SaveUtf8Text fullPath, json

    If writeBack Then ThisWorkbook.Save

    MsgBox exportedCount & " tutor record(s) exported to:" & vbCrLf & fullPath, vbInformation, "MSH Tutor Website Sync"

CleanExit:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "Tutor JSON export failed: " & Err.Description, vbCritical, "MSH Tutor Website Sync"
End Sub

Public Sub ExportTutorsJsonAndOpenFolder()
    ExportTutorsJson
    OpenTutorWebsiteFolder
End Sub

Public Sub OpenTutorWebsiteFolder()
    Dim folderPath As String
    folderPath = GetConfigValue(CFG_FOLDER_CELL, ThisWorkbook.Path)

    If Len(folderPath) = 0 Then
        MsgBox "No website folder has been set yet.", vbExclamation, "MSH Tutor Website Sync"
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
        Err.Raise vbObjectError + 513, , "Required column not found in TUTOR_REGISTER: " & headerName
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

Private Function MakeDisplayName(ByVal fullName As String) As String
    Dim cleanName As String
    Dim parts() As String
    Dim firstPart As String, lastPart As String

    cleanName = Application.WorksheetFunction.Trim(fullName)
    If Len(cleanName) = 0 Then
        MakeDisplayName = ""
        Exit Function
    End If

    parts = Split(cleanName, " ")
    If UBound(parts) >= 1 Then
        firstPart = parts(LBound(parts))
        lastPart = parts(UBound(parts))
        MakeDisplayName = UCase$(Left$(firstPart, 1)) & ". " & StrConv(lastPart, vbProperCase)
    Else
        MakeDisplayName = cleanName
    End If
End Function

Private Function JsonArrayFromSubjects(ByVal subjectText As String) As String
    Dim cleaned As String
    Dim parts() As String
    Dim i As Long
    Dim oneSubject As String
    Dim outText As String
    Dim firstItem As Boolean

    cleaned = Trim$(subjectText)
    cleaned = Replace(cleaned, ";", ",")
    cleaned = Replace(cleaned, "/", ",")
    cleaned = Replace(cleaned, "|", ",")

    If Len(cleaned) = 0 Then
        JsonArrayFromSubjects = "[]"
        Exit Function
    End If

    parts = Split(cleaned, ",")
    outText = "["
    firstItem = True

    For i = LBound(parts) To UBound(parts)
        oneSubject = Trim$(parts(i))
        If Len(oneSubject) > 0 Then
            If Not firstItem Then outText = outText & ", "
            outText = outText & """" & EscapeJson(oneSubject) & """"
            firstItem = False
        End If
    Next i

    outText = outText & "]"
    JsonArrayFromSubjects = outText
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
