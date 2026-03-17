Attribute VB_Name = "modTutorSyncButtons"
Option Explicit

Private Const CONTROL_SHEET As String = "TUTOR_WEBSITE_SYNC"

Public Sub InstallTutorSyncButtons()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & CONTROL_SHEET & "' was not found. Run SetupTutorWebsiteSync first or use the buttons-ready workbook.", vbExclamation, "MSH Buttons"
        Exit Sub
    End If

    RemoveTutorSyncButtons

    AddOneButton ws, "btnSetupTutorSync", ws.Range("D5").Left, ws.Range("D5").Top, 180, 28, "Setup Tutor Sync", "SetupTutorWebsiteSync"
    AddOneButton ws, "btnChooseTutorFolder", ws.Range("D7").Left, ws.Range("D7").Top, 180, 28, "Choose Website Folder", "ChooseTutorJsonFolder"
    AddOneButton ws, "btnExportTutorJson", ws.Range("D9").Left, ws.Range("D9").Top, 180, 28, "Export Tutors JSON", "ExportTutorsJson"
    AddOneButton ws, "btnOpenTutorFolder", ws.Range("D11").Left, ws.Range("D11").Top, 180, 28, "Open Website Folder", "OpenTutorWebsiteFolder"
    AddOneButton ws, "btnExportAndOpen", ws.Range("D13").Left, ws.Range("D13").Top, 180, 28, "Export + Open Folder", "ExportTutorsJsonAndOpenFolder"

    MsgBox "Tutor sync buttons installed on '" & CONTROL_SHEET & "'.", vbInformation, "MSH Buttons"
End Sub

Public Sub RemoveTutorSyncButtons()
    Dim ws As Worksheet
    Dim i As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then Exit Sub

    For i = ws.Buttons.Count To 1 Step -1
        If Left$(ws.Buttons(i).Name, 3) = "btn" Then
            ws.Buttons(i).Delete
        End If
    Next i
End Sub

Private Sub AddOneButton(ByVal ws As Worksheet, ByVal btnName As String, ByVal btnLeft As Double, ByVal btnTop As Double, ByVal btnWidth As Double, ByVal btnHeight As Double, ByVal captionText As String, ByVal macroName As String)
    Dim btn As Button

    Set btn = ws.Buttons.Add(btnLeft, btnTop, btnWidth, btnHeight)
    btn.Name = btnName
    btn.OnAction = macroName
    btn.Characters.Text = captionText
End Sub
