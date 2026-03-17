Attribute VB_Name = "modLearnerSyncButtons"
Option Explicit

Private Const CONTROL_SHEET As String = "LEARNER_WEBSITE_SYNC"

Public Sub InstallLearnerSyncButtons()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CONTROL_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Sheet '" & CONTROL_SHEET & "' was not found. Run SetupLearnerWebsiteSync first or use the learner-backend-ready workbook.", vbExclamation, "MSH Buttons"
        Exit Sub
    End If

    RemoveLearnerSyncButtons

    AddOneButton ws, "btnSetupLearnerSync", ws.Range("D5").Left, ws.Range("D5").Top, 195, 28, "Setup Learner Sync", "SetupLearnerWebsiteSync"
    AddOneButton ws, "btnChooseLearnerFolder", ws.Range("D7").Left, ws.Range("D7").Top, 195, 28, "Choose Data Folder", "ChooseLearnerJsonFolder"
    AddOneButton ws, "btnExportLearnerJson", ws.Range("D9").Left, ws.Range("D9").Top, 195, 28, "Export Learner JSON", "ExportLearnersJson"
    AddOneButton ws, "btnOpenLearnerFolder", ws.Range("D11").Left, ws.Range("D11").Top, 195, 28, "Open Data Folder", "OpenLearnerWebsiteFolder"
    AddOneButton ws, "btnExportLearnerAndOpen", ws.Range("D13").Left, ws.Range("D13").Top, 195, 28, "Export + Open Folder", "ExportLearnersJsonAndOpenFolder"

    MsgBox "Learner sync buttons installed on '" & CONTROL_SHEET & "'.", vbInformation, "MSH Buttons"
End Sub

Public Sub RemoveLearnerSyncButtons()
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
