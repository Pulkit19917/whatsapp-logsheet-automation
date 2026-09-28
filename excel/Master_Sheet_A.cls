VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Tracks the last (date, source) pair the viewer was refreshed for - so
' moving between DIFFERENT ROWS that happen to point at the SAME photo
' folder (e.g. Day/Night shift rows for one log sheet) does NOT reset
' the viewer's current zoom/rotation/photo selection. Only an actual
' change in date or source name triggers a reload.
Private lastDateText As String
Private lastSourceName As String


Private Sub Worksheet_BeforeDoubleClick( _
    ByVal Target As Range, _
    Cancel As Boolean)

    On Error GoTo ErrorHandler


    '-----------------------------------------------------
    ' ONLY DATA ROWS
    '-----------------------------------------------------

    If Target.Row < 2 Then Exit Sub


    '-----------------------------------------------------
    ' PREVENT NORMAL EXCEL EDIT MODE
    '-----------------------------------------------------

    Cancel = True


    '-----------------------------------------------------
    ' OPEN PHOTO VIEWER FOR THIS ROW
    '-----------------------------------------------------

    OpenPhotoViewerForRow "MASTER_LOG_SHEET_A", Target.Row

    ' Remember what this row resolves to, so a click on another cell in
    ' the SAME row right after opening doesn't immediately reload it again.
    RememberRowIdentity Target.Row


    Exit Sub


ErrorHandler:

    MsgBox "Error opening photo viewer:" & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Photo Viewer Error"

End Sub


'=========================================================
' AUTO-REFRESH PHOTO VIEWER WHEN THE ROW'S DATE/SOURCE CHANGES
'
' Fires on every click / arrow-key move. Only acts if the
' viewer window is already open. Compares the resolved
' Date + Source Name (not just the row number) against the
' last refresh - two different rows for the same photo
' folder won't reset the viewer's zoom/rotation/selection.
'=========================================================

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

    On Error Resume Next

    If Target.Row < 2 Then Exit Sub

    If Not IsPhotoViewerOpen() Then Exit Sub

    Dim newDateText As String
    Dim newSourceName As String

    newDateText = GetDateFolderName(Me.Cells(Target.Row, "B").value)
    newSourceName = Trim(CStr(Me.Cells(Target.Row, "C").value))

    ' Same folder as last time (even though the row changed) - leave the
    ' viewer exactly as it is, don't reload.
    If newDateText = lastDateText And newSourceName = lastSourceName Then
        Exit Sub
    End If

    RefreshPhotoViewerForRow "MASTER_LOG_SHEET_A", Target.Row

    lastDateText = newDateText
    lastSourceName = newSourceName

End Sub


'=========================================================
' REMEMBER THIS ROW'S DATE/SOURCE AS THE CURRENT IDENTITY
'=========================================================

Private Sub RememberRowIdentity(ByVal rowNumber As Long)

    On Error Resume Next

    lastDateText = GetDateFolderName(Me.Cells(rowNumber, "B").value)
    lastSourceName = Trim(CStr(Me.Cells(rowNumber, "C").value))

End Sub

