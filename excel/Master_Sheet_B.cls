VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'=========================================================
' MASTER_LOG_SHEET_B  -  worksheet module
'
' WHY THIS SHEET HAS ITS OWN OPEN/REFRESH CODE
' -------------------------------------------
' The shared module (modPhotoViewer) reads the date and the
' source-folder name from fixed columns:
'
'     Public Const DATE_COLUMN          As String = "B"
'     Public Const SOURCE_FOLDER_COLUMN As String = "C"
'
' Those are SHEET A's columns. This sheet keeps the date in
' column A and the source name in column B, so calling
' OpenPhotoViewerForRow "MASTER_LOG_SHEET_B" made the shared
' code read the wrong two cells - it picked up a blank or the
' wrong text as the source name, so the "Mundra-1 / Mundra-2 /
' Mundra-3 all merge into Mundra" lookup never had anything
' valid to match on.
'
' Rather than change the shared constants (which would break
' Sheet A), this sheet reads its OWN columns and then calls the
' exact same public helpers Sheet A ends up using:
'
'     GetDateFolderName
'     FolderExists
'     FindFolderByName
'     FindAllMatchingFolders   <- the multi-folder merge
'     frmPhotoViewer.LoadPhotos
'
' So the behaviour is identical to Sheet A: every folder whose
' name is the source name, with or without a -1 / -2 / -3
' suffix, is found and all of their photos are shown together
' in one viewer.
'=========================================================


'---------------------------------------------------------
' THIS SHEET'S COLUMNS - change here if they ever move
'---------------------------------------------------------

Private Const DATE_COL As String = "C"

Private Const SOURCE_COL As String = "E"


' Tracks the last (date, source) pair the viewer was refreshed for - so
' moving between DIFFERENT ROWS that happen to point at the SAME photo
' folder (e.g. Day/Night shift rows for one log sheet) does NOT reset
' the viewer's current zoom/rotation/photo selection. Only an actual
' change in date or source name triggers a reload.
Private lastDateText As String
Private lastSourceName As String


'=========================================================
' DOUBLE-CLICK A DATA ROW -> OPEN THE VIEWER
'=========================================================

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

    OpenViewerForRowOnThisSheet Target.Row

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

    newDateText = GetDateFolderName(Me.Cells(Target.Row, DATE_COL).value)
    newSourceName = Trim$(CStr(Me.Cells(Target.Row, SOURCE_COL).value))

    ' Same folder as last time (even though the row changed) - leave the
    ' viewer exactly as it is, don't reload.
    If newDateText = lastDateText And newSourceName = lastSourceName Then
        Exit Sub
    End If

    RefreshViewerForRowOnThisSheet Target.Row

    lastDateText = newDateText
    lastSourceName = newSourceName

End Sub


'=========================================================
' REMEMBER THIS ROW'S DATE/SOURCE AS THE CURRENT IDENTITY
'=========================================================

Private Sub RememberRowIdentity(ByVal rowNumber As Long)

    On Error Resume Next

    lastDateText = GetDateFolderName(Me.Cells(rowNumber, DATE_COL).value)
    lastSourceName = Trim$(CStr(Me.Cells(rowNumber, SOURCE_COL).value))

End Sub


'=========================================================
' RESOLVE THIS ROW -> DATE FOLDER + ALL MATCHING SOURCE FOLDERS
'
' Returns True only if everything resolved. On success:
'     dateTextOut      = "2026-09-17"
'     sourceNameOut    = "Mundra"
'     matchedFoldersOut = collection of full paths, e.g.
'                         ...\2026-09-17\Mundra
'                         ...\2026-09-17\Mundra-2
'                         ...\2026-09-17\Mundra-3
'
' showMessages:
'     True  - used by the double-click path, explains what
'             went wrong (empty date, missing folder, etc.)
'     False - used by the silent auto-refresh path, which must
'             never pop a message box while you arrow around
'             the sheet.
'=========================================================

Private Function ResolveRow( _
    ByVal rowNumber As Long, _
    ByRef dateTextOut As String, _
    ByRef sourceNameOut As String, _
    ByRef matchedFoldersOut As Collection, _
    ByVal showMessages As Boolean) As Boolean

    Dim dateFolder As String

    ResolveRow = False


    '-----------------------------------------------------
    ' READ DATE  (column A on this sheet)
    '-----------------------------------------------------

    If Trim$(CStr(Me.Cells(rowNumber, DATE_COL).value)) = "" Then

        If showMessages Then
            MsgBox "Date is empty in row " & rowNumber, _
                   vbExclamation, _
                   "Date Missing"
        End If

        Exit Function

    End If


    dateTextOut = GetDateFolderName(Me.Cells(rowNumber, DATE_COL).value)


    If dateTextOut = "" Then

        If showMessages Then
            MsgBox "Invalid date in cell " & _
                   DATE_COL & rowNumber & "." & vbCrLf & _
                   "Expected an Excel date such as 17/09/2026.", _
                   vbExclamation, _
                   "Invalid Date"
        End If

        Exit Function

    End If


    '-----------------------------------------------------
    ' READ SOURCE FILE NAME / FOLDER NAME  (column B here)
    '-----------------------------------------------------

    sourceNameOut = Trim$(CStr(Me.Cells(rowNumber, SOURCE_COL).value))


    If sourceNameOut = "" Then

        If showMessages Then
            MsgBox "Source File Name / Folder is empty in row " & _
                   rowNumber, _
                   vbExclamation, _
                   "Source Folder Missing"
        End If

        Exit Function

    End If


    '-----------------------------------------------------
    ' CHECK MAIN PHOTO FOLDER
    '-----------------------------------------------------

    If Not FolderExists(MAIN_PHOTO_FOLDER) Then

        If showMessages Then
            MsgBox "Main photo folder not found:" & vbCrLf & _
                   MAIN_PHOTO_FOLDER & vbCrLf & vbCrLf & _
                   "Please update MAIN_PHOTO_FOLDER in the VBA module.", _
                   vbCritical, _
                   "Main Folder Not Found"
        End If

        Exit Function

    End If


    '-----------------------------------------------------
    ' FIND DATE FOLDER
    '
    ' Example:
    ' Excel date = 17/09/2026
    ' Folder     = 2026-09-17
    '-----------------------------------------------------

    dateFolder = FindFolderByName(MAIN_PHOTO_FOLDER, dateTextOut)


    If dateFolder = "" Then

        If showMessages Then
            MsgBox "Date folder not found:" & vbCrLf & _
                   dateTextOut & vbCrLf & vbCrLf & _
                   "Searched in:" & vbCrLf & _
                   MAIN_PHOTO_FOLDER, _
                   vbExclamation, _
                   "Date Folder Not Found"
        End If

        Exit Function

    End If


    '-----------------------------------------------------
    ' FIND ALL MATCHING SOURCE FOLDERS
    '
    ' This is the same merge Sheet A gets:
    '
    ' Excel value    = Mundra
    ' Actual folders = Mundra, Mundra-2, Mundra-3
    '
    ' ALL matching folders are found and their photos are
    ' merged together into the one viewer.
    '-----------------------------------------------------

    Set matchedFoldersOut = FindAllMatchingFolders(dateFolder, sourceNameOut)


    If matchedFoldersOut.Count = 0 Then

        If showMessages Then
            MsgBox "Source folder not found:" & vbCrLf & _
                   sourceNameOut & vbCrLf & vbCrLf & _
                   "Date folder:" & vbCrLf & _
                   dateFolder & vbCrLf & vbCrLf & _
                   "The macro also checked suffix folders such as:" & _
                   vbCrLf & _
                   sourceNameOut & "-2" & vbCrLf & _
                   sourceNameOut & "-3" & vbCrLf & _
                   sourceNameOut & "-4", _
                   vbExclamation, _
                   "Source Folder Not Found"
        End If

        Exit Function

    End If


    ResolveRow = True

End Function


'=========================================================
' OPEN THE VIEWER FOR A ROW ON THIS SHEET
'=========================================================

Private Sub OpenViewerForRowOnThisSheet(ByVal rowNumber As Long)

    On Error GoTo ErrorHandler

    Dim dateText As String
    Dim sourceName As String
    Dim matchedFolders As Collection

    If Not ResolveRow(rowNumber, dateText, sourceName, matchedFolders, True) Then
        Exit Sub
    End If


    Load frmPhotoViewer

    frmPhotoViewer.LoadPhotos _
        dateText, _
        sourceName, _
        matchedFolders, _
        rowNumber

    frmPhotoViewer.Show vbModeless

    Exit Sub


ErrorHandler:

    MsgBox "Error while opening photo viewer:" & vbCrLf & _
           Err.Description, _
           vbCritical, _
           "Photo Viewer Error"

End Sub


'=========================================================
' REFRESH AN ALREADY-OPEN VIEWER FOR A ROW ON THIS SHEET
'
' Silent by design - no message boxes, since this runs on
' every selection change.
'=========================================================

Private Sub RefreshViewerForRowOnThisSheet(ByVal rowNumber As Long)

    On Error Resume Next

    Dim dateText As String
    Dim sourceName As String
    Dim matchedFolders As Collection

    If Not ResolveRow(rowNumber, dateText, sourceName, matchedFolders, False) Then
        Exit Sub
    End If

    frmPhotoViewer.LoadPhotos _
        dateText, _
        sourceName, _
        matchedFolders, _
        rowNumber

End Sub


