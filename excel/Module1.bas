Attribute VB_Name = "Module1"
Option Explicit

'=========================================================
' EKTA SITE PHOTO VIEWER
'
' Main folder:
'     Date folder
'         Source File Name folder
'             Photos
'
' Matching supports:
'
' Excel value:       Ennore
' Actual folder:     ENNORE
' Actual folder:     ENNORE-2
' Actual folder:     ENNORE-3
' Actual folder:     ENNORE_4
' Actual folder:     ENNORE 5
'
' Exact folder is preferred first.
' If exact folder is not found, suffix folder is used.
'=========================================================


'=========================================================
' CONFIGURATION
'=========================================================

Public Const MAIN_PHOTO_FOLDER As String = _
    "\\MANISH\Users\Nitesh\Desktop\WA Folder\output"


'---------------------------------------------------------
' CHANGE THESE ONLY IF YOUR COLUMNS ARE DIFFERENT
'---------------------------------------------------------

Public Const DATE_COLUMN As String = "C"

Public Const SOURCE_FOLDER_COLUMN As String = "E"


'=========================================================
' OPEN PHOTO VIEWER FOR SELECTED EXCEL ROW
'=========================================================

Public Sub OpenPhotoViewerForRow( _
    ByVal sheetName As String, _
    ByVal rowNumber As Long)

    On Error GoTo ErrorHandler

    Dim ws As Worksheet

    Dim dateText As String
    Dim sourceName As String

    Dim dateFolder As String
    Dim matchedFolders As Collection

    Set ws = ThisWorkbook.Worksheets(sheetName)


    '-----------------------------------------------------
    ' READ DATE
    '-----------------------------------------------------

    If Trim$(CStr(ws.Cells(rowNumber, DATE_COLUMN).value)) = "" Then

        MsgBox "Date is empty in row " & rowNumber, _
               vbExclamation, _
               "Date Missing"

        Exit Sub

    End If


    dateText = GetDateFolderName( _
                    ws.Cells(rowNumber, DATE_COLUMN).value)


    If dateText = "" Then

        MsgBox "Invalid date in cell " & _
               DATE_COLUMN & rowNumber & "." & vbCrLf & _
               "Expected an Excel date such as 17/09/2026.", _
               vbExclamation, _
               "Invalid Date"

        Exit Sub

    End If


    '-----------------------------------------------------
    ' READ SOURCE FILE NAME / FOLDER NAME
    '-----------------------------------------------------

    sourceName = Trim$( _
                    CStr(ws.Cells( _
                        rowNumber, _
                        SOURCE_FOLDER_COLUMN).value))


    If sourceName = "" Then

        MsgBox "Source File Name / Folder is empty in row " & _
               rowNumber, _
               vbExclamation, _
               "Source Folder Missing"

        Exit Sub

    End If


    '-----------------------------------------------------
    ' CHECK MAIN PHOTO FOLDER
    '-----------------------------------------------------

    If Not FolderExists(MAIN_PHOTO_FOLDER) Then

        MsgBox "Main photo folder not found:" & vbCrLf & _
               MAIN_PHOTO_FOLDER & vbCrLf & vbCrLf & _
               "Please update MAIN_PHOTO_FOLDER in the VBA module.", _
               vbCritical, _
               "Main Folder Not Found"

        Exit Sub

    End If


    '-----------------------------------------------------
    ' FIND DATE FOLDER
    '
    ' Example:
    ' Excel date = 17/09/2026
    ' Folder     = 2026-09-17
    '-----------------------------------------------------

    dateFolder = FindFolderByName( _
                    MAIN_PHOTO_FOLDER, _
                    dateText)


    If dateFolder = "" Then

        MsgBox "Date folder not found:" & vbCrLf & _
               dateText & vbCrLf & vbCrLf & _
               "Searched in:" & vbCrLf & _
               MAIN_PHOTO_FOLDER, _
               vbExclamation, _
               "Date Folder Not Found"

        Exit Sub

    End If


    '-----------------------------------------------------
    ' FIND ALL MATCHING SOURCE FOLDERS
    '
    ' Example:
    ' Excel value    = Mundra
    ' Actual folders = Mundra-1, Mundra-2, Mundra-3
    '
    ' ALL matching folders are found and their photos are
    ' merged together in the viewer.
    '-----------------------------------------------------

    Set matchedFolders = FindAllMatchingFolders( _
                            dateFolder, _
                            sourceName)


    If matchedFolders.Count = 0 Then

        MsgBox "Source folder not found:" & vbCrLf & _
               sourceName & vbCrLf & vbCrLf & _
               "Date folder:" & vbCrLf & _
               dateFolder & vbCrLf & vbCrLf & _
               "The macro also checked suffix folders such as:" & _
               vbCrLf & _
               sourceName & "-2" & vbCrLf & _
               sourceName & "-3" & vbCrLf & _
               sourceName & "-4", _
               vbExclamation, _
               "Source Folder Not Found"

        Exit Sub

    End If


    '-----------------------------------------------------
    ' OPEN PHOTO VIEWER
    '-----------------------------------------------------

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
' CHECK IF PHOTO VIEWER IS CURRENTLY OPEN
'=========================================================

Public Function IsPhotoViewerOpen() As Boolean

    Dim frm As Object

    For Each frm In VBA.UserForms

        If TypeName(frm) = "frmPhotoViewer" Then

            IsPhotoViewerOpen = True
            Exit Function

        End If

    Next frm

    IsPhotoViewerOpen = False

End Function


'=========================================================
' REFRESH ALREADY OPEN PHOTO VIEWER
'=========================================================

Public Sub RefreshPhotoViewerForRow( _
    ByVal sheetName As String, _
    ByVal rowNumber As Long)

    On Error Resume Next

    Dim ws As Worksheet

    Dim dateText As String
    Dim sourceName As String

    Dim dateFolder As String
    Dim matchedFolders As Collection

    Set ws = ThisWorkbook.Worksheets(sheetName)


    '-----------------------------------------------------
    ' READ DATE
    '-----------------------------------------------------

    If Trim$(CStr(ws.Cells(rowNumber, DATE_COLUMN).value)) = "" Then
        Exit Sub
    End If


    dateText = GetDateFolderName( _
                    ws.Cells(rowNumber, DATE_COLUMN).value)


    If dateText = "" Then
        Exit Sub
    End If


    '-----------------------------------------------------
    ' READ SOURCE FOLDER NAME
    '-----------------------------------------------------

    sourceName = Trim$( _
                    CStr(ws.Cells( _
                        rowNumber, _
                        SOURCE_FOLDER_COLUMN).value))


    If sourceName = "" Then
        Exit Sub
    End If


    If Not FolderExists(MAIN_PHOTO_FOLDER) Then
        Exit Sub
    End If


    '-----------------------------------------------------
    ' FIND DATE FOLDER
    '-----------------------------------------------------

    dateFolder = FindFolderByName( _
                    MAIN_PHOTO_FOLDER, _
                    dateText)


    If dateFolder = "" Then
        Exit Sub
    End If


    '-----------------------------------------------------
    ' FIND ALL MATCHING SOURCE FOLDERS
    '
    ' This uses the same multi-folder merge logic as
    ' OpenPhotoViewerForRow.
    '-----------------------------------------------------

    Set matchedFolders = FindAllMatchingFolders( _
                            dateFolder, _
                            sourceName)


    If matchedFolders.Count = 0 Then
        Exit Sub
    End If


    '-----------------------------------------------------
    ' REFRESH EXISTING VIEWER
    '-----------------------------------------------------

    frmPhotoViewer.LoadPhotos _
        dateText, _
        sourceName, _
        matchedFolders, _
        rowNumber

End Sub


'=========================================================
' CONVERT EXCEL DATE TO YYYY-MM-DD
'=========================================================

Public Function GetDateFolderName( _
    ByVal cellValue As Variant) As String

    On Error GoTo InvalidDate

    Dim d As Date

    If IsDate(cellValue) Then

        d = CDate(cellValue)

        GetDateFolderName = Format$(d, "yyyy-mm-dd")

    Else

        GetDateFolderName = ""

    End If

    Exit Function


InvalidDate:

    GetDateFolderName = ""

End Function


'=========================================================
' CHECK FOLDER EXISTS
'=========================================================

Public Function FolderExists( _
    ByVal folderPath As String) As Boolean

    On Error GoTo FolderError

    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")

    FolderExists = fso.FolderExists(folderPath)

    Exit Function


FolderError:

    FolderExists = False

End Function


'=========================================================
' NORMALIZE FOLDER NAME
'
' Examples:
'
' Ennore          -> ennore
' ENNORE-2        -> ennore
' ENNORE-3        -> ennore
' ENNORE_4        -> ennore
' ENNORE 5        -> ennore
'
' Kattupalli      -> kattupalli
' KATTUPALLI-2    -> kattupalli
'
' Kattupalli-2025-26 -> kattupalli
'
' Case, spaces, hyphens and underscores are ignored.
'=========================================================

Private Function NormalizeFolderKey( _
    ByVal folderName As String) As String

    Dim re As Object

    Dim oldValue As String
    Dim newValue As String

    newValue = LCase$(Trim$(folderName))


    Set re = CreateObject("VBScript.RegExp")

    With re

        .Global = True
        .IgnoreCase = True

        'Remove trailing numeric suffixes.
        '
        'Examples:
        '   ennore-2
        '   ennore-3
        '   kattupalli-2025-26
        '
        .Pattern = "[\s_-]+\d+$"

    End With


    'Keep removing trailing numeric parts
    'until no more are found.

    Do

        oldValue = newValue

        newValue = re.Replace(newValue, "")

    Loop While newValue <> oldValue


    'Remove spaces, hyphens and underscores.

    newValue = Replace$(newValue, " ", "")
    newValue = Replace$(newValue, "-", "")
    newValue = Replace$(newValue, "_", "")


    NormalizeFolderKey = newValue

End Function


'=========================================================
' FIND FOLDER BY NAME
'
' MATCHING ORDER:
'
' 1. Exact direct folder name
' 2. Exact folder one level deeper
' 3. Base-name matching in direct folders
' 4. Base-name matching one level deeper
'
' Examples:
'
' Excel value: Ennore
'
' Actual folders:
'   ENNORE
'   ENNORE-2
'   ENNORE-3
'
' Result:
'   ENNORE is selected first.
'
' If ENNORE does not exist:
'   ENNORE-2 or ENNORE-3 is selected.
'
' Matching is case-insensitive.
' Spaces, hyphens and underscores are ignored.
'=========================================================

Public Function FindFolderByName( _
    ByVal parentPath As String, _
    ByVal wantedName As String) As String

    On Error GoTo SearchError

    Dim fso As Object

    Dim parentFolder As Object
    Dim subFolder As Object
    Dim nestedFolder As Object

    Dim wantedKey As String
    Dim folderKey As String

    Dim fallbackPath As String
    Dim fallbackLength As Long

    Set fso = CreateObject("Scripting.FileSystemObject")

    Set parentFolder = fso.GetFolder(parentPath)


    wantedName = Trim$(wantedName)


    If wantedName = "" Then

        FindFolderByName = ""

        Exit Function

    End If


    wantedKey = NormalizeFolderKey(wantedName)


    '=====================================================
    ' STEP 1
    ' EXACT MATCH IN DIRECT SUBFOLDERS
    '=====================================================

    For Each subFolder In parentFolder.SubFolders

        If StrComp( _
            Trim$(subFolder.Name), _
            wantedName, _
            vbTextCompare) = 0 Then

            FindFolderByName = subFolder.Path

            Exit Function

        End If

    Next subFolder


    '=====================================================
    ' STEP 2
    ' EXACT MATCH ONE LEVEL DEEPER
    '=====================================================

    For Each subFolder In parentFolder.SubFolders

        For Each nestedFolder In subFolder.SubFolders

            If StrComp( _
                Trim$(nestedFolder.Name), _
                wantedName, _
                vbTextCompare) = 0 Then

                FindFolderByName = nestedFolder.Path

                Exit Function

            End If

        Next nestedFolder

    Next subFolder


    '=====================================================
    ' STEP 3
    ' BASE-NAME MATCH IN DIRECT SUBFOLDERS
    '
    ' Ennore matches:
    '   ENNORE-2
    '   ENNORE-3
    '   ENNORE-4
    '
    ' Kattupalli matches:
    '   KATTUPALLI-2
    '   KATTUPALLI_3
    '   KATTUPALLI 4
    '=====================================================

    For Each subFolder In parentFolder.SubFolders

        folderKey = NormalizeFolderKey(subFolder.Name)


        If folderKey = wantedKey Then

            'Prefer the shortest matching name.
            '
            'Example:
            '   ENNORE-2 is preferred over ENNORE-2026

            If fallbackPath = "" _
               Or Len(subFolder.Name) < fallbackLength Then

                fallbackPath = subFolder.Path

                fallbackLength = Len(subFolder.Name)

            End If

        End If

    Next subFolder


    '=====================================================
    ' STEP 4
    ' BASE-NAME MATCH ONE LEVEL DEEPER
    '=====================================================

    For Each subFolder In parentFolder.SubFolders

        For Each nestedFolder In subFolder.SubFolders

            folderKey = NormalizeFolderKey(nestedFolder.Name)


            If folderKey = wantedKey Then

                If fallbackPath = "" _
                   Or Len(nestedFolder.Name) < fallbackLength Then

                    fallbackPath = nestedFolder.Path

                    fallbackLength = Len(nestedFolder.Name)

                End If

            End If

        Next nestedFolder

    Next subFolder


    '=====================================================
    ' STEP 5
    ' RETURN MATCHING SUFFIX FOLDER
    '=====================================================

    If fallbackPath <> "" Then

        FindFolderByName = fallbackPath

    Else

        FindFolderByName = ""

    End If

    Exit Function


SearchError:

    FindFolderByName = ""

End Function


'=========================================================
' FIND ALL FOLDERS MATCHING A NAME
'
' Unlike FindFolderByName above (which picks a single best
' match), this returns EVERY folder whose name normalizes to
' the same base name - so a site's duplicate/batch upload
' folders for the same date are ALL included, not just the
' first one found. Example:
'
'   Wanted:  Mundra
'   Matches: Mundra-1, Mundra-2, Mundra-3   (all three)
'
' Uses the same NormalizeFolderKey rules as FindFolderByName
' (case, spaces, hyphens, underscores, and trailing numeric
' suffixes are all ignored for matching purposes).
'
' Searches direct subfolders first; only falls back to one
' level deeper if nothing matched directly.
'=========================================================

Public Function FindAllMatchingFolders( _
    ByVal parentPath As String, _
    ByVal wantedName As String) As Collection

    Dim results As New Collection

    On Error GoTo SearchError

    Dim fso As Object
    Dim parentFolder As Object
    Dim subFolder As Object
    Dim nestedFolder As Object

    Dim wantedKey As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set parentFolder = fso.GetFolder(parentPath)

    wantedName = Trim$(wantedName)

    If wantedName = "" Then

        Set FindAllMatchingFolders = results
        Exit Function

    End If

    wantedKey = NormalizeFolderKey(wantedName)


    '-----------------------------------------------------
    ' DIRECT SUBFOLDERS
    '-----------------------------------------------------

    For Each subFolder In parentFolder.SubFolders

        If NormalizeFolderKey(subFolder.Name) = wantedKey Then

            results.Add subFolder.Path

        End If

    Next subFolder


    '-----------------------------------------------------
    ' IF NONE FOUND DIRECTLY, SEARCH ONE LEVEL DEEPER
    '-----------------------------------------------------

    If results.Count = 0 Then

        For Each subFolder In parentFolder.SubFolders

            For Each nestedFolder In subFolder.SubFolders

                If NormalizeFolderKey(nestedFolder.Name) = wantedKey Then

                    results.Add nestedFolder.Path

                End If

            Next nestedFolder

        Next subFolder

    End If

    Set FindAllMatchingFolders = results
    Exit Function

SearchError:

    Set FindAllMatchingFolders = results

End Function


'=========================================================
' GET PHOTO FILES FROM FOLDER
'
' Supported:
'   JPG
'   JPEG
'   BMP
'   GIF
'   PNG
'
' Includes photos inside subfolders.
'=========================================================

Public Sub GetPhotoFiles( _
    ByVal folderPath As String, _
    ByRef photoFiles() As String, _
    ByRef photoCount As Long)

    On Error GoTo SearchError

    Dim fso As Object

    Dim folder As Object
    Dim file As Object
    Dim subFolder As Object

    Dim ext As String


    Set fso = CreateObject("Scripting.FileSystemObject")

    Set folder = fso.GetFolder(folderPath)


    '-----------------------------------------------------
    ' SEARCH FILES IN CURRENT FOLDER
    '-----------------------------------------------------

    For Each file In folder.Files

        ext = LCase$(fso.GetExtensionName(file.Name))


        If ext = "jpg" _
           Or ext = "jpeg" _
           Or ext = "bmp" _
           Or ext = "gif" _
           Or ext = "png" Then


            photoCount = photoCount + 1


            ReDim Preserve photoFiles(1 To photoCount)


            photoFiles(photoCount) = file.Path

        End If

    Next file


    '-----------------------------------------------------
    ' SEARCH SUBFOLDERS RECURSIVELY
    '-----------------------------------------------------

    For Each subFolder In folder.SubFolders

        GetPhotoFiles _
            subFolder.Path, _
            photoFiles, _
            photoCount

    Next subFolder


    Exit Sub


SearchError:

    'Ignore inaccessible folders and continue.

End Sub


'=========================================================
' OPEN PHOTO FOLDER IN WINDOWS EXPLORER
'=========================================================

Public Sub OpenFolder( _
    ByVal folderPath As String)

    Shell "explorer.exe " & _
          Chr$(34) & folderPath & Chr$(34), _
          vbNormalFocus

End Sub

