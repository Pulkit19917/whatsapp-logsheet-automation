VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPhotoViewer 
   Caption         =   "UserForm1"
   ClientHeight    =   6345
   ClientLeft      =   -75
   ClientTop       =   9465.001
   ClientWidth     =   11265
   OleObjectBlob   =   "frmPhotoViewer.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmPhotoViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'=========================================================
' EKTA LOGBOOK PHOTO VIEWER (frmPhotoViewer)
'=========================================================

'=========================================================
' WINDOWS / GDI+ DECLARATIONS
'=========================================================

#If VBA7 Then


    Private Declare PtrSafe Function GdiplusStartup Lib "gdiplus" ( _
        ByRef token As LongPtr, _
        ByRef inputbuf As GdiplusStartupInput, _
        ByVal outputbuf As LongPtr) As Long

    Private Declare PtrSafe Sub GdiplusShutdown Lib "gdiplus" ( _
        ByVal token As LongPtr)

    Private Declare PtrSafe Function GdipCreateBitmapFromFile Lib "gdiplus" ( _
        ByVal filename As LongPtr, _
        ByRef bitmap As LongPtr) As Long

    Private Declare PtrSafe Function GdipDisposeImage Lib "gdiplus" ( _
        ByVal image As LongPtr) As Long

    Private Declare PtrSafe Function GdipImageRotateFlip Lib "gdiplus" ( _
        ByVal image As LongPtr, _
        ByVal rfType As Long) As Long

    Private Declare PtrSafe Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
        ByVal bitmap As LongPtr, _
        ByRef hbmReturn As LongPtr, _
        ByVal background As Long) As Long

    Private Declare PtrSafe Function DeleteObject Lib "gdi32" ( _
        ByVal hObject As LongPtr) As Long

    Private Declare PtrSafe Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
        ByRef PicDesc As PICTDESC, _
        ByRef RefIID As GUID, _
        ByVal fPictureOwnsHandle As Long, _
        ByRef IPic As IPicture) As Long

#Else

    Private Declare Function GdiplusStartup Lib "gdiplus" ( _
        ByRef token As Long, _
        ByRef inputbuf As GdiplusStartupInput, _
        ByVal outputbuf As Long) As Long

    Private Declare Sub GdiplusShutdown Lib "gdiplus" ( _
        ByVal token As Long)

    Private Declare Function GdipCreateBitmapFromFile Lib "gdiplus" ( _
        ByVal filename As Long, _
        ByRef bitmap As Long) As Long

    Private Declare Function GdipDisposeImage Lib "gdiplus" ( _
        ByVal image As Long) As Long

    Private Declare Function GdipImageRotateFlip Lib "gdiplus" ( _
        ByVal image As Long, _
        ByVal rfType As Long) As Long

    Private Declare Function GdipCreateHBITMAPFromBitmap Lib "gdiplus" ( _
        ByVal bitmap As Long, _
        ByRef hbmReturn As Long) As Long

    Private Declare Function DeleteObject Lib "gdi32" ( _
        ByVal hObject As Long) As Long

    Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" ( _
        ByRef PicDesc As PICTDESC, _
        ByRef RefIID As GUID, _
        ByVal fPictureOwnsHandle As Long, _
        ByRef IPic As IPicture) As Long

#End If


'=========================================================
' FOCUS API DECLARATIONS
'=========================================================

#If VBA7 Then

    Private Declare PtrSafe Function SetFocusAPI Lib "user32" Alias "SetFocus" ( _
        ByVal hwnd As LongPtr) As LongPtr

#Else

    Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" ( _
        ByVal hwnd As Long) As Long

#End If


'=========================================================
' GDI+ TYPES
'=========================================================

Private Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As LongPtr
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type PICTDESC
    cbSizeOfStruct As Long
    picType As Long
#If VBA7 Then
    hImage As LongPtr
    xExt As Long
    yExt As Long
#Else
    hImage As Long
    xExt As Long
    yExt As Long
#End If
End Type


'=========================================================
' MODULE VARIABLES
'=========================================================

Private photoFiles() As String
Private photoCount As Long
Private currentPhotoIndex As Long

Private currentDateText As String
Private currentSourceName As String
Private currentSourceFolder As String
Private currentFolderCount As Long
Private currentRowNumber As Long

Private currentRotation As Long
Private rotationFile As String

' --- ZOOM ---
Private currentZoomFactor As Double
Private baseImageWidth As Single
Private baseImageHeight As Single
Private baseImageLeft As Single
Private baseImageTop As Single

Private Const ZOOM_STEP As Double = 1.25
Private Const ZOOM_MIN As Double = 0.3
Private Const ZOOM_MAX As Double = 5#

' --- PAN / DRAG (FIXED ABSOLUTE POSITIONING) ---
Private panOffsetX As Single
Private panOffsetY As Single
Private isDragging As Boolean
Private dragStartFormX As Single
Private dragStartFormY As Single

#If VBA7 Then
    Private gdiplusToken As LongPtr
#Else
    Private gdiplusToken As Long
#End If

Private gdiplusStarted As Boolean


'=========================================================
' HELPER PROCEDURES
'=========================================================

Private Sub ReturnFocusToExcel()
    On Error Resume Next
    #If VBA7 Then
        SetFocusAPI CLngPtr(Application.hwnd)
    #Else
        SetFocusAPI Application.hwnd
    #End If
End Sub


'=========================================================
' USERFORM INITIALIZE
'=========================================================

Private Sub UserForm_Initialize()
    On Error Resume Next

    Me.Caption = "Ekta Logbook - Photo Viewer"

    StartMouseWheel Me

    With lstPhotos
        .Clear
        .ColumnCount = 1
        .IntegralHeight = False
    End With

    cmdPrevious.Caption = ChrW$(9664)
    cmdNext.Caption = ChrW$(9654)

    With cmdPrevious
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 14
        .Font.Bold = False
        .ForeColor = RGB(150, 150, 150)
    End With

    With cmdNext
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 14
        .Font.Bold = False
        .ForeColor = RGB(150, 150, 150)
    End With

    cmdRotateLeft.Caption = ChrW$(8634)
    cmdRotateRight.Caption = ChrW$(8635)

    With cmdRotateLeft
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 14
        .Font.Bold = False
        .ForeColor = RGB(150, 150, 150)
    End With

    With cmdRotateRight
        .Font.Name = "Segoe UI Symbol"
        .Font.Size = 14
        .Font.Bold = False
        .ForeColor = RGB(150, 150, 150)
    End With

    cmdOpenFolder.Caption = "Open Folder"
    cmdZoomIn.Caption = "Zoom In (+)"
    cmdZoomOut.Caption = "Zoom Out (-)"

    imgphoto.PictureSizeMode = 3

    baseImageWidth = imgphoto.Width
    baseImageHeight = imgphoto.Height
    baseImageLeft = imgphoto.Left
    baseImageTop = imgphoto.Top
    currentZoomFactor = 1#

    StartGDIPlus

    lblInfo.Caption = "No photo selected."
End Sub


'=========================================================
' LOAD PHOTOS
'=========================================================

Public Sub LoadPhotos( _
    ByVal dateText As String, _
    ByVal sourceName As String, _
    ByVal matchedFolders As Collection, _
    ByVal rowNumber As Long)

    On Error GoTo LoadError

    currentDateText = dateText
    currentSourceName = sourceName
    currentSourceFolder = CStr(matchedFolders(1))
    currentFolderCount = matchedFolders.Count
    currentRowNumber = rowNumber

    rotationFile = Environ$("APPDATA") & "\EktaLogbookPhotoRotations.txt"
    currentPhotoIndex = 0
    currentRotation = 0

    lstPhotos.Clear
    Erase photoFiles
    photoCount = 0

    Dim oneFolder As Variant
    For Each oneFolder In matchedFolders
        GetPhotoFiles CStr(oneFolder), photoFiles, photoCount
    Next oneFolder

    If photoCount = 0 Then
        imgphoto.Picture = Nothing
        lblInfo.Caption = "No photos found." & vbCrLf & sourceName
        Exit Sub
    End If

    Dim i As Long
    For i = 1 To photoCount
        lstPhotos.AddItem GetFileNameOnly(photoFiles(i))
    Next i

    currentPhotoIndex = 1
    lstPhotos.ListIndex = 0
    DisplayCurrentPhoto
    Exit Sub

LoadError:
    MsgBox "Unable to load photos." & vbCrLf & vbCrLf & Err.Description, vbCritical, "Photo Viewer Error"
End Sub


'=========================================================
' DISPLAY PHOTO
'=========================================================

Private Sub DisplayCurrentPhoto()
    On Error GoTo DisplayError

    If photoCount = 0 Then Exit Sub
    If currentPhotoIndex < 1 Or currentPhotoIndex > photoCount Then Exit Sub

    currentRotation = LoadRotation(photoFiles(currentPhotoIndex))
    SetImageFromFile imgphoto, photoFiles(currentPhotoIndex)

    ResetZoom

    If lstPhotos.ListCount >= currentPhotoIndex Then
        lstPhotos.ListIndex = currentPhotoIndex - 1
    End If

    UpdateInfo
    Exit Sub

DisplayError:
    imgphoto.Picture = Nothing
    lblInfo.Caption = "Unable to display photo:" & vbCrLf & GetFileNameOnly(photoFiles(currentPhotoIndex)) & vbCrLf & vbCrLf & Err.Description
End Sub


'=========================================================
' CONTROLS & NAVIGATION
'=========================================================

Private Sub lstPhotos_Click()
    On Error GoTo ErrorHandler
    If lstPhotos.ListIndex < 0 Then Exit Sub
    currentPhotoIndex = lstPhotos.ListIndex + 1
    DisplayCurrentPhoto
    Exit Sub
ErrorHandler:
    MsgBox "Unable to open selected photo." & vbCrLf & Err.Description, vbExclamation, "Photo Viewer"
End Sub

Private Sub lstPhotos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If lstPhotos.ListIndex >= 0 Then
        currentPhotoIndex = lstPhotos.ListIndex + 1
        DisplayCurrentPhoto
    End If
End Sub

Private Sub cmdPrevious_Click()
    If photoCount = 0 Then Exit Sub
    currentPhotoIndex = currentPhotoIndex - 1
    If currentPhotoIndex < 1 Then currentPhotoIndex = photoCount
    DisplayCurrentPhoto
    ReturnFocusToExcel
End Sub

Private Sub cmdNext_Click()
    If photoCount = 0 Then Exit Sub
    currentPhotoIndex = currentPhotoIndex + 1
    If currentPhotoIndex > photoCount Then currentPhotoIndex = 1
    DisplayCurrentPhoto
    ReturnFocusToExcel
End Sub

Private Sub cmdRotateLeft_Click()
    If photoCount = 0 Or currentPhotoIndex = 0 Then Exit Sub
    currentRotation = currentRotation - 90
    If currentRotation < 0 Then currentRotation = currentRotation + 360
    SaveRotation photoFiles(currentPhotoIndex), currentRotation
    ApplyRotationToCurrentPhoto
    ReturnFocusToExcel
End Sub

Private Sub cmdRotateRight_Click()
    If photoCount = 0 Or currentPhotoIndex = 0 Then Exit Sub
    currentRotation = currentRotation + 90
    If currentRotation >= 360 Then currentRotation = currentRotation - 360
    SaveRotation photoFiles(currentPhotoIndex), currentRotation
    ApplyRotationToCurrentPhoto
    ReturnFocusToExcel
End Sub


'=========================================================
' ZOOM MANAGEMENT
'=========================================================

Private Sub cmdZoomIn_Click()
    ChangeZoom True
    ReturnFocusToExcel
End Sub

Private Sub cmdZoomOut_Click()
    ChangeZoom False
    ReturnFocusToExcel
End Sub

Public Sub MouseWheelZoom(ByVal wheelDelta As Long)
    On Error Resume Next
    If photoCount = 0 Or wheelDelta = 0 Then Exit Sub
    ChangeZoom (wheelDelta > 0)
End Sub

Private Sub ChangeZoom(ByVal zoomIn As Boolean)
    If photoCount = 0 Then Exit Sub

    If zoomIn Then
        currentZoomFactor = currentZoomFactor * ZOOM_STEP
    Else
        currentZoomFactor = currentZoomFactor / ZOOM_STEP
    End If

    If currentZoomFactor > ZOOM_MAX Then currentZoomFactor = ZOOM_MAX
    If currentZoomFactor < ZOOM_MIN Then currentZoomFactor = ZOOM_MIN

    ApplyZoom
End Sub

Private Sub ResetZoom()
    currentZoomFactor = 1#
    panOffsetX = 0
    panOffsetY = 0
    ApplyZoom
End Sub

Private Sub ApplyZoom()
    Dim newWidth As Single, newHeight As Single
    Dim centerX As Single, centerY As Single

    centerX = baseImageLeft + (baseImageWidth / 2)
    centerY = baseImageTop + (baseImageHeight / 2)

    newWidth = baseImageWidth * currentZoomFactor
    newHeight = baseImageHeight * currentZoomFactor

    imgphoto.Width = newWidth
    imgphoto.Height = newHeight

    imgphoto.Left = centerX - (newWidth / 2) + panOffsetX
    imgphoto.Top = centerY - (newHeight / 2) + panOffsetY

    UpdateInfo
End Sub




'=========================================================
' DRAG TO PAN (CONTAINER ABSOLUTE POSITIONING)
'=========================================================

Private Sub imgPhoto_MouseDown( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    ' Enable dragging only when left click is held and image is zoomed in
    If Button = 1 And currentZoomFactor > 1# Then
        isDragging = True
        
        ' Store the initial mouse anchor point relative to the UserForm
        dragStartFormX = imgphoto.Left + X
        dragStartFormY = imgphoto.Top + Y
    End If
End Sub

Private Sub imgPhoto_MouseMove( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    If Not isDragging Then Exit Sub

    ' Update image position based on displacement from the initial click
    Dim newLeft As Single
    Dim newTop As Single

    newLeft = imgphoto.Left + X - (dragStartFormX - imgphoto.Left)
    newTop = imgphoto.Top + Y - (dragStartFormY - imgphoto.Top)

    ' Directly update control coordinates
    imgphoto.Left = imgphoto.Left + (X - (dragStartFormX - imgphoto.Left))
    imgphoto.Top = imgphoto.Top + (Y - (dragStartFormY - imgphoto.Top))

    ' Update anchor point to current position for the next move delta
    dragStartFormX = imgphoto.Left + X
    dragStartFormY = imgphoto.Top + Y

    ' Track offset relative to base center point for zoom scaling
    Dim centerX As Single
    Dim centerY As Single

    centerX = baseImageLeft + (baseImageWidth / 2)
    centerY = baseImageTop + (baseImageHeight / 2)

    panOffsetX = imgphoto.Left - (centerX - (imgphoto.Width / 2))
    panOffsetY = imgphoto.Top - (centerY - (imgphoto.Height / 2))
End Sub

Private Sub imgPhoto_MouseUp( _
    ByVal Button As Integer, _
    ByVal Shift As Integer, _
    ByVal X As Single, _
    ByVal Y As Single)

    isDragging = False
End Sub


'=========================================================
' ROTATION IMPLEMENTATION
'=========================================================

Private Sub ApplyRotationToCurrentPhoto()
    On Error GoTo RotationError

    If Not gdiplusStarted Then StartGDIPlus

    Dim rotatedPicture As IPicture
    Set rotatedPicture = GetRotatedPicture(photoFiles(currentPhotoIndex), currentRotation)

    If rotatedPicture Is Nothing Then
        MsgBox "Unable to rotate this image.", vbExclamation, "Rotation Error"
        Exit Sub
    End If

    Set imgphoto.Picture = rotatedPicture
    ResetZoom
    UpdateInfo
    Exit Sub

RotationError:
    MsgBox "Rotation failed." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "Rotation Error"
End Sub

Private Function GetRotatedPicture(ByVal filePath As String, ByVal rotation As Long) As IPicture
    On Error GoTo RotationError

    #If VBA7 Then
        Dim bitmap As LongPtr, hBitmap As LongPtr
    #Else
        Dim bitmap As Long, hBitmap As Long
    #End If

    Dim result As Long
    result = GdipCreateBitmapFromFile(StrPtr(filePath), bitmap)
    If result <> 0 Then Exit Function

    Select Case rotation
        Case 0: result = 0
        Case 90: result = GdipImageRotateFlip(bitmap, 1)
        Case 180: result = GdipImageRotateFlip(bitmap, 2)
        Case 270: result = GdipImageRotateFlip(bitmap, 3)
    End Select

    If result <> 0 Then
        GdipDisposeImage bitmap
        Exit Function
    End If

    result = GdipCreateHBITMAPFromBitmap(bitmap, hBitmap, 0)
    If result <> 0 Then
        GdipDisposeImage bitmap
        Exit Function
    End If

    Set GetRotatedPicture = HBitmapToPicture(hBitmap)
    GdipDisposeImage bitmap
    Exit Function

RotationError:
    Set GetRotatedPicture = Nothing
End Function

Private Function HBitmapToPicture(ByVal hBitmap As LongPtr) As IPicture
    On Error GoTo ErrorHandler

    Dim Pic As IPicture
    Dim PicDesc As PICTDESC
    Dim IID As GUID
    Dim result As Long

    With IID
        .Data1 = &H7BF80980
        .Data2 = &HBF32
        .Data3 = &H101A
        .Data4(0) = &H8B: .Data4(1) = &HBB: .Data4(2) = &H0: .Data4(3) = &HAA
        .Data4(4) = &H0: .Data4(5) = &H30: .Data4(6) = &HC: .Data4(7) = &HAB
    End With

    With PicDesc
        .cbSizeOfStruct = LenB(PicDesc)
        .picType = 1
        .hImage = hBitmap
        .xExt = 0
        .yExt = 0
    End With

    result = OleCreatePictureIndirect(PicDesc, IID, 1, Pic)

    If result = 0 Then
        Set HBitmapToPicture = Pic
    Else
        DeleteObject hBitmap
        Set HBitmapToPicture = Nothing
    End If
    Exit Function

ErrorHandler:
    DeleteObject hBitmap
    Set HBitmapToPicture = Nothing
End Function


'=========================================================
' IMAGE HELPERS & GDI+
'=========================================================

Private Sub SetImageFromFile(ByVal imageControl As MSForms.image, ByVal filePath As String)
    On Error GoTo ErrorHandler
    If Len(Dir(filePath)) = 0 Then
        imageControl.Picture = Nothing
        Exit Sub
    End If

    Dim pictureObject As IPicture
    Set pictureObject = GetPictureFromFile(filePath)

    If pictureObject Is Nothing Then
        imageControl.Picture = LoadPicture(filePath)
    Else
        Set imageControl.Picture = pictureObject
    End If
    Exit Sub

ErrorHandler:
    On Error Resume Next
    imageControl.Picture = LoadPicture(filePath)
End Sub

Private Function GetPictureFromFile(ByVal filePath As String) As IPicture
    On Error GoTo ErrorHandler

    #If VBA7 Then
        Dim bitmap As LongPtr, hBitmap As LongPtr
    #Else
        Dim bitmap As Long, hBitmap As Long
    #End If

    Dim result As Long
    result = GdipCreateBitmapFromFile(StrPtr(filePath), bitmap)
    If result <> 0 Then Exit Function

    result = GdipCreateHBITMAPFromBitmap(bitmap, hBitmap, 0)
    GdipDisposeImage bitmap
    If result <> 0 Then Exit Function

    Set GetPictureFromFile = HBitmapToPicture(hBitmap)
    Exit Function

ErrorHandler:
    Set GetPictureFromFile = Nothing
End Function

Private Sub StartGDIPlus()
    On Error GoTo ErrorHandler
    If gdiplusStarted Then Exit Sub

    Dim startupInput As GdiplusStartupInput
    Dim result As Long

    startupInput.GdiplusVersion = 1
    result = GdiplusStartup(gdiplusToken, startupInput, 0)
    gdiplusStarted = (result = 0)
    Exit Sub

ErrorHandler:
    gdiplusStarted = False
End Sub

Private Sub StopGDIPlus()
    On Error Resume Next
    If gdiplusStarted Then
        GdiplusShutdown gdiplusToken
        gdiplusStarted = False
    End If
End Sub


'=========================================================
' LABELS & UTILITIES
'=========================================================

Private Sub UpdateInfo()
    On Error Resume Next

    Dim rotationText As String
    Select Case currentRotation
        Case 0: rotationText = "0°"
        Case 90: rotationText = "90°"
        Case 180: rotationText = "180°"
        Case 270: rotationText = "270°"
    End Select

    Dim folderNote As String
    If currentFolderCount > 1 Then
        folderNote = "  |  " & currentFolderCount & " folders merged"
    End If

    lblInfo.Caption = _
        "Photo " & currentPhotoIndex & " of " & photoCount & _
        "     |     Rotation: " & rotationText & _
        "     |     Zoom: " & Format(currentZoomFactor * 100, "0") & "%" & vbCrLf & _
        currentDateText & "  |  " & currentSourceName & folderNote
End Sub

Private Sub cmdOpenFolder_Click()
    On Error GoTo ErrorHandler
    If Len(currentSourceFolder) = 0 Then Exit Sub

    OpenFolder currentSourceFolder

    If currentFolderCount > 1 Then
        MsgBox "This site has " & currentFolderCount & " folders for this date." & vbCrLf & _
               "Opening the first one - check the date folder for the others.", vbInformation, "Multiple Folders"
    End If
    Exit Sub

ErrorHandler:
    MsgBox "Unable to open folder." & vbCrLf & Err.Description, vbExclamation, "Open Folder"
End Sub

Private Function GetFileNameOnly(ByVal filePath As String) As String
    Dim p As Long
    p = InStrRev(filePath, "\")
    If p > 0 Then
        GetFileNameOnly = Mid$(filePath, p + 1)
    Else
        GetFileNameOnly = filePath
    End If
End Function


'=========================================================
' ROTATION PERSISTENCE
'=========================================================

Private Function LoadRotation(ByVal filePath As String) As Long
    On Error GoTo ErrorHandler
    Dim fileNo As Integer, oneLine As String, parts() As String

    LoadRotation = 0
    If Len(Dir(rotationFile)) = 0 Then Exit Function

    fileNo = FreeFile
    Open rotationFile For Input As #fileNo

    Do While Not EOF(fileNo)
        Line Input #fileNo, oneLine
        If InStr(1, oneLine, "|") > 0 Then
            parts = Split(oneLine, "|", 2)
            If StrComp(parts(0), filePath, vbTextCompare) = 0 Then
                LoadRotation = CLng(parts(1))
                Exit Do
            End If
        End If
    Loop

    Close #fileNo
    Exit Function

ErrorHandler:
    On Error Resume Next
    Close #fileNo
    LoadRotation = 0
End Function

Private Sub SaveRotation(ByVal filePath As String, ByVal rotation As Long)
    On Error GoTo ErrorHandler
    Dim fileNo As Integer, tempFile As String, oneLine As String
    Dim parts() As String, found As Boolean, allText As String

    If Len(rotationFile) = 0 Then
        rotationFile = Environ$("APPDATA") & "\EktaLogbookPhotoRotations.txt"
    End If

    rotation = rotation Mod 360
    If rotation < 0 Then rotation = rotation + 360

    If Len(Dir(rotationFile)) > 0 Then
        fileNo = FreeFile
        Open rotationFile For Input As #fileNo

        Do While Not EOF(fileNo)
            Line Input #fileNo, oneLine
            If InStr(1, oneLine, "|") > 0 Then
                parts = Split(oneLine, "|", 2)
                If StrComp(parts(0), filePath, vbTextCompare) = 0 Then
                    allText = allText & filePath & "|" & rotation & vbCrLf
                    found = True
                Else
                    allText = allText & oneLine & vbCrLf
                End If
            End If
        Loop

        Close #fileNo
    End If

    If Not found Then
        allText = allText & filePath & "|" & rotation & vbCrLf
    End If

    tempFile = rotationFile & ".tmp"
    fileNo = FreeFile
    Open tempFile For Output As #fileNo
    Print #fileNo, allText
    Close #fileNo

    If Len(Dir(rotationFile)) > 0 Then Kill rotationFile
    Name tempFile As rotationFile
    Exit Sub

ErrorHandler:
    On Error Resume Next
    Close #fileNo
End Sub


'=========================================================
' TERMINATION CLEANUP
'=========================================================

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    StopMouseWheel
    StopGDIPlus
End Sub

Private Sub UserForm_Terminate()
    StopMouseWheel
    StopGDIPlus
End Sub

