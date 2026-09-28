Attribute VB_Name = "modMouseWheel"
Option Explicit

'=========================================================
' MOUSE-WHEEL ZOOM SUPPORT
' Excel 64-bit / VBA7
'
' Standard module name:
' modMouseWheel
'=========================================================

#If VBA7 Then

    Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
        ByVal idHook As Long, _
        ByVal lpfn As LongPtr, _
        ByVal hMod As LongPtr, _
        ByVal dwThreadId As Long) As LongPtr

    Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" ( _
        ByVal hHook As LongPtr) As Long

    Private Declare PtrSafe Function CallNextHookEx Lib "user32" ( _
        ByVal hHook As LongPtr, _
        ByVal nCode As Long, _
        ByVal wParam As LongPtr, _
        ByVal lParam As LongPtr) As LongPtr

    Private Declare PtrSafe Function GetCursorPos Lib "user32" ( _
        ByRef lpPoint As POINTAPI) As Long

    Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function GetWindowRect Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByRef lpRect As RECT) As Long

    Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, _
        ByRef Source As Any, _
        ByVal Length As LongPtr)

    Private gHook As LongPtr

#Else

    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" ( _
        ByVal idHook As Long, _
        ByVal lpfn As Long, _
        ByVal hMod As Long, _
        ByVal dwThreadId As Long) As Long

    Private Declare Function UnhookWindowsHookEx Lib "user32" ( _
        ByVal hHook As Long) As Long

    Private Declare Function CallNextHookEx Lib "user32" ( _
        ByVal hHook As Long, _
        ByVal nCode As Long, _
        ByVal wParam As Long, _
        ByVal lParam As Long) As Long

    Private Declare Function GetCursorPos Lib "user32" ( _
        ByRef lpPoint As POINTAPI) As Long

    Private Declare Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

    Private Declare Function GetWindowRect Lib "user32" ( _
        ByVal hwnd As Long, _
        ByRef lpRect As RECT) As Long

    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
        ByRef Destination As Any, _
        ByRef Source As Any, _
        ByVal Length As Long)

    Private gHook As Long

#End If

Private Const WH_MOUSE_LL As Long = 14
Private Const WM_MOUSEWHEEL As Long = &H20A
Private Const HC_ACTION As Long = 0

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mouseData As Long
    flags As Long
    time As Long

#If VBA7 Then
    dwExtraInfo As LongPtr
#Else
    dwExtraInfo As Long
#End If

End Type

Private gViewer As frmPhotoViewer
Private gViewerCaption As String

'=========================================================
' START MOUSE-WHEEL HOOK
'=========================================================

Public Sub StartMouseWheel(ByVal viewer As frmPhotoViewer)

    On Error Resume Next

    StopMouseWheel

    Set gViewer = viewer
    gViewerCaption = viewer.Caption

    gHook = SetWindowsHookEx( _
                WH_MOUSE_LL, _
                AddressOf MouseWheelProc, _
                0, _
                0)

End Sub

'=========================================================
' STOP MOUSE-WHEEL HOOK
'=========================================================

Public Sub StopMouseWheel()

    On Error Resume Next

    If gHook <> 0 Then
        UnhookWindowsHookEx gHook
    End If

    gHook = 0

    Set gViewer = Nothing
    gViewerCaption = vbNullString

End Sub

'=========================================================
' MOUSE-WHEEL CALLBACK
'=========================================================

#If VBA7 Then

Public Function MouseWheelProc( _
    ByVal nCode As Long, _
    ByVal wParam As LongPtr, _
    ByVal lParam As LongPtr) As LongPtr

#Else

Public Function MouseWheelProc( _
    ByVal nCode As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long) As Long

#End If

    On Error GoTo ContinueHook

    If nCode = HC_ACTION Then

        If wParam = WM_MOUSEWHEEL Then

            If FormIsUnderMouse() Then

                Dim hookInfo As MSLLHOOKSTRUCT
                Dim wheelDelta As Long

                CopyMemory _
                    hookInfo, _
                    ByVal lParam, _
                    LenB(hookInfo)

                wheelDelta = HighWordSigned(hookInfo.mouseData)

                If wheelDelta <> 0 Then

                    If Not gViewer Is Nothing Then
                        gViewer.MouseWheelZoom wheelDelta
                    End If

                End If

            End If

        End If

    End If

ContinueHook:

    MouseWheelProc = CallNextHookEx( _
                        gHook, _
                        nCode, _
                        wParam, _
                        lParam)

End Function

'=========================================================
' CHECK WHETHER MOUSE IS OVER THE PHOTO VIEWER
'=========================================================

Private Function FormIsUnderMouse() As Boolean

    On Error GoTo NotUnderForm

    If gViewer Is Nothing Then Exit Function
    If Len(gViewerCaption) = 0 Then Exit Function

#If VBA7 Then

    Dim hwndForm As LongPtr

#Else

    Dim hwndForm As Long

#End If

    Dim pt As POINTAPI
    Dim rc As RECT

    hwndForm = FindWindowA( _
                    vbNullString, _
                    gViewerCaption)

    If hwndForm = 0 Then Exit Function

    If GetCursorPos(pt) = 0 Then Exit Function

    If GetWindowRect(hwndForm, rc) = 0 Then Exit Function

    FormIsUnderMouse = _
        (pt.X >= rc.Left And _
         pt.X <= rc.Right And _
         pt.Y >= rc.Top And _
         pt.Y <= rc.Bottom)

    Exit Function

NotUnderForm:

    FormIsUnderMouse = False

End Function

'=========================================================
' GET SIGNED HIGH WORD FROM MOUSE DATA
'=========================================================

Private Function HighWordSigned(ByVal value As Long) As Long

    Dim result As Long

    result = (value And &HFFFF0000) \ &H10000

    If result >= 32768 Then
        result = result - 65536
    End If

    HighWordSigned = result

End Function

