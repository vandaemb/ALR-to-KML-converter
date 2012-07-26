Attribute VB_Name = "modDragDrop"
Private Const WM_DROPFILES = &H233
Private Const GWL_WNDPROC = (-4)
Private Declare Sub DragAcceptFiles Lib "shell32.dll" (ByVal hwnd As Long, ByVal fAccept As Long)
Private Declare Function DragQueryFile Lib "shell32.dll" Alias "DragQueryFileA" (ByVal HDROP As Long, ByVal UINT As Long, ByVal lpStr As String, ByVal ch As Long) As Long
Private Declare Sub DragFinish Lib "shell32.dll" (ByVal HDROP As Long)
Private PrevProc As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Function HookForm(ByVal hwnd As Long)
    PrevProc = SetWindowLong(hwnd, GWL_WNDPROC, AddressOf WindowProc)
End Function
Private Function UnHookForm(ByVal hwnd As Long)
    If PrevProc <> 0 Then
        SetWindowLong hwnd, GWL_WNDPROC, PrevProc
        PrevProc = 0
    End If
End Function

Private Function WindowProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If uMsg = WM_DROPFILES Then
        Dropped wParam
    End If
    WindowProc = CallWindowProc(PrevProc, hwnd, uMsg, wParam, lParam)
End Function

Public Sub EnableDragDrop(ByVal hwnd As Long)
    DragAcceptFiles hwnd, 1
    HookForm (hwnd)
End Sub

Public Sub DisableDragDrop(ByVal hwnd As Long)
    DragAcceptFiles hwnd, 0
    UnHookForm hwnd
End Sub

Public Sub Dropped(ByVal HDROP As Long)
    Dim strFilename As String * 511
    Call DragQueryFile(HDROP, 0, strFilename, 511)
    
    frmALRtoKML.GotADrop (strFilename)
    Call DragQueryFile(HDROP, 2, strFilename, 511)

    
End Sub
