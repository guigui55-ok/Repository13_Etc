Attribute VB_Name = "Module7"
Option Explicit
 
Private Declare Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function CallWindowProc Lib "user32.dll" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetClipboardViewer Lib "user32.dll" (ByVal hWndNewViewer As Long) As Long
Private Declare Function ChangeClipboardChain Lib "user32.dll" (ByVal hWndRemove As Long, ByVal hWndNewNext As Long) As Long
Private Declare Function IsClipboardFormatAvailable Lib "user32.dll" (ByVal format As Long) As Long
 
Private Const GWL_WNDPROC As Long = -4
 
Private Const WM_DRAWCLIPBOARD As Long = &H308
Private Const WM_CHANGECBCHAIN As Long = &H30D
Private Const WM_NCHITTEST As Long = &H84
 
Private Const CF_BITMAP As Long = 2
 
Private Const ROW_HEIGHT As Double = 13.5
 
Private hWndForm As Long
Private wpWindowProcOrg As Long
Private hWndNextViewer As Long
Private firstFired As Boolean
 
Public Sub catchClipboard()
    hWndForm = FindWindow("ThunderDFrame", UserForm1.Caption)
    wpWindowProcOrg = SetWindowLong(hWndForm, GWL_WNDPROC, AddressOf WindowProc)
    firstFired = False
    hWndNextViewer = SetClipboardViewer(hWndForm)
End Sub
 
Public Sub releaseClipboard()
    Call ChangeClipboardChain(hWndForm, hWndNextViewer)
    Call SetWindowLong(hWndForm, GWL_WNDPROC, wpWindowProcOrg)
End Sub
 
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Select Case uMsg
        Case WM_DRAWCLIPBOARD
            If Not firstFired Then
                firstFired = True
            ElseIf IsClipboardFormatAvailable(CF_BITMAP) <> 0 Then
                pasteToSheet
            End If
            If hWndNextViewer <> 0 Then
                Call SendMessage(hWndNextViewer, uMsg, wParam, lParam)
            End If
            WindowProc = 0
        Case WM_CHANGECBCHAIN
            If wParam = hWndNextViewer Then
                hWndNextViewer = lParam
            ElseIf hWndNextViewer <> 0 Then
                Call SendMessage(hWndNextViewer, uMsg, wParam, lParam)
            End If
            WindowProc = 0
        Case WM_NCHITTEST
            WindowProc = 0
        Case Else
            WindowProc = CallWindowProc(wpWindowProcOrg, hWndForm, uMsg, wParam, lParam)
    End Select
End Function
 
Public Sub pasteToSheet()
    Dim rowIdx As Integer
    
    With Sheet1
        If .Shapes.Count > 0 Then
            With .Shapes(.Shapes.Count)
                rowIdx = (.Top + .Height) / ROW_HEIGHT + 4
            End With
        Else
            rowIdx = 1
        End If
        .Cells(rowIdx, 1).PasteSpecial
    End With
End Sub
