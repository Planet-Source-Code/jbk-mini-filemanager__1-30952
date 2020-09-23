Attribute VB_Name = "mdMisc"

Option Explicit
Const SPACE = 5
Const BAR_WIDTH = 50
Public Const HWND_TOPMOST = -1&
Public Const HWND_NOTOPMOST = -2&
Public Const SWP_NOSIZE = &H1&
Public Const SWP_NOMOVE = &H2&
Public Const SWP_NOACTIVATE = &H10&
Public Const SWP_SHOWWINDOW = &H40&
Public Const THREAD_BASE_PRIORITY_MAX = 2
Public Const HIGH_PRIORITY_CLASS = &H80
Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type
Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function GetCurrentThread Lib "kernel32" () As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Public Declare Sub SetWindowPos Lib "User32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public GraphPoints(0 To 99) As Long
Sub DrawUsage(lUsage As Long, picPercent As PictureBox, picGraph As PictureBox)
    Dim Cnt As Long
    picPercent.ScaleMode = vbPixels
    For Cnt = 0 To 10
        picPercent.Line (SPACE, SPACE + Cnt * 3)-(SPACE + BAR_WIDTH, SPACE + Cnt * 3 + 1), IIf(lUsage >= 100 - Cnt * 10 And lUsage <> 0, &HFFFFFF, &H808080), BF
    Next Cnt
    ShiftPoints
    GraphPoints(UBound(GraphPoints)) = lUsage
    picGraph.Cls
    For Cnt = LBound(GraphPoints) To UBound(GraphPoints) - 1
        picGraph.Line (Cnt, 100 - GraphPoints(Cnt))-(Cnt + 1, 100 - GraphPoints(Cnt + 1)), &HFFFFFF
    Next Cnt
End Sub
Sub DrawUsage2(lUsage As Long, picPercent As PictureBox)
    Dim Cnt As Long
    picPercent.ScaleMode = vbPixels
    For Cnt = 0 To 99
        picPercent.Line (SPACE, SPACE + Cnt * 3)-(SPACE + BAR_WIDTH, SPACE + Cnt * 3 + 1), IIf(lUsage >= 100 - Cnt And lUsage <> 0, &HFFFFFF, &H808080), BF
    Next Cnt
End Sub
Sub DrawUsage3(lUsage As Long, picPercent As PictureBox)
    Dim Cnt As Long
    picPercent.ScaleMode = vbPixels
    For Cnt = 0 To 10
        picPercent.Line (SPACE, SPACE + Cnt * 3)-(SPACE + BAR_WIDTH, SPACE + Cnt * 3 + 1), IIf(lUsage >= 100 - Cnt * 10 And lUsage <> 0, &HFFFFFF, &H808080), BF
    Next Cnt
End Sub
Sub DrawUsage4(lUsage As Long, picGraph As PictureBox)
    Dim Cnt As Long
    ShiftPoints
    GraphPoints(UBound(GraphPoints)) = lUsage
    picGraph.Cls
    For Cnt = LBound(GraphPoints) To UBound(GraphPoints) - 1
        picGraph.Line (Cnt, 100 - GraphPoints(Cnt))-(Cnt + 1, 100 - GraphPoints(Cnt + 1)), &HFFFFFF
    Next Cnt
End Sub
'Shift all the points from the graph one place to the left
Sub ShiftPoints()
    Dim Cnt As Long
    For Cnt = LBound(GraphPoints) To UBound(GraphPoints) - 1
        GraphPoints(Cnt) = GraphPoints(Cnt + 1)
    Next Cnt
End Sub
'return True is the OS is WindowsNT3.5(1), NT4.0, 2000 or XP
Public Function IsWinNT() As Boolean
    Dim OSInfo As OSVERSIONINFO
    OSInfo.dwOSVersionInfoSize = Len(OSInfo)
    'retrieve OS version info
    GetVersionEx OSInfo
    'if we're on NT, return True
    IsWinNT = (OSInfo.dwPlatformId = 2)
End Function
