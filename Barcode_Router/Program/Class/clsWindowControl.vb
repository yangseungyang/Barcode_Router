Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsWindowControl
'Function   : 다른 프로그램 창 / Router 창을 맨 앞으로 올린다 (Win32)
'Framwork   : VB2022 Framework4.8
'Memo       : - 분기 대상 프로그램은 COM 설정의 프로세스명(예 : AIEC_HIPOT, FAM3_MAIN_HIPOT) 또는 exe 전체 경로로 찾는다
'             - 숨김 상태이면 Process.MainWindowHandle 이 0 이므로, EnumWindows 로 그 프로세스의 최상위 창을 찾는다
'             - Router 가 앞에 있을 때 바코드를 보내므로 Router 는 다른 창을 앞으로 보낼 수 있다
'             - 검사가 끝나면 검사 프로그램이 AllowSetForegroundWindow 로 권한을 넘긴 뒤 END 를 보내므로,
'               END 를 받은 Router 는 자기 창을 앞으로 올릴 수 있다 (구조설계 P-08)
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-7항, 7-9항
'============================================================================================================================
Imports System.Runtime.InteropServices

Public Class clsWindowControl

    Private Const SW_SHOW As Integer = 5
    Private Const SW_RESTORE As Integer = 9

    Private Delegate Function EnumWindowsProc(ByVal hWnd As IntPtr, ByVal lParam As IntPtr) As Boolean

    <DllImport("user32.dll")>
    Private Shared Function EnumWindows(ByVal lpEnumFunc As EnumWindowsProc, ByVal lParam As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindowThreadProcessId(ByVal hWnd As IntPtr, ByRef lpdwProcessId As Integer) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Shared Function GetWindow(ByVal hWnd As IntPtr, ByVal uCmd As Integer) As IntPtr
    End Function
    <DllImport("user32.dll", CharSet:=CharSet.Unicode)>
    Private Shared Function GetWindowTextLength(ByVal hWnd As IntPtr) As Integer
    End Function
    <DllImport("user32.dll")>
    Private Shared Function ShowWindow(ByVal hWnd As IntPtr, ByVal nCmdShow As Integer) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function IsIconic(ByVal hWnd As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function SetForegroundWindow(ByVal hWnd As IntPtr) As Boolean
    End Function
    <DllImport("user32.dll")>
    Private Shared Function BringWindowToTop(ByVal hWnd As IntPtr) As Boolean
    End Function

    Public Shared Function BringProcessToFront(ByVal processName As String) As Boolean
        '==========================================================================================
        'Function : 프로세스명으로 창을 찾아 숨김 / 최소화를 풀고 맨 앞으로
        'Return   : 창을 찾았으면 True
        '==========================================================================================
        If processName.Trim = "" Then Return False
        Dim procs() As Process = FindProcesses(processName)
        If procs.Length = 0 Then Return False

        Dim pid As Integer = procs(0).Id
        Dim hWnd As IntPtr = procs(0).MainWindowHandle

        If hWnd = IntPtr.Zero Then
            '숨김 상태 : 제목이 있는 최상위 창(소유자 없음)을 찾는다
            EnumWindows(Function(h As IntPtr, l As IntPtr) As Boolean
                            Dim wpid As Integer = 0
                            GetWindowThreadProcessId(h, wpid)
                            If wpid = pid AndAlso GetWindow(h, 4) = IntPtr.Zero AndAlso GetWindowTextLength(h) > 0 Then   '4 = GW_OWNER
                                hWnd = h
                                Return False
                            End If
                            Return True
                        End Function, IntPtr.Zero)
        End If
        If hWnd = IntPtr.Zero Then Return False

        Return ShowFront(hWnd)
    End Function

    Public Shared Function ShowFront(ByVal hWnd As IntPtr) As Boolean
        If IsIconic(hWnd) Then
            ShowWindow(hWnd, SW_RESTORE)
        Else
            ShowWindow(hWnd, SW_SHOW)
        End If
        BringWindowToTop(hWnd)
        Return SetForegroundWindow(hWnd)
    End Function

    Public Shared Function IsProcessRunning(ByVal processName As String) As Boolean
        If processName.Trim = "" Then Return False
        Return FindProcesses(processName).Length > 0
    End Function

    Public Shared Function FindProcesses(ByVal nameOrPath As String) As Process()
        '==========================================================================================
        'Function : 프로세스명 또는 exe 전체 경로로 프로세스 찾기 (2026-10-28)
        'Remarks  : \ 가 있으면 exe 경로로 본다 - 이름이 같은 exe 가 둘이어도 구분 (CPU검사장치 A8 / FAM3 'Inspection System')
        '           경로를 읽지 못하는 프로세스는 건너뛴다
        '==========================================================================================
        Dim key As String = nameOrPath.Trim
        If key = "" Then Return New Process() {}
        If key.IndexOf("\"c) < 0 Then Return Process.GetProcessesByName(key.Replace(".exe", "").Replace(".EXE", ""))

        Dim lst As New List(Of Process)
        For Each p As Process In Process.GetProcessesByName(IO.Path.GetFileNameWithoutExtension(key))
            Try
                If String.Equals(IO.Path.GetFullPath(p.MainModule.FileName), IO.Path.GetFullPath(key), StringComparison.OrdinalIgnoreCase) Then lst.Add(p)
            Catch
            End Try
        Next
        Return lst.ToArray
    End Function

End Class
