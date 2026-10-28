Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterLog
'Function   : 바코드 분배 프로그램 System Log 기록
'Framwork   : VB2022 Framework4.8
'Memo       : 기존 검사 프로그램(A8 / FAM3) 의 System_Log 와 같은 위치 규칙을 쓴다
'             {프로그램 폴더}\System_Log\yyyy-MM\yyyy-MM-dd.txt
'             기록 내용 : 수신 바코드, MSCODE, 맞은 규칙, 전송 COM, END 수신, 오류 (구조설계 R-15)
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7항 / A8_FAM3_Router_구조설계_v1.0.xlsx
'============================================================================================================================

Public Class clsRouterLog

    Private Shared ReadOnly mLock As New Object

    Public Shared Sub Write(ByVal appPath As String, ByVal msg As String)
        '==========================================================================================
        'Function : 로그 한 줄 기록 (여러 스레드에서 불러도 되도록 잠근다)
        '==========================================================================================
        Try
            SyncLock mLock
                Dim folder As String = IO.Path.Combine(appPath, "System_Log", Format(DateTime.Now, "yyyy-MM"))
                If Not IO.Directory.Exists(folder) Then IO.Directory.CreateDirectory(folder)
                Dim file As String = IO.Path.Combine(folder, Format(DateTime.Now, "yyyy-MM-dd") & ".txt")
                IO.File.AppendAllText(file, Format(DateTime.Now, "yyyy-MM-dd HH:mm:ss") & vbTab & msg & vbCrLf, System.Text.Encoding.UTF8)
            End SyncLock
        Catch ex As Exception
            Debug.Print("Log write NG : " & ex.Message)
        End Try
    End Sub

End Class
