Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterLog
'Function   : 바코드 분배 프로그램 System Log 기록
'Framwork   : VB2022 Framework4.8
'Memo       : 기존 검사 프로그램(A8 / FAM3) 의 System_Log 와 같은 위치 규칙을 쓴다
'             {프로그램 폴더}\System_Log\yyyy-MM\yyyy-MM-dd.txt   (월 폴더가 없으면 만든다)
'             시작할 때 tblConfig LogKeepMonths 보다 오래된 월 폴더를 지운다 (DeleteOld, 2026-10-28)
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

    Public Shared Function DeleteOld(ByVal appPath As String, ByVal keepMonths As Integer) As List(Of String)
        '==========================================================================================
        'Function : System_Log 의 월 폴더(yyyy-MM) 중 keepMonths 개월보다 오래된 것을 지운다 (2026-10-28 지시)
        'Data     : keepMonths = tblConfig LogKeepMonths. 0 이하(설정 없음)면 지우지 않는다
        'Return   : 지운 폴더 이름
        'Remarks  : 예 : 6 이고 지금이 2026-10 이면 2026-04 이전 폴더를 지운다 (2026-05 ~ 2026-10 남김)
        '==========================================================================================
        Dim deleted As New List(Of String)
        If keepMonths <= 0 Then Return deleted
        Dim root As String = IO.Path.Combine(appPath, "System_Log")
        If Not IO.Directory.Exists(root) Then Return deleted

        Dim limit As New DateTime(DateTime.Now.Year, DateTime.Now.Month, 1)
        limit = limit.AddMonths(-(keepMonths - 1))
        For Each dir As String In IO.Directory.GetDirectories(root)
            Dim name As String = IO.Path.GetFileName(dir)
            Dim d As DateTime
            If Not DateTime.TryParseExact(name, "yyyy-MM", Globalization.CultureInfo.InvariantCulture, Globalization.DateTimeStyles.None, d) Then Continue For
            If d >= limit Then Continue For
            Try
                IO.Directory.Delete(dir, True)
                deleted.Add(name)
            Catch ex As Exception
                Debug.Print("Log delete NG : " & ex.Message)
            End Try
        Next
        Return deleted
    End Function

End Class
