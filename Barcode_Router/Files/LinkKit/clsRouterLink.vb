Option Strict On : Option Explicit On
'================================================================================================================================
'Class Name : clsRouterLink
'Function   : 바코드 분배 프로그램(Barcode_Router) 연동 - 공통판 (어느 VB.NET 검사 프로그램에도 그대로 넣는다)
'Framework  : .NET Framework 4.x (System.Data.OleDb, System.IO.Ports 만 사용 - 프로그램 고유 클래스에 의존하지 않음)
'Memo       : ● 설정 파일 : <설정 폴더>\RouterLink.mdb > tblRouterLink (1행)
'                 RouterLink  : "Yes" / "No"  - 첫 글자가 Y 이면 사용
'               바코드 포트는 이 파일에 넣지 않는다 → 원래 설정 <설정 폴더>\EquipmentSetting.mdb > tblEquipmentSetting >
'               BarcodeReaderPort / BarcodeReaderSetting (예 COM9 / 9600-N-8-1)
'                 Memo        : 자유 입력
'               파일이 없거나 / 읽기 오류 / Y 가 아니면 미사용 → 아무것도 하지 않음 (원래 프로그램과 같은 동작)
'               처음 한 번만 읽고 결과는 메시지 없이 로그로만 남긴다
'             ● 사용법 (모든 프로그램 같은 방식, 호출 1줄)
'               원래 바코드 COM (EquipmentSetting.mdb BarcodeReaderPort) 을 SerialPort1 로 받는 프로그램에서
'               바코드 수신 → 검사 처리가 끝난 "뒤" (SerialPort1_DataReceived 끝, OK / NG / 중지 모두 지나가는 곳) 에
'                   Call clsRouterLink.SendEnd(SerialPort1, mAppPath & "Files\", Sub(m) SysLogWrite(True, m))
'               마지막 인수(로그)는 생략 가능. 설정 폴더 = RouterLink.mdb 가 있는 폴더 (EquipmentSetting.mdb 와 같은 폴더) (끝의 \ 는 있어도 없어도 됨)
'             ● Router 규약 : 바코드 + CR 수신 → 검사 → "END" + CR 회신. 창을 앞으로 띄우는 것은 Router 가 한다
'               END 를 보내기 전에 AllowSetForegroundWindow 로 Router 가 다시 맨 앞으로 올 수 있게 권한을 넘긴다
'             ● 이 클래스는 예외를 밖으로 내보내지 않는다 (연동 실패가 검사를 멈추지 않게)
'Issue      : 2026/11/13  Ver0.00    YSH       New Release (Barcode_Router 연동 가이드 배포판)
'             2026/11/16  Ver0.00    YSH       연동 방식을 1 가지로 통일 (COM 수신 + END 회신). HID 형 Start 삭제
'             근거 : Barcode_Router 연동 공통 규약.md Rev 14
'================================================================================================================================
Imports System.Collections.Generic
Imports System.Data.OleDb
Imports System.IO
Imports System.IO.Ports
Imports System.Runtime.InteropServices

Public NotInheritable Class clsRouterLink

    Public Const FILE_NAME As String = "RouterLink.mdb"
    Private Const TABLE_NAME As String = "tblRouterLink"
    Private Const ROUTER_END As String = "END"
    Private Const ASFW_ANY As Integer = -1
    Private Shared ReadOnly PROVIDERS() As String = {"Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0"}

    <DllImport("user32.dll")>
    Private Shared Function AllowSetForegroundWindow(ByVal dwProcessId As Integer) As Boolean
    End Function

    Private Shared mChecked As Boolean = False
    Private Shared mEnabled As Boolean = False

    Private Sub New()
    End Sub

    '============================================================================================================================
    ' 사용 여부
    '============================================================================================================================
    Public Shared Function IsEnabled(ByVal dataFolder As String, Optional ByVal log As Action(Of String) = Nothing) As Boolean
        '==========================================================================================
        'Function : RouterLink.mdb 를 처음 한 번 읽고 사용 여부를 돌려준다
        'Return   : RouterLink 첫 글자가 Y 이면 True. 파일 없음 / 읽기 오류 / 그 밖의 값이면 False
        '==========================================================================================
        If mChecked Then Return mEnabled

        Dim path As String = ""
        Dim msg As String
        Try
            path = System.IO.Path.Combine(If(dataFolder, ""), FILE_NAME)
            If Not File.Exists(path) Then
                msg = "RouterLink : No (" & FILE_NAME & " not found : " & path & ")"
            Else
                Dim v As Dictionary(Of String, String) = ReadRow(path, TABLE_NAME)
                Dim s As String = Value(v, "RouterLink")
                mEnabled = s.ToUpper.StartsWith("Y")
                msg = "RouterLink : " & If(mEnabled, "Yes", "No") & " (value '" & s & "', " & path & ")"
            End If
        Catch ex As Exception
            mEnabled = False
            msg = "RouterLink : No (read error " & ex.Message & ", " & path & ")"
        End Try

        mChecked = True
        Call WriteLog(log, msg)
        Return mEnabled

    End Function

    '============================================================================================================================
    ' (A) COM 수신형 : 검사 끝 회신
    '============================================================================================================================
    Public Shared Sub SendEnd(ByVal port As SerialPort, ByVal dataFolder As String, Optional ByVal log As Action(Of String) = Nothing)
        '==========================================================================================
        'Function : "END + CR" 을 Router 로 회신. 미사용이면 아무것도 하지 않는다
        'Data     : port = 바코드를 받는 SerialPort (Router 와 이어진 가상 COM 이 열려 있어야 한다)
        '==========================================================================================
        If Not IsEnabled(dataFolder, log) Then Exit Sub

        Dim msg As String
        Try
            If port Is Nothing OrElse Not port.IsOpen Then
                msg = "Router END Send skip : port closed"
            Else
                Call AllowSetForegroundWindow(ASFW_ANY)
                port.Write(ROUTER_END & vbCr)
                msg = "Router END Send OK : " & port.PortName
            End If
        Catch ex As Exception
            msg = "Router END Send NG : " & ex.Message
        End Try

        Call WriteLog(log, msg)

    End Sub

    '============================================================================================================================
    ' 내부
    '============================================================================================================================
    Private Shared Function ReadRow(ByVal path As String, ByVal table As String) As Dictionary(Of String, String)
        Dim lastErr As Exception = Nothing
        For Each p As String In PROVIDERS
            Try
                Using cn As New OleDbConnection("Provider=" & p & ";Data Source=" & path & ";Mode=Read;")
                    cn.Open()
                    Using cmd As New OleDbCommand("SELECT * FROM " & table, cn)
                        Using rd As OleDbDataReader = cmd.ExecuteReader()
                            Dim d As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
                            If rd.Read() Then
                                For i As Integer = 0 To rd.FieldCount - 1
                                    d(rd.GetName(i)) = If(rd.IsDBNull(i), "", Convert.ToString(rd.GetValue(i)).Trim)
                                Next
                            End If
                            Return d
                        End Using
                    End Using
                End Using
            Catch ex As InvalidOperationException
                lastErr = ex            'provider 없음 → 다음 provider
            End Try
        Next
        Throw New ApplicationException("OLEDB provider not found (Jet 4.0 / ACE)" & If(lastErr Is Nothing, "", " : " & lastErr.Message))
    End Function

    Private Shared Function Value(ByVal d As Dictionary(Of String, String), ByVal key As String) As String
        Dim s As String = ""
        If d IsNot Nothing AndAlso d.TryGetValue(key, s) Then Return s
        Return ""
    End Function

    Private Shared Sub WriteLog(ByVal log As Action(Of String), ByVal msg As String)
        Try
            If log IsNot Nothing Then log.Invoke(msg)
        Catch
            '로그 실패는 검사에 영향을 주지 않는다
        End Try
    End Sub

End Class
