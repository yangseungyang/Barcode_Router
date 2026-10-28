Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterPath
'Function   : Router 의 데이터 폴더 (설정 MDB, System_Log) 를 exe 위치와 분리해서 정한다
'Framwork   : VB2022 Framework4.8
'Memo       : exe 는 어디에 있어도 된다 (바탕화면, 복사본, 바로가기 모두 같은 설정 / 로그를 쓴다)
'             데이터 폴더 : 기본 C:\Barcode_Router\
'                           실행 인수 /data:폴더 로 바꿀 수 있다 (예 : 바로가기 대상 끝에  /data:D:\Barcode_Router)
'                 Files\RouterSetting.mdb   : 설정
'                 System_Log\yyyy-MM\       : 로그
'             설정 파일이 없을 때 (처음 실행)
'                 1) 예전 위치(Ver0.00 규칙 : exe 경로의 \bin 한 단계 위, 없으면 exe 폴더)의 Files\RouterSetting.mdb 가 있으면 복사
'                 2) 없으면 exe 안에 넣어 둔 빈 설정(RouterSetting_Template.mdb)으로 새로 만든다
'Issue      : 2026/10/28  Ver0.01    YSH       New Release
'             근거 : Barcode_Router 연동 공통 규약.md Rev 1 7항 (exe 와 설정 / 로그 분리)
'============================================================================================================================

Public Class clsRouterPath

    Public Const DEFAULT_DATA_FOLDER As String = "C:\Barcode_Router\"
    Private Const ARG_DATA As String = "/data:"
    Private Const SETTING_RELATIVE As String = "Files\RouterSetting.mdb"
    Private Const TEMPLATE_RESOURCE As String = "RouterSetting_Template.mdb"

    Private Shared mDataFolder As String = Nothing

    Public Shared ReadOnly Property DataFolder() As String
        '==========================================================================================
        'Function : 데이터 폴더 (끝에 \ 포함). 실행 인수 /data: 가 있으면 그 폴더, 없으면 C:\Barcode_Router\
        '==========================================================================================
        Get
            If mDataFolder Is Nothing Then
                Dim f As String = DEFAULT_DATA_FOLDER
                For Each a As String In Environment.GetCommandLineArgs()
                    If a.StartsWith(ARG_DATA, StringComparison.OrdinalIgnoreCase) AndAlso a.Length > ARG_DATA.Length Then
                        f = a.Substring(ARG_DATA.Length).Trim().Trim(""""c)
                    End If
                Next
                If Not f.EndsWith("\") Then f &= "\"
                mDataFolder = f
            End If
            Return mDataFolder
        End Get
    End Property

    Public Shared ReadOnly Property SettingFile() As String
        Get
            Return DataFolder & SETTING_RELATIVE
        End Get
    End Property

    Public Shared Function LegacyAppPath() As String
        '==========================================================================================
        'Function : Ver0.00 의 데이터 위치 규칙 (exe 경로의 \bin 한 단계 위, \bin 이 없으면 exe 폴더)
        '==========================================================================================
        Dim p As String = Application.StartupPath
        Dim i As Integer = InStr(p, "\bin", CompareMethod.Text)
        If i > 0 Then Return Mid(p, 1, i)
        Return p & "\"
    End Function

    Public Shared Function PrepareSettingFile() As String
        '==========================================================================================
        'Function : 설정 파일이 없으면 준비한다 (예전 위치에서 복사 → 없으면 빈 설정으로 새로 만듦)
        'Return   : 한 일 (로그 / 안내용). 이미 있으면 ""
        '==========================================================================================
        Dim target As String = SettingFile
        If IO.File.Exists(target) Then Return ""

        IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(target))

        Dim legacy As String = LegacyAppPath() & SETTING_RELATIVE
        If IO.File.Exists(legacy) AndAlso Not String.Equals(IO.Path.GetFullPath(legacy), IO.Path.GetFullPath(target), StringComparison.OrdinalIgnoreCase) Then
            IO.File.Copy(legacy, target)
            Return "예전 위치의 설정을 복사했습니다 : " & legacy & " → " & target
        End If

        Using s As IO.Stream = GetType(clsRouterPath).Assembly.GetManifestResourceStream(TEMPLATE_RESOURCE)
            If s Is Nothing Then Throw New ApplicationException("빈 설정 파일(내장)을 찾지 못했습니다 : " & TEMPLATE_RESOURCE)
            Using fs As New IO.FileStream(target, IO.FileMode.CreateNew)
                s.CopyTo(fs)
            End Using
        End Using
        Return "설정 파일을 새로 만들었습니다 (빈 설정) : " & target
    End Function

End Class
