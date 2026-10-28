Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterPath
'Function   : Router 의 데이터 폴더 (Files\RouterSetting.mdb, System_Log) 를 정한다
'Framwork   : VB2022 Framework4.8
'Memo       : 저장소 프로젝트 폴더의 Files\RouterSetting.mdb 를 그대로 읽는다 (exe 를 바탕화면 등으로 복사해도 같은 파일)
'             찾는 순서 1) 실행 인수 /data:폴더
'                       2) exe 경로의 \bin 한 단계 위에 Files\RouterSetting.mdb 가 있으면 그 폴더 (저장소 bin\Release 에서 실행)
'                       3) 그 외 C:\Barcode_Router\Barcode_Router\ (설비 PC 저장소 위치)
'Issue      : 2026/10/28  Rev 0.00    YSH       New Release
'             근거 : Barcode_Router 연동 공통 규약.md 7항
'============================================================================================================================

Public Class clsRouterPath

    Public Const REPO_DATA_FOLDER As String = "C:\Barcode_Router\Barcode_Router\"     '설비 PC 저장소(C:\Barcode_Router) 의 프로젝트 폴더
    Private Const ARG_DATA As String = "/data:"
    Private Const SETTING_RELATIVE As String = "Files\RouterSetting.mdb"

    Private Shared mDataFolder As String = Nothing

    Public Shared ReadOnly Property DataFolder() As String
        '==========================================================================================
        'Function : 데이터 폴더 (끝에 \ 포함)
        '==========================================================================================
        Get
            If mDataFolder Is Nothing Then
                Dim f As String = ""
                For Each a As String In Environment.GetCommandLineArgs()
                    If a.StartsWith(ARG_DATA, StringComparison.OrdinalIgnoreCase) AndAlso a.Length > ARG_DATA.Length Then
                        f = a.Substring(ARG_DATA.Length).Trim().Trim(""""c)
                    End If
                Next
                If f = "" Then
                    If IO.File.Exists(ExeAppPath() & SETTING_RELATIVE) Then
                        f = ExeAppPath()
                    Else
                        f = REPO_DATA_FOLDER
                    End If
                End If
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

    Private Shared Function ExeAppPath() As String
        '==========================================================================================
        'Function : exe 경로의 \bin 한 단계 위 (\bin 이 없으면 exe 폴더)
        '==========================================================================================
        Dim p As String = Application.StartupPath
        Dim i As Integer = InStr(p, "\bin", CompareMethod.Text)
        If i > 0 Then Return Mid(p, 1, i)
        Return p & "\"
    End Function

End Class
