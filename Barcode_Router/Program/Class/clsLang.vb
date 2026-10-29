Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsLang
'Function   : 화면 / 알림 문구 언어 (한국어 PC / 일본어 PC)
'Framwork   : VB2022 Framework4.8
'Memo       : 2026-10-28 지시 - 한국어 / 일본어 PC 를 유저가 고르고 (RouterSetting.mdb > tblConfig > UiLanguage = KO / JA),
'             그 언어로 깨지지 않게 표시한다. 문구는 코드에 한국어 / 일본어 두 가지를 같이 적고 설정에 따라 고른다
'               예 : clsLang.L("저장", "保存")
'             - UiLanguage 가 JA 이면 일본어, 그 밖(KO / 비어 있음)은 한국어 (코드 원문이 한국어이므로)
'             - 글꼴 : tblConfig UiFont 가 비어 있으면 언어에 맞는 Windows 글꼴 (KO = Malgun Gothic, JA = Meiryo UI)
'               → 한글 / 일본어 글자가 없는 글꼴로 그려져 □ 로 보이는 것을 막는다 (clsUiTheme.SetFontName)
'             - Designer 에 적힌 한국어 글자는 각 화면의 ApplyLanguageTexts 에서 바꾼다 (Designer 파일은 그대로)
'             - 소스 파일은 모두 UTF-8(BOM) 이라 빌드하는 PC 의 Windows 언어와 관계없이 같은 글자가 들어간다 (공통 규약 10-6)
'Issue      : 2026/10/28  Rev 0.00    YSH       New Release
'============================================================================================================================

Public Class clsLang

    Public Const LANG_KO As String = "KO"
    Public Const LANG_JA As String = "JA"

    Private Shared mLang As String = LANG_KO

    Public Shared Sub SetLanguage(ByVal value As String)
        mLang = If(value IsNot Nothing AndAlso value.Trim.ToUpperInvariant = LANG_JA, LANG_JA, LANG_KO)
    End Sub

    Public Shared ReadOnly Property Language() As String
        Get
            Return mLang
        End Get
    End Property

    Public Shared ReadOnly Property IsJA() As Boolean
        Get
            Return mLang = LANG_JA
        End Get
    End Property

    Public Shared Function L(ByVal ko As String, ByVal ja As String) As String
        Return If(mLang = LANG_JA, ja, ko)
    End Function

    Public Shared Function LanguageFont() As String
        '언어에 맞는 Windows 기본 글꼴 (영문 이름 - 어느 언어 Windows 에서도 같은 이름으로 찾음)
        Return If(mLang = LANG_JA, "Meiryo UI", "Malgun Gothic")
    End Function

End Class
