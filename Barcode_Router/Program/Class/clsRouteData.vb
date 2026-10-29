Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouteData
'Function   : 바코드 종류 판단, 받아오는 정보 항목 값 만들기 (tblBarcodeType, tblField)
'Framwork   : VB2022 Framework4.8
'Memo       : 2026-10-28 지시 - UI / 분기 규칙에서 쓰는 정보를 유저가 설정 (하드코딩 하지 않음)
'             - 바코드 종류 : 번호 순으로 보고 처음 맞는 종류를 돌려준다. 맞는 것이 없으면 Nothing (보내지 않음)
'                             길이(최소 / 최대, 0 = 제한 없음) + 글자 종류(ANY / NUM / ALPHA / ALNUM) + 패턴(VB Like, 비우면 보지 않음)
'             - 항목 값     : INPUT  = 찍은 바코드
'                             PLASMA = PLASMA 응답에서 PlasmaTag 뒤 ~ 다음 '$' 앞 (cls100_PLASMA_Order.DividePlasma 와 같은 방법)
'                                      → 새 항목은 tblField 에 PLASMA 항목명만 넣으면 코드 수정 없이 쓸 수 있다
'                             DERIVE = BaseField 의 값을 구분자로 나눈 몇 번째, 또는 시작 위치 + 길이로 자른 값
'                                      (예 : MODEL = MSCODE 를 '-' 로 나눈 1 번째)
'             - 설정에 없는 항목의 값은 빈 문자열
'Issue      : 2026/10/28  Rev 0.00    YSH       New Release
'============================================================================================================================

Public Class clsRouteData

    Public Shared Function ClassifyBarcode(ByVal types As List(Of clsBarcodeType), ByVal code As String) As clsBarcodeType
        For Each t As clsBarcodeType In types.OrderBy(Function(x) x.TypeNo)
            If Not t.Enabled Then Continue For
            If IsTypeMatch(t, code) Then Return t
        Next
        Return Nothing
    End Function

    Public Shared Function IsTypeMatch(ByVal t As clsBarcodeType, ByVal code As String) As Boolean
        If t.MinLen > 0 AndAlso code.Length < t.MinLen Then Return False
        If t.MaxLen > 0 AndAlso code.Length > t.MaxLen Then Return False

        Select Case t.CharKind
            Case clsBarcodeType.KIND_NUM
                If Not code.All(Function(c) Char.IsDigit(c)) Then Return False
            Case clsBarcodeType.KIND_ALPHA
                If Not code.All(Function(c) IsAsciiLetter(c)) Then Return False
            Case clsBarcodeType.KIND_ALNUM
                If Not code.All(Function(c) Char.IsDigit(c) OrElse IsAsciiLetter(c)) Then Return False
        End Select

        If t.Pattern.Trim <> "" Then
            Try
                If Not (code.ToUpperInvariant Like t.Pattern.Trim.ToUpperInvariant) Then Return False
            Catch
                Return False        '패턴이 잘못된 경우
            End Try
        End If
        Return True
    End Function

    Private Shared Function IsAsciiLetter(ByVal c As Char) As Boolean
        Return (c >= "A"c AndAlso c <= "Z"c) OrElse (c >= "a"c AndAlso c <= "z"c)
    End Function

    Public Shared Function BuildValues(ByVal fields As List(Of clsField), ByVal barcode As String, ByVal plasmaRaw As String) As Dictionary(Of String, String)
        '==========================================================================================
        'Function : 항목ID → 값. INPUT / PLASMA 를 먼저 채우고 DERIVE 를 만든다 (DERIVE 가 DERIVE 를 쓰면 순서대로 여러 번 계산)
        '==========================================================================================
        Dim v As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each f As clsField In fields
            Select Case f.SourceType
                Case clsField.SRC_INPUT : v(f.FieldID) = barcode
                Case clsField.SRC_PLASMA : v(f.FieldID) = DividePlasma(f.PlasmaTag, plasmaRaw)
            End Select
        Next

        Dim derives As List(Of clsField) = fields.Where(Function(x) x.SourceType = clsField.SRC_DERIVE).ToList
        For pass As Integer = 1 To Math.Max(1, derives.Count)
            For Each f As clsField In derives
                Dim baseValue As String = ""
                v.TryGetValue(f.BaseField, baseValue)
                v(f.FieldID) = Derive(f, If(baseValue, ""))
            Next
        Next
        Return v
    End Function

    Public Shared Function Derive(ByVal f As clsField, ByVal baseValue As String) As String
        If f.Delimiter <> "" Then
            Dim parts() As String = baseValue.Split(New String() {f.Delimiter}, StringSplitOptions.None)
            Dim i As Integer = Math.Max(1, f.PartNo) - 1
            If i >= parts.Length Then Return ""
            Return parts(i).Trim
        End If
        If f.StartPos <= 0 Then Return baseValue.Trim
        If f.StartPos > baseValue.Length Then Return ""
        If f.PartLen <= 0 Then Return baseValue.Substring(f.StartPos - 1).Trim
        Return baseValue.Substring(f.StartPos - 1, Math.Min(f.PartLen, baseValue.Length - f.StartPos + 1)).Trim
    End Function

    Public Shared Function DividePlasma(ByVal tag As String, ByVal source As String) As String
        '==========================================================================================
        'Function : PLASMA 응답에서 tag 뒤 ~ 다음 '$' 앞 (cls100_PLASMA_Order.DividePlasma 와 같은 방법)
        '==========================================================================================
        If tag Is Nothing OrElse tag.Trim = "" OrElse source Is Nothing OrElse source = "" Then Return ""
        Dim p As Integer = source.IndexOf(tag.Trim, StringComparison.Ordinal)
        If p < 0 Then Return ""
        Dim s As Integer = p + tag.Trim.Length
        Dim e As Integer = source.IndexOf("$"c, s)
        If e < 0 Then Return source.Substring(s).Trim
        Return source.Substring(s, e - s).Trim
    End Function

End Class
