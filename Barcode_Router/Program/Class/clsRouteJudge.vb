Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouteJudge
'Function   : 분기 규칙 판단
'Framwork   : VB2022 Framework4.8
'Memo       : 판단 방법 (구조설계 R-04 ~ R-07, 10_확인항목 Q2 / Q3 제안대로 승인 2026-10-26)
'             - MODEL = MSCODE 의 첫 '-' 앞부분 (예 : F3SP71-4S → F3SP71)
'             - 규칙은 우선순위가 작은 것부터 보고, 처음 맞는 규칙을 돌려준다 (사용 안 함 / 분기 대상 없는 규칙은 건너뜀)
'             - 한 규칙 안의 조건 줄끼리는 줄마다 고른 연결(AND / OR)로 계산. AND 를 먼저 묶는다
'               (예 : A AND B OR C = (A 그리고 B) 또는 C)  - 2026-10-26 지시로 AND 고정에서 변경
'             - 한 조건의 입력값 여러 줄끼리는 OR
'             - EQUAL : 값이 완전히 같음
'               LIKE  : 값을 포함. 입력값에 * ? # [ 가 있으면 VB Like 패턴으로 판단 (예 : F3* = F3 로 시작)
'             - 대소문자는 구분하지 않는다
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-3항, 7-4항
'============================================================================================================================

Public Class clsRouteJudge

    Public Shared Function GetModel(ByVal mscode As String) As String
        Dim p As Integer = mscode.IndexOf("-"c)
        If p < 0 Then Return mscode.Trim
        Return mscode.Substring(0, p).Trim
    End Function

    Public Shared Function Judge(ByVal rules As List(Of clsRouteRule), ByVal mscode As String) As clsRouteRule
        Dim model As String = GetModel(mscode)

        For Each ru As clsRouteRule In rules.OrderBy(Function(x) x.Priority).ThenBy(Function(x) x.RuleID)
            If Not ru.Enabled Then Continue For
            If ru.Conds.Count = 0 Then Continue For

            If IsRuleMatch(ru, mscode, model) Then Return ru
        Next

        Return Nothing
    End Function

    Public Shared Function IsRuleMatch(ByVal ru As clsRouteRule, ByVal mscode As String, ByVal model As String) As Boolean
        '==========================================================================================
        'Function : 조건 줄을 연결(AND / OR)대로 계산 - AND 를 먼저 묶고, OR 로 나뉜 묶음 중 하나라도 맞으면 True
        '           예 : A AND B OR C = (A 그리고 B) 또는 C
        '==========================================================================================
        Dim groupOk As Boolean = True
        For i As Integer = 0 To ru.Conds.Count - 1
            Dim c As clsRuleCond = ru.Conds(i)
            If i > 0 AndAlso c.JoinOp = clsRuleCond.JOIN_OR Then
                If groupOk Then Return True         '앞 묶음이 모두 맞음
                groupOk = True                      '새 묶음 시작
            End If
            If Not groupOk Then Continue For        '이 묶음은 이미 틀림 - 다음 OR 까지 건너뜀

            Dim target As String = If(c.Target = clsRuleCond.TARGET_MODEL, model, mscode)
            Dim anyOk As Boolean = False
            For Each v As String In c.ValueList()
                If IsMatch(target, c.Operator, v) Then anyOk = True : Exit For
            Next
            If Not anyOk Then groupOk = False
        Next
        Return groupOk
    End Function

    Public Shared Function IsMatch(ByVal value As String, ByVal op As String, ByVal pattern As String) As Boolean
        Dim v As String = value.Trim.ToUpperInvariant
        Dim p As String = pattern.Trim.ToUpperInvariant
        If p = "" Then Return False

        If op = clsRuleCond.OP_EQUAL Then
            Return v = p
        End If

        'LIKE
        If p.IndexOfAny(New Char() {"*"c, "?"c, "#"c, "["c}) >= 0 Then
            Try
                Return v Like p
            Catch
                Return False        '패턴이 잘못된 경우
            End Try
        End If
        Return v.Contains(p)
    End Function

End Class
