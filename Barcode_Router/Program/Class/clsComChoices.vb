Option Strict On : Option Explicit On
'==========================================================
'Class Name : clsComChoices
'Function   : COM ComboBox 목록 (COM 설정 화면, 분기 COM 신규 / 수정 화면에서 같이 씀)
'Framwork   : VB2022 Framework4.8
'Memo       : - COM : PC 에 있는 COM + COM1 ~ COM20 + 지금 설정된 값 (번호 순)
'                     설비 PC 의 가상 COM 을 개발 PC 에서 미리 설정할 수도 있으므로 PC 에 없는 번호도 넣는다
'             - 지금 설정된 값은 목록에 없어도 항상 넣는다 (기존 설정이 화면에서 사라지지 않게)
'             - 통신설정은 ComboBox 로 하지 않고 직접 입력한다 (2026-10-28 지시)
'Issue      : 2026/10/28  Rev 0.00    YSH       New Release
'==========================================================

Public Class clsComChoices

    Private Const COM_MAX As Integer = 20       '목록에 기본으로 넣는 COM 번호 끝

    Public Shared Function PcPorts() As String()
        Dim ports() As String = IO.Ports.SerialPort.GetPortNames()
        Array.Sort(ports, AddressOf ComparePort)
        Return ports
    End Function

    Public Shared Function PortList(ByVal ParamArray currentValues() As String) As String()
        '==========================================================================================
        'Function : COM ComboBox 목록 (PC 의 COM + COM1 ~ COM20 + 지금 값, 번호 순, 중복 없음)
        '==========================================================================================
        Dim lst As New List(Of String)
        For Each p As String In PcPorts()
            Call AddUnique(lst, p)
        Next
        For i As Integer = 1 To COM_MAX
            Call AddUnique(lst, "COM" & i)
        Next
        For Each v As String In currentValues
            Call AddUnique(lst, v)
        Next
        lst.Sort(AddressOf ComparePort)
        Return lst.ToArray
    End Function

    Private Shared Sub AddUnique(ByVal lst As List(Of String), ByVal value As String)
        If value Is Nothing Then Exit Sub
        Dim v As String = value.Trim.ToUpperInvariant
        If v = "" OrElse lst.Contains(v) Then Exit Sub
        lst.Add(v)
    End Sub

    Private Shared Function ComparePort(ByVal a As String, ByVal b As String) As Integer
        'COM2 가 COM10 보다 앞에 오도록 번호로 비교 (번호가 없으면 글자 순)
        Dim na As Integer = PortNo(a), nb As Integer = PortNo(b)
        If na >= 0 AndAlso nb >= 0 Then Return na.CompareTo(nb)
        Return String.Compare(a, b, StringComparison.OrdinalIgnoreCase)
    End Function

    Private Shared Function PortNo(ByVal p As String) As Integer
        Dim n As Integer
        If p IsNot Nothing AndAlso p.ToUpperInvariant.StartsWith("COM") AndAlso Integer.TryParse(p.Substring(3), n) Then Return n
        Return -1
    End Function

End Class
