Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmDestEdit
'Function   : 분기 COM (분기 대상) 1개 신규 / 수정 화면
'Framwork   : VB2022 Framework4.8
'Memo       : 분기 조건 설정 화면에서 COM 설정 화면으로 가지 않고 바로 분기 COM 을 새로 등록하거나 이름 / COM 을 고친다
'             (2026-10-26 지시 : 분기 COM 이 COM8 / COM11 로 고정되어 보임 - 새 등록, 이름 수정, 신규 추가)
'             - COM 은 PC 의 COM 목록에서 고르거나 직접 입력 (가상 COM 을 나중에 만들 수도 있으므로)
'             - 다른 분기 대상 / 바코드리더와 같은 COM 은 쓸 수 없다
'             - 저장은 호출한 화면이 RouterSetting.mdb 에 한다
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'==========================================================

Public Class frmDestEdit

    Private mDest As clsDestination
    Private mUsedPorts As List(Of String)       '다른 곳에서 쓰는 COM (대문자)

    Public ReadOnly Property Result() As clsDestination
        Get
            Return mDest
        End Get
    End Property

    Public Sub New(ByVal dest As clsDestination, ByVal isNew As Boolean, ByVal usedPorts As List(Of String))
        InitializeComponent()
        mDest = dest
        mUsedPorts = usedPorts
        Label_Header.Text = If(isNew, "＋  분기 COM 신규", "✎  분기 COM 수정") & "   (No " & dest.DestID & ")"
    End Sub

    Private Sub frmDestEdit_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        clsUiTheme.StyleButton(Button_PortTest, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Save, clsUiTheme.ButtonKind.Primary)
        clsUiTheme.StyleButton(Button_Close, clsUiTheme.ButtonKind.Normal)

        Dim pcPorts() As String = IO.Ports.SerialPort.GetPortNames()
        Array.Sort(pcPorts)
        ComboBox_Port.Items.AddRange(pcPorts)
        Label_PcPorts.Text = "직접 입력 가능 · PC : " & If(pcPorts.Length = 0, "(없음)", String.Join(", ", pcPorts))

        TextBox_Name.Text = mDest.DestName
        ComboBox_Port.Text = mDest.Port
        TextBox_Setting.Text = If(mDest.Setting = "", "9600-N-8-1", mDest.Setting)
        TextBox_Process.Text = mDest.ProcessName
        NumericUpDown_Timeout.Value = Math.Max(0, Math.Min(3600, mDest.EndTimeoutSec))
        CheckBox_Enabled.Checked = mDest.Enabled
    End Sub

    Private Sub Button_PortTest_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_PortTest.Click
        Dim port As String = ComboBox_Port.Text.Trim.ToUpperInvariant
        Dim p As New clsRouterPort
        Try
            p.Open(port, TextBox_Setting.Text.Trim)
            p.Close()
            MessageBox.Show(port & " 열기 / 닫기 OK", "포트 시험", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(port & " 열기 NG" & vbCrLf & ex.Message, "포트 시험", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        Dim name As String = TextBox_Name.Text.Trim
        Dim port As String = ComboBox_Port.Text.Trim.ToUpperInvariant
        Dim setting As String = TextBox_Setting.Text.Trim.ToUpperInvariant

        If name = "" Then
            MessageBox.Show("이름을 넣어 주세요.", "분기 COM", MessageBoxButtons.OK, MessageBoxIcon.Warning) : TextBox_Name.Focus() : Exit Sub
        End If
        If Not System.Text.RegularExpressions.Regex.IsMatch(port, "^COM\d+$") Then
            MessageBox.Show("COM 을 넣어 주세요 (예 : COM15).", "분기 COM", MessageBoxButtons.OK, MessageBoxIcon.Warning) : ComboBox_Port.Focus() : Exit Sub
        End If
        If setting.Split("-"c).Length < 4 Then
            MessageBox.Show("통신설정을 넣어 주세요 (예 : 9600-N-8-1).", "분기 COM", MessageBoxButtons.OK, MessageBoxIcon.Warning) : TextBox_Setting.Focus() : Exit Sub
        End If
        If CheckBox_Enabled.Checked AndAlso mUsedPorts.Contains(port) Then
            MessageBox.Show(port & " 는 다른 분기 대상이나 바코드리더에서 쓰고 있습니다.", "분기 COM", MessageBoxButtons.OK, MessageBoxIcon.Warning) : ComboBox_Port.Focus() : Exit Sub
        End If

        mDest.DestName = name
        mDest.Port = port
        mDest.Setting = setting
        mDest.ProcessName = TextBox_Process.Text.Trim.Replace(".exe", "").Replace(".EXE", "")
        mDest.EndTimeoutSec = CInt(NumericUpDown_Timeout.Value)
        mDest.Enabled = CheckBox_Enabled.Checked
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
