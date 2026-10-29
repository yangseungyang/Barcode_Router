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

    '종료 신호 / 끝 문자 (2026-10-28 추가 - 분기 대상마다 다를 수 있음. Designer 를 건드리지 않고 코드에서 만든다)
    Private WithEvents TextBox_EndMessage As New TextBox With {.BorderStyle = BorderStyle.FixedSingle, .Font = clsUiTheme.UiFont(10.0!), .Size = New Size(80, 25)}
    Private ComboBox_LineEnd As New ComboBox With {.DropDownStyle = ComboBoxStyle.DropDownList, .FlatStyle = FlatStyle.Flat, .Font = clsUiTheme.UiFont(10.0!), .Size = New Size(80, 25)}

    Private Sub AddEndControls()
        '"사용" 줄 앞에 한 줄 넣고 아래를 내린다
        Const ROW_H As Integer = 36
        Dim y As Integer = CheckBox_Enabled.Top
        For Each c As Control In New Control() {CheckBox_Enabled, Button_PortTest, Button_Save, Button_Close}
            c.Top += ROW_H
        Next
        Me.ClientSize = New Size(Me.ClientSize.Width, Me.ClientSize.Height + ROW_H)

        Dim capEnd As New Label With {.Text = clsLang.L("종료 신호", "終了信号"), .ForeColor = Label_TimeoutCap.ForeColor, .Font = Label_TimeoutCap.Font, .Location = New Point(Label_TimeoutCap.Left, y + 3), .Size = New Size(90, 22)}
        TextBox_EndMessage.Location = New Point(NumericUpDown_Timeout.Left, y)
        Dim capLine As New Label With {.Text = clsLang.L("끝 문자", "終端文字"), .ForeColor = Label_TimeoutCap.ForeColor, .Font = Label_TimeoutCap.Font, .Location = New Point(NumericUpDown_Timeout.Left + 96, y + 3), .Size = New Size(60, 22)}
        ComboBox_LineEnd.Location = New Point(NumericUpDown_Timeout.Left + 160, y)
        Dim hint As New Label With {.Text = clsLang.L("검사 프로그램이 끝나면 보내는 글자 / 보낼 때 붙이는 끝 문자", "検査プログラムが終了時に送る文字 / 送信時に付ける終端文字"), .ForeColor = Label_TimeoutHint.ForeColor, .Font = Label_TimeoutHint.Font,
                                    .Location = New Point(NumericUpDown_Timeout.Left + 246, y + 3), .Size = New Size(Me.ClientSize.Width - NumericUpDown_Timeout.Left - 256, 22), .AutoEllipsis = True}
        ComboBox_LineEnd.Items.AddRange(New Object() {clsDestination.LINE_CR, clsDestination.LINE_LF, clsDestination.LINE_CRLF})
        Me.Controls.AddRange(New Control() {capEnd, TextBox_EndMessage, capLine, ComboBox_LineEnd, hint})
    End Sub

    Public ReadOnly Property Result() As clsDestination
        Get
            Return mDest
        End Get
    End Property

    Public Sub New(ByVal dest As clsDestination, ByVal isNew As Boolean, ByVal usedPorts As List(Of String))
        InitializeComponent()
        Call ApplyLanguageTexts()       '화면 글자를 언어 설정(clsLang)에 맞게 (2026-10-28). 아래 제목보다 먼저
        mDest = dest
        mUsedPorts = usedPorts
        Label_Header.Text = If(isNew, clsLang.L("＋  분기 COM 신규", "＋  分岐COM 新規"), clsLang.L("✎  분기 COM 수정", "✎  分岐COM 修正")) & "   (No " & dest.DestID & ")"
    End Sub

    Private Sub frmDestEdit_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        clsUiTheme.StyleButton(Button_PortTest, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Save, clsUiTheme.ButtonKind.Primary)
        clsUiTheme.StyleButton(Button_Close, clsUiTheme.ButtonKind.Normal)

        Dim pcPorts() As String = clsComChoices.PcPorts()
        ComboBox_Port.Items.AddRange(clsComChoices.PortList(mDest.Port))              'COM 목록 (clsComChoices, 2026-10-28 지시)
        Label_PcPorts.Text = clsLang.L("직접 입력 가능 · PC : ", "直接入力可 · PC : ") & If(pcPorts.Length = 0, clsLang.L("(없음)", "(なし)"), String.Join(", ", pcPorts))

        TextBox_Name.Text = mDest.DestName
        ComboBox_Port.Text = mDest.Port.ToUpperInvariant
        TextBox_Setting.Text = mDest.Setting
        TextBox_Process.Text = mDest.ProcessName
        NumericUpDown_Timeout.Value = Math.Max(0, Math.Min(3600, mDest.EndTimeoutSec))
        CheckBox_Enabled.Checked = mDest.Enabled

        Call AddEndControls()
        TextBox_EndMessage.Text = mDest.EndMessage
        ComboBox_LineEnd.SelectedItem = If(ComboBox_LineEnd.Items.Contains(mDest.LineEnd.ToUpperInvariant), mDest.LineEnd.ToUpperInvariant, Nothing)
        Label_ProcessHint.Text = clsLang.L("바코드를 보낼 때 앞으로 띄울 프로그램 (.exe 없이 이름, 또는 exe 전체 경로)", "バーコード送信時に前面に出すプログラム (.exeなしの名前、またはexeのフルパス)")
        clsUiTheme.ApplyFont(Me)
    End Sub

    Private Sub Button_PortTest_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_PortTest.Click
        Dim port As String = ComboBox_Port.Text.Trim.ToUpperInvariant
        Dim p As New clsRouterPort
        Try
            p.Open(port, TextBox_Setting.Text.Trim)
            p.Close()
            MessageBox.Show(port & clsLang.L(" 열기 / 닫기 OK", " オープン / クローズ OK"), clsLang.L("포트 시험", "ポート試験"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(port & clsLang.L(" 열기 NG", " オープン NG") & vbCrLf & ex.Message, clsLang.L("포트 시험", "ポート試験"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        Dim name As String = TextBox_Name.Text.Trim
        Dim port As String = ComboBox_Port.Text.Trim.ToUpperInvariant
        Dim setting As String = TextBox_Setting.Text.Trim.ToUpperInvariant

        If name = "" Then
            MessageBox.Show(clsLang.L("이름을 넣어 주세요.", "名前を入力してください。"), clsLang.L("분기 COM", "分岐COM"), MessageBoxButtons.OK, MessageBoxIcon.Warning) : TextBox_Name.Focus() : Exit Sub
        End If
        If Not System.Text.RegularExpressions.Regex.IsMatch(port, "^COM\d+$") Then
            MessageBox.Show(clsLang.L("COM 을 넣어 주세요 (예 : COM15).", "COMを入力してください (例 : COM15)。"), clsLang.L("분기 COM", "分岐COM"), MessageBoxButtons.OK, MessageBoxIcon.Warning) : ComboBox_Port.Focus() : Exit Sub
        End If
        If setting.Split("-"c).Length < 4 Then
            MessageBox.Show(clsLang.L("통신설정을 넣어 주세요 (예 : 9600-N-8-1).", "通信設定を入力してください (例 : 9600-N-8-1)。"), clsLang.L("분기 COM", "分岐COM"), MessageBoxButtons.OK, MessageBoxIcon.Warning) : TextBox_Setting.Focus() : Exit Sub
        End If
        If CheckBox_Enabled.Checked AndAlso mUsedPorts.Contains(port) Then
            MessageBox.Show(port & clsLang.L(" 는 다른 분기 대상이나 바코드리더에서 쓰고 있습니다.", " は他の分岐先かバーコードリーダで使われています。"), clsLang.L("분기 COM", "分岐COM"), MessageBoxButtons.OK, MessageBoxIcon.Warning) : ComboBox_Port.Focus() : Exit Sub
        End If

        If TextBox_EndMessage.Text.Trim = "" OrElse ComboBox_LineEnd.SelectedItem Is Nothing Then
            MessageBox.Show(clsLang.L("종료 신호(예 : END)와 끝 문자를 넣어 주세요.", "終了信号(例 : END)と終端文字を入力してください。") & vbCrLf & clsLang.L("종료 신호가 없으면 검사가 끝난 것을 알 수 없습니다.", "終了信号がないと検査の終了が分かりません。"), clsLang.L("분기 COM", "分岐COM"), MessageBoxButtons.OK, MessageBoxIcon.Warning) : TextBox_EndMessage.Focus() : Exit Sub
        End If

        mDest.DestName = name
        mDest.Port = port
        mDest.Setting = setting
        Dim proc As String = TextBox_Process.Text.Trim
        mDest.ProcessName = If(proc.Contains("\"), proc, proc.Replace(".exe", "").Replace(".EXE", ""))      'exe 전체 경로는 그대로 (같은 이름 exe 구분)
        mDest.EndMessage = TextBox_EndMessage.Text.Trim
        mDest.LineEnd = ComboBox_LineEnd.SelectedItem.ToString
        mDest.EndTimeoutSec = CInt(NumericUpDown_Timeout.Value)
        mDest.Enabled = CheckBox_Enabled.Checked
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub


    Private Sub ApplyLanguageTexts()
        '==========================================================================================
        'Function : Designer 에 적힌 한국어 글자를 언어 설정(clsLang)에 맞게 바꾼다 (Designer 파일은 그대로, 2026-10-28)
        '==========================================================================================
        Me.Label_Header.Text = clsLang.L("⇄  분기 COM", "⇄  分岐COM")
        Me.Label_NameCap.Text = clsLang.L("이름", "名前")
        Me.Label_PcPorts.Text = clsLang.L("직접 입력 가능 (예 : COM15)", "直接入力可 (例 : COM15)")
        Me.Label_SettingCap.Text = clsLang.L("통신설정", "通信設定")
        Me.Label_ProcessCap.Text = clsLang.L("프로세스명", "プロセス名")
        Me.Label_ProcessHint.Text = clsLang.L("바코드를 보낼 때 앞으로 띄울 프로그램 (.exe 없이, 예 : AIEC_HIPOT)", "バーコード送信時に前面に出すプログラム (.exeなし、例 : AIEC_HIPOT)")
        Me.Label_TimeoutCap.Text = clsLang.L("END 대기", "END待ち")
        Me.Label_TimeoutHint.Text = clsLang.L("초  (0 = 무제한)", "秒  (0 = 無制限)")
        Me.CheckBox_Enabled.Text = clsLang.L("사용", "使用")
        Me.Button_PortTest.Text = clsLang.L("⇄  포트 시험", "⇄  ポート試験")
        Me.Button_Save.Text = clsLang.L("저장", "保存")
        Me.Button_Close.Text = clsLang.L("닫기", "閉じる")
        Me.Text = clsLang.L("분기 COM", "分岐COM")
    End Sub

End Class
