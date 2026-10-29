Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmComSetting
'Function   : COM 설정 화면 (바코드리더 입력 방식 / 포트, 분기 대상별 COM)
'Framwork   : VB2022 Framework4.8
'Memo       : - 설비명     : 이 Router 가 있는 설비 이름 (메인 화면 제목줄 가운데 표시, 2026-10-26 추가)
'             - 바코드리더 : 입력 방식 HID(바코드 칸 키보드 입력) / COM, 리더 포트, 통신설정 (구조설계 C-01, C-02)
'             - 분기 대상 : 이름, Router 쪽 COM, 통신설정, 프로세스명(창 앞으로), END 대기 초, 사용 (구조설계 C-03)
'             - 포트 시험 : 선택한 행의 COM 을 열었다 닫아 본다 (구조설계 C-04)
'             - COM 은 ComboBox 로 고른다 (목록 : clsComChoices, 2026-10-28 지시). 바코드리더 칸은 직접 입력도 된다
'               통신설정은 직접 입력 (예 : 9600-N-8-1)
'             - 검사 프로그램 쪽 COM 은 각 프로그램의 EquipmentSetting.mdb > tblEquipmentSetting > BarcodeReaderPort 로 설정
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-10항 (COM 은 따로 설정할 수 있는 설정 화면)
'==========================================================

Public Class frmComSetting

    Private mSetting As clsRouterSetting

    '종료 신호 / 끝 문자 열 (2026-10-28 추가, Designer 를 건드리지 않고 코드에서 '사용' 열 앞에 넣는다)
    Private colEndMessage As New DataGridViewTextBoxColumn With {.HeaderText = clsLang.L("종료 신호", "終了信号"), .Name = "colEndMessage", .Width = 80}
    Private colLineEnd As New DataGridViewComboBoxColumn With {.HeaderText = clsLang.L("끝 문자", "終端文字"), .Name = "colLineEnd", .Width = 70, .FlatStyle = FlatStyle.Flat}

    Public Sub New(ByVal setting As clsRouterSetting)
        InitializeComponent()
        mSetting = setting
    End Sub

    Private Sub frmComSetting_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Call ApplyLanguageTexts()       '화면 글자를 언어 설정(clsLang)에 맞게 (2026-10-28)
        Dim pcPorts() As String = clsComChoices.PcPorts()
        Label_PcPorts.Text = clsLang.L("PC 의 COM : ", "PCのCOM : ") & If(pcPorts.Length = 0, clsLang.L("(없음)", "(なし)"), String.Join(", ", pcPorts))

        '화면 모양 (Router_UI_시안_v1.html)
        For Each rb As RadioButton In {RadioButton_HID, RadioButton_COM}
            rb.FlatStyle = FlatStyle.Flat
            rb.FlatAppearance.BorderColor = Color.FromArgb(136, 135, 128)
            rb.FlatAppearance.CheckedBackColor = clsUiTheme.PrimaryBack
            rb.FlatAppearance.MouseOverBackColor = clsUiTheme.CardBack
            rb.Cursor = Cursors.Hand
        Next
        clsUiTheme.StyleButton(Button_Add, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Delete, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_PortTest, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Save, clsUiTheme.ButtonKind.Primary)
        clsUiTheme.StyleButton(Button_Close, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleGrid(DataGridView_Dest)
        colLineEnd.Items.AddRange(New Object() {clsDestination.LINE_CR, clsDestination.LINE_LF, clsDestination.LINE_CRLF})
        DataGridView_Dest.Columns.Insert(colEnabled.Index, colEndMessage)
        DataGridView_Dest.Columns.Insert(colEnabled.Index, colLineEnd)
        colProcess.HeaderText = clsLang.L("프로세스명 / exe 경로", "プロセス名 / exeパス")
        colPort.Width = 90

        'COM ComboBox 목록 (clsComChoices) - 지금 설정된 값도 목록에 넣는다. 통신설정은 직접 입력 (2026-10-28 지시)
        Dim curPorts As New List(Of String) From {mSetting.Reader.Port}
        For Each d As clsDestination In mSetting.Destinations
            curPorts.Add(d.Port)
        Next
        ComboBox_ReaderPort.Items.AddRange(clsComChoices.PortList(curPorts.ToArray))
        colPort.Items.AddRange(clsComChoices.PortList(curPorts.ToArray))
        TextBox_Equipment.Text = mSetting.EquipmentName

        RadioButton_COM.Checked = (mSetting.Reader.InputMode = clsReaderSetting.MODE_COM)
        RadioButton_HID.Checked = Not RadioButton_COM.Checked
        ComboBox_ReaderPort.Text = mSetting.Reader.Port.ToUpperInvariant
        TextBox_ReaderSetting.Text = mSetting.Reader.Setting

        DataGridView_Dest.Rows.Clear()
        For Each d As clsDestination In mSetting.Destinations
            Dim lineEnd As String = d.LineEnd.ToUpperInvariant
            DataGridView_Dest.Rows.Add(d.DestID, d.DestName, d.Port.ToUpperInvariant, d.Setting, d.ProcessName, d.EndTimeoutSec, d.EndMessage,
                                       If(colLineEnd.Items.Contains(lineEnd), lineEnd, Nothing), d.Enabled)     'COM 은 ComboBox 열 목록과 같게 대문자로
        Next
        Call ModeChanged()
        clsUiTheme.ApplyFont(Me)
    End Sub

    Private Sub RadioButton_Mode_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles RadioButton_HID.CheckedChanged, RadioButton_COM.CheckedChanged
        Call ModeChanged()
    End Sub

    Private Function SelectedMode() As String
        Return If(RadioButton_COM.Checked, clsReaderSetting.MODE_COM, clsReaderSetting.MODE_HID)
    End Function

    Private Sub ModeChanged()
        Dim isCom As Boolean = RadioButton_COM.Checked
        ComboBox_ReaderPort.Enabled = isCom
        TextBox_ReaderSetting.Enabled = isCom
        '선택된 쪽은 진한 파랑 + 흰 글자
        RadioButton_HID.ForeColor = If(RadioButton_HID.Checked, clsUiTheme.HeaderText, clsUiTheme.TextMain)
        RadioButton_COM.ForeColor = If(RadioButton_COM.Checked, clsUiTheme.HeaderText, clsUiTheme.TextMain)
    End Sub

    Private Sub Button_Add_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Add.Click
        Dim maxID As Integer = 0
        For Each r As DataGridViewRow In DataGridView_Dest.Rows
            maxID = Math.Max(maxID, CInt(r.Cells(colID.Index).Value))
        Next
        '새 행의 COM 은 아직 아무도 안 쓰는 첫 번호 (ComboBox 열은 목록에 없는 값을 넣을 수 없으므로 빈 값 대신)
        Dim used As New List(Of String)
        If RadioButton_COM.Checked Then used.Add(ComboBox_ReaderPort.Text.Trim.ToUpperInvariant)
        For Each r As DataGridViewRow In DataGridView_Dest.Rows
            used.Add(CellText(r, colPort.Index).ToUpperInvariant)
        Next
        Dim newPort As String = ""
        For Each p As Object In colPort.Items
            If Not used.Contains(p.ToString) Then newPort = p.ToString : Exit For
        Next
        DataGridView_Dest.Rows.Add(maxID + 1, "", If(newPort = "", Nothing, newPort), "", "", 0, "", Nothing, True)      '통신설정 / 종료 신호 / 끝 문자는 비워 둔다 (기본값을 두지 않음)
    End Sub

    Private Sub DataGridView_Dest_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles DataGridView_Dest.DataError
        'ComboBox 열에 목록 밖의 값이 들어온 경우 - 기본 오류 창 대신 무시 (저장할 때 값 확인을 다시 한다)
        e.ThrowException = False
    End Sub

    Private Sub Button_Delete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Delete.Click
        If DataGridView_Dest.CurrentRow Is Nothing Then Exit Sub
        Dim id As Integer = CInt(DataGridView_Dest.CurrentRow.Cells(colID.Index).Value)
        For Each ru As clsRouteRule In mSetting.Rules
            If ru.DestID = id Then
                MessageBox.Show(clsLang.L("분기 조건 '", "分岐条件 '") & ru.Description & clsLang.L("' 에서 쓰고 있어서 지울 수 없습니다.", "' で使用中のため削除できません。") & vbCrLf & clsLang.L("먼저 분기 조건 설정에서 바꿔 주세요.", "先に分岐条件設定で変更してください。"), clsLang.L("COM 설정", "COM設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Next
        DataGridView_Dest.Rows.Remove(DataGridView_Dest.CurrentRow)
    End Sub

    Private Sub Button_PortTest_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_PortTest.Click
        DataGridView_Dest.EndEdit()
        If DataGridView_Dest.CurrentRow Is Nothing Then Exit Sub
        Dim port As String = CellText(DataGridView_Dest.CurrentRow, colPort.Index)
        Dim setting As String = CellText(DataGridView_Dest.CurrentRow, colSetting.Index)
        Dim p As New clsRouterPort
        Try
            p.Open(port, setting)
            p.Close()
            MessageBox.Show(port & clsLang.L(" 열기 / 닫기 OK", " オープン / クローズ OK"), clsLang.L("포트 시험", "ポート試験"), MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(port & clsLang.L(" 열기 NG", " オープン NG") & vbCrLf & ex.Message, clsLang.L("포트 시험", "ポート試験"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        DataGridView_Dest.EndEdit()

        Dim reader As New clsReaderSetting
        reader.InputMode = SelectedMode()
        reader.Port = ComboBox_ReaderPort.Text.Trim.ToUpperInvariant
        reader.Setting = TextBox_ReaderSetting.Text.Trim
        If reader.InputMode = clsReaderSetting.MODE_COM AndAlso (reader.Port = "" OrElse reader.Setting.Split("-"c).Length < 4) Then
            MessageBox.Show(clsLang.L("바코드리더 COM 과 통신설정(예 : 9600-N-8-1)을 넣어 주세요.", "バーコードリーダのCOMと通信設定(例 : 9600-N-8-1)を入力してください。"), clsLang.L("COM 설정", "COM設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        Dim lst As New List(Of clsDestination)
        Dim usedPorts As New List(Of String)
        If reader.InputMode = clsReaderSetting.MODE_COM Then usedPorts.Add(reader.Port.ToUpperInvariant)

        For Each r As DataGridViewRow In DataGridView_Dest.Rows
            Dim d As New clsDestination
            d.DestID = CInt(r.Cells(colID.Index).Value)
            d.DestName = CellText(r, colName.Index)
            d.Port = CellText(r, colPort.Index)
            d.Setting = CellText(r, colSetting.Index)
            d.ProcessName = CellText(r, colProcess.Index)
            Dim t As Integer = 0
            Integer.TryParse(CellText(r, colTimeout.Index), t)
            d.EndTimeoutSec = Math.Max(0, t)
            d.Enabled = CBool(If(r.Cells(colEnabled.Index).Value, False))
            d.EndMessage = CellText(r, colEndMessage.Index)
            d.LineEnd = CellText(r, colLineEnd.Index)
            If d.ProcessName.IndexOf("\"c) < 0 Then d.ProcessName = d.ProcessName.Replace(".exe", "").Replace(".EXE", "")

            If d.DestName = "" OrElse d.Port = "" OrElse d.Setting.Split("-"c).Length < 4 Then
                MessageBox.Show("No " & d.DestID & clsLang.L(" : 이름, COM, 통신설정(예 : 9600-N-8-1)을 넣어 주세요.", " : 名前、COM、通信設定(例 : 9600-N-8-1)を入力してください。"), clsLang.L("COM 설정", "COM設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If d.EndMessage = "" OrElse d.LineEnd = "" Then
                MessageBox.Show("No " & d.DestID & clsLang.L(" : 종료 신호(예 : END)와 끝 문자(CR / LF / CRLF)를 넣어 주세요.", " : 終了信号(例 : END)と終端文字(CR / LF / CRLF)を入力してください。") & vbCrLf & clsLang.L("종료 신호가 없으면 검사가 끝난 것을 알 수 없습니다.", "終了信号がないと検査の終了が分かりません。"), clsLang.L("COM 설정", "COM設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If d.Enabled Then
                If usedPorts.Contains(d.Port.ToUpperInvariant) Then
                    MessageBox.Show(d.Port & clsLang.L(" 를 두 곳에서 쓰고 있습니다.", " が2か所で使われています。"), clsLang.L("COM 설정", "COM設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
                usedPorts.Add(d.Port.ToUpperInvariant)
            End If
            lst.Add(d)
        Next

        Try
            mSetting.EquipmentName = TextBox_Equipment.Text.Trim
            mSetting.Reader = reader
            mSetting.Destinations = lst
            mSetting.SaveReaderAndDestinations()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(clsLang.L("저장 NG", "保存 NG") & vbCrLf & ex.Message, clsLang.L("COM 설정", "COM設定"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Function CellText(ByVal r As DataGridViewRow, ByVal col As Integer) As String
        Dim v As Object = r.Cells(col).Value
        If v Is Nothing Then Return ""
        Return v.ToString.Trim
    End Function


    Private Sub ApplyLanguageTexts()
        Me.colTimeout.ToolTipText = clsLang.L("0 = 무제한", "0 = 無制限")
        '==========================================================================================
        'Function : Designer 에 적힌 한국어 글자를 언어 설정(clsLang)에 맞게 바꾼다 (Designer 파일은 그대로, 2026-10-28)
        '==========================================================================================
        Me.Label_Header.Text = clsLang.L("⇄  COM 설정", "⇄  COM設定")
        Me.Label_PcPorts.Text = clsLang.L("PC 의 COM :", "PCのCOM :")
        Me.Label_EquipCap.Text = clsLang.L("설비명", "設備名")
        Me.Label_EquipHint.Text = clsLang.L("이 Router 가 있는 설비 (메인 화면 제목줄 가운데에 표시, 예 : 내전압 검사장치 2호기)", "このRouterがある設備 (メイン画面タイトル中央に表示、例 : 耐電圧検査装置 2号機)")
        Me.Label_ReaderCap.Text = clsLang.L("바코드리더", "バーコードリーダ")
        Me.Label_ReaderSettingCap.Text = clsLang.L("통신설정", "通信設定")
        Me.Label_DestCap.Text = clsLang.L("분기 대상  · Router 쪽 COM (가상 COM 쌍의 한쪽)", "分岐先  · Router側COM (仮想COMペアの片側)")
        Me.colName.HeaderText = clsLang.L("이름", "名前")
        Me.colSetting.HeaderText = clsLang.L("통신설정", "通信設定")
        Me.colProcess.HeaderText = clsLang.L("프로세스명 (창 앞으로)", "プロセス名 (ウィンドウを前面へ)")
        Me.colTimeout.HeaderText = clsLang.L("END 대기(초)", "END待ち(秒)")
        Me.colEnabled.HeaderText = clsLang.L("사용", "使用")
        Me.Button_Add.Text = clsLang.L("＋  추가", "＋  追加")
        Me.Button_Delete.Text = clsLang.L("－  삭제", "－  削除")
        Me.Button_PortTest.Text = clsLang.L("⇄  포트 시험", "⇄  ポート試験")
        Me.Button_Save.Text = clsLang.L("저장", "保存")
        Me.Button_Close.Text = clsLang.L("닫기", "閉じる")
        Me.Text = clsLang.L("COM 설정", "COM設定")
    End Sub

End Class
