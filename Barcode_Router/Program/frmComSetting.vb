Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmComSetting
'Function   : COM 설정 화면 (바코드리더 입력 방식 / 포트, 분기 대상별 COM)
'Framwork   : VB2022 Framework4.8
'Memo       : - 설비명     : 이 Router 가 있는 설비 이름 (메인 화면 제목줄 가운데 표시, 2026-10-26 추가)
'             - 바코드리더 : 입력 방식 HID(바코드 칸 키보드 입력) / COM, 리더 포트, 통신설정 (구조설계 C-01, C-02)
'             - 분기 대상 : 이름, Router 쪽 COM, 통신설정, 프로세스명(창 앞으로), END 대기 초, 사용 (구조설계 C-03)
'             - 포트 시험 : 선택한 행의 COM 을 열었다 닫아 본다 (구조설계 C-04)
'             - 검사 프로그램 쪽 COM 은 각 프로그램의 EquipmentSetting.mdb > tblEquipmentSetting > BarcodeReaderPort 로 설정
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-10항 (COM 은 따로 설정할 수 있는 설정 화면)
'==========================================================

Public Class frmComSetting

    Private mSetting As clsRouterSetting

    Public Sub New(ByVal setting As clsRouterSetting)
        InitializeComponent()
        mSetting = setting
    End Sub

    Private Sub frmComSetting_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Dim pcPorts() As String = IO.Ports.SerialPort.GetPortNames()
        Array.Sort(pcPorts)
        Label_PcPorts.Text = "PC 의 COM : " & If(pcPorts.Length = 0, "(없음)", String.Join(", ", pcPorts))

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

        ComboBox_ReaderPort.Items.AddRange(pcPorts)
        TextBox_Equipment.Text = mSetting.EquipmentName

        RadioButton_COM.Checked = (mSetting.Reader.InputMode = clsReaderSetting.MODE_COM)
        RadioButton_HID.Checked = Not RadioButton_COM.Checked
        ComboBox_ReaderPort.Text = mSetting.Reader.Port
        TextBox_ReaderSetting.Text = mSetting.Reader.Setting

        DataGridView_Dest.Rows.Clear()
        For Each d As clsDestination In mSetting.Destinations
            DataGridView_Dest.Rows.Add(d.DestID, d.DestName, d.Port, d.Setting, d.ProcessName, d.EndTimeoutSec, d.Enabled)
        Next
        Call ModeChanged()
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
        DataGridView_Dest.Rows.Add(maxID + 1, "", "", "9600-N-8-1", "", 0, True)
    End Sub

    Private Sub Button_Delete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Delete.Click
        If DataGridView_Dest.CurrentRow Is Nothing Then Exit Sub
        Dim id As Integer = CInt(DataGridView_Dest.CurrentRow.Cells(colID.Index).Value)
        For Each ru As clsRouteRule In mSetting.Rules
            If ru.DestID = id Then
                MessageBox.Show("분기 조건 '" & ru.Description & "' 에서 쓰고 있어서 지울 수 없습니다." & vbCrLf & "먼저 분기 조건 설정에서 바꿔 주세요.", "COM 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            MessageBox.Show(port & " 열기 / 닫기 OK", "포트 시험", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show(port & " 열기 NG" & vbCrLf & ex.Message, "포트 시험", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        DataGridView_Dest.EndEdit()

        Dim reader As New clsReaderSetting
        reader.InputMode = SelectedMode()
        reader.Port = ComboBox_ReaderPort.Text.Trim
        reader.Setting = TextBox_ReaderSetting.Text.Trim
        If reader.InputMode = clsReaderSetting.MODE_COM AndAlso (reader.Port = "" OrElse reader.Setting.Split("-"c).Length < 4) Then
            MessageBox.Show("바코드리더 COM 과 통신설정(예 : 9600-N-8-1)을 넣어 주세요.", "COM 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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

            If d.DestName = "" OrElse d.Port = "" OrElse d.Setting.Split("-"c).Length < 4 Then
                MessageBox.Show("No " & d.DestID & " : 이름, COM, 통신설정(예 : 9600-N-8-1)을 넣어 주세요.", "COM 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If d.Enabled Then
                If usedPorts.Contains(d.Port.ToUpperInvariant) Then
                    MessageBox.Show(d.Port & " 를 두 곳에서 쓰고 있습니다.", "COM 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            MessageBox.Show("저장 NG" & vbCrLf & ex.Message, "COM 설정", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Shared Function CellText(ByVal r As DataGridViewRow, ByVal col As Integer) As String
        Dim v As Object = r.Cells(col).Value
        If v Is Nothing Then Return ""
        Return v.ToString.Trim
    End Function

End Class
