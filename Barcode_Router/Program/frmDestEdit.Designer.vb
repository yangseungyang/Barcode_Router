<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmDestEdit
    Inherits System.Windows.Forms.Form

    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    Private components As System.ComponentModel.IContainer

    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Panel_Header = New System.Windows.Forms.Panel()
        Me.Label_Header = New System.Windows.Forms.Label()
        Me.Label_NameCap = New System.Windows.Forms.Label()
        Me.TextBox_Name = New System.Windows.Forms.TextBox()
        Me.Label_PortCap = New System.Windows.Forms.Label()
        Me.ComboBox_Port = New System.Windows.Forms.ComboBox()
        Me.Label_PcPorts = New System.Windows.Forms.Label()
        Me.Label_SettingCap = New System.Windows.Forms.Label()
        Me.TextBox_Setting = New System.Windows.Forms.TextBox()
        Me.Label_ProcessCap = New System.Windows.Forms.Label()
        Me.TextBox_Process = New System.Windows.Forms.TextBox()
        Me.Label_ProcessHint = New System.Windows.Forms.Label()
        Me.Label_TimeoutCap = New System.Windows.Forms.Label()
        Me.NumericUpDown_Timeout = New System.Windows.Forms.NumericUpDown()
        Me.Label_TimeoutHint = New System.Windows.Forms.Label()
        Me.CheckBox_Enabled = New System.Windows.Forms.CheckBox()
        Me.Button_PortTest = New System.Windows.Forms.Button()
        Me.Button_Save = New System.Windows.Forms.Button()
        Me.Button_Close = New System.Windows.Forms.Button()
        Me.Panel_Header.SuspendLayout()
        CType(Me.NumericUpDown_Timeout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel_Header
        '
        Me.Panel_Header.BackColor = System.Drawing.Color.FromArgb(CType(CType(12, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(124, Byte), Integer))
        Me.Panel_Header.Controls.Add(Me.Label_Header)
        Me.Panel_Header.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Header.Name = "Panel_Header"
        Me.Panel_Header.Size = New System.Drawing.Size(520, 42)
        '
        'Label_Header
        '
        Me.Label_Header.AutoSize = True
        Me.Label_Header.Font = New System.Drawing.Font("맑은 고딕", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label_Header.ForeColor = System.Drawing.Color.FromArgb(CType(CType(230, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(251, Byte), Integer))
        Me.Label_Header.Location = New System.Drawing.Point(12, 9)
        Me.Label_Header.Name = "Label_Header"
        Me.Label_Header.Text = "⇄  분기 COM"
        '
        'Label_NameCap
        '
        Me.Label_NameCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_NameCap.Location = New System.Drawing.Point(16, 60)
        Me.Label_NameCap.Name = "Label_NameCap"
        Me.Label_NameCap.Size = New System.Drawing.Size(90, 22)
        Me.Label_NameCap.Text = "이름"
        '
        'TextBox_Name
        '
        Me.TextBox_Name.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_Name.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.TextBox_Name.Location = New System.Drawing.Point(110, 57)
        Me.TextBox_Name.Name = "TextBox_Name"
        Me.TextBox_Name.Size = New System.Drawing.Size(392, 25)
        Me.TextBox_Name.TabIndex = 0
        '
        'Label_PortCap
        '
        Me.Label_PortCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_PortCap.Location = New System.Drawing.Point(16, 96)
        Me.Label_PortCap.Name = "Label_PortCap"
        Me.Label_PortCap.Size = New System.Drawing.Size(90, 22)
        Me.Label_PortCap.Text = "COM"
        '
        'ComboBox_Port
        '
        Me.ComboBox_Port.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboBox_Port.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.ComboBox_Port.Location = New System.Drawing.Point(110, 93)
        Me.ComboBox_Port.Name = "ComboBox_Port"
        Me.ComboBox_Port.Size = New System.Drawing.Size(130, 25)
        Me.ComboBox_Port.TabIndex = 1
        '
        'Label_PcPorts
        '
        Me.Label_PcPorts.Font = New System.Drawing.Font("맑은 고딕", 8.5!)
        Me.Label_PcPorts.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_PcPorts.Location = New System.Drawing.Point(248, 96)
        Me.Label_PcPorts.Name = "Label_PcPorts"
        Me.Label_PcPorts.Size = New System.Drawing.Size(254, 22)
        Me.Label_PcPorts.Text = "직접 입력 가능 (예 : COM15)"
        '
        'Label_SettingCap
        '
        Me.Label_SettingCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_SettingCap.Location = New System.Drawing.Point(16, 132)
        Me.Label_SettingCap.Name = "Label_SettingCap"
        Me.Label_SettingCap.Size = New System.Drawing.Size(90, 22)
        Me.Label_SettingCap.Text = "통신설정"
        '
        'TextBox_Setting
        '
        Me.TextBox_Setting.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_Setting.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.TextBox_Setting.Location = New System.Drawing.Point(110, 129)
        Me.TextBox_Setting.Name = "TextBox_Setting"
        Me.TextBox_Setting.Size = New System.Drawing.Size(130, 25)
        Me.TextBox_Setting.TabIndex = 2
        '
        'Label_ProcessCap
        '
        Me.Label_ProcessCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_ProcessCap.Location = New System.Drawing.Point(16, 168)
        Me.Label_ProcessCap.Name = "Label_ProcessCap"
        Me.Label_ProcessCap.Size = New System.Drawing.Size(90, 22)
        Me.Label_ProcessCap.Text = "프로세스명"
        '
        'TextBox_Process
        '
        Me.TextBox_Process.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_Process.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.TextBox_Process.Location = New System.Drawing.Point(110, 165)
        Me.TextBox_Process.Name = "TextBox_Process"
        Me.TextBox_Process.Size = New System.Drawing.Size(392, 25)
        Me.TextBox_Process.TabIndex = 3
        '
        'Label_ProcessHint
        '
        Me.Label_ProcessHint.Font = New System.Drawing.Font("맑은 고딕", 8.5!)
        Me.Label_ProcessHint.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_ProcessHint.Location = New System.Drawing.Point(110, 192)
        Me.Label_ProcessHint.Name = "Label_ProcessHint"
        Me.Label_ProcessHint.Size = New System.Drawing.Size(392, 18)
        Me.Label_ProcessHint.Text = "바코드를 보낼 때 앞으로 띄울 프로그램 (.exe 없이, 예 : AIEC_HIPOT)"
        '
        'Label_TimeoutCap
        '
        Me.Label_TimeoutCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_TimeoutCap.Location = New System.Drawing.Point(16, 222)
        Me.Label_TimeoutCap.Name = "Label_TimeoutCap"
        Me.Label_TimeoutCap.Size = New System.Drawing.Size(90, 22)
        Me.Label_TimeoutCap.Text = "END 대기"
        '
        'NumericUpDown_Timeout
        '
        Me.NumericUpDown_Timeout.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.NumericUpDown_Timeout.Location = New System.Drawing.Point(110, 219)
        Me.NumericUpDown_Timeout.Maximum = New Decimal(New Integer() {3600, 0, 0, 0})
        Me.NumericUpDown_Timeout.Name = "NumericUpDown_Timeout"
        Me.NumericUpDown_Timeout.Size = New System.Drawing.Size(80, 25)
        Me.NumericUpDown_Timeout.TabIndex = 4
        '
        'Label_TimeoutHint
        '
        Me.Label_TimeoutHint.Font = New System.Drawing.Font("맑은 고딕", 8.5!)
        Me.Label_TimeoutHint.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_TimeoutHint.Location = New System.Drawing.Point(198, 222)
        Me.Label_TimeoutHint.Name = "Label_TimeoutHint"
        Me.Label_TimeoutHint.Size = New System.Drawing.Size(200, 22)
        Me.Label_TimeoutHint.Text = "초  (0 = 무제한)"
        '
        'CheckBox_Enabled
        '
        Me.CheckBox_Enabled.Location = New System.Drawing.Point(110, 254)
        Me.CheckBox_Enabled.Name = "CheckBox_Enabled"
        Me.CheckBox_Enabled.Size = New System.Drawing.Size(120, 24)
        Me.CheckBox_Enabled.TabIndex = 5
        Me.CheckBox_Enabled.Text = "사용"
        '
        'Button_PortTest
        '
        Me.Button_PortTest.Location = New System.Drawing.Point(16, 300)
        Me.Button_PortTest.Name = "Button_PortTest"
        Me.Button_PortTest.Size = New System.Drawing.Size(120, 36)
        Me.Button_PortTest.TabIndex = 6
        Me.Button_PortTest.Text = "⇄  포트 시험"
        '
        'Button_Save
        '
        Me.Button_Save.Font = New System.Drawing.Font("맑은 고딕", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_Save.Location = New System.Drawing.Point(296, 298)
        Me.Button_Save.Name = "Button_Save"
        Me.Button_Save.Size = New System.Drawing.Size(100, 40)
        Me.Button_Save.TabIndex = 7
        Me.Button_Save.Text = "저장"
        '
        'Button_Close
        '
        Me.Button_Close.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Close.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Close.Location = New System.Drawing.Point(402, 298)
        Me.Button_Close.Name = "Button_Close"
        Me.Button_Close.Size = New System.Drawing.Size(100, 40)
        Me.Button_Close.TabIndex = 8
        Me.Button_Close.Text = "닫기"
        '
        'frmDestEdit
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.Button_Close
        Me.ClientSize = New System.Drawing.Size(520, 352)
        Me.Controls.Add(Me.Button_Close)
        Me.Controls.Add(Me.Button_Save)
        Me.Controls.Add(Me.Button_PortTest)
        Me.Controls.Add(Me.CheckBox_Enabled)
        Me.Controls.Add(Me.Label_TimeoutHint)
        Me.Controls.Add(Me.NumericUpDown_Timeout)
        Me.Controls.Add(Me.Label_TimeoutCap)
        Me.Controls.Add(Me.Label_ProcessHint)
        Me.Controls.Add(Me.TextBox_Process)
        Me.Controls.Add(Me.Label_ProcessCap)
        Me.Controls.Add(Me.TextBox_Setting)
        Me.Controls.Add(Me.Label_SettingCap)
        Me.Controls.Add(Me.Label_PcPorts)
        Me.Controls.Add(Me.ComboBox_Port)
        Me.Controls.Add(Me.Label_PortCap)
        Me.Controls.Add(Me.TextBox_Name)
        Me.Controls.Add(Me.Label_NameCap)
        Me.Controls.Add(Me.Panel_Header)
        Me.Font = New System.Drawing.Font("맑은 고딕", 9.5!)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDestEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "분기 COM"
        Me.Panel_Header.ResumeLayout(False)
        Me.Panel_Header.PerformLayout()
        CType(Me.NumericUpDown_Timeout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel_Header As System.Windows.Forms.Panel
    Friend WithEvents Label_Header As System.Windows.Forms.Label
    Friend WithEvents Label_NameCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_Name As System.Windows.Forms.TextBox
    Friend WithEvents Label_PortCap As System.Windows.Forms.Label
    Friend WithEvents ComboBox_Port As System.Windows.Forms.ComboBox
    Friend WithEvents Label_PcPorts As System.Windows.Forms.Label
    Friend WithEvents Label_SettingCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_Setting As System.Windows.Forms.TextBox
    Friend WithEvents Label_ProcessCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_Process As System.Windows.Forms.TextBox
    Friend WithEvents Label_ProcessHint As System.Windows.Forms.Label
    Friend WithEvents Label_TimeoutCap As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown_Timeout As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label_TimeoutHint As System.Windows.Forms.Label
    Friend WithEvents CheckBox_Enabled As System.Windows.Forms.CheckBox
    Friend WithEvents Button_PortTest As System.Windows.Forms.Button
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents Button_Close As System.Windows.Forms.Button
End Class
