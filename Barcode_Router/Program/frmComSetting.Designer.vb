<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmComSetting
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
        Me.Label_PcPorts = New System.Windows.Forms.Label()
        Me.Label_Header = New System.Windows.Forms.Label()
        Me.Label_EquipCap = New System.Windows.Forms.Label()
        Me.TextBox_Equipment = New System.Windows.Forms.TextBox()
        Me.Label_EquipHint = New System.Windows.Forms.Label()
        Me.Label_ReaderCap = New System.Windows.Forms.Label()
        Me.RadioButton_HID = New System.Windows.Forms.RadioButton()
        Me.RadioButton_COM = New System.Windows.Forms.RadioButton()
        Me.Label_ReaderPortCap = New System.Windows.Forms.Label()
        Me.ComboBox_ReaderPort = New System.Windows.Forms.ComboBox()
        Me.Label_ReaderSettingCap = New System.Windows.Forms.Label()
        Me.TextBox_ReaderSetting = New System.Windows.Forms.TextBox()
        Me.Label_DestCap = New System.Windows.Forms.Label()
        Me.DataGridView_Dest = New System.Windows.Forms.DataGridView()
        Me.colID = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colPort = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.colSetting = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colProcess = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colTimeout = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.colEnabled = New System.Windows.Forms.DataGridViewCheckBoxColumn()
        Me.Button_Add = New System.Windows.Forms.Button()
        Me.Button_Delete = New System.Windows.Forms.Button()
        Me.Button_PortTest = New System.Windows.Forms.Button()
        Me.Button_Save = New System.Windows.Forms.Button()
        Me.Button_Close = New System.Windows.Forms.Button()
        Me.Panel_Header.SuspendLayout()
        CType(Me.DataGridView_Dest, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Panel_Header
        '
        Me.Panel_Header.BackColor = System.Drawing.Color.FromArgb(CType(CType(12, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(124, Byte), Integer))
        Me.Panel_Header.Controls.Add(Me.Label_PcPorts)
        Me.Panel_Header.Controls.Add(Me.Label_Header)
        Me.Panel_Header.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Header.Name = "Panel_Header"
        Me.Panel_Header.Size = New System.Drawing.Size(900, 42)
        '
        'Label_Header
        '
        Me.Label_Header.AutoSize = True
        Me.Label_Header.Font = New System.Drawing.Font("맑은 고딕", 12.0!, System.Drawing.FontStyle.Bold)
        Me.Label_Header.ForeColor = System.Drawing.Color.FromArgb(CType(CType(230, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(251, Byte), Integer))
        Me.Label_Header.Location = New System.Drawing.Point(12, 9)
        Me.Label_Header.Name = "Label_Header"
        Me.Label_Header.Text = "⇄  COM 설정"
        '
        'Label_PcPorts
        '
        Me.Label_PcPorts.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Label_PcPorts.ForeColor = System.Drawing.Color.FromArgb(CType(CType(181, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(244, Byte), Integer))
        Me.Label_PcPorts.Location = New System.Drawing.Point(300, 11)
        Me.Label_PcPorts.Name = "Label_PcPorts"
        Me.Label_PcPorts.Size = New System.Drawing.Size(586, 20)
        Me.Label_PcPorts.Text = "PC 의 COM :"
        Me.Label_PcPorts.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label_EquipCap
        '
        Me.Label_EquipCap.AutoSize = True
        Me.Label_EquipCap.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_EquipCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_EquipCap.Location = New System.Drawing.Point(12, 60)
        Me.Label_EquipCap.Name = "Label_EquipCap"
        Me.Label_EquipCap.Text = "설비명"
        '
        'TextBox_Equipment
        '
        Me.TextBox_Equipment.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_Equipment.Font = New System.Drawing.Font("맑은 고딕", 11.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_Equipment.Location = New System.Drawing.Point(86, 55)
        Me.TextBox_Equipment.MaxLength = 50
        Me.TextBox_Equipment.Name = "TextBox_Equipment"
        Me.TextBox_Equipment.Size = New System.Drawing.Size(330, 27)
        Me.TextBox_Equipment.TabIndex = 0
        '
        'Label_EquipHint
        '
        Me.Label_EquipHint.Font = New System.Drawing.Font("맑은 고딕", 8.5!)
        Me.Label_EquipHint.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_EquipHint.Location = New System.Drawing.Point(424, 60)
        Me.Label_EquipHint.Name = "Label_EquipHint"
        Me.Label_EquipHint.Size = New System.Drawing.Size(462, 20)
        Me.Label_EquipHint.Text = "이 Router 가 있는 설비 (메인 화면 제목줄 가운데에 표시, 예 : 내전압 검사장치 2호기)"
        '
        'Label_ReaderCap
        '
        Me.Label_ReaderCap.AutoSize = True
        Me.Label_ReaderCap.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_ReaderCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_ReaderCap.Location = New System.Drawing.Point(12, 100)
        Me.Label_ReaderCap.Name = "Label_ReaderCap"
        Me.Label_ReaderCap.Text = "바코드리더"
        '
        'RadioButton_HID
        '
        Me.RadioButton_HID.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton_HID.Location = New System.Drawing.Point(14, 124)
        Me.RadioButton_HID.Name = "RadioButton_HID"
        Me.RadioButton_HID.Size = New System.Drawing.Size(72, 30)
        Me.RadioButton_HID.TabIndex = 0
        Me.RadioButton_HID.Text = "HID"
        Me.RadioButton_HID.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RadioButton_COM
        '
        Me.RadioButton_COM.Appearance = System.Windows.Forms.Appearance.Button
        Me.RadioButton_COM.Location = New System.Drawing.Point(85, 124)
        Me.RadioButton_COM.Name = "RadioButton_COM"
        Me.RadioButton_COM.Size = New System.Drawing.Size(72, 30)
        Me.RadioButton_COM.TabIndex = 1
        Me.RadioButton_COM.Text = "COM"
        Me.RadioButton_COM.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_ReaderPortCap
        '
        Me.Label_ReaderPortCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_ReaderPortCap.Location = New System.Drawing.Point(176, 128)
        Me.Label_ReaderPortCap.Name = "Label_ReaderPortCap"
        Me.Label_ReaderPortCap.Size = New System.Drawing.Size(40, 22)
        Me.Label_ReaderPortCap.Text = "COM"
        '
        'ComboBox_ReaderPort
        '
        Me.ComboBox_ReaderPort.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboBox_ReaderPort.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.ComboBox_ReaderPort.Location = New System.Drawing.Point(218, 126)
        Me.ComboBox_ReaderPort.Name = "ComboBox_ReaderPort"
        Me.ComboBox_ReaderPort.Size = New System.Drawing.Size(120, 25)
        Me.ComboBox_ReaderPort.TabIndex = 2
        '
        'Label_ReaderSettingCap
        '
        Me.Label_ReaderSettingCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_ReaderSettingCap.Location = New System.Drawing.Point(356, 128)
        Me.Label_ReaderSettingCap.Name = "Label_ReaderSettingCap"
        Me.Label_ReaderSettingCap.Size = New System.Drawing.Size(64, 22)
        Me.Label_ReaderSettingCap.Text = "통신설정"
        '
        'TextBox_ReaderSetting
        '
        Me.TextBox_ReaderSetting.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_ReaderSetting.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.TextBox_ReaderSetting.Location = New System.Drawing.Point(422, 126)
        Me.TextBox_ReaderSetting.Name = "TextBox_ReaderSetting"
        Me.TextBox_ReaderSetting.Size = New System.Drawing.Size(140, 25)
        Me.TextBox_ReaderSetting.TabIndex = 3
        '
        'Label_DestCap
        '
        Me.Label_DestCap.AutoSize = True
        Me.Label_DestCap.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_DestCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_DestCap.Location = New System.Drawing.Point(12, 172)
        Me.Label_DestCap.Name = "Label_DestCap"
        Me.Label_DestCap.Text = "분기 대상  · Router 쪽 COM (가상 COM 쌍의 한쪽)"
        '
        'DataGridView_Dest
        '
        Me.DataGridView_Dest.AllowUserToAddRows = False
        Me.DataGridView_Dest.AllowUserToDeleteRows = False
        Me.DataGridView_Dest.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colID, Me.colName, Me.colPort, Me.colSetting, Me.colProcess, Me.colTimeout, Me.colEnabled})
        Me.DataGridView_Dest.Location = New System.Drawing.Point(14, 196)
        Me.DataGridView_Dest.MultiSelect = False
        Me.DataGridView_Dest.Name = "DataGridView_Dest"
        Me.DataGridView_Dest.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView_Dest.Size = New System.Drawing.Size(872, 220)
        Me.DataGridView_Dest.TabIndex = 4
        '
        'colID
        '
        Me.colID.HeaderText = "No"
        Me.colID.Name = "colID"
        Me.colID.ReadOnly = True
        Me.colID.Width = 44
        '
        'colName
        '
        Me.colName.HeaderText = "이름"
        Me.colName.Name = "colName"
        Me.colName.Width = 200
        '
        'colPort
        '
        Me.colPort.HeaderText = "COM"
        Me.colPort.Name = "colPort"
        Me.colPort.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.colPort.Width = 80
        '
        'colSetting
        '
        Me.colSetting.HeaderText = "통신설정"
        Me.colSetting.Name = "colSetting"
        Me.colSetting.Width = 110
        '
        'colProcess
        '
        Me.colProcess.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.colProcess.HeaderText = "프로세스명 (창 앞으로)"
        Me.colProcess.Name = "colProcess"
        '
        'colTimeout
        '
        Me.colTimeout.HeaderText = "END 대기(초)"
        Me.colTimeout.Name = "colTimeout"
        Me.colTimeout.ToolTipText = "0 = 무제한"
        Me.colTimeout.Width = 100
        '
        'colEnabled
        '
        Me.colEnabled.HeaderText = "사용"
        Me.colEnabled.Name = "colEnabled"
        Me.colEnabled.Width = 50
        '
        'Button_Add
        '
        Me.Button_Add.Location = New System.Drawing.Point(14, 426)
        Me.Button_Add.Name = "Button_Add"
        Me.Button_Add.Size = New System.Drawing.Size(96, 32)
        Me.Button_Add.TabIndex = 5
        Me.Button_Add.Text = "＋  추가"
        '
        'Button_Delete
        '
        Me.Button_Delete.Location = New System.Drawing.Point(116, 426)
        Me.Button_Delete.Name = "Button_Delete"
        Me.Button_Delete.Size = New System.Drawing.Size(96, 32)
        Me.Button_Delete.TabIndex = 6
        Me.Button_Delete.Text = "－  삭제"
        '
        'Button_PortTest
        '
        Me.Button_PortTest.Location = New System.Drawing.Point(218, 426)
        Me.Button_PortTest.Name = "Button_PortTest"
        Me.Button_PortTest.Size = New System.Drawing.Size(120, 32)
        Me.Button_PortTest.TabIndex = 7
        Me.Button_PortTest.Text = "⇄  포트 시험"
        '
        'Button_Save
        '
        Me.Button_Save.Font = New System.Drawing.Font("맑은 고딕", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_Save.Location = New System.Drawing.Point(658, 466)
        Me.Button_Save.Name = "Button_Save"
        Me.Button_Save.Size = New System.Drawing.Size(110, 40)
        Me.Button_Save.TabIndex = 8
        Me.Button_Save.Text = "저장"
        '
        'Button_Close
        '
        Me.Button_Close.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Close.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Close.Location = New System.Drawing.Point(776, 466)
        Me.Button_Close.Name = "Button_Close"
        Me.Button_Close.Size = New System.Drawing.Size(110, 40)
        Me.Button_Close.TabIndex = 9
        Me.Button_Close.Text = "닫기"
        '
        'frmComSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.Button_Close
        Me.ClientSize = New System.Drawing.Size(900, 520)
        Me.Controls.Add(Me.Button_Close)
        Me.Controls.Add(Me.Button_Save)
        Me.Controls.Add(Me.Button_PortTest)
        Me.Controls.Add(Me.Button_Delete)
        Me.Controls.Add(Me.Button_Add)
        Me.Controls.Add(Me.DataGridView_Dest)
        Me.Controls.Add(Me.Label_DestCap)
        Me.Controls.Add(Me.TextBox_ReaderSetting)
        Me.Controls.Add(Me.Label_ReaderSettingCap)
        Me.Controls.Add(Me.ComboBox_ReaderPort)
        Me.Controls.Add(Me.Label_ReaderPortCap)
        Me.Controls.Add(Me.RadioButton_COM)
        Me.Controls.Add(Me.RadioButton_HID)
        Me.Controls.Add(Me.Label_ReaderCap)
        Me.Controls.Add(Me.Label_EquipHint)
        Me.Controls.Add(Me.TextBox_Equipment)
        Me.Controls.Add(Me.Label_EquipCap)
        Me.Controls.Add(Me.Panel_Header)
        Me.Font = New System.Drawing.Font("맑은 고딕", 9.5!)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmComSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "COM 설정"
        Me.Panel_Header.ResumeLayout(False)
        Me.Panel_Header.PerformLayout()
        CType(Me.DataGridView_Dest, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel_Header As System.Windows.Forms.Panel
    Friend WithEvents Label_Header As System.Windows.Forms.Label
    Friend WithEvents Label_PcPorts As System.Windows.Forms.Label
    Friend WithEvents Label_EquipCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_Equipment As System.Windows.Forms.TextBox
    Friend WithEvents Label_EquipHint As System.Windows.Forms.Label
    Friend WithEvents Label_ReaderCap As System.Windows.Forms.Label
    Friend WithEvents RadioButton_HID As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton_COM As System.Windows.Forms.RadioButton
    Friend WithEvents Label_ReaderPortCap As System.Windows.Forms.Label
    Friend WithEvents ComboBox_ReaderPort As System.Windows.Forms.ComboBox
    Friend WithEvents Label_ReaderSettingCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_ReaderSetting As System.Windows.Forms.TextBox
    Friend WithEvents Label_DestCap As System.Windows.Forms.Label
    Friend WithEvents DataGridView_Dest As System.Windows.Forms.DataGridView
    Friend WithEvents colID As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colPort As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colSetting As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colProcess As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colTimeout As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents colEnabled As System.Windows.Forms.DataGridViewCheckBoxColumn
    Friend WithEvents Button_Add As System.Windows.Forms.Button
    Friend WithEvents Button_Delete As System.Windows.Forms.Button
    Friend WithEvents Button_PortTest As System.Windows.Forms.Button
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents Button_Close As System.Windows.Forms.Button
End Class
