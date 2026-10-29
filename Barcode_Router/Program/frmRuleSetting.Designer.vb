<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmRuleSetting
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
        Me.Label_RuleCap = New System.Windows.Forms.Label()
        Me.ComboBox_Rule = New System.Windows.Forms.ComboBox()
        Me.Button_New = New System.Windows.Forms.Button()
        Me.Button_Up = New System.Windows.Forms.Button()
        Me.Button_Down = New System.Windows.Forms.Button()
        Me.Button_Delete = New System.Windows.Forms.Button()
        Me.Panel_Rule = New System.Windows.Forms.Panel()
        Me.CheckBox_Enabled = New System.Windows.Forms.CheckBox()
        Me.Label_Priority = New System.Windows.Forms.Label()
        Me.Label_DestCap = New System.Windows.Forms.Label()
        Me.ComboBox_Dest = New System.Windows.Forms.ComboBox()
        Me.Button_DestNew = New System.Windows.Forms.Button()
        Me.Button_DestEdit = New System.Windows.Forms.Button()
        Me.Button_DestDelete = New System.Windows.Forms.Button()
        Me.colJoin = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.Label_DescCap = New System.Windows.Forms.Label()
        Me.TextBox_Desc = New System.Windows.Forms.TextBox()
        Me.Label_CondCap = New System.Windows.Forms.Label()
        Me.DataGridView_Cond = New System.Windows.Forms.DataGridView()
        Me.colTarget = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.colOperator = New System.Windows.Forms.DataGridViewComboBoxColumn()
        Me.colValues = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.Button_CondAdd = New System.Windows.Forms.Button()
        Me.Button_CondDelete = New System.Windows.Forms.Button()
        Me.Label_ValuesTitle = New System.Windows.Forms.Label()
        Me.Label_ValuesCap = New System.Windows.Forms.Label()
        Me.Panel_ValuesBorder = New System.Windows.Forms.Panel()
        Me.TextBox_Values = New System.Windows.Forms.TextBox()
        Me.Panel_Test = New System.Windows.Forms.Panel()
        Me.Label_TestCap = New System.Windows.Forms.Label()
        Me.TextBox_TestMSCODE = New System.Windows.Forms.TextBox()
        Me.Button_Test = New System.Windows.Forms.Button()
        Me.Label_TestResult = New System.Windows.Forms.Label()
        Me.Button_Save = New System.Windows.Forms.Button()
        Me.Button_Close = New System.Windows.Forms.Button()
        Me.Panel_Header.SuspendLayout()
        Me.Panel_Rule.SuspendLayout()
        CType(Me.DataGridView_Cond, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel_ValuesBorder.SuspendLayout()
        Me.Panel_Test.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel_Header
        '
        Me.Panel_Header.BackColor = System.Drawing.Color.FromArgb(CType(CType(12, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(124, Byte), Integer))
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
        Me.Label_Header.Text = "⑂  분기 조건 설정"
        '
        'Label_RuleCap
        '
        Me.Label_RuleCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_RuleCap.Location = New System.Drawing.Point(14, 60)
        Me.Label_RuleCap.Name = "Label_RuleCap"
        Me.Label_RuleCap.Size = New System.Drawing.Size(70, 22)
        Me.Label_RuleCap.Text = "분기 규칙"
        Me.Label_RuleCap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ComboBox_Rule
        '
        Me.ComboBox_Rule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Rule.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboBox_Rule.Font = New System.Drawing.Font("맑은 고딕", 11.0!, System.Drawing.FontStyle.Bold)
        Me.ComboBox_Rule.Location = New System.Drawing.Point(88, 56)
        Me.ComboBox_Rule.Name = "ComboBox_Rule"
        Me.ComboBox_Rule.Size = New System.Drawing.Size(470, 28)
        Me.ComboBox_Rule.TabIndex = 0
        '
        'Button_New
        '
        Me.Button_New.Location = New System.Drawing.Point(568, 55)
        Me.Button_New.Name = "Button_New"
        Me.Button_New.Size = New System.Drawing.Size(92, 32)
        Me.Button_New.TabIndex = 1
        Me.Button_New.Text = "＋  신규"
        '
        'Button_Up
        '
        Me.Button_Up.Location = New System.Drawing.Point(666, 55)
        Me.Button_Up.Name = "Button_Up"
        Me.Button_Up.Size = New System.Drawing.Size(44, 32)
        Me.Button_Up.TabIndex = 2
        Me.Button_Up.Text = "▲"
        '
        'Button_Down
        '
        Me.Button_Down.Location = New System.Drawing.Point(714, 55)
        Me.Button_Down.Name = "Button_Down"
        Me.Button_Down.Size = New System.Drawing.Size(44, 32)
        Me.Button_Down.TabIndex = 3
        Me.Button_Down.Text = "▼"
        '
        'Button_Delete
        '
        Me.Button_Delete.Location = New System.Drawing.Point(786, 55)
        Me.Button_Delete.Name = "Button_Delete"
        Me.Button_Delete.Size = New System.Drawing.Size(100, 32)
        Me.Button_Delete.TabIndex = 4
        Me.Button_Delete.Text = "✕  삭제"
        '
        'Panel_Rule
        '
        Me.Panel_Rule.Controls.Add(Me.Panel_ValuesBorder)
        Me.Panel_Rule.Controls.Add(Me.Label_ValuesCap)
        Me.Panel_Rule.Controls.Add(Me.Label_ValuesTitle)
        Me.Panel_Rule.Controls.Add(Me.Button_CondDelete)
        Me.Panel_Rule.Controls.Add(Me.Button_CondAdd)
        Me.Panel_Rule.Controls.Add(Me.DataGridView_Cond)
        Me.Panel_Rule.Controls.Add(Me.Label_CondCap)
        Me.Panel_Rule.Controls.Add(Me.TextBox_Desc)
        Me.Panel_Rule.Controls.Add(Me.Label_DescCap)
        Me.Panel_Rule.Controls.Add(Me.Button_DestDelete)
        Me.Panel_Rule.Controls.Add(Me.Button_DestEdit)
        Me.Panel_Rule.Controls.Add(Me.Button_DestNew)
        Me.Panel_Rule.Controls.Add(Me.ComboBox_Dest)
        Me.Panel_Rule.Controls.Add(Me.Label_DestCap)
        Me.Panel_Rule.Controls.Add(Me.Label_Priority)
        Me.Panel_Rule.Controls.Add(Me.CheckBox_Enabled)
        Me.Panel_Rule.Location = New System.Drawing.Point(14, 100)
        Me.Panel_Rule.Name = "Panel_Rule"
        Me.Panel_Rule.Size = New System.Drawing.Size(872, 380)
        '
        'CheckBox_Enabled
        '
        Me.CheckBox_Enabled.Location = New System.Drawing.Point(0, 2)
        Me.CheckBox_Enabled.Name = "CheckBox_Enabled"
        Me.CheckBox_Enabled.Size = New System.Drawing.Size(70, 24)
        Me.CheckBox_Enabled.TabIndex = 5
        Me.CheckBox_Enabled.Text = "사용"
        '
        'Label_Priority
        '
        Me.Label_Priority.ForeColor = System.Drawing.Color.FromArgb(CType(CType(39, Byte), Integer), CType(CType(80, Byte), Integer), CType(CType(10, Byte), Integer))
        Me.Label_Priority.Location = New System.Drawing.Point(76, 4)
        Me.Label_Priority.Name = "Label_Priority"
        Me.Label_Priority.Size = New System.Drawing.Size(200, 20)
        Me.Label_Priority.Text = "우선순위 -"
        Me.Label_Priority.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Label_DestCap
        '
        Me.Label_DestCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_DestCap.Location = New System.Drawing.Point(0, 38)
        Me.Label_DestCap.Name = "Label_DestCap"
        Me.Label_DestCap.Size = New System.Drawing.Size(70, 22)
        Me.Label_DestCap.Text = "분기 COM"
        Me.Label_DestCap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'ComboBox_Dest
        '
        Me.ComboBox_Dest.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox_Dest.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.ComboBox_Dest.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.ComboBox_Dest.Location = New System.Drawing.Point(74, 36)
        Me.ComboBox_Dest.Name = "ComboBox_Dest"
        Me.ComboBox_Dest.Size = New System.Drawing.Size(260, 25)
        Me.ComboBox_Dest.TabIndex = 6
        '
        'Button_DestNew
        '
        Me.Button_DestNew.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Button_DestNew.Location = New System.Drawing.Point(340, 35)
        Me.Button_DestNew.Name = "Button_DestNew"
        Me.Button_DestNew.Size = New System.Drawing.Size(66, 27)
        Me.Button_DestNew.TabStop = False
        Me.Button_DestNew.Text = "＋ 신규"
        '
        'Button_DestEdit
        '
        Me.Button_DestEdit.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Button_DestEdit.Location = New System.Drawing.Point(410, 35)
        Me.Button_DestEdit.Name = "Button_DestEdit"
        Me.Button_DestEdit.Size = New System.Drawing.Size(66, 27)
        Me.Button_DestEdit.TabStop = False
        Me.Button_DestEdit.Text = "✎ 수정"
        '
        'Button_DestDelete
        '
        Me.Button_DestDelete.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Button_DestDelete.Location = New System.Drawing.Point(480, 35)
        Me.Button_DestDelete.Name = "Button_DestDelete"
        Me.Button_DestDelete.Size = New System.Drawing.Size(60, 27)
        Me.Button_DestDelete.TabStop = False
        Me.Button_DestDelete.Text = "✕ 삭제"
        '
        'Label_DescCap
        '
        Me.Label_DescCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_DescCap.Location = New System.Drawing.Point(0, 72)
        Me.Label_DescCap.Name = "Label_DescCap"
        Me.Label_DescCap.Size = New System.Drawing.Size(70, 22)
        Me.Label_DescCap.Text = "설명"
        Me.Label_DescCap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_Desc
        '
        Me.TextBox_Desc.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_Desc.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.TextBox_Desc.Location = New System.Drawing.Point(74, 70)
        Me.TextBox_Desc.Name = "TextBox_Desc"
        Me.TextBox_Desc.Size = New System.Drawing.Size(466, 25)
        Me.TextBox_Desc.TabIndex = 7
        '
        'Label_CondCap
        '
        Me.Label_CondCap.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_CondCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_CondCap.Location = New System.Drawing.Point(0, 108)
        Me.Label_CondCap.Name = "Label_CondCap"
        Me.Label_CondCap.Size = New System.Drawing.Size(540, 22)
        Me.Label_CondCap.Text = "조건  · 줄마다 연결 AND / OR 선택 (AND 를 먼저 묶음)"
        Me.Label_CondCap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'DataGridView_Cond
        '
        Me.DataGridView_Cond.AllowUserToAddRows = False
        Me.DataGridView_Cond.AllowUserToDeleteRows = False
        Me.DataGridView_Cond.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.colJoin, Me.colTarget, Me.colOperator, Me.colValues})
        Me.DataGridView_Cond.EditMode = System.Windows.Forms.DataGridViewEditMode.EditOnEnter
        Me.DataGridView_Cond.Location = New System.Drawing.Point(0, 132)
        Me.DataGridView_Cond.MultiSelect = False
        Me.DataGridView_Cond.Name = "DataGridView_Cond"
        Me.DataGridView_Cond.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView_Cond.Size = New System.Drawing.Size(540, 200)
        Me.DataGridView_Cond.TabIndex = 8
        '
        'colJoin
        '
        Me.colJoin.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.colJoin.HeaderText = "연결"
        Me.colJoin.Name = "colJoin"
        Me.colJoin.ToolTipText = "앞 줄과의 연결 (AND / OR). 첫 줄은 없음"
        Me.colJoin.Width = 72
        '
        'colTarget
        '
        Me.colTarget.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.colTarget.HeaderText = "대상"
        Me.colTarget.Name = "colTarget"
        Me.colTarget.Width = 108
        '
        'colOperator
        '
        Me.colOperator.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.colOperator.HeaderText = "조건"
        Me.colOperator.Name = "colOperator"
        Me.colOperator.Width = 92
        '
        'colValues
        '
        Me.colValues.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        Me.colValues.HeaderText = "입력값 (오른쪽 칸에서 입력)"
        Me.colValues.Name = "colValues"
        Me.colValues.ReadOnly = True
        '
        'Button_CondAdd
        '
        Me.Button_CondAdd.Location = New System.Drawing.Point(0, 340)
        Me.Button_CondAdd.Name = "Button_CondAdd"
        Me.Button_CondAdd.Size = New System.Drawing.Size(112, 32)
        Me.Button_CondAdd.TabIndex = 9
        Me.Button_CondAdd.Text = "＋  조건 추가"
        '
        'Button_CondDelete
        '
        Me.Button_CondDelete.Location = New System.Drawing.Point(118, 340)
        Me.Button_CondDelete.Name = "Button_CondDelete"
        Me.Button_CondDelete.Size = New System.Drawing.Size(112, 32)
        Me.Button_CondDelete.TabIndex = 10
        Me.Button_CondDelete.Text = "－  조건 삭제"
        '
        'Label_ValuesTitle
        '
        Me.Label_ValuesTitle.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_ValuesTitle.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_ValuesTitle.Location = New System.Drawing.Point(560, 2)
        Me.Label_ValuesTitle.Name = "Label_ValuesTitle"
        Me.Label_ValuesTitle.Size = New System.Drawing.Size(312, 22)
        Me.Label_ValuesTitle.Text = "입력값 (선택한 조건 줄)"
        '
        'Label_ValuesCap
        '
        Me.Label_ValuesCap.Font = New System.Drawing.Font("맑은 고딕", 8.5!)
        Me.Label_ValuesCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_ValuesCap.Location = New System.Drawing.Point(560, 26)
        Me.Label_ValuesCap.Name = "Label_ValuesCap"
        Me.Label_ValuesCap.Size = New System.Drawing.Size(312, 36)
        Me.Label_ValuesCap.Text = "Enter 로 줄 나눔 (직접 입력) · 하나라도 맞으면 OK" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "LIKE 예 : F3* = F3 로 시작"
        '
        'Panel_ValuesBorder
        '
        Me.Panel_ValuesBorder.BackColor = System.Drawing.Color.FromArgb(CType(CType(55, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(221, Byte), Integer))
        Me.Panel_ValuesBorder.Controls.Add(Me.TextBox_Values)
        Me.Panel_ValuesBorder.Location = New System.Drawing.Point(560, 66)
        Me.Panel_ValuesBorder.Name = "Panel_ValuesBorder"
        Me.Panel_ValuesBorder.Padding = New System.Windows.Forms.Padding(2)
        Me.Panel_ValuesBorder.Size = New System.Drawing.Size(312, 306)
        '
        'TextBox_Values
        '
        Me.TextBox_Values.AcceptsReturn = True
        Me.TextBox_Values.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox_Values.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TextBox_Values.Font = New System.Drawing.Font("Consolas", 12.0!)
        Me.TextBox_Values.ForeColor = System.Drawing.Color.FromArgb(CType(CType(4, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.TextBox_Values.Multiline = True
        Me.TextBox_Values.Name = "TextBox_Values"
        Me.TextBox_Values.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox_Values.TabIndex = 11
        '
        'Panel_Test
        '
        Me.Panel_Test.BackColor = System.Drawing.Color.FromArgb(CType(CType(241, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(232, Byte), Integer))
        Me.Panel_Test.Controls.Add(Me.Label_TestResult)
        Me.Panel_Test.Controls.Add(Me.Button_Test)
        Me.Panel_Test.Controls.Add(Me.TextBox_TestMSCODE)
        Me.Panel_Test.Controls.Add(Me.Label_TestCap)
        Me.Panel_Test.Location = New System.Drawing.Point(14, 492)
        Me.Panel_Test.Name = "Panel_Test"
        Me.Panel_Test.Size = New System.Drawing.Size(872, 50)
        '
        'Label_TestCap
        '
        Me.Label_TestCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_TestCap.Location = New System.Drawing.Point(10, 14)
        Me.Label_TestCap.Name = "Label_TestCap"
        Me.Label_TestCap.Size = New System.Drawing.Size(70, 22)
        Me.Label_TestCap.Text = "판단 시험"
        Me.Label_TestCap.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'TextBox_TestMSCODE
        '
        Me.TextBox_TestMSCODE.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox_TestMSCODE.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.TextBox_TestMSCODE.Location = New System.Drawing.Point(84, 12)
        Me.TextBox_TestMSCODE.Name = "TextBox_TestMSCODE"
        Me.TextBox_TestMSCODE.Size = New System.Drawing.Size(190, 25)
        Me.TextBox_TestMSCODE.TabIndex = 12
        '
        'Button_Test
        '
        Me.Button_Test.Location = New System.Drawing.Point(282, 9)
        Me.Button_Test.Name = "Button_Test"
        Me.Button_Test.Size = New System.Drawing.Size(70, 31)
        Me.Button_Test.TabIndex = 13
        Me.Button_Test.Text = "시험"
        '
        'Label_TestResult
        '
        Me.Label_TestResult.Location = New System.Drawing.Point(362, 12)
        Me.Label_TestResult.Name = "Label_TestResult"
        Me.Label_TestResult.Size = New System.Drawing.Size(498, 26)
        Me.Label_TestResult.Text = "항목ID=값; 항목ID=값  (= 없으면 고른 조건 줄 항목)"
        Me.Label_TestResult.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_TestResult.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Button_Save
        '
        Me.Button_Save.Font = New System.Drawing.Font("맑은 고딕", 10.0!, System.Drawing.FontStyle.Bold)
        Me.Button_Save.Location = New System.Drawing.Point(658, 556)
        Me.Button_Save.Name = "Button_Save"
        Me.Button_Save.Size = New System.Drawing.Size(110, 40)
        Me.Button_Save.TabIndex = 14
        Me.Button_Save.Text = "저장"
        '
        'Button_Close
        '
        Me.Button_Close.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button_Close.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Close.Location = New System.Drawing.Point(776, 556)
        Me.Button_Close.Name = "Button_Close"
        Me.Button_Close.Size = New System.Drawing.Size(110, 40)
        Me.Button_Close.TabIndex = 15
        Me.Button_Close.Text = "닫기"
        '
        'frmRuleSetting
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.Button_Close
        Me.ClientSize = New System.Drawing.Size(900, 610)
        Me.Controls.Add(Me.Button_Close)
        Me.Controls.Add(Me.Button_Save)
        Me.Controls.Add(Me.Panel_Test)
        Me.Controls.Add(Me.Panel_Rule)
        Me.Controls.Add(Me.Button_Delete)
        Me.Controls.Add(Me.Button_Down)
        Me.Controls.Add(Me.Button_Up)
        Me.Controls.Add(Me.Button_New)
        Me.Controls.Add(Me.ComboBox_Rule)
        Me.Controls.Add(Me.Label_RuleCap)
        Me.Controls.Add(Me.Panel_Header)
        Me.Font = New System.Drawing.Font("맑은 고딕", 9.5!)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmRuleSetting"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "분기 조건 설정"
        Me.Panel_Header.ResumeLayout(False)
        Me.Panel_Header.PerformLayout()
        Me.Panel_Rule.ResumeLayout(False)
        Me.Panel_Rule.PerformLayout()
        CType(Me.DataGridView_Cond, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel_ValuesBorder.ResumeLayout(False)
        Me.Panel_ValuesBorder.PerformLayout()
        Me.Panel_Test.ResumeLayout(False)
        Me.Panel_Test.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents Panel_Header As System.Windows.Forms.Panel
    Friend WithEvents Label_Header As System.Windows.Forms.Label
    Friend WithEvents Label_RuleCap As System.Windows.Forms.Label
    Friend WithEvents ComboBox_Rule As System.Windows.Forms.ComboBox
    Friend WithEvents Button_New As System.Windows.Forms.Button
    Friend WithEvents Button_Up As System.Windows.Forms.Button
    Friend WithEvents Button_Down As System.Windows.Forms.Button
    Friend WithEvents Button_Delete As System.Windows.Forms.Button
    Friend WithEvents Panel_Rule As System.Windows.Forms.Panel
    Friend WithEvents CheckBox_Enabled As System.Windows.Forms.CheckBox
    Friend WithEvents Label_Priority As System.Windows.Forms.Label
    Friend WithEvents Label_DestCap As System.Windows.Forms.Label
    Friend WithEvents ComboBox_Dest As System.Windows.Forms.ComboBox
    Friend WithEvents Button_DestNew As System.Windows.Forms.Button
    Friend WithEvents Button_DestEdit As System.Windows.Forms.Button
    Friend WithEvents Button_DestDelete As System.Windows.Forms.Button
    Friend WithEvents colJoin As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents Label_DescCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_Desc As System.Windows.Forms.TextBox
    Friend WithEvents Label_CondCap As System.Windows.Forms.Label
    Friend WithEvents DataGridView_Cond As System.Windows.Forms.DataGridView
    Friend WithEvents colTarget As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colOperator As System.Windows.Forms.DataGridViewComboBoxColumn
    Friend WithEvents colValues As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents Button_CondAdd As System.Windows.Forms.Button
    Friend WithEvents Button_CondDelete As System.Windows.Forms.Button
    Friend WithEvents Label_ValuesTitle As System.Windows.Forms.Label
    Friend WithEvents Label_ValuesCap As System.Windows.Forms.Label
    Friend WithEvents Panel_ValuesBorder As System.Windows.Forms.Panel
    Friend WithEvents TextBox_Values As System.Windows.Forms.TextBox
    Friend WithEvents Panel_Test As System.Windows.Forms.Panel
    Friend WithEvents Label_TestCap As System.Windows.Forms.Label
    Friend WithEvents TextBox_TestMSCODE As System.Windows.Forms.TextBox
    Friend WithEvents Button_Test As System.Windows.Forms.Button
    Friend WithEvents Label_TestResult As System.Windows.Forms.Label
    Friend WithEvents Button_Save As System.Windows.Forms.Button
    Friend WithEvents Button_Close As System.Windows.Forms.Button
End Class
