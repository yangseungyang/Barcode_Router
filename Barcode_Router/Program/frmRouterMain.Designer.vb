<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmRouterMain
    Inherits System.Windows.Forms.Form

    'Form은 Dispose를 재정의하여 구성 요소 목록을 정리합니다.
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

    'Windows Form 디자이너에 필요합니다.
    Private components As System.ComponentModel.IContainer

    '참고: 다음 프로시저는 Windows Form 디자이너에 필요합니다.
    '수정하려면 Windows Form 디자이너를 사용하십시오.
    '코드 편집기에서는 수정하지 마세요.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Panel_Header = New System.Windows.Forms.Panel()
        Me.Label_HeaderInfo = New System.Windows.Forms.Label()
        Me.Label_Equipment = New System.Windows.Forms.Label()
        Me.Label_Ver = New System.Windows.Forms.Label()
        Me.Label_Title = New System.Windows.Forms.Label()
        Me.Panel_StatusBorder = New System.Windows.Forms.Panel()
        Me.Panel_Status = New System.Windows.Forms.Panel()
        Me.Label_Status = New System.Windows.Forms.Label()
        Me.Label_StatusHint = New System.Windows.Forms.Label()
        Me.Label_StatusSub = New System.Windows.Forms.Label()
        Me.Panel_BarCodeBorder = New System.Windows.Forms.Panel()
        Me.Panel_BarCode = New System.Windows.Forms.Panel()
        Me.Label_InputMode = New System.Windows.Forms.Label()
        Me.TextBox_BarCode = New System.Windows.Forms.TextBox()
        Me.Label_ScanIcon = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel_Cards = New System.Windows.Forms.FlowLayoutPanel()
        Me.Button_Manual = New System.Windows.Forms.Button()
        Me.Button_Display = New System.Windows.Forms.Button()
        Me.Label_DestCap = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel_Dest = New System.Windows.Forms.FlowLayoutPanel()
        Me.Label_LogCap = New System.Windows.Forms.Label()
        Me.ListBox_Log = New System.Windows.Forms.ListBox()
        Me.Button_Rule = New System.Windows.Forms.Button()
        Me.Button_Com = New System.Windows.Forms.Button()
        Me.Button_Release = New System.Windows.Forms.Button()
        Me.Timer_Wait = New System.Windows.Forms.Timer(Me.components)
        Me.Panel_Header.SuspendLayout()
        Me.Panel_StatusBorder.SuspendLayout()
        Me.Panel_Status.SuspendLayout()
        Me.Panel_BarCodeBorder.SuspendLayout()
        Me.Panel_BarCode.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel_Header
        '
        Me.Panel_Header.BackColor = System.Drawing.Color.FromArgb(CType(CType(12, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(124, Byte), Integer))
        Me.Panel_Header.Controls.Add(Me.Button_Manual)
        Me.Panel_Header.Controls.Add(Me.Label_Equipment)
        Me.Panel_Header.Controls.Add(Me.Label_HeaderInfo)
        Me.Panel_Header.Controls.Add(Me.Label_Ver)
        Me.Panel_Header.Controls.Add(Me.Label_Title)
        Me.Panel_Header.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Header.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Header.Name = "Panel_Header"
        Me.Panel_Header.Size = New System.Drawing.Size(900, 48)
        '
        'Label_Title
        '
        Me.Label_Title.AutoSize = True
        Me.Label_Title.Font = New System.Drawing.Font("맑은 고딕", 14.0!, System.Drawing.FontStyle.Bold)
        Me.Label_Title.ForeColor = System.Drawing.Color.FromArgb(CType(CType(230, Byte), Integer), CType(CType(241, Byte), Integer), CType(CType(251, Byte), Integer))
        Me.Label_Title.Location = New System.Drawing.Point(12, 10)
        Me.Label_Title.Name = "Label_Title"
        Me.Label_Title.Text = "▥  Barcode Router"
        '
        'Label_Ver
        '
        Me.Label_Ver.AutoSize = True
        Me.Label_Ver.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Label_Ver.ForeColor = System.Drawing.Color.FromArgb(CType(CType(181, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(244, Byte), Integer))
        Me.Label_Ver.Location = New System.Drawing.Point(206, 18)
        Me.Label_Ver.Name = "Label_Ver"
        Me.Label_Ver.Text = "Ver 0.00"
        '
        'Label_HeaderInfo
        '
        Me.Label_HeaderInfo.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Label_HeaderInfo.ForeColor = System.Drawing.Color.FromArgb(CType(CType(181, Byte), Integer), CType(CType(212, Byte), Integer), CType(CType(244, Byte), Integer))
        Me.Label_HeaderInfo.Location = New System.Drawing.Point(556, 14)
        Me.Label_HeaderInfo.Name = "Label_HeaderInfo"
        Me.Label_HeaderInfo.Size = New System.Drawing.Size(240, 20)
        Me.Label_HeaderInfo.Text = "작업자 - · 2026-10-26 00:00"
        Me.Label_HeaderInfo.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'Label_Equipment (설비명 - 제목줄 가운데)
        '
        Me.Label_Equipment.AutoEllipsis = True
        Me.Label_Equipment.BackColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(95, Byte), Integer), CType(CType(165, Byte), Integer))
        Me.Label_Equipment.Font = New System.Drawing.Font("맑은 고딕", 13.0!, System.Drawing.FontStyle.Bold)
        Me.Label_Equipment.ForeColor = System.Drawing.Color.White
        Me.Label_Equipment.Location = New System.Drawing.Point(262, 7)
        Me.Label_Equipment.Name = "Label_Equipment"
        Me.Label_Equipment.Size = New System.Drawing.Size(290, 34)
        Me.Label_Equipment.Text = "설비명"
        Me.Label_Equipment.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Button_Manual (사용자 설명서 - 제목줄 오른쪽)
        '
        Me.Button_Manual.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Button_Manual.Location = New System.Drawing.Point(804, 9)
        Me.Button_Manual.Name = "Button_Manual"
        Me.Button_Manual.Size = New System.Drawing.Size(84, 30)
        Me.Button_Manual.TabStop = False
        Me.Button_Manual.Text = "ⓘ 설명서"
        '
        'Panel_StatusBorder
        '
        Me.Panel_StatusBorder.BackColor = System.Drawing.Color.FromArgb(CType(CType(151, Byte), Integer), CType(CType(196, Byte), Integer), CType(CType(89, Byte), Integer))
        Me.Panel_StatusBorder.Controls.Add(Me.Panel_Status)
        Me.Panel_StatusBorder.Location = New System.Drawing.Point(14, 60)
        Me.Panel_StatusBorder.Name = "Panel_StatusBorder"
        Me.Panel_StatusBorder.Padding = New System.Windows.Forms.Padding(2)
        Me.Panel_StatusBorder.Size = New System.Drawing.Size(872, 112)
        '
        'Panel_Status
        '
        Me.Panel_Status.BackColor = System.Drawing.Color.FromArgb(CType(CType(234, Byte), Integer), CType(CType(243, Byte), Integer), CType(CType(222, Byte), Integer))
        Me.Panel_Status.Controls.Add(Me.Label_Status)
        Me.Panel_Status.Controls.Add(Me.Label_StatusHint)
        Me.Panel_Status.Controls.Add(Me.Label_StatusSub)
        Me.Panel_Status.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_Status.Name = "Panel_Status"
        Me.Panel_Status.Padding = New System.Windows.Forms.Padding(8, 6, 8, 6)
        '
        'Label_Status
        '
        Me.Label_Status.AutoEllipsis = True
        Me.Label_Status.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Label_Status.Font = New System.Drawing.Font("맑은 고딕", 24.0!, System.Drawing.FontStyle.Bold)
        Me.Label_Status.Name = "Label_Status"
        Me.Label_Status.Text = "바코드를 찍어 주세요"
        Me.Label_Status.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_StatusHint
        '
        Me.Label_StatusHint.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label_StatusHint.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Label_StatusHint.Name = "Label_StatusHint"
        Me.Label_StatusHint.Size = New System.Drawing.Size(852, 20)
        Me.Label_StatusHint.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label_StatusSub
        '
        Me.Label_StatusSub.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label_StatusSub.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Label_StatusSub.Name = "Label_StatusSub"
        Me.Label_StatusSub.Size = New System.Drawing.Size(852, 22)
        Me.Label_StatusSub.Text = "대기"
        Me.Label_StatusSub.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Panel_BarCodeBorder
        '
        Me.Panel_BarCodeBorder.BackColor = System.Drawing.Color.FromArgb(CType(CType(55, Byte), Integer), CType(CType(138, Byte), Integer), CType(CType(221, Byte), Integer))
        Me.Panel_BarCodeBorder.Controls.Add(Me.Panel_BarCode)
        Me.Panel_BarCodeBorder.Location = New System.Drawing.Point(14, 182)
        Me.Panel_BarCodeBorder.Name = "Panel_BarCodeBorder"
        Me.Panel_BarCodeBorder.Padding = New System.Windows.Forms.Padding(2)
        Me.Panel_BarCodeBorder.Size = New System.Drawing.Size(872, 56)
        '
        'Panel_BarCode
        '
        Me.Panel_BarCode.BackColor = System.Drawing.Color.White
        Me.Panel_BarCode.Controls.Add(Me.Label_InputMode)
        Me.Panel_BarCode.Controls.Add(Me.TextBox_BarCode)
        Me.Panel_BarCode.Controls.Add(Me.Label_ScanIcon)
        Me.Panel_BarCode.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel_BarCode.Name = "Panel_BarCode"
        '
        'Label_ScanIcon
        '
        Me.Label_ScanIcon.Font = New System.Drawing.Font("맑은 고딕", 16.0!)
        Me.Label_ScanIcon.ForeColor = System.Drawing.Color.FromArgb(CType(CType(24, Byte), Integer), CType(CType(95, Byte), Integer), CType(CType(165, Byte), Integer))
        Me.Label_ScanIcon.Location = New System.Drawing.Point(8, 8)
        Me.Label_ScanIcon.Name = "Label_ScanIcon"
        Me.Label_ScanIcon.Size = New System.Drawing.Size(34, 34)
        Me.Label_ScanIcon.Text = "▤"
        Me.Label_ScanIcon.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'TextBox_BarCode
        '
        Me.TextBox_BarCode.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox_BarCode.Font = New System.Drawing.Font("맑은 고딕", 20.0!, System.Drawing.FontStyle.Bold)
        Me.TextBox_BarCode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(4, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(83, Byte), Integer))
        Me.TextBox_BarCode.Location = New System.Drawing.Point(48, 7)
        Me.TextBox_BarCode.Name = "TextBox_BarCode"
        Me.TextBox_BarCode.Size = New System.Drawing.Size(700, 36)
        Me.TextBox_BarCode.TabIndex = 0
        '
        'Label_InputMode
        '
        Me.Label_InputMode.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.Label_InputMode.ForeColor = System.Drawing.Color.FromArgb(CType(CType(95, Byte), Integer), CType(CType(94, Byte), Integer), CType(CType(90, Byte), Integer))
        Me.Label_InputMode.Location = New System.Drawing.Point(756, 14)
        Me.Label_InputMode.Name = "Label_InputMode"
        Me.Label_InputMode.Size = New System.Drawing.Size(104, 22)
        Me.Label_InputMode.Text = "HID 입력"
        Me.Label_InputMode.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'FlowLayoutPanel_Cards (제품 정보 카드 - tblCard 설정대로 코드에서 만든다)
        '
        Me.FlowLayoutPanel_Cards.Location = New System.Drawing.Point(14, 248)
        Me.FlowLayoutPanel_Cards.Name = "FlowLayoutPanel_Cards"
        Me.FlowLayoutPanel_Cards.Size = New System.Drawing.Size(880, 58)
        Me.FlowLayoutPanel_Cards.Margin = New System.Windows.Forms.Padding(0)
        '
        'Label_DestCap
        '
        Me.Label_DestCap.AutoSize = True
        Me.Label_DestCap.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_DestCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_DestCap.Location = New System.Drawing.Point(12, 316)
        Me.Label_DestCap.Name = "Label_DestCap"
        Me.Label_DestCap.Text = "분기 대상"
        '
        'FlowLayoutPanel_Dest
        '
        Me.FlowLayoutPanel_Dest.AutoScroll = True
        Me.FlowLayoutPanel_Dest.Location = New System.Drawing.Point(14, 338)
        Me.FlowLayoutPanel_Dest.Name = "FlowLayoutPanel_Dest"
        Me.FlowLayoutPanel_Dest.Size = New System.Drawing.Size(886, 136)
        '
        'Label_LogCap
        '
        Me.Label_LogCap.AutoSize = True
        Me.Label_LogCap.Font = New System.Drawing.Font("맑은 고딕", 9.5!, System.Drawing.FontStyle.Bold)
        Me.Label_LogCap.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.Label_LogCap.Location = New System.Drawing.Point(12, 480)
        Me.Label_LogCap.Name = "Label_LogCap"
        Me.Label_LogCap.Text = "로그"
        '
        'ListBox_Log
        '
        Me.ListBox_Log.BackColor = System.Drawing.Color.FromArgb(CType(CType(241, Byte), Integer), CType(CType(239, Byte), Integer), CType(CType(232, Byte), Integer))
        Me.ListBox_Log.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ListBox_Log.Font = New System.Drawing.Font("Consolas", 9.5!)
        Me.ListBox_Log.ForeColor = System.Drawing.Color.FromArgb(CType(CType(68, Byte), Integer), CType(CType(68, Byte), Integer), CType(CType(65, Byte), Integer))
        Me.ListBox_Log.FormattingEnabled = True
        Me.ListBox_Log.HorizontalScrollbar = True
        Me.ListBox_Log.IntegralHeight = False
        Me.ListBox_Log.ItemHeight = 15
        Me.ListBox_Log.Location = New System.Drawing.Point(14, 502)
        Me.ListBox_Log.Name = "ListBox_Log"
        Me.ListBox_Log.Size = New System.Drawing.Size(872, 118)
        Me.ListBox_Log.TabStop = False
        '
        'Button_Rule
        '
        Me.Button_Rule.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Rule.Location = New System.Drawing.Point(14, 632)
        Me.Button_Rule.Name = "Button_Rule"
        Me.Button_Rule.Size = New System.Drawing.Size(170, 42)
        Me.Button_Rule.TabStop = False
        Me.Button_Rule.Text = "⑂  분기 조건 설정"
        '
        'Button_Com
        '
        Me.Button_Com.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Com.Location = New System.Drawing.Point(192, 632)
        Me.Button_Com.Name = "Button_Com"
        Me.Button_Com.Size = New System.Drawing.Size(150, 42)
        Me.Button_Com.TabStop = False
        Me.Button_Com.Text = "⇄  COM 설정"
        '
        'Button_Display (화면 설정 : 항목 / 카드 / 바코드 종류 / 기타, 설정 내보내기 / 가져오기)
        '
        Me.Button_Display.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Display.Location = New System.Drawing.Point(350, 632)
        Me.Button_Display.Name = "Button_Display"
        Me.Button_Display.Size = New System.Drawing.Size(150, 42)
        Me.Button_Display.TabStop = False
        Me.Button_Display.Text = "⚙  화면 설정"
        '
        'Button_Release
        '
        Me.Button_Release.Font = New System.Drawing.Font("맑은 고딕", 10.0!)
        Me.Button_Release.Location = New System.Drawing.Point(736, 632)
        Me.Button_Release.Name = "Button_Release"
        Me.Button_Release.Size = New System.Drawing.Size(150, 42)
        Me.Button_Release.TabStop = False
        Me.Button_Release.Text = "■  대기 해제"
        '
        'Timer_Wait
        '
        Me.Timer_Wait.Interval = 1000
        '
        'frmRouterMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.ClientSize = New System.Drawing.Size(900, 688)
        Me.Controls.Add(Me.Button_Release)
        Me.Controls.Add(Me.Button_Com)
        Me.Controls.Add(Me.Button_Rule)
        Me.Controls.Add(Me.ListBox_Log)
        Me.Controls.Add(Me.Label_LogCap)
        Me.Controls.Add(Me.FlowLayoutPanel_Dest)
        Me.Controls.Add(Me.Label_DestCap)
        Me.Controls.Add(Me.FlowLayoutPanel_Cards)
        Me.Controls.Add(Me.Button_Display)
        Me.Controls.Add(Me.Panel_BarCodeBorder)
        Me.Controls.Add(Me.Panel_StatusBorder)
        Me.Controls.Add(Me.Panel_Header)
        Me.Font = New System.Drawing.Font("맑은 고딕", 9.0!)
        Me.ForeColor = System.Drawing.Color.FromArgb(CType(CType(44, Byte), Integer), CType(CType(44, Byte), Integer), CType(CType(42, Byte), Integer))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "frmRouterMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Barcode Router"
        Me.Panel_Header.ResumeLayout(False)
        Me.Panel_Header.PerformLayout()
        Me.Panel_StatusBorder.ResumeLayout(False)
        Me.Panel_Status.ResumeLayout(False)
        Me.Panel_BarCodeBorder.ResumeLayout(False)
        Me.Panel_BarCode.ResumeLayout(False)
        Me.Panel_BarCode.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Panel_Header As System.Windows.Forms.Panel
    Friend WithEvents Label_Title As System.Windows.Forms.Label
    Friend WithEvents Label_Ver As System.Windows.Forms.Label
    Friend WithEvents Label_HeaderInfo As System.Windows.Forms.Label
    Friend WithEvents Label_Equipment As System.Windows.Forms.Label
    Friend WithEvents Panel_StatusBorder As System.Windows.Forms.Panel
    Friend WithEvents Panel_Status As System.Windows.Forms.Panel
    Friend WithEvents Label_Status As System.Windows.Forms.Label
    Friend WithEvents Label_StatusHint As System.Windows.Forms.Label
    Friend WithEvents Label_StatusSub As System.Windows.Forms.Label
    Friend WithEvents Panel_BarCodeBorder As System.Windows.Forms.Panel
    Friend WithEvents Panel_BarCode As System.Windows.Forms.Panel
    Friend WithEvents Label_ScanIcon As System.Windows.Forms.Label
    Friend WithEvents TextBox_BarCode As System.Windows.Forms.TextBox
    Friend WithEvents Label_InputMode As System.Windows.Forms.Label
    Friend WithEvents FlowLayoutPanel_Cards As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents Button_Manual As System.Windows.Forms.Button
    Friend WithEvents Button_Display As System.Windows.Forms.Button
    Friend WithEvents Label_DestCap As System.Windows.Forms.Label
    Friend WithEvents FlowLayoutPanel_Dest As System.Windows.Forms.FlowLayoutPanel
    Friend WithEvents Label_LogCap As System.Windows.Forms.Label
    Friend WithEvents ListBox_Log As System.Windows.Forms.ListBox
    Friend WithEvents Button_Rule As System.Windows.Forms.Button
    Friend WithEvents Button_Com As System.Windows.Forms.Button
    Friend WithEvents Button_Release As System.Windows.Forms.Button
    Friend WithEvents Timer_Wait As System.Windows.Forms.Timer
End Class
