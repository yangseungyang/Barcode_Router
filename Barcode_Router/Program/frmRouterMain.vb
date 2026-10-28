Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmRouterMain
'Function   : 바코드 분배 프로그램 (Barcode_Router) 메인 화면
'Framwork   : VB2022 Framework4.8
'Reference  : EtherComN (PLASMA 조회, 기존 검사 프로그램과 같음)
'Memo       : 흐름 (구조설계 8_전체Flow R1 ~ R7)
'             R1  바코드 수신 (HID : 바코드 칸 Enter / COM : 바코드리더 COM)
'             R2  8자리가 아니면 PLASMA 조회 (MSCODE, SERIAL, ORDER)
'             R2a 8자리(작업자 바코드)는 사용 중인 모든 분기 대상으로 전송 (10_확인항목 Q1 = A, 2026-10-26)
'                 작업자 바코드는 필요할 때만 찍는다 (검사 중이 아니면 언제든 받음)
'             R3  MODEL = MSCODE 첫 '-' 앞
'             R4  규칙 판단 (clsRouteJudge)
'             R5  분기 COM 으로 "바코드 + CR" 전송, 대상 프로그램 창을 앞으로
'             R6  검사 중 표시, END 대기 (대상별 타임아웃, 0 = 무제한)
'             R7  "END + CR" 수신 → 대기 → Router 창을 앞으로
'             화면 : Router_UI_시안_v1.html (2026-10-26 승인) - 제목줄, 큰 상태 패널(대기 초록 / 검사 중 주황 /
'                    오류 빨강 / 조회 중 파랑), 바코드 칸, 제품 정보 카드 4개, 분기 대상 카드, 로그, 평면 버튼
'                    제목줄 가운데 설비명 (COM 설정에서 입력, 2026-10-26 추가)
'             설정 / 로그 : 저장소 프로젝트 폴더의 Files\RouterSetting.mdb, System_Log\ (clsRouterPath, 공통 규약 7항)
'                    exe 를 바탕화면 등으로 복사해도 C:\Barcode_Router\Barcode_Router\ 를 읽는다. EtherComN.dll 은 exe 에 내장
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7항 / A8_FAM3_Router_구조설계_v1.1.xlsx
'==========================================================

Public Class frmRouterMain

    'Program Ver ==========================================
    Private Const Version As String = "0.00"    '2026.10.26 Ver 0.00 YSH

    Private Const OPERATOR_LEN As Integer = 8       '작업자 바코드 길이 (A8 / FAM3 frmMain 과 같은 판단)
    Private Const MSG_END As String = "END"         '검사 프로그램이 보내는 종료 신호 (구조설계 P-04)
    Private Const HINT_BUSY As String = "검사가 끝나면 이 화면이 다시 앞으로 옵니다"

    Dim mAppPath As String = clsRouterPath.DataFolder   '저장소 프로젝트 폴더 (bin 의 한 단계 위, bin 밖에서 실행하면 C:\Barcode_Router\Barcode_Router\)
    Dim mSetting As clsRouterSetting
    Dim mReaderPort As clsRouterPort                                        '바코드리더 COM (COM 모드일 때만)
    Dim mDestPorts As New Dictionary(Of Integer, clsRouterPort)             'DestID → 분기 대상 COM
    Dim mPending As New Dictionary(Of Integer, DateTime)                    'END 를 기다리는 DestID → 전송 시각
    Dim mLastEnd As New Dictionary(Of Integer, DateTime)                    'DestID → 마지막 END 수신 시각
    Dim mPendingText As String = ""
    Dim mOperator As String = "-"
    Dim mTick As Integer = 0

    Private Delegate Sub LineDelegate(ByVal port As clsRouterPort, ByVal line As String)

    'Load / Close ==========================================================================================================
    Private Sub frmRouterMain_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Me.Text = "Barcode Router  Ver : " & Version
        Label_Ver.Text = "Ver " & Version
        Label_Ver.Left = Label_Title.Right + 6
        clsUiTheme.StyleButton(Button_Rule, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Com, clsUiTheme.ButtonKind.Accent)
        Call UpdateReleaseButton()
        WriteLog("Program Start Ver " & Version)
        WriteLog("Data Folder : " & mAppPath & "  (exe : " & Application.ExecutablePath & ")")     '어느 폴더를 쓰는지 로그로 남김

        Call LoadAndOpen()
        Call SetReady()
        Call UpdateHeader()
        Timer_Wait.Start()
    End Sub

    Private Sub frmRouterMain_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Shown
        TextBox_BarCode.Focus()
    End Sub

    Private Sub frmRouterMain_FormClosed(ByVal sender As Object, ByVal e As FormClosedEventArgs) Handles MyBase.FormClosed
        Timer_Wait.Stop()
        Call CloseAllPorts()
        WriteLog("Program End")
    End Sub

    Private Sub LoadAndOpen()
        '==========================================================================================
        'Function : 설정 읽기 → 바코드리더 COM / 분기 대상 COM 열기
        '==========================================================================================
        Call CloseAllPorts()

        Try
            mSetting = New clsRouterSetting(mAppPath & "Files\RouterSetting.mdb")
            mSetting.Load()
            WriteLog("Setting Load OK : 분기 대상 " & mSetting.Destinations.Count & " / 규칙 " & mSetting.Rules.Count)
        Catch ex As Exception
            WriteLog("Setting Load NG : " & ex.Message)
            MessageBox.Show("설정을 읽지 못했습니다." & vbCrLf & ex.Message, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Error)
            mSetting = New clsRouterSetting(mAppPath & "Files\RouterSetting.mdb")
        End Try

        Call ShowEquipment()

        '바코드리더 COM
        Label_InputMode.Text = "HID 입력"
        If mSetting.Reader.InputMode = clsReaderSetting.MODE_COM Then
            Label_InputMode.Text = mSetting.Reader.Port & " 입력"
            Try
                mReaderPort = New clsRouterPort
                AddHandler mReaderPort.LineReceived, AddressOf Port_LineReceived
                mReaderPort.Open(mSetting.Reader.Port, mSetting.Reader.Setting)
                WriteLog("Reader COM Open OK : " & mSetting.Reader.Port)
            Catch ex As Exception
                WriteLog("Reader COM Open NG : " & mSetting.Reader.Port & " " & ex.Message)
                MessageBox.Show("바코드리더 COM 을 열지 못했습니다 : " & mSetting.Reader.Port & vbCrLf & ex.Message, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End Try
        End If

        '분기 대상 COM
        For Each d As clsDestination In mSetting.Destinations
            If Not d.Enabled Then Continue For
            Dim p As New clsRouterPort
            p.Tag = d
            AddHandler p.LineReceived, AddressOf Port_LineReceived
            Try
                p.Open(d.Port, d.Setting)
                WriteLog("Dest COM Open OK : " & d.DisplayName)
            Catch ex As Exception
                WriteLog("Dest COM Open NG : " & d.DisplayName & " " & ex.Message)
            End Try
            mDestPorts(d.DestID) = p
        Next

        Call RefreshDestList()
    End Sub

    Private Sub CloseAllPorts()
        If mReaderPort IsNot Nothing Then
            RemoveHandler mReaderPort.LineReceived, AddressOf Port_LineReceived
            mReaderPort.Close()
            mReaderPort = Nothing
        End If
        For Each p As clsRouterPort In mDestPorts.Values
            RemoveHandler p.LineReceived, AddressOf Port_LineReceived
            p.Close()
        Next
        mDestPorts.Clear()
    End Sub

    'Barcode input =========================================================================================================
    Private Sub TextBox_BarCode_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs) Handles TextBox_BarCode.KeyDown
        If e.KeyCode <> Keys.Return Then Exit Sub
        e.SuppressKeyPress = True
        Dim code As String = TextBox_BarCode.Text.Trim
        If code <> "" Then Call HandleBarcode(code)
    End Sub

    Private Sub Port_LineReceived(ByVal port As clsRouterPort, ByVal line As String)
        '수신 스레드 → 화면 스레드
        Try
            Me.BeginInvoke(New LineDelegate(AddressOf OnPortLine), port, line)
        Catch ex As Exception
            Debug.Print("BeginInvoke NG : " & ex.Message)
        End Try
    End Sub

    Private Sub OnPortLine(ByVal port As clsRouterPort, ByVal line As String)
        If port Is mReaderPort Then
            TextBox_BarCode.Text = line
            Call HandleBarcode(line)
            Exit Sub
        End If

        Dim d As clsDestination = TryCast(port.Tag, clsDestination)
        If d Is Nothing Then Exit Sub

        If line.ToUpperInvariant = MSG_END Then
            Call OnEnd(d)
        Else
            WriteLog("Recv " & d.DisplayName & " : " & line)
        End If
    End Sub

    Private Sub HandleBarcode(ByVal code As String)
        '==========================================================================================
        'Function : 바코드 한 건 처리 (R1 ~ R6)
        '==========================================================================================
        If mPending.Count > 0 Then
            WriteLog("Barcode skip (검사 중) : " & code)
            Call SetStatus(clsUiTheme.StatusKind.Busy, mPendingText, "검사 중 · 끝난 뒤 다시 찍어 주세요", "지금 찍은 바코드는 보내지 않았습니다 : " & code)
            TextBox_BarCode.SelectAll()
            Exit Sub
        End If

        WriteLog("Barcode : " & code)

        '작업자 바코드 (R2a) ----------------------------------------------------
        If code.Length = OPERATOR_LEN Then
            Call SendOperator(code)
            Exit Sub
        End If

        'PLASMA (R2) --------------------------------------------------------------
        Call ClearInfo()
        Dim cPsm As New cls100_PLASMA_Order
        Try
            Call SetStatus(clsUiTheme.StatusKind.Info, "PLASMA 조회 중", "조회 중", code)
            Application.DoEvents()
            cPsm.BARCODE_TEXT = code
            Call cPsm.GET_ORDER_INFORMAION()
        Catch ex As Exception
            WriteLog("PLASMA NG : " & code & " " & ex.Message)
            Call SetStatus(clsUiTheme.StatusKind.Error, "PLASMA 조회 실패", "오류", ex.Message)
            TextBox_BarCode.SelectAll()
            Exit Sub
        End Try

        Dim mscode As String = cPsm.MSCODE.Trim
        Label_MSCODE.Text = mscode
        Label_MODEL.Text = clsRouteJudge.GetModel(mscode)
        Label_SERIAL.Text = cPsm.SERIAL
        WriteLog("PLASMA OK : MSCODE=" & mscode & " SERIAL=" & cPsm.SERIAL & " ORDER=" & cPsm.ORDER)

        'Judge (R3, R4) -----------------------------------------------------------
        Dim ru As clsRouteRule = clsRouteJudge.Judge(mSetting.Rules, mscode)
        If ru Is Nothing Then
            WriteLog("Judge NG : 분기 대상 없음 MSCODE=" & mscode)
            Call SetStatus(clsUiTheme.StatusKind.Error, "분기 대상 없음 : " & mscode, "오류", "분기 조건 설정을 확인해 주세요")
            TextBox_BarCode.SelectAll()
            Exit Sub
        End If

        Dim d As clsDestination = mSetting.FindDestination(ru.DestID)
        If d Is Nothing OrElse Not d.Enabled OrElse Not mDestPorts.ContainsKey(d.DestID) Then
            WriteLog("Judge NG : 규칙 '" & ru.Description & "' 의 분기 대상이 없거나 사용 안 함")
            Call SetStatus(clsUiTheme.StatusKind.Error, "분기 대상 설정 오류", "오류", "규칙 '" & ru.Description & "' 의 분기 COM 을 확인해 주세요")
            TextBox_BarCode.SelectAll()
            Exit Sub
        End If
        Label_Rule.Text = ru.Priority & ". " & d.Port & "  " & ru.Description
        WriteLog("Judge OK : " & ru.ToString & " → " & d.DisplayName)

        'Send (R5, R6) ------------------------------------------------------------
        If Not SendTo(d, code) Then
            Call SetStatus(clsUiTheme.StatusKind.Error, "전송 실패 : " & d.DestName, "오류", d.Port & " 를 확인해 주세요")
            TextBox_BarCode.SelectAll()
            Exit Sub
        End If

        mPendingText = d.DestName
        Call SetStatus(clsUiTheme.StatusKind.Busy, d.DestName, "검사 중 · " & Format(DateTime.Now, "HH:mm:ss") & " 전송", HINT_BUSY)
        Call RefreshDestList()

        If Not clsWindowControl.BringProcessToFront(d.ProcessName) Then
            WriteLog("Window NG : 프로세스 창을 찾지 못함 " & d.ProcessName)
        End If
    End Sub

    Private Sub SendOperator(ByVal code As String)
        '==========================================================================================
        'Function : 작업자 바코드를 사용 중인 모든 분기 대상으로 전송 (10_확인항목 Q1 = A)
        '           A8 : 작업자 갱신 / FAM3 : 작업자 갱신 + 주명판 송신. 각 대상의 END 를 기다린다
        '==========================================================================================
        Call ClearInfo()
        Dim names As New List(Of String)

        For Each d As clsDestination In mSetting.Destinations
            If Not d.Enabled OrElse Not mDestPorts.ContainsKey(d.DestID) Then Continue For
            If SendTo(d, code) Then names.Add(d.DestName)
        Next

        If names.Count = 0 Then
            Call SetStatus(clsUiTheme.StatusKind.Error, "작업자 바코드를 보낼 곳이 없습니다", "오류", "COM 설정을 확인해 주세요")
            TextBox_BarCode.SelectAll()
            Exit Sub
        End If

        mOperator = code
        Call UpdateHeader()
        mPendingText = "작업자 갱신"
        Call SetStatus(clsUiTheme.StatusKind.Busy, "작업자 갱신 : " & code, "검사 중 · " & String.Join(", ", names), HINT_BUSY)
        Call RefreshDestList()
    End Sub

    Private Function SendTo(ByVal d As clsDestination, ByVal code As String) As Boolean
        Dim p As clsRouterPort = mDestPorts(d.DestID)
        Try
            If Not p.IsOpen Then p.Open(d.Port, d.Setting) '닫혀 있으면 한 번 더 열어 본다
            p.SendLine(code)
            mPending(d.DestID) = DateTime.Now
            WriteLog("Send OK : " & d.DisplayName & " ← " & code)
            Return True
        Catch ex As Exception
            WriteLog("Send NG : " & d.DisplayName & " " & ex.Message)
            Return False
        End Try
    End Function

    'END ===================================================================================================================
    Private Sub OnEnd(ByVal d As clsDestination)
        '==========================================================================================
        'Function : END 수신 (R7)
        '==========================================================================================
        WriteLog("END : " & d.DisplayName)
        If Not mPending.ContainsKey(d.DestID) Then Exit Sub

        mPending.Remove(d.DestID)
        mLastEnd(d.DestID) = DateTime.Now
        Call RefreshDestList()

        If mPending.Count = 0 Then
            Call SetReady()
            Call ShowRouterFront()
        End If
    End Sub

    Private Sub Timer_Wait_Tick(ByVal sender As Object, ByVal e As EventArgs) Handles Timer_Wait.Tick
        '==========================================================================================
        'Function : 1초마다 - 제목줄 시각, END 대기 타임아웃(대상별 EndTimeoutSec, 0 = 무제한),
        '           5초마다 분기 대상 카드(프로그램 실행 여부) 갱신
        '==========================================================================================
        mTick += 1
        Call UpdateHeader()
        If mTick Mod 5 = 0 Then Call RefreshDestList()
        If mPending.Count = 0 Then Exit Sub

        Dim changed As Boolean = False
        For Each id As Integer In mPending.Keys.ToList
            Dim d As clsDestination = mSetting.FindDestination(id)
            If d Is Nothing Then mPending.Remove(id) : changed = True : Continue For
            If d.EndTimeoutSec <= 0 Then Continue For
            If (DateTime.Now - mPending(id)).TotalSeconds > d.EndTimeoutSec Then
                mPending.Remove(id)
                changed = True
                WriteLog("END Timeout : " & d.DisplayName & " (" & d.EndTimeoutSec & "s)")
                MessageBox.Show("END 를 받지 못했습니다 (타임아웃 " & d.EndTimeoutSec & "초) : " & d.DisplayName, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Next

        If Not changed Then Exit Sub
        Call RefreshDestList()
        If mPending.Count = 0 Then
            Call SetReady()
            Call ShowRouterFront()
        End If
    End Sub

    Private Sub Button_Release_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Release.Click
        If mPending.Count = 0 Then
            TextBox_BarCode.Focus()
            Exit Sub
        End If
        If MessageBox.Show("END 를 기다리지 않고 대기 상태로 돌아갑니까?", "Barcode Router", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub
        WriteLog("Wait Release (수동) : " & mPendingText)
        mPending.Clear()
        Call RefreshDestList()
        Call SetReady()
    End Sub

    'Setting ===============================================================================================================
    Private Sub Button_Rule_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Rule.Click
        If Not CanOpenSetting() Then Exit Sub
        Call CloseAllPorts()            '분기 COM 신규 / 수정 화면에서 포트 시험을 할 수 있도록 Router 의 포트를 모두 닫는다
        Using f As New frmRuleSetting(mSetting)
            If f.ShowDialog(Me) = DialogResult.OK Then WriteLog("Rule Setting Saved")
        End Using
        Call LoadAndOpen()
        TextBox_BarCode.Focus()
    End Sub

    Private Sub Button_Com_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Com.Click
        If Not CanOpenSetting() Then Exit Sub
        Call CloseAllPorts()            '포트 시험을 할 수 있도록 Router 의 포트를 모두 닫는다
        Using f As New frmComSetting(mSetting)
            If f.ShowDialog(Me) = DialogResult.OK Then WriteLog("COM Setting Saved")
        End Using
        Call LoadAndOpen()
        TextBox_BarCode.Focus()
    End Sub

    Private Function CanOpenSetting() As Boolean
        If mPending.Count = 0 Then Return True
        MessageBox.Show("검사 중에는 설정을 바꿀 수 없습니다.", "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return False
    End Function

    'Display ===============================================================================================================
    Private Sub SetReady()
        mPendingText = ""
        Call UpdateReleaseButton()
        Call SetStatus(clsUiTheme.StatusKind.Ready, "바코드를 찍어 주세요", "대기", "")
        TextBox_BarCode.Text = ""
        TextBox_BarCode.Focus()
    End Sub

    Private Sub SetStatus(ByVal kind As clsUiTheme.StatusKind, ByVal main As String, ByVal subText As String, ByVal hint As String)
        '==========================================================================================
        'Function : 상태 패널 (대기 초록 / 검사 중 주황 / 오류 빨강 / 조회 중 파랑)
        '==========================================================================================
        Dim c As clsUiTheme.StatusColor = clsUiTheme.GetStatusColor(kind)
        Panel_StatusBorder.BackColor = c.Border
        Panel_Status.BackColor = c.Back
        Label_Status.ForeColor = c.Main
        Label_StatusSub.ForeColor = c.Sub
        Label_StatusHint.ForeColor = c.Sub
        Label_Status.Text = main
        Label_StatusSub.Text = subText
        Label_StatusHint.Text = hint
        If kind = clsUiTheme.StatusKind.Busy Then TextBox_BarCode.Text = ""
        Panel_Status.Refresh()
    End Sub

    Private Sub UpdateReleaseButton()
        '==========================================================================================
        'Function : 대기 해제 버튼은 END 를 기다리는 동안에만 빨갛게 켜고, 평소에는 회색으로 끈다
        '           (2026-10-26 지시 : 처음 실행하면 빨간 대기 해제 버튼이 떠 있어 오류처럼 보임)
        '==========================================================================================
        Dim waiting As Boolean = (mPending.Count > 0)
        If waiting Then
            clsUiTheme.StyleButton(Button_Release, clsUiTheme.ButtonKind.Danger)
            Button_Release.Cursor = Cursors.Hand
        Else
            clsUiTheme.StyleButton(Button_Release, clsUiTheme.ButtonKind.Normal)
            Button_Release.BackColor = clsUiTheme.CardBack
            Button_Release.FlatAppearance.BorderColor = clsUiTheme.LineLight
            Button_Release.Cursor = Cursors.Default
        End If
        Button_Release.Enabled = waiting
    End Sub

    Private Sub ShowEquipment()
        '==========================================================================================
        'Function : 설비명을 제목줄 가운데에 표시 (COM 설정 > 설비명, 2026-10-26 추가)
        '==========================================================================================
        Dim name As String = mSetting.EquipmentName.Trim
        If name = "" Then
            Label_Equipment.Text = "설비명 미설정 (COM 설정에서 입력)"
            Label_Equipment.Font = New Font("맑은 고딕", 9.5!)
            Label_Equipment.ForeColor = clsUiTheme.HeaderSub
            Label_Equipment.BackColor = clsUiTheme.HeaderBack
            Me.Text = "Barcode Router  Ver : " & Version
        Else
            Label_Equipment.Text = name
            Label_Equipment.Font = New Font("맑은 고딕", 13.0!, FontStyle.Bold)
            Label_Equipment.ForeColor = Color.White
            Label_Equipment.BackColor = clsUiTheme.PrimaryBack
            Me.Text = "Barcode Router  Ver : " & Version & "  [" & name & "]"     '작업표시줄에서도 설비를 알 수 있게
        End If
    End Sub

    Private Sub UpdateHeader()
        Label_HeaderInfo.Text = "작업자 " & mOperator & "  ·  " & Format(DateTime.Now, "yyyy-MM-dd HH:mm")
    End Sub

    Private Sub ClearInfo()
        Label_MSCODE.Text = ""
        Label_MODEL.Text = ""
        Label_SERIAL.Text = ""
        Label_Rule.Text = ""
    End Sub

    Private Sub ShowRouterFront()
        If Me.WindowState = FormWindowState.Minimized Then Me.WindowState = FormWindowState.Normal
        Me.Show()
        Me.TopMost = True
        Me.TopMost = False
        Me.Activate()
        clsWindowControl.ShowFront(Me.Handle)
        TextBox_BarCode.Focus()
    End Sub

    Private Sub RefreshDestList()
        '==========================================================================================
        'Function : 분기 대상 카드 (왼쪽 색 띠 + 상태 표시 + COM / 프로그램 / 마지막 END)
        '==========================================================================================
        Call UpdateReleaseButton()      '기다리는 검사가 생기거나 없어질 때마다 여기를 지난다
        If mSetting Is Nothing Then Exit Sub
        FlowLayoutPanel_Dest.SuspendLayout()
        For Each c As Control In FlowLayoutPanel_Dest.Controls.Cast(Of Control).ToList
            FlowLayoutPanel_Dest.Controls.Remove(c)
            c.Dispose()
        Next

        For Each d As clsDestination In mSetting.Destinations
            Dim kind As clsUiTheme.StatusKind
            Dim state As String
            If Not d.Enabled Then
                kind = clsUiTheme.StatusKind.Info : state = "사용 안 함"
            ElseIf Not mDestPorts.ContainsKey(d.DestID) OrElse Not mDestPorts(d.DestID).IsOpen Then
                kind = clsUiTheme.StatusKind.Error : state = "COM 실패"
            ElseIf mPending.ContainsKey(d.DestID) Then
                kind = clsUiTheme.StatusKind.Busy : state = "검사 중"
            Else
                kind = clsUiTheme.StatusKind.Ready : state = "대기"
            End If

            Dim running As Boolean = clsWindowControl.IsProcessRunning(d.ProcessName)
            If d.Enabled AndAlso Not running AndAlso kind = clsUiTheme.StatusKind.Ready Then
                kind = clsUiTheme.StatusKind.Error : state = "프로그램 꺼짐"
            End If

            Dim detail As String = d.Port & "  ·  " & If(d.ProcessName = "", "-", d.ProcessName) & If(running, " 실행 중", " 꺼짐")
            If mPending.ContainsKey(d.DestID) Then detail &= "  ·  " & Format(mPending(d.DestID), "HH:mm:ss") & " 전송"
            If mLastEnd.ContainsKey(d.DestID) Then detail &= "  ·  마지막 END " & Format(mLastEnd(d.DestID), "HH:mm:ss")

            FlowLayoutPanel_Dest.Controls.Add(MakeDestCard(d.DestName, state, detail, kind))
        Next
        FlowLayoutPanel_Dest.ResumeLayout()
    End Sub

    Private Function MakeDestCard(ByVal name As String, ByVal state As String, ByVal detail As String, ByVal kind As clsUiTheme.StatusKind) As Panel
        Dim c As clsUiTheme.StatusColor = clsUiTheme.GetStatusColor(kind)

        Dim card As New Panel With {.Size = New Size(428, 60), .Margin = New Padding(0, 0, 12, 10), .BackColor = Color.White}
        AddHandler card.Paint, Sub(s As Object, pe As PaintEventArgs)
                                   Dim p As Panel = DirectCast(s, Panel)
                                   Using pen As New Pen(clsUiTheme.Line)
                                       pe.Graphics.DrawRectangle(pen, 0, 0, p.Width - 1, p.Height - 1)
                                   End Using
                               End Sub

        Dim bar As New Panel With {.Dock = DockStyle.Left, .Width = 5, .BackColor = c.Border}
        Dim lblName As New Label With {.Text = name, .Font = New Font("맑은 고딕", 11.0!, FontStyle.Bold), .ForeColor = clsUiTheme.TextMain,
                                       .Location = New Point(14, 7), .Size = New Size(290, 24), .AutoEllipsis = True}
        Dim lblState As New Label With {.Text = state, .Font = New Font("맑은 고딕", 8.5!), .ForeColor = c.Main, .BackColor = c.Back,
                                        .TextAlign = ContentAlignment.MiddleCenter, .Location = New Point(318, 9), .Size = New Size(100, 20)}
        Dim lblDetail As New Label With {.Text = detail, .Font = New Font("맑은 고딕", 8.5!), .ForeColor = clsUiTheme.TextSub,
                                         .Location = New Point(14, 34), .Size = New Size(404, 20), .AutoEllipsis = True}
        card.Controls.Add(lblDetail)
        card.Controls.Add(lblState)
        card.Controls.Add(lblName)
        card.Controls.Add(bar)
        Return card
    End Function

    Private Sub WriteLog(ByVal msg As String)
        clsRouterLog.Write(mAppPath, msg)
        ListBox_Log.Items.Insert(0, Format(DateTime.Now, "HH:mm:ss") & "  " & msg)
        Do While ListBox_Log.Items.Count > 100
            ListBox_Log.Items.RemoveAt(ListBox_Log.Items.Count - 1)
        Loop
    End Sub

End Class
