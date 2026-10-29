Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmRouterMain
'Function   : 바코드 분배 프로그램 (Barcode_Router) 메인 화면
'Framwork   : VB2022 Framework4.8
'Reference  : EtherComN (PLASMA 조회, 기존 검사 프로그램과 같음)
'Memo       : 흐름 (구조설계 8_전체Flow R1 ~ R7)
'             R1  바코드 수신 (HID : 바코드 칸 Enter / COM : 바코드리더 COM)
'             R2  바코드 종류 판단 (tblBarcodeType : 길이 / 글자 종류 / 패턴, 2026-10-28) → 처리
'                 ROUTE     : PLASMA 조회 (서버 = tblConfig) → 항목 값 (tblField) → 카드 표시 (tblCard)
'                 BROADCAST : 사용 중인 모든 분기 대상으로 전송 (예 : 작업자 바코드, 10_확인항목 Q1 = A)
'                 REJECT / 맞는 종류 없음 : 보내지 않음
'             R3  항목 값 : PLASMA 항목 / 다른 항목에서 잘라 만든 항목(예 : MODEL) / 입력 바코드 (clsRouteData)
'             R4  규칙 판단 (clsRouteJudge, 조건 대상 = 항목)
'             R5  분기 COM 으로 "바코드 + 끝 문자(대상별 CR / LF / CRLF)" 전송, 대상 프로그램 창을 앞으로
'             R6  검사 중 표시, 종료 신호 대기 (대상별 타임아웃, 0 = 무제한)
'             R7  대상별 종료 신호(EndMessage, 예 : END) 수신 → 대기 → Router 창을 앞으로
'             화면 : Router_UI_시안_v1.html (2026-10-26 승인) - 제목줄, 큰 상태 패널(대기 초록 / 검사 중 주황 /
'                    오류 빨강 / 조회 중 파랑), 바코드 칸, 제품 정보 카드(tblCard 설정대로), 분기 대상 카드, 로그, 평면 버튼
'                    제목줄 [설명서] (tblConfig ManualPath), 아래 [화면 설정] (frmDisplaySetting, 2026-10-28)
'                    분기 규칙은 검사에 필요한 정보가 아니라 로그에만 남긴다 (2026-10-28 지시)
'             대상 확인 : 보내기 전에 분기 대상 프로그램(프로세스명)이 실행 중인지 확인, 꺼져 있으면 보내지 않음.
'                    END 를 기다리는 중 대상이 꺼지면 그 대상은 기다리지 않음. 기다리는 동안 바코드 칸 잠금 (2026-10-28 지시)
'                    제목줄 가운데 설비명 (COM 설정에서 입력, 2026-10-26 추가)
'             설정 / 로그 : 저장소 프로젝트 폴더의 Files\RouterSetting.mdb, System_Log\ (clsRouterPath, 공통 규약 7항)
'                    exe 를 바탕화면 등으로 복사해도 C:\Barcode_Router\Barcode_Router\ 를 읽는다. EtherComN.dll 은 exe 에 내장
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7항 / A8_FAM3_Router_구조설계_v1.1.xlsx
'==========================================================

Public Class frmRouterMain

    'Program Ver ==========================================
    Private Const Version As String = "0.00"    '2026.10.26 Ver 0.00 YSH

    '바코드 종류(작업자 / 제품 등)는 tblBarcodeType, 종료 신호는 분기 대상별 EndMessage 로 정한다 (2026-10-28, 하드코딩 하지 않음)
    Private Shared ReadOnly Property HINT_BUSY() As String
        Get
            Return clsLang.L("검사가 끝나면 이 화면이 다시 앞으로 옵니다", "検査が終わるとこの画面が再び前面に出ます")
        End Get
    End Property

    Dim mAppPath As String = clsRouterPath.DataFolder   '저장소 프로젝트 폴더 (bin 의 한 단계 위, bin 밖에서 실행하면 C:\Barcode_Router\Barcode_Router\)
    Dim mSetting As clsRouterSetting
    Dim mReaderPort As clsRouterPort                                        '바코드리더 COM (COM 모드일 때만)
    Dim mDestPorts As New Dictionary(Of Integer, clsRouterPort)             'DestID → 분기 대상 COM
    Dim mPending As New Dictionary(Of Integer, DateTime)                    'END 를 기다리는 DestID → 전송 시각
    Dim mLastEnd As New Dictionary(Of Integer, DateTime)                    'DestID → 마지막 END 수신 시각
    Dim mPendingText As String = ""
    Dim mOperator As String = "-"                                           '마지막으로 모든 대상에 보낸 바코드 (종류명 + 값)
    Dim mTick As Integer = 0
    Dim mCardValues As New List(Of KeyValuePair(Of String, Label))          '카드 값 칸 (항목ID, Label) - tblCard 로 만든다
    Dim mCardsHeight As Integer = -1                                        '지금 카드 영역 높이 (아래 컨트롤을 그만큼 내린다)

    Private Delegate Sub LineDelegate(ByVal port As clsRouterPort, ByVal line As String)

    'Load / Close ==========================================================================================================
    Private Sub frmRouterMain_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Me.Text = "Barcode Router  Ver : " & Version
        Label_Ver.Text = "Ver " & Version
        Label_Ver.Left = Label_Title.Right + 6
        clsUiTheme.StyleButton(Button_Rule, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Com, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Display, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Manual, clsUiTheme.ButtonKind.Accent)
        Call UpdateReleaseButton()
        WriteLog("Program Start Ver " & Version)
        WriteLog("Data Folder : " & mAppPath & "  (exe : " & Application.ExecutablePath & ")")     '어느 폴더를 쓰는지 로그로 남김

        Call LoadAndOpen()

        '오래된 로그 월 폴더 삭제 (tblConfig LogKeepMonths, 비우면 지우지 않음)
        Dim deleted As List(Of String) = clsRouterLog.DeleteOld(mAppPath, mSetting.GetConfigInt(clsRouterSetting.CFG_LOG_KEEP_MONTHS))
        If deleted.Count > 0 Then WriteLog("Log Delete : " & String.Join(", ", deleted))
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
            WriteLog(clsLang.L("Setting Load OK : 분기 대상 ", "Setting Load OK : 分岐先 ") & mSetting.Destinations.Count & clsLang.L(" / 규칙 ", " / ルール ") & mSetting.Rules.Count)
        Catch ex As Exception
            WriteLog("Setting Load NG : " & ex.Message)
            MessageBox.Show(clsLang.L("설정을 읽지 못했습니다.", "設定を読めませんでした。") & vbCrLf & ex.Message, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Error)
            mSetting = New clsRouterSetting(mAppPath & "Files\RouterSetting.mdb")
        End Try
        For Each note As String In mSetting.LoadNotes
            WriteLog("Setting : " & note)
        Next

        '화면 설정 (언어 / 글꼴 / 카드) - 언어는 한국어 PC / 일본어 PC (tblConfig UiLanguage, 2026-10-28)
        clsLang.SetLanguage(mSetting.GetConfig(clsRouterSetting.CFG_UI_LANGUAGE))
        clsUiTheme.SetFontName(mSetting.GetConfig(clsRouterSetting.CFG_UI_FONT))
        Call ApplyLanguageTexts()
        Call UpdateHeader()             'ApplyLanguageTexts 가 Designer 글자로 되돌린 머리줄을 다시 그림
        Call clsUiTheme.ApplyFont(Me)
        Call BuildCards()

        Call ShowEquipment()

        '바코드리더 COM
        Label_InputMode.Text = clsLang.L("HID 입력", "HID入力")
        If mSetting.Reader.InputMode = clsReaderSetting.MODE_COM Then
            Label_InputMode.Text = mSetting.Reader.Port & clsLang.L(" 입력", " 入力")
            Try
                mReaderPort = New clsRouterPort
                AddHandler mReaderPort.LineReceived, AddressOf Port_LineReceived
                mReaderPort.Open(mSetting.Reader.Port, mSetting.Reader.Setting)
                WriteLog("Reader COM Open OK : " & mSetting.Reader.Port)
            Catch ex As Exception
                WriteLog("Reader COM Open NG : " & mSetting.Reader.Port & " " & ex.Message)
                MessageBox.Show(clsLang.L("바코드리더 COM 을 열지 못했습니다 : ", "バーコードリーダのCOMを開けませんでした : ") & mSetting.Reader.Port & vbCrLf & ex.Message, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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

        If d.EndMessage.Trim <> "" AndAlso String.Equals(line.Trim, d.EndMessage.Trim, StringComparison.OrdinalIgnoreCase) Then     '분기 대상별 종료 신호 (tblDestination EndMessage)
            Call OnEnd(d)
        Else
            WriteLog("Recv " & d.DisplayName & " : " & line)
        End If
    End Sub

    Private Sub HandleBarcode(ByVal code As String)
        '==========================================================================================
        'Function : 바코드 한 건 처리 (R1 ~ R6)
        '==========================================================================================
        TextBox_BarCode.Clear()         '찍은 바코드는 바로 지운다 - 오류여도 다시 찍을 수 있게 (2026-10-29 지시). 찍은 값은 로그 / 상태 패널에 남는다
        If mPending.Count > 0 Then
            WriteLog(clsLang.L("Barcode skip (검사 중) : ", "Barcode skip (検査中) : ") & code)
            Call SetStatus(clsUiTheme.StatusKind.Busy, mPendingText, clsLang.L("검사 중 · 끝난 뒤 다시 찍어 주세요", "検査中 · 終了後にもう一度読んでください"), clsLang.L("지금 찍은 바코드는 보내지 않았습니다 : ", "今読んだバーコードは送っていません : ") & code)
            TextBox_BarCode.Clear()
            Exit Sub
        End If

        WriteLog("Barcode : " & code)

        '바코드 종류 (tblBarcodeType) ---------------------------------------------
        Dim bt As clsBarcodeType = clsRouteData.ClassifyBarcode(mSetting.BarcodeTypes, code)
        If bt Is Nothing Then
            WriteLog(clsLang.L("Barcode NG : 맞는 바코드 종류 없음 (", "Barcode NG : 合うバーコード種類なし (") & code.Length & clsLang.L("자리)", "桁)"))
            Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("등록되지 않은 바코드입니다", "登録されていないバーコードです"), clsLang.L("오류", "エラー"), code & " (" & code.Length & clsLang.L("자리) · 화면 설정 > 바코드 종류를 확인해 주세요", "桁) · 画面設定 > バーコード種類を確認してください"))
            TextBox_BarCode.Clear()
            Exit Sub
        End If
        WriteLog("Barcode Type : " & bt.TypeName & " (" & bt.ActionType & ")")

        Select Case bt.ActionType
            Case clsBarcodeType.ACT_BROADCAST
                Call SendBroadcast(code, bt)
            Case clsBarcodeType.ACT_ROUTE
                Call SendRoute(code, bt)
            Case Else
                Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("보내지 않는 바코드입니다 : ", "送らないバーコードです : ") & bt.TypeName, clsLang.L("오류", "エラー"), code)
                TextBox_BarCode.Clear()
        End Select
    End Sub

    Private Sub SendRoute(ByVal code As String, ByVal bt As clsBarcodeType)
        '==========================================================================================
        'Function : PLASMA 조회 → 항목 값 (tblField) → 카드 표시 → 분기 규칙 → 한 대상에 전송 (R2 ~ R6)
        '==========================================================================================
        Call ClearInfo()
        Dim cPsm As New cls100_PLASMA_Order
        Try
            Call SetStatus(clsUiTheme.StatusKind.Info, clsLang.L("PLASMA 조회 중", "PLASMA照会中"), clsLang.L("조회 중 · ", "照会中 · ") & bt.TypeName, code)
            Application.DoEvents()
            cPsm.SetServer(mSetting.GetConfig(clsRouterSetting.CFG_PLASMA_HOST1), mSetting.GetConfig(clsRouterSetting.CFG_PLASMA_HOST2),
                           mSetting.GetConfigInt(clsRouterSetting.CFG_PLASMA_PORT), mSetting.GetConfig(clsRouterSetting.CFG_PLASMA_COMMAND))
            cPsm.BARCODE_TEXT = code
            Call cPsm.GET_ORDER_INFORMAION()
        Catch ex As Exception
            WriteLog("PLASMA NG : " & code & " " & ex.Message)
            Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("PLASMA 조회 실패", "PLASMA照会失敗"), clsLang.L("오류", "エラー") & " · " & code, ex.Message)     '바코드 칸은 비우므로 찍은 값은 여기에 보인다
            TextBox_BarCode.Clear()
            Exit Sub
        End Try

        Dim values As Dictionary(Of String, String) = clsRouteData.BuildValues(mSetting.Fields, code, cPsm.RAW_DATA)
        Call ShowCards(values)
        WriteLog("PLASMA OK : " & String.Join(" ", mCardValues.Select(Function(kv) kv.Key & "=" & ValueOf(values, kv.Key))))

        'Judge (R3, R4) -----------------------------------------------------------
        Dim ru As clsRouteRule = clsRouteJudge.Judge(mSetting.Rules, values)
        If ru Is Nothing Then
            WriteLog(clsLang.L("Judge NG : 분기 대상 없음", "Judge NG : 分岐先なし"))
            Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("분기 대상 없음", "分岐先なし"), clsLang.L("오류", "エラー"), clsLang.L("분기 조건 설정을 확인해 주세요", "分岐条件設定を確認してください"))
            TextBox_BarCode.Clear()
            Exit Sub
        End If

        Dim d As clsDestination = mSetting.FindDestination(ru.DestID)
        If d Is Nothing OrElse Not d.Enabled OrElse Not mDestPorts.ContainsKey(d.DestID) Then
            WriteLog(clsLang.L("Judge NG : 규칙 '", "Judge NG : ルール '") & ru.Description & clsLang.L("' 의 분기 대상이 없거나 사용 안 함", "' の分岐先がないか未使用"))
            Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("분기 대상 설정 오류", "分岐先設定エラー"), clsLang.L("오류", "エラー"), clsLang.L("규칙 '", "ルール '") & ru.Description & clsLang.L("' 의 분기 COM 을 확인해 주세요", "' の分岐COMを確認してください"))
            TextBox_BarCode.Clear()
            Exit Sub
        End If
        WriteLog("Judge OK : " & ru.ToString & " → " & d.DisplayName)     '분기 규칙은 검사에 필요한 정보가 아니라 화면에 표시하지 않고 로그에만 남긴다

        '대상 프로그램 실행 확인 ---------------------------------------------------
        If Not IsDestRunning(d) Then
            WriteLog(clsLang.L("Send skip : 프로그램 꺼짐 ", "Send skip : プログラム停止中 ") & d.DisplayName & " (" & d.ProcessName & ")")
            Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("프로그램이 실행되어 있지 않습니다 : ", "プログラムが起動していません : ") & d.DestName, clsLang.L("오류", "エラー"), d.ProcessName & clsLang.L(" 을(를) 실행한 뒤 다시 찍어 주세요", " を起動してからもう一度読んでください"))
            TextBox_BarCode.Clear()
            Exit Sub
        End If

        'Send (R5, R6) ------------------------------------------------------------
        If Not SendTo(d, code) Then
            Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("전송 실패 : ", "送信失敗 : ") & d.DestName, clsLang.L("오류", "エラー"), d.Port & clsLang.L(" 를 확인해 주세요", " を確認してください"))
            TextBox_BarCode.Clear()
            Exit Sub
        End If

        mPendingText = d.DestName
        Call SetStatus(clsUiTheme.StatusKind.Busy, d.DestName, clsLang.L("검사 중 · ", "検査中 · ") & Format(DateTime.Now, "HH:mm:ss") & clsLang.L(" 전송", " 送信"), HINT_BUSY)
        Call RefreshDestList()

        If Not clsWindowControl.BringProcessToFront(d.ProcessName) Then
            WriteLog(clsLang.L("Window NG : 프로세스 창을 찾지 못함 ", "Window NG : プロセスのウィンドウが見つからない ") & d.ProcessName)
        End If
    End Sub

    Private Shared Function ValueOf(ByVal values As Dictionary(Of String, String), ByVal key As String) As String
        Dim v As String = ""
        If values IsNot Nothing Then values.TryGetValue(key, v)
        Return If(v, "")
    End Function

    Private Sub SendBroadcast(ByVal code As String, ByVal bt As clsBarcodeType)
        '==========================================================================================
        'Function : 바코드를 사용 중인 모든 분기 대상으로 전송 (예 : 작업자 바코드, 10_확인항목 Q1 = A)
        '           A8 : 작업자 갱신 / FAM3 : 작업자 갱신 + 주명판 송신. 각 대상의 종료 신호를 기다린다
        '==========================================================================================
        Call ClearInfo()
        Dim names As New List(Of String)

        For Each d As clsDestination In mSetting.Destinations
            If Not d.Enabled OrElse Not mDestPorts.ContainsKey(d.DestID) Then Continue For
            If Not IsDestRunning(d) Then
                WriteLog(clsLang.L("Send skip : 프로그램 꺼짐 ", "Send skip : プログラム停止中 ") & d.DisplayName & " (" & d.ProcessName & ")")     '꺼진 프로그램은 종료 신호가 오지 않으므로 보내지 않는다
                Continue For
            End If
            If SendTo(d, code) Then names.Add(d.DestName)
        Next

        If names.Count = 0 Then
            Call SetStatus(clsUiTheme.StatusKind.Error, bt.TypeName & clsLang.L(" 바코드를 보낼 곳이 없습니다", " バーコードの送り先がありません"), clsLang.L("오류", "エラー"), clsLang.L("검사 프로그램 실행 / COM 설정을 확인해 주세요", "検査プログラムの起動 / COM設定を確認してください"))
            TextBox_BarCode.Clear()
            Exit Sub
        End If

        mOperator = bt.TypeName & " " & code
        Call UpdateHeader()
        mPendingText = bt.TypeName
        Call SetStatus(clsUiTheme.StatusKind.Busy, bt.TypeName & " : " & code, clsLang.L("검사 중 · ", "検査中 · ") & String.Join(", ", names), HINT_BUSY)
        Call RefreshDestList()
    End Sub

    Private Function IsDestRunning(ByVal d As clsDestination) As Boolean
        '==========================================================================================
        'Function : 분기 대상 프로그램이 실행 중인지 (분기 대상의 프로세스명으로 확인)
        'Remarks  : 프로세스명을 비워 둔 분기 대상은 확인하지 않고 보낸다 (확인할 수 없으므로)
        '           꺼진 프로그램에 보내면 END 가 오지 않아 '검사 중' 으로 멈춰 있게 된다 (2026-10-28 지시, A안)
        '==========================================================================================
        If d.ProcessName.Trim = "" Then Return True
        Return clsWindowControl.IsProcessRunning(d.ProcessName)
    End Function

    Private Function SendTo(ByVal d As clsDestination, ByVal code As String) As Boolean
        Dim p As clsRouterPort = mDestPorts(d.DestID)
        Try
            If Not p.IsOpen Then p.Open(d.Port, d.Setting) '닫혀 있으면 한 번 더 열어 본다
            p.SendLine(code, d.LineEndChars())      '끝 문자는 분기 대상별 설정 (tblDestination LineEnd)
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
        Dim closed As New List(Of String)
        For Each id As Integer In mPending.Keys.ToList
            Dim d As clsDestination = mSetting.FindDestination(id)
            If d Is Nothing Then mPending.Remove(id) : changed = True : Continue For
            If Not IsDestRunning(d) Then
                '검사 중에 대상 프로그램이 꺼지면 END 가 오지 않으므로 그 대상은 기다리지 않는다
                mPending.Remove(id)
                changed = True
                closed.Add(d.DestName)
                WriteLog(clsLang.L("Wait Release (프로그램 꺼짐) : ", "Wait Release (プログラム停止) : ") & d.DisplayName & " (" & d.ProcessName & ")")
                Continue For
            End If
            If d.EndTimeoutSec <= 0 Then Continue For
            If (DateTime.Now - mPending(id)).TotalSeconds > d.EndTimeoutSec Then
                mPending.Remove(id)
                changed = True
                WriteLog("END Timeout : " & d.DisplayName & " (" & d.EndTimeoutSec & "s)")
                MessageBox.Show(clsLang.L("END 를 받지 못했습니다 (타임아웃 ", "ENDを受信できませんでした (タイムアウト ") & d.EndTimeoutSec & clsLang.L("초) : ", "秒) : ") & d.DisplayName, "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Next

        If Not changed Then Exit Sub
        Call RefreshDestList()
        If mPending.Count = 0 Then
            Call SetReady()
            Call ShowRouterFront()
            If closed.Count > 0 Then Call SetStatus(clsUiTheme.StatusKind.Error, clsLang.L("검사 프로그램이 꺼져 대기를 풀었습니다", "検査プログラムが停止したため待機を解除しました"), clsLang.L("오류", "エラー"), String.Join(", ", closed) & clsLang.L(" 을(를) 실행한 뒤 다시 찍어 주세요", " を起動してからもう一度読んでください"))
        End If
    End Sub

    Private Sub Button_Release_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Release.Click
        If mPending.Count = 0 Then
            TextBox_BarCode.Focus()
            Exit Sub
        End If
        If MessageBox.Show(clsLang.L("END 를 기다리지 않고 대기 상태로 돌아갑니까?", "ENDを待たずに待機状態に戻りますか?"), "Barcode Router", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub
        WriteLog(clsLang.L("Wait Release (수동) : ", "Wait Release (手動) : ") & mPendingText)
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
        Call SetReady()                 '설정을 바꾼 뒤 상태 패널 / 바코드 칸을 대기로 (언어가 바뀌었을 수 있음)
    End Sub

    Private Sub Button_Com_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Com.Click
        If Not CanOpenSetting() Then Exit Sub
        Call CloseAllPorts()            '포트 시험을 할 수 있도록 Router 의 포트를 모두 닫는다
        Using f As New frmComSetting(mSetting)
            If f.ShowDialog(Me) = DialogResult.OK Then WriteLog("COM Setting Saved")
        End Using
        Call LoadAndOpen()
        Call SetReady()                 '설정을 바꾼 뒤 상태 패널 / 바코드 칸을 대기로 (언어가 바뀌었을 수 있음)
    End Sub

    Private Sub Button_Display_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Display.Click
        If Not CanOpenSetting() Then Exit Sub
        Call CloseAllPorts()            '설정 가져오기로 설정 파일을 바꿀 수 있으므로 포트를 닫는다
        Using f As New frmDisplaySetting(mSetting)
            If f.ShowDialog(Me) = DialogResult.OK Then WriteLog("Display Setting Saved")
        End Using
        Call LoadAndOpen()
        Call SetReady()                 '설정을 바꾼 뒤 상태 패널 / 바코드 칸을 대기로 (언어가 바뀌었을 수 있음)
    End Sub

    Private Sub Button_Manual_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Manual.Click
        '==========================================================================================
        'Function : 사용자 설명서 열기 (tblConfig ManualPath - 전체 경로 또는 데이터 폴더 기준 상대 경로)
        '==========================================================================================
        Dim path As String = mSetting.GetConfig(clsRouterSetting.CFG_MANUAL_PATH)
        If path = "" Then
            MessageBox.Show(clsLang.L("사용자 설명서 파일이 설정되어 있지 않습니다.", "取扱説明書ファイルが設定されていません。") & vbCrLf & clsLang.L("화면 설정 > 기타 > ManualPath 에 넣어 주세요.", "画面設定 > その他 > ManualPath に入力してください。"), clsLang.L("사용자 설명서", "取扱説明書"), MessageBoxButtons.OK, MessageBoxIcon.Information)
            Exit Sub
        End If
        If Not IO.Path.IsPathRooted(path) Then path = IO.Path.Combine(mAppPath, path)
        If Not IO.File.Exists(path) Then
            MessageBox.Show(clsLang.L("사용자 설명서 파일이 없습니다.", "取扱説明書ファイルがありません。") & vbCrLf & path, clsLang.L("사용자 설명서", "取扱説明書"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Try
            Process.Start(New ProcessStartInfo(path) With {.UseShellExecute = True})
            WriteLog("Manual Open : " & path)
        Catch ex As Exception
            MessageBox.Show(clsLang.L("사용자 설명서를 열지 못했습니다.", "取扱説明書を開けませんでした。") & vbCrLf & ex.Message, clsLang.L("사용자 설명서", "取扱説明書"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function CanOpenSetting() As Boolean
        If mPending.Count = 0 Then Return True
        MessageBox.Show(clsLang.L("검사 중에는 설정을 바꿀 수 없습니다.", "検査中は設定を変更できません。"), "Barcode Router", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Return False
    End Function

    'Display ===============================================================================================================
    Private Sub SetReady()
        mPendingText = ""
        Call UpdateReleaseButton()
        Call SetStatus(clsUiTheme.StatusKind.Ready, clsLang.L("바코드를 찍어 주세요", "バーコードを読んでください"), clsLang.L("대기", "待機"), "")
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
        '           END 를 기다리는 동안은 바코드 칸도 막는다 (재입력 방지, 2026-10-28 지시)
        '==========================================================================================
        Dim waiting As Boolean = (mPending.Count > 0)
        TextBox_BarCode.ReadOnly = waiting
        TextBox_BarCode.BackColor = If(waiting, clsUiTheme.CardBack, Color.White)
        Panel_BarCode.BackColor = TextBox_BarCode.BackColor
        Label_InputMode.Text = If(waiting, clsLang.L("입력 잠금", "入力ロック"), If(mReaderPort IsNot Nothing, mSetting.Reader.Port & clsLang.L(" 입력", " 入力"), clsLang.L("HID 입력", "HID入力")))
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
            Label_Equipment.Text = clsLang.L("설비명 미설정 (COM 설정에서 입력)", "設備名未設定 (COM設定で入力)")
            Label_Equipment.Font = clsUiTheme.UiFont(9.5!)
            Label_Equipment.ForeColor = clsUiTheme.HeaderSub
            Label_Equipment.BackColor = clsUiTheme.HeaderBack
            Me.Text = "Barcode Router  Ver : " & Version
        Else
            Label_Equipment.Text = name
            Label_Equipment.Font = clsUiTheme.UiFont(13.0!, FontStyle.Bold)
            Label_Equipment.ForeColor = Color.White
            Label_Equipment.BackColor = clsUiTheme.PrimaryBack
            Me.Text = "Barcode Router  Ver : " & Version & "  [" & name & "]"     '작업표시줄에서도 설비를 알 수 있게
        End If
    End Sub

    Private Sub UpdateHeader()
        Label_HeaderInfo.Text = mOperator & "  ·  " & Format(DateTime.Now, "yyyy-MM-dd HH:mm")       'mOperator = 모든 대상에 보낸 마지막 바코드 (바코드 종류명 + 값)
    End Sub

    Private Sub ClearInfo()
        For Each kv As KeyValuePair(Of String, Label) In mCardValues
            kv.Value.Text = ""
        Next
    End Sub

    Private Sub ShowCards(ByVal values As Dictionary(Of String, String))
        For Each kv As KeyValuePair(Of String, Label) In mCardValues
            kv.Value.Text = ValueOf(values, kv.Key)
        Next
    End Sub

    Private Sub BuildCards()
        '==========================================================================================
        'Function : 제품 정보 카드를 tblCard 설정대로 만든다 (2026-10-28 지시 : 카드는 유저가 설정)
        '           - 사용인 카드를 번호 순으로, 한 줄 CardsPerRow 개 (설정이 없으면 한 줄에 모두), 폭은 나눠 씀
        '           - 줄 수가 바뀌면 아래 컨트롤(분기 대상 / 로그 / 버튼)을 내리거나 올리고 창 높이를 맞춘다
        '           - 설정이 없으면 빈 카드 영역
        '==========================================================================================
        Const CARD_H As Integer = 58, GAP As Integer = 8, AREA_W As Integer = 872
        mCardValues.Clear()
        FlowLayoutPanel_Cards.SuspendLayout()
        For Each c As Control In FlowLayoutPanel_Cards.Controls.Cast(Of Control).ToList
            FlowLayoutPanel_Cards.Controls.Remove(c)
            c.Dispose()
        Next

        Dim cards As List(Of clsCard) = mSetting.Cards.Where(Function(x) x.Enabled AndAlso x.FieldID <> "").OrderBy(Function(x) x.CardNo).ToList
        Dim perRow As Integer = mSetting.GetConfigInt(clsRouterSetting.CFG_CARDS_PER_ROW)
        If perRow <= 0 Then perRow = Math.Max(1, cards.Count)
        Dim rows As Integer = Math.Max(1, CInt(Math.Ceiling(cards.Count / perRow)))
        Dim w As Integer = (AREA_W - (perRow - 1) * GAP) \ perRow

        For i As Integer = 0 To cards.Count - 1
            Dim cd As clsCard = cards(i)
            Dim f As clsField = mSetting.FindField(cd.FieldID)
            Dim cap As String = If(cd.Caption.Trim <> "", cd.Caption, If(f Is Nothing, cd.FieldID, f.Caption))
            Dim lastCol As Boolean = ((i + 1) Mod perRow = 0)
            Dim lastRow As Boolean = (i \ perRow = rows - 1)
            Dim p As New Panel With {.Size = New Size(w, CARD_H), .BackColor = clsUiTheme.CardBack,
                                     .Margin = New Padding(0, 0, If(lastCol, 0, GAP), If(lastRow, 0, GAP))}
            Dim lc As New Label With {.Text = cap, .Font = clsUiTheme.UiFont(8.5!), .ForeColor = clsUiTheme.TextSub,
                                      .Location = New Point(10, 6), .Size = New Size(w - 20, 18), .AutoEllipsis = True}
            Dim lv As New Label With {.Text = "", .Font = clsUiTheme.UiFont(12.0!, FontStyle.Bold), .ForeColor = clsUiTheme.TextMain,
                                      .Location = New Point(10, 26), .Size = New Size(w - 16, 26), .AutoEllipsis = True}
            p.Controls.Add(lv)
            p.Controls.Add(lc)
            FlowLayoutPanel_Cards.Controls.Add(p)
            mCardValues.Add(New KeyValuePair(Of String, Label)(cd.FieldID, lv))
        Next
        FlowLayoutPanel_Cards.ResumeLayout()

        '카드 줄 수로 높이가 바뀐 만큼 아래 컨트롤을 옮기고 창 높이를 맞춘다
        Dim newH As Integer = rows * CARD_H + (rows - 1) * GAP
        If mCardsHeight < 0 Then mCardsHeight = FlowLayoutPanel_Cards.Height
        Dim delta As Integer = newH - mCardsHeight
        If delta <> 0 Then
            Dim bottomOld As Integer = FlowLayoutPanel_Cards.Top + mCardsHeight
            For Each c As Control In Me.Controls
                If c IsNot FlowLayoutPanel_Cards AndAlso c.Top >= bottomOld Then c.Top += delta
            Next
            Me.ClientSize = New Size(Me.ClientSize.Width, Me.ClientSize.Height + delta)
        End If
        FlowLayoutPanel_Cards.Height = newH
        mCardsHeight = newH
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
                kind = clsUiTheme.StatusKind.Info : state = clsLang.L("사용 안 함", "未使用")
            ElseIf Not mDestPorts.ContainsKey(d.DestID) OrElse Not mDestPorts(d.DestID).IsOpen Then
                kind = clsUiTheme.StatusKind.Error : state = clsLang.L("COM 실패", "COM失敗")
            ElseIf mPending.ContainsKey(d.DestID) Then
                kind = clsUiTheme.StatusKind.Busy : state = clsLang.L("검사 중", "検査中")
            Else
                kind = clsUiTheme.StatusKind.Ready : state = clsLang.L("대기", "待機")
            End If

            Dim running As Boolean = clsWindowControl.IsProcessRunning(d.ProcessName)
            If d.Enabled AndAlso Not running AndAlso kind = clsUiTheme.StatusKind.Ready Then
                kind = clsUiTheme.StatusKind.Error : state = clsLang.L("프로그램 꺼짐", "プログラム停止")
            End If

            Dim detail As String = d.Port & "  ·  " & If(d.ProcessName = "", "-", d.ProcessName) & If(running, clsLang.L(" 실행 중", " 起動中"), clsLang.L(" 꺼짐", " 停止"))
            If mPending.ContainsKey(d.DestID) Then detail &= "  ·  " & Format(mPending(d.DestID), "HH:mm:ss") & clsLang.L(" 전송", " 送信")
            If mLastEnd.ContainsKey(d.DestID) Then detail &= clsLang.L("  ·  마지막 END ", "  ·  最終END ") & Format(mLastEnd(d.DestID), "HH:mm:ss")

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
        Dim lblName As New Label With {.Text = name, .Font = clsUiTheme.UiFont(11.0!, FontStyle.Bold), .ForeColor = clsUiTheme.TextMain,
                                       .Location = New Point(14, 7), .Size = New Size(290, 24), .AutoEllipsis = True}
        Dim lblState As New Label With {.Text = state, .Font = clsUiTheme.UiFont(8.5!), .ForeColor = c.Main, .BackColor = c.Back,
                                        .TextAlign = ContentAlignment.MiddleCenter, .Location = New Point(318, 9), .Size = New Size(100, 20)}
        Dim lblDetail As New Label With {.Text = detail, .Font = clsUiTheme.UiFont(8.5!), .ForeColor = clsUiTheme.TextSub,
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


    Private Sub ApplyLanguageTexts()
        '==========================================================================================
        'Function : Designer 에 적힌 한국어 글자를 언어 설정(clsLang)에 맞게 바꾼다 (Designer 파일은 그대로, 2026-10-28)
        '==========================================================================================
        Me.Label_HeaderInfo.Text = clsLang.L("작업자 - · 2026-10-26 00:00", "- · 2026-10-26 00:00")
        Me.Label_Equipment.Text = clsLang.L("설비명", "設備名")
        Me.Button_Manual.Text = clsLang.L("ⓘ 설명서", "ⓘ 説明書")
        Me.Label_Status.Text = clsLang.L("바코드를 찍어 주세요", "バーコードを読んでください")
        Me.Label_StatusSub.Text = clsLang.L("대기", "待機")
        Me.Label_InputMode.Text = clsLang.L("HID 입력", "HID入力")
        Me.Label_DestCap.Text = clsLang.L("분기 대상", "分岐先")
        Me.Label_LogCap.Text = clsLang.L("로그", "ログ")
        Me.Button_Rule.Text = clsLang.L("⑂  분기 조건 설정", "⑂  分岐条件設定")
        Me.Button_Com.Text = clsLang.L("⇄  COM 설정", "⇄  COM設定")
        Me.Button_Display.Text = clsLang.L("⚙  화면 설정", "⚙  画面設定")
        Me.Button_Release.Text = clsLang.L("■  대기 해제", "■  待機解除")
    End Sub

End Class
