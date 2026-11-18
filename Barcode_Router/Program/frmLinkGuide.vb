Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmLinkGuide
'Function   : 연동 가이드 - 검사 프로그램을 Router 에 붙이는 순서를 단계 버튼으로 보여 주고, 필요한 파일 / 명령 / 코드를 만들어 준다
'Framwork   : VB2022 Framework4.8
'Memo       : 2026-11-13 지시 - 각 검사 프로그램에 같은 수정(설정 파일, 가상 COM, 클래스, 코드 몇 줄)을 유저가 직접 따라 할 수 있게
'             - 단계 : ① 가상 COM  ② RouterLink.mdb  ③ clsRouterLink.vb  ④ 코드 추가  ⑤ Router 등록  ⑥ 확인
'             - 단계 설명 글은 Files\LinkKit\Guide_KO.txt (일본어 화면이면 Guide_JA.txt, 없으면 KO) 에서 읽는다 → 글을 고쳐도 다시 빌드할 필요 없음
'             - 배포 파일 : Files\LinkKit\RouterLink.mdb (빈 양식), Files\LinkKit\clsRouterLink.vb (공통판)
'             - com0com 설치 / 명령 실행은 하지 않는다 (관리자 권한 · 드라이버) - 명령을 만들어 복사만 한다
'             - 화면은 Designer 없이 코드로 만든다
'Issue      : 2026/11/13  Rev 0.00    YSH       New Release
'==========================================================
Imports System.Data.OleDb
Imports System.IO

Public Class frmLinkGuide
    Inherits Form

    Private ReadOnly mKitFolder As String = clsRouterPath.DataFolder & "Files\LinkKit\"
    Private ReadOnly mOpenCom As Action
    Private ReadOnly mOpenRule As Action
    Private ReadOnly mSteps As New List(Of Button)
    Private mGuide As New Dictionary(Of Integer, String)
    Private mStep As Integer = 0

    Private ReadOnly STEP_TITLES() As String = {
        clsLang.L("① 가상 COM 만들기", "① 仮想COM作成"),
        clsLang.L("② RouterLink.mdb 넣기", "② RouterLink.mdb 配置"),
        clsLang.L("③ clsRouterLink.vb 추가", "③ clsRouterLink.vb 追加"),
        clsLang.L("④ 코드 추가", "④ コード追加"),
        clsLang.L("⑤ Router 에 등록", "⑤ Routerに登録"),
        clsLang.L("⑥ 연동 확인", "⑥ 連携確認")}

    Private Panel_Steps As New Panel With {.Location = New Point(14, 56), .Size = New Size(220, 570)}
    Private Label_StepTitle As New Label With {.Location = New Point(250, 56), .Size = New Size(736, 30), .ForeColor = clsUiTheme.InputText}
    Private TextBox_Guide As New TextBox With {.Location = New Point(250, 90), .Size = New Size(736, 170), .Multiline = True, .ReadOnly = True,
                                               .ScrollBars = ScrollBars.Vertical, .BackColor = clsUiTheme.CardBack, .BorderStyle = BorderStyle.None}
    Private Panel_Action As New Panel With {.Location = New Point(250, 270), .Size = New Size(736, 356)}

    Public Sub New(ByVal openCom As Action, ByVal openRule As Action)
        mOpenCom = openCom
        mOpenRule = openRule
        Me.Text = clsLang.L("연동 가이드", "連携ガイド")
        Me.ClientSize = New Size(1000, 640)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False : Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.Font = clsUiTheme.UiFont(9.5!)

        Dim header As New Panel With {.Dock = DockStyle.Top, .Height = 42, .BackColor = clsUiTheme.HeaderBack}
        header.Controls.Add(New Label With {.Text = clsLang.L("⛓  연동 가이드", "⛓  連携ガイド"), .AutoSize = True, .Location = New Point(12, 9), .ForeColor = clsUiTheme.HeaderText,
                                            .Font = clsUiTheme.UiFont(12.0!, FontStyle.Bold)})
        header.Controls.Add(New Label With {.Text = clsLang.L("검사 프로그램을 Router 에 붙이는 순서 - 위에서부터 차례로", "検査プログラムをRouterにつなぐ手順 - 上から順に"),
                                            .AutoSize = True, .Location = New Point(190, 14), .ForeColor = clsUiTheme.HeaderSub})
        Me.Controls.Add(header)

        For i As Integer = 0 To STEP_TITLES.Length - 1
            Dim b As New Button With {.Text = STEP_TITLES(i), .Tag = i, .Location = New Point(0, i * 68), .Size = New Size(220, 60),
                                      .TextAlign = ContentAlignment.MiddleLeft, .Padding = New Padding(10, 0, 0, 0), .Font = clsUiTheme.UiFont(10.5!, FontStyle.Bold)}
            AddHandler b.Click, AddressOf Step_Click
            mSteps.Add(b)
            Panel_Steps.Controls.Add(b)
        Next
        Panel_Steps.Controls.Add(New Label With {.Text = clsLang.L("파일 위치 : ", "ファイル : ") & vbCrLf & mKitFolder, .Location = New Point(0, 420), .Size = New Size(220, 80),
                                                 .ForeColor = clsUiTheme.TextSub})
        Label_StepTitle.Font = clsUiTheme.UiFont(13.0!, FontStyle.Bold)
        TextBox_Guide.Font = clsUiTheme.UiFont(10.0!)
        Me.Controls.AddRange(New Control() {Panel_Steps, Label_StepTitle, TextBox_Guide, Panel_Action})
    End Sub

    Private Sub frmLinkGuide_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Call LoadGuide()
        Call ShowStep(0)
    End Sub

    Private Sub Step_Click(ByVal sender As Object, ByVal e As EventArgs)
        Call ShowStep(CInt(DirectCast(sender, Button).Tag))
    End Sub

    '==========================================================================================
    ' 단계 설명 글 (Guide_KO.txt / Guide_JA.txt) - [1] ~ [6] 로 나눈다
    '==========================================================================================
    Private Sub LoadGuide()
        mGuide.Clear()
        Dim f As String = mKitFolder & If(clsLang.IsJA AndAlso File.Exists(mKitFolder & "Guide_JA.txt"), "Guide_JA.txt", "Guide_KO.txt")
        If Not File.Exists(f) Then Exit Sub
        Dim cur As Integer = -1
        Dim sb As New System.Text.StringBuilder
        For Each ln As String In File.ReadAllLines(f, System.Text.Encoding.UTF8)
            Dim t As String = ln.Trim
            If t.Length >= 3 AndAlso t.StartsWith("[") AndAlso t.EndsWith("]") AndAlso IsNumeric(t.Substring(1, t.Length - 2)) Then
                If cur >= 0 Then mGuide(cur) = sb.ToString.TrimEnd
                cur = CInt(t.Substring(1, t.Length - 2)) - 1
                sb.Clear()
            ElseIf cur >= 0 Then
                sb.AppendLine(ln)
            End If
        Next
        If cur >= 0 Then mGuide(cur) = sb.ToString.TrimEnd
    End Sub

    Private Sub ShowStep(ByVal n As Integer)
        mStep = n
        For i As Integer = 0 To mSteps.Count - 1
            clsUiTheme.StyleButton(mSteps(i), If(i = n, clsUiTheme.ButtonKind.Primary, clsUiTheme.ButtonKind.Normal))
            mSteps(i).TextAlign = ContentAlignment.MiddleLeft
        Next
        Label_StepTitle.Text = STEP_TITLES(n)
        Dim g As String = ""
        If Not mGuide.TryGetValue(n, g) Then g = clsLang.L("설명 파일이 없습니다 : ", "説明ファイルがありません : ") & mKitFolder & "Guide_KO.txt"
        TextBox_Guide.Text = g.Replace(vbCrLf, vbLf).Replace(vbLf, vbCrLf)
        TextBox_Guide.Select(0, 0)

        Panel_Action.Controls.Clear()
        Select Case n
            Case 0 : Call BuildStepCom()
            Case 1 : Call BuildStepMdb()
            Case 2 : Call BuildStepCls()
            Case 3 : Call BuildStepCode()
            Case 4 : Call BuildStepRouter()
            Case 5 : Call BuildStepCheck()
        End Select
    End Sub

    '==========================================================================================
    ' 화면 부품
    '==========================================================================================
    Private Function AddLabel(ByVal text As String, ByVal x As Integer, ByVal y As Integer, Optional ByVal w As Integer = 160) As Label
        Dim l As New Label With {.Text = text, .Location = New Point(x, y + 4), .Size = New Size(w, 22), .ForeColor = clsUiTheme.TextSection}
        Panel_Action.Controls.Add(l)
        Return l
    End Function

    Private Function AddText(ByVal value As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer) As TextBox
        Dim t As New TextBox With {.Text = value, .Location = New Point(x, y), .Size = New Size(w, 26)}
        Panel_Action.Controls.Add(t)
        Return t
    End Function

    Private Function AddButton(ByVal text As String, ByVal x As Integer, ByVal y As Integer, ByVal w As Integer, ByVal kind As clsUiTheme.ButtonKind, ByVal onClick As EventHandler) As Button
        Dim b As New Button With {.Text = text, .Location = New Point(x, y), .Size = New Size(w, 34)}
        clsUiTheme.StyleButton(b, kind)
        AddHandler b.Click, onClick
        Panel_Action.Controls.Add(b)
        Return b
    End Function

    Private Function AddOutput(ByVal y As Integer, ByVal h As Integer) As TextBox
        Dim t As New TextBox With {.Location = New Point(0, y), .Size = New Size(736 - 96, h), .Multiline = True, .ReadOnly = True, .WordWrap = False,
                                   .ScrollBars = ScrollBars.Both, .Font = New Font("Consolas", 10.0!), .BackColor = Color.White}
        Panel_Action.Controls.Add(t)
        AddButton(clsLang.L("⧉ 복사", "⧉ コピー"), 736 - 88, y, 88, clsUiTheme.ButtonKind.Accent,
                  Sub(s, ev)
                      If t.Text <> "" Then Clipboard.SetText(t.Text) : MessageBox.Show(clsLang.L("복사했습니다.", "コピーしました。"), Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                  End Sub)
        Return t
    End Function

    '==========================================================================================
    ' ① 가상 COM : 안의 버튼(설치 / 짝 만들기 / 기본 짝 지우기 / 목록 / 짝 지우기)을 누르면 그 cmd 명령을 보여 준다
    '   명령은 관리자 권한 cmd 에 통째로 붙여 넣는 형태. Router 는 실행하지 않는다
    '==========================================================================================
    Private Const CD_COM0COM As String = "if exist ""C:\Program Files (x86)\com0com\setupc.exe"" (cd /d ""C:\Program Files (x86)\com0com"") else (cd /d ""C:\Program Files\com0com"")"

    Private Sub BuildStepCom()
        Dim ports() As String = IO.Ports.SerialPort.GetPortNames()
        Array.Sort(ports)
        Dim installed As Boolean = File.Exists("C:\Program Files (x86)\com0com\setupc.exe") OrElse File.Exists("C:\Program Files\com0com\setupc.exe")

        Dim subNames() As String = {clsLang.L("1. com0com 설치", "1. com0com インストール"), clsLang.L("2. COM 짝 만들기", "2. COMペア作成"),
                                    clsLang.L("3. 기본 짝 지우기", "3. 既定ペア削除"), clsLang.L("4. 목록 확인", "4. 一覧確認"), clsLang.L("5. 짝 지우기", "5. ペア削除")}
        Dim subs As New List(Of Button)
        For i As Integer = 0 To subNames.Length - 1
            Dim b As New Button With {.Text = subNames(i), .Tag = i, .Location = New Point(i * 148, 0), .Size = New Size(142, 34)}
            subs.Add(b)
            Panel_Action.Controls.Add(b)
        Next

        '1 설치 : 압축 푼 폴더
        Dim lZip As Label = AddLabel(clsLang.L("압축 푼 폴더", "解凍フォルダ"), 0, 42, 110)
        Dim tZip As TextBox = AddText("", 114, 42, 520)
        Dim bZip As New Button With {.Text = clsLang.L("폴더...", "フォルダ..."), .Location = New Point(640, 40), .Size = New Size(96, 30)}
        clsUiTheme.StyleButton(bZip, clsUiTheme.ButtonKind.Normal)
        Panel_Action.Controls.Add(bZip)
        '2 짝 만들기 : COM 번호
        Dim lRouter As Label = AddLabel(clsLang.L("Router 쪽 COM", "Router側COM"), 0, 42, 110)
        Dim tRouter As TextBox = AddText("", 114, 42, 90)
        Dim lInsp As Label = AddLabel(clsLang.L("검사 프로그램 쪽 COM", "検査プログラム側COM"), 230, 42, 160)
        Dim tInsp As TextBox = AddText("", 394, 42, 90)
        '5 짝 지우기 : 번호
        Dim lNo As Label = AddLabel(clsLang.L("짝 번호 (목록의 CNCA 뒤 숫자)", "ペア番号 (一覧のCNCAの後の数字)"), 0, 42, 220)
        Dim tNo As TextBox = AddText("", 224, 42, 60)

        Dim lInfo As Label = AddLabel(clsLang.L("이 PC 의 COM : ", "このPCのCOM : ") & If(ports.Length = 0, "-", String.Join(", ", ports)) & "   /   com0com : " &
                                      If(installed, clsLang.L("설치됨", "インストール済"), clsLang.L("설치 안 됨", "未インストール")), 0, 74, 736)
        lInfo.ForeColor = clsUiTheme.TextSub
        Dim lWarn As Label = AddLabel("", 0, 96, 736)
        lWarn.ForeColor = Color.FromArgb(163, 45, 45)
        Dim tOut As TextBox = AddOutput(124, 168)

        Dim mode As Integer = If(installed, 1, 0)
        Dim make As EventHandler =
            Sub(s, ev)
                For Each c As Control In New Control() {lZip, tZip, bZip}
                    c.Visible = (mode = 0)
                Next
                For Each c As Control In New Control() {lRouter, tRouter, lInsp, tInsp}
                    c.Visible = (mode = 1)
                Next
                lNo.Visible = (mode = 4) : tNo.Visible = (mode = 4)
                For i As Integer = 0 To subs.Count - 1
                    clsUiTheme.StyleButton(subs(i), If(i = mode, clsUiTheme.ButtonKind.Primary, clsUiTheme.ButtonKind.Normal))
                Next

                Dim warn As New List(Of String)
                Dim sb As New System.Text.StringBuilder
                Select Case mode
                    Case 0
                        Dim zip As String = tZip.Text.Trim
                        Dim setupExe As String = "setup.exe"
                        If zip <> "" AndAlso Directory.Exists(zip) Then
                            Dim found() As String = Directory.GetFiles(zip, "setup*.exe")
                            If found.Length > 0 Then
                                setupExe = Path.GetFileName(found(0))
                            Else
                                warn.Add(clsLang.L("그 폴더에 setup*.exe 가 없습니다.", "そのフォルダに setup*.exe がありません。"))
                            End If
                        End If
                        If installed Then warn.Add(clsLang.L("이 PC 에는 이미 설치되어 있습니다 (2 로).", "このPCには既にインストール済みです (2へ)。"))
                        If zip = "" Then warn.Add(clsLang.L("[폴더...] 로 압축 푼 폴더를 고르면 명령의 <...> 가 바뀝니다.", "[フォルダ...] で解凍フォルダを選ぶとコマンドの<...>が置き換わります。"))
                        sb.AppendLine("REM ===== " & clsLang.L("com0com 설치 (PC 마다 처음 한 번) : 관리자 권한 cmd 에 통째로 붙여 넣기", "com0com インストール (PCごとに初回のみ) : 管理者権限のcmdにまとめて貼り付け") & " =====")
                        sb.AppendLine("REM " & clsLang.L("준비 : com0com-2.2.2.0-x64-fre-signed.zip (서명판 64bit) 을 압축 풀기", "準備 : com0com-2.2.2.0-x64-fre-signed.zip (署名版64bit) を解凍"))
                        sb.AppendLine("cd /d """ & If(zip = "", "<" & clsLang.L("압축 푼 폴더", "解凍フォルダ") & ">", zip) & """")
                        sb.AppendLine(setupExe & " /S")
                        sb.AppendLine("REM " & clsLang.L("/S = 화면 없이 설치 (설치 화면으로 하려면 /S 를 지움). 드라이버 설치 확인 창이 뜨면 [설치]", "/S = 画面なしでインストール (画面で行う場合は /S を削除)。ドライバ確認が出たら [インストール]"))
                        sb.AppendLine("ping -n 6 127.0.0.1 >nul")
                        sb.AppendLine(CD_COM0COM)
                        sb.AppendLine("setupc list")
                        sb.AppendLine("REM " & clsLang.L("목록에 CNCA0 / CNCB0 이 보이면 설치 끝 → 3. 기본 짝 지우기 → 2. COM 짝 만들기", "一覧に CNCA0 / CNCB0 があれば完了 → 3. 既定ペア削除 → 2. COMペア作成"))
                    Case 1
                        Dim a As String = NormCom(tRouter.Text), b As String = NormCom(tInsp.Text)
                        If Not installed Then warn.Add(clsLang.L("com0com 이 설치되어 있지 않습니다 (1. 먼저).", "com0comがインストールされていません (1.を先に)。"))
                        For Each c As String In {a, b}
                            If c <> "" AndAlso ports.Contains(c, StringComparer.OrdinalIgnoreCase) Then warn.Add(c & clsLang.L(" 은 이미 이 PC 에 있습니다.", " はこのPCに既にあります。"))
                        Next
                        If a <> "" AndAlso a.Equals(b, StringComparison.OrdinalIgnoreCase) Then warn.Add(clsLang.L("두 번호가 같습니다.", "2つの番号が同じです。"))
                        If a = "" OrElse b = "" Then warn.Add(clsLang.L("두 칸에 COM 번호를 넣으면 명령의 <...> 가 바뀝니다.", "2つにCOM番号を入れるとコマンドの<...>が置き換わります。"))
                        sb.AppendLine("REM ===== " & clsLang.L("가상 COM 짝 만들기 (검사 프로그램 1개당 1번) : 관리자 권한 cmd 에 통째로 붙여 넣기", "仮想COMペア作成 (検査プログラム1つにつき1回) : 管理者権限のcmdにまとめて貼り付け") & " =====")
                        sb.AppendLine(CD_COM0COM)
                        sb.AppendLine("setupc install PortName=" & If(a = "", "<" & clsLang.L("Router쪽COM", "Router側COM") & ">", a) & " PortName=" & If(b = "", "<" & clsLang.L("검사쪽COM", "検査側COM") & ">", b))
                        sb.AppendLine("setupc list")
                        sb.AppendLine("REM " & clsLang.L("Router 쪽 COM → ⑤ 분기 대상 COM / 검사 프로그램 쪽 COM → ② EquipmentSetting.mdb BarcodeReaderPort", "Router側COM → ⑤ 分岐先COM / 検査プログラム側COM → ② EquipmentSetting.mdb BarcodeReaderPort"))
                    Case 2
                        sb.AppendLine("REM ===== " & clsLang.L("com0com 기본 짝 CNCA0 / CNCB0 지우기 (설치 직후 한 번) : 관리자 권한 cmd 에 붙여 넣기", "com0com 既定ペア CNCA0 / CNCB0 削除 (インストール直後に1回) : 管理者権限のcmdに貼り付け") & " =====")
                        sb.AppendLine(CD_COM0COM)
                        sb.AppendLine("setupc list")
                        sb.AppendLine("setupc remove 0")
                        sb.AppendLine("setupc list")
                        sb.AppendLine("REM " & clsLang.L("기본 짝은 0 번 (CNCA0 / CNCB0). 이미 지웠으면 remove 0 은 아무것도 하지 않음", "既定ペアは0番 (CNCA0 / CNCB0)。削除済みなら remove 0 は何もしない"))
                    Case 3
                        sb.AppendLine("REM ===== " & clsLang.L("가상 COM 목록 확인", "仮想COM一覧確認") & " =====")
                        sb.AppendLine(CD_COM0COM)
                        sb.AppendLine("setupc list")
                        sb.AppendLine("REM " & clsLang.L("예) CNCA1 PortName=COM8 / CNCB1 PortName=COM9 → 1 번 짝이 COM8 ⇄ COM9", "例) CNCA1 PortName=COM8 / CNCB1 PortName=COM9 → 1番ペアが COM8 ⇄ COM9"))
                    Case 4
                        Dim n As String = tNo.Text.Trim
                        If n <> "" AndAlso Not IsNumeric(n) Then warn.Add(clsLang.L("번호는 숫자만 (CNCA1 이면 1).", "番号は数字のみ (CNCA1 なら 1)。"))
                        If n = "" Then warn.Add(clsLang.L("4. 목록 확인 에서 지울 짝의 번호를 확인해 넣으세요.", "4. 一覧確認 で削除するペアの番号を確認して入力してください。"))
                        sb.AppendLine("REM ===== " & clsLang.L("가상 COM 짝 지우기 : 관리자 권한 cmd 에 붙여 넣기", "仮想COMペア削除 : 管理者権限のcmdに貼り付け") & " =====")
                        sb.AppendLine(CD_COM0COM)
                        sb.AppendLine("setupc list")
                        sb.AppendLine("setupc remove " & If(n = "", "<" & clsLang.L("번호", "番号") & ">", n))
                        sb.AppendLine("setupc list")
                End Select
                lWarn.Text = String.Join("  ", warn)
                tOut.Text = sb.ToString
            End Sub

        For Each b As Button In subs
            AddHandler b.Click, Sub(s, ev)
                                    mode = CInt(DirectCast(s, Button).Tag)
                                    Call make(Nothing, EventArgs.Empty)
                                End Sub
        Next
        AddHandler bZip.Click, Sub(s, ev)
                                   Using d As New FolderBrowserDialog With {.Description = clsLang.L("com0com zip 을 압축 푼 폴더 (setup.exe 가 있는 곳)", "com0com zip を解凍したフォルダ (setup.exe がある所)")}
                                       If d.ShowDialog(Me) = DialogResult.OK Then tZip.Text = d.SelectedPath
                                   End Using
                               End Sub
        For Each t As TextBox In {tZip, tRouter, tInsp, tNo}
            AddHandler t.TextChanged, make
        Next
        Call make(Nothing, EventArgs.Empty)
    End Sub

    Private Shared Function NormCom(ByVal s As String) As String
        s = s.Trim.ToUpper
        If s = "" Then Return ""
        If IsNumeric(s) Then s = "COM" & s
        Return s
    End Function

    '==========================================================================================
    ' 단계 안의 버튼 (① 과 같은 모양) - 누른 버튼만 강조하고 onSelect 를 부른다
    '==========================================================================================
    Private Function AddSubButtons(ByVal names() As String, ByVal onSelect As Action(Of Integer), ByVal first As Integer) As List(Of Button)
        Dim subs As New List(Of Button)
        Dim w As Integer = Math.Min(142, CInt(736 / names.Length) - 6)
        For i As Integer = 0 To names.Length - 1
            Dim b As New Button With {.Text = names(i), .Tag = i, .Location = New Point(i * (w + 6), 0), .Size = New Size(w, 34)}
            subs.Add(b)
            Panel_Action.Controls.Add(b)
        Next
        Dim pick As Action(Of Integer) =
            Sub(n)
                For i As Integer = 0 To subs.Count - 1
                    clsUiTheme.StyleButton(subs(i), If(i = n, clsUiTheme.ButtonKind.Primary, clsUiTheme.ButtonKind.Normal))
                Next
                onSelect.Invoke(n)
            End Sub
        For Each b As Button In subs
            AddHandler b.Click, Sub(s, ev) pick.Invoke(CInt(DirectCast(s, Button).Tag))
        Next
        pick.Invoke(first)
        Return subs
    End Function

    '==========================================================================================
    ' ② RouterLink.mdb : RouterLink (Yes / No) + Memo 만. 바코드 포트는 원래 설정 EquipmentSetting.mdb
    '==========================================================================================
    Private Sub BuildStepMdb()
        Dim cUse As New CheckBox With {.Text = clsLang.L("RouterLink = Yes (Router 와 연동)", "RouterLink = Yes (Routerと連携)"), .Checked = True, .Location = New Point(0, 0), .Size = New Size(400, 26)}
        Panel_Action.Controls.Add(cUse)
        AddLabel("Memo", 0, 36, 60)
        Dim tMemo As TextBox = AddText("", 64, 36, 670)
        AddButton(clsLang.L("⇩ RouterLink.mdb 저장...", "⇩ RouterLink.mdb 保存..."), 0, 76, 240, clsUiTheme.ButtonKind.Primary,
                  Sub(s, ev) Call SaveMdb(cUse.Checked, tMemo.Text.Trim))
        AddLabel(clsLang.L("저장 위치 : EquipmentSetting.mdb 와 같은 폴더", "保存先 : EquipmentSetting.mdbと同じフォルダ"), 250, 84, 486).ForeColor = clsUiTheme.TextSub
        Dim box As New TextBox With {.Location = New Point(0, 126), .Size = New Size(736, 80), .Multiline = True, .ReadOnly = True, .BackColor = Color.White,
                                     .Font = clsUiTheme.UiFont(10.0!),
                                     .Text = clsLang.L("바코드 포트" & vbCrLf & "  EquipmentSetting.mdb > BarcodeReaderPort    = ① 의 검사 프로그램 쪽 COM" & vbCrLf & "  EquipmentSetting.mdb > BarcodeReaderSetting = ⑤ 의 통신설정과 같게 (예 9600-N-8-1)",
                                                       "バーコードポート" & vbCrLf & "  EquipmentSetting.mdb > BarcodeReaderPort    = ① の検査プログラム側COM" & vbCrLf & "  EquipmentSetting.mdb > BarcodeReaderSetting = ⑤ の通信設定と同じ (例 9600-N-8-1)")}
        Panel_Action.Controls.Add(box)
    End Sub
    Private Sub SaveMdb(ByVal use As Boolean, ByVal memo As String)
        Dim src As String = mKitFolder & clsLinkKit.MDB_NAME
        If Not File.Exists(src) Then
            MessageBox.Show(clsLang.L("양식 파일이 없습니다 : ", "テンプレートファイルがありません : ") & src, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Using d As New SaveFileDialog With {.FileName = clsLinkKit.MDB_NAME, .Filter = "RouterLink.mdb|RouterLink.mdb",
                                             .Title = clsLang.L("검사 프로그램의 EquipmentSetting.mdb 가 있는 폴더에 저장", "検査プログラムのEquipmentSetting.mdbと同じフォルダに保存")}
            If d.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
            Try
                File.Copy(src, d.FileName, True)
                File.SetAttributes(d.FileName, FileAttributes.Normal)
                Call clsLinkKit.WriteMdb(d.FileName, If(use, "Yes", "No"), memo)
                MessageBox.Show(clsLang.L("저장했습니다.", "保存しました。") & vbCrLf & d.FileName, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(clsLang.L("저장하지 못했습니다.", "保存できませんでした。") & vbCrLf & ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    '==========================================================================================
    ' ③ clsRouterLink.vb : 공통판 파일 저장
    '==========================================================================================
    Private Sub BuildStepCls()
        AddButton(clsLang.L("⇩ clsRouterLink.vb 저장...", "⇩ clsRouterLink.vb 保存..."), 0, 0, 240, clsUiTheme.ButtonKind.Primary,
                  Sub(s, ev) Call SaveKitFile(clsLinkKit.CLS_NAME, "VB (*.vb)|*.vb"))
        AddButton(clsLang.L("▤ LinkKit 폴더 열기", "▤ LinkKitフォルダを開く"), 250, 0, 200, clsUiTheme.ButtonKind.Normal,
                  Sub(s, ev) Call OpenFolder(mKitFolder))
    End Sub

    Private Sub SaveKitFile(ByVal name As String, ByVal filter As String)
        Dim src As String = mKitFolder & name
        If Not File.Exists(src) Then
            MessageBox.Show(clsLang.L("파일이 없습니다 : ", "ファイルがありません : ") & src, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        Using d As New SaveFileDialog With {.FileName = name, .Filter = filter}
            If d.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
            Try
                File.Copy(src, d.FileName, True)
                File.SetAttributes(d.FileName, FileAttributes.Normal)
                MessageBox.Show(clsLang.L("저장했습니다.", "保存しました。") & vbCrLf & d.FileName, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                MessageBox.Show(clsLang.L("저장하지 못했습니다.", "保存できませんでした。") & vbCrLf & ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub OpenFolder(ByVal f As String)
        Try
            Process.Start(New ProcessStartInfo(f) With {.UseShellExecute = True})
        Catch ex As Exception
            MessageBox.Show(ex.Message, Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    '==========================================================================================
    ' ④ 코드 추가 : 모든 검사 프로그램 같은 방식 (바코드 COM 수신 → 검사 → END 회신) - SerialPort 이름 → 넣을 1줄
    '==========================================================================================
    Private Sub BuildStepCode()
        AddLabel("SerialPort", 0, 44, 110)
        Dim t1 As TextBox = AddText("Me.SerialPort1", 114, 44, 250)
        AddLabel(clsLang.L("설정 폴더", "設定フォルダ"), 0, 76, 110)
        Dim tFolder As TextBox = AddText("mAppPath & ""Files\""", 114, 76, 250)
        AddLabel(clsLang.L("로그", "ログ"), 380, 76, 100)
        Dim tLog As TextBox = AddText("SysLogWrite(True, m)", 484, 76, 250)
        AddLabel(clsLang.L("주석", "コメント"), 0, 108, 110)
        Dim tCmt As TextBox = AddText("'VerX.XX YSH ADD : Barcode_Router 연동 (clsRouterLink)", 114, 108, 620)
        Dim tOut As TextBox = AddOutput(142, 150)

        Dim make As EventHandler =
            Sub(s, ev)
                Dim logArg As String = If(tLog.Text.Trim = "", "", ", Sub(m) " & tLog.Text.Trim)
                Dim sb As New System.Text.StringBuilder
                sb.AppendLine("' " & clsLang.L("넣는 곳 : SerialPort1_DataReceived 의 끝 - 검사 처리가 끝나고 포트를 다시 연 다음 줄 (OK / NG / 중지 모두 지나가는 곳)", "位置 : SerialPort1_DataReceived の最後 - 検査処理が終わりポートを開き直した次の行 (OK / NG / 中止 すべて通る所)"))
                sb.AppendLine("Call clsRouterLink.SendEnd(" & t1.Text.Trim & ", " & tFolder.Text.Trim & logArg & ")     " & tCmt.Text.Trim)
                tOut.Text = sb.ToString
            End Sub
        For Each t As TextBox In {t1, tFolder, tLog, tCmt}
            AddHandler t.TextChanged, make
        Next
        Call make(Nothing, EventArgs.Empty)
    End Sub
    '==========================================================================================
    ' ⑤ Router 등록 : COM 설정 / 분기 조건 설정 창 열기
    '==========================================================================================
    Private Sub BuildStepRouter()
        AddButton(clsLang.L("⇄  COM 설정 열기", "⇄  COM設定を開く"), 0, 0, 220, clsUiTheme.ButtonKind.Accent, Sub(s, ev) If mOpenCom IsNot Nothing Then mOpenCom.Invoke())
        AddButton(clsLang.L("⑂  분기 조건 설정 열기", "⑂  分岐条件設定を開く"), 230, 0, 220, clsUiTheme.ButtonKind.Accent, Sub(s, ev) If mOpenRule IsNot Nothing Then mOpenRule.Invoke())
    End Sub

    '==========================================================================================
    ' ⑥ 확인 : 확인 항목 버튼 → 무엇을 보면 되는지 / 안 되면 어디를
    '==========================================================================================
    Private Sub BuildStepCheck()
        Dim items(,) As String = {
            {clsLang.L("1. 설정 읽기", "1. 設定読込"),
             clsLang.L("검사 프로그램을 켜고 그 프로그램 로그를 봅니다." & vbCrLf & vbCrLf & "정상 : RouterLink : Yes (value 'Yes', ...\RouterLink.mdb)" & vbCrLf & vbCrLf & "No / not found → ② (파일 위치, 값)",
                       "検査プログラムを起動し、そのログを見ます。" & vbCrLf & vbCrLf & "正常 : RouterLink : Yes (value 'Yes', ...\RouterLink.mdb)" & vbCrLf & vbCrLf & "No / not found → ② (ファイルの場所、値)")},
            {clsLang.L("2. Router 카드", "2. Routerカード"),
             clsLang.L("Router 메인 화면의 분기 대상 카드를 봅니다." & vbCrLf & vbCrLf & "대기 → 정상" & vbCrLf & "COM 실패 → ① / ⑤ 의 COM 번호" & vbCrLf & "프로그램 꺼짐 → ⑤ 의 프로세스명",
                       "Routerメイン画面の分岐先カードを見ます。" & vbCrLf & vbCrLf & "待機 → 正常" & vbCrLf & "COM失敗 → ① / ⑤ のCOM番号" & vbCrLf & "プログラム停止 → ⑤ のプロセス名")},
            {clsLang.L("3. 바코드 찍기", "3. バーコード読取"),
             clsLang.L("바코드를 찍습니다." & vbCrLf & vbCrLf & "정상 : 검사 프로그램 창이 앞으로 오고 검사 시작" & vbCrLf & vbCrLf & "반응 없음 → ② 포트 (짝의 반대쪽 COM), 통신설정, 끝 문자 (CR / LF)",
                       "バーコードを読みます。" & vbCrLf & vbCrLf & "正常 : 検査プログラムが前面に来て検査開始" & vbCrLf & vbCrLf & "反応なし → ② ポート (ペアの反対側COM)、通信設定、終端文字 (CR / LF)")},
            {clsLang.L("4. 검사 끝", "4. 検査終了"),
             clsLang.L("검사를 끝냅니다 (OK / NG 모두)." & vbCrLf & vbCrLf & "정상 : 검사 프로그램 로그 Router END Send OK → Router 가 앞으로, 대기" & vbCrLf & vbCrLf & "계속 검사 중 → ④ 넣은 위치, 포트가 열려 있는지",
                       "検査を終えます (OK / NG とも)。" & vbCrLf & vbCrLf & "正常 : 検査プログラムログ Router END Send OK → Routerが前面、待機" & vbCrLf & vbCrLf & "検査中のまま → ④ の挿入位置、ポートが開いているか")}}
        Dim tInfo As New TextBox With {.Location = New Point(0, 44), .Size = New Size(736, 196), .Multiline = True, .ReadOnly = True,
                                       .BackColor = Color.White, .Font = clsUiTheme.UiFont(10.5!)}
        Panel_Action.Controls.Add(tInfo)
        AddButton(clsLang.L("▤ Router 로그 폴더 열기", "▤ Routerログフォルダを開く"), 0, 252, 240, clsUiTheme.ButtonKind.Normal,
                  Sub(s, ev) Call OpenFolder(clsRouterPath.DataFolder & "System_Log"))
        Dim names(items.GetLength(0) - 1) As String
        For i As Integer = 0 To names.Length - 1
            names(i) = items(i, 0)
        Next
        AddSubButtons(names, Sub(n)
                                 tInfo.Text = items(n, 1)
                                 tInfo.Select(0, 0)
                             End Sub, 0)
    End Sub
End Class

'==========================================================
'Class Name : clsLinkKit
'Function   : 연동 가이드 배포 파일 (Files\LinkKit) - RouterLink.mdb 값 쓰기
'==========================================================
Public NotInheritable Class clsLinkKit

    Public Const MDB_NAME As String = "RouterLink.mdb"
    Public Const CLS_NAME As String = "clsRouterLink.vb"
    Private Shared ReadOnly PROVIDERS() As String = {"Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0"}

    Private Sub New()
    End Sub

    Public Shared Sub WriteMdb(ByVal path As String, ByVal link As String, ByVal memo As String)
        Dim lastErr As Exception = Nothing
        For Each p As String In PROVIDERS
            Try
                Using cn As New OleDbConnection("Provider=" & p & ";Data Source=" & path)
                    cn.Open()
                    Using cmd As New OleDbCommand("UPDATE tblRouterLink SET RouterLink = ?, [Memo] = ?", cn)
                        cmd.Parameters.AddWithValue("@l", link)
                        cmd.Parameters.AddWithValue("@m", memo)
                        cmd.ExecuteNonQuery()
                    End Using
                End Using
                Exit Sub
            Catch ex As InvalidOperationException
                lastErr = ex
            End Try
        Next
        Throw New ApplicationException("OLEDB provider not found (Jet 4.0 / ACE)" & If(lastErr Is Nothing, "", " : " & lastErr.Message))
    End Sub

End Class
