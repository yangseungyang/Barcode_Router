Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmDisplaySetting
'Function   : 화면 설정 (받아오는 정보 항목 / 메인 화면 카드 / 바코드 종류 / 기타) + 설정 내보내기 / 가져오기
'Framwork   : VB2022 Framework4.8
'Memo       : 2026-10-28 지시 - UI 를 유저가 동적으로 설정 (하드코딩 하지 않음)
'             - 항목      (tblField)       : PLASMA 항목명 / 다른 항목에서 잘라 만듦(DERIVE) / 입력 바코드(INPUT). 분기 규칙 대상, 카드에 쓴다
'             - 카드      (tblCard)        : 메인 화면 제품 정보 카드. 항목은 ComboBox 로 고른다
'             - 바코드 종류(tblBarcodeType) : 길이 / 글자 종류 / 패턴 → 처리 (ROUTE 분기 / BROADCAST 모두에 보냄 / REJECT 거부)
'             - 기타      (tblConfig)      : PLASMA 서버, 글꼴, 설명서 경로, 로그 보관 개월, 카드 한 줄 개수
'             - 표의 줄 순서가 번호(순서)가 된다 (▲ / ▼)
'             - 내보내기 : 저장된 RouterSetting.mdb 를 다른 이름으로 복사 / 가져오기 : 고른 mdb 로 바꾼다 (지금 파일은 Files\Backup 에 보관)
'             - 화면은 Designer 없이 코드로 만든다 (표 4개 + 버튼)
'Issue      : 2026/10/28  Rev 0.00    YSH       New Release
'==========================================================

Public Class frmDisplaySetting
    Inherits Form

    Private mSetting As clsRouterSetting

    Private WithEvents Tab_Main As New TabControl With {.Location = New Point(14, 54), .Size = New Size(972, 470)}
    Private WithEvents Grid_Field As New DataGridView
    Private WithEvents Grid_Card As New DataGridView
    Private WithEvents Grid_Type As New DataGridView
    Private WithEvents Grid_Config As New DataGridView
    Private WithEvents Button_Add As New Button With {.Text = clsLang.L("＋ 추가", "＋ 追加"), .Size = New Size(90, 34)}
    Private WithEvents Button_Delete As New Button With {.Text = clsLang.L("✕ 삭제", "✕ 削除"), .Size = New Size(90, 34)}
    Private WithEvents Button_Up As New Button With {.Text = "▲", .Size = New Size(44, 34)}
    Private WithEvents Button_Down As New Button With {.Text = "▼", .Size = New Size(44, 34)}
    Private WithEvents Button_Export As New Button With {.Text = clsLang.L("⇪ 설정 내보내기", "⇪ 設定エクスポート"), .Size = New Size(140, 38)}
    Private WithEvents Button_Import As New Button With {.Text = clsLang.L("⇩ 설정 가져오기", "⇩ 設定インポート"), .Size = New Size(140, 38)}
    Private WithEvents Button_Save As New Button With {.Text = clsLang.L("저장", "保存"), .Size = New Size(100, 38)}
    Private WithEvents Button_Close As New Button With {.Text = clsLang.L("닫기", "閉じる"), .Size = New Size(100, 38)}
    Private Label_Hint As New Label With {.Location = New Point(14, 576), .Size = New Size(972, 20), .ForeColor = clsUiTheme.TextSub}

    'tblConfig 키 설명 (값은 넣지 않는다 - 비어 있는 키도 표에 보이게만 한다)
    Private ReadOnly CONFIG_KEYS As String(,) = {
        {clsRouterSetting.CFG_UI_LANGUAGE, clsLang.L("화면 언어 (KO = 한국어 PC / JA = 일본어 PC). 비우면 한국어", "画面の言語 (KO = 韓国語PC / JA = 日本語PC)。空欄なら韓国語")},
        {clsRouterSetting.CFG_PLASMA_HOST1, clsLang.L("PLASMA 서버 1 (EtherComN.ComHost 첫 번째 주소)", "PLASMAサーバ 1 (EtherComN.ComHost 1番目のアドレス)")},
        {clsRouterSetting.CFG_PLASMA_HOST2, clsLang.L("PLASMA 서버 2 (EtherComN.ComHost 두 번째 주소)", "PLASMAサーバ 2 (EtherComN.ComHost 2番目のアドレス)")},
        {clsRouterSetting.CFG_PLASMA_PORT, clsLang.L("PLASMA 포트", "PLASMAポート")},
        {clsRouterSetting.CFG_PLASMA_COMMAND, clsLang.L("PLASMA 조회 명령 (앞에 붙는 글자)", "PLASMA照会コマンド (先頭に付く文字)")},
        {clsRouterSetting.CFG_UI_FONT, clsLang.L("화면 글꼴 (PC 에 없으면 Windows 기본 글꼴). 일본어 PC 예 : Meiryo UI", "画面フォント (PCにない場合はWindows標準フォント)。空欄なら言語に合わせて自動")},
        {clsRouterSetting.CFG_MANUAL_PATH, clsLang.L("사용자 설명서 파일 (전체 경로 또는 데이터 폴더 기준 상대 경로)", "取扱説明書ファイル (フルパス、またはデータフォルダからの相対パス)")},
        {clsRouterSetting.CFG_LOG_KEEP_MONTHS, clsLang.L("로그 보관 개월 (오래된 월 폴더는 시작할 때 삭제, 비우면 삭제 안 함)", "ログ保管月数 (古い月フォルダは起動時に削除、空欄なら削除しない)")},
        {clsRouterSetting.CFG_CARDS_PER_ROW, clsLang.L("제품 정보 카드 한 줄 개수 (비우면 한 줄에 모두)", "製品情報カードの1行の数 (空欄なら1行に全部)")}}

    Public Sub New(ByVal setting As clsRouterSetting)
        mSetting = setting
        Me.Text = clsLang.L("화면 설정", "画面設定")
        Me.ClientSize = New Size(1000, 610)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False : Me.MinimizeBox = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.Font = clsUiTheme.UiFont(9.0!)

        Dim header As New Panel With {.Dock = DockStyle.Top, .Height = 42, .BackColor = clsUiTheme.HeaderBack}
        header.Controls.Add(New Label With {.Text = clsLang.L("⚙  화면 설정", "⚙  画面設定"), .AutoSize = True, .Location = New Point(12, 9), .ForeColor = clsUiTheme.HeaderText,
                                            .Font = clsUiTheme.UiFont(12.0!, FontStyle.Bold)})
        Me.Controls.Add(header)

        Call AddPage(clsLang.L("항목 (받아오는 정보)", "項目 (取得する情報)"), Grid_Field)
        Call AddPage(clsLang.L("카드 (메인 화면)", "カード (メイン画面)"), Grid_Card)
        Call AddPage(clsLang.L("바코드 종류", "バーコード種類"), Grid_Type)
        Call AddPage(clsLang.L("기타", "その他"), Grid_Config)
        Me.Controls.Add(Tab_Main)

        Button_Add.Location = New Point(14, 532) : Button_Delete.Location = New Point(110, 532)
        Button_Up.Location = New Point(206, 532) : Button_Down.Location = New Point(254, 532)
        Button_Export.Location = New Point(424, 530) : Button_Import.Location = New Point(570, 530)
        Button_Save.Location = New Point(780, 530) : Button_Close.Location = New Point(886, 530)
        Me.Controls.AddRange(New Control() {Button_Add, Button_Delete, Button_Up, Button_Down, Button_Export, Button_Import, Button_Save, Button_Close, Label_Hint})
    End Sub

    Private Sub AddPage(ByVal title As String, ByVal g As DataGridView)
        Dim pg As New TabPage(title) With {.BackColor = Color.White, .Padding = New Padding(6)}
        g.Dock = DockStyle.Fill
        g.AllowUserToAddRows = False
        g.AllowUserToDeleteRows = False
        g.MultiSelect = False
        g.SelectionMode = DataGridViewSelectionMode.CellSelect
        pg.Controls.Add(g)
        Tab_Main.TabPages.Add(pg)
    End Sub

    Private Shared Function TextCol(ByVal name As String, ByVal header As String, ByVal width As Integer) As DataGridViewTextBoxColumn
        Return New DataGridViewTextBoxColumn With {.Name = name, .HeaderText = header, .Width = width}
    End Function

    Private Shared Function ComboCol(ByVal name As String, ByVal header As String, ByVal width As Integer, ByVal items() As String) As DataGridViewComboBoxColumn
        Dim c As New DataGridViewComboBoxColumn With {.Name = name, .HeaderText = header, .Width = width, .FlatStyle = FlatStyle.Flat}
        c.Items.AddRange(items)
        Return c
    End Function

    Private Shared Function CheckCol(ByVal name As String, ByVal header As String) As DataGridViewCheckBoxColumn
        Return New DataGridViewCheckBoxColumn With {.Name = name, .HeaderText = header, .Width = 50}
    End Function

    Private Sub frmDisplaySetting_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        For Each b As Button In {Button_Add, Button_Up, Button_Down, Button_Close}
            clsUiTheme.StyleButton(b, clsUiTheme.ButtonKind.Normal)
        Next
        clsUiTheme.StyleButton(Button_Delete, clsUiTheme.ButtonKind.Danger)
        clsUiTheme.StyleButton(Button_Export, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Import, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Save, clsUiTheme.ButtonKind.Primary)
        For Each g As DataGridView In {Grid_Field, Grid_Card, Grid_Type, Grid_Config}
            clsUiTheme.StyleGrid(g)
        Next

        '항목 ---------------------------------------------------------------------------------
        Grid_Field.Columns.AddRange(New DataGridViewColumn() {
            TextCol("FieldID", clsLang.L("항목ID", "項目ID"), 100), TextCol("DisplayName", clsLang.L("표시명", "表示名"), 120),
            ComboCol("SourceType", clsLang.L("종류", "種類"), 100, {clsField.SRC_PLASMA, clsField.SRC_DERIVE, clsField.SRC_INPUT}),
            TextCol("PlasmaTag", clsLang.L("PLASMA 항목명", "PLASMA項目名"), 170), TextCol("BaseField", clsLang.L("원 항목", "元項目"), 90), TextCol("Delimiter", clsLang.L("구분자", "区切り"), 60),
            TextCol("PartNo", clsLang.L("몇 번째", "何番目"), 60), TextCol("StartPos", clsLang.L("시작", "開始"), 50), TextCol("PartLen", clsLang.L("길이", "長さ"), 50), CheckCol("Enabled", clsLang.L("사용", "使用"))})
        Grid_Field.Columns("PlasmaTag").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        For Each f As clsField In mSetting.Fields.OrderBy(Function(x) x.SortNo)
            Grid_Field.Rows.Add(f.FieldID, f.DisplayName, If(IsIn(f.SourceType, clsField.SRC_PLASMA, clsField.SRC_DERIVE, clsField.SRC_INPUT), f.SourceType, Nothing),
                                f.PlasmaTag, f.BaseField, f.Delimiter, IntText(f.PartNo), IntText(f.StartPos), IntText(f.PartLen), f.Enabled)
        Next

        '카드 ---------------------------------------------------------------------------------
        Grid_Card.Columns.AddRange(New DataGridViewColumn() {
            ComboCol("FieldID", clsLang.L("항목", "項目"), 160, FieldIDs()), TextCol("Caption", clsLang.L("표시명 (비우면 항목 표시명)", "表示名 (空欄なら項目の表示名)"), 260), CheckCol("Enabled", clsLang.L("사용", "使用"))})
        Grid_Card.Columns("Caption").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        For Each c As clsCard In mSetting.Cards.OrderBy(Function(x) x.CardNo)
            Call EnsureComboItem(Grid_Card, "FieldID", c.FieldID)
            Grid_Card.Rows.Add(If(c.FieldID = "", Nothing, c.FieldID), c.Caption, c.Enabled)
        Next

        '바코드 종류 ---------------------------------------------------------------------------
        Grid_Type.Columns.AddRange(New DataGridViewColumn() {
            TextCol("TypeName", clsLang.L("종류 이름", "種類名"), 160), TextCol("MinLen", clsLang.L("최소 길이", "最小長さ"), 90), TextCol("MaxLen", clsLang.L("최대 길이", "最大長さ"), 90),
            ComboCol("CharKind", clsLang.L("글자 종류", "文字種類"), 90, {clsBarcodeType.KIND_ANY, clsBarcodeType.KIND_NUM, clsBarcodeType.KIND_ALPHA, clsBarcodeType.KIND_ALNUM}),
            TextCol("Pattern", clsLang.L("패턴 (VB Like, 예 : A*)", "パターン (VB Like、例 : A*)"), 160),
            ComboCol("ActionType", clsLang.L("처리", "処理"), 130, {clsBarcodeType.ACT_ROUTE, clsBarcodeType.ACT_BROADCAST, clsBarcodeType.ACT_REJECT}), CheckCol("Enabled", clsLang.L("사용", "使用"))})
        Grid_Type.Columns("Pattern").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        For Each t As clsBarcodeType In mSetting.BarcodeTypes.OrderBy(Function(x) x.TypeNo)
            Grid_Type.Rows.Add(t.TypeName, IntText(t.MinLen), IntText(t.MaxLen),
                               If(IsIn(t.CharKind, clsBarcodeType.KIND_ANY, clsBarcodeType.KIND_NUM, clsBarcodeType.KIND_ALPHA, clsBarcodeType.KIND_ALNUM), t.CharKind, Nothing),
                               t.Pattern, If(IsIn(t.ActionType, clsBarcodeType.ACT_ROUTE, clsBarcodeType.ACT_BROADCAST, clsBarcodeType.ACT_REJECT), t.ActionType, Nothing), t.Enabled)
        Next

        '기타 ---------------------------------------------------------------------------------
        Grid_Config.Columns.AddRange(New DataGridViewColumn() {TextCol("ConfigKey", clsLang.L("키", "キー"), 160), TextCol("ConfigValue", clsLang.L("값", "値"), 260), TextCol("Description", clsLang.L("설명", "説明"), 400)})
        Grid_Config.Columns("Description").AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill
        Dim shown As New List(Of String)
        For Each k As String In mSetting.Config.Keys
            Dim desc As String = ""
            mSetting.ConfigDesc.TryGetValue(k, desc)
            For i As Integer = 0 To CONFIG_KEYS.GetLength(0) - 1      '알려진 키는 화면 언어에 맞는 설명으로 보인다 (저장된 설명은 다른 언어일 수 있음)
                If String.Equals(CONFIG_KEYS(i, 0), k, StringComparison.OrdinalIgnoreCase) Then desc = CONFIG_KEYS(i, 1)
            Next
            Grid_Config.Rows.Add(k, mSetting.Config(k), If(desc, ""))
            shown.Add(k.ToUpperInvariant)
        Next
        For i As Integer = 0 To CONFIG_KEYS.GetLength(0) - 1
            If Not shown.Contains(CONFIG_KEYS(i, 0).ToUpperInvariant) Then Grid_Config.Rows.Add(CONFIG_KEYS(i, 0), "", CONFIG_KEYS(i, 1))     '키만 보이고 값은 비움
        Next

        '화면 언어는 KO / JA 중에서 고른다 (2026-10-28)
        For Each r As DataGridViewRow In Grid_Config.Rows
            If Not String.Equals(Cell(r, "ConfigKey"), clsRouterSetting.CFG_UI_LANGUAGE, StringComparison.OrdinalIgnoreCase) Then Continue For
            Dim cur As String = Cell(r, "ConfigValue").ToUpperInvariant
            Dim cb As New DataGridViewComboBoxCell With {.FlatStyle = FlatStyle.Flat}
            cb.Items.AddRange(New Object() {"", clsLang.LANG_KO, clsLang.LANG_JA})
            r.Cells("ConfigValue") = cb
            cb.Value = If(cur = clsLang.LANG_KO OrElse cur = clsLang.LANG_JA, cur, "")
        Next

        Call UpdateHint()
        clsUiTheme.ApplyFont(Me)
    End Sub

    Private Shared Function IsIn(ByVal v As String, ByVal ParamArray items() As String) As Boolean
        Return items.Contains(v)
    End Function

    Private Shared Function IntText(ByVal n As Integer) As String
        Return If(n = 0, "", n.ToString)
    End Function

    Private Function FieldIDs() As String()
        Dim lst As New List(Of String)
        For Each r As DataGridViewRow In Grid_Field.Rows
            Dim id As String = Cell(r, "FieldID")
            If id <> "" AndAlso Not lst.Contains(id) Then lst.Add(id)
        Next
        Return lst.ToArray
    End Function

    Private Shared Sub EnsureComboItem(ByVal g As DataGridView, ByVal col As String, ByVal value As String)
        Dim c As DataGridViewComboBoxColumn = DirectCast(g.Columns(col), DataGridViewComboBoxColumn)
        If value <> "" AndAlso Not c.Items.Contains(value) Then c.Items.Add(value)
    End Sub

    Private Shared Function Cell(ByVal r As DataGridViewRow, ByVal col As String) As String
        Dim v As Object = r.Cells(col).Value
        If v Is Nothing Then Return ""
        Return v.ToString.Trim
    End Function

    Private Shared Function CellInt(ByVal r As DataGridViewRow, ByVal col As String) As Integer
        Dim n As Integer = 0
        Integer.TryParse(Cell(r, col), n)
        Return n
    End Function

    Private Shared Function CellBool(ByVal r As DataGridViewRow, ByVal col As String) As Boolean
        Return CBool(If(r.Cells(col).Value, False))
    End Function

    Private ReadOnly Property CurGrid() As DataGridView
        Get
            Return DirectCast(Tab_Main.SelectedTab.Controls(0), DataGridView)
        End Get
    End Property

    Private Sub Tab_Main_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles Tab_Main.SelectedIndexChanged
        '카드 탭으로 오면 항목 목록을 다시 채운다 (항목 탭에서 새로 넣은 항목도 고를 수 있게)
        If CurGrid Is Grid_Card Then
            For Each id As String In FieldIDs()
                Call EnsureComboItem(Grid_Card, "FieldID", id)
            Next
        End If
        Call UpdateHint()
    End Sub

    Private Sub UpdateHint()
        Dim g As DataGridView = CurGrid
        If g Is Grid_Field Then
            Label_Hint.Text = clsLang.L("PLASMA = PLASMA 항목명 뒤 값 / DERIVE = 원 항목을 구분자로 나눈 몇 번째 (구분자 없으면 시작·길이) / INPUT = 찍은 바코드", "PLASMA = PLASMA項目名の後ろの値 / DERIVE = 元項目を区切りで分けた何番目 (区切りなしは開始·長さ) / INPUT = 読んだバーコード")
        ElseIf g Is Grid_Card Then
            Label_Hint.Text = clsLang.L("위에서부터 메인 화면 카드 순서. 한 줄 개수는 기타 > CardsPerRow", "上からメイン画面のカード順。1行の数は その他 > CardsPerRow")
        ElseIf g Is Grid_Type Then
            Label_Hint.Text = clsLang.L("위에서부터 보고 처음 맞는 종류로 처리. 길이 0 = 제한 없음. ROUTE = PLASMA 조회 후 분기 / BROADCAST = 모든 대상 / REJECT = 보내지 않음", "上から見て最初に合う種類で処理。長さ0 = 制限なし。ROUTE = PLASMA照会後に分岐 / BROADCAST = 全分岐先 / REJECT = 送らない")
        Else
            Label_Hint.Text = clsLang.L("값을 비우면 그 기능은 동작하지 않습니다 (프로그램에 기본값 없음)", "値が空欄の機能は動作しません (プログラムに初期値なし)")
        End If
    End Sub

    '줄 추가 / 삭제 / 순서 =================================================================================================
    Private Sub Button_Add_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Add.Click
        Dim g As DataGridView = CurGrid
        Dim i As Integer
        If g Is Grid_Field Then
            i = g.Rows.Add("", "", Nothing, "", "", "", "", "", "", True)
        ElseIf g Is Grid_Card Then
            i = g.Rows.Add(Nothing, "", True)
        ElseIf g Is Grid_Type Then
            i = g.Rows.Add("", "", "", Nothing, "", Nothing, True)
        Else
            i = g.Rows.Add("", "", "")
        End If
        g.CurrentCell = g.Rows(i).Cells(0)
        g.BeginEdit(True)
    End Sub

    Private Sub Button_Delete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Delete.Click
        Dim g As DataGridView = CurGrid
        If g.CurrentRow Is Nothing Then Exit Sub
        If g Is Grid_Field Then
            '분기 규칙 / 카드에서 쓰는 항목은 지우지 않는다
            Dim id As String = Cell(g.CurrentRow, "FieldID")
            Dim users As New List(Of String)
            For Each ru As clsRouteRule In mSetting.Rules
                If ru.Conds.Any(Function(c) String.Equals(c.Target, id, StringComparison.OrdinalIgnoreCase)) Then users.Add(clsLang.L("규칙 '", "ルール '") & ru.Description & "'")
            Next
            For Each r As DataGridViewRow In Grid_Card.Rows
                If String.Equals(Cell(r, "FieldID"), id, StringComparison.OrdinalIgnoreCase) Then users.Add(clsLang.L("카드", "カード")) : Exit For
            Next
            If id <> "" AndAlso users.Count > 0 Then
                MessageBox.Show("'" & id & clsLang.L("' 는 ", "' は ") & String.Join(", ", users) & clsLang.L(" 에서 쓰고 있어서 지울 수 없습니다.", " で使用中のため削除できません。") & vbCrLf & clsLang.L("사용을 끄거나, 먼저 그쪽을 바꿔 주세요.", "使用をオフにするか、先にそちらを変更してください。"), clsLang.L("화면 설정", "画面設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        End If
        g.Rows.Remove(g.CurrentRow)
    End Sub

    Private Sub Button_Up_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Up.Click
        Call MoveRow(-1)
    End Sub

    Private Sub Button_Down_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Down.Click
        Call MoveRow(1)
    End Sub

    Private Sub MoveRow(ByVal dir As Integer)
        Dim g As DataGridView = CurGrid
        If g.CurrentRow Is Nothing Then Exit Sub
        g.EndEdit()
        Dim i As Integer = g.CurrentRow.Index
        Dim j As Integer = i + dir
        If j < 0 OrElse j >= g.Rows.Count Then Exit Sub
        Dim col As Integer = g.CurrentCell.ColumnIndex
        Dim r As DataGridViewRow = g.Rows(i)
        g.Rows.RemoveAt(i)
        g.Rows.Insert(j, r)
        g.CurrentCell = g.Rows(j).Cells(col)
    End Sub

    Private Sub Grid_DataError(ByVal sender As Object, ByVal e As DataGridViewDataErrorEventArgs) Handles Grid_Field.DataError, Grid_Card.DataError, Grid_Type.DataError, Grid_Config.DataError
        e.ThrowException = False        'ComboBox 열에 목록 밖 값 - 저장할 때 다시 확인한다
    End Sub

    '저장 ===================================================================================================================
    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        For Each g As DataGridView In {Grid_Field, Grid_Card, Grid_Type, Grid_Config}
            g.EndEdit()
        Next

        '항목
        Dim fields As New List(Of clsField)
        Dim ids As New List(Of String)
        For Each r As DataGridViewRow In Grid_Field.Rows
            Dim f As New clsField With {.FieldID = Cell(r, "FieldID"), .DisplayName = Cell(r, "DisplayName"), .SourceType = Cell(r, "SourceType"),
                                        .PlasmaTag = Cell(r, "PlasmaTag"), .BaseField = Cell(r, "BaseField"), .Delimiter = If(r.Cells("Delimiter").Value, "").ToString,
                                        .PartNo = CellInt(r, "PartNo"), .StartPos = CellInt(r, "StartPos"), .PartLen = CellInt(r, "PartLen"),
                                        .Enabled = CellBool(r, "Enabled"), .SortNo = r.Index + 1}
            If f.FieldID = "" OrElse f.SourceType = "" Then Return_Warn(clsLang.L("항목 ", "項目 ") & (r.Index + 1) & clsLang.L(" 줄 : 항목ID 와 종류를 넣어 주세요.", " 行 : 項目IDと種類を入力してください。"), Grid_Field) : Exit Sub
            If ids.Contains(f.FieldID.ToUpperInvariant) Then Return_Warn(clsLang.L("항목ID '", "項目ID '") & f.FieldID & clsLang.L("' 가 두 번 있습니다.", "' が2回あります。"), Grid_Field) : Exit Sub
            If f.SourceType = clsField.SRC_PLASMA AndAlso f.PlasmaTag = "" Then Return_Warn("'" & f.FieldID & clsLang.L("' : PLASMA 항목명을 넣어 주세요 (예 : $MS_CODE).", "' : PLASMA項目名を入力してください (例 : $MS_CODE)。"), Grid_Field) : Exit Sub
            If f.SourceType = clsField.SRC_DERIVE AndAlso f.BaseField = "" Then Return_Warn("'" & f.FieldID & clsLang.L("' : 원 항목을 넣어 주세요.", "' : 元項目を入力してください。"), Grid_Field) : Exit Sub
            ids.Add(f.FieldID.ToUpperInvariant)
            fields.Add(f)
        Next
        For Each f As clsField In fields.Where(Function(x) x.SourceType = clsField.SRC_DERIVE)
            If Not ids.Contains(f.BaseField.ToUpperInvariant) Then Return_Warn("'" & f.FieldID & clsLang.L("' 의 원 항목 '", "' の元項目 '") & f.BaseField & clsLang.L("' 가 항목에 없습니다.", "' が項目にありません。"), Grid_Field) : Exit Sub
        Next

        '카드
        Dim cards As New List(Of clsCard)
        For Each r As DataGridViewRow In Grid_Card.Rows
            Dim c As New clsCard With {.CardNo = r.Index + 1, .FieldID = Cell(r, "FieldID"), .Caption = Cell(r, "Caption"), .Enabled = CellBool(r, "Enabled")}
            If c.FieldID = "" Then Return_Warn(clsLang.L("카드 ", "カード ") & c.CardNo & clsLang.L(" 줄 : 항목을 골라 주세요.", " 行 : 項目を選んでください。"), Grid_Card) : Exit Sub
            If Not ids.Contains(c.FieldID.ToUpperInvariant) Then Return_Warn(clsLang.L("카드 ", "カード ") & c.CardNo & clsLang.L(" 줄 : 항목 '", " 行 : 項目 '") & c.FieldID & clsLang.L("' 가 항목 탭에 없습니다.", "' が項目タブにありません。"), Grid_Card) : Exit Sub
            cards.Add(c)
        Next

        '바코드 종류
        Dim types As New List(Of clsBarcodeType)
        For Each r As DataGridViewRow In Grid_Type.Rows
            Dim t As New clsBarcodeType With {.TypeNo = r.Index + 1, .TypeName = Cell(r, "TypeName"), .MinLen = CellInt(r, "MinLen"), .MaxLen = CellInt(r, "MaxLen"),
                                              .CharKind = Cell(r, "CharKind"), .Pattern = Cell(r, "Pattern"), .ActionType = Cell(r, "ActionType"), .Enabled = CellBool(r, "Enabled")}
            If t.TypeName = "" OrElse t.CharKind = "" OrElse t.ActionType = "" Then Return_Warn(clsLang.L("바코드 종류 ", "バーコード種類 ") & t.TypeNo & clsLang.L(" 줄 : 종류 이름, 글자 종류, 처리를 넣어 주세요.", " 行 : 種類名、文字種類、処理を入力してください。"), Grid_Type) : Exit Sub
            If t.MaxLen > 0 AndAlso t.MinLen > t.MaxLen Then Return_Warn("'" & t.TypeName & clsLang.L("' : 최소 길이가 최대 길이보다 큽니다.", "' : 最小長さが最大長さより大きいです。"), Grid_Type) : Exit Sub
            types.Add(t)
        Next

        '기타
        Dim cfg As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        Dim cfgDesc As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        For Each r As DataGridViewRow In Grid_Config.Rows
            Dim k As String = Cell(r, "ConfigKey")
            If k = "" Then Continue For
            If cfg.ContainsKey(k) Then Return_Warn(clsLang.L("키 '", "キー '") & k & clsLang.L("' 가 두 번 있습니다.", "' が2回あります。"), Grid_Config) : Exit Sub
            cfg(k) = Cell(r, "ConfigValue")
            cfgDesc(k) = Cell(r, "Description")
        Next

        Try
            mSetting.Fields = fields
            mSetting.Cards = cards
            mSetting.BarcodeTypes = types
            mSetting.Config = cfg
            mSetting.ConfigDesc = cfgDesc
            mSetting.SaveDisplay()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(clsLang.L("저장 NG", "保存 NG") & vbCrLf & ex.Message, clsLang.L("화면 설정", "画面設定"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub Return_Warn(ByVal msg As String, ByVal g As DataGridView)
        For Each pg As TabPage In Tab_Main.TabPages
            If pg.Controls.Contains(g) Then Tab_Main.SelectedTab = pg
        Next
        MessageBox.Show(msg, clsLang.L("화면 설정", "画面設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
    End Sub

    Private Sub Button_Close_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Close.Click
        Me.Close()
    End Sub

    '설정 내보내기 / 가져오기 ================================================================================================
    Private Sub Button_Export_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Export.Click
        '==========================================================================================
        'Function : 저장된 설정 파일(RouterSetting.mdb)을 다른 곳에 복사 (다른 설비로 옮길 때)
        'Remarks  : 이 화면에서 고치고 저장하지 않은 내용은 들어가지 않는다
        '==========================================================================================
        Dim name As String = "RouterSetting_" & If(mSetting.EquipmentName.Trim = "", "Router", MakeFileName(mSetting.EquipmentName.Trim)) & "_" & Format(DateTime.Now, "yyyyMMdd") & ".mdb"
        Using dlg As New SaveFileDialog With {.Filter = clsLang.L("설정 파일 (*.mdb)|*.mdb", "設定ファイル (*.mdb)|*.mdb"), .FileName = name, .Title = clsLang.L("설정 내보내기 (저장된 설정)", "設定エクスポート (保存済みの設定)")}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
            Try
                IO.File.Copy(mSetting.DbPath, dlg.FileName, True)
                MessageBox.Show(clsLang.L("내보냈습니다.", "エクスポートしました。") & vbCrLf & dlg.FileName, clsLang.L("설정 내보내기", "設定エクスポート"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                clsRouterLog.Write(IO.Path.GetDirectoryName(IO.Path.GetDirectoryName(mSetting.DbPath)) & "\", "Setting Export : " & dlg.FileName)
            Catch ex As Exception
                MessageBox.Show(clsLang.L("내보내기 NG", "エクスポート NG") & vbCrLf & ex.Message, clsLang.L("설정 내보내기", "設定エクスポート"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Sub Button_Import_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Import.Click
        '==========================================================================================
        'Function : 고른 설정 파일로 바꾼다. 지금 파일은 Files\Backup\RouterSetting_yyyyMMdd_HHmmss.mdb 로 보관
        '==========================================================================================
        Using dlg As New OpenFileDialog With {.Filter = clsLang.L("설정 파일 (*.mdb)|*.mdb", "設定ファイル (*.mdb)|*.mdb"), .Title = clsLang.L("설정 가져오기", "設定インポート")}
            If dlg.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
            If String.Equals(IO.Path.GetFullPath(dlg.FileName), IO.Path.GetFullPath(mSetting.DbPath), StringComparison.OrdinalIgnoreCase) Then Exit Sub

            '먼저 읽어 본다 (Router 설정 파일인지)
            Dim test As New clsRouterSetting(dlg.FileName)
            Try
                test.Load()
            Catch ex As Exception
                MessageBox.Show(clsLang.L("Router 설정 파일로 읽을 수 없습니다.", "Router設定ファイルとして読めません。") & vbCrLf & ex.Message, clsLang.L("설정 가져오기", "設定インポート"), MessageBoxButtons.OK, MessageBoxIcon.Error)
                Exit Sub
            End Try
            If MessageBox.Show(clsLang.L("지금 설정을 이 파일로 바꿉니까?", "現在の設定をこのファイルに置き換えますか?") & vbCrLf & dlg.FileName & vbCrLf & vbCrLf &
                               clsLang.L("설비명 : ", "設備名 : ") & test.EquipmentName & clsLang.L(" / 분기 대상 ", " / 分岐先 ") & test.Destinations.Count & clsLang.L(" / 규칙 ", " / ルール ") & test.Rules.Count & vbCrLf &
                               clsLang.L("(지금 파일은 Files\Backup 에 보관합니다)", "(現在のファイルは Files\Backup に保管します)"), clsLang.L("설정 가져오기", "設定インポート"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub

            Try
                Dim backupDir As String = IO.Path.Combine(IO.Path.GetDirectoryName(mSetting.DbPath), "Backup")
                If Not IO.Directory.Exists(backupDir) Then IO.Directory.CreateDirectory(backupDir)
                Dim backup As String = IO.Path.Combine(backupDir, "RouterSetting_" & Format(DateTime.Now, "yyyyMMdd_HHmmss") & ".mdb")
                IO.File.Copy(mSetting.DbPath, backup, False)
                IO.File.Copy(dlg.FileName, mSetting.DbPath, True)
                clsRouterLog.Write(IO.Path.GetDirectoryName(IO.Path.GetDirectoryName(mSetting.DbPath)) & "\", "Setting Import : " & dlg.FileName & " (backup " & backup & ")")
                MessageBox.Show(clsLang.L("가져왔습니다. 설정을 다시 읽습니다.", "インポートしました。設定を読み直します。") & vbCrLf & clsLang.L("보관 : ", "保管 : ") & backup, clsLang.L("설정 가져오기", "設定インポート"), MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.DialogResult = DialogResult.OK
                Me.Close()
            Catch ex As Exception
                MessageBox.Show(clsLang.L("가져오기 NG", "インポート NG") & vbCrLf & ex.Message, clsLang.L("설정 가져오기", "設定インポート"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    Private Shared Function MakeFileName(ByVal s As String) As String
        For Each c As Char In IO.Path.GetInvalidFileNameChars()
            s = s.Replace(c, "_"c)
        Next
        Return s.Replace(" "c, "_"c)
    End Function

End Class
