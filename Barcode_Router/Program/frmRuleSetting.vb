Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmRuleSetting
'Function   : 분기 조건 설정 화면
'Framwork   : VB2022 Framework4.8
'Memo       : 지시 (2026-10-26) : 조건은 유저가 설정한 것을 ComboBox 로, 그것과 연결되는 변수(입력값)까지 설정,
'                                 대상이 여러 개면 Enter 키로 분기 (직접 작성)
'             - ComboBox_Rule  : 저장된 규칙을 고르면 아래 칸에 연결된 값(사용, 분기 COM, 설명, 조건 줄, 입력값)이 나온다
'             - 분기 COM       : 등록된 분기 대상 ComboBox + [신규] [수정] [삭제] (2026-10-26 추가, frmDestEdit, 바로 저장)
'             - 조건 줄        : 연결 ComboBox (AND / OR, 첫 줄 없음, 2026-10-26 추가 - AND 를 먼저 묶음),
'                                대상 ComboBox (화면 설정 > 항목 의 사용 항목, 2026-10-28), 조건 ComboBox (LIKE / EQUAL)
'             - 입력값         : 선택한 조건 줄의 값. 여러 줄 TextBox 에 Enter 로 줄을 나눠 직접 입력 (한 줄 = 값 하나, OR)
'             - 우선순위       : 목록 순서 (▲ / ▼ 로 바꿈). 저장할 때 1 부터 다시 매긴다
'             - 판단 시험      : "항목ID=값; ..." 을 넣으면 화면 내용(저장 전) 기준으로 어떤 규칙에 맞는지 보여 준다
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-3항 / 구조설계 U-01 ~ U-08
'==========================================================

Public Class frmRuleSetting

    Private mSetting As clsRouterSetting
    Private mRules As New List(Of clsRouteRule)     '화면에서 고치는 사본 (저장할 때만 설정에 반영)
    Private mLoading As Boolean = False
    Private mDefaultTarget As String = ""            '새 조건 줄의 대상 (항목 목록의 첫 번째)

    Public Sub New(ByVal setting As clsRouterSetting)
        InitializeComponent()
        mSetting = setting
    End Sub

    Private Sub frmRuleSetting_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Call ApplyLanguageTexts()       '화면 글자를 언어 설정(clsLang)에 맞게 (2026-10-28)
        '화면 모양 (Router_UI_시안_v1.html)
        clsUiTheme.StyleButton(Button_New, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_Up, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Down, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Delete, clsUiTheme.ButtonKind.Danger)
        clsUiTheme.StyleButton(Button_CondAdd, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_CondDelete, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Test, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_Save, clsUiTheme.ButtonKind.Primary)
        clsUiTheme.StyleButton(Button_Close, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleGrid(DataGridView_Cond)
        clsUiTheme.StyleButton(Button_DestNew, clsUiTheme.ButtonKind.Accent)
        clsUiTheme.StyleButton(Button_DestEdit, clsUiTheme.ButtonKind.Normal)
        clsUiTheme.StyleButton(Button_DestDelete, clsUiTheme.ButtonKind.Danger)

        colJoin.Items.AddRange(New Object() {clsRuleCond.JOIN_AND, clsRuleCond.JOIN_OR})
        colOperator.Items.AddRange(New Object() {clsRuleCond.OP_LIKE, clsRuleCond.OP_EQUAL})

        '조건 대상 = 화면 설정 > 항목 (tblField) 의 사용 항목 (2026-10-28 지시 : 받아오는 정보를 분기 규칙에서 모두 사용)
        '저장된 규칙이 쓰는 항목은 지금 사용 안 함이어도 목록에 넣는다 (기존 규칙이 화면에서 깨지지 않게)
        Dim targets As New List(Of String)
        For Each f As clsField In mSetting.EnabledFields()
            If Not targets.Contains(f.FieldID) Then targets.Add(f.FieldID)
        Next
        For Each ru As clsRouteRule In mSetting.Rules
            For Each c As clsRuleCond In ru.Conds
                If c.Target <> "" AndAlso Not targets.Contains(c.Target) Then targets.Add(c.Target)
            Next
        Next
        colTarget.Items.AddRange(targets.ToArray)
        mDefaultTarget = If(targets.Count > 0, targets(0), "")
        If targets.Count = 0 Then Label_TestResult.Text = clsLang.L("  화면 설정 > 항목 이 없습니다 (조건 대상을 고를 수 없음)", "  画面設定 > 項目 がありません (条件の対象を選べません)")

        Call RefreshDestCombo(-1)

        For Each ru As clsRouteRule In mSetting.Rules.OrderBy(Function(x) x.Priority).ThenBy(Function(x) x.RuleID)
            Dim c As New clsRouteRule With {.RuleID = ru.RuleID, .Priority = ru.Priority, .Enabled = ru.Enabled, .DestID = ru.DestID, .Description = ru.Description}
            For Each rc As clsRuleCond In ru.Conds
                c.Conds.Add(New clsRuleCond With {.JoinOp = rc.JoinOp, .Target = rc.Target, .Operator = rc.Operator, .InputValues = rc.InputValues})
            Next
            mRules.Add(c)
        Next

        Call Renumber()
        Call RefreshRuleList(0)
        clsUiTheme.ApplyFont(Me)
    End Sub

    '규칙 목록 ==============================================================================================================
    Private ReadOnly Property CurRule() As clsRouteRule
        Get
            Dim i As Integer = ComboBox_Rule.SelectedIndex
            If i < 0 OrElse i >= mRules.Count Then Return Nothing
            Return mRules(i)
        End Get
    End Property

    Private Sub Renumber()
        For i As Integer = 0 To mRules.Count - 1
            mRules(i).Priority = i + 1
        Next
    End Sub

    Private Sub RefreshRuleList(ByVal selectIndex As Integer)
        mLoading = True
        ComboBox_Rule.Items.Clear()
        For Each ru As clsRouteRule In mRules
            ComboBox_Rule.Items.Add(ru.ToString)
        Next
        mLoading = False

        If mRules.Count = 0 Then
            Call ShowRule(Nothing)
        Else
            ComboBox_Rule.SelectedIndex = Math.Max(0, Math.Min(selectIndex, mRules.Count - 1))
        End If
    End Sub

    Private Sub ComboBox_Rule_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox_Rule.SelectedIndexChanged
        If mLoading Then Exit Sub
        Call ShowRule(CurRule)
    End Sub

    Private Sub ShowRule(ByVal ru As clsRouteRule)
        '==========================================================================================
        'Function : 고른 규칙에 연결된 값을 아래 칸에 표시
        '==========================================================================================
        mLoading = True
        Panel_Rule.Enabled = (ru IsNot Nothing)
        DataGridView_Cond.Rows.Clear()
        TextBox_Values.Text = ""

        If ru Is Nothing Then
            Label_Priority.Text = clsLang.L("우선순위 -", "優先順位 -")
            CheckBox_Enabled.Checked = False
            ComboBox_Dest.SelectedIndex = -1
            TextBox_Desc.Text = ""
        Else
            Label_Priority.Text = clsLang.L("우선순위 ", "優先順位 ") & ru.Priority
            CheckBox_Enabled.Checked = ru.Enabled
            ComboBox_Dest.SelectedIndex = -1
            For i As Integer = 0 To ComboBox_Dest.Items.Count - 1
                If DirectCast(ComboBox_Dest.Items(i), clsDestination).DestID = ru.DestID Then ComboBox_Dest.SelectedIndex = i
            Next
            TextBox_Desc.Text = ru.Description
            For Each c As clsRuleCond In ru.Conds
                DataGridView_Cond.Rows.Add(c.JoinOp, c.Target, c.Operator, ValuesDisplay(c))
            Next
            Call FixFirstJoinCell()
        End If
        mLoading = False

        If DataGridView_Cond.Rows.Count > 0 Then
            DataGridView_Cond.CurrentCell = DataGridView_Cond.Rows(0).Cells(colValues.Index)
            Call ShowValues()
        End If
    End Sub

    Private Shared Function ValuesDisplay(ByVal c As clsRuleCond) As String
        Return String.Join(" | ", c.ValueList())
    End Function

    Private Sub FixFirstJoinCell()
        '==========================================================================================
        'Function : 첫 조건 줄은 앞 줄이 없으므로 연결(AND / OR) 칸을 비우고 고칠 수 없게 한다
        '==========================================================================================
        If DataGridView_Cond.Rows.Count = 0 Then Exit Sub
        Dim r As DataGridViewRow = DataGridView_Cond.Rows(0)
        If TypeOf r.Cells(colJoin.Index) Is DataGridViewComboBoxCell Then
            r.Cells(colJoin.Index) = New DataGridViewTextBoxCell With {.Value = "－"}
        End If
        r.Cells(colJoin.Index).ReadOnly = True
        r.Cells(colJoin.Index).Style.ForeColor = clsUiTheme.TextSub
    End Sub

    '분기 COM (분기 대상) 신규 / 수정 / 삭제 ================================================================================
    '  2026-10-26 지시 : 분기 COM 이 COM8 / COM11 로 고정되어 보임 → 이 화면에서 새로 등록, 이름 / COM 수정, 신규 추가
    '  분기 COM 은 규칙과 따로 바로 RouterSetting.mdb (tblDestination) 에 저장한다
    Private Sub RefreshDestCombo(ByVal selectDestID As Integer)
        Dim keep As Boolean = mLoading
        mLoading = True
        ComboBox_Dest.Items.Clear()
        For Each d As clsDestination In mSetting.Destinations
            ComboBox_Dest.Items.Add(d)
            If d.DestID = selectDestID Then ComboBox_Dest.SelectedIndex = ComboBox_Dest.Items.Count - 1
        Next
        mLoading = keep
    End Sub

    Private Function UsedPorts(ByVal exceptDestID As Integer) As List(Of String)
        Dim lst As New List(Of String)
        If mSetting.Reader.InputMode = clsReaderSetting.MODE_COM Then lst.Add(mSetting.Reader.Port.ToUpperInvariant)
        For Each d As clsDestination In mSetting.Destinations
            If d.DestID <> exceptDestID AndAlso d.Enabled Then lst.Add(d.Port.ToUpperInvariant)
        Next
        Return lst
    End Function

    Private Function SaveDestinations() As Boolean
        Try
            mSetting.SaveReaderAndDestinations()
            Return True
        Catch ex As Exception
            MessageBox.Show(clsLang.L("분기 COM 저장 NG", "分岐COM 保存 NG") & vbCrLf & ex.Message, clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub Button_DestNew_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_DestNew.Click
        Dim newID As Integer = 1
        If mSetting.Destinations.Count > 0 Then newID = mSetting.Destinations.Max(Function(x) x.DestID) + 1
        Dim d As New clsDestination With {.DestID = newID, .Setting = "", .Enabled = True}     '통신설정 / 종료 신호는 비워 두고 입력받는다 (기본값을 두지 않음)

        Using f As New frmDestEdit(d, True, UsedPorts(newID))
            If f.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
        End Using

        mSetting.Destinations.Add(d)
        If Not SaveDestinations() Then mSetting.Destinations.Remove(d) : Exit Sub
        Call RefreshDestCombo(d.DestID)
        If CurRule IsNot Nothing Then CurRule.DestID = d.DestID          '고르고 있던 규칙의 분기 COM 으로 바로 지정
    End Sub

    Private Sub Button_DestEdit_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_DestEdit.Click
        Dim cur As clsDestination = TryCast(ComboBox_Dest.SelectedItem, clsDestination)
        If cur Is Nothing Then Exit Sub
        Dim work As New clsDestination With {.DestID = cur.DestID, .DestName = cur.DestName, .Port = cur.Port, .Setting = cur.Setting,
                                             .ProcessName = cur.ProcessName, .EndTimeoutSec = cur.EndTimeoutSec, .Enabled = cur.Enabled,
                                             .EndMessage = cur.EndMessage, .LineEnd = cur.LineEnd}

        Using f As New frmDestEdit(work, False, UsedPorts(cur.DestID))
            If f.ShowDialog(Me) <> DialogResult.OK Then Exit Sub
        End Using

        Dim i As Integer = mSetting.Destinations.IndexOf(cur)
        mSetting.Destinations(i) = work
        If Not SaveDestinations() Then mSetting.Destinations(i) = cur : Exit Sub
        Call RefreshDestCombo(work.DestID)
    End Sub

    Private Sub Button_DestDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_DestDelete.Click
        Dim cur As clsDestination = TryCast(ComboBox_Dest.SelectedItem, clsDestination)
        If cur Is Nothing Then Exit Sub
        '화면의 규칙이나 저장된 규칙에서 쓰고 있으면 지우지 않는다 (먼저 그 규칙의 분기 COM 을 바꾸고 저장)
        Dim users As New List(Of String)
        For Each ru As clsRouteRule In mRules.Concat(mSetting.Rules)
            If ru.DestID = cur.DestID AndAlso Not users.Contains(ru.Description) Then users.Add(ru.Description)
        Next
        If users.Count > 0 Then
            MessageBox.Show(clsLang.L("이 분기 COM 을 쓰는 규칙이 있어서 지울 수 없습니다 : ", "この分岐COMを使うルールがあるため削除できません : ") & String.Join(", ", users) & vbCrLf &
                            clsLang.L("규칙의 분기 COM 을 바꾸고 저장한 뒤 지워 주세요.", "ルールの分岐COMを変更・保存してから削除してください。"), clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If MessageBox.Show(clsLang.L("분기 COM '", "分岐COM '") & cur.DisplayName & clsLang.L("' 을 지웁니까?", "' を削除しますか?"), clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub

        Dim i As Integer = mSetting.Destinations.IndexOf(cur)
        mSetting.Destinations.RemoveAt(i)
        If Not SaveDestinations() Then mSetting.Destinations.Insert(i, cur) : Exit Sub
        Call RefreshDestCombo(If(CurRule Is Nothing, -1, CurRule.DestID))
    End Sub

    Private Sub Button_New_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_New.Click
        Dim ru As New clsRouteRule With {.Enabled = True, .Description = clsLang.L("새 규칙", "新しいルール")}
        If ComboBox_Dest.Items.Count > 0 Then ru.DestID = DirectCast(ComboBox_Dest.Items(0), clsDestination).DestID
        ru.Conds.Add(New clsRuleCond With {.Target = mDefaultTarget})
        mRules.Add(ru)
        Call Renumber()
        Call RefreshRuleList(mRules.Count - 1)
        TextBox_Desc.Focus()
        TextBox_Desc.SelectAll()
    End Sub

    Private Sub Button_Delete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Delete.Click
        Dim ru As clsRouteRule = CurRule
        If ru Is Nothing Then Exit Sub
        If MessageBox.Show("'" & ru.Description & clsLang.L("' 규칙을 지웁니까?", "' ルールを削除しますか?"), clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub
        Dim i As Integer = ComboBox_Rule.SelectedIndex
        mRules.RemoveAt(i)
        Call Renumber()
        Call RefreshRuleList(i)
    End Sub

    Private Sub Button_Up_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Up.Click
        Call MoveRule(-1)
    End Sub

    Private Sub Button_Down_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Down.Click
        Call MoveRule(1)
    End Sub

    Private Sub MoveRule(ByVal dir As Integer)
        Dim i As Integer = ComboBox_Rule.SelectedIndex
        Dim j As Integer = i + dir
        If i < 0 OrElse j < 0 OrElse j >= mRules.Count Then Exit Sub
        Dim t As clsRouteRule = mRules(i)
        mRules(i) = mRules(j)
        mRules(j) = t
        Call Renumber()
        Call RefreshRuleList(j)
    End Sub

    Private Sub RefreshCurrentRuleText()
        Dim i As Integer = ComboBox_Rule.SelectedIndex
        If i < 0 Then Exit Sub
        mLoading = True
        ComboBox_Rule.Items(i) = mRules(i).ToString
        mLoading = False
    End Sub

    '규칙 항목 ==============================================================================================================
    Private Sub CheckBox_Enabled_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles CheckBox_Enabled.CheckedChanged
        If mLoading OrElse CurRule Is Nothing Then Exit Sub
        CurRule.Enabled = CheckBox_Enabled.Checked
        Call RefreshCurrentRuleText()
    End Sub

    Private Sub ComboBox_Dest_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ComboBox_Dest.SelectedIndexChanged
        If mLoading OrElse CurRule Is Nothing OrElse ComboBox_Dest.SelectedItem Is Nothing Then Exit Sub
        CurRule.DestID = DirectCast(ComboBox_Dest.SelectedItem, clsDestination).DestID
    End Sub

    Private Sub TextBox_Desc_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox_Desc.TextChanged
        If mLoading OrElse CurRule Is Nothing Then Exit Sub
        CurRule.Description = TextBox_Desc.Text.Trim
        Call RefreshCurrentRuleText()
    End Sub

    '조건 줄 ================================================================================================================
    Private ReadOnly Property CurCond() As clsRuleCond
        Get
            Dim ru As clsRouteRule = CurRule
            If ru Is Nothing OrElse DataGridView_Cond.CurrentRow Is Nothing Then Return Nothing
            Dim i As Integer = DataGridView_Cond.CurrentRow.Index
            If i < 0 OrElse i >= ru.Conds.Count Then Return Nothing
            Return ru.Conds(i)
        End Get
    End Property

    Private Sub Button_CondAdd_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_CondAdd.Click
        Dim ru As clsRouteRule = CurRule
        If ru Is Nothing Then Exit Sub
        Dim c As New clsRuleCond With {.Target = mDefaultTarget}
        ru.Conds.Add(c)
        mLoading = True
        DataGridView_Cond.Rows.Add(c.JoinOp, c.Target, c.Operator, "")
        Call FixFirstJoinCell()
        mLoading = False
        DataGridView_Cond.CurrentCell = DataGridView_Cond.Rows(DataGridView_Cond.Rows.Count - 1).Cells(colValues.Index)
        Call ShowValues()
        TextBox_Values.Focus()
    End Sub

    Private Sub Button_CondDelete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_CondDelete.Click
        Dim ru As clsRouteRule = CurRule
        If ru Is Nothing OrElse DataGridView_Cond.CurrentRow Is Nothing Then Exit Sub
        Dim i As Integer = DataGridView_Cond.CurrentRow.Index
        ru.Conds.RemoveAt(i)
        Call ShowRule(ru)
    End Sub

    Private Sub DataGridView_Cond_CurrentCellDirtyStateChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView_Cond.CurrentCellDirtyStateChanged
        'ComboBox 칸은 고르는 즉시 값을 반영한다
        If DataGridView_Cond.IsCurrentCellDirty Then DataGridView_Cond.CommitEdit(DataGridViewDataErrorContexts.Commit)
    End Sub

    Private Sub DataGridView_Cond_CellValueChanged(ByVal sender As Object, ByVal e As DataGridViewCellEventArgs) Handles DataGridView_Cond.CellValueChanged
        If mLoading OrElse e.RowIndex < 0 Then Exit Sub
        Dim ru As clsRouteRule = CurRule
        If ru Is Nothing OrElse e.RowIndex >= ru.Conds.Count Then Exit Sub
        Dim v As Object = DataGridView_Cond.Rows(e.RowIndex).Cells(e.ColumnIndex).Value
        If v Is Nothing Then Exit Sub
        If e.ColumnIndex = colJoin.Index AndAlso e.RowIndex > 0 Then ru.Conds(e.RowIndex).JoinOp = v.ToString
        If e.ColumnIndex = colTarget.Index Then ru.Conds(e.RowIndex).Target = v.ToString
        If e.ColumnIndex = colOperator.Index Then ru.Conds(e.RowIndex).Operator = v.ToString
    End Sub

    Private Sub DataGridView_Cond_SelectionChanged(ByVal sender As Object, ByVal e As EventArgs) Handles DataGridView_Cond.SelectionChanged
        If mLoading Then Exit Sub
        Call ShowValues()
    End Sub

    Private Sub ShowValues()
        Dim c As clsRuleCond = CurCond
        mLoading = True
        TextBox_Values.Enabled = (c IsNot Nothing)
        TextBox_Values.Text = If(c Is Nothing, "", c.InputValues)
        mLoading = False
    End Sub

    Private Sub TextBox_Values_TextChanged(ByVal sender As Object, ByVal e As EventArgs) Handles TextBox_Values.TextChanged
        If mLoading Then Exit Sub
        Dim c As clsRuleCond = CurCond
        If c Is Nothing Then Exit Sub
        c.InputValues = TextBox_Values.Text
        mLoading = True
        DataGridView_Cond.CurrentRow.Cells(colValues.Index).Value = ValuesDisplay(c)
        mLoading = False
    End Sub

    '판단 시험 / 저장 =======================================================================================================
    Private Sub Button_Test_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Test.Click
        '==========================================================================================
        'Function : 판단 시험 - 입력 형식 (2026-10-28, 항목이 설정으로 바뀌어서)
        '           "항목ID=값; 항목ID=값"  예 : MSCODE=F3SP71-4S; SERIAL=C3AA0001
        '           = 이 없으면 고른 조건 줄의 대상 항목 값으로 본다. DERIVE 항목(예 : MODEL)은 설정대로 만든다
        '==========================================================================================
        Dim text As String = TextBox_TestMSCODE.Text.Trim
        If text = "" Then Exit Sub
        Call Renumber()

        Dim given As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)
        If text.Contains("=") Then
            For Each part As String In text.Split(";"c)
                Dim p As Integer = part.IndexOf("="c)
                If p > 0 Then given(part.Substring(0, p).Trim) = part.Substring(p + 1).Trim
            Next
        Else
            Dim key As String = If(CurCond IsNot Nothing AndAlso CurCond.Target <> "", CurCond.Target, mDefaultTarget)
            given(key) = text
        End If

        '입력한 값을 항목 값으로 넣고, 입력하지 않은 DERIVE 항목은 설정대로 만든다 (PLASMA 조회 없이)
        Dim values As New Dictionary(Of String, String)(given, StringComparer.OrdinalIgnoreCase)
        Dim derives As List(Of clsField) = mSetting.Fields.Where(Function(x) x.SourceType = clsField.SRC_DERIVE AndAlso Not given.ContainsKey(x.FieldID)).ToList
        For pass As Integer = 1 To Math.Max(1, derives.Count)       'DERIVE 가 DERIVE 를 쓰는 경우
            For Each f As clsField In derives
                values(f.FieldID) = clsRouteData.Derive(f, ValueOrEmpty(values, f.BaseField))
            Next
        Next
        Dim shown As String = String.Join(", ", values.Where(Function(x) x.Value <> "").Select(Function(x) x.Key & "=" & x.Value))

        Dim ru As clsRouteRule = clsRouteJudge.Judge(mRules, values)
        Dim c As clsUiTheme.StatusColor
        If ru Is Nothing Then
            c = clsUiTheme.GetStatusColor(clsUiTheme.StatusKind.Error)
            Label_TestResult.Text = clsLang.L("  ✕  분기 대상 없음  (", "  ✕  分岐先なし  (") & shown & ")"
        Else
            c = clsUiTheme.GetStatusColor(clsUiTheme.StatusKind.Ready)
            Dim d As clsDestination = mSetting.FindDestination(ru.DestID)
            Label_TestResult.Text = "  ✓  " & ru.ToString & " → " & If(d Is Nothing, clsLang.L("(분기 COM 없음)", "(分岐COMなし)"), d.DisplayName)
        End If
        Label_TestResult.BackColor = c.Back
        Label_TestResult.ForeColor = c.Main
    End Sub

    Private Shared Function ValueOrEmpty(ByVal values As Dictionary(Of String, String), ByVal key As String) As String
        Dim v As String = ""
        values.TryGetValue(key, v)
        Return If(v, "")
    End Function

    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        DataGridView_Cond.EndEdit()

        For Each ru As clsRouteRule In mRules
            If ru.Description = "" Then
                MessageBox.Show(clsLang.L("우선순위 ", "優先順位 ") & ru.Priority & clsLang.L(" : 설명을 넣어 주세요.", " : 説明を入力してください。"), clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If mSetting.FindDestination(ru.DestID) Is Nothing Then
                MessageBox.Show("'" & ru.Description & clsLang.L("' : 분기 COM 을 골라 주세요.", "' : 分岐COMを選んでください。"), clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If ru.Conds.Count = 0 OrElse ru.Conds.Any(Function(c) c.ValueList().Count = 0) Then
                MessageBox.Show("'" & ru.Description & clsLang.L("' : 조건 줄마다 입력값을 한 개 이상 넣어 주세요.", "' : 条件行ごとに入力値を1つ以上入力してください。"), clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
        Next

        Call Renumber()
        For i As Integer = 0 To mRules.Count - 1
            mRules(i).RuleID = i + 1
        Next

        Try
            mSetting.Rules = mRules
            mSetting.SaveRules()
            Me.DialogResult = DialogResult.OK
            Me.Close()
        Catch ex As Exception
            MessageBox.Show(clsLang.L("저장 NG", "保存 NG") & vbCrLf & ex.Message, clsLang.L("분기 조건 설정", "分岐条件設定"), MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    Private Sub ApplyLanguageTexts()
        Me.Label_ValuesCap.Text = clsLang.L("Enter 로 줄 나눔 (직접 입력) · 하나라도 맞으면 OK", "Enterで改行 (直接入力) · 1つでも合えばOK") & vbCrLf &
                                  clsLang.L("LIKE 예 : F3* = F3 로 시작", "LIKE 例 : F3* = F3で始まる")
        Me.colJoin.ToolTipText = clsLang.L("앞 줄과의 연결 (AND / OR). 첫 줄은 없음", "前の行との接続 (AND / OR)。最初の行はなし")
        '==========================================================================================
        'Function : Designer 에 적힌 한국어 글자를 언어 설정(clsLang)에 맞게 바꾼다 (Designer 파일은 그대로, 2026-10-28)
        '==========================================================================================
        Me.Label_Header.Text = clsLang.L("⑂  분기 조건 설정", "⑂  分岐条件設定")
        Me.Label_RuleCap.Text = clsLang.L("분기 규칙", "分岐ルール")
        Me.Button_New.Text = clsLang.L("＋  신규", "＋  新規")
        Me.Button_Delete.Text = clsLang.L("✕  삭제", "✕  削除")
        Me.CheckBox_Enabled.Text = clsLang.L("사용", "使用")
        Me.Label_Priority.Text = clsLang.L("우선순위 -", "優先順位 -")
        Me.Label_DestCap.Text = clsLang.L("분기 COM", "分岐COM")
        Me.Button_DestNew.Text = clsLang.L("＋ 신규", "＋ 新規")
        Me.Button_DestEdit.Text = clsLang.L("✎ 수정", "✎ 修正")
        Me.Button_DestDelete.Text = clsLang.L("✕ 삭제", "✕ 削除")
        Me.Label_DescCap.Text = clsLang.L("설명", "説明")
        Me.Label_CondCap.Text = clsLang.L("조건  · 줄마다 연결 AND / OR 선택 (AND 를 먼저 묶음)", "条件  · 行ごとに接続 AND / OR を選択 (ANDを先にまとめる)")
        Me.colJoin.HeaderText = clsLang.L("연결", "接続")
        Me.colTarget.HeaderText = clsLang.L("대상", "対象")
        Me.colOperator.HeaderText = clsLang.L("조건", "条件")
        Me.colValues.HeaderText = clsLang.L("입력값 (오른쪽 칸에서 입력)", "入力値 (右の欄で入力)")
        Me.Button_CondAdd.Text = clsLang.L("＋  조건 추가", "＋  条件追加")
        Me.Button_CondDelete.Text = clsLang.L("－  조건 삭제", "－  条件削除")
        Me.Label_ValuesTitle.Text = clsLang.L("입력값 (선택한 조건 줄)", "入力値 (選択した条件行)")
        Me.Label_TestCap.Text = clsLang.L("판단 시험", "判定試験")
        Me.Button_Test.Text = clsLang.L("시험", "試験")
        Me.Label_TestResult.Text = clsLang.L("항목ID=값; 항목ID=값  (= 없으면 고른 조건 줄 항목)", "項目ID=値; 項目ID=値  (= なしは選択した条件行の項目)")
        Me.Button_Save.Text = clsLang.L("저장", "保存")
        Me.Button_Close.Text = clsLang.L("닫기", "閉じる")
        Me.Text = clsLang.L("분기 조건 설정", "分岐条件設定")
    End Sub

End Class
