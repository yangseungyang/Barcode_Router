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
'                                대상 ComboBox (MSCODE / MODEL), 조건 ComboBox (LIKE / EQUAL)
'             - 입력값         : 선택한 조건 줄의 값. 여러 줄 TextBox 에 Enter 로 줄을 나눠 직접 입력 (한 줄 = 값 하나, OR)
'             - 우선순위       : 목록 순서 (▲ / ▼ 로 바꿈). 저장할 때 1 부터 다시 매긴다
'             - 판단 시험      : MSCODE 를 넣으면 화면 내용(저장 전) 기준으로 어떤 규칙에 맞는지 보여 준다
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-3항 / 구조설계 U-01 ~ U-08
'==========================================================

Public Class frmRuleSetting

    Private mSetting As clsRouterSetting
    Private mRules As New List(Of clsRouteRule)     '화면에서 고치는 사본 (저장할 때만 설정에 반영)
    Private mLoading As Boolean = False

    Public Sub New(ByVal setting As clsRouterSetting)
        InitializeComponent()
        mSetting = setting
    End Sub

    Private Sub frmRuleSetting_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
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
        colTarget.Items.AddRange(New Object() {clsRuleCond.TARGET_MSCODE, clsRuleCond.TARGET_MODEL})
        colOperator.Items.AddRange(New Object() {clsRuleCond.OP_LIKE, clsRuleCond.OP_EQUAL})

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
            Label_Priority.Text = "우선순위 -"
            CheckBox_Enabled.Checked = False
            ComboBox_Dest.SelectedIndex = -1
            TextBox_Desc.Text = ""
        Else
            Label_Priority.Text = "우선순위 " & ru.Priority
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
            MessageBox.Show("분기 COM 저장 NG" & vbCrLf & ex.Message, "분기 조건 설정", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Sub Button_DestNew_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_DestNew.Click
        Dim newID As Integer = 1
        If mSetting.Destinations.Count > 0 Then newID = mSetting.Destinations.Max(Function(x) x.DestID) + 1
        Dim d As New clsDestination With {.DestID = newID, .Setting = "9600-N-8-1", .Enabled = True}

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
                                             .ProcessName = cur.ProcessName, .EndTimeoutSec = cur.EndTimeoutSec, .Enabled = cur.Enabled}

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
            MessageBox.Show("이 분기 COM 을 쓰는 규칙이 있어서 지울 수 없습니다 : " & String.Join(", ", users) & vbCrLf &
                            "규칙의 분기 COM 을 바꾸고 저장한 뒤 지워 주세요.", "분기 조건 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If
        If MessageBox.Show("분기 COM '" & cur.DisplayName & "' 을 지웁니까?", "분기 조건 설정", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub

        Dim i As Integer = mSetting.Destinations.IndexOf(cur)
        mSetting.Destinations.RemoveAt(i)
        If Not SaveDestinations() Then mSetting.Destinations.Insert(i, cur) : Exit Sub
        Call RefreshDestCombo(If(CurRule Is Nothing, -1, CurRule.DestID))
    End Sub

    Private Sub Button_New_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_New.Click
        Dim ru As New clsRouteRule With {.Enabled = True, .Description = "새 규칙"}
        If ComboBox_Dest.Items.Count > 0 Then ru.DestID = DirectCast(ComboBox_Dest.Items(0), clsDestination).DestID
        ru.Conds.Add(New clsRuleCond)
        mRules.Add(ru)
        Call Renumber()
        Call RefreshRuleList(mRules.Count - 1)
        TextBox_Desc.Focus()
        TextBox_Desc.SelectAll()
    End Sub

    Private Sub Button_Delete_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Delete.Click
        Dim ru As clsRouteRule = CurRule
        If ru Is Nothing Then Exit Sub
        If MessageBox.Show("'" & ru.Description & "' 규칙을 지웁니까?", "분기 조건 설정", MessageBoxButtons.YesNo, MessageBoxIcon.Question) <> DialogResult.Yes Then Exit Sub
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
        Dim c As New clsRuleCond
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
        Dim mscode As String = TextBox_TestMSCODE.Text.Trim
        If mscode = "" Then Exit Sub
        Call Renumber()
        Dim ru As clsRouteRule = clsRouteJudge.Judge(mRules, mscode)
        Dim c As clsUiTheme.StatusColor
        If ru Is Nothing Then
            c = clsUiTheme.GetStatusColor(clsUiTheme.StatusKind.Error)
            Label_TestResult.Text = "  ✕  분기 대상 없음  (MODEL = " & clsRouteJudge.GetModel(mscode) & ")"
        Else
            c = clsUiTheme.GetStatusColor(clsUiTheme.StatusKind.Ready)
            Dim d As clsDestination = mSetting.FindDestination(ru.DestID)
            Label_TestResult.Text = "  ✓  " & ru.ToString & " → " & If(d Is Nothing, "(분기 COM 없음)", d.DisplayName)
        End If
        Label_TestResult.BackColor = c.Back
        Label_TestResult.ForeColor = c.Main
    End Sub

    Private Sub Button_Save_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Save.Click
        DataGridView_Cond.EndEdit()

        For Each ru As clsRouteRule In mRules
            If ru.Description = "" Then
                MessageBox.Show("우선순위 " & ru.Priority & " : 설명을 넣어 주세요.", "분기 조건 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If mSetting.FindDestination(ru.DestID) Is Nothing Then
                MessageBox.Show("'" & ru.Description & "' : 분기 COM 을 골라 주세요.", "분기 조건 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Exit Sub
            End If
            If ru.Conds.Count = 0 OrElse ru.Conds.Any(Function(c) c.ValueList().Count = 0) Then
                MessageBox.Show("'" & ru.Description & "' : 조건 줄마다 입력값을 한 개 이상 넣어 주세요.", "분기 조건 설정", MessageBoxButtons.OK, MessageBoxIcon.Warning)
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
            MessageBox.Show("저장 NG" & vbCrLf & ex.Message, "분기 조건 설정", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

End Class
