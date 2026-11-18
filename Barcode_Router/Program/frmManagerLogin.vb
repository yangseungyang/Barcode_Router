Option Strict On : Option Explicit On
'==========================================================
'Form Name  : frmManagerLogin
'Function   : 담당자 로그인 - 사번을 넣으면 tblManager 명단에서 찾아 이름을 보여 주고 로그인한다
'Framwork   : VB2022 Framework4.8
'Memo       : 2026-11-18 지시 - 작업자가 설정을 바꾸지 못하게, 담당자가 로그인했을 때만 설정 / 설명서 버튼을 보인다
'             - 비밀번호 없음 (사번만). 명단은 RouterSetting.mdb tblManager (사번, 이름, 사용) - 프로그램에 사번을 넣지 않는다
'             - 명단이 비어 있으면 로그인할 수 없다 (Access 로 tblManager 에 담당자를 넣어야 함)
'             - 화면은 Designer 없이 코드로 만든다
'Issue      : 2026/11/18  Rev 0.00    YSH       New Release
'==========================================================

Public Class frmManagerLogin
    Inherits Form

    Private ReadOnly mSetting As clsRouterSetting
    Private mManager As clsManager = Nothing

    Private TextBox_EmpNo As New TextBox With {.Location = New Point(24, 84), .Size = New Size(312, 30)}
    Private Label_Result As New Label With {.Location = New Point(24, 122), .Size = New Size(312, 40), .ForeColor = clsUiTheme.TextSub}
    Private WithEvents Button_Login As New Button With {.Text = clsLang.L("로그인", "ログイン"), .Location = New Point(130, 172), .Size = New Size(100, 36)}
    Private WithEvents Button_Cancel As New Button With {.Text = clsLang.L("취소", "キャンセル"), .Location = New Point(236, 172), .Size = New Size(100, 36), .DialogResult = DialogResult.Cancel}

    ''' <summary>로그인한 담당자 (취소하면 Nothing)</summary>
    Public ReadOnly Property Manager() As clsManager
        Get
            Return mManager
        End Get
    End Property

    Public Sub New(ByVal setting As clsRouterSetting)
        mSetting = setting
        Me.Text = clsLang.L("담당자 로그인", "担当者ログイン")
        Me.ClientSize = New Size(360, 224)
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False : Me.MinimizeBox = False
        Me.ShowInTaskbar = False
        Me.StartPosition = FormStartPosition.CenterParent
        Me.BackColor = Color.White
        Me.Font = clsUiTheme.UiFont(9.5!)

        Dim header As New Panel With {.Dock = DockStyle.Top, .Height = 42, .BackColor = clsUiTheme.HeaderBack}
        header.Controls.Add(New Label With {.Text = clsLang.L("담당자 로그인", "担当者ログイン"), .AutoSize = True, .Location = New Point(12, 9), .ForeColor = clsUiTheme.HeaderText,
                                            .Font = clsUiTheme.UiFont(12.0!, FontStyle.Bold)})
        Me.Controls.Add(header)
        Me.Controls.Add(New Label With {.Text = clsLang.L("사번", "社員番号"), .Location = New Point(24, 60), .AutoSize = True, .ForeColor = clsUiTheme.TextMain})
        TextBox_EmpNo.Font = clsUiTheme.UiFont(13.0!)
        Me.Controls.Add(TextBox_EmpNo)
        Me.Controls.Add(Label_Result)
        Me.Controls.Add(Button_Login)
        Me.Controls.Add(Button_Cancel)
        clsUiTheme.StyleButton(Button_Login, clsUiTheme.ButtonKind.Primary)
        clsUiTheme.StyleButton(Button_Cancel, clsUiTheme.ButtonKind.Normal)
        Me.AcceptButton = Button_Login
        Me.CancelButton = Button_Cancel

        If mSetting.Managers.Where(Function(x) x.Enabled).Count = 0 Then
            Label_Result.Text = clsLang.L("등록된 담당자가 없습니다." & vbCrLf & "RouterSetting.mdb 의 tblManager 에 사번 / 이름을 넣어 주세요.",
                                          "登録された担当者がいません。" & vbCrLf & "RouterSetting.mdb の tblManager に社員番号 / 名前を入力してください。")
            Label_Result.ForeColor = clsUiTheme.GetStatusColor(clsUiTheme.StatusKind.Error).Main
            TextBox_EmpNo.Enabled = False
            Button_Login.Enabled = False
        End If
    End Sub

    Private Sub frmManagerLogin_Shown(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Shown
        If TextBox_EmpNo.Enabled Then TextBox_EmpNo.Focus() Else Button_Cancel.Focus()
    End Sub

    Private Sub Button_Login_Click(ByVal sender As Object, ByVal e As EventArgs) Handles Button_Login.Click
        Dim no As String = TextBox_EmpNo.Text.Trim
        If no = "" Then
            TextBox_EmpNo.Focus()
            Exit Sub
        End If
        Dim m As clsManager = mSetting.FindManager(no)
        If m Is Nothing Then
            Label_Result.Text = clsLang.L("등록되지 않은 사번입니다 : ", "登録されていない社員番号です : ") & no
            Label_Result.ForeColor = clsUiTheme.GetStatusColor(clsUiTheme.StatusKind.Error).Main
            TextBox_EmpNo.SelectAll()
            TextBox_EmpNo.Focus()
            Exit Sub
        End If
        mManager = m
        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub

End Class
