Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsUiTheme
'Function   : 바코드 분배 프로그램 화면 색 / 버튼 / 표 모양 (화면 3개 공통)
'Framwork   : VB2022 Framework4.8
'Memo       : 화면 시안 0_신작\1_분석\Router_UI_시안_v1.html (2026-10-26 지시자 승인) 의 색을 그대로 쓴다
'             - 제목줄 진한 파랑, 상태 패널 : 대기 초록 / 검사 중 주황 / 오류 빨강 / 조회 중 파랑
'             - 버튼은 평면(Flat), 표는 회색 머리줄 + 파란 선택줄
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'============================================================================================================================

Public Class clsUiTheme

    '색 =====================================================================================================================
    Public Shared ReadOnly HeaderBack As Color = ColorTranslator.FromHtml("#0C447C")
    Public Shared ReadOnly HeaderText As Color = ColorTranslator.FromHtml("#E6F1FB")
    Public Shared ReadOnly HeaderSub As Color = ColorTranslator.FromHtml("#B5D4F4")
    Public Shared ReadOnly TextMain As Color = ColorTranslator.FromHtml("#2C2C2A")
    Public Shared ReadOnly TextSub As Color = ColorTranslator.FromHtml("#5F5E5A")
    Public Shared ReadOnly TextSection As Color = ColorTranslator.FromHtml("#444441")
    Public Shared ReadOnly CardBack As Color = ColorTranslator.FromHtml("#F1EFE8")
    Public Shared ReadOnly Line As Color = ColorTranslator.FromHtml("#B4B2A9")
    Public Shared ReadOnly LineLight As Color = ColorTranslator.FromHtml("#D3D1C7")
    Public Shared ReadOnly InputLine As Color = ColorTranslator.FromHtml("#378ADD")
    Public Shared ReadOnly InputText As Color = ColorTranslator.FromHtml("#042C53")
    Public Shared ReadOnly SelectBack As Color = ColorTranslator.FromHtml("#E6F1FB")
    Public Shared ReadOnly SelectText As Color = ColorTranslator.FromHtml("#0C447C")
    Public Shared ReadOnly PrimaryBack As Color = ColorTranslator.FromHtml("#185FA5")

    Public Enum StatusKind
        Ready       '대기 (초록)
        Busy        '검사 중 (주황)
        [Error]     '오류 (빨강)
        Info        '조회 중 (파랑)
    End Enum

    Public Structure StatusColor
        Public Back As Color
        Public Border As Color
        Public Main As Color
        Public [Sub] As Color
    End Structure

    Public Shared Function GetStatusColor(ByVal k As StatusKind) As StatusColor
        Dim c As New StatusColor
        Select Case k
            Case StatusKind.Busy
                c.Back = ColorTranslator.FromHtml("#FAEEDA") : c.Border = ColorTranslator.FromHtml("#EF9F27")
                c.Main = ColorTranslator.FromHtml("#633806") : c.Sub = ColorTranslator.FromHtml("#854F0B")
            Case StatusKind.Error
                c.Back = ColorTranslator.FromHtml("#FCEBEB") : c.Border = ColorTranslator.FromHtml("#E24B4A")
                c.Main = ColorTranslator.FromHtml("#791F1F") : c.Sub = ColorTranslator.FromHtml("#A32D2D")
            Case StatusKind.Info
                c.Back = ColorTranslator.FromHtml("#E6F1FB") : c.Border = ColorTranslator.FromHtml("#378ADD")
                c.Main = ColorTranslator.FromHtml("#0C447C") : c.Sub = ColorTranslator.FromHtml("#185FA5")
            Case Else
                c.Back = ColorTranslator.FromHtml("#EAF3DE") : c.Border = ColorTranslator.FromHtml("#97C459")
                c.Main = ColorTranslator.FromHtml("#27500A") : c.Sub = ColorTranslator.FromHtml("#3B6D11")
        End Select
        Return c
    End Function

    '글꼴 ===================================================================================================================
    '  tblConfig UiFont (2026-10-28) : 한국어 / 일본어 Windows 에 따라 있는 글꼴이 다를 수 있어 설정으로 고른다
    '  비어 있으면 언어(clsLang, 한국어 / 일본어 PC)에 맞는 글꼴, PC 에 없는 글꼴이면 Windows 기본 글꼴
    Private Shared mFontName As String = ""

    Public Shared Sub SetFontName(ByVal name As String)
        '설정 글꼴(tblConfig UiFont) → 비어 있으면 언어 글꼴(clsLang.LanguageFont) → PC 에 없으면 Windows 기본 글꼴
        Dim want As String = If(name Is Nothing OrElse name.Trim = "", clsLang.LanguageFont(), name.Trim)
        Try
            Using ff As New FontFamily(want)        '없는 글꼴이면 예외
                mFontName = ff.Name
            End Using
        Catch
            mFontName = SystemFonts.MessageBoxFont.Name
        End Try
    End Sub

    Public Shared ReadOnly Property FontName() As String
        Get
            Return If(mFontName = "", SystemFonts.MessageBoxFont.Name, mFontName)
        End Get
    End Property

    Public Shared Function UiFont(ByVal size As Single, Optional ByVal style As FontStyle = FontStyle.Regular) As Font
        Return New Font(FontName, size, style)
    End Function

    Public Shared Sub ApplyFont(ByVal root As Control)
        '==========================================================================================
        'Function : 화면의 모든 컨트롤 글꼴을 설정한 글꼴로 바꾼다 (크기 / 굵기는 그대로). 설정이 없으면 아무것도 안 함
        '==========================================================================================
        If mFontName = "" Then Exit Sub
        If root.Font.Name <> mFontName Then root.Font = New Font(mFontName, root.Font.Size, root.Font.Style)
        For Each c As Control In root.Controls
            Call ApplyFont(c)
        Next
    End Sub

    '버튼 ===================================================================================================================
    Public Enum ButtonKind
        Normal      '흰 바탕 회색 테두리
        Accent      '연한 파랑 (주요 기능)
        Primary     '진한 파랑 (저장)
        Danger      '연한 빨강 (삭제 / 대기 해제)
    End Enum

    Public Shared Sub StyleButton(ByVal b As Button, ByVal k As ButtonKind)
        b.FlatStyle = FlatStyle.Flat
        b.FlatAppearance.BorderSize = 1
        b.Cursor = Cursors.Hand
        b.UseVisualStyleBackColor = False
        Select Case k
            Case ButtonKind.Accent
                b.BackColor = SelectBack : b.ForeColor = SelectText
                b.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#85B7EB")
                b.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#B5D4F4")
            Case ButtonKind.Primary
                b.BackColor = PrimaryBack : b.ForeColor = HeaderText
                b.FlatAppearance.BorderColor = PrimaryBack
                b.FlatAppearance.MouseOverBackColor = HeaderBack
            Case ButtonKind.Danger
                b.BackColor = ColorTranslator.FromHtml("#FCEBEB") : b.ForeColor = ColorTranslator.FromHtml("#791F1F")
                b.FlatAppearance.BorderColor = ColorTranslator.FromHtml("#F09595")
                b.FlatAppearance.MouseOverBackColor = ColorTranslator.FromHtml("#F7C1C1")
            Case Else
                b.BackColor = Color.White : b.ForeColor = TextMain
                b.FlatAppearance.BorderColor = Line
                b.FlatAppearance.MouseOverBackColor = CardBack
        End Select
    End Sub

    '표 =====================================================================================================================
    Public Shared Sub StyleGrid(ByVal g As DataGridView)
        g.BackgroundColor = Color.White
        g.BorderStyle = BorderStyle.FixedSingle
        g.GridColor = LineLight
        g.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal
        g.EnableHeadersVisualStyles = False
        g.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None
        g.ColumnHeadersDefaultCellStyle.BackColor = CardBack
        g.ColumnHeadersDefaultCellStyle.ForeColor = TextSub
        g.ColumnHeadersDefaultCellStyle.SelectionBackColor = CardBack
        g.ColumnHeadersDefaultCellStyle.Font = UiFont(9.0!)
        g.ColumnHeadersDefaultCellStyle.Padding = New Padding(4, 4, 4, 4)
        g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        g.ColumnHeadersHeight = 30
        g.DefaultCellStyle.Font = UiFont(10.0!)
        g.DefaultCellStyle.ForeColor = TextMain
        g.DefaultCellStyle.SelectionBackColor = SelectBack
        g.DefaultCellStyle.SelectionForeColor = SelectText
        g.DefaultCellStyle.Padding = New Padding(4, 0, 4, 0)
        g.RowTemplate.Height = 30
        g.RowHeadersVisible = False
        g.AllowUserToResizeRows = False
    End Sub

End Class
