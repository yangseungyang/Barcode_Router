Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterSetting
'Function   : 바코드 분배 프로그램 설정 읽기 / 저장 (Files\RouterSetting.mdb)
'Framwork   : VB2022 Framework4.8
'Memo       : 설정 MDB 는 기존 검사 프로그램과 같은 ACE OLEDB 12.0 으로 읽고 쓴다 (구조설계 4_설정데이터)
'             tblEquipment   : 이 Router 가 있는 설비 이름 (메인 화면 제목줄 가운데 표시)
'             tblReader      : 바코드리더 입력 방식(HID / COM), 포트, 통신설정
'             tblDestination : 분기 대상 (이름, Router 쪽 COM, 통신설정, 프로세스명, END 대기 초, 사용)
'             tblRule        : 분기 규칙 (우선순위, 사용, 분기 대상, 설명)
'             tblRuleCond    : 규칙의 조건 줄 (연결 AND/OR, 대상 MSCODE/MODEL, 조건 LIKE/EQUAL, 입력값 - Enter 로 나눈 여러 줄)
'             저장은 테이블 내용을 지우고 화면의 내용으로 다시 넣는다
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-3항, 7-10항 / 구조설계 R-13, R-14
'============================================================================================================================
Imports System.Data.OleDb

'바코드리더 설정 ===========================================================================================================
Public Class clsReaderSetting
    Public Const MODE_HID As String = "HID"
    Public Const MODE_COM As String = "COM"
    Public InputMode As String = MODE_HID
    Public Port As String = ""
    Public Setting As String = "9600-N-8-1"
End Class

'분기 대상 =================================================================================================================
Public Class clsDestination
    Public DestID As Integer
    Public DestName As String = ""
    Public Port As String = ""
    Public Setting As String = "9600-N-8-1"
    Public ProcessName As String = ""
    Public EndTimeoutSec As Integer = 0         '0 = 무제한
    Public Enabled As Boolean = True

    Public ReadOnly Property DisplayName() As String
        Get
            Return Port & "  " & DestName
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return DisplayName
    End Function
End Class

'규칙의 조건 줄 ============================================================================================================
Public Class clsRuleCond
    Public Const TARGET_MSCODE As String = "MSCODE"
    Public Const TARGET_MODEL As String = "MODEL"
    Public Const OP_LIKE As String = "LIKE"
    Public Const OP_EQUAL As String = "EQUAL"

    Public Const JOIN_AND As String = "AND"
    Public Const JOIN_OR As String = "OR"

    Public JoinOp As String = JOIN_AND          '앞 조건 줄과의 연결 (첫 줄은 쓰지 않음). AND 를 먼저 묶고 OR 로 나눈다
    Public Target As String = TARGET_MSCODE
    Public [Operator] As String = OP_LIKE
    Public InputValues As String = ""           'Enter(CRLF) 로 나눈 값 그대로

    Public Function ValueList() As List(Of String)
        '==========================================================================================
        'Function : 입력값을 줄 단위로 나눠 돌려준다 (빈 줄 제외, 앞뒤 공백 제거)
        '==========================================================================================
        Dim lst As New List(Of String)
        For Each s As String In InputValues.Replace(vbCrLf, vbLf).Replace(vbCr, vbLf).Split(CChar(vbLf))
            If s.Trim <> "" Then lst.Add(s.Trim)
        Next
        Return lst
    End Function
End Class

'분기 규칙 =================================================================================================================
Public Class clsRouteRule
    Public RuleID As Integer
    Public Priority As Integer = 1
    Public Enabled As Boolean = True
    Public DestID As Integer
    Public Description As String = ""
    Public Conds As New List(Of clsRuleCond)

    Public Overrides Function ToString() As String
        Return Priority.ToString & ". " & Description & If(Enabled, "", "  (사용 안 함)")
    End Function
End Class

'설정 읽기 / 저장 ==========================================================================================================
Public Class clsRouterSetting

    Private mDbPath As String
    Public EquipmentName As String = ""             '이 Router 가 있는 설비 이름 (메인 화면 제목줄 가운데 표시, 2026-10-26 추가)
    Public Reader As New clsReaderSetting
    Public Destinations As New List(Of clsDestination)
    Public Rules As New List(Of clsRouteRule)

    Public Sub New(ByVal dbPath As String)
        mDbPath = dbPath
    End Sub

    Private Function Conn() As OleDbConnection
        Return New OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & mDbPath)
    End Function

    Private Shared Function S(ByVal v As Object) As String
        If v Is Nothing OrElse IsDBNull(v) Then Return ""
        Return v.ToString
    End Function

    Private Shared Function N(ByVal v As Object) As Integer
        If v Is Nothing OrElse IsDBNull(v) Then Return 0
        Return CInt(v)
    End Function

    Private Shared Function B(ByVal v As Object) As Boolean
        If v Is Nothing OrElse IsDBNull(v) Then Return False
        Return CBool(v)
    End Function

    Public Function FindDestination(ByVal destID As Integer) As clsDestination
        For Each d As clsDestination In Destinations
            If d.DestID = destID Then Return d
        Next
        Return Nothing
    End Function

    Public Sub Load()
        '==========================================================================================
        'Function : RouterSetting.mdb 에서 모든 설정을 읽는다
        '==========================================================================================
        If Not IO.File.Exists(mDbPath) Then Throw New ApplicationException("설정 파일이 없습니다 : " & mDbPath)

        Reader = New clsReaderSetting
        Destinations = New List(Of clsDestination)
        Rules = New List(Of clsRouteRule)

        EquipmentName = ""

        Using cn As OleDbConnection = Conn()
            cn.Open()

            Call EnsureEquipmentTable(cn)
            Using cmd As New OleDbCommand("SELECT EquipmentName FROM tblEquipment", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    If r.Read() Then EquipmentName = S(r("EquipmentName"))
                End Using
            End Using

            Using cmd As New OleDbCommand("SELECT InputMode, Port, Setting FROM tblReader", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    If r.Read() Then
                        Reader.InputMode = S(r("InputMode"))
                        Reader.Port = S(r("Port"))
                        Reader.Setting = S(r("Setting"))
                    End If
                End Using
            End Using

            Using cmd As New OleDbCommand("SELECT DestID, DestName, Port, Setting, ProcessName, EndTimeoutSec, Enabled FROM tblDestination ORDER BY DestID", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Dim d As New clsDestination
                        d.DestID = N(r("DestID"))
                        d.DestName = S(r("DestName"))
                        d.Port = S(r("Port"))
                        d.Setting = S(r("Setting"))
                        d.ProcessName = S(r("ProcessName"))
                        d.EndTimeoutSec = N(r("EndTimeoutSec"))
                        d.Enabled = B(r("Enabled"))
                        Destinations.Add(d)
                    End While
                End Using
            End Using

            Using cmd As New OleDbCommand("SELECT RuleID, Priority, Enabled, DestID, Description FROM tblRule ORDER BY Priority, RuleID", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Dim ru As New clsRouteRule
                        ru.RuleID = N(r("RuleID"))
                        ru.Priority = N(r("Priority"))
                        ru.Enabled = B(r("Enabled"))
                        ru.DestID = N(r("DestID"))
                        ru.Description = S(r("Description"))
                        Rules.Add(ru)
                    End While
                End Using
            End Using

            Call EnsureJoinColumn(cn)
            Using cmd As New OleDbCommand("SELECT RuleID, CondNo, JoinOp, Target, Operator, InputValues FROM tblRuleCond ORDER BY RuleID, CondNo", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Dim id As Integer = N(r("RuleID"))
                        Dim c As New clsRuleCond
                        c.JoinOp = If(S(r("JoinOp")).ToUpperInvariant = clsRuleCond.JOIN_OR, clsRuleCond.JOIN_OR, clsRuleCond.JOIN_AND)
                        c.Target = S(r("Target"))
                        c.Operator = S(r("Operator"))
                        c.InputValues = S(r("InputValues"))
                        For Each ru As clsRouteRule In Rules
                            If ru.RuleID = id Then ru.Conds.Add(c) : Exit For
                        Next
                    End While
                End Using
            End Using
        End Using

    End Sub

    Public Sub SaveReaderAndDestinations()
        '==========================================================================================
        'Function : COM 설정 화면 내용 저장 (tblReader, tblDestination)
        '==========================================================================================
        Using cn As OleDbConnection = Conn()
            cn.Open()
            Call EnsureEquipmentTable(cn)
            Using tr As OleDbTransaction = cn.BeginTransaction()
                Exec(cn, tr, "DELETE FROM tblEquipment")
                Exec(cn, tr, "INSERT INTO tblEquipment (EquipmentName) VALUES (?)", EquipmentName)

                Exec(cn, tr, "DELETE FROM tblReader")
                Exec(cn, tr, "INSERT INTO tblReader (InputMode, Port, Setting) VALUES (?,?,?)", Reader.InputMode, Reader.Port, Reader.Setting)

                Exec(cn, tr, "DELETE FROM tblDestination")
                For Each d As clsDestination In Destinations
                    Exec(cn, tr, "INSERT INTO tblDestination (DestID, DestName, Port, Setting, ProcessName, EndTimeoutSec, Enabled) VALUES (?,?,?,?,?,?,?)",
                         d.DestID, d.DestName, d.Port, d.Setting, d.ProcessName, d.EndTimeoutSec, d.Enabled)
                Next
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Shared Sub EnsureEquipmentTable(ByVal cn As OleDbConnection)
        '==========================================================================================
        'Function : tblEquipment(설비명) 테이블이 없으면 만든다 (2026-10-26 추가, 예전 설정 파일도 그대로 사용)
        '==========================================================================================
        Dim t As DataTable = cn.GetSchema("Tables", New String() {Nothing, Nothing, "tblEquipment", "TABLE"})
        If t.Rows.Count > 0 Then Exit Sub
        Using cmd As New OleDbCommand("CREATE TABLE tblEquipment (EquipmentName TEXT(100))", cn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Private Shared Sub EnsureJoinColumn(ByVal cn As OleDbConnection)
        '==========================================================================================
        'Function : tblRuleCond 에 JoinOp(AND / OR) 컬럼이 없으면 추가 (조건 AND / OR 선택 기능, 2026-10-26)
        '           예전 설정 파일도 그대로 쓸 수 있게 한다 (없던 줄은 AND)
        '==========================================================================================
        Dim cols As DataTable = cn.GetSchema("Columns", New String() {Nothing, Nothing, "tblRuleCond", "JoinOp"})
        If cols.Rows.Count > 0 Then Exit Sub
        Using cmd As New OleDbCommand("ALTER TABLE tblRuleCond ADD COLUMN JoinOp TEXT(5)", cn)
            cmd.ExecuteNonQuery()
        End Using
        Using cmd As New OleDbCommand("UPDATE tblRuleCond SET JoinOp = 'AND'", cn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub SaveRules()
        '==========================================================================================
        'Function : 분기 조건 설정 화면 내용 저장 (tblRule, tblRuleCond)
        '==========================================================================================
        Using cn As OleDbConnection = Conn()
            cn.Open()
            Using tr As OleDbTransaction = cn.BeginTransaction()
                Exec(cn, tr, "DELETE FROM tblRuleCond")
                Exec(cn, tr, "DELETE FROM tblRule")
                For Each ru As clsRouteRule In Rules
                    Exec(cn, tr, "INSERT INTO tblRule (RuleID, Priority, Enabled, DestID, Description) VALUES (?,?,?,?,?)",
                         ru.RuleID, ru.Priority, ru.Enabled, ru.DestID, ru.Description)
                    Dim no As Integer = 1
                    For Each c As clsRuleCond In ru.Conds
                        Exec(cn, tr, "INSERT INTO tblRuleCond (RuleID, CondNo, JoinOp, Target, Operator, InputValues) VALUES (?,?,?,?,?,?)",
                             ru.RuleID, no, c.JoinOp, c.Target, c.Operator, c.InputValues)
                        no += 1
                    Next
                Next
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Shared Sub Exec(ByVal cn As OleDbConnection, ByVal tr As OleDbTransaction, ByVal sql As String, ByVal ParamArray args() As Object)
        Using cmd As New OleDbCommand(sql, cn, tr)
            For i As Integer = 0 To args.Length - 1
                cmd.Parameters.AddWithValue("p" & i.ToString, If(args(i), DBNull.Value))
            Next
            cmd.ExecuteNonQuery()
        End Using
    End Sub

End Class
