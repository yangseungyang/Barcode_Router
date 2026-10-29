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
'             tblRuleCond    : 규칙의 조건 줄 (연결 AND/OR, 대상 = tblField 의 항목ID, 조건 LIKE/EQUAL, 입력값 - Enter 로 나눈 여러 줄)
'             --- 화면 설정 (2026-10-28 지시 : UI 를 유저가 설정, 하드코딩 하지 않음) ---
'             tblConfig      : 키 / 값 (PLASMA 서버, 글꼴, 설명서 경로, 로그 보관 개월, 카드 한 줄 개수)
'             tblField       : 받아오는 정보 항목 (PLASMA 항목명 / 다른 항목에서 잘라 만든 항목 / 입력 바코드)
'             tblCard        : 메인 화면 제품 정보 카드 (순서, 항목, 표시명)
'             tblBarcodeType : 바코드 종류 (길이 / 글자 종류 / 패턴 → 처리 : 분기 / 모두에 보냄 / 거부)
'             tblDestination : EndMessage(종료 신호), LineEnd(끝 문자 CR / LF / CRLF) 열 추가
'             - 프로그램에는 기본값을 넣지 않는다. 테이블이나 값이 없으면 빈 설정으로 동작하고 로그를 남긴다
'             - 새 테이블 / 열이 없는 예전 설정 파일은 저장할 때 빈 구조만 만든다 (값은 넣지 않음)
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
    Public ProcessName As String = ""           '프로세스명 또는 exe 전체 경로 (\ 가 있으면 경로로 찾음)
    Public EndTimeoutSec As Integer = 0         '0 = 무제한
    Public Enabled As Boolean = True
    Public EndMessage As String = ""            '검사 프로그램이 보내는 종료 신호 (예 : END). 비어 있으면 종료를 알 수 없음
    Public LineEnd As String = ""               '보낼 때 붙이는 끝 문자 CR / LF / CRLF (비어 있으면 붙이지 않음)

    Public Const LINE_CR As String = "CR"
    Public Const LINE_LF As String = "LF"
    Public Const LINE_CRLF As String = "CRLF"

    Public Function LineEndChars() As String
        Select Case LineEnd.Trim.ToUpperInvariant
            Case LINE_CR : Return vbCr
            Case LINE_LF : Return vbLf
            Case LINE_CRLF : Return vbCrLf
            Case Else : Return ""
        End Select
    End Function

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
    Public Const OP_LIKE As String = "LIKE"
    Public Const OP_EQUAL As String = "EQUAL"

    Public Const JOIN_AND As String = "AND"
    Public Const JOIN_OR As String = "OR"

    Public JoinOp As String = JOIN_AND          '앞 조건 줄과의 연결 (첫 줄은 쓰지 않음). AND 를 먼저 묶고 OR 로 나눈다
    Public Target As String = ""                '대상 항목 (tblField 의 FieldID)
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
        Return Priority.ToString & ". " & Description & If(Enabled, "", clsLang.L("  (사용 안 함)", "  (未使用)"))
    End Function
End Class

'받아오는 정보 항목 (tblField) ===============================================================================================
Public Class clsField
    Public Const SRC_PLASMA As String = "PLASMA"    'PLASMA 응답의 항목 (PlasmaTag 뒤 ~ 다음 $ 앞)
    Public Const SRC_DERIVE As String = "DERIVE"    '다른 항목(BaseField)에서 잘라 만듦 : 구분자 + 몇 번째, 또는 시작 위치 + 길이
    Public Const SRC_INPUT As String = "INPUT"      '찍은 바코드 그대로

    Public FieldID As String = ""
    Public DisplayName As String = ""
    Public SourceType As String = ""
    Public PlasmaTag As String = ""
    Public BaseField As String = ""
    Public Delimiter As String = ""
    Public PartNo As Integer = 0                    '구분자로 나눈 몇 번째 (1 부터)
    Public StartPos As Integer = 0                  '구분자가 없을 때 : 시작 위치 (1 부터)
    Public PartLen As Integer = 0                   '구분자가 없을 때 : 길이 (0 = 끝까지)
    Public Enabled As Boolean = True
    Public SortNo As Integer = 0

    Public ReadOnly Property Caption() As String
        Get
            Return If(DisplayName.Trim = "", FieldID, DisplayName)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return FieldID
    End Function
End Class

'메인 화면 카드 (tblCard) ==================================================================================================
Public Class clsCard
    Public CardNo As Integer
    Public FieldID As String = ""
    Public Caption As String = ""                   '비어 있으면 항목 표시명
    Public Enabled As Boolean = True
End Class

'바코드 종류 (tblBarcodeType) ==============================================================================================
Public Class clsBarcodeType
    Public Const KIND_ANY As String = "ANY"         '글자 종류 : 아무거나
    Public Const KIND_NUM As String = "NUM"         '숫자만
    Public Const KIND_ALPHA As String = "ALPHA"     '영문만
    Public Const KIND_ALNUM As String = "ALNUM"     '영문 / 숫자만

    Public Const ACT_ROUTE As String = "ROUTE"          'PLASMA 조회 → 분기 규칙 → 한 대상에 보냄 (제품 바코드)
    Public Const ACT_BROADCAST As String = "BROADCAST"  '사용 중인 모든 대상에 보냄 (작업자 바코드)
    Public Const ACT_REJECT As String = "REJECT"        '보내지 않음

    Public TypeNo As Integer
    Public TypeName As String = ""
    Public MinLen As Integer = 0                    '0 = 제한 없음
    Public MaxLen As Integer = 0                    '0 = 제한 없음
    Public CharKind As String = KIND_ANY
    Public Pattern As String = ""                   'VB Like 패턴 (비우면 보지 않음, 예 : A*)
    Public ActionType As String = ""
    Public Enabled As Boolean = True
End Class

'설정 읽기 / 저장 ==========================================================================================================
Public Class clsRouterSetting

    Private mDbPath As String
    Public EquipmentName As String = ""             '이 Router 가 있는 설비 이름 (메인 화면 제목줄 가운데 표시, 2026-10-26 추가)
    Public Reader As New clsReaderSetting
    Public Destinations As New List(Of clsDestination)
    Public Rules As New List(Of clsRouteRule)
    Public Config As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase)     'tblConfig 키 → 값
    Public ConfigDesc As New Dictionary(Of String, String)(StringComparer.OrdinalIgnoreCase) 'tblConfig 키 → 설명
    Public Fields As New List(Of clsField)
    Public Cards As New List(Of clsCard)
    Public BarcodeTypes As New List(Of clsBarcodeType)
    Public LoadNotes As New List(Of String)         '읽다가 없던 테이블 / 열 (로그용)

    'tblConfig 키
    Public Const CFG_PLASMA_HOST1 As String = "PlasmaHost1"
    Public Const CFG_PLASMA_HOST2 As String = "PlasmaHost2"
    Public Const CFG_PLASMA_PORT As String = "PlasmaPort"
    Public Const CFG_PLASMA_COMMAND As String = "PlasmaCommand"
    Public Const CFG_UI_FONT As String = "UiFont"
    Public Const CFG_UI_LANGUAGE As String = "UiLanguage"     'KO / JA (clsLang)
    Public Const CFG_MANUAL_PATH As String = "ManualPath"
    Public Const CFG_LOG_KEEP_MONTHS As String = "LogKeepMonths"
    Public Const CFG_CARDS_PER_ROW As String = "CardsPerRow"

    Public ReadOnly Property DbPath() As String
        Get
            Return mDbPath
        End Get
    End Property

    Public Sub New(ByVal dbPath As String)
        mDbPath = dbPath
    End Sub

    Public Function GetConfig(ByVal key As String) As String
        Dim v As String = ""
        If Config.TryGetValue(key, v) Then Return v.Trim
        Return ""
    End Function

    Public Function GetConfigInt(ByVal key As String) As Integer
        Dim n As Integer = 0
        Integer.TryParse(GetConfig(key), n)
        Return n
    End Function

    Public Function FindField(ByVal fieldID As String) As clsField
        For Each f As clsField In Fields
            If String.Equals(f.FieldID, fieldID, StringComparison.OrdinalIgnoreCase) Then Return f
        Next
        Return Nothing
    End Function

    Public Function EnabledFields() As List(Of clsField)
        Return Fields.Where(Function(x) x.Enabled).OrderBy(Function(x) x.SortNo).ToList
    End Function

    'MDB 엔진 (2026-10-29) : Router 는 32bit(x86) 로 빌드하고, Windows 에 기본으로 있는 Jet 을 먼저 쓴다
    '  64bit 빌드는 Access Database Engine x64 가 설치된 PC 에서만 MDB 를 읽어서, 다른 PC 에서 '공급자 등록 안 됨' 오류가 났다
    '  위에서부터 써 보고 등록된 첫 엔진을 쓴다 (FAM3 내전압 / CPU검사장치 검사 프로그램과 같은 Jet 방식)
    Private Shared ReadOnly PROVIDERS() As String = {"Microsoft.Jet.OLEDB.4.0", "Microsoft.ACE.OLEDB.12.0", "Microsoft.ACE.OLEDB.16.0"}
    Private Shared mProvider As String = ""

    Public Shared ReadOnly Property ProviderName() As String
        Get
            Return mProvider
        End Get
    End Property

    Private Function Conn() As OleDbConnection
        If mProvider = "" Then mProvider = FindProvider(mDbPath)
        Return New OleDbConnection("Provider=" & mProvider & ";Data Source=" & mDbPath)
    End Function

    Private Shared Function FindProvider(ByVal dbPath As String) As String
        '==========================================================================================
        'Function : 이 PC 에 등록된 MDB 엔진 찾기 (PROVIDERS 순서). 등록 안 됨(InvalidOperationException)이면 다음 것
        '           파일 오류 등 다른 오류는 엔진은 있다는 뜻이므로 그 엔진을 쓴다 (오류는 실제로 읽을 때 그대로 나옴)
        '==========================================================================================
        Dim tried As New List(Of String)
        For Each p As String In PROVIDERS
            Try
                Using cn As New OleDbConnection("Provider=" & p & ";Data Source=" & dbPath)
                    cn.Open()
                End Using
                Return p
            Catch ex As InvalidOperationException
                tried.Add(p)                '이 PC 에 등록되지 않은 엔진
            Catch
                Return p
            End Try
        Next
        Throw New ApplicationException(clsLang.L("MDB 엔진이 없습니다 (" & If(Environment.Is64BitProcess, "64bit", "32bit") & " 프로그램) : " & String.Join(", ", tried) & vbCrLf &
                                                 "32bit 프로그램은 Windows 기본 Jet 엔진을 씁니다. Router 가 32bit(x86) 로 빌드되었는지 확인해 주세요.",
                                                 "MDBエンジンがありません (" & If(Environment.Is64BitProcess, "64bit", "32bit") & " プログラム) : " & String.Join(", ", tried) & vbCrLf &
                                                 "32bitプログラムはWindows標準のJetエンジンを使います。Routerが32bit(x86)でビルドされているか確認してください。"))
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
        If Not IO.File.Exists(mDbPath) Then Throw New ApplicationException(clsLang.L("설정 파일이 없습니다 : ", "設定ファイルがありません : ") & mDbPath)

        Reader = New clsReaderSetting
        Destinations = New List(Of clsDestination)
        Rules = New List(Of clsRouteRule)
        Config.Clear()
        ConfigDesc.Clear()
        Fields = New List(Of clsField)
        Cards = New List(Of clsCard)
        BarcodeTypes = New List(Of clsBarcodeType)
        LoadNotes = New List(Of String)

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

            Dim hasEndCols As Boolean = HasColumn(cn, "tblDestination", "EndMessage") AndAlso HasColumn(cn, "tblDestination", "LineEnd")
            If Not hasEndCols Then LoadNotes.Add(clsLang.L("tblDestination.EndMessage / LineEnd 없음 (빈 값)", "tblDestination.EndMessage / LineEnd なし (空欄)"))
            Using cmd As New OleDbCommand("SELECT * FROM tblDestination ORDER BY DestID", cn)
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
                        If hasEndCols Then
                            d.EndMessage = S(r("EndMessage"))
                            d.LineEnd = S(r("LineEnd"))
                        End If
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

            Call LoadDisplay(cn)
        End Using

    End Sub

    Private Sub LoadDisplay(ByVal cn As OleDbConnection)
        '==========================================================================================
        'Function : 화면 설정 읽기 (tblConfig, tblField, tblCard, tblBarcodeType) - 테이블이 없으면 빈 설정
        '==========================================================================================
        If HasTable(cn, "tblConfig") Then
            Using cmd As New OleDbCommand("SELECT ConfigKey, ConfigValue, Description FROM tblConfig", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Dim k As String = S(r("ConfigKey")).Trim
                        If k = "" Then Continue While
                        Config(k) = S(r("ConfigValue"))
                        ConfigDesc(k) = S(r("Description"))
                    End While
                End Using
            End Using
        Else
            LoadNotes.Add(clsLang.L("tblConfig 없음", "tblConfig なし"))
        End If

        If HasTable(cn, "tblField") Then
            Using cmd As New OleDbCommand("SELECT * FROM tblField ORDER BY SortNo", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Dim f As New clsField
                        f.FieldID = S(r("FieldID")).Trim
                        If f.FieldID = "" Then Continue While
                        f.DisplayName = S(r("DisplayName"))
                        f.SourceType = S(r("SourceType")).Trim.ToUpperInvariant
                        f.PlasmaTag = S(r("PlasmaTag"))
                        f.BaseField = S(r("BaseField")).Trim
                        f.Delimiter = S(r("Delimiter"))
                        f.PartNo = N(r("PartNo"))
                        f.StartPos = N(r("StartPos"))
                        f.PartLen = N(r("PartLen"))
                        f.Enabled = B(r("Enabled"))
                        f.SortNo = N(r("SortNo"))
                        Fields.Add(f)
                    End While
                End Using
            End Using
        Else
            LoadNotes.Add(clsLang.L("tblField 없음", "tblField なし"))
        End If

        If HasTable(cn, "tblCard") Then
            Using cmd As New OleDbCommand("SELECT CardNo, FieldID, Caption, Enabled FROM tblCard ORDER BY CardNo", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Cards.Add(New clsCard With {.CardNo = N(r("CardNo")), .FieldID = S(r("FieldID")).Trim, .Caption = S(r("Caption")), .Enabled = B(r("Enabled"))})
                    End While
                End Using
            End Using
        Else
            LoadNotes.Add(clsLang.L("tblCard 없음", "tblCard なし"))
        End If

        If HasTable(cn, "tblBarcodeType") Then
            Using cmd As New OleDbCommand("SELECT * FROM tblBarcodeType ORDER BY TypeNo", cn)
                Using r As OleDbDataReader = cmd.ExecuteReader()
                    While r.Read()
                        Dim t As New clsBarcodeType
                        t.TypeNo = N(r("TypeNo"))
                        t.TypeName = S(r("TypeName"))
                        t.MinLen = N(r("MinLen"))
                        t.MaxLen = N(r("MaxLen"))
                        t.CharKind = S(r("CharKind")).Trim.ToUpperInvariant
                        t.Pattern = S(r("Pattern"))
                        t.ActionType = S(r("ActionType")).Trim.ToUpperInvariant
                        t.Enabled = B(r("Enabled"))
                        BarcodeTypes.Add(t)
                    End While
                End Using
            End Using
        Else
            LoadNotes.Add(clsLang.L("tblBarcodeType 없음", "tblBarcodeType なし"))
        End If
    End Sub

    Public Sub SaveDisplay()
        '==========================================================================================
        'Function : 화면 설정 저장 (tblConfig, tblField, tblCard, tblBarcodeType)
        '           테이블이 없는 예전 설정 파일이면 빈 구조만 만든 뒤 화면 내용을 넣는다
        '==========================================================================================
        Using cn As OleDbConnection = Conn()
            cn.Open()
            Call EnsureDisplaySchema(cn)
            Using tr As OleDbTransaction = cn.BeginTransaction()
                Exec(cn, tr, "DELETE FROM tblConfig")
                For Each k As String In Config.Keys
                    Dim desc As String = ""
                    ConfigDesc.TryGetValue(k, desc)
                    Exec(cn, tr, "INSERT INTO tblConfig (ConfigKey, ConfigValue, Description) VALUES (?,?,?)", k, Config(k), desc)
                Next

                Exec(cn, tr, "DELETE FROM tblField")
                For Each f As clsField In Fields
                    Exec(cn, tr, "INSERT INTO tblField (FieldID, DisplayName, SourceType, PlasmaTag, BaseField, Delimiter, PartNo, StartPos, PartLen, Enabled, SortNo) VALUES (?,?,?,?,?,?,?,?,?,?,?)",
                         f.FieldID, f.DisplayName, f.SourceType, f.PlasmaTag, f.BaseField, f.Delimiter, f.PartNo, f.StartPos, f.PartLen, f.Enabled, f.SortNo)
                Next

                Exec(cn, tr, "DELETE FROM tblCard")
                For Each c As clsCard In Cards
                    Exec(cn, tr, "INSERT INTO tblCard (CardNo, FieldID, Caption, Enabled) VALUES (?,?,?,?)", c.CardNo, c.FieldID, c.Caption, c.Enabled)
                Next

                Exec(cn, tr, "DELETE FROM tblBarcodeType")
                For Each t As clsBarcodeType In BarcodeTypes
                    Exec(cn, tr, "INSERT INTO tblBarcodeType (TypeNo, TypeName, MinLen, MaxLen, CharKind, Pattern, ActionType, Enabled) VALUES (?,?,?,?,?,?,?,?)",
                         t.TypeNo, t.TypeName, t.MinLen, t.MaxLen, t.CharKind, t.Pattern, t.ActionType, t.Enabled)
                Next
                tr.Commit()
            End Using
        End Using
    End Sub

    Private Shared Sub EnsureDisplaySchema(ByVal cn As OleDbConnection)
        '빈 구조만 만든다 (값은 넣지 않음 - 프로그램에 기본값을 두지 않는다)
        If Not HasTable(cn, "tblConfig") Then DDL(cn, "CREATE TABLE tblConfig (ConfigKey TEXT(50), ConfigValue TEXT(255), Description TEXT(255))")
        If Not HasTable(cn, "tblField") Then DDL(cn, "CREATE TABLE tblField (FieldID TEXT(30), DisplayName TEXT(50), SourceType TEXT(10), PlasmaTag TEXT(50), BaseField TEXT(30), Delimiter TEXT(10), PartNo INTEGER, StartPos INTEGER, PartLen INTEGER, Enabled YESNO, SortNo INTEGER)")
        If Not HasTable(cn, "tblCard") Then DDL(cn, "CREATE TABLE tblCard (CardNo INTEGER, FieldID TEXT(30), Caption TEXT(50), Enabled YESNO)")
        If Not HasTable(cn, "tblBarcodeType") Then DDL(cn, "CREATE TABLE tblBarcodeType (TypeNo INTEGER, TypeName TEXT(50), MinLen INTEGER, MaxLen INTEGER, CharKind TEXT(10), Pattern TEXT(100), ActionType TEXT(10), Enabled YESNO)")
    End Sub

    Private Shared Sub EnsureDestinationColumns(ByVal cn As OleDbConnection)
        If Not HasColumn(cn, "tblDestination", "EndMessage") Then DDL(cn, "ALTER TABLE tblDestination ADD COLUMN EndMessage TEXT(20)")
        If Not HasColumn(cn, "tblDestination", "LineEnd") Then DDL(cn, "ALTER TABLE tblDestination ADD COLUMN LineEnd TEXT(10)")
    End Sub

    Private Shared Function HasTable(ByVal cn As OleDbConnection, ByVal table As String) As Boolean
        Return cn.GetSchema("Tables", New String() {Nothing, Nothing, table, "TABLE"}).Rows.Count > 0
    End Function

    Private Shared Function HasColumn(ByVal cn As OleDbConnection, ByVal table As String, ByVal column As String) As Boolean
        Return cn.GetSchema("Columns", New String() {Nothing, Nothing, table, column}).Rows.Count > 0
    End Function

    Private Shared Sub DDL(ByVal cn As OleDbConnection, ByVal sql As String)
        Using cmd As New OleDbCommand(sql, cn)
            cmd.ExecuteNonQuery()
        End Using
    End Sub

    Public Sub SaveReaderAndDestinations()
        '==========================================================================================
        'Function : COM 설정 화면 내용 저장 (tblReader, tblDestination)
        '==========================================================================================
        Using cn As OleDbConnection = Conn()
            cn.Open()
            Call EnsureEquipmentTable(cn)
            Call EnsureDestinationColumns(cn)
            Using tr As OleDbTransaction = cn.BeginTransaction()
                Exec(cn, tr, "DELETE FROM tblEquipment")
                Exec(cn, tr, "INSERT INTO tblEquipment (EquipmentName) VALUES (?)", EquipmentName)

                Exec(cn, tr, "DELETE FROM tblReader")
                Exec(cn, tr, "INSERT INTO tblReader (InputMode, Port, Setting) VALUES (?,?,?)", Reader.InputMode, Reader.Port, Reader.Setting)

                Exec(cn, tr, "DELETE FROM tblDestination")
                For Each d As clsDestination In Destinations
                    Exec(cn, tr, "INSERT INTO tblDestination (DestID, DestName, Port, Setting, ProcessName, EndTimeoutSec, Enabled, EndMessage, LineEnd) VALUES (?,?,?,?,?,?,?,?,?)",
                         d.DestID, d.DestName, d.Port, d.Setting, d.ProcessName, d.EndTimeoutSec, d.Enabled, d.EndMessage, d.LineEnd)
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
