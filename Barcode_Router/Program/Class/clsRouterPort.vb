Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : clsRouterPort
'Function   : Router 의 COM 포트 1개 (바코드리더 COM 또는 분기 대상 COM)
'Framwork   : VB2022 Framework4.8
'Memo       : 통신 규약 (구조설계 5_통신규약)
'             - 9600-N-8-1 형식의 통신설정, 종단 CR (FAM3 cls101_ComRS232C 기존 설정과 같음)
'             - 송신 : 문자열 + 끝 문자 (분기 대상별 설정 CR / LF / CRLF, 2026-10-28 - CPU검사장치 A8 은 LF 가 필요)
'             - 수신 : CR 또는 LF 까지 모아서 한 줄이 되면 LineReceived 이벤트 (수신 스레드에서 발생하므로 화면은 Invoke 로 처리)
'Issue      : 2026/10/26  Rev 0.00    YSH       New Release
'             근거 : FAM3 증산 대응_내전압 설비 ADD-ON.md Rev 3  7-4항, 7-9항, 7-10항
'============================================================================================================================

Public Class clsRouterPort

    Private WithEvents mPort As New IO.Ports.SerialPort
    Private mBuf As New System.Text.StringBuilder
    Private ReadOnly mLock As New Object

    Public Property Tag As Object               '분기 대상(clsDestination) 등 호출 쪽에서 쓰는 값

    Public Event LineReceived(ByVal sender As clsRouterPort, ByVal line As String)

    Public ReadOnly Property PortName() As String
        Get
            Return mPort.PortName
        End Get
    End Property

    Public ReadOnly Property IsOpen() As Boolean
        Get
            Return mPort.IsOpen
        End Get
    End Property

    Public Sub Open(ByVal portName As String, ByVal setting As String)
        '==========================================================================================
        'Function : 포트 열기   setting 예 : 9600-N-8-1
        '==========================================================================================
        Dim ps() As String = setting.Trim.Split("-"c)
        If ps.Length < 4 Then Throw New ApplicationException(clsLang.L("통신설정 형식이 잘못되었습니다 (예 : 9600-N-8-1) : ", "通信設定の形式が正しくありません (例 : 9600-N-8-1) : ") & setting)

        Close()
        mPort.PortName = portName.Trim
        mPort.BaudRate = CInt(ps(0))
        Select Case ps(1).ToUpperInvariant
            Case "E" : mPort.Parity = IO.Ports.Parity.Even
            Case "M" : mPort.Parity = IO.Ports.Parity.Mark
            Case "O" : mPort.Parity = IO.Ports.Parity.Odd
            Case "S" : mPort.Parity = IO.Ports.Parity.Space
            Case Else : mPort.Parity = IO.Ports.Parity.None
        End Select
        mPort.DataBits = CInt(ps(2))
        Select Case ps(3)
            Case "1.5" : mPort.StopBits = IO.Ports.StopBits.OnePointFive
            Case "2" : mPort.StopBits = IO.Ports.StopBits.Two
            Case Else : mPort.StopBits = IO.Ports.StopBits.One
        End Select
        mPort.Handshake = IO.Ports.Handshake.None
        mPort.RtsEnable = True
        mPort.ReceivedBytesThreshold = 1
        mPort.WriteTimeout = 3000
        mPort.Open()
        mPort.DiscardInBuffer()
        mPort.DiscardOutBuffer()
        SyncLock mLock
            mBuf.Clear()
        End SyncLock
    End Sub

    Public Sub Close()
        Try
            If mPort.IsOpen Then mPort.Close()
        Catch ex As Exception
            Debug.Print("Port close NG : " & ex.Message)
        End Try
    End Sub

    Public Sub SendLine(ByVal data As String, ByVal lineEnd As String)
        If Not mPort.IsOpen Then Throw New ApplicationException(mPort.PortName & clsLang.L(" 포트가 열려 있지 않습니다.", " ポートが開いていません。"))
        mPort.Write(data & lineEnd)
    End Sub

    Private Sub mPort_DataReceived(ByVal sender As Object, ByVal e As IO.Ports.SerialDataReceivedEventArgs) Handles mPort.DataReceived
        Dim lines As New List(Of String)
        Try
            SyncLock mLock
                mBuf.Append(mPort.ReadExisting())
                Do
                    Dim s As String = mBuf.ToString
                    Dim p As Integer = s.IndexOfAny(New Char() {CChar(vbCr), CChar(vbLf)})     'CR / LF / CRLF 어느 것이든 한 줄 (CRLF 의 LF 는 빈 줄이 되어 버려진다)
                    If p < 0 Then Exit Do
                    lines.Add(s.Substring(0, p).Trim)
                    mBuf.Remove(0, p + 1)
                Loop
            End SyncLock
        Catch ex As Exception
            Debug.Print("Port receive NG : " & ex.Message)
        End Try

        For Each l As String In lines
            If l <> "" Then RaiseEvent LineReceived(Me, l)
        Next
    End Sub

End Class
