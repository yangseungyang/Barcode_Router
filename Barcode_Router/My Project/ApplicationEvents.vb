Option Strict On : Option Explicit On
'============================================================================================================================
'Class Name : MyApplication (ApplicationEvents)
'Function   : 프로그램 시작 처리 - exe 하나만 있어도 실행되도록 EtherComN.dll 을 exe 안에서 읽는다
'Framwork   : VB2022 Framework4.8
'Memo       : EtherComN.dll 이 exe 옆에 있으면 그 파일을 쓰고 (기존과 같음),
'             없을 때만 AssemblyResolve 로 exe 에 넣어 둔 EtherComN.dll 을 읽는다
'Issue      : 2026/10/28  Rev 0.00    YSH       New Release
'             근거 : Barcode_Router 연동 공통 규약.md 7항 (exe 단독 실행)
'============================================================================================================================

Namespace My

    Partial Friend Class MyApplication

        Private Const ETHERCOM_RESOURCE As String = "EtherComN.dll"

        Private Sub MyApplication_Startup(ByVal sender As Object, ByVal e As ApplicationServices.StartupEventArgs) Handles Me.Startup
            AddHandler AppDomain.CurrentDomain.AssemblyResolve, AddressOf ResolveEmbedded
        End Sub

        Private Shared Function ResolveEmbedded(ByVal sender As Object, ByVal args As ResolveEventArgs) As Reflection.Assembly
            '==========================================================================================
            'Function : exe 옆에서 찾지 못한 EtherComN 을 exe 내장 리소스에서 읽는다
            '==========================================================================================
            If Not args.Name.StartsWith("EtherComN,", StringComparison.OrdinalIgnoreCase) Then Return Nothing

            Using s As IO.Stream = GetType(MyApplication).Assembly.GetManifestResourceStream(ETHERCOM_RESOURCE)
                If s Is Nothing Then Return Nothing
                Dim buf(CInt(s.Length) - 1) As Byte
                s.Read(buf, 0, buf.Length)
                Return Reflection.Assembly.Load(buf)
            End Using
        End Function

    End Class

End Namespace
