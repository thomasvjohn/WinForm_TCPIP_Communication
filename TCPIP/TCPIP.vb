Imports System.Net.Sockets
Imports System.Threading
Imports System.Net
Imports System.Text

Public Class TCPIP
    ' Event data send back to calling form
    Public Event Datareceived(ByVal txt As String)

    ' connection status back to form True: ok
    Public Event connection(ByVal cStatus As Boolean)

    ' data send successfull (True)
    Public Event sendOK(ByVal sStatus As Boolean)

    ' data receive successfull (True)
    Public Event recOK(ByVal sReceive As Boolean)

    Private serverRuns As Boolean
    Private server As TcpListener
    Private sc As SynchronizationContext
    Private isConnected, receiveStatus, sendStatus As Boolean
    Private iRemote, pLocal As EndPoint

    ' reads endpoints
    Public ReadOnly Property Remote() As EndPoint
        Get
            Return iRemote
        End Get
    End Property

    ' reads local point
    Public ReadOnly Property Local() As EndPoint
        Get
            Return pLocal
        End Get
    End Property

    ' TCP connect with server
    Public Sub connect(ByVal hostAdress As String, ByVal hostPort As Integer)

        sc = SynchronizationContext.Current

        Try
            server = New TcpListener(IPAddress.Parse(hostAdress), hostPort)
        Catch ex As Exception
            MsgBox("server create: " & ex.Message, MsgBoxStyle.Exclamation)
        End Try

        Try
            With server
                .Start()
                .BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAccept), server)
                isConnected = True
            End With
        Catch ex As Exception
            MsgBox("server listen: " & ex.Message, MsgBoxStyle.Exclamation)
            isConnected = False
        Finally
            RaiseEvent connection(isConnected)
        End Try

    End Sub

    ' disConnect server
    Public Sub disconnect()
        Try
            isConnected = False
            server.Stop()
        Catch ex As Exception
            MsgBox("disConnect server: " & ex.Message, MsgBoxStyle.Exclamation)
            isConnected = True
        Finally
            RaiseEvent connection(isConnected)
        End Try
    End Sub

    ' TCP send data
    Public Sub SendData(ByVal txt As String, ByVal remoteAddress As String, ByVal remotePort As Integer)

        Dim clientSocket = New TcpClient
        Dim iP As IPAddress = IPAddress.Any
        Dim isIp As Boolean = IPAddress.TryParse(remoteAddress, iP)

        With clientSocket
            Try
                If isIp Then    ' ip address
                    .Connect(IPAddress.Parse(remoteAddress), remotePort)
                Else            ' DNS name
                    .Connect(remoteAddress, remotePort)
                End If

                Dim data() As Byte = Encoding.ASCII.GetBytes(txt)
                .NoDelay = True
                .GetStream().Write(data, 0, data.Length)
                .GetStream().Close()

                .Close()
                sendStatus = True
            Catch ex As Exception
                MsgBox("sendData: " & ex.Message, MsgBoxStyle.Exclamation)
                sendStatus = False
            Finally
                RaiseEvent sendOK(sendStatus)
            End Try

        End With
    End Sub

    ' TCP asynchronous receive on secondary thread
    Private Sub DoAccept(ByVal ar As IAsyncResult)

        Dim sb As New StringBuilder
        Dim buf() As Byte
        Dim datalen As Integer

        Dim listener As TcpListener
        Dim clientSocket As TcpClient
        If Not isConnected Then Exit Sub
        Try
            listener = CType(ar.AsyncState, TcpListener)
            clientSocket = listener.EndAcceptTcpClient(ar)
            clientSocket.ReceiveTimeout = 5000
            iRemote = clientSocket.Client.RemoteEndPoint
            pLocal = clientSocket.Client.LocalEndPoint

        Catch ex As ObjectDisposedException
            MsgBox("DoAccept ObjectDisposedException " & ex.Message, MsgBoxStyle.Exclamation)
            ' after server.stop() AsyncCallback is also active, but the object server is disposed
            Exit Sub
        End Try

        Try
            With clientSocket
                datalen = 0
                ' somtimes it occurs that .available returns the value 0 also data in buffer exists
                While datalen = 0
                    ' data in read Buffer
                    datalen = .Available
                End While
                buf = New Byte(datalen - 1) {}
                'get entire bytes at once
                .GetStream().Read(buf, 0, buf.Length)
                sb.Append(Encoding.ASCII.GetString(buf, 0, buf.Length))
                .Close()
            End With
            receiveStatus = True
        Catch ex As TimeoutException
            MsgBox("doAcceptData timeout: " & ex.Message, MsgBoxStyle.Exclamation)
            receiveStatus = False
            clientSocket.Close()
            Exit Sub
        Catch ex As Exception
            MsgBox("doAcceptData: " & ex.Message, MsgBoxStyle.Exclamation)
            receiveStatus = False
            clientSocket.Close()
            Exit Sub
        Finally
            RaiseEvent recOK(receiveStatus)
        End Try
        ' post data
        sc.Post(New SendOrPostCallback(AddressOf OnDatareceived), sb.ToString)
        ' start new read
        server.BeginAcceptTcpClient(New AsyncCallback(AddressOf DoAccept), server)
    End Sub

    ' now data to calling class and UI thread
    Private Sub OnDatareceived(ByVal state As Object)
        RaiseEvent Datareceived(state.ToString)
    End Sub

End Class
