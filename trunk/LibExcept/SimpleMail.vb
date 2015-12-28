Imports System.Net
Imports System.Net.Sockets
Imports System.Threading

'--
'-- a simple class for trivial SMTP mail support
'-- 
'-- basic features:
'--
'--   ** trivial SMTP implementation
'--   ** HTML body
'--   ** plain text body
'--   ** one file attachment
'--   ** basic retry mechanism
'--
'-- Jeff Atwood
'-- http://www.codinghorror.com
'--
Public Class SimpleMail

    Public Class SMTPMailMessage
        Public From As String
        Public [To] As String
        Public Subject As String
        Public Body As String
        Public BodyHTML As String
        Public AttachmentPath As String
    End Class

    Public Class SMTPClient

        Private Const _intBufferSize As Integer = 1024
        Private Const _intResponseTimeExpected As Integer = 10
        Private Const _intResponseTimeMax As Integer = 750
        Private Const _strAddressSeperator As String = ";"
        Private Const _intMaxRetries As Integer = 5        
        Private Const _blnDebugMode As Boolean = False
        Private Const _blnPlainTextOnly As Boolean = False

        Private _strDefaultDomain As String = "vendteksys.com"
        Private _strServer As String = "vsi.ods.org"
        Private _intPort As Integer = 25

        Private _intRetries As Integer = 1
        Private _strLastResponse As String

        Public Property Port() As Integer
            Get
                Return _intPort
            End Get
            Set(ByVal Value As Integer)
                _intPort = Value
            End Set
        End Property

        Public Property Server() As String
            Get
                Return _strServer
            End Get
            Set(ByVal Value As String)
                _strServer = Value
            End Set
        End Property

        Public Property DefaultDomain() As String
            Get
                Return _strDefaultDomain
            End Get
            Set(ByVal Value As String)
                _strDefaultDomain = Value
            End Set
        End Property

        '-- 
        '-- send data over the current network connection
        '--
        Private Sub SendData(ByVal objTcpClient As TcpClient, ByVal strData As String)
            Dim objNetworkStream As NetworkStream = objTcpClient.GetStream()
            Dim bytWriteBuffer(strData.Length) As Byte
            Dim en As New System.Text.ASCIIEncoding

            bytWriteBuffer = en.GetBytes(strData)
            objNetworkStream.Write(bytWriteBuffer, 0, bytWriteBuffer.Length)
        End Sub

        '--
        '-- get data from the current network connection
        '--
        Private Function GetData(ByVal objTcpClient As TcpClient) As String
            Dim objNetworkStream As System.Net.Sockets.NetworkStream = objTcpClient.GetStream()

            If objNetworkStream.DataAvailable Then
                Dim bytReadBuffer() As Byte
                Dim intStreamSize As Integer
                bytReadBuffer = New Byte(_intBufferSize) {}
                intStreamSize = objNetworkStream.Read(bytReadBuffer, 0, bytReadBuffer.Length)
                Dim en As New System.Text.ASCIIEncoding
                Return en.GetString(bytReadBuffer)
            Else
                Return ""
            End If
        End Function

        '--
        '-- issue a SMTP command
        '--
        Private Function Command(ByVal objTcpClient As TcpClient, ByVal strCommand As String, _
            Optional ByVal strExpectedResponse As String = "250", _
            Optional ByVal intExpectedResponseTimeMax As Integer = _intResponseTimeMax) As Boolean

            Dim intResponseTime As Integer

            '-- send the command over the socket with a trailing cr/lf
            If strCommand.Length > 0 Then
                SendData(objTcpClient, strCommand & Environment.NewLine)
            End If

            '-- wait until we get a response, or time out
            _strLastResponse = ""
            intResponseTime = 0
            Do While (_strLastResponse = "") And (intResponseTime <= intExpectedResponseTimeMax)
                intResponseTime += _intResponseTimeExpected
                _strLastResponse = GetData(objTcpClient)
                Thread.CurrentThread.Sleep(_intResponseTimeExpected)
            Loop

            '-- this is helpful for debugging SMTP problems
            If _blnDebugMode Then
                Debug.WriteLine("SMTP >> " + strCommand + " (after " + intResponseTime.ToString + "ms)")
                Debug.WriteLine("SMTP << " + _strLastResponse)
            End If

            '-- if we have a response, check the first 10 characters for the expected response code
            If _strLastResponse = "" Then
                If _blnDebugMode Then
                    Debug.WriteLine("** EXPECTED RESPONSE " + strExpectedResponse + " NOT RETURNED **")
                End If
                Return False
            Else
                Return (_strLastResponse.IndexOf(strExpectedResponse, 0, 10) <> -1)
            End If
        End Function

        '--
        '-- send mail with integrated retry mechanism
        '--
        Public Function SendMail(ByVal objSMTPMailMessage As SMTPMailMessage) As Boolean
            Dim intRetryInterval As Integer = 333
            Try
                If Not SendMailInternal(objSMTPMailMessage) Then
                    _intRetries += 1
                    If _intRetries <= _intMaxRetries Then
                        Thread.CurrentThread.Sleep(intRetryInterval)
                        SendMail(objSMTPMailMessage)
                    Else
                        Return False
                    End If
                End If
            Catch ex As Exception
                _intRetries += 1
                If _intRetries <= _intMaxRetries Then
                    Thread.CurrentThread.Sleep(intRetryInterval)
                    SendMail(objSMTPMailMessage)
                Else
                    Throw ex
                End If
            End Try
            'Console.WriteLine("sent after " + _intRetries.ToString)
            _intRetries = 1
            Return True
        End Function

        '--
        '-- send an email via trivial SMTP
        '--
        Private Function SendMailInternal(ByVal objSMTPMailMessage As SMTPMailMessage) As Boolean
            Dim objSMTPClient As New SMTPClient
            Dim objIPHostEntry As IPHostEntry
            Dim objTcpClient As New TcpClient

            '-- resolve server text name to an IP address
            Try
                objIPHostEntry = Dns.GetHostByName(_strServer)
            Catch e As Exception
                Throw New Exception("Unable to resolve server name " + _strServer, e)
            End Try

            '-- attempt to connect to the server by IP address and port number
            Try
                objTcpClient.Connect(objIPHostEntry.AddressList(0), _intPort)
            Catch e As Exception
                Throw New Exception("Unable to connect to SMTP server at " + _strServer.ToString + ":" + _intPort.ToString, e)
            End Try

            '-- make sure we get the SMTP welcome message
            If Not objSMTPClient.Command(objTcpClient, "", "220") Then
                Throw New Exception("SMTP server at " + _strServer.ToString + ":" + _intPort.ToString + _
                    " did not return expected SMTP welcome:" _
                    + Environment.NewLine + objSMTPClient._strLastResponse)
            End If

            If Not objSMTPClient.Command(objTcpClient, "HELO " + Environment.MachineName) Then
                objTcpClient.Close()
                Return False
            End If

            If objSMTPMailMessage.From = "" Then
                objSMTPMailMessage.From = System.AppDomain.CurrentDomain.FriendlyName.ToLower + "@" + Environment.MachineName.ToLower + "." + _strDefaultDomain
            End If

            If Not objSMTPClient.Command(objTcpClient, "MAIL FROM: <" + objSMTPMailMessage.From + ">") Then
                objTcpClient.Close()
                Return False
            End If

            '-- send email to more than one recipient
            Dim arRecipients() As String = objSMTPMailMessage.To.Split(_strAddressSeperator.ToCharArray)
            Dim strRecipient As String
            For Each strRecipient In arRecipients
                If Not objSMTPClient.Command(objTcpClient, "RCPT TO: <" + strRecipient + ">") Then
                    objTcpClient.Close()
                    Return False
                End If
            Next

            If Not objSMTPClient.Command(objTcpClient, "DATA", "354") Then
                objTcpClient.Close()
                Return False
            End If

            Dim objStringBuilder As New Text.StringBuilder
            With objStringBuilder
                '-- write common email headers
                .Append("To: " + objSMTPMailMessage.To + Environment.NewLine)
                .Append("From: " + objSMTPMailMessage.From + Environment.NewLine)
                .Append("Subject: " + objSMTPMailMessage.Subject + Environment.NewLine)
                If _blnPlainTextOnly Then
                    '-- write plain text body
                    .Append(Environment.NewLine + objSMTPMailMessage.Body + Environment.NewLine)
                Else
                    Dim strContentType As String
                    '-- typical case; mixed content will be displayed side-by-side
                    strContentType = "multipart/mixed"
                    '-- unusual case; text and HTML body are both included, let the reader determine which it can handle
                    If objSMTPMailMessage.Body <> "" And objSMTPMailMessage.BodyHTML <> "" Then
                        strContentType = "multipart/alternative"
                    End If

                    .Append("MIME-Version: 1.0" + Environment.NewLine)
                    .Append("Content-Type: " + strContentType + "; boundary=""NextMimePart""" + Environment.NewLine)
                    .Append("Content-Transfer-Encoding: 7bit" + Environment.NewLine)
                    ' -- default content (for non-MIME compliant email clients, should be extremely rare)
                    .Append("This message is in MIME format. Since your mail reader does not understand " + Environment.NewLine)
                    .Append("this format, some or all of this message may not be legible." + Environment.NewLine)
                    '-- handle text body (if any)
                    If objSMTPMailMessage.Body <> "" Then
                        .Append(Environment.NewLine + "--NextMimePart" + Environment.NewLine)
                        .Append("Content-Type: text/plain; charset=us-ascii" + Environment.NewLine)
                        .Append(Environment.NewLine + objSMTPMailMessage.Body + Environment.NewLine)
                    End If
                    ' -- handle HTML body (if any)
                    If objSMTPMailMessage.BodyHTML <> "" Then
                        .Append(Environment.NewLine + "--NextMimePart" + Environment.NewLine)
                        .Append("Content-Type: text/html; charset=iso-8859-1" + Environment.NewLine)
                        .Append(Environment.NewLine + objSMTPMailMessage.BodyHTML + Environment.NewLine)
                    End If
                    '-- handle attachment (if any)
                    If objSMTPMailMessage.AttachmentPath <> "" Then
                        .Append(FileToMimeString(objSMTPMailMessage.AttachmentPath))
                    End If
                End If
                '-- <crlf>.<crlf> marks end of message content
                .Append(Environment.NewLine + "." + Environment.NewLine)
            End With

            If Not objSMTPClient.Command(objTcpClient, objStringBuilder.ToString) Then
                objTcpClient.Close()
                Return False
            End If

            objSMTPClient.Command(objTcpClient, "QUIT")
            objTcpClient.Close()

            Return True
        End Function

        '--
        '-- turn a file into a base 64 string
        '--
        Private Function FileToMimeString(ByVal strFilepath As String) As String

            Dim objFilestream As System.IO.FileStream
            Dim objStringBuilder As New Text.StringBuilder
            '-- note that chunk size is equal to maximum line width
            Const intChunkSize As Integer = 75
            Dim bytRead(intChunkSize) As Byte
            Dim intRead As Integer
            Dim strFilename As String

            '-- get just the filename out of the path
            strFilename = System.IO.Path.GetFileName(strFilepath)

            With objStringBuilder
                .Append(Environment.NewLine & "--NextMimePart" + Environment.NewLine)
                .Append("Content-Type: application/octet-stream; name=""" + strFilename & """" + Environment.NewLine)
                .Append("Content-Transfer-Encoding: base64" + Environment.NewLine)
                .Append("Content-Disposition: attachment; filename=""" + strFilename & """" + Environment.NewLine)
                .Append(Environment.NewLine)
            End With

            objFilestream = New System.IO.FileStream(strFilepath, System.IO.FileMode.Open, System.IO.FileAccess.Read)
            intRead = objFilestream.Read(bytRead, 0, intChunkSize)
            Do While intRead > 0
                objStringBuilder.Append(System.Convert.ToBase64String(bytRead, 0, intRead))
                objStringBuilder.Append(Environment.NewLine)
                intRead = objFilestream.Read(bytRead, 0, intChunkSize)
            Loop
            objFilestream.Close()

            Return objStringBuilder.ToString
        End Function

    End Class

End Class