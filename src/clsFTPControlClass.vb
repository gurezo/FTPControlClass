' @(h) clsFTPControlClass.vb          ver 01.00.00
'
' @(s)
' 
'
Option Strict Off
Option Explicit On 

Imports System
Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Net.Sockets

Public Class FTPControl

    Private m_sRemoteHost, m_sRemotePath, m_sRemoteUser As String
    Private m_sRemotePassword, m_sMess As String
    Private m_iRemotePort, m_iBytes As Int32
    Private m_objClientSocket As Socket
    Private m_iRetValue As Int32
    Private m_bLoggedIn As Boolean
    Private m_sMes, m_sReply As String

    'Set the size of the packet that is used to read and to write data to the FTP server
    'to the following specified size.
    Public Const BLOCK_SIZE = 512
    Private m_aBuffer(BLOCK_SIZE) As Byte
    Private ASCII As Encoding = Encoding.ASCII
    Public flag_bool As Boolean
    'General variable declaration
    Private m_sMessageString As String

    ' @(f)
    '
    ' 機能　　 :New処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Sub New()
        m_sRemoteHost = "microsoft"
        m_sRemotePath = "."
        m_sRemoteUser = "anonymous"
        m_sRemotePassword = ""
        m_sMessageString = ""
        m_iRemotePort = 21
        m_bLoggedIn = False
    End Sub

    ' @(f)
    '
    ' 機能　　 :New処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :Parameterized constructor
    '
    ' 備考　　 :
    '
    Public Sub New(ByVal sRemoteHost As String, _
                   ByVal sRemotePath As String, _
                   ByVal sRemoteUser As String, _
                   ByVal sRemotePassword As String, _
                   ByVal iRemotePort As Int32)
        m_sRemoteHost = sRemoteHost
        m_sRemotePath = sRemotePath
        m_sRemoteUser = sRemoteUser
        m_sRemotePassword = sRemotePassword
        m_sMessageString = ""
        m_iRemotePort = 21
        m_bLoggedIn = False
    End Sub

    ' @(f)
    '
    ' 機能　　 :ﾘﾓｰﾄﾎｽﾄﾌﾟﾛﾊﾟﾃｨ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾘﾓｰﾄﾎｽﾄの設定/取得を行う
    '
    ' 備考　　 :
    '
    Public Property RemoteHostFTPServer() As String
        Get
            Return m_sRemoteHost
        End Get
        Set(ByVal Value As String)
            m_sRemoteHost = Value
        End Set
    End Property


    ' @(f)
    '
    ' 機能　　 :ﾘﾓｰﾄﾎﾟｰﾄﾌﾟﾛﾊﾟﾃｨ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾘﾓｰﾄﾎﾟｰﾄの設定/取得を行う
    '
    ' 備考　　 :
    '
    Public Property RemotePort() As Int32
        Get
            Return m_iRemotePort
        End Get
        Set(ByVal Value As Int32)
            m_iRemotePort = Value
        End Set
    End Property

    ' @(f)
    '
    ' 機能　　 :ﾘﾓｰﾄﾊﾟｽﾌﾟﾛﾊﾟﾃｨ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾘﾓｰﾄﾊﾟｽの設定/取得を行う
    '
    ' 備考　　 :
    '
    Public Property RemotePath() As String
        Get
            Return m_sRemotePath
        End Get
        Set(ByVal Value As String)
            m_sRemotePath = Value
        End Set
    End Property

    ' @(f)
    '
    ' 機能　　 :ﾘﾓｰﾄﾊﾟｽﾜｰﾄﾞﾌﾟﾛﾊﾟﾃｨ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾘﾓｰﾄﾊﾟｽﾜｰﾄﾞの設定/取得を行う
    '
    ' 備考　　 :
    '
    Public Property RemotePassword() As String
        Get
            Return m_sRemotePassword
        End Get
        Set(ByVal Value As String)
            m_sRemotePassword = Value
        End Set
    End Property

    ' @(f)
    '
    ' 機能　　 :ﾘﾓｰﾄﾕｰｻﾞｰﾌﾟﾛﾊﾟﾃｨ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾘﾓｰﾄﾕｰｻﾞｰの設定/取得を行う
    '
    ' 備考　　 :
    '
    Public Property RemoteUser() As String
        Get
            Return m_sRemoteUser
        End Get
        Set(ByVal Value As String)
            m_sRemoteUser = Value
        End Set
    End Property

    ' @(f)
    '
    ' 機能　　 :ﾒｯｾｰｼﾞ文字列ﾌﾟﾛﾊﾟﾃｨ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾒｯｾｰｼﾞ文字列の設定/取得を行う
    '
    ' 備考　　 :
    '
    Public Property MessageString() As String
        Get
            Return m_sMessageString
        End Get
        Set(ByVal Value As String)
            m_sMessageString = Value
        End Set
    End Property

    ' @(f)
    '
    ' 機能　　 :ﾌｧｲﾙﾘｽﾄ取得処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :文字列配列にﾌｧｲﾙﾘｽﾄを格納
    '
    ' 備考　　 :
    '
    Public Function GetFileList(ByVal sMask As String) As String()

        Dim cSocket As Socket
        Dim bytes As Int32
        Dim seperator As Char = ControlChars.Lf
        Dim mess() As String

        m_sMes = ""
        ''ﾛｸﾞｲﾝﾁｪｯｸ
        If (Not (m_bLoggedIn)) Then
            Login()
        End If
        cSocket = CreateDataSocket()

        ''FTPｺﾏﾝﾄﾞ送信,
        SendCommand("NLST " & sMask)
        If (Not (m_iRetValue = 150 Or m_iRetValue = 125)) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        m_sMes = ""
        Do While (True)
            m_aBuffer.Clear(m_aBuffer, 0, m_aBuffer.Length)
            bytes = cSocket.Receive(m_aBuffer, m_aBuffer.Length, 0)
            m_sMes += ASCII.GetString(m_aBuffer, 0, bytes)
            If (bytes < m_aBuffer.Length) Then
                Exit Do
            End If
        Loop

        mess = m_sMes.Split(seperator)
        cSocket.Close()
        ReadReply()
        If (m_iRetValue <> 226) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If
        Return mess

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾌｧｲﾙｻｲｽﾞ取得処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Function GetFileSize(ByVal sFileName As String) As Long
        Dim size As Long
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTPｺﾏﾝﾄﾞ送信,
        SendCommand("SIZE " & sFileName)
        size = 0
        If (m_iRetValue = 213) Then
            size = Int64.Parse(m_sReply.Substring(4))
        Else
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If
        Return size

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾛｸﾞｲﾝ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Function Login() As Boolean
        m_objClientSocket = _
        New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        Dim ep As New IPEndPoint(Dns.Resolve(m_sRemoteHost).AddressList(0), m_iRemotePort)
        Try
            m_objClientSocket.Connect(ep)
        Catch ex As Exception
            MessageString = m_sReply
            Throw New IOException("Cannot connect to remote server")
        End Try

        ReadReply()
        If (m_iRetValue <> 220) Then
            CloseConnection()
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''FTPｺﾏﾝﾄﾞ送信　ﾕｰｻﾞｰﾛｸﾞｲﾝID
        SendCommand("USER " & m_sRemoteUser)
        If (Not (m_iRetValue = 331 Or m_iRetValue = 230)) Then
            Cleanup()
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        If (m_iRetValue <> 230) Then
            '''FTPｺﾏﾝﾄﾞ送信　ﾕｰｻﾞｰﾛｸﾞｲﾝﾊﾟｽﾜｰﾄﾞ
            SendCommand("PASS " & m_sRemotePassword)
            If (Not (m_iRetValue = 230 Or m_iRetValue = 202)) Then
                Cleanup()
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If
        End If
        m_bLoggedIn = True

        ''ChangeDirectoryﾙｰﾁﾝでﾕｰｻﾞｰ定義された
        ''FTPﾌｫﾙﾀﾞへ割り当てます
        ChangeDirectory(m_sRemotePath)

        ''戻り値ｾｯﾄ
        Return m_bLoggedIn

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾊﾞｲﾅﾘｰﾓｰﾄﾞ切替え処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - 切替えﾌﾗｸﾞ
    '
    ' 機能説明 :ﾌﾗｸﾞがtrueの時ﾊﾞｲﾅﾘｰﾓｰﾄﾞでﾀﾞｳﾝﾛｰﾄﾞします
    ' 　　　　  上記以外はASCIIです
    '
    ' 備考　　 :
    '
    Public Sub SetBinaryMode(ByVal bMode As Boolean)
        If (bMode = True) Then
            '''ﾊﾞｲﾅﾘｰﾓｰﾄﾞで'FTPｺﾏﾝﾄﾞ送信
            '''TYPEはﾓｰﾄﾞの指定をする時に使用
            SendCommand("TYPE I")
        Else
            '''ASCIIﾓｰﾄﾞでFTPｺﾏﾝﾄﾞ送信
            '''TYPEはﾓｰﾄﾞの指定をする時に使用
            SendCommand("TYPE A")
        End If
        If (m_iRetValue <> 200) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If
    End Sub

    ' @(f)
    '
    ' 機能　　 :ﾀﾞｳﾝﾛｰﾄﾞﾌｧｲﾙ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - ﾌｧｲﾙ文字列
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のﾛｰｶﾙﾌｫﾙﾀﾞにﾀﾞｳﾝﾛｰﾄﾞ
    '
    ' 備考　　 :元ﾌｧｲﾙ名と同じ
    '
    Public Sub DownloadFile(ByVal sFileName As String)
        DownloadFile(sFileName, "", False)
    End Sub

    ' @(f)
    '
    ' 機能　　 :ﾀﾞｳﾝﾛｰﾄﾞﾌｧｲﾙ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - 元ﾌｧｲﾙ文字列
    ' 　　　　  ARG2 - 変更ﾌｧｲﾙ文字列
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のﾛｰｶﾙﾌｫﾙﾀﾞにﾀﾞｳﾝﾛｰﾄﾞ
    '
    ' 備考　　 :ﾌｧｲﾙ名を変更してﾀﾞｳﾝﾛｰﾄﾞ
    '
    Public Sub DownloadFile(ByVal sFileName As String, _
                            ByVal bResume As Boolean)
        DownloadFile(sFileName, "", bResume)
    End Sub

    ' @(f)
    '
    ' 機能　　 :ﾀﾞｳﾝﾛｰﾄﾞﾌｧｲﾙ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - 元ﾌｧｲﾙ文字列
    ' 　　　　  ARG2 - ﾛｰｶﾙﾌｧｲﾙ文字列
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のﾛｰｶﾙﾌｫﾙﾀﾞにﾀﾞｳﾝﾛｰﾄﾞ
    '
    ' 備考　　 :存在するﾌｫﾙﾀﾞﾊﾟｽを付記する事
    '
    Public Sub DownloadFile(ByVal sFileName As String, _
                            ByVal sLocalFileName As String)
        DownloadFile(sFileName, sLocalFileName, False)
    End Sub

    ' @(f)
    '
    ' 機能　　 :ﾀﾞｳﾝﾛｰﾄﾞﾌｧｲﾙ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - 元ﾌｧｲﾙ文字列
    ' 　　　　  ARG2 - ﾛｰｶﾙﾌｧｲﾙ文字列
    ' 　　　　  ARG3 - 変更ﾌｧｲﾙ文字列
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のﾛｰｶﾙﾌｫﾙﾀﾞにﾀﾞｳﾝﾛｰﾄﾞ
    '
    ' 備考　　 :存在するﾌｫﾙﾀﾞﾊﾟｽを付記する事
    '
    Public Sub DownloadFile(ByVal sFileName As String, _
                            ByVal sLocalFileName As String, _
                            ByVal bResume As Boolean)

        Dim st As Stream
        Dim output As FileStream
        Dim cSocket As Socket
        Dim offset, npos As Long

        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        SetBinaryMode(True)
        If (sLocalFileName.Equals("")) Then
            sLocalFileName = sFileName
        End If

        If (Not (File.Exists(sLocalFileName))) Then
            st = File.Create(sLocalFileName)
            st.Close()
        End If

        output = New FileStream(sLocalFileName, FileMode.Open)
        cSocket = CreateDataSocket()
        offset = 0

        If (bResume) Then
            offset = output.Length
            If (offset > 0) Then
                '''FTPｺﾏﾝﾄﾞ送信　ﾘｽﾀｰﾄ
                SendCommand("REST " & offset)
                If (m_iRetValue <> 350) Then
                    offset = 0
                End If
            End If
            If (offset > 0) Then
                npos = output.Seek(offset, SeekOrigin.Begin)
            End If
        End If

        ''FTPｺﾏﾝﾄﾞ送信　ﾘﾄﾗｲ
        SendCommand("RETR " & sFileName)
        If (Not (m_iRetValue = 150 Or m_iRetValue = 125)) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        Do While (True)
            m_aBuffer.Clear(m_aBuffer, 0, m_aBuffer.Length)
            m_iBytes = cSocket.Receive(m_aBuffer, m_aBuffer.Length, 0)
            output.Write(m_aBuffer, 0, m_iBytes)
            If (m_iBytes <= 0) Then
                Exit Do
            End If
        Loop
        output.Close()

        If (cSocket.Connected) Then
            cSocket.Close()
        End If

        ReadReply()
        If (Not (m_iRetValue = 226 Or m_iRetValue = 250)) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If
    End Sub

    ' @(f)
    '
    ' 機能　　 :ｱｯﾌﾟﾛｰﾄﾞﾌｧｲﾙ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - ﾌｧｲﾙ文字列
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のFTPｻｲﾄにｱｯﾌﾟﾛｰﾄﾞ
    '
    ' 備考　　 :
    '
    Public Sub UploadFile(ByVal sFileName As String)
        UploadFile(sFileName, False)
    End Sub

    ' @(f)
    '
    ' 機能　　 :ｱｯﾌﾟﾛｰﾄﾞﾌｧｲﾙ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - ﾌｧｲﾙ文字列
    ' 　　　　  ARG2 - OKNGﾌﾗｸﾞ
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のFTPｻｲﾄにｱｯﾌﾟﾛｰﾄﾞ
    '
    ' 備考　　 :
    '
    Public Sub UploadFile(ByVal sFileName As String, _
                          ByVal bResume As Boolean)
        Dim cSocket As Socket
        Dim offset As Long
        Dim input As FileStream
        Dim bFileNotFound As Boolean

        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        cSocket = CreateDataSocket()
        offset = 0
        If (bResume) Then
            Try
                SetBinaryMode(True)
                offset = GetFileSize(sFileName)
            Catch ex As Exception
                offset = 0
            End Try
        End If

        If (offset > 0) Then
            SendCommand("REST " & offset)
            If (m_iRetValue <> 350) Then
                ''''ﾘﾓｰﾄｻ-ﾊﾞｰが再試行をｻﾎﾟｰﾄしていない場合
                offset = 0
            End If
        End If

        ''FTPｺﾏﾝﾄﾞ送信　ﾌｧｲﾙ保存
        SendCommand("STOR " & Path.GetFileName(sFileName))
        If (Not (m_iRetValue = 125 Or m_iRetValue = 150)) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''ｱｯﾌﾟﾛｰﾄﾞ前にﾌｧｲﾙの存在ﾁｪｯｸ
        bFileNotFound = False
        If (File.Exists(sFileName)) Then
            '''ﾌｧｲﾙｵｰﾌﾟﾝ
            input = New FileStream(sFileName, FileMode.Open)
            If (offset <> 0) Then
                input.Seek(offset, SeekOrigin.Begin)
            End If

            '''ﾌｧｲﾙｱｯﾌﾟﾛｰﾄﾞ
            m_iBytes = input.Read(m_aBuffer, 0, m_aBuffer.Length)
            Do While (m_iBytes > 0)
                cSocket.Send(m_aBuffer, m_iBytes, 0)
                m_iBytes = input.Read(m_aBuffer, 0, m_aBuffer.Length)
            Loop
            input.Close()
        Else
            bFileNotFound = True
        End If

        If (cSocket.Connected) Then
            cSocket.Close()
        End If

        ''ﾌｧｲﾙｱｯﾌﾟﾛｰﾄﾞ後の存在ﾁｪｯｸ
        If (bFileNotFound) Then
            MessageString = m_sReply
            Throw New IOException("The file: " & sFileName & " was not found." & _
            " Cannot upload the file to the FTP site.")
        End If

        ReadReply()
        If (Not (m_iRetValue = 226 Or m_iRetValue = 250)) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

    End Sub

    ' @(f)
    '
    ' 機能　　 :ﾌｧｲﾙ削除処理
    '
    ' 返り値　 :正常終了 - True
    ' 　　　    ｴﾗｰ終了 - False
    '
    ' 引き数　 :ARG1 - ﾌｧｲﾙ文字列
    '
    ' 機能説明 :指定のﾌｧｲﾙを任意のFTPｻｲﾄから削除します
    '
    ' 備考　　 :
    '
    Public Function DeleteFile(ByVal sFileName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTPｺﾏﾝﾄﾞ送信　ﾌｧｲﾙ削除
        SendCommand("DELE " & sFileName)
        If (m_iRetValue <> 250) Then
            bResult = False
            MessageString = m_sReply
        End If

        ''戻り値ｾｯﾄ
        Return bResult

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾌｧｲﾙ名変更処理
    '
    ' 返り値　 :正常終了 - True
    ' 　　　    ｴﾗｰ終了 - False
    '
    ' 引き数　 :ARG1 - 変更前ﾌｧｲﾙ文字列
    ' 　　　　  ARG2 - 変更後ﾌｧｲﾙ文字列
    '
    ' 機能説明 :FTPｻｲﾄ上にあるﾌｧｲﾙ名を変更します
    '
    ' 備考　　 :
    '
    Public Function RenameFile(ByVal sOldFileName As String, _
                               ByVal sNewFileName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTPｺﾏﾝﾄﾞ送信　RNFR
        ''指定したﾌｧｲﾙ名を変更する。変更元ﾌｧｲﾙ名の指定である。
        ''RNTOを続けて実行しなくてはならない
        SendCommand("RNFR " & sOldFileName)
        If (m_iRetValue <> 350) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''FTPｺﾏﾝﾄﾞ送信　RNTO
        ''RNFRの後に実行される。
        ''RNFRｺﾏﾝﾄﾞで指定したﾌｧｲﾙを、指定したﾌｧｲﾙ名に変更する。
        ''変更先ﾌｧｲﾙ名の指定である()
        SendCommand("RNTO " & sNewFileName)
        If (m_iRetValue <> 250) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''戻り値ｾｯﾄ
        Return bResult

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾌｫﾙﾀﾞ作成処理
    '
    ' 返り値　 :正常終了 - True
    ' 　　　    ｴﾗｰ終了 - False
    '
    ' 引き数　 :ARG1 - ﾃﾞｨﾚｸﾄﾘ文字列
    '
    ' 機能説明 :FTPｻｲﾄ上にﾌｫﾙﾀﾞを作成します
    '
    ' 備考　　 :
    '
    Public Function CreateDirectory(ByVal sDirName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTPｺﾏﾝﾄﾞ送信　MKD
        SendCommand("MKD " & sDirName)
        If (m_iRetValue <> 257) Then
            bResult = False
            MessageString = m_sReply
        End If

        ''戻り値ｾｯﾄ
        Return bResult

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾌｫﾙﾀﾞ削除処理
    '
    ' 返り値　 :正常終了 - True
    ' 　　　    ｴﾗｰ終了 - False
    '
    ' 引き数　 :ARG1 - ﾃﾞｨﾚｸﾄﾘ文字列
    '
    ' 機能説明 :FTPｻｲﾄ上にﾌｫﾙﾀﾞを削除します
    '
    ' 備考　　 :
    '
    Public Function RemoveDirectory(ByVal sDirName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTPｺﾏﾝﾄﾞ送信　RMD
        SendCommand("RMD " & sDirName)
        If (m_iRetValue <> 250) Then
            bResult = False
            MessageString = m_sReply
        End If

        ''戻り値ｾｯﾄ
        Return bResult

    End Function

    ' @(f)
    '
    ' 機能　　 :ﾌｫﾙﾀﾞ移動処理
    '
    ' 返り値　 :正常終了 - True
    ' 　　　    ｴﾗｰ終了 - False
    '
    ' 引き数　 :ARG1 - ﾃﾞｨﾚｸﾄﾘ文字列
    '
    ' 機能説明 :FTPｻｲﾄ上にﾌｫﾙﾀﾞから任意ﾌｫﾙﾀﾞへ移動
    '
    ' 備考　　 :
    '
    Public Function ChangeDirectory(ByVal sDirName As String) As Boolean

        Dim bResult As Boolean

        bResult = True

        ''現在位置確認
        If (sDirName.Equals(".")) Then
            Exit Function
        End If

        ''ﾛｸﾞｲﾝﾁｪｯｸ
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTPｺﾏﾝﾄﾞ送信　ﾌｫﾙﾀﾞ移動
        SendCommand("CWD " & sDirName)
        If (m_iRetValue <> 250) Then
            bResult = False
            MessageString = m_sReply
        End If
        Me.m_sRemotePath = sDirName

        ''戻り値ｾｯﾄ
        Return bResult

    End Function

    ' @(f)
    '
    ' 機能　　 :FTP切断処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Public Sub CloseConnection()

        If (Not (m_objClientSocket Is Nothing)) Then
            '''FTPｺﾏﾝﾄﾞ送信　切断
            SendCommand("QUIT")
        End If
        Cleanup()

    End Sub

    ' @(f)
    '
    ' 機能　　 :読込みﾘﾄﾗｲ処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Private Sub ReadReply()
        m_sMes = ""
        m_sReply = ReadLine()
        m_iRetValue = Int32.Parse(m_sReply.Substring(0, 3))
    End Sub

    ' @(f)
    '
    ' 機能　　 :変数初期化処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Private Sub Cleanup()
        If Not (m_objClientSocket Is Nothing) Then
            m_objClientSocket.Close()
            m_objClientSocket = Nothing
        End If
        m_bLoggedIn = False
    End Sub

    ' @(f)
    '
    ' 機能　　 :読込み処理
    '
    ' 返り値　 :正常終了 - 読込み文字列
    ' 　　　    ｴﾗｰ終了 - ””空文字
    '
    ' 引き数　 :ARG1 - ﾒｯｾｰｼﾞﾌﾗｸﾞ
    '
    ' 機能説明 :FTPｻｲﾄより読込みを行います。
    '
    ' 備考　　 :
    '
    Private Function ReadLine(Optional ByVal bClearMes As Boolean = False) As String

        Dim seperator As Char = ControlChars.Lf
        Dim mess() As String

        If (bClearMes) Then
            m_sMes = ""
        End If
        Do While (True)
            m_aBuffer.Clear(m_aBuffer, 0, BLOCK_SIZE)
            m_iBytes = m_objClientSocket.Receive(m_aBuffer, m_aBuffer.Length, 0)
            m_sMes += ASCII.GetString(m_aBuffer, 0, m_iBytes)
            If (m_iBytes < m_aBuffer.Length) Then
                Exit Do
            End If
        Loop

        mess = m_sMes.Split(seperator)
        If (m_sMes.Length > 2) Then
            m_sMes = mess(mess.Length - 2)
        Else
            m_sMes = mess(0)
        End If

        If (Not (m_sMes.Substring(3, 1).Equals(" "))) Then
            Return ReadLine(True)
        End If
        Return m_sMes

    End Function

    ' @(f)
    '
    ' 機能　　 :FTPｺﾏﾝﾄﾞ送信処理
    '
    ' 返り値　 :なし
    '
    ' 引き数　 :ARG1 - ｺﾏﾝﾄﾞ文字列
    '
    ' 機能説明 :
    '
    ' 備考　　 :
    '
    Private Sub SendCommand(ByVal sCommand As String)

        sCommand = sCommand & ControlChars.CrLf
        Dim cmdbytes As Byte() = ASCII.GetBytes(sCommand)

        m_objClientSocket.Send(cmdbytes, cmdbytes.Length, 0)
        ReadReply()

    End Sub

    ' @(f)
    '
    ' 機能　　 :読込み処理
    '
    ' 返り値　 :正常終了 - ｿｹｯﾄ変数
    '
    ' 引き数　 :なし
    '
    ' 機能説明 :ﾃﾞｰﾀｿｹｯﾄを作成します
    '
    ' 備考　　 :
    '
    Private Function CreateDataSocket() As Socket

        Dim index1, index2, len As Int32
        Dim partCount, i, port As Int32
        Dim ipData, buf, ipAddress As String
        Dim parts(6) As Int32
        Dim ch As Char
        Dim s As Socket
        Dim ep As IPEndPoint

        ''FTPｺﾏﾝﾄﾞ送信　PASV
        SendCommand("PASV")
        If (m_iRetValue <> 227) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        index1 = m_sReply.IndexOf("(")
        index2 = m_sReply.IndexOf(")")
        ipData = m_sReply.Substring(index1 + 1, index2 - index1 - 1)
        len = ipData.Length

        partCount = 0
        buf = ""
        For i = 0 To ((len - 1) And partCount <= 6)
            ch = Char.Parse(ipData.Substring(i, 1))
            If (Char.IsDigit(ch)) Then
                buf += ch
            ElseIf (ch <> ",") Then
                MessageString = m_sReply
                Throw New IOException("Malformed PASV reply: " & m_sReply)
            End If
            If ((ch = ",") Or (i + 1 = len)) Then
                Try
                    parts(partCount) = Int32.Parse(buf)
                    partCount += 1
                    buf = ""
                Catch ex As Exception
                    MessageString = m_sReply
                    Throw New IOException("Malformed PASV reply: " & m_sReply)
                End Try
            End If
        Next

        ipAddress = parts(0) & "." & parts(1) & "." & parts(2) & "." & parts(3)
        port = parts(4) << 8

        ''ﾃﾞｰﾀﾎﾟｰﾄﾅﾝﾊﾞｰを測定してください。
        port = port + parts(5)
        s = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
        ep = New IPEndPoint(Dns.Resolve(ipAddress).AddressList(0), port)

        Try
            s.Connect(ep)
        Catch ex As Exception
            MessageString = m_sReply
            Throw New IOException("Cannot connect to remote server")
            'If you cannot connect to the FTP
            'server that is specified, make the boolean variable false.
            flag_bool = False
        End Try

        ''指定されるFTPｻｰﾊﾞｰに接続できたなら、変数をTrueにします
        flag_bool = True
        Return s

    End Function

End Class
