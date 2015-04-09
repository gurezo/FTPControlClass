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
    ' �@�\�@�@ :New����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :New����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :Parameterized constructor
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :�Ӱ�ν������è����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :�Ӱ�νĂ̐ݒ�/�擾���s��
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :�Ӱ��߰������è����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :�Ӱ��߰Ă̐ݒ�/�擾���s��
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :�Ӱ��߽�����è����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :�Ӱ��߽�̐ݒ�/�擾���s��
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :�Ӱ��߽ܰ�������è����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :�Ӱ��߽ܰ�ނ̐ݒ�/�擾���s��
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :�Ӱ�հ�ް�����è����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :�Ӱ�հ�ް�̐ݒ�/�擾���s��
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :ү���ޕ����������è����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :ү���ޕ�����̐ݒ�/�擾���s��
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :̧��ؽĎ擾����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :������z���̧��ؽĂ��i�[
    '
    ' ���l�@�@ :
    '
    Public Function GetFileList(ByVal sMask As String) As String()

        Dim cSocket As Socket
        Dim bytes As Int32
        Dim seperator As Char = ControlChars.Lf
        Dim mess() As String

        m_sMes = ""
        ''۸޲�����
        If (Not (m_bLoggedIn)) Then
            Login()
        End If
        cSocket = CreateDataSocket()

        ''FTP����ޑ��M,
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
    ' �@�\�@�@ :̧�ٻ��ގ擾����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Public Function GetFileSize(ByVal sFileName As String) As Long
        Dim size As Long
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTP����ޑ��M,
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
    ' �@�\�@�@ :۸޲ݏ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
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

        ''FTP����ޑ��M�@հ�ް۸޲�ID
        SendCommand("USER " & m_sRemoteUser)
        If (Not (m_iRetValue = 331 Or m_iRetValue = 230)) Then
            Cleanup()
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        If (m_iRetValue <> 230) Then
            '''FTP����ޑ��M�@հ�ް۸޲��߽ܰ��
            SendCommand("PASS " & m_sRemotePassword)
            If (Not (m_iRetValue = 230 Or m_iRetValue = 202)) Then
                Cleanup()
                MessageString = m_sReply
                Throw New IOException(m_sReply.Substring(4))
            End If
        End If
        m_bLoggedIn = True

        ''ChangeDirectoryٰ�݂�հ�ް��`���ꂽ
        ''FTP̫��ނ֊��蓖�Ă܂�
        ChangeDirectory(m_sRemotePath)

        ''�߂�l���
        Return m_bLoggedIn

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :�޲�ذӰ�ސؑւ�����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - �ؑւ��׸�
    '
    ' �@�\���� :�׸ނ�true�̎��޲�ذӰ�ނ��޳�۰�ނ��܂�
    ' �@�@�@�@  ��L�ȊO��ASCII�ł�
    '
    ' ���l�@�@ :
    '
    Public Sub SetBinaryMode(ByVal bMode As Boolean)
        If (bMode = True) Then
            '''�޲�ذӰ�ނ�'FTP����ޑ��M
            '''TYPE��Ӱ�ނ̎w������鎞�Ɏg�p
            SendCommand("TYPE I")
        Else
            '''ASCIIӰ�ނ�FTP����ޑ��M
            '''TYPE��Ӱ�ނ̎w������鎞�Ɏg�p
            SendCommand("TYPE A")
        End If
        If (m_iRetValue <> 200) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�޳�۰��̧�ُ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ̧�ٕ�����
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�۰��̫��ނ��޳�۰��
    '
    ' ���l�@�@ :��̧�ٖ��Ɠ���
    '
    Public Sub DownloadFile(ByVal sFileName As String)
        DownloadFile(sFileName, "", False)
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�޳�۰��̧�ُ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ��̧�ٕ�����
    ' �@�@�@�@  ARG2 - �ύX̧�ٕ�����
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�۰��̫��ނ��޳�۰��
    '
    ' ���l�@�@ :̧�ٖ���ύX�����޳�۰��
    '
    Public Sub DownloadFile(ByVal sFileName As String, _
                            ByVal bResume As Boolean)
        DownloadFile(sFileName, "", bResume)
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�޳�۰��̧�ُ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ��̧�ٕ�����
    ' �@�@�@�@  ARG2 - ۰��̧�ٕ�����
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�۰��̫��ނ��޳�۰��
    '
    ' ���l�@�@ :���݂���̫����߽��t�L���鎖
    '
    Public Sub DownloadFile(ByVal sFileName As String, _
                            ByVal sLocalFileName As String)
        DownloadFile(sFileName, sLocalFileName, False)
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�޳�۰��̧�ُ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ��̧�ٕ�����
    ' �@�@�@�@  ARG2 - ۰��̧�ٕ�����
    ' �@�@�@�@  ARG3 - �ύX̧�ٕ�����
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�۰��̫��ނ��޳�۰��
    '
    ' ���l�@�@ :���݂���̫����߽��t�L���鎖
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
                '''FTP����ޑ��M�@ؽ���
                SendCommand("REST " & offset)
                If (m_iRetValue <> 350) Then
                    offset = 0
                End If
            End If
            If (offset > 0) Then
                npos = output.Seek(offset, SeekOrigin.Begin)
            End If
        End If

        ''FTP����ޑ��M�@��ײ
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
    ' �@�\�@�@ :����۰��̧�ُ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ̧�ٕ�����
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�FTP��Ăɱ���۰��
    '
    ' ���l�@�@ :
    '
    Public Sub UploadFile(ByVal sFileName As String)
        UploadFile(sFileName, False)
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :����۰��̧�ُ���
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ̧�ٕ�����
    ' �@�@�@�@  ARG2 - OKNG�׸�
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�FTP��Ăɱ���۰��
    '
    ' ���l�@�@ :
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
                ''''�ӰĻ-�ް���Ď��s���߰Ă��Ă��Ȃ��ꍇ
                offset = 0
            End If
        End If

        ''FTP����ޑ��M�@̧�ٕۑ�
        SendCommand("STOR " & Path.GetFileName(sFileName))
        If (Not (m_iRetValue = 125 Or m_iRetValue = 150)) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''����۰�ޑO��̧�ق̑�������
        bFileNotFound = False
        If (File.Exists(sFileName)) Then
            '''̧�ٵ����
            input = New FileStream(sFileName, FileMode.Open)
            If (offset <> 0) Then
                input.Seek(offset, SeekOrigin.Begin)
            End If

            '''̧�ٱ���۰��
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

        ''̧�ٱ���۰�ތ�̑�������
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
    ' �@�\�@�@ :̧�ٍ폜����
    '
    ' �Ԃ�l�@ :����I�� - True
    ' �@�@�@    �װ�I�� - False
    '
    ' �������@ :ARG1 - ̧�ٕ�����
    '
    ' �@�\���� :�w���̧�ق�C�ӂ�FTP��Ă���폜���܂�
    '
    ' ���l�@�@ :
    '
    Public Function DeleteFile(ByVal sFileName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTP����ޑ��M�@̧�ٍ폜
        SendCommand("DELE " & sFileName)
        If (m_iRetValue <> 250) Then
            bResult = False
            MessageString = m_sReply
        End If

        ''�߂�l���
        Return bResult

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :̧�ٖ��ύX����
    '
    ' �Ԃ�l�@ :����I�� - True
    ' �@�@�@    �װ�I�� - False
    '
    ' �������@ :ARG1 - �ύX�O̧�ٕ�����
    ' �@�@�@�@  ARG2 - �ύX��̧�ٕ�����
    '
    ' �@�\���� :FTP��ď�ɂ���̧�ٖ���ύX���܂�
    '
    ' ���l�@�@ :
    '
    Public Function RenameFile(ByVal sOldFileName As String, _
                               ByVal sNewFileName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTP����ޑ��M�@RNFR
        ''�w�肵��̧�ٖ���ύX����B�ύX��̧�ٖ��̎w��ł���B
        ''RNTO�𑱂��Ď��s���Ȃ��Ă͂Ȃ�Ȃ�
        SendCommand("RNFR " & sOldFileName)
        If (m_iRetValue <> 350) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''FTP����ޑ��M�@RNTO
        ''RNFR�̌�Ɏ��s�����B
        ''RNFR����ނŎw�肵��̧�ق��A�w�肵��̧�ٖ��ɕύX����B
        ''�ύX��̧�ٖ��̎w��ł���()
        SendCommand("RNTO " & sNewFileName)
        If (m_iRetValue <> 250) Then
            MessageString = m_sReply
            Throw New IOException(m_sReply.Substring(4))
        End If

        ''�߂�l���
        Return bResult

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :̫��ލ쐬����
    '
    ' �Ԃ�l�@ :����I�� - True
    ' �@�@�@    �װ�I�� - False
    '
    ' �������@ :ARG1 - �ިڸ�ؕ�����
    '
    ' �@�\���� :FTP��ď��̫��ނ��쐬���܂�
    '
    ' ���l�@�@ :
    '
    Public Function CreateDirectory(ByVal sDirName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTP����ޑ��M�@MKD
        SendCommand("MKD " & sDirName)
        If (m_iRetValue <> 257) Then
            bResult = False
            MessageString = m_sReply
        End If

        ''�߂�l���
        Return bResult

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :̫��ލ폜����
    '
    ' �Ԃ�l�@ :����I�� - True
    ' �@�@�@    �װ�I�� - False
    '
    ' �������@ :ARG1 - �ިڸ�ؕ�����
    '
    ' �@�\���� :FTP��ď��̫��ނ��폜���܂�
    '
    ' ���l�@�@ :
    '
    Public Function RemoveDirectory(ByVal sDirName As String) As Boolean

        Dim bResult As Boolean

        bResult = True
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTP����ޑ��M�@RMD
        SendCommand("RMD " & sDirName)
        If (m_iRetValue <> 250) Then
            bResult = False
            MessageString = m_sReply
        End If

        ''�߂�l���
        Return bResult

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :̫��ވړ�����
    '
    ' �Ԃ�l�@ :����I�� - True
    ' �@�@�@    �װ�I�� - False
    '
    ' �������@ :ARG1 - �ިڸ�ؕ�����
    '
    ' �@�\���� :FTP��ď��̫��ނ���C��̫��ނֈړ�
    '
    ' ���l�@�@ :
    '
    Public Function ChangeDirectory(ByVal sDirName As String) As Boolean

        Dim bResult As Boolean

        bResult = True

        ''���݈ʒu�m�F
        If (sDirName.Equals(".")) Then
            Exit Function
        End If

        ''۸޲�����
        If (Not (m_bLoggedIn)) Then
            Login()
        End If

        ''FTP����ޑ��M�@̫��ވړ�
        SendCommand("CWD " & sDirName)
        If (m_iRetValue <> 250) Then
            bResult = False
            MessageString = m_sReply
        End If
        Me.m_sRemotePath = sDirName

        ''�߂�l���
        Return bResult

    End Function

    ' @(f)
    '
    ' �@�\�@�@ :FTP�ؒf����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Public Sub CloseConnection()

        If (Not (m_objClientSocket Is Nothing)) Then
            '''FTP����ޑ��M�@�ؒf
            SendCommand("QUIT")
        End If
        Cleanup()

    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�Ǎ�����ײ����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Private Sub ReadReply()
        m_sMes = ""
        m_sReply = ReadLine()
        m_iRetValue = Int32.Parse(m_sReply.Substring(0, 3))
    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�ϐ�����������
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :�Ǎ��ݏ���
    '
    ' �Ԃ�l�@ :����I�� - �Ǎ��ݕ�����
    ' �@�@�@    �װ�I�� - �h�h�󕶎�
    '
    ' �������@ :ARG1 - ү�����׸�
    '
    ' �@�\���� :FTP��Ă��Ǎ��݂��s���܂��B
    '
    ' ���l�@�@ :
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
    ' �@�\�@�@ :FTP����ޑ��M����
    '
    ' �Ԃ�l�@ :�Ȃ�
    '
    ' �������@ :ARG1 - ����ޕ�����
    '
    ' �@�\���� :
    '
    ' ���l�@�@ :
    '
    Private Sub SendCommand(ByVal sCommand As String)

        sCommand = sCommand & ControlChars.CrLf
        Dim cmdbytes As Byte() = ASCII.GetBytes(sCommand)

        m_objClientSocket.Send(cmdbytes, cmdbytes.Length, 0)
        ReadReply()

    End Sub

    ' @(f)
    '
    ' �@�\�@�@ :�Ǎ��ݏ���
    '
    ' �Ԃ�l�@ :����I�� - ���ĕϐ�
    '
    ' �������@ :�Ȃ�
    '
    ' �@�\���� :�ް����Ă��쐬���܂�
    '
    ' ���l�@�@ :
    '
    Private Function CreateDataSocket() As Socket

        Dim index1, index2, len As Int32
        Dim partCount, i, port As Int32
        Dim ipData, buf, ipAddress As String
        Dim parts(6) As Int32
        Dim ch As Char
        Dim s As Socket
        Dim ep As IPEndPoint

        ''FTP����ޑ��M�@PASV
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

        ''�ް��߰����ް�𑪒肵�Ă��������B
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

        ''�w�肳���FTP���ް�ɐڑ��ł����Ȃ�A�ϐ���True�ɂ��܂�
        flag_bool = True
        Return s

    End Function

End Class
