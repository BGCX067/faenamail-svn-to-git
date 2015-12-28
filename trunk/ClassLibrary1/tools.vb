Imports Microsoft.Win32
Imports System.IO
Public Class tools
    Public Function LoadDSNList() As String()
        Dim a() As String

        Dim RegrootKey As RegistryKey
        Dim RegsubKey As RegistryKey

        RegrootKey = Registry.LocalMachine 'To access Key_Local_Machine registry

        RegsubKey = RegrootKey.OpenSubKey("SOFTWARE\\ODBC\\ODBC.INI\\ODBC Data Sources")

        Dim SNameData As Object
        Dim ODBCDSNName As String
        ' To provide this code to Net 1 developers, I made link this.
        ' Or else you can directly declare dsnNames is For Each Loop.
        Dim i As Integer = 0
        If Not RegsubKey Is Nothing Then
            For Each ODBCDSNName In RegsubKey.GetValueNames() 'Get all the values in the Key
                SNameData = RegsubKey.GetValue(ODBCDSNName) ' Get the data from the Key of Valu
                ReDim Preserve a(i)
                a(i) = ODBCDSNName
                i = i + 1
            Next
            RegsubKey.Close()
            RegsubKey = Nothing
        End If
        RegrootKey.Close()
        RegrootKey = Nothing
        Return a
    End Function

    Public Sub Shell(ByVal parm As String)
        Microsoft.VisualBasic.Shell(parm)
    End Sub
    Public Function IsWin9x() As Boolean
        Dim osInfo As OperatingSystem
        Dim sAns As String
        osInfo = System.Environment.OSVersion
        With osInfo
            Select Case .Platform
                Case .Platform.Win32Windows
                    Return True
                Case Else
                    Return False
            End Select
        End With
    End Function
    Public Function OS_Version() As String
        Dim osInfo As OperatingSystem
        Dim sAns As String
        osInfo = System.Environment.OSVersion
        With osInfo
            Select Case .Platform
                Case .Platform.Win32Windows
                    Select Case (.Version.Minor)
                        Case 0
                            sAns = "Windows 95"
                        Case 10
                            If .Version.Revision.ToString() = "2222A" Then
                                sAns = "Windows 98 Second Edition"
                            Else
                                sAns = "Windows 98"
                            End If
                        Case 90
                            sAns = "Windows Me"
                    End Select
                Case .Platform.Win32NT
                    Select Case (.Version.Major)
                        Case 3
                            sAns = "Windows NT 3.51"
                        Case 4
                            sAns = "Windows NT 4.0"
                        Case 5
                            If .Version.Minor = 0 Then
                                sAns = "Windows 2000"
                            ElseIf .Version.Minor = 1 Then
                                sAns = "Windows XP"
                            ElseIf .Version.Minor = 2 Then
                                sAns = "Windows Server 2003"
                            Else 'Future version maybe update as needed
                                sAns = "Unknown Windows Version"
                            End If
                    End Select
            End Select
        End With
        Return sAns
    End Function

End Class
