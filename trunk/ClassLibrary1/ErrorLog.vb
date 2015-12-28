'************************************CONFIDENTIAL**************************************
' All information in this document and associated documents are confidential.
' It may not be disclosed in any manner without the express prior written
' permission of Faena Technologies.
'
'
'*********************************COPYRIGHT NOTICE*************************************
'
'  Copyright (C) by Faena Technologies. All rights reserved.
'
' No part of this software may be used, stored, compiled, reproduced,
' modified, transcribed, translated, transmitted, or transferred, in
' any form or by any means  whether electronic, mechanical,  magnetic,
' optical, or otherwise, without the express prior written permission
'                          of Faena Technologies.
' The license and distribution terms for any available version or derivative of
' this code cannot be changed.
'
'*********************************OTHER INFO*******************************************
'Requires Visual Studio .NET 1.1 (or greater).
'**************************************************************************************
Imports System.IO
Imports System.Windows.Forms
Public Class ErrorLog

    Public Function LogException(ByVal ex As Exception, ByVal writeFile As Boolean) As String
        Dim strErr As String = ""
        Try
            strErr = vbNewLine + "===========Error Log Entry=======" + Now.ToString + vbNewLine
            strErr += ex.ToString + vbNewLine + vbNewLine
            strErr += "ProductName:" + Application.ProductName.ToString + vbNewLine + vbNewLine
            strErr += "ProductVersion:" + Application.ProductVersion.ToString + vbNewLine + vbNewLine
            strErr += "ExecutablePath:" + Application.ExecutablePath.ToString + vbNewLine + vbNewLine
            strErr += "StartupPath:" + Application.StartupPath.ToString + vbNewLine + vbNewLine
            strErr += "Culture:" + Application.CurrentCulture.ToString + vbNewLine + vbNewLine
            strErr += "OS:" + OS_Version() + vbNewLine + vbNewLine
            strErr += "OS Version:" + System.Environment.OSVersion.ToString + vbNewLine
            If writeFile Then
                Dim fw As StreamWriter = New StreamWriter(Application.StartupPath + "\Error.log", True)
                fw.WriteLine(strErr)
                fw.Close()
            End If
        Catch ex1 As Exception
            MessageBox.Show("Error writing log file: " + ex1.Message, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
        Return strErr
    End Function

    Public Function OS_Version() As String
        Try
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
                                Else 'Future version maybe update
                                    'as needed
                                    sAns = "Unknown Windows Version"
                                End If
                        End Select
                End Select
            End With
            Return sAns
        Catch ex As Exception
            Return ex.ToString
        End Try
    End Function

End Class
