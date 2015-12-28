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

Option Strict On
Option Explicit On 
Imports Microsoft.Win32
Imports System.IO
Module Module1
    Public mainStorage As String
    Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Integer, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Integer) As Integer
    Public Sub RegSaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As String)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\FaenaMail\" + subKey)
        regSubKey.SetValue(name, value)
    End Sub

    Public Sub RegSaveSetting(ByVal subKey As String, ByVal name As String, ByVal value As Long)
        Dim regSubKey As RegistryKey = Registry.LocalMachine.CreateSubKey("Software\Faena\FaenaMail\" + subKey)
        regSubKey.SetValue(name, value)
    End Sub

    Public Function RegGetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As String) As String
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\FaenaMail\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return regSubKey.GetValue(name, defaultValue).ToString
        End If
    End Function
    Public Function RegGetSetting(ByVal subKey As String, ByVal name As String, ByVal defaultValue As Integer) As Integer
        Dim regSubKey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\FaenaMail\" + subKey, False)
        If regSubKey Is Nothing Then
            Return defaultValue
        Else
            Return CInt(regSubKey.GetValue(name, defaultValue))
        End If
    End Function
    Public Function PrepareStr(ByVal strValue As String) As String
        ' This function accepts a string and creates a string that can
        ' be used in a SQL statement by adding single quotes around
        ' it and handling empty values.
        strValue.Replace("'", "''")
        If strValue.Trim() = "" Then
            Return "NULL"
        Else
            Return "'" & strValue.Trim() & "'"
        End If
    End Function

End Module
