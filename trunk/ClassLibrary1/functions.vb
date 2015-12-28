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
Imports System.Data.Odbc
Public Class Functions
    Dim objConnection As OdbcConnection
    Public Function GetCommandObject(ByVal strCommand As String, ByVal ExecuteIt As Boolean) As OdbcCommand
        Dim objCommand As New OdbcCommand(strCommand, objConnection)
        If ExecuteIt = False Then
            Return objCommand
        Else
            objCommand.ExecuteNonQuery()
        End If
    End Function

    Public Function GetConnectionObject() As OdbcConnection
        Return objConnection
    End Function

    Public Function GetReaderObject(ByVal strCommand As String) As OdbcDataReader
        Dim objCommand As OdbcCommand = GetCommandObject(strCommand, False)
        Dim objReader As OdbcDataReader = objCommand.ExecuteReader(CommandBehavior.CloseConnection)
        Return objReader
        objReader.Close()
    End Function

    Public Function EscapeString(ByVal strIn As String) As String
        Return strIn.Replace("'", "''")
    End Function

    Public Sub New(Optional ByVal odbcDSN As String = "")
        If odbcDSN = "" Then odbcDSN = reg.GetSetting("Faena Mail", "ODBCDSN", "")
        objConnection = New OdbcConnection("DSN=" + odbcDSN)
        objConnection.Open()
    End Sub

End Class

