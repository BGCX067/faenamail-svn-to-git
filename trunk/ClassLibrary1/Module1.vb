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
Imports VB = Microsoft.VisualBasic
Module Module1
    Public reg As New ClassRegistry
    Private lastcheck1 As Double
    Private lastcheck2 As Double
    Public mainStorage As String
    Public Function EncPW(ByVal strIn As String) As String
        Dim i As Integer
        EncPW = "!"
        For i = 1 To Len(strIn)
            Select Case i Mod 3
                Case 0
                    EncPW = EncPW & Chr(Asc(Mid$(strIn, i, 1)) + 1)
                Case 1
                    EncPW = EncPW & Chr(Asc(Mid$(strIn, i, 1)) - 2)
                Case 2
                    EncPW = EncPW & Chr(Asc(Mid$(strIn, i, 1)) + 3)
                Case Else
                    EncPW = EncPW & Chr(Asc(Mid$(strIn, i, 1)) - 4)
            End Select
        Next
    End Function

    Public Function DecPW(ByVal strIn As String) As String
        Dim i As Integer
        If Left$(strIn, 1) = "!" Then
            strIn = Mid$(strIn, 2)
            For i = 1 To Len(strIn)
                Select Case i Mod 3
                    Case 0
                        DecPW = DecPW & Chr(Asc(Mid$(strIn, i, 1)) - 1)
                    Case 1
                        DecPW = DecPW & Chr(Asc(Mid$(strIn, i, 1)) + 2)
                    Case 2
                        DecPW = DecPW & Chr(Asc(Mid$(strIn, i, 1)) - 3)
                    Case Else
                        DecPW = DecPW & Chr(Asc(Mid$(strIn, i, 1)) + 4)
                End Select
            Next
        Else
            DecPW = strIn
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
    Function CheckSN1(ByRef strIn As String) As Boolean
        On Error Resume Next
        Dim i As Short
        Dim l As Integer
        Dim strTemp As String
        l = 2002 'the Seed #1
        CheckSN1 = False
        If Len(strIn) <> 14 Then
            GoTo Label1Exit
        End If
        If System.Math.Abs(VB.Timer() - lastcheck1) < 0.1 Then
            If Rnd() < 0.1 Then CheckSN1 = True
            GoTo Label1Exit
        End If
        For i = 1 To 3
            l = l + Asc(Mid(strIn, i, 1))
        Next
        l = l Mod 26
        If 65 + l <> Asc(Mid(strIn, 4, 1)) Then
            GoTo Label1Exit
        End If
        For i = 6 To 8
            l = l + Asc(Mid(strIn, i, 1))
        Next
        l = l Mod 26
        If 65 + l <> Asc(Mid(strIn, 9, 1)) Then
            GoTo Label1Exit
        End If
        l = 0 'the Seed #2
        For i = 11 To 13
            l = l + Asc(Mid(strIn, i, 1))
        Next
        l = l Mod 26
        If 65 + l <> Asc(Mid(strIn, 14, 1)) Then
            GoTo Label1Exit
        End If
        strTemp = Chr(Asc(Mid(strIn, 5, 1)) + 1)
        strTemp = Chr(Asc(strTemp) - 1)
        If strTemp <> "-" Then GoTo Label1Exit
        CheckSN1 = True
Label1Exit:
        lastcheck1 = VB.Timer()
    End Function

    Function CheckSN2(ByRef strIn As String) As Boolean
        On Error Resume Next
        Dim i As Integer
        Dim l As Integer
        Dim strTemp As String
        CheckSN2 = False
        If System.Math.Abs(VB.Timer() - lastcheck2) < 0.1 Then
            If Rnd() < 0.1 Then
                CheckSN2 = True
            End If
            GoTo Label1Exit
        End If
        l = 2002 'the Seed #1
        For i = 1 To 3
            l = l + Asc(Mid(strIn, i, 1))
        Next
        l = l Mod 26
        If 65 + l <> Asc(Mid(strIn, 4, 1)) Then
            GoTo Label1Exit
        End If
        For i = 6 To 8
            l = l + Asc(Mid(strIn, i, 1))
        Next
        l = l Mod 26
        If 65 + l <> Asc(Mid(strIn, 9, 1)) Then
            GoTo Label1Exit
        End If
        l = 0 'the Seed #2
        For i = 11 To 13
            l = l + Asc(Mid(strIn, i, 1))
        Next
        l = l Mod 26
        If 65 + l <> Asc(Mid(strIn, 14, 1)) Then
            GoTo Label1Exit
        End If
        strTemp = Chr(Asc(Mid(strIn, 5, 1)) + 1)
        strTemp = Chr(Asc(strTemp) - 1)
        If strTemp <> "-" Then GoTo Label1Exit
        CheckSN2 = True
Label1Exit:
        lastcheck2 = VB.Timer()
    End Function

    Public Function TextFromMime(ByVal strMime As String) As String
        Try
            If strMime.IndexOf("This is a multi-part message in MIME format.") >= 0 _
            Or strMime.IndexOf("This is a multipart message in MIME format") >= 0 _
            Or strMime.IndexOf("-----=_Part_") >= 0 Then
                Dim pos As Integer = 0
                Dim strPart As String = ""
                Dim charset As String = ""
                Dim encoding As String = ""
                Dim noattach As String = ""
                pos = strMime.IndexOf("Content-Type: text/plain;")
                If pos >= 0 Then
                    strPart = strMime.Substring(pos)
                    pos = strPart.IndexOf("charset")
                    If pos >= 0 Then
                        charset = strPart.Substring(pos)
                        charset = charset.Substring(0, charset.IndexOf(vbLf)).Trim
                        strPart = strPart.Substring(pos)
                        strPart = strPart.Substring(strPart.IndexOf(vbLf))
                    End If
                    pos = strPart.IndexOf("Content-Transfer-Encoding")
                    If pos >= 0 Then
                        encoding = strPart.Substring(pos)
                        encoding = encoding.Substring(0, encoding.IndexOf(vbLf)).Trim
                        strPart = strPart.Substring(pos)
                        strPart = strPart.Substring(strPart.IndexOf(vbLf))
                    End If
                    pos = strPart.IndexOf(vbLf & "------=_NextPart")
                    If pos > 4 Then
                        strPart = strPart.Substring(3, pos - 4)
                    Else
                        pos = strPart.IndexOf(vbLf & "------=_Part_")
                        If pos > 4 Then
                            strPart = strPart.Substring(3, pos - 4)
                        Else
                            strPart = strPart.Substring(3)
                        End If
                    End If
                    pos = strPart.IndexOf(vbLf & "Content-Type: application")
                    If pos > 0 Then
                        strPart = strPart.Substring(0, strPart.Substring(pos).IndexOf(vbLf & vbCr & vbLf) + pos)
                        strPart = strPart & "( ... Attachment blocked.)"
                    End If
                    Return strPart
                Else
                    Return strMime
                End If
            Else
                Return strMime
            End If
        Catch
            Return strMime
        End Try
    End Function
    'Public Sub Debuglog(ByVal ex As Exception)
    '    Dim el As New ErrorLog
    '    el.LogException(ex, True)
    'End Sub

    Public Function DBStr(ByVal field As Object) As String
        If IsDBNull(field) Then
            Return ""
        ElseIf field Is Nothing Then
            Return ""
        Else
            Return field.ToString.Trim
        End If
    End Function

End Module
