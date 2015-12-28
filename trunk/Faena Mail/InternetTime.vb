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
Option Explicit On 
Module InternetTime

    Private Structure SYSTEMTIME
        Dim wYear As Integer
        Dim wMonth As Integer
        Dim wDayOfWeek As Integer
        Dim wDay As Integer
        Dim wHour As Integer
        Dim wMinute As Integer
        Dim wSecond As Integer
        Dim wMilliseconds As Integer
    End Structure

    Private Structure FILETIME
        Dim dwLowDateTime As Long
        Dim dwHighDateTime As Long
    End Structure
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Declare Function InternetTimeToSystemTime Lib "wininet.dll" _
           (ByVal lpszTime As String, _
            ByRef pst As SYSTEMTIME, _
            ByVal dwReserved As Long) _
            As Long
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '############################################################################################
    'Experimental
    Private Declare Function FileTimeToSystemTime Lib "kernel32" _
           (ByVal lpFileTime As FILETIME, _
    ByVal lpSystemTime As SYSTEMTIME) _
            As Long
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    Private Declare Function MimeOleInetDateToFileTime Lib "inetcomm.dll" _
           (ByVal lpszTime As String, _
            ByRef FT As FILETIME) _
            As Long
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '############################################################################################
    Public Function InternetTimeToVBTime(ByVal sInternetTime As String) As Date
        '********************************************************************************
        'Purpose     :Just convert internet date/time to VB Date data type
        'Arguments   :string with date/time in internet format such as
        '             "Mon, 06 Jul 1998 11:10:46 GMT"
        '            :returns a GMT date
        '********************************************************************************
        Dim uSystemTime As SYSTEMTIME
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        Call InternetTimeToSystemTime(sInternetTime, uSystemTime, 0&)
        With uSystemTime
            InternetTimeToVBTime = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & _
                                         TimeSerial(.wHour, .wMinute, .wSecond))
        End With
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End Function
    '############################################################################################
    Public Function MailTimeToVBTime1(ByVal sMailTime As String) As Date
        Dim uSystemTime As SYSTEMTIME
        Dim sWWW As String
        Dim dGMT As Long
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        sWWW = sMailTime
        If (InStr(1, sWWW, ",") = 0) Then sWWW = "Thu, " & sWWW
        Call InternetTimeToSystemTime(sWWW, uSystemTime, 0&)
        With uSystemTime
            MailTimeToVBTime1 = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & _
                                         TimeSerial(.wHour, .wMinute, .wSecond))
        End With
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        If ((Mid(sWWW, 27, 1) = "-") Or (Mid(sWWW, 27, 1) = "+")) Then dGMT = CLng(Val(Mid(sWWW, 28, 2)) + (Val(Mid(sWWW, 30, 2)) * 100 / 60))
        If (Mid(sWWW, 27, 1) = "-") Then MailTimeToVBTime1 = DateAdd("H", dGMT, MailTimeToVBTime1)
        If (Mid(sWWW, 27, 1) = "+") Then MailTimeToVBTime1 = DateAdd("H", -dGMT, MailTimeToVBTime1)
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End Function
    '############################################################################################
    'Use this for email 
    Public Function MailTimeToVBTime2(ByVal sMailTime As String) As Date
        Dim uFileTime As FILETIME
        Dim uSystemTime As SYSTEMTIME
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        MimeOleInetDateToFileTime(sMailTime, uFileTime)
        FileTimeToSystemTime(uFileTime, uSystemTime)
        With uSystemTime
            MailTimeToVBTime2 = CDate(DateSerial(.wYear, .wMonth, .wDay) & " " & _
                                         TimeSerial(.wHour, .wMinute, .wSecond))
        End With
        '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    End Function
    '############################################################################################

End Module
