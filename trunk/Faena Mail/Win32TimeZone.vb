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
Option Strict On
Option Compare Binary

Imports System
Imports System.Globalization
Imports Microsoft.Win32

' Win32TimeZone
' by Michael R. Brumm
'
' For updates and more information, visit:
' http://www.michaelbrumm.com/simpletimezone.html
'
' or contact me@michaelbrumm.com
'
' Please do not modify this code and re-release it. If you
' require changes to this class, please derive your own class
' from SimpleTimeZone, and add (or override) the methods and
' properties on your own derived class. You never know when 
' your code might need to be version compatible with another
' class that uses Win32TimeZone.

' This should have been part of Microsoft.Win32, so that is
' where I located it.
Namespace Win32


  <Serializable()> Public Class Win32TimeZone
    Inherits SimpleTimeZone


    Private _index As Int32
    Private _displayName As String


    Public Sub New( _
      ByVal index As Int32, _
      ByVal displayName As String, _
      ByVal standardOffset As TimeSpan, _
      ByVal standardName As String, _
      ByVal standardAbbreviation As String _
    )

      MyBase.New( _
        standardOffset, _
        standardName, _
        standardAbbreviation _
        )

      _index = index
      _displayName = displayName

    End Sub

    Public Sub New( _
      ByVal index As Int32, _
      ByVal displayName As String, _
      ByVal standardOffset As TimeSpan, _
      ByVal standardName As String, _
      ByVal standardAbbreviation As String, _
      ByVal daylightDelta As TimeSpan, _
      ByVal daylightName As String, _
      ByVal daylightAbbreviation As String, _
      ByVal daylightTimeChangeStart As DaylightTimeChange, _
      ByVal daylightTimeChangeEnd As DaylightTimeChange _
    )

      MyBase.New( _
        standardOffset, _
        standardName, _
        standardAbbreviation, _
        daylightDelta, _
        daylightName, _
        daylightAbbreviation, _
        daylightTimeChangeStart, _
        daylightTimeChangeEnd _
        )

      _index = index
      _displayName = displayName

    End Sub


    Public ReadOnly Property Index() As Int32
      Get
        Return _index
      End Get
    End Property


    Public ReadOnly Property DisplayName() As String
      Get
        Return _displayName
      End Get
    End Property


    Public Overrides Function ToString() As String
      Return _displayName
    End Function


  End Class


End Namespace
