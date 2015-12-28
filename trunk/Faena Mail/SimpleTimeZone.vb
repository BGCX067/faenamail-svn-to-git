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
Imports System.Globalization

' SimpleTimeZone
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
' class that uses SimpleTimeZone.

' This should have been part of System.Globalization, so that is
' where I located it.
'Namespace Globalization


' IMPORTANT:
' This class is immutable, and any derived classes
' should also be immutable.
<Serializable()> Public Class DaylightTimeChange


    Private Const NUM_DAYS_IN_WEEK As Int32 = 7

    Private _month As Int32
    Private _dayOfWeek As DayOfWeek
    Private _dayOfWeekIndex As Int32
    Private _timeOfDay As TimeSpan


    ' Constructor without parameters is not allowed.
    Private Sub New()
    End Sub

    ' Constructor allows the definition of a time change
    ' for most time zones using daylight savings time. These
    ' time zones often define the start or end of daylight
    ' savings as "the first Sunday of April, at 2:00am". This
    ' would be constructed as:
    '
    ' New DaylightTimeChange( _
    '   4, _                      ' 4th month: April
    '   DayOfWeek.Sunday, 0, _    ' 1st Sunday
    '   New TimeSpan(2, 0, 0) _   ' at 2:00am
    ' )
    '
    ' "The last Sunday of October, at 2:00am" would be
    ' constructed as:
    '
    ' New DaylightTimeChange( _
    '   10, _                     ' 10th month: October
    '   DayOfWeek.Sunday, 4, _    ' 5th (last) Sunday
    '   New TimeSpan(2, 0, 0) _   ' at 2:00am
    ' )
    '
    Public Sub New( _
      ByVal month As Int32, _
      ByVal dayOfWeek As DayOfWeek, _
      ByVal dayOfWeekIndex As Int32, _
      ByVal timeOfDay As TimeSpan _
    )

        ' Parameter checking
        If ((month < 1) Or (month > 12)) Then
            Throw New ArgumentOutOfRangeException("month", month, "The month must be between 1 and 12, inclusive.")
        End If

        If ( _
            (dayOfWeek < dayOfWeek.Sunday) Or _
            (dayOfWeek > dayOfWeek.Saturday) _
        ) Then
            Throw New ArgumentOutOfRangeException("dayOfWeek", dayOfWeek, "The day of week must be between Sunday and Saturday.")
        End If

        ' 0 = 1st
        ' 1 = 2nd
        ' 2 = 3rd
        ' 3 = 4th
        ' 4 = 5th (last)
        If ((dayOfWeekIndex < 0) Or (dayOfWeekIndex > 4)) Then
            Throw New ArgumentOutOfRangeException("dayOfWeekIndex", dayOfWeekIndex, "The day of week index must be between 0 and 4, inclusive.")
        End If

        If ( _
            (timeOfDay.Ticks < 0) Or _
            (timeOfDay.Ticks >= TimeSpan.TicksPerDay) _
        ) Then
            Throw New ArgumentOutOfRangeException("timeOfDay", timeOfDay, "The time of the day must be less than one day, and not negative.")
        End If

        ' Initialize private storage
        _month = month
        _dayOfWeek = dayOfWeek
        _dayOfWeekIndex = dayOfWeekIndex
        _timeOfDay = timeOfDay

    End Sub


    ' Returns the time and date of the daylight savings change
    ' for a particular year. For example:
    '   "the 1st Sunday of April at 2:00am" for the year "2000"
    '   is "2000/04/02 02:00"
    Public Function GetDate( _
      ByVal year As Int32 _
    ) As DateTime

        If (year < 1) Or (year > DateTime.MaxValue.Year) Then
            Throw New ArgumentOutOfRangeException("year")
        End If

        ' Get the first day of the change month for the specified year.
        Dim resultDate As New DateTime(year, _month, 1)

        ' Get the first day of the month that falls on the
        ' day of the week for this change.
        If (resultDate.DayOfWeek > _dayOfWeek) Then
            resultDate = resultDate.AddDays(NUM_DAYS_IN_WEEK - (resultDate.DayOfWeek - _dayOfWeek))
        ElseIf (resultDate.DayOfWeek < _dayOfWeek) Then
            resultDate = resultDate.AddDays(_dayOfWeek - resultDate.DayOfWeek)
        End If

        ' Get the nth weekday (3rd Tuesday, for example)
        resultDate = resultDate.AddDays(NUM_DAYS_IN_WEEK * _dayOfWeekIndex)

        ' If the date has passed the month, then go back a week. This allows
        ' the 5th weekday to always be the last weekday.
        While (resultDate.Month > _month)
            resultDate = resultDate.AddDays(-NUM_DAYS_IN_WEEK)
        End While

        ' Add the time of day that daylight savings begins.
        resultDate = resultDate.Add(_timeOfDay)

        ' Return the date and time of the change.
        Return resultDate

    End Function


End Class


<Serializable()> Public MustInherit Class SimpleTimeZone
    Inherits TimeZone


    Private _standardAlways As Boolean
    Private _daylightAlwaysWithinStandard As Boolean
    Private _standardAlwaysWithinDaylight As Boolean

    Private _standardOffset As TimeSpan
    Private _standardName As String
    Private _standardAbbreviation As String

    Private _daylightDelta As TimeSpan
    Private _daylightOffset As TimeSpan
    Private _daylightName As String
    Private _daylightAbbreviation As String
    Private _daylightTimeChangeStart As DaylightTimeChange
    Private _daylightTimeChangeEnd As DaylightTimeChange


    ' Constructor without parameters is not allowed.
    Private Sub New()
    End Sub

    ' Constructor for time zone without daylight savings time.
    Public Sub New( _
      ByVal standardOffset As TimeSpan, _
      ByVal standardName As String, _
      ByVal standardAbbreviation As String _
    )

        ' Initialize private storage
        _standardAlways = True

        _standardOffset = standardOffset
        _standardName = standardName
        _standardAbbreviation = standardAbbreviation

    End Sub

    ' Constructor for time zone with or without daylight savings time.
    Public Sub New( _
      ByVal standardOffset As TimeSpan, _
      ByVal standardName As String, _
      ByVal standardAbbreviation As String, _
      ByVal daylightDelta As TimeSpan, _
      ByVal daylightName As String, _
      ByVal daylightAbbreviation As String, _
      ByVal daylightTimeChangeStart As DaylightTimeChange, _
      ByVal daylightTimeChangeEnd As DaylightTimeChange _
    )

        ' Allow non-daylight savings time zones to be created
        ' using this constructor.
        If ( _
          (daylightTimeChangeStart Is Nothing) And _
          (daylightTimeChangeEnd Is Nothing) _
          ) Then

            ' Initialize private storage
            _standardAlways = True

            _standardOffset = standardOffset
            _standardName = standardName
            _standardAbbreviation = standardAbbreviation

            Exit Sub

        End If

        ' If the time zone has a start OR an end, then it
        ' must have a start AND an end.
        If (daylightTimeChangeStart Is Nothing) Then
            Throw New ArgumentNullException("daylightTimeChangeStart")
        End If

        If (daylightTimeChangeEnd Is Nothing) Then
            Throw New ArgumentNullException("daylightTimeChangeEnd")
        End If

        ' Initialize private storage
        _standardAlways = False

        _standardOffset = standardOffset
        _standardName = standardName
        _standardAbbreviation = standardAbbreviation

        _daylightDelta = daylightDelta
        _daylightOffset = _standardOffset.Add(daylightDelta)
        _daylightName = daylightName
        _daylightAbbreviation = daylightAbbreviation

        ' These referance types are immutable, so they cannot be
        ' changed outside this class' scope, and thus can be
        ' permanently referenced.
        _daylightTimeChangeStart = daylightTimeChangeStart
        _daylightTimeChangeEnd = daylightTimeChangeEnd

    End Sub


    Public Overrides ReadOnly Property StandardName() As String
        Get
            Return _standardName
        End Get
    End Property


    Public Overridable ReadOnly Property StandardAbbreviation() As String
        Get
            Return _standardAbbreviation
        End Get
    End Property


    Public Overrides ReadOnly Property DaylightName() As String
        Get
            Return _daylightName
        End Get
    End Property


    Public Overridable ReadOnly Property DaylightAbbreviation() As String
        Get
            Return _daylightAbbreviation
        End Get
    End Property


    ' The name is dependant on whether the time zone is in daylight
    ' savings time or not. This method can be ambiguous during
    ' daylight changes.
    Public Overridable Function GetNameLocalTime( _
      ByVal time As DateTime _
    ) As String

        If (_standardAlways) Then
            Return _standardName
        ElseIf (IsDaylightSavingTime(time)) Then
            Return _daylightName
        Else
            Return _standardName
        End If

    End Function

    ' This method is unambiguous during daylight changes.
    Public Overridable Function GetNameUniversalTime( _
      ByVal time As DateTime _
    ) As String

        If (IsDaylightSavingTimeUniversalTime(time)) Then
            Return _daylightName
        Else
            Return _standardName
        End If

    End Function


    ' The abbreviation is dependant on whether the time zone is in
    ' daylight savings time or not. This method can be ambiguous during
    ' daylight changes.
    Public Overridable Function GetAbbreviationLocalTime( _
      ByVal time As DateTime _
    ) As String

        If (_standardAlways) Then
            Return _standardAbbreviation
        ElseIf (IsDaylightSavingTime(time)) Then
            Return _daylightAbbreviation
        Else
            Return _standardAbbreviation
        End If

    End Function

    ' This method is unambiguous during daylight changes.
    Public Overridable Function GetAbbreviationUniversalTime( _
      ByVal time As DateTime _
    ) As String

        If (IsDaylightSavingTimeUniversalTime(time)) Then
            Return _daylightAbbreviation
        Else
            Return _standardAbbreviation
        End If

    End Function


    Public Overrides Function GetDaylightChanges( _
      ByVal year As Int32 _
    ) As DaylightTime

        If (year < 1) Or (year > DateTime.MaxValue.Year) Then
            Throw New ArgumentOutOfRangeException("year")
        End If

        If (_standardAlways) Then
            Return Nothing

        Else
            Return New DaylightTime( _
             _daylightTimeChangeStart.GetDate(year), _
             _daylightTimeChangeEnd.GetDate(year), _
             _daylightDelta _
            )
        End If

    End Function


    ' This method can be ambiguous during daylight changes.
    Public Overloads Overrides Function IsDaylightSavingTime( _
      ByVal time As DateTime _
    ) As Boolean

        Return IsDaylightSavingTime(time, False)

    End Function


    ' This method is unambiguous during daylight changes.
    Public Overridable Function IsDaylightSavingTimeUniversalTime( _
      ByVal time As DateTime _
    ) As Boolean

        time = time.Add(_standardOffset)
        Return IsDaylightSavingTime(time, True)

    End Function


    Private Overloads Function IsDaylightSavingTime( _
      ByVal time As DateTime, _
      ByVal fromUtcTime As Boolean _
    ) As Boolean

        ' If this time zone is never in daylight savings, then
        ' return false.
        If (_standardAlways) Then
            Return False
        End If

        ' Get the daylight savings time start and end for this
        ' time's year.
        Dim daylightTimes As DaylightTime
        daylightTimes = GetDaylightChanges(time.Year)

        ' Return whether the time is within the daylight savings
        ' time for this year.
        Return IsDaylightSavingTime(time, daylightTimes, fromUtcTime)

    End Function


    Public Overloads Shared Function IsDaylightSavingTime( _
       ByVal time As DateTime, _
       ByVal daylightTimes As DaylightTime _
    ) As Boolean

        Return IsDaylightSavingTime(time, daylightTimes, False)

    End Function


    Private Overloads Shared Function IsDaylightSavingTime( _
      ByVal time As DateTime, _
      ByVal daylightTimes As DaylightTime, _
      ByVal fromUtcTime As Boolean _
    ) As Boolean

        ' Mirrors .NET Framework TimeZone functionality, which 
        ' does not throw an exception.
        If (daylightTimes Is Nothing) Then
            Return False
        End If

        Dim daylightStart As DateTime
        Dim daylightEnd As DateTime
        Dim daylightDelta As TimeSpan
        daylightStart = daylightTimes.Start
        daylightEnd = daylightTimes.End
        daylightDelta = daylightTimes.Delta

        ' Northern hemisphere (normally)
        ' The daylight savings time of the year falls between the
        ' start and the end dates.
        If (daylightStart < daylightEnd) Then

            ' The daylight savings time of the year falls between the
            ' start and the end dates.
            If ( _
              (time >= daylightStart) And _
              (time < daylightEnd) _
              ) Then

                ' If the time was taken from a UTC time, then do not apply
                ' the backward compatibility.
                If (fromUtcTime) Then
                    Return True

                    ' Backward compatiblity with .NET Framework TimeZone.
                    ' If the daylight savings delta is positive, then there is a
                    ' period of time which does not exist (between 2am and 3am in
                    ' most daylight savings time zones) at the beginning of the
                    ' daylight savings. This period of non-existant time should be 
                    ' considered standard time (not daylight savings).
                Else

                    If (daylightDelta.Ticks > 0) Then
                        If (time < (daylightStart.Add(daylightDelta))) Then
                            Return False
                        Else
                            Return True
                        End If
                    Else
                        Return True
                    End If

                End If

                ' Otherwise, the time and date is not within daylight
                ' savings time.
            Else

                ' If the time was taken from a UTC time, then do not apply
                ' the backward compatibility.
                If (fromUtcTime) Then
                    Return False

                    ' Backward compatiblity with .NET Framework TimeZone.
                    ' If the daylight savings delta is negative (which shouldn't
                    ' happen), then there is a period of time which does not exist
                    ' (between 2am and 3am in most daylight savings time zones).
                    ' at the end of daylight savings. This period of
                    ' non-existant time should be considered daylight savings.
                Else

                    If (daylightDelta.Ticks < 0) Then

                        If ( _
                          (time >= daylightEnd) And _
                          (time < daylightEnd.Subtract(daylightDelta)) _
                          ) Then
                            Return True
                        Else
                            Return False
                        End If

                    Else
                        Return False
                    End If

                End If

            End If

            ' Southern hemisphere (normally)
            ' The daylight savings time of the year is after the start,
            ' or before the end, but not between the two dates.
        Else

            ' The daylight savings time of the year is after the start,
            ' or before the end, but not between the two dates.
            If (time >= daylightStart) Then

                ' If the time was taken from a UTC time, then do not apply
                ' the backward compatibility.
                If (fromUtcTime) Then
                    Return True

                    ' Backward compatiblity with .NET Framework TimeZone.
                    ' If the daylight savings delta is positive, then there is a
                    ' period of time which does not exist (between 2am and 3am in
                    ' most daylight savings time zones) at the beginning of the
                    ' daylight savings. This period of non-existant time should be 
                    ' considered standard time (not daylight savings).
                Else

                    If (daylightDelta.Ticks > 0) Then
                        If (time < (daylightStart.Add(daylightDelta))) Then
                            Return False
                        Else
                            Return True
                        End If
                    Else
                        Return True
                    End If

                End If

                ' The current time is before the end of daylight savings, so
                ' it is during daylight savings.
            ElseIf (time < daylightEnd) Then
                Return True

                ' Otherwise, the time and date is not within daylight
                ' savings time.
            Else

                ' If the time was taken from a UTC time, then do not apply
                ' the backward compatibility.
                If (fromUtcTime) Then
                    Return False

                    ' Backward compatiblity with .NET Framework TimeZone.
                    ' If the daylight savings delta is negative (which shouldn't
                    ' happen), then there is a period of time which does not exist
                    ' (between 2am and 3am in most daylight savings time zones).
                    ' at the end of daylight savings. This period of
                    ' non-existant time should be considered daylight savings.
                Else

                    If (daylightDelta.Ticks < 0) Then

                        If ( _
                          (time >= daylightEnd) And _
                          (time < daylightEnd.Subtract(daylightDelta)) _
                          ) Then
                            Return True
                        Else
                            Return False
                        End If

                    Else
                        Return False
                    End If

                End If

            End If

        End If

    End Function


    Public Overrides Function GetUtcOffset( _
      ByVal time As DateTime _
    ) As TimeSpan

        ' If this time zone is never in daylight savings, then
        ' return the standard offset.
        If (_standardAlways) Then
            Return _standardOffset

            ' If the time zone is in daylight savings, then return
            ' the daylight savings offset.
        ElseIf (IsDaylightSavingTime(time)) Then
            Return _daylightOffset

            ' Otherwise, return the standard offset.
        Else
            Return _standardOffset
        End If

    End Function


    Public Overrides Function ToLocalTime( _
      ByVal time As DateTime _
    ) As DateTime

        time = time.Add(_standardOffset)

        If (Not (_standardAlways)) Then
            If (IsDaylightSavingTime(time, True)) Then
                time = time.Add(_daylightDelta)
            End If
        End If

        Return time

    End Function


    ' This can return an incorrect time during the time change
    ' between standard and daylight savings time, because
    ' times near the daylight savings switch can be ambiguous.
    '
    ' For example, if daylight savings ends at:
    ' "2000/10/29 02:00", and fall back an hour, then is:
    ' "2000/10/29 01:30", during daylight savings, or not?
    '
    ' Consequently, this function is provided for backwards
    ' compatiblity only, and should be deprecated and replaced
    ' with the overload that allows daylight savings to be
    ' specified.
    Public Overloads Overrides Function ToUniversalTime( _
      ByVal time As DateTime _
    ) As DateTime

        If (_standardAlways) Then
            Return time.Subtract(_standardOffset)

        Else

            If (IsDaylightSavingTime(time)) Then
                Return time.Subtract(_daylightOffset)
            Else
                Return time.Subtract(_standardOffset)
            End If

        End If


    End Function


    ' This overload allows the status of daylight savings
    ' to be specified along with the time. This conversion
    ' is unambiguous and always correct.
    Public Overloads Function ToUniversalTime( _
      ByVal time As DateTime, _
      ByVal daylightSavings As Boolean _
    ) As DateTime

        If (_standardAlways) Then
            Return time.Subtract(_standardOffset)

        Else

            If (daylightSavings) Then
                Return time.Subtract(_daylightOffset)
            Else
                Return time.Subtract(_standardOffset)
            End If

        End If

    End Function


End Class

'End Namespace
