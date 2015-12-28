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
Option Strict Off
Public Class Win32TimeZoneComparer
    Implements IComparer


    Public Function Compare(ByVal x As Object, ByVal y As Object) As Int32 _
      Implements IComparer.Compare

        If (x Is Nothing) Then

            If (y Is Nothing) Then
                Compare = 0
            Else
                Compare = -1
            End If

        ElseIf (y Is Nothing) Then
            Compare = 1

        Else

            Dim comparison As Int32
            comparison = x.Index.CompareTo(y.Index)

            If (comparison = 0) Then
                comparison = String.Compare(x.DisplayName, y.DisplayName)
            End If

            Return comparison

        End If

    End Function


End Class
