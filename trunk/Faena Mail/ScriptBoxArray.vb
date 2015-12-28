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
Public Class ScriptBoxArray
    Inherits System.Collections.CollectionBase
    Private ReadOnly HostPanel As System.Windows.Forms.Panel
    Public Function AddNewScriptBox() As ScriptBox
        ' Create a new instance of the Button class.
        Dim aScriptBox As New ScriptBox()
        aScriptBox.Top = 100
        aScriptBox.Left = 100
        Me.List.Add(aScriptBox)
        aScriptBox.Tag = Me.Count
        aScriptBox.Label1.Text = "ScriptBox " & Me.Count.ToString
        aScriptBox.Name = aScriptBox.Label1.Text
        ' Add the button to the controls collection of the form 
        ' referenced by the HostForm field.
        HostPanel.Controls.Add(aScriptBox)
        aScriptBox.BringToFront()
        ' Add the button to the collection's internal list.
        ' Set intial properties for the button object.
        Return aScriptBox
    End Function

    Public Sub New(ByVal host As System.Windows.Forms.Panel)
        HostPanel = host
    End Sub

    Default Public ReadOnly Property Item(ByVal Index As Integer) As ScriptBox
        Get
            Return CType(Me.List.Item(Index), ScriptBox)
        End Get
    End Property

    Public Sub Remove()
        ' Check to be sure there is a button to remove.
        Dim i As Integer
        For i = Me.Count - 1 To 0 Step -1

            If Me(i).active Then
                ' Remove the last button added to the array from the host form 
                ' controls collection. Note the use of the default property in 
                ' accessing the array.
                HostPanel.Controls.Remove(Me(i))
                Me.List.RemoveAt(i)
            End If
        Next
    End Sub

End Class
