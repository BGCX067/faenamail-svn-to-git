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
Imports System.Drawing
Public Class ThreeDLine
    Inherits System.Windows.Forms.UserControl

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'UserControl overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        '
        'ThreeDLine
        '
        Me.Name = "ThreeDLine"
        Me.Size = New System.Drawing.Size(472, 152)

    End Sub

#End Region

    Public Enum tOrientation
        Horizontal
        Vertical
    End Enum

    Public Property Orientation() As tOrientation
        Get
            Return m_Orientation
        End Get
        Set(ByVal Value As tOrientation)
            m_Orientation = Value
            'Set a "default" height and width
            If m_Orientation = tOrientation.Horizontal Then
                Me.Width = 120
                Me.Height = 2
            Else
                Me.Width = 2
                Me.Height = 100
            End If
        End Set
    End Property

    Private m_Orientation As tOrientation

    Private Sub ThreeDLine_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles MyBase.Paint

        'Draws the line

        'Point Variables
        Dim pL1 As Point
        Dim pL2 As Point
        Dim pL3 As Point
        Dim pL4 As Point

        'Create a pen, you can change the colours if they're wrong.
        Dim DP As New Pen(System.Drawing.SystemColors.ControlDark)
        Dim LP As New Pen(SystemColors.ControlLight)

        'Determine orientation then set the height of the control
        'Set the starting points for first and second lines:
        If m_Orientation = tOrientation.Horizontal Then
            Me.Height = 2 '(2 pixels)
            '-----------------------------
            '|o  -pL1 ------------ pl3- o |
            '| o -pL2 ------------ pl4-  o|
            '-----------------------------
            pL1 = New Point(0, 0)
            pL2 = New Point(1, 1)
            pL3 = New Point(Me.Width - 1, 0)
            pL4 = New Point(Me.Width, 1)
        Else
            Me.Width = 2
            '----------
            '| o-pL1  |
            '| |pL2-o |
            '| |    | |
            '| |    | |
            '| o-pL3| |
            '|  pL4-o |
            '----------
            pL1 = New Point(0, 0)
            pL2 = New Point(1, 1)
            pL3 = New Point(0, Me.Height - 1)
            pL4 = New Point(1, Me.Height)
        End If

        'Draw the lines. Simple.
        e.Graphics.DrawLine(DP, pL1, pL3)
        e.Graphics.DrawLine(LP, pL2, pL4)

    End Sub

    Private Sub ThreeDLine_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
End Class
