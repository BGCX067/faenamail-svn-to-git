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
Public Class ScriptBox
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel1.ContextMenu = Me.ContextMenu1
        Me.Panel1.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.Label1, Me.TextBox1})
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(144, 83)
        Me.Panel1.TabIndex = 1
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.Text = "&Start form here"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.Text = "&Delete"
        '
        'Button1
        '
        Me.Button1.ContextMenu = Me.ContextMenu1
        Me.Button1.Dock = System.Windows.Forms.DockStyle.Right
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button1.Location = New System.Drawing.Point(73, 0)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(69, 17)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Button1"
        '
        'Label1
        '
        Me.Label1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Label1.ContextMenu = Me.ContextMenu1
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(76, 17)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Move me"
        '
        'TextBox1
        '
        Me.TextBox1.AcceptsTab = True
        Me.TextBox1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.TextBox1.Location = New System.Drawing.Point(0, 17)
        Me.TextBox1.Multiline = True
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox1.Size = New System.Drawing.Size(142, 64)
        Me.TextBox1.TabIndex = 6
        Me.TextBox1.Text = "TextBox1"
        '
        'ScriptBox
        '
        Me.ContextMenu = Me.ContextMenu1
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Name = "ScriptBox"
        Me.Size = New System.Drawing.Size(144, 83)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public modified As Boolean
    Public active As Boolean 'see if the script box is highlighted.
    Private posX, posY As Integer
    Public mSettingIf, mSettingAction As Integer
    Public mSettingIfDes, mSettingActionDes, mSettingPath, mSettingMsg, mSettingAttachment, mSettingBcc As String

    Private Sub Label1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseDown
        Me.Cursor.Current = System.Windows.Forms.Cursors.SizeAll
        Me.Focus()
        posX = e.X
        posY = e.Y
    End Sub

    Private Sub Label1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles Label1.MouseMove
        If e.Button = MouseButtons.Left Then
            Me.Left = Me.Left + e.X - posX
            Me.Top = Me.Top + e.Y - posY
        End If
    End Sub

    Private Sub ScriptBox_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Label1.BackColor = Color.LightGray
    End Sub

    Private Sub ScriptBox_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Enter
        Label1.BackColor = Color.LightBlue
        active = True
    End Sub

    Private Sub ScriptBox_Leave(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Leave
        Label1.BackColor = Color.LightGray
        active = False
    End Sub

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim ctl, scriptCtl As Control
        Dim paths() As String
        Dim i As Integer = -1
        For Each ctl In Me.ParentForm.Controls
            If TypeOf ctl Is Panel Then
                For Each scriptCtl In ctl.Controls
                    If TypeOf scriptCtl Is ScriptBox Then
                        i = i + 1
                        ReDim Preserve paths(i)
                        paths(i) = scriptCtl.Name
                    End If
                Next
            End If
        Next
        If Not paths Is Nothing Then
            Select Case Button1.Text
                Case "IfScript"
                    SetScriptIf(paths)
                Case "SQL"
                    SetScriptSQL(paths)
                Case "email"
                    SetScriptEmail(paths)
                Case "MailList"
                    SetMailList(paths)
                Case Else
                    MsgBox("[Button1_Click]Unhandled case:" & Button1.Text, MsgBoxStyle.Critical)
            End Select
        End If
    End Sub

    Private Sub SetMailList(ByVal paths() As String)
        Dim frm As New FormSetMailList()
        frm.TextBox1.Text = Label1.Text
        If mSettingIf = 0 Then
            frm.RadioButton1.Checked = True
        Else
            frm.RadioButton2.Checked = True
        End If
        frm.TextBox2.Text = mSettingIfDes
        Dim path As String
        For Each path In paths
            If Not path Is Nothing Then
                frm.ComboBox5.Items.Add(path)
            End If
        Next
        frm.ComboBox5.Text = mSettingPath
        If frm.ShowDialog() = DialogResult.OK Then
            modified = True
            Label1.Text = frm.TextBox1.Text
            Me.Name = Label1.Text
            TextBox1.Text = frm.TextBox2.Text
            If frm.RadioButton1.Checked = True Then
                mSettingIf = 0
            Else
                mSettingIf = 1
            End If
            mSettingIfDes = frm.TextBox2.Text
            mSettingPath = frm.ComboBox5.Text
        End If
    End Sub

    Private Sub SetScriptIf(ByVal paths() As String)
        Dim frm As New FormSetFilter()
        frm.TextBox1.Text = Label1.Text
        frm.ComboBox1.SelectedIndex = mSettingIf
        frm.ComboBox3.Text = mSettingIfDes
        frm.ComboBox2.SelectedIndex = mSettingAction
        frm.ComboBox4.Text = mSettingActionDes
        Dim path As String
        For Each path In paths
            If Not path Is Nothing Then
                frm.ComboBox5.Items.Add(path)
            End If
        Next
        frm.ComboBox5.Text = mSettingPath
        If frm.ShowDialog() = DialogResult.OK Then
            modified = True
            Label1.Text = frm.TextBox1.Text
            Me.Name = Label1.Text
            TextBox1.Text = frm.TextBox2.Text
            mSettingIf = frm.ComboBox1.SelectedIndex
            mSettingIfDes = frm.ComboBox3.Text
            mSettingAction = frm.ComboBox2.SelectedIndex
            mSettingActionDes = frm.ComboBox4.Text
            mSettingPath = frm.ComboBox5.Text
        End If
    End Sub

    Private Sub SetScriptSQL(ByVal paths() As String)
        Dim frm As New FormSetSQL()
        frm.TextBox1.Text = Label1.Text
        frm.TextBox4.Text = mSettingIfDes
        Dim path As String
        For Each path In paths
            If Not path Is Nothing Then
                frm.ComboBox5.Items.Add(path)
            End If
        Next
        frm.ComboBox5.Text = mSettingPath
        If frm.ShowDialog() = DialogResult.OK Then
            modified = True
            Label1.Text = frm.TextBox1.Text
            Me.Name = Label1.Text
            mSettingIfDes = frm.TextBox4.Text
            mSettingPath = frm.ComboBox5.Text
            TextBox1.Text = frm.TextBox4.Text
        End If
    End Sub

    Private Sub SetScriptEmail(ByVal paths() As String)
        Dim frm As New FormSetemail()
        frm.TextBox1.Text = Label1.Text
        frm.TextBox2.Text = mSettingIfDes
        frm.TextBox3.Text = mSettingActionDes
        frm.TextBox7.Text = mSettingBcc
        Dim path As String
        For Each path In paths
            If Not path Is Nothing Then
                frm.ComboBox5.Items.Add(path)
            End If
        Next
        frm.ComboBox5.Text = mSettingPath
        frm.TextBox4.Text = TextBox1.Text
        frm.TextBox5.Text = mSettingMsg
        frm.TextBox6.Text = mSettingAttachment
        If frm.ShowDialog() = DialogResult.OK Then
            modified = True
            Label1.Text = frm.TextBox1.Text
            Me.Name = Label1.Text
            mSettingIfDes = frm.TextBox2.Text
            mSettingActionDes = frm.TextBox3.Text
            mSettingBcc = frm.TextBox7.Text
            mSettingPath = frm.ComboBox5.Text
            TextBox1.Text = frm.TextBox4.Text
            mSettingMsg = frm.TextBox5.Text
            mSettingAttachment = frm.TextBox6.Text
        End If
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        modified = True
    End Sub

    Private Sub MenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem1.Click
        Me.ParentForm.Tag = Me.Name
        modified = True
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click

        CType(ParentForm, FormScriptWindow).MenuItem14.PerformClick()
    End Sub
End Class
