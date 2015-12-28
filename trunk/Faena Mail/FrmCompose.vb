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
Option Compare Text
Option Explicit On 
Imports System.Xml
Imports System.IO
Imports System.Text
Public Class FrmCompose
    Inherits System.Windows.Forms.Form
    Private oldMessageID As String
    Private _mailType As String
    'Private _origtop As Integer
    Private _origHeight As Integer

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal strMessageID As String, ByVal mailtype As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        oldMessageID = strMessageID
        _mailType = mailtype
    End Sub

    'Form overrides dispose to clean up the component list.
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
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents BtnAttach As System.Windows.Forms.Button
    Friend WithEvents ComboBox5 As System.Windows.Forms.ComboBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox4 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox3 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox2 As System.Windows.Forms.ComboBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.BtnAttach = New System.Windows.Forms.Button
        Me.ComboBox5 = New System.Windows.Forms.ComboBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label6 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.ComboBox4 = New System.Windows.Forms.ComboBox
        Me.ComboBox3 = New System.Windows.Forms.ComboBox
        Me.ComboBox2 = New System.Windows.Forms.ComboBox
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 248)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(420, 22)
        Me.StatusBar1.TabIndex = 0
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem4})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem6, Me.MenuItem8, Me.MenuItem7})
        Me.MenuItem1.Text = "&File"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 0
        Me.MenuItem6.Text = "Save to schedular"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.Text = "-"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 2
        Me.MenuItem7.Text = "&Close"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5, Me.MenuItem9})
        Me.MenuItem4.Text = "&Insert"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem14, Me.MenuItem13})
        Me.MenuItem5.Text = "&Text"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 0
        Me.MenuItem14.Text = "&Select text file..."
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 1
        Me.MenuItem13.Text = "-"
        Me.MenuItem13.Visible = False
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 1
        Me.MenuItem9.Text = "&File attachment..."
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        Me.ErrorProvider1.DataMember = ""
        '
        'RichTextBox1
        '
        Me.RichTextBox1.AcceptsTab = True
        Me.RichTextBox1.AllowDrop = True
        Me.RichTextBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.RichTextBox1.Location = New System.Drawing.Point(0, 160)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ShowSelectionMargin = True
        Me.RichTextBox1.Size = New System.Drawing.Size(420, 88)
        Me.RichTextBox1.TabIndex = 17
        Me.RichTextBox1.Text = ""
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.Filter = "All files|*.*"
        Me.OpenFileDialog1.Multiselect = True
        Me.OpenFileDialog1.Title = "Insert File"
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem2, Me.MenuItem3})
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 0
        Me.MenuItem2.Text = "Open"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 1
        Me.MenuItem3.Text = "Remove"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.ListView1)
        Me.Panel1.Controls.Add(Me.BtnAttach)
        Me.Panel1.Controls.Add(Me.ComboBox5)
        Me.Panel1.Controls.Add(Me.TextBox4)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.TextBox2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.Label4)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.TextBox3)
        Me.Panel1.Controls.Add(Me.TextBox1)
        Me.Panel1.Controls.Add(Me.ComboBox4)
        Me.Panel1.Controls.Add(Me.ComboBox3)
        Me.Panel1.Controls.Add(Me.ComboBox2)
        Me.Panel1.Controls.Add(Me.ComboBox1)
        Me.Panel1.Controls.Add(Me.DateTimePicker1)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel1.Location = New System.Drawing.Point(0, 0)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(420, 160)
        Me.Panel1.TabIndex = 21
        '
        'ListView1
        '
        Me.ListView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.ContextMenu = Me.ContextMenu1
        Me.ListView1.Location = New System.Drawing.Point(88, 136)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(316, 19)
        Me.ListView1.TabIndex = 37
        Me.ListView1.View = System.Windows.Forms.View.List
        '
        'BtnAttach
        '
        Me.BtnAttach.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BtnAttach.Location = New System.Drawing.Point(8, 136)
        Me.BtnAttach.Name = "BtnAttach"
        Me.BtnAttach.Size = New System.Drawing.Size(75, 20)
        Me.BtnAttach.TabIndex = 36
        Me.BtnAttach.Text = "Attach..."
        '
        'ComboBox5
        '
        Me.ComboBox5.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox5.Items.AddRange(New Object() {"Email Address(es)", "Mail List"})
        Me.ComboBox5.Location = New System.Drawing.Point(64, 68)
        Me.ComboBox5.Name = "ComboBox5"
        Me.ComboBox5.Size = New System.Drawing.Size(107, 21)
        Me.ComboBox5.TabIndex = 29
        '
        'TextBox4
        '
        Me.TextBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox4.Location = New System.Drawing.Point(176, 68)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(226, 20)
        Me.TextBox4.TabIndex = 30
        Me.TextBox4.Text = ""
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label5.Location = New System.Drawing.Point(8, 68)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(26, 16)
        Me.Label5.TabIndex = 28
        Me.Label5.Text = "Bcc:"
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.Location = New System.Drawing.Point(176, 44)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(226, 20)
        Me.TextBox2.TabIndex = 27
        Me.TextBox2.Text = ""
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Location = New System.Drawing.Point(8, 48)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(21, 16)
        Me.Label1.TabIndex = 25
        Me.Label1.Text = "Cc:"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.Location = New System.Drawing.Point(8, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(34, 16)
        Me.Label3.TabIndex = 20
        Me.Label3.Text = "From:"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label6.Location = New System.Drawing.Point(8, 112)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(45, 16)
        Me.Label6.TabIndex = 34
        Me.Label6.Text = "Subject:"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.Location = New System.Drawing.Point(8, 92)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(55, 16)
        Me.Label4.TabIndex = 31
        Me.Label4.Text = "Schedule:"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label2.Location = New System.Drawing.Point(8, 28)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(21, 16)
        Me.Label2.TabIndex = 22
        Me.Label2.Text = "To:"
        '
        'TextBox3
        '
        Me.TextBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox3.Location = New System.Drawing.Point(64, 112)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(340, 20)
        Me.TextBox3.TabIndex = 35
        Me.TextBox3.Text = ""
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.Location = New System.Drawing.Point(176, 24)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(226, 20)
        Me.TextBox1.TabIndex = 24
        Me.TextBox1.Text = ""
        '
        'ComboBox4
        '
        Me.ComboBox4.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox4.Items.AddRange(New Object() {"Email Address(es)", "Mail List"})
        Me.ComboBox4.Location = New System.Drawing.Point(64, 44)
        Me.ComboBox4.Name = "ComboBox4"
        Me.ComboBox4.Size = New System.Drawing.Size(107, 21)
        Me.ComboBox4.TabIndex = 26
        '
        'ComboBox3
        '
        Me.ComboBox3.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox3.Items.AddRange(New Object() {"Email Address(es)", "Mail List"})
        Me.ComboBox3.Location = New System.Drawing.Point(64, 24)
        Me.ComboBox3.Name = "ComboBox3"
        Me.ComboBox3.Size = New System.Drawing.Size(107, 21)
        Me.ComboBox3.TabIndex = 23
        '
        'ComboBox2
        '
        Me.ComboBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox2.DisplayMember = "AccountName"
        Me.ComboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox2.Location = New System.Drawing.Point(64, 0)
        Me.ComboBox2.Name = "ComboBox2"
        Me.ComboBox2.Size = New System.Drawing.Size(340, 21)
        Me.ComboBox2.TabIndex = 21
        '
        'ComboBox1
        '
        Me.ComboBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ComboBox1.Items.AddRange(New Object() {"One time only", "Everyday", "Every week", "Every month"})
        Me.ComboBox1.Location = New System.Drawing.Point(64, 88)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(153, 21)
        Me.ComboBox1.TabIndex = 32
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.DateTimePicker1.CustomFormat = "yyyy/MM/dd HH:mm:ss"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(220, 88)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.ShowUpDown = True
        Me.DateTimePicker1.Size = New System.Drawing.Size(183, 20)
        Me.DateTimePicker1.TabIndex = 33
        '
        'FrmCompose
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(420, 270)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.StatusBar1)
        Me.KeyPreview = True
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmCompose"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Message"
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub WriteMsglog(ByVal msg As String, Optional ByVal level As MsgBoxStyle = MsgBoxStyle.Information)
        StatusBar1.Text = level.ToString & ":" & msg
    End Sub

    Private Function LoadAccountBox() As Boolean
        Dim mxDoc As New XmlDocument
        Dim actFile As String = mainStorage & "accounts.xml"
        If File.Exists(actFile) Then
            mxDoc.Load(actFile)
            Dim mxNodeList As XmlNodeList = mxDoc.SelectNodes("/Accounts/Account")
            Dim itmx As ListViewItem
            ComboBox2.Items.Clear()
            Dim mxNode As XmlNode
            Dim actName As String
            Dim mailAdr As String
            For Each mxNode In mxNodeList
                If Not (mxNode.SelectSingleNode("Name") Is Nothing) Then
                    actName = mxNode.SelectSingleNode("Name").InnerText
                    mailAdr = mxNode.SelectSingleNode("ReplyAdr").InnerText
                    If mailAdr = "" Then
                        mailAdr = mxNode.SelectSingleNode("EmailAdr").InnerText
                    End If
                    ComboBox2.Items.Add(actName & "<" & mailAdr & ">")
                End If
            Next mxNode
            Return True
        Else
            MessageBox.Show("Account configuration missing, cannot continue.", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
        End If
    End Function

    Private Sub FrmCompose_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If LoadAccountBox() Then
            If _mailType = "MailList" Then
                ComboBox3.SelectedIndex = 1
            Else
                ComboBox3.SelectedIndex = 0
            End If
            ComboBox1.SelectedIndex = 0
            If ComboBox2.Items.Count > 0 Then ComboBox2.SelectedIndex = 0
            ComboBox4.SelectedIndex = 0
            If oldMessageID = "" Then
                ComboBox1.SelectedIndex = 0
                If ComboBox3.SelectedIndex = 1 Then TextBox1.Text = "Select * From users Where EmailList=Yes"
            Else
                MessageBox.Show("Unhandled 494", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                Try
                    Dim f As New FileStream(mainStorage & "Outbox\" & oldMessageID & ".mail", FileMode.Open)
                    Dim s As New StreamReader(f)
                    Dim sl As String
                    While sl = s.ReadLine()
                        If sl.StartsWith("To: ") Then TextBox1.Text = s.ReadLine()
                        If sl.StartsWith("Subject: ") Then
                        End If
                        s.ReadLine()
                        s.ReadLine()

                        TextBox3.Text = s.ReadLine()
                        ComboBox1.Text = s.ReadLine()
                        DateTimePicker1.Value = CType(s.ReadLine, Date)
                        s.ReadLine()
                        RichTextBox1.Text = s.ReadLine()
                    End While
                    s.Close()
                    f.Close()
                Catch ex As Exception
                    MsgBox("Error load MailList." & ex.Message, MsgBoxStyle.Critical)
                End Try
            End If
            Dim txts() As String = Split(reg.GetSetting("Faena Mail\Compose", "TextHistory", ""), ";")
            Dim txt As String
            MenuItem13.Visible = False
            For Each txt In txts
                MenuItem13.Visible = True
                If txt <> "" Then
                    MenuItem5.MenuItems.Add(txt, _
                         New EventHandler(AddressOf MenuInsertText_Click))
                End If
            Next
        Else
            Me.Close()
        End If
        If ListView1.Items.Count = 0 Then
            'hide attach box
            '_origtop = RichTextBox1.Top
            _origHeight = Panel1.Height
            'RichTextBox1.Height = RichTextBox1.Height + Math.Abs(RichTextBox1.Top - BtnAttach.Top)
            'RichTextBox1.Top = BtnAttach.Top
            Panel1.Height = BtnAttach.Top
            BtnAttach.Visible = False
            ListView1.Visible = False
        End If
    End Sub

    Private Sub ValidateEmailAddress(ByVal txt As TextBox)
        ' Confirm there is text in the control.
        If txt.TextLength = 0 Then
            Throw New Exception("Email address is a required field")
        Else
            ' Confirm that there is a "." and an "@" in the e-mail address.
            If txt.Text.IndexOf(".") = -1 Or txt.Text.IndexOf("@") = -1 Then
                Throw New Exception("E-mail address must be valid e-mail " & _
                    "address format." & ControlChars.Cr & "For example " & _
                    "'someone@isp.com'")
            End If
        End If
    End Sub

    Private Function UpdateMessage() As Boolean
        If Not Directory.Exists(mainStorage & "Outbox\") Then Directory.CreateDirectory(mainStorage & "Outbox\")
        Dim f As New FileStream(mainStorage & "Outbox\" & oldMessageID & ".mail", FileMode.Create)
        Dim s As New StreamWriter(f)
        Dim actName As String = ComboBox2.Text.Substring(0, ComboBox2.Text.IndexOf("<"))
        'actName = actName.Substring(1, actName.Length - 2)
        s.WriteLine("Account: " & actName)
        s.WriteLine("Date: " & Now.UtcNow.ToLongDateString & " " & Now.UtcNow.ToLongTimeString)
        s.WriteLine("MailType: " & ComboBox3.Text)
        s.WriteLine("From: " & ComboBox2.Text)
        s.WriteLine("To: " & TextBox1.Text)
        s.WriteLine("Cc: " & TextBox2.Text)
        s.WriteLine("Bcc: " & TextBox4.Text)
        s.WriteLine("Subject: " & TextBox3.Text)
        s.WriteLine("ScheduleType: " & ComboBox1.Text)
        s.WriteLine("ScheduleAt: " & DateTimePicker1.Value.ToString)
        For Each itm As ListViewItem In ListView1.Items
            s.WriteLine("Attachment: " & itm.Tag.ToString)
        Next
        s.WriteLine("Status: N")
        s.WriteLine()
        s.WriteLine(RichTextBox1.Text)
        s.WriteLine()
        s.Close()
        f.Close()
        UpdateMessage = True
    End Function

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Dim ok As Boolean = False
        If TextBox1.Text Like ("*" & EmailPart(ComboBox2.Text) & "*") Then
            If MsgBox("The From address and To address are the same. Do you want to continue?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo Or MsgBoxStyle.DefaultButton2) = MsgBoxResult.Yes Then
                ok = True
            Else
                TextBox1.Focus()
            End If
        Else
            ok = True
        End If
        If ok Then
            If ComboBox3.SelectedIndex = 0 Then
                Try
                    ValidateEmailAddress(TextBox1)
                Catch ex As Exception
                    TextBox1.Select(0, TextBox1.Text.Length)
                    ErrorProvider1.SetError(TextBox1, ex.Message)
                    Exit Sub
                End Try
            End If
            If oldMessageID = "" Then
                Do
                    oldMessageID = Path.GetFileName(Path.GetTempFileName)
                Loop While File.Exists(mainStorage & "Outbox\" & oldMessageID & ".mail")
            End If
            If UpdateMessage() Then Me.Close()
        End If
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Me.Close()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        If TextBox3.Text = "" Then
            Me.Text = "New Message"
        Else
            Me.Text = TextBox3.Text
        End If
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New frmAbout
        frm.ShowDialog(Me)
        frm.Dispose()
    End Sub
    Private Sub InserTextFile(ByVal fileName As String)
        If File.Exists(fileName) Then
            Dim FS As New FileStream(fileName, FileMode.Open, FileAccess.Read)
            Dim RD As New StreamReader(FS)
            If RD.Peek > 0 Then
                Dim strtemp As String = RD.ReadToEnd
                Dim str As New StringBuilder
                str.Append(RichTextBox1.Text)
                str.Insert(RichTextBox1.SelectionStart, strtemp)
                RichTextBox1.Text = str.ToString
                str = Nothing
            End If
            RD.Close()
            FS.Close()
        End If
    End Sub

    Private Sub MenuInsertText_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim itemClicked As New MenuItem
        itemClicked = CType(sender, MenuItem)
        InserTextFile(itemClicked.Text)
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        OpenFileDialog1.DefaultExt = ".txt"
        OpenFileDialog1.Filter = "Plain Text File (*.txt)|*.txt|All files (*.*)|*.*"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim strFullName As String = OpenFileDialog1.FileName
            InserTextFile(strFullName)
            Dim itm As MenuItem
            Dim itmCnt As Integer
            Dim found As Boolean = False
            For Each itm In MenuItem5.MenuItems
                itmCnt = itmCnt + 1
                If itm.Text = strFullName Then
                    found = True
                    Exit For
                End If
            Next
            If Not found Then
                Dim txts() As String = Split(reg.GetSetting("Faena Mail\Compose", "TextHistory", ""), ";")
                Dim txt, txtOut As String
                Dim txtCol As New Collection
                For Each txt In txts
                    If txt <> "" Then txtCol.Add(txt)
                Next
                If itmCnt > 20 Then
                    MenuItem5.MenuItems.RemoveAt(2)
                    txtCol.Remove(0)
                End If
                MenuItem5.MenuItems.Add(strFullName, _
                     New EventHandler(AddressOf MenuInsertText_Click))
                txtCol.Add(strFullName)
                For Each txt In txtCol
                    txtOut = txtOut & txt & ";"
                Next
                reg.SaveSetting("Faena Mail\Compose", "TextHistory", txtOut)
            End If
        End If
    End Sub

    Private Sub FrmCompose_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress
        If (Control.ModifierKeys And Keys.Shift) = Keys.Shift And (Control.ModifierKeys And Keys.Enter) = Keys.Enter Then
            e.Handled = True
            MenuItem6.PerformClick()
        End If
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            For Each fn As String In OpenFileDialog1.FileNames
                Dim fi As FileInfo = New FileInfo(fn)
                Dim itm As ListViewItem = ListView1.Items.Add(Path.GetFileName(fn) & "(" & (fi.Length \ 1024) & "KB)")
                itm.Tag = fn
            Next
            BtnAttach.Visible = True
            ListView1.Visible = True
            'show attach box
            'RichTextBox1.Height = RichTextBox1.Height - Math.Abs(RichTextBox1.Top - _origtop)
            'RichTextBox1.Top = _origtop
            Panel1.Height = _origHeight
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MenuItem9.PerformClick()
    End Sub

    Private Sub MenuItem3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        For Each itm As ListViewItem In ListView1.SelectedItems
            itm.Remove()
        Next
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        For Each itm As ListViewItem In ListView1.SelectedItems
            Shell(itm.Tag.ToString)
        Next
    End Sub

    Private Sub ListView1_MouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Dim lvitem As ListViewItem = CType(ListView1.GetItemAt(e.X, e.Y), ListViewItem)
        If (Not lvitem Is Nothing) Then
            ' Verify that the tag property is not "null".
            If Not lvitem.Tag Is Nothing Then
                ' Change the ToolTip only if the pointer moved to a new node.
                If lvitem.Tag.ToString() <> Me.ToolTip1.GetToolTip(Me.ListView1) Then
                    Me.ToolTip1.SetToolTip(Me.ListView1, lvitem.Tag.ToString())
                End If
            Else
                Me.ToolTip1.SetToolTip(Me.ListView1, "")
            End If
        Else  ' Pointer is not over a node so clear the ToolTip. Then 
            Me.ToolTip1.SetToolTip(Me.ListView1, "")
        End If
    End Sub

    Private Sub BtnAttach_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnAttach.Click
        MenuItem9.PerformClick()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub ListView1_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles ListView1.KeyPress
    End Sub

    Private Sub ListView1_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles ListView1.KeyUp
        If e.KeyCode = Keys.Delete Then
            MenuItem3.PerformClick()
            e.Handled = True
        End If

    End Sub
End Class
