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
Imports System.IO
Public Class FrmViewEmail
    Inherits System.Windows.Forms.Form
    Private mailFilePath As String
    Private actName As String

#Region " Windows Form Designer generated code "

    Public Sub New(ByVal strPath As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mailFilePath = strPath
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        RichTextBox1.Text = ""
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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem3 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem4 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem5 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem6 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem7 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem8 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem9 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem10 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem11 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem12 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem26 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem28 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FrmViewEmail))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.MenuItem31 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem26 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem28 = New System.Windows.Forms.MenuItem
        Me.MenuItem27 = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.Label1 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem6, Me.MenuItem8, Me.MenuItem4})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem21, Me.MenuItem23, Me.MenuItem22, Me.MenuItem2, Me.MenuItem25, Me.MenuItem26, Me.MenuItem24, Me.MenuItem28, Me.MenuItem27, Me.MenuItem29, Me.MenuItem3})
        Me.MenuItem1.Text = "&File"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 0
        Me.MenuItem21.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem30, Me.MenuItem31})
        Me.MenuItem21.Text = "&New"
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 0
        Me.MenuItem30.Text = "Single Mail Message"
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 1
        Me.MenuItem31.Text = "Mailing List Message"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 1
        Me.MenuItem23.Text = "-"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 2
        Me.MenuItem22.Text = "&Save As ..."
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "-"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 4
        Me.MenuItem25.Text = "Move to Folder..."
        '
        'MenuItem26
        '
        Me.MenuItem26.Index = 5
        Me.MenuItem26.Text = "Copy to Folder..."
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 6
        Me.MenuItem24.Text = "&Delete Message"
        '
        'MenuItem28
        '
        Me.MenuItem28.Index = 7
        Me.MenuItem28.Text = "-"
        '
        'MenuItem27
        '
        Me.MenuItem27.Index = 8
        Me.MenuItem27.Text = "Print..."
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 9
        Me.MenuItem29.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 10
        Me.MenuItem3.Text = "&Close"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 1
        Me.MenuItem6.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem19, Me.MenuItem20, Me.MenuItem18, Me.MenuItem7})
        Me.MenuItem6.Text = "&View"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 0
        Me.MenuItem19.Text = "&Previous Message"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 1
        Me.MenuItem20.Text = "&Next Message"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 2
        Me.MenuItem18.Text = "-"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 3
        Me.MenuItem7.Text = "&Source"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 2
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem11, Me.MenuItem12, Me.MenuItem13, Me.MenuItem14, Me.MenuItem15, Me.MenuItem16, Me.MenuItem17})
        Me.MenuItem8.Text = "&Message"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "&New Single Message"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "New Mailing List Message"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 2
        Me.MenuItem11.Text = "-"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 3
        Me.MenuItem12.Text = "&Reply to Sender"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 4
        Me.MenuItem13.Text = "Reply to &All"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 5
        Me.MenuItem14.Text = "&Forward"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 6
        Me.MenuItem15.Text = "Forwar&d As Attachment"
        Me.MenuItem15.Visible = False
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 7
        Me.MenuItem16.Text = "-"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 8
        Me.MenuItem17.Text = "&Flag Message"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 3
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem5})
        Me.MenuItem4.Text = "&Help"
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 0
        Me.MenuItem5.Text = "&About"
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 254)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Size = New System.Drawing.Size(493, 16)
        Me.StatusBar1.TabIndex = 1
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(4, 4)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 16)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "From:"
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox1.Location = New System.Drawing.Point(58, 6)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.ReadOnly = True
        Me.TextBox1.Size = New System.Drawing.Size(236, 13)
        Me.TextBox1.TabIndex = 3
        Me.TextBox1.Text = ""
        '
        'Label2
        '
        Me.Label2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(307, 7)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(32, 16)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Date:"
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox2.Location = New System.Drawing.Point(346, 7)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.ReadOnly = True
        Me.TextBox2.Size = New System.Drawing.Size(130, 13)
        Me.TextBox2.TabIndex = 5
        Me.TextBox2.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(4, 29)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(21, 16)
        Me.Label3.TabIndex = 6
        Me.Label3.Text = "To:"
        '
        'TextBox3
        '
        Me.TextBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox3.Location = New System.Drawing.Point(58, 29)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ReadOnly = True
        Me.TextBox3.Size = New System.Drawing.Size(418, 13)
        Me.TextBox3.TabIndex = 7
        Me.TextBox3.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(4, 72)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(47, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "Subject:"
        '
        'TextBox4
        '
        Me.TextBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox4.Location = New System.Drawing.Point(56, 72)
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ReadOnly = True
        Me.TextBox4.Size = New System.Drawing.Size(422, 13)
        Me.TextBox4.TabIndex = 9
        Me.TextBox4.Text = ""
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RichTextBox1.Location = New System.Drawing.Point(0, 112)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(491, 140)
        Me.RichTextBox1.TabIndex = 11
        Me.RichTextBox1.Text = ""
        '
        'Timer1
        '
        Me.Timer1.Interval = 3000
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(4, 50)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(22, 16)
        Me.Label5.TabIndex = 12
        Me.Label5.Text = "Cc:"
        '
        'TextBox5
        '
        Me.TextBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox5.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox5.Location = New System.Drawing.Point(57, 52)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.ReadOnly = True
        Me.TextBox5.Size = New System.Drawing.Size(236, 13)
        Me.TextBox5.TabIndex = 13
        Me.TextBox5.Text = ""
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.AutoSize = True
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(309, 52)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(27, 16)
        Me.Label6.TabIndex = 14
        Me.Label6.Text = "Bcc:"
        '
        'TextBox6
        '
        Me.TextBox6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.TextBox6.Location = New System.Drawing.Point(350, 53)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.ReadOnly = True
        Me.TextBox6.Size = New System.Drawing.Size(130, 13)
        Me.TextBox6.TabIndex = 15
        Me.TextBox6.Text = ""
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label7.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(4, 92)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(67, 16)
        Me.Label7.TabIndex = 8
        Me.Label7.Text = "Attachment:"
        '
        'ListView1
        '
        Me.ListView1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.ListView1.Location = New System.Drawing.Point(76, 92)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(404, 16)
        Me.ListView1.TabIndex = 16
        '
        'FrmViewEmail
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(493, 270)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Controls.Add(Me.Label7)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmViewEmail"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "FrmViewEmail"
        Me.ResumeLayout(False)

    End Sub

#End Region

    'Private Sub DispHtml(ByVal strIn As String)
    '    Dim fileName As String
    '    Do
    '        fileName = Path.GetDirectoryName(mailFilePath) & Path.GetFileNameWithoutExtension(Path.GetTempFileName) & ".htm"
    '    Loop While File.Exists(fileName)
    '    Dim FS As New FileStream(fileName, FileMode.CreateNew, FileAccess.Write)
    '    Dim SW As New StreamWriter(FS)
    '    SW.Write(strIn)
    '    SW.Close()
    '    FS.Close()

    '    AxWebBrowser1.Navigate2("file:///" & fileName.Replace("\", "/").Replace(" ", "%20"), )
    '    'File.Delete(fileName)
    'End Sub

    Private Sub FrmViewEmail_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        StatusBar1.Text = "Loading email..."
        Try
            Dim FS As FileStream
            Dim SR As StreamReader
            Dim strLine As String
            Dim strLineTemp As String
            FS = New FileStream(mailFilePath, FileMode.Open, FileAccess.Read)
            SR = New StreamReader(FS)

            While SR.Peek > 0
                strLine = SR.ReadLine
                If strLine = "" Then
                    Exit While
                ElseIf strLine.StartsWith("Account:") Then
                    If strLine.Length >= 9 Then
                        actName = strLine.Substring(9).ToString
                    Else
                        actName = ""
                    End If
                ElseIf strLine.StartsWith("From:") Then
                    If strLine.Length >= 6 Then
                        TextBox1.Text = strLine.Substring(6).ToString
                    Else
                        TextBox1.Text = ""
                    End If
                ElseIf strLine.StartsWith("Date:") Then
                    If strLine.Length >= 6 Then
                        TextBox2.Text = strLine.Substring(6).ToString
                    Else
                        TextBox2.Text = ""
                    End If
                ElseIf strLine.StartsWith("To:") Then
                    If strLine.Length >= 4 Then
                        TextBox3.Text = strLine.Substring(4).ToString
                    Else
                        TextBox3.Text = ""
                    End If
                ElseIf strLine.StartsWith("Cc:") Then
                    If strLine.Length >= 4 Then
                        TextBox5.Text = strLine.Substring(4).ToString
                    Else
                        TextBox5.Text = ""
                    End If
                ElseIf strLine.StartsWith("Bcc:") Then
                    If strLine.Length >= 5 Then
                        TextBox6.Text = strLine.Substring(5).ToString
                    Else
                        TextBox6.Text = ""
                    End If
                ElseIf strLine.StartsWith("Subject:") Then
                    If strLine.Length > 9 Then
                        TextBox4.Text = strLine.Substring(9).ToString
                        Me.Text = TextBox4.Text
                    Else
                        TextBox4.Text = ""
                        Me.Text = "(No Subject)"
                    End If
                ElseIf strLine.StartsWith("Attachment:") Then
                    If strLine.Length > Len("Attachment:") Then
                        Dim fn As String = strLine.Substring(Len("Attachment:"))
                        Dim itm As ListViewItem
                        If File.Exists(fn) Then
                            Dim fi As FileInfo = New FileInfo(fn)
                            itm = ListView1.Items.Add(Path.GetFileNameWithoutExtension(fn) & "(" & (fi.Length \ 1024) & "KB)")
                            itm.Tag = fn
                        Else
                            itm = ListView1.Items.Add(Path.GetFileNameWithoutExtension(fn) & "(Not found)")
                            itm.Tag = "Not exists: " & fn
                        End If
                    End If
                    End If
            End While

            'strLine = SR.ReadLine
            'If strLine = "This is a multi-part message in MIME format." Then
            '    While SR.Peek > 0
            '        strLine = SR.ReadLine
            '        If strLine = "Content-Transfer-Encoding: quoted-printable" Then Exit While
            '    End While
            '    strLine = ""
            '    While SR.Peek > 0
            '        strLineTemp = SR.ReadLine
            '        If strLineTemp.StartsWith("------=") Then Exit While
            '        strLine = strLine & strLineTemp & vbNewLine
            '    End While
            'Else
            '    strLine = strLine & vbNewLine & SR.ReadToEnd
            'End If
            strLine = SR.ReadToEnd
            RichTextBox1.Text = textfrommime(strLine)
            'DispHtml(strLine)
            SR.Close()
            FS.Close()
        Catch ex As Exception
            MsgBox("Error open email. " & ex.Message, MsgBoxStyle.Exclamation)
        End Try
        StatusBar1.Text = "Size " & FileLen(mailFilePath).ToString("#,#") & " byte(s)"
        Timer1.Interval = reg.GetSetting("Faena Mail\", "MarkReadDelay", 3) * 1000
        Timer1.Enabled = CType(reg.GetSetting("Faena Mail\", "MarkRead", 1), Boolean)
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Me.Close()
    End Sub

    Private Sub MenuItem30_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem30.Click
        Dim frm As New FrmCompose("", "")
        frm.Show()
    End Sub

    Private Sub MenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem31.Click
        Dim frm As New FrmCompose("", "MailList")
        frm.Show()
    End Sub

    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Dim frm As New FrmCompose("", "")
        frm.Show()
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Dim frm As New FrmCompose("", "MailList")
        frm.Show()
    End Sub

    Private Sub MenuItem12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem12.Click
        Dim frm As New FrmCompose("", "")
        Dim i As Integer
        For i = 0 To frm.ComboBox2.Items.Count - 1
            If frm.ComboBox2.Items(i).ToString Like "*" & actName & "*" Then
                frm.ComboBox2.SelectedIndex = i
            End If
        Next
        frm.TextBox3.Text = "Re: " & Me.TextBox4.Text
        frm.TextBox1.Text = Me.TextBox1.Text
        frm.RichTextBox1.Text = QuoteReplyText(RichTextBox1.Text)
        frm.Show()
        frm.RichTextBox1.Focus()
        Me.Close()
    End Sub

    Private Function QuoteReplyText(ByVal body As String) As String
        Dim sBody As String
        QuoteReplyText = vbNewLine
        QuoteReplyText = QuoteReplyText & "----- Original Message -----" & vbNewLine
        QuoteReplyText = QuoteReplyText & "From: " & TextBox1.Text & vbNewLine
        QuoteReplyText = QuoteReplyText & "To: " & TextBox3.Text & vbNewLine
        QuoteReplyText = QuoteReplyText & "Sent: " & TextBox2.Text & vbNewLine
        QuoteReplyText = QuoteReplyText & "Subject: " & TextBox4.Text & vbNewLine
        QuoteReplyText = QuoteReplyText & vbNewLine
        QuoteReplyText = QuoteReplyText & body
        'For Each sBody In Split(body, vbLf)
        '    QuoteReplyText = QuoteReplyText & "> " & sBody & vbLf
        'Next
    End Function

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Dim frm As New frmAbout()
        frm.ShowDialog(Me)
        frm.Dispose()
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        MarkMailStatus(mailFilePath, "R")
    End Sub

    Private Sub MenuItem7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem7.Click
        Dim frm As New FrmView(mailFilePath)
        frm.Show()
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Dim frm As New FrmCompose("", "")
        Dim i As Integer
        For i = 0 To frm.ComboBox2.Items.Count - 1
            If frm.ComboBox2.Items(i).ToString Like "*" & actName & "*" Then
                frm.ComboBox2.SelectedIndex = i
            End If
        Next
        frm.TextBox3.Text = "Re: " & Me.TextBox4.Text
        frm.TextBox1.Text = Me.TextBox1.Text & ";" & Me.TextBox3.Text
        frm.TextBox2.Text = Me.TextBox5.Text
        frm.RichTextBox1.Text = QuoteReplyText(RichTextBox1.Text)
        frm.Show()
        frm.RichTextBox1.Focus()
        Me.Close()
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        Dim frm As New FrmCompose("", "")
        Dim i As Integer
        For i = 0 To frm.ComboBox2.Items.Count - 1
            If frm.ComboBox2.Items(i).ToString Like "*" & actName & "*" Then
                frm.ComboBox2.SelectedIndex = i
            End If
        Next
        frm.TextBox3.Text = "Fw: " & Me.TextBox4.Text
        frm.RichTextBox1.Text = QuoteReplyText(RichTextBox1.Text)
        frm.Show()
        frm.TextBox1.Focus()
        Me.Close()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged

    End Sub

    Private Sub ListView1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles ListView1.MouseMove
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
End Class
