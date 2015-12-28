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
Imports System.Drawing.Drawing2D
Imports System.Xml
Public Class FormScriptWindow
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New(Optional ByVal scriptFileName As String = "")
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        mScriptFileName = scriptFileName
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
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
    Friend WithEvents MenuItem1 As System.Windows.Forms.MenuItem
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
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents MenuItem2 As System.Windows.Forms.MenuItem
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.MainMenu1 = New System.Windows.Forms.MainMenu()
        Me.MenuItem1 = New System.Windows.Forms.MenuItem()
        Me.MenuItem3 = New System.Windows.Forms.MenuItem()
        Me.MenuItem4 = New System.Windows.Forms.MenuItem()
        Me.MenuItem5 = New System.Windows.Forms.MenuItem()
        Me.MenuItem6 = New System.Windows.Forms.MenuItem()
        Me.MenuItem7 = New System.Windows.Forms.MenuItem()
        Me.MenuItem8 = New System.Windows.Forms.MenuItem()
        Me.MenuItem9 = New System.Windows.Forms.MenuItem()
        Me.MenuItem10 = New System.Windows.Forms.MenuItem()
        Me.MenuItem11 = New System.Windows.Forms.MenuItem()
        Me.MenuItem12 = New System.Windows.Forms.MenuItem()
        Me.MenuItem13 = New System.Windows.Forms.MenuItem()
        Me.MenuItem15 = New System.Windows.Forms.MenuItem()
        Me.MenuItem16 = New System.Windows.Forms.MenuItem()
        Me.MenuItem17 = New System.Windows.Forms.MenuItem()
        Me.MenuItem2 = New System.Windows.Forms.MenuItem()
        Me.MenuItem14 = New System.Windows.Forms.MenuItem()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.SuspendLayout()
        '
        'Panel1
        '
        Me.Panel1.AccessibleName = ""
        Me.Panel1.AutoScroll = True
        Me.Panel1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(292, 273)
        Me.Panel1.TabIndex = 0
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem8, Me.MenuItem12})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem3, Me.MenuItem4, Me.MenuItem5, Me.MenuItem6, Me.MenuItem7})
        Me.MenuItem1.MergeType = System.Windows.Forms.MenuMerge.MergeItems
        Me.MenuItem1.Text = "&File"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 0
        Me.MenuItem3.Text = "&Save"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 1
        Me.MenuItem4.Text = "Save &As..."
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 2
        Me.MenuItem5.Text = "&Close"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "Previe&w"
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 4
        Me.MenuItem7.Text = "&Print"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 1
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem9, Me.MenuItem10, Me.MenuItem11})
        Me.MenuItem8.Text = "&Edit"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 0
        Me.MenuItem9.Text = "&Copy"
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 1
        Me.MenuItem10.Text = "Cu&t"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 2
        Me.MenuItem11.Text = "&Paste"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 2
        Me.MenuItem12.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem13, Me.MenuItem14})
        Me.MenuItem12.Text = "&Scripts"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 0
        Me.MenuItem13.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem15, Me.MenuItem16, Me.MenuItem17, Me.MenuItem2})
        Me.MenuItem13.Text = "&Insert"
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 0
        Me.MenuItem15.Text = "IfScript"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 1
        Me.MenuItem16.Text = "SQLScript"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 2
        Me.MenuItem17.Text = "Send email"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 3
        Me.MenuItem2.Text = "Mail List"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 1
        Me.MenuItem14.Text = "&Delete"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.FileName = "doc1"
        '
        'FormScriptWindow
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(292, 273)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Panel1})
        Me.Menu = Me.MainMenu1
        Me.Name = "FormScriptWindow"
        Me.Text = "FormScriptWindow"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private modified As Boolean
    Private mScriptFileName As String
    Private mctlScriptBox As ScriptBoxArray

    Private Sub Panel1_Paint(ByVal sender As System.Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles Panel1.Paint
        Dim grp As Graphics = Panel1.CreateGraphics()
        Dim myArrow As New AdjustableArrowCap(4, 4, True)
        Dim customArrow As CustomLineCap = myArrow
        grp.Clear(Color.WhiteSmoke)
        Dim thePen As New Pen(Color.Black)
        thePen.CustomEndCap = myArrow
        thePen.LineJoin = Drawing.Drawing2D.LineJoin.Bevel
        Dim i, ii As Integer
        For i = 0 To mctlScriptBox.Count - 1
            If mctlScriptBox(i).mSettingPath <> "" Then
                For ii = 0 To mctlScriptBox.Count - 1
                    If mctlScriptBox(ii).Label1.Text = mctlScriptBox(i).mSettingPath Then
                        grp.DrawLine(thePen, CInt(mctlScriptBox(i).Left + mctlScriptBox(i).Width / 2), CInt(mctlScriptBox(i).Top + mctlScriptBox(i).Height), _
                            CInt(mctlScriptBox(ii).Left + mctlScriptBox(ii).Width / 2), CInt(mctlScriptBox(ii).Top))
                    End If
                Next
            End If
        Next
    End Sub

    Private Function LoadNodeText(ByVal xnode As XmlNode, ByVal strname As String) As String
        Dim oneNode As XmlNode
        oneNode = xnode.SelectSingleNode(strname)
        If Not oneNode Is Nothing Then
            Return oneNode.InnerText
        Else
            Return ""
        End If
    End Function
    Private Sub FormScriptWindow_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim frm As New FormWait()
        frm.Show()
        frm.Refresh()
        Dim ok As Boolean = True
        mctlScriptBox = New ScriptBoxArray(Panel1)
        If mScriptFileName <> "" Then
            Dim xDoc As New XmlDocument()
            Try
                xDoc.Load(mScriptFileName)
                Dim xnodeList As XmlNodeList
                Dim xnode As XmlNode
                xnode = xDoc.SelectSingleNode("/scriptfile")
                Me.Tag = xnode.Attributes("startupbox").InnerText
                xnodeList = xDoc.SelectNodes("/scriptfile/scriptbox")
                For Each xnode In xnodeList
                    Dim aScript As ScriptBox = mctlScriptBox.AddNewScriptBox()
                    aScript.Button1.Text = xnode.Attributes("type").InnerText
                    aScript.Label1.Text = LoadNodeText(xnode, "name")
                    aScript.Name = aScript.Label1.Text
                    aScript.Left = XmlConvert.ToInt32(LoadNodeText(xnode, "left"))
                    If aScript.Left < 0 Then aScript.Left = 0
                    aScript.Top = XmlConvert.ToInt32(LoadNodeText(xnode, "top"))
                    If aScript.Top < 0 Then aScript.Top = 0
                    aScript.TextBox1.Text = LoadNodeText(xnode, "textbox")
                    aScript.mSettingIf = XmlConvert.ToInt32(LoadNodeText(xnode, "settingif"))
                    aScript.mSettingIfDes = LoadNodeText(xnode, "settingifdes")
                    aScript.mSettingAction = XmlConvert.ToInt32(LoadNodeText(xnode, "settingaction"))
                    aScript.mSettingActionDes = LoadNodeText(xnode, "settingactiondes")
                    aScript.mSettingBcc = LoadNodeText(xnode, "settingbcc")
                    aScript.mSettingMsg = LoadNodeText(xnode, "settingmsg")
                    aScript.mSettingAttachment = LoadNodeText(xnode, "settingattachment")
                    aScript.mSettingPath = LoadNodeText(xnode, "settingpath")
                    aScript.modified = False
                Next
            Catch ex As Exception
                MsgBox("Unable to load file:" & mScriptFileName & vbNewLine & ex.Message, MsgBoxStyle.Exclamation)
            End Try
        End If
        frm.Hide()
        frm.Dispose()
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        mctlScriptBox.Remove()
        modified = True
    End Sub

    Private Sub MenuItem15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem15.Click
        Dim aScript As ScriptBox = mctlScriptBox.AddNewScriptBox()
        aScript.Button1.Text = "IfScript"
        aScript.TextBox1.Text = "IfScript act as an email filter. It also lets you setup script path."
        modified = True
    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click
        Dim aScript As ScriptBox = mctlScriptBox.AddNewScriptBox()
        aScript.Button1.Text = "SQL"
        aScript.TextBox1.Text = "Enter SQL query here. Such as SELECT/UPDATE/INSERT/DELETE."
        modified = True
    End Sub

    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click
        Dim aScript As ScriptBox = mctlScriptBox.AddNewScriptBox()
        aScript.Button1.Text = "email"
        aScript.TextBox1.Text = "Send email message and attachments."
        modified = True
    End Sub

    Private Function MakeBackup(ByVal orgFile As String, ByVal bakFile As String) As Boolean
        MakeBackup = True
        Try
            If IO.File.Exists(bakFile) Then Kill(bakFile)
            Rename(orgFile, bakFile)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Backup")
            MakeBackup = False
        End Try
    End Function

    Private Sub SaveScript()
        If MakeBackup(mScriptFileName, mScriptFileName & ".bak") Then
            Dim i As Integer
            Try
                Dim xWriter As XmlTextWriter = New XmlTextWriter(mScriptFileName, System.Text.Encoding.UTF8)
                xWriter.WriteStartDocument(True)
                xWriter.WriteComment("This is a script file composed by Faena Mail. http://www.faena.com")
                xWriter.WriteStartElement("scriptfile")
                xWriter.WriteAttributeString("version", "1")
                If CheckStartUpBox() Then
                    xWriter.WriteAttributeString("startupbox", Me.Tag.ToString)
                Else
                    xWriter.WriteAttributeString("startupbox", "")
                End If
                For i = 0 To mctlScriptBox.Count - 1
                    xWriter.WriteStartElement("scriptbox")
                    xWriter.WriteAttributeString("type", mctlScriptBox(i).Button1.Text)
                    xWriter.WriteElementString("name", mctlScriptBox(i).Label1.Text)
                    xWriter.WriteElementString("left", XmlConvert.ToString(mctlScriptBox(i).Left))
                    xWriter.WriteElementString("top", XmlConvert.ToString(mctlScriptBox(i).Top))
                    xWriter.WriteElementString("textbox", mctlScriptBox(i).TextBox1.Text)
                    xWriter.WriteElementString("settingif", XmlConvert.ToString(mctlScriptBox(i).mSettingIf))
                    xWriter.WriteElementString("settingifdes", mctlScriptBox(i).mSettingIfDes)
                    xWriter.WriteElementString("settingaction", XmlConvert.ToString(mctlScriptBox(i).mSettingAction))
                    xWriter.WriteElementString("settingactiondes", mctlScriptBox(i).mSettingActionDes)
                    xWriter.WriteElementString("settingbcc", mctlScriptBox(i).mSettingBcc)
                    xWriter.WriteElementString("settingmsg", mctlScriptBox(i).mSettingMsg)
                    xWriter.WriteElementString("settingattachment", mctlScriptBox(i).mSettingAttachment)
                    xWriter.WriteElementString("settingpath", mctlScriptBox(i).mSettingPath)
                    xWriter.WriteEndElement()
                Next
                xWriter.WriteEndElement()
                xWriter.WriteEndDocument()
                xWriter.Close()
                modified = False
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Saving File")
                MakeBackup(mScriptFileName & ".bak", mScriptFileName)
            End Try
        Else
                MsgBox("Unable to make backup file." & vbNewLine & "To prevent data lost, use menu File->Save As another file name.", MsgBoxStyle.Critical)
        End If
    End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        If mScriptFileName = "" Then
            MenuItem4.PerformClick()
        Else
            SaveScript()
        End If
    End Sub

    Private Sub MenuItem4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem4.Click
        SaveFileDialog1.FileName = Me.Text
        SaveFileDialog1.AddExtension = True
        SaveFileDialog1.DefaultExt = ".Script"
        SaveFileDialog1.Filter = "Faena Mail Script File (*.Script)|*.Script|All files (*.*)|*.*"

        SaveFileDialog1.OverwritePrompt = True
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            mScriptFileName = SaveFileDialog1.FileName
            Me.Text = mScriptFileName
            SaveScript()
        End If
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Me.Close()
    End Sub

    Private Function CheckStartUpBox() As Boolean
        Dim i As Integer
        Dim ok As Boolean = False
        If Not Me.Tag Is Nothing Then
            For i = 0 To mctlScriptBox.Count - 1

                If mctlScriptBox(i).Name.ToString = Me.Tag.ToString Then
                    ok = True
                    Exit For
                End If
            Next
        End If
        If Not ok Then
            Dim frm As New FormSet1stScript()
            For i = 0 To mctlScriptBox.Count - 1
                frm.ComboBox1.Items.Add(mctlScriptBox(i).Name)
            Next
            If frm.ShowDialog = DialogResult.OK Then
                Me.Tag = frm.ComboBox1.Text
                modified = True
                ok = True
            End If
        End If
        Return ok
    End Function

    Private Sub FormScriptWindow_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        Dim i As Integer
        For i = 0 To mctlScriptBox.Count - 1
            If mctlScriptBox(i).modified = True Then
                modified = True
                Exit For
            End If
        Next
        If modified Then
            Select Case MsgBox("The script is modified. Do you want to save it?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNoCancel, Me.Text.Substring(Me.Text.LastIndexOf("\") + 1))
                Case MsgBoxResult.Yes
                    MenuItem3.PerformClick()
                Case MsgBoxResult.No
                Case Else
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Sub MenuItem19_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim aScript As ScriptBox = mctlScriptBox.AddNewScriptBox()
        aScript.Button1.Text = "GetField"
        aScript.TextBox1.Text = "Get field from DB or file"
        modified = True
    End Sub

    Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
        Dim aScript As ScriptBox = mctlScriptBox.AddNewScriptBox()
        aScript.Button1.Text = "MailList"
        aScript.TextBox1.Text = "Add to/Delete from mailing list"
        modified = True
    End Sub


End Class
