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
Imports Microsoft.Win32
Imports System.IO
Imports System.ServiceProcess
Public Class FormOptions
    Inherits System.Windows.Forms.Form
    Private myService As ServiceController
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()
        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

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
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPage1 As System.Windows.Forms.TabPage
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents TabPage3 As System.Windows.Forms.TabPage
    Friend WithEvents TabPage4 As System.Windows.Forms.TabPage
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnContinue As System.Windows.Forms.Button
    Friend WithEvents btnPause As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents BtnShowODBC As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(FormOptions))
        Me.Button1 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPage3 = New System.Windows.Forms.TabPage
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.TabPage1 = New System.Windows.Forms.TabPage
        Me.BtnShowODBC = New System.Windows.Forms.Button
        Me.ComboBox1 = New System.Windows.Forms.ComboBox
        Me.Button5 = New System.Windows.Forms.Button
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.TabPage4 = New System.Windows.Forms.TabPage
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.btnStop = New System.Windows.Forms.Button
        Me.btnContinue = New System.Windows.Forms.Button
        Me.btnPause = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog
        Me.TabControl1.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabPage1.SuspendLayout()
        Me.TabPage4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button1.Location = New System.Drawing.Point(243, 299)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(81, 25)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&OK"
        '
        'Button2
        '
        Me.Button2.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button2.Location = New System.Drawing.Point(351, 299)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(80, 25)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "&Cancel"
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.TabPage3)
        Me.TabControl1.Controls.Add(Me.TabPage1)
        Me.TabControl1.Controls.Add(Me.TabPage4)
        Me.TabControl1.Location = New System.Drawing.Point(5, 6)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(464, 282)
        Me.TabControl1.TabIndex = 0
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.CheckBox6)
        Me.TabPage3.Controls.Add(Me.Label8)
        Me.TabPage3.Controls.Add(Me.NumericUpDown2)
        Me.TabPage3.Controls.Add(Me.CheckBox3)
        Me.TabPage3.Controls.Add(Me.CheckBox4)
        Me.TabPage3.Controls.Add(Me.CheckBox2)
        Me.TabPage3.Controls.Add(Me.Label5)
        Me.TabPage3.Controls.Add(Me.NumericUpDown1)
        Me.TabPage3.Controls.Add(Me.CheckBox1)
        Me.TabPage3.Location = New System.Drawing.Point(4, 22)
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.Size = New System.Drawing.Size(456, 256)
        Me.TabPage3.TabIndex = 2
        Me.TabPage3.Text = "Preferences"
        '
        'CheckBox6
        '
        Me.CheckBox6.Checked = True
        Me.CheckBox6.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox6.Location = New System.Drawing.Point(32, 40)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(216, 16)
        Me.CheckBox6.TabIndex = 0
        Me.CheckBox6.Text = "Check for new version at startup"
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label8.Location = New System.Drawing.Point(356, 171)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(54, 16)
        Me.Label8.TabIndex = 8
        Me.Label8.Text = "second(s)"
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Location = New System.Drawing.Point(276, 169)
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(65, 20)
        Me.NumericUpDown2.TabIndex = 7
        '
        'CheckBox3
        '
        Me.CheckBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox3.Location = New System.Drawing.Point(32, 166)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(232, 26)
        Me.CheckBox3.TabIndex = 6
        Me.CheckBox3.Text = "Mark message read after displaying for "
        '
        'CheckBox4
        '
        Me.CheckBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox4.Location = New System.Drawing.Point(32, 66)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(300, 20)
        Me.CheckBox4.TabIndex = 1
        Me.CheckBox4.Text = "Send and receive message at startup"
        '
        'CheckBox2
        '
        Me.CheckBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox2.Location = New System.Drawing.Point(32, 132)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(266, 24)
        Me.CheckBox2.TabIndex = 5
        Me.CheckBox2.Text = "Notify when new message arrives"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label5.Location = New System.Drawing.Point(356, 101)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(51, 16)
        Me.Label5.TabIndex = 4
        Me.Label5.Text = "minute(s)"
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Location = New System.Drawing.Point(276, 99)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(65, 20)
        Me.NumericUpDown1.TabIndex = 3
        '
        'CheckBox1
        '
        Me.CheckBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox1.Location = New System.Drawing.Point(32, 96)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(199, 26)
        Me.CheckBox1.TabIndex = 2
        Me.CheckBox1.Text = "Check for new messages every"
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.BtnShowODBC)
        Me.TabPage1.Controls.Add(Me.ComboBox1)
        Me.TabPage1.Controls.Add(Me.Button5)
        Me.TabPage1.Controls.Add(Me.TextBox3)
        Me.TabPage1.Controls.Add(Me.Label4)
        Me.TabPage1.Controls.Add(Me.Button4)
        Me.TabPage1.Controls.Add(Me.TextBox2)
        Me.TabPage1.Controls.Add(Me.Label3)
        Me.TabPage1.Controls.Add(Me.Label1)
        Me.TabPage1.Location = New System.Drawing.Point(4, 22)
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.Size = New System.Drawing.Size(456, 256)
        Me.TabPage1.TabIndex = 0
        Me.TabPage1.Text = "System"
        '
        'BtnShowODBC
        '
        Me.BtnShowODBC.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.BtnShowODBC.Location = New System.Drawing.Point(332, 56)
        Me.BtnShowODBC.Name = "BtnShowODBC"
        Me.BtnShowODBC.Size = New System.Drawing.Size(80, 23)
        Me.BtnShowODBC.TabIndex = 8
        Me.BtnShowODBC.Text = "Show ODBC"
        '
        'ComboBox1
        '
        Me.ComboBox1.Location = New System.Drawing.Point(16, 56)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(312, 21)
        Me.ComboBox1.TabIndex = 1
        Me.ComboBox1.Text = "ComboBox1"
        '
        'Button5
        '
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button5.Location = New System.Drawing.Point(332, 220)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(81, 25)
        Me.Button5.TabIndex = 7
        Me.Button5.Text = "Select"
        '
        'TextBox3
        '
        Me.TextBox3.Location = New System.Drawing.Point(16, 220)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(305, 20)
        Me.TextBox3.TabIndex = 6
        Me.TextBox3.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.Location = New System.Drawing.Point(20, 188)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(165, 16)
        Me.Label4.TabIndex = 5
        Me.Label4.Text = "Location of your message store:"
        '
        'Button4
        '
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button4.Location = New System.Drawing.Point(332, 140)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(81, 25)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Select"
        '
        'TextBox2
        '
        Me.TextBox2.Location = New System.Drawing.Point(16, 140)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(305, 20)
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.Text = ""
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.Location = New System.Drawing.Point(16, 116)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(75, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Select log file:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Location = New System.Drawing.Point(16, 26)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(139, 16)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "ODBC Data Source Name:"
        '
        'TabPage4
        '
        Me.TabPage4.Controls.Add(Me.Label6)
        Me.TabPage4.Controls.Add(Me.GroupBox3)
        Me.TabPage4.Location = New System.Drawing.Point(4, 22)
        Me.TabPage4.Name = "TabPage4"
        Me.TabPage4.Size = New System.Drawing.Size(456, 256)
        Me.TabPage4.TabIndex = 3
        Me.TabPage4.Text = "Service"
        '
        'Label6
        '
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label6.Location = New System.Drawing.Point(24, 34)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(408, 39)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "FaenaMailService Install"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.Label7)
        Me.GroupBox3.Controls.Add(Me.btnStop)
        Me.GroupBox3.Controls.Add(Me.btnContinue)
        Me.GroupBox3.Controls.Add(Me.btnPause)
        Me.GroupBox3.Controls.Add(Me.btnStart)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(22, 100)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(411, 120)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Status"
        '
        'Label7
        '
        Me.Label7.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label7.Location = New System.Drawing.Point(46, 27)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(331, 37)
        Me.Label7.TabIndex = 0
        Me.Label7.Text = "FaenaMailService Status"
        '
        'btnStop
        '
        Me.btnStop.AccessibleDescription = "Button used to Stop the Windows Service."
        Me.btnStop.AccessibleName = "Stop Button"
        Me.btnStop.Enabled = False
        Me.btnStop.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnStop.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnStop.Location = New System.Drawing.Point(310, 77)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(87, 27)
        Me.btnStop.TabIndex = 4
        Me.btnStop.Text = "Sto&p"
        '
        'btnContinue
        '
        Me.btnContinue.AccessibleDescription = "Button used to Continue the Windows Service."
        Me.btnContinue.AccessibleName = "Continue Button"
        Me.btnContinue.Enabled = False
        Me.btnContinue.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnContinue.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnContinue.Location = New System.Drawing.Point(209, 77)
        Me.btnContinue.Name = "btnContinue"
        Me.btnContinue.Size = New System.Drawing.Size(88, 27)
        Me.btnContinue.TabIndex = 3
        Me.btnContinue.Text = "&Continue"
        '
        'btnPause
        '
        Me.btnPause.AccessibleDescription = "Button used to Pause the Windows Service."
        Me.btnPause.AccessibleName = "Pause Button"
        Me.btnPause.Enabled = False
        Me.btnPause.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnPause.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnPause.Location = New System.Drawing.Point(109, 77)
        Me.btnPause.Name = "btnPause"
        Me.btnPause.Size = New System.Drawing.Size(87, 27)
        Me.btnPause.TabIndex = 2
        Me.btnPause.Text = "&Pause"
        '
        'btnStart
        '
        Me.btnStart.AccessibleDescription = "Button used to Start the Windows Service."
        Me.btnStart.AccessibleName = "Start Button"
        Me.btnStart.Enabled = False
        Me.btnStart.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.btnStart.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.btnStart.Location = New System.Drawing.Point(8, 77)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(88, 27)
        Me.btnStart.TabIndex = 1
        Me.btnStart.Text = "&Start"
        '
        'SaveFileDialog1
        '
        Me.SaveFileDialog1.Filter = "Log files|*.log|All files|*.*"
        Me.SaveFileDialog1.OverwritePrompt = False
        Me.SaveFileDialog1.RestoreDirectory = True
        Me.SaveFileDialog1.Title = "Select log file"
        '
        'Timer1
        '
        Me.Timer1.Interval = 1000
        '
        'FolderBrowserDialog1
        '
        Me.FolderBrowserDialog1.Description = "Select Main Storage Folder"
        '
        'FormOptions
        '
        Me.AcceptButton = Me.Button1
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.Button2
        Me.ClientSize = New System.Drawing.Size(481, 333)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Options"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPage3.ResumeLayout(False)
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



    Private Sub FormOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CheckBox6.Checked = Not CType(reg.GetSetting("Faena Mail\", "DonotCheckUpdate", 0), Boolean)
        Dim tools As New FaenaMailLibrary.tools
        Dim a() As String = tools.LoadDSNList()
        ComboBox1.Items.Clear()
        If Not a Is Nothing Then ComboBox1.Items.AddRange(a)
        ComboBox1.Text = reg.GetSetting("Faena Mail\", "ODBCDSN", "")
        TextBox2.Text = reg.GetSetting("Faena Mail\", "LogFile", AppDomain.CurrentDomain.BaseDirectory & "FaenaMailDebug.log")
        TextBox3.Text = reg.GetSetting("Faena Mail\", "MsgStore", AppDomain.CurrentDomain.BaseDirectory)
        NumericUpDown1.Value = reg.GetSetting("Faena Mail\", "CheckEmailTimer", 30)
        CheckBox1.Checked = reg.GetSetting("Faena Mail\", "CheckEmail", True)
        CheckBox2.Checked = reg.GetSetting("Faena Mail\", "ShowNotify", True)
        CheckBox4.Checked = CType(reg.GetSetting("Faena Mail\", "CheckAtStartup", 1), Boolean)
        NumericUpDown2.Value = reg.GetSetting("Faena Mail\", "MarkReadDelay", 3)
        CheckBox3.Checked = CType(reg.GetSetting("Faena Mail\", "MarkRead", 1), Boolean)
        If CheckServiceInstallation() Then
            Label6.Text = "FaenaMail Service is installed."
        Else
            Label6.Text = "FaenaMail Service is not installed. This requires Windows NT,2000,XP."
        End If
        MyCheckBox()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            If Not TextBox3.Text.EndsWith("\") Then
                TextBox3.Text = TextBox3.Text & "\"
            End If
            If Not Directory.Exists(TextBox3.Text) Then
                Directory.CreateDirectory(TextBox3.Text)
            End If
            reg.SaveSetting("Faena Mail\", "ODBCDSN", ComboBox1.Text)

            reg.SaveSetting("Faena Mail\", "LogFile", TextBox2.Text)
            reg.SaveSetting("Faena Mail\", "MsgStore", TextBox3.Text)
            mainStorage = TextBox3.Text
            reg.SaveSetting("Faena Mail\", "CheckEmailTimer", CInt(NumericUpDown1.Value))
            reg.SaveSetting("Faena Mail\", "CheckEmail", CInt(CheckBox1.Checked))
            reg.SaveSetting("Faena Mail\", "ShowNotify", CInt(CheckBox2.Checked))
            reg.SaveSetting("Faena Mail\", "CheckAtStartup", CInt(CheckBox4.Checked))
            reg.SaveSetting("Faena Mail\", "MarkReadDelay", CInt(NumericUpDown2.Value))
            reg.SaveSetting("Faena Mail\", "MarkRead", CInt(CheckBox3.Checked))
            reg.SaveSetting("Faena Mail\", "DonotCheckUpdate", CLng(Not CheckBox6.Checked))

            If CheckServiceInstallation() Then
                Select Case myService.Status
                    Case ServiceControllerStatus.Stopped
                    Case ServiceControllerStatus.Running
                        myService.Stop()
                        Dim tmTimeout As Date = Now
                        Do Until (myService.Status = ServiceControllerStatus.Stopped)
                            Application.DoEvents()
                            If DateDiff(DateInterval.Second, tmTimeout, Now) > 30 Then
                                MsgBox("Timeout when stopping FaenaMailServie. If this error repeats, please contact Faena technical support.", MsgBoxStyle.Critical)
                                Exit Sub
                            End If
                        Loop
                        myService.Start()
                    Case ServiceControllerStatus.Paused
                    Case Else
                End Select
            End If
        Catch ex As Exception
            MsgBox("Error when saving setting. " & ex.Message, MsgBoxStyle.Critical)
        End Try
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SaveFileDialog1.RestoreDirectory = True
        SaveFileDialog1.DefaultExt = ".log"
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            TextBox2.Text = SaveFileDialog1.FileName
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyCheckBox()
    End Sub

    Private Sub MyCheckBox()
        If CheckBox1.Checked Then
            NumericUpDown1.Enabled = True
        Else
            NumericUpDown1.Enabled = False
        End If
        If CheckBox3.Checked Then
            NumericUpDown2.Enabled = True
        Else
            NumericUpDown2.Enabled = False
        End If
    End Sub

    Private Function CheckServiceInstallation() As Boolean
        ' Verify to see if the service is installed.
        Timer1.Enabled = False
        Dim installedServices() As ServiceController
        Dim tmpService As ServiceController
        Dim i As Integer = 0
        Try
            installedServices = ServiceController.GetServices()

            For Each tmpService In installedServices
                If tmpService.DisplayName = "FaenaMailService" Then
                    myService = tmpService
                    CheckServiceInstallation = True
                    Exit For
                End If
            Next tmpService
            Timer1.Enabled = True
        Catch
            'Win9x service not supported.
        End Try
    End Function

    Private Sub UpdateServiceStatus()

        ' Shut off the timer so it doesn't raise events while were
        '   in this code
        Timer1.Enabled = False

        If Not (myService Is Nothing) Then

            ' Recreate myServer, otherwise the status for myServer never
            '   changes, despite changes in the status of the 
            '   windows service
            CheckServiceInstallation()

            Select Case myService.Status
                Case ServiceControllerStatus.Stopped
                    Me.btnContinue.Enabled = False
                    Me.btnPause.Enabled = False
                    Me.btnStart.Enabled = True
                    Me.btnStop.Enabled = False
                Case ServiceControllerStatus.Running
                    Me.btnContinue.Enabled = False
                    Me.btnPause.Enabled = True
                    Me.btnStart.Enabled = False
                    Me.btnStop.Enabled = True
                Case ServiceControllerStatus.Paused
                    Me.btnContinue.Enabled = True
                    Me.btnPause.Enabled = False
                    Me.btnStart.Enabled = False
                    Me.btnStop.Enabled = True
                Case Else
                    ' This can occur when an action is pending. In this case
                    '   don't allow the user to push any buttons.
                    Me.btnContinue.Enabled = False
                    Me.btnPause.Enabled = False
                    Me.btnStart.Enabled = False
                    Me.btnStop.Enabled = False
            End Select

            ' Output the status to the Status Bar
            Label7.Text = "Service Status: " + myService.Status.ToString()
        Else
            ' The service isn't running so dim everything but refresh.
            Label7.Text = "Service not found"
            Me.btnContinue.Enabled = False
            Me.btnPause.Enabled = False
            Me.btnStart.Enabled = False
            Me.btnStop.Enabled = False
        End If
        Timer1.Enabled = True
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        UpdateServiceStatus()
    End Sub

    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click
        btnStart.Enabled = False
        myService.Start()
    End Sub

    Private Sub btnPause_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPause.Click
        btnPause.Enabled = False
        myService.Pause()
    End Sub

    Private Sub btnContinue_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnContinue.Click
        btnContinue.Enabled = False
        myService.[Continue]()
    End Sub

    Private Sub btnStop_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStop.Click
        btnStop.Enabled = False
        myService.Stop()
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MyCheckBox()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        FolderBrowserDialog1.SelectedPath = TextBox3.Text
        If FolderBrowserDialog1.ShowDialog = DialogResult.OK Then
            TextBox3.Text = FolderBrowserDialog1.SelectedPath
        End If
    End Sub



    Private Sub BtnShowODBC_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnShowODBC.Click
        Shell("Rundll32.exe shell32.dll,Control_RunDLL Odbccp32.cpl")
        Dim tools As New FaenaMailLibrary.tools
        Dim a() As String = tools.LoadDSNList()
        ComboBox1.Items.Clear()
        If Not a Is Nothing Then ComboBox1.Items.AddRange(a)
        ComboBox1.Text = reg.GetSetting("Faena Mail\", "ODBCDSN", "")

    End Sub
End Class
