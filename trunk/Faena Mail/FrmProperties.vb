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
Imports System.IO
Imports System.Xml
Public Class FormProperties
    Inherits System.Windows.Forms.Form
    Private oldAccountName As String 'Current AccountName
    'Private mActName As String
#Region " Windows Form Designer generated code "

    Public Sub New(ByVal actName As String)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        oldAccountName = actName
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
    Friend WithEvents TabControl1 As System.Windows.Forms.TabControl
    Friend WithEvents TabPageGeneral As System.Windows.Forms.TabPage
    Friend WithEvents TabPageServers As System.Windows.Forms.TabPage
    Friend WithEvents TabPageAdvanced As System.Windows.Forms.TabPage
    Friend WithEvents ButtonOK As System.Windows.Forms.Button
    Friend WithEvents ButtonCancel As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents NumericUpDown1 As System.Windows.Forms.NumericUpDown
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox5 As System.Windows.Forms.GroupBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox6 As System.Windows.Forms.GroupBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents TextBox9 As System.Windows.Forms.TextBox
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents TextBox10 As System.Windows.Forms.TextBox
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents TextBox11 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents TabScripts As System.Windows.Forms.TabPage
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ErrorProvider1 As System.Windows.Forms.ErrorProvider
    Friend WithEvents GroupBox7 As System.Windows.Forms.GroupBox
    Friend WithEvents NumericUpDown2 As System.Windows.Forms.NumericUpDown
    Friend WithEvents NumericUpDown3 As System.Windows.Forms.NumericUpDown
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox6 As System.Windows.Forms.CheckBox
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    Friend WithEvents GroupBox8 As System.Windows.Forms.GroupBox
    Friend WithEvents TextBox12 As System.Windows.Forms.TextBox
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents TextBox13 As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Validator1 As ValidatorControl.Validator
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.TabControl1 = New System.Windows.Forms.TabControl
        Me.TabPageGeneral = New System.Windows.Forms.TabPage
        Me.CheckBox4 = New System.Windows.Forms.CheckBox
        Me.GroupBox6 = New System.Windows.Forms.GroupBox
        Me.TextBox11 = New System.Windows.Forms.TextBox
        Me.Label12 = New System.Windows.Forms.Label
        Me.TextBox10 = New System.Windows.Forms.TextBox
        Me.Label11 = New System.Windows.Forms.Label
        Me.TextBox9 = New System.Windows.Forms.TextBox
        Me.Label10 = New System.Windows.Forms.Label
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.Label9 = New System.Windows.Forms.Label
        Me.GroupBox5 = New System.Windows.Forms.GroupBox
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.Label8 = New System.Windows.Forms.Label
        Me.TabPageServers = New System.Windows.Forms.TabPage
        Me.CheckBox3 = New System.Windows.Forms.CheckBox
        Me.GroupBox4 = New System.Windows.Forms.GroupBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.Label7 = New System.Windows.Forms.Label
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.Label6 = New System.Windows.Forms.Label
        Me.GroupBox3 = New System.Windows.Forms.GroupBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.Label5 = New System.Windows.Forms.Label
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.Label4 = New System.Windows.Forms.Label
        Me.GroupBox8 = New System.Windows.Forms.GroupBox
        Me.RadioButton1 = New System.Windows.Forms.RadioButton
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label15 = New System.Windows.Forms.Label
        Me.TextBox13 = New System.Windows.Forms.TextBox
        Me.Label14 = New System.Windows.Forms.Label
        Me.TextBox12 = New System.Windows.Forms.TextBox
        Me.RadioButton2 = New System.Windows.Forms.RadioButton
        Me.TabPageAdvanced = New System.Windows.Forms.TabPage
        Me.GroupBox7 = New System.Windows.Forms.GroupBox
        Me.CheckBox5 = New System.Windows.Forms.CheckBox
        Me.NumericUpDown3 = New System.Windows.Forms.NumericUpDown
        Me.NumericUpDown2 = New System.Windows.Forms.NumericUpDown
        Me.CheckBox6 = New System.Windows.Forms.CheckBox
        Me.GroupBox2 = New System.Windows.Forms.GroupBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.GroupBox1 = New System.Windows.Forms.GroupBox
        Me.Label1 = New System.Windows.Forms.Label
        Me.NumericUpDown1 = New System.Windows.Forms.NumericUpDown
        Me.CheckBox2 = New System.Windows.Forms.CheckBox
        Me.CheckBox1 = New System.Windows.Forms.CheckBox
        Me.TabScripts = New System.Windows.Forms.TabPage
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.Button6 = New System.Windows.Forms.Button
        Me.Label13 = New System.Windows.Forms.Label
        Me.Button4 = New System.Windows.Forms.Button
        Me.Button3 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.ButtonOK = New System.Windows.Forms.Button
        Me.ButtonCancel = New System.Windows.Forms.Button
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog
        Me.ErrorProvider1 = New System.Windows.Forms.ErrorProvider
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Validator1 = New ValidatorControl.Validator(Me.components)
        Me.TabControl1.SuspendLayout()
        Me.TabPageGeneral.SuspendLayout()
        Me.GroupBox6.SuspendLayout()
        Me.GroupBox5.SuspendLayout()
        Me.TabPageServers.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox8.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.TabPageAdvanced.SuspendLayout()
        Me.GroupBox7.SuspendLayout()
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TabScripts.SuspendLayout()
        Me.SuspendLayout()
        '
        'TabControl1
        '
        Me.TabControl1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TabControl1.Controls.Add(Me.TabPageGeneral)
        Me.TabControl1.Controls.Add(Me.TabPageServers)
        Me.TabControl1.Controls.Add(Me.TabPageAdvanced)
        Me.TabControl1.Controls.Add(Me.TabScripts)
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(441, 338)
        Me.TabControl1.TabIndex = 0
        '
        'TabPageGeneral
        '
        Me.TabPageGeneral.Controls.Add(Me.CheckBox4)
        Me.TabPageGeneral.Controls.Add(Me.GroupBox6)
        Me.TabPageGeneral.Controls.Add(Me.GroupBox5)
        Me.TabPageGeneral.Location = New System.Drawing.Point(4, 22)
        Me.TabPageGeneral.Name = "TabPageGeneral"
        Me.TabPageGeneral.Size = New System.Drawing.Size(433, 312)
        Me.TabPageGeneral.TabIndex = 0
        Me.TabPageGeneral.Text = "General"
        '
        'CheckBox4
        '
        Me.CheckBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox4.Location = New System.Drawing.Point(11, 237)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(343, 28)
        Me.CheckBox4.TabIndex = 2
        Me.CheckBox4.Text = "Include this account when receiving mail or synchronizing"
        '
        'GroupBox6
        '
        Me.GroupBox6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox6.Controls.Add(Me.TextBox11)
        Me.GroupBox6.Controls.Add(Me.Label12)
        Me.GroupBox6.Controls.Add(Me.TextBox10)
        Me.GroupBox6.Controls.Add(Me.Label11)
        Me.GroupBox6.Controls.Add(Me.TextBox9)
        Me.GroupBox6.Controls.Add(Me.Label10)
        Me.GroupBox6.Controls.Add(Me.TextBox8)
        Me.GroupBox6.Controls.Add(Me.Label9)
        Me.GroupBox6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox6.Location = New System.Drawing.Point(8, 96)
        Me.GroupBox6.Name = "GroupBox6"
        Me.GroupBox6.Size = New System.Drawing.Size(416, 128)
        Me.GroupBox6.TabIndex = 1
        Me.GroupBox6.TabStop = False
        Me.GroupBox6.Text = "User Information"
        '
        'TextBox11
        '
        Me.TextBox11.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox11, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox11, False)
        Me.TextBox11.Location = New System.Drawing.Point(100, 97)
        Me.Validator1.SetMaxValue(Me.TextBox11, "")
        Me.Validator1.SetMinValue(Me.TextBox11, "")
        Me.TextBox11.Name = "TextBox11"
        Me.Validator1.SetRegularExpression(Me.TextBox11, "")
        Me.Validator1.SetRequired(Me.TextBox11, False)
        Me.TextBox11.Size = New System.Drawing.Size(203, 20)
        Me.TextBox11.TabIndex = 7
        Me.TextBox11.Text = ""
        '
        'Label12
        '
        Me.Label12.AutoSize = True
        Me.Label12.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label12.Location = New System.Drawing.Point(12, 101)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(80, 16)
        Me.Label12.TabIndex = 6
        Me.Label12.Text = "Reply address:"
        '
        'TextBox10
        '
        Me.TextBox10.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox10, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox10, True)
        Me.TextBox10.Location = New System.Drawing.Point(100, 70)
        Me.Validator1.SetMaxValue(Me.TextBox10, "")
        Me.Validator1.SetMinValue(Me.TextBox10, "")
        Me.TextBox10.Name = "TextBox10"
        Me.Validator1.SetRegularExpression(Me.TextBox10, "")
        Me.Validator1.SetRequired(Me.TextBox10, True)
        Me.TextBox10.Size = New System.Drawing.Size(203, 20)
        Me.TextBox10.TabIndex = 5
        Me.TextBox10.Text = ""
        '
        'Label11
        '
        Me.Label11.AutoSize = True
        Me.Label11.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label11.Location = New System.Drawing.Point(12, 73)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(83, 16)
        Me.Label11.TabIndex = 4
        Me.Label11.Text = "E-mail address:"
        '
        'TextBox9
        '
        Me.TextBox9.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox9, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox9, False)
        Me.TextBox9.Location = New System.Drawing.Point(100, 44)
        Me.Validator1.SetMaxValue(Me.TextBox9, "")
        Me.Validator1.SetMinValue(Me.TextBox9, "")
        Me.TextBox9.Name = "TextBox9"
        Me.Validator1.SetRegularExpression(Me.TextBox9, "")
        Me.Validator1.SetRequired(Me.TextBox9, False)
        Me.TextBox9.Size = New System.Drawing.Size(203, 20)
        Me.TextBox9.TabIndex = 3
        Me.TextBox9.Text = ""
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label10.Location = New System.Drawing.Point(12, 45)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(72, 16)
        Me.Label10.TabIndex = 2
        Me.Label10.Text = "Organization:"
        '
        'TextBox8
        '
        Me.TextBox8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox8, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox8, False)
        Me.TextBox8.Location = New System.Drawing.Point(100, 19)
        Me.Validator1.SetMaxValue(Me.TextBox8, "")
        Me.Validator1.SetMinValue(Me.TextBox8, "")
        Me.TextBox8.Name = "TextBox8"
        Me.Validator1.SetRegularExpression(Me.TextBox8, "")
        Me.Validator1.SetRequired(Me.TextBox8, False)
        Me.TextBox8.Size = New System.Drawing.Size(203, 20)
        Me.TextBox8.TabIndex = 1
        Me.TextBox8.Text = ""
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label9.Location = New System.Drawing.Point(12, 20)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(38, 16)
        Me.Label9.TabIndex = 0
        Me.Label9.Text = "Name:"
        '
        'GroupBox5
        '
        Me.GroupBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox5.Controls.Add(Me.TextBox7)
        Me.GroupBox5.Controls.Add(Me.Label8)
        Me.GroupBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox5.Location = New System.Drawing.Point(8, 5)
        Me.GroupBox5.Name = "GroupBox5"
        Me.GroupBox5.Size = New System.Drawing.Size(416, 88)
        Me.GroupBox5.TabIndex = 0
        Me.GroupBox5.TabStop = False
        Me.GroupBox5.Text = "Mail Account"
        '
        'TextBox7
        '
        Me.TextBox7.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox7, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox7, False)
        Me.TextBox7.Location = New System.Drawing.Point(22, 55)
        Me.Validator1.SetMaxValue(Me.TextBox7, "")
        Me.Validator1.SetMinValue(Me.TextBox7, "")
        Me.TextBox7.Name = "TextBox7"
        Me.Validator1.SetRegularExpression(Me.TextBox7, "")
        Me.Validator1.SetRequired(Me.TextBox7, False)
        Me.TextBox7.Size = New System.Drawing.Size(282, 20)
        Me.TextBox7.TabIndex = 1
        Me.TextBox7.Text = ""
        '
        'Label8
        '
        Me.Label8.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label8.Location = New System.Drawing.Point(22, 19)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(346, 33)
        Me.Label8.TabIndex = 0
        Me.Label8.Text = "Type the name by which you would refer to this account.  For example: ""Work"" or """ & _
        "Home Address""."
        '
        'TabPageServers
        '
        Me.TabPageServers.Controls.Add(Me.CheckBox3)
        Me.TabPageServers.Controls.Add(Me.GroupBox4)
        Me.TabPageServers.Controls.Add(Me.GroupBox3)
        Me.TabPageServers.Controls.Add(Me.GroupBox8)
        Me.TabPageServers.Location = New System.Drawing.Point(4, 22)
        Me.TabPageServers.Name = "TabPageServers"
        Me.TabPageServers.Size = New System.Drawing.Size(433, 312)
        Me.TabPageServers.TabIndex = 1
        Me.TabPageServers.Text = "Servers"
        '
        'CheckBox3
        '
        Me.CheckBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox3.Location = New System.Drawing.Point(16, 168)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(284, 24)
        Me.CheckBox3.TabIndex = 4
        Me.CheckBox3.Text = "My outgoing server (SMTP) requires authentication"
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.TextBox6)
        Me.GroupBox4.Controls.Add(Me.Label7)
        Me.GroupBox4.Controls.Add(Me.TextBox5)
        Me.GroupBox4.Controls.Add(Me.Label6)
        Me.GroupBox4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox4.Location = New System.Drawing.Point(7, 96)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(402, 72)
        Me.GroupBox4.TabIndex = 1
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Incoming Mail Server"
        '
        'TextBox6
        '
        Me.TextBox6.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox6, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox6, False)
        Me.TextBox6.Location = New System.Drawing.Point(132, 47)
        Me.Validator1.SetMaxValue(Me.TextBox6, "")
        Me.Validator1.SetMinValue(Me.TextBox6, "")
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.Validator1.SetRegularExpression(Me.TextBox6, "")
        Me.Validator1.SetRequired(Me.TextBox6, False)
        Me.TextBox6.Size = New System.Drawing.Size(221, 20)
        Me.TextBox6.TabIndex = 3
        Me.TextBox6.Text = ""
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label7.Location = New System.Drawing.Point(12, 49)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(57, 16)
        Me.Label7.TabIndex = 2
        Me.Label7.Text = "Password:"
        '
        'TextBox5
        '
        Me.TextBox5.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox5, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox5, False)
        Me.TextBox5.Location = New System.Drawing.Point(132, 20)
        Me.Validator1.SetMaxValue(Me.TextBox5, "")
        Me.Validator1.SetMinValue(Me.TextBox5, "")
        Me.TextBox5.Name = "TextBox5"
        Me.Validator1.SetRegularExpression(Me.TextBox5, "")
        Me.Validator1.SetRequired(Me.TextBox5, False)
        Me.TextBox5.Size = New System.Drawing.Size(221, 20)
        Me.TextBox5.TabIndex = 1
        Me.TextBox5.Text = ""
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label6.Location = New System.Drawing.Point(12, 22)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 16)
        Me.Label6.TabIndex = 0
        Me.Label6.Text = "Account name:"
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.TextBox4)
        Me.GroupBox3.Controls.Add(Me.Label5)
        Me.GroupBox3.Controls.Add(Me.TextBox3)
        Me.GroupBox3.Controls.Add(Me.Label4)
        Me.GroupBox3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox3.Location = New System.Drawing.Point(7, 4)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(402, 84)
        Me.GroupBox3.TabIndex = 0
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Server Information"
        '
        'TextBox4
        '
        Me.TextBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox4, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox4, False)
        Me.TextBox4.Location = New System.Drawing.Point(132, 58)
        Me.Validator1.SetMaxValue(Me.TextBox4, "")
        Me.Validator1.SetMinValue(Me.TextBox4, "")
        Me.TextBox4.Name = "TextBox4"
        Me.Validator1.SetRegularExpression(Me.TextBox4, "")
        Me.Validator1.SetRequired(Me.TextBox4, False)
        Me.TextBox4.Size = New System.Drawing.Size(220, 20)
        Me.TextBox4.TabIndex = 3
        Me.TextBox4.Text = ""
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label5.Location = New System.Drawing.Point(12, 60)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(119, 16)
        Me.Label5.TabIndex = 2
        Me.Label5.Text = "Outgoing mail (SMTP):"
        '
        'TextBox3
        '
        Me.TextBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox3, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox3, False)
        Me.TextBox3.Location = New System.Drawing.Point(132, 26)
        Me.Validator1.SetMaxValue(Me.TextBox3, "")
        Me.Validator1.SetMinValue(Me.TextBox3, "")
        Me.TextBox3.Name = "TextBox3"
        Me.Validator1.SetRegularExpression(Me.TextBox3, "")
        Me.Validator1.SetRequired(Me.TextBox3, False)
        Me.TextBox3.Size = New System.Drawing.Size(220, 20)
        Me.TextBox3.TabIndex = 1
        Me.TextBox3.Text = ""
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label4.Location = New System.Drawing.Point(12, 28)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(118, 16)
        Me.Label4.TabIndex = 0
        Me.Label4.Text = "Incoming mail (POP3):"
        '
        'GroupBox8
        '
        Me.GroupBox8.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox8.Controls.Add(Me.RadioButton1)
        Me.GroupBox8.Controls.Add(Me.Panel1)
        Me.GroupBox8.Controls.Add(Me.RadioButton2)
        Me.GroupBox8.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox8.Location = New System.Drawing.Point(7, 172)
        Me.GroupBox8.Name = "GroupBox8"
        Me.GroupBox8.Size = New System.Drawing.Size(402, 136)
        Me.GroupBox8.TabIndex = 1
        Me.GroupBox8.TabStop = False
        '
        'RadioButton1
        '
        Me.RadioButton1.Checked = True
        Me.RadioButton1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.RadioButton1.Location = New System.Drawing.Point(28, 24)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(284, 24)
        Me.RadioButton1.TabIndex = 6
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Use same settings as my incoming mail server"
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label15)
        Me.Panel1.Controls.Add(Me.TextBox13)
        Me.Panel1.Controls.Add(Me.Label14)
        Me.Panel1.Controls.Add(Me.TextBox12)
        Me.Panel1.Enabled = False
        Me.Panel1.Location = New System.Drawing.Point(24, 72)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(356, 56)
        Me.Panel1.TabIndex = 5
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label15.Location = New System.Drawing.Point(16, 8)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(80, 16)
        Me.Label15.TabIndex = 0
        Me.Label15.Text = "Account name:"
        '
        'TextBox13
        '
        Me.TextBox13.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox13, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox13, False)
        Me.TextBox13.Location = New System.Drawing.Point(112, 4)
        Me.Validator1.SetMaxValue(Me.TextBox13, "")
        Me.Validator1.SetMinValue(Me.TextBox13, "")
        Me.TextBox13.Name = "TextBox13"
        Me.Validator1.SetRegularExpression(Me.TextBox13, "")
        Me.Validator1.SetRequired(Me.TextBox13, False)
        Me.TextBox13.Size = New System.Drawing.Size(212, 20)
        Me.TextBox13.TabIndex = 1
        Me.TextBox13.Text = ""
        '
        'Label14
        '
        Me.Label14.AutoSize = True
        Me.Label14.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label14.Location = New System.Drawing.Point(16, 32)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(57, 16)
        Me.Label14.TabIndex = 2
        Me.Label14.Text = "Password:"
        '
        'TextBox12
        '
        Me.TextBox12.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox12, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox12, False)
        Me.TextBox12.Location = New System.Drawing.Point(112, 32)
        Me.Validator1.SetMaxValue(Me.TextBox12, "")
        Me.Validator1.SetMinValue(Me.TextBox12, "")
        Me.TextBox12.Name = "TextBox12"
        Me.TextBox12.PasswordChar = Microsoft.VisualBasic.ChrW(42)
        Me.Validator1.SetRegularExpression(Me.TextBox12, "")
        Me.Validator1.SetRequired(Me.TextBox12, False)
        Me.TextBox12.Size = New System.Drawing.Size(212, 20)
        Me.TextBox12.TabIndex = 3
        Me.TextBox12.Text = ""
        '
        'RadioButton2
        '
        Me.RadioButton2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.RadioButton2.Location = New System.Drawing.Point(28, 48)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(284, 24)
        Me.RadioButton2.TabIndex = 6
        Me.RadioButton2.Text = "Log on using"
        '
        'TabPageAdvanced
        '
        Me.TabPageAdvanced.Controls.Add(Me.GroupBox7)
        Me.TabPageAdvanced.Controls.Add(Me.GroupBox2)
        Me.TabPageAdvanced.Controls.Add(Me.GroupBox1)
        Me.TabPageAdvanced.Location = New System.Drawing.Point(4, 22)
        Me.TabPageAdvanced.Name = "TabPageAdvanced"
        Me.TabPageAdvanced.Size = New System.Drawing.Size(433, 312)
        Me.TabPageAdvanced.TabIndex = 2
        Me.TabPageAdvanced.Text = "Advanced"
        '
        'GroupBox7
        '
        Me.GroupBox7.Controls.Add(Me.CheckBox5)
        Me.GroupBox7.Controls.Add(Me.NumericUpDown3)
        Me.GroupBox7.Controls.Add(Me.NumericUpDown2)
        Me.GroupBox7.Controls.Add(Me.CheckBox6)
        Me.GroupBox7.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox7.Location = New System.Drawing.Point(16, 196)
        Me.GroupBox7.Name = "GroupBox7"
        Me.GroupBox7.Size = New System.Drawing.Size(390, 88)
        Me.GroupBox7.TabIndex = 2
        Me.GroupBox7.TabStop = False
        Me.GroupBox7.Text = "Messages"
        '
        'CheckBox5
        '
        Me.CheckBox5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox5.Location = New System.Drawing.Point(12, 24)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(256, 24)
        Me.CheckBox5.TabIndex = 4
        Me.CheckBox5.Text = "Limit number of messages per connection to:"
        '
        'NumericUpDown3
        '
        Me.NumericUpDown3.Enabled = False
        Me.NumericUpDown3.Location = New System.Drawing.Point(267, 52)
        Me.NumericUpDown3.Maximum = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.NumericUpDown3.Name = "NumericUpDown3"
        Me.NumericUpDown3.Size = New System.Drawing.Size(68, 20)
        Me.NumericUpDown3.TabIndex = 3
        '
        'NumericUpDown2
        '
        Me.NumericUpDown2.Enabled = False
        Me.NumericUpDown2.Location = New System.Drawing.Point(267, 24)
        Me.NumericUpDown2.Maximum = New Decimal(New Integer() {999999, 0, 0, 0})
        Me.NumericUpDown2.Name = "NumericUpDown2"
        Me.NumericUpDown2.Size = New System.Drawing.Size(68, 20)
        Me.NumericUpDown2.TabIndex = 1
        '
        'CheckBox6
        '
        Me.CheckBox6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox6.Location = New System.Drawing.Point(12, 52)
        Me.CheckBox6.Name = "CheckBox6"
        Me.CheckBox6.Size = New System.Drawing.Size(244, 24)
        Me.CheckBox6.TabIndex = 4
        Me.CheckBox6.Text = "Limit number of recipients per message to:"
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.TextBox2)
        Me.GroupBox2.Controls.Add(Me.Label3)
        Me.GroupBox2.Controls.Add(Me.TextBox1)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox2.Location = New System.Drawing.Point(16, 4)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(397, 104)
        Me.GroupBox2.TabIndex = 1
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Server Port Numbers"
        '
        'TextBox2
        '
        Me.TextBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox2, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox2, False)
        Me.TextBox2.Location = New System.Drawing.Point(156, 30)
        Me.Validator1.SetMaxValue(Me.TextBox2, "")
        Me.Validator1.SetMinValue(Me.TextBox2, "")
        Me.TextBox2.Name = "TextBox2"
        Me.Validator1.SetRegularExpression(Me.TextBox2, "")
        Me.Validator1.SetRequired(Me.TextBox2, False)
        Me.TextBox2.Size = New System.Drawing.Size(98, 20)
        Me.TextBox2.TabIndex = 3
        Me.TextBox2.Text = "110"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label3.Location = New System.Drawing.Point(14, 33)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(118, 16)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Incoming mail (POP3):"
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Validator1.SetDataType(Me.TextBox1, ValidatorControl.DataTypeConstants.StringType)
        Me.Validator1.SetIfValidate(Me.TextBox1, False)
        Me.TextBox1.Location = New System.Drawing.Point(156, 65)
        Me.Validator1.SetMaxValue(Me.TextBox1, "")
        Me.Validator1.SetMinValue(Me.TextBox1, "")
        Me.TextBox1.Name = "TextBox1"
        Me.Validator1.SetRegularExpression(Me.TextBox1, "")
        Me.Validator1.SetRequired(Me.TextBox1, False)
        Me.TextBox1.Size = New System.Drawing.Size(98, 20)
        Me.TextBox1.TabIndex = 1
        Me.TextBox1.Text = "25"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label2.Location = New System.Drawing.Point(12, 68)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(119, 16)
        Me.Label2.TabIndex = 0
        Me.Label2.Text = "Outgoing mail (SMTP):"
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.NumericUpDown1)
        Me.GroupBox1.Controls.Add(Me.CheckBox2)
        Me.GroupBox1.Controls.Add(Me.CheckBox1)
        Me.GroupBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.GroupBox1.Location = New System.Drawing.Point(17, 108)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(397, 84)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Incoming Mail"
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label1.Location = New System.Drawing.Point(351, 56)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(36, 16)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "day(s)"
        '
        'NumericUpDown1
        '
        Me.NumericUpDown1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NumericUpDown1.Location = New System.Drawing.Point(232, 48)
        Me.NumericUpDown1.Name = "NumericUpDown1"
        Me.NumericUpDown1.Size = New System.Drawing.Size(110, 20)
        Me.NumericUpDown1.TabIndex = 2
        '
        'CheckBox2
        '
        Me.CheckBox2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox2.Location = New System.Drawing.Point(34, 48)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(163, 18)
        Me.CheckBox2.TabIndex = 1
        Me.CheckBox2.Text = "Remove from server after"
        '
        'CheckBox1
        '
        Me.CheckBox1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.CheckBox1.Location = New System.Drawing.Point(17, 24)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(247, 17)
        Me.CheckBox1.TabIndex = 0
        Me.CheckBox1.Text = "Leave a copy of messages on server"
        '
        'TabScripts
        '
        Me.TabScripts.Controls.Add(Me.ListView1)
        Me.TabScripts.Controls.Add(Me.Button6)
        Me.TabScripts.Controls.Add(Me.Label13)
        Me.TabScripts.Controls.Add(Me.Button4)
        Me.TabScripts.Controls.Add(Me.Button3)
        Me.TabScripts.Controls.Add(Me.Button2)
        Me.TabScripts.Controls.Add(Me.Button1)
        Me.TabScripts.Location = New System.Drawing.Point(4, 22)
        Me.TabScripts.Name = "TabScripts"
        Me.TabScripts.Size = New System.Drawing.Size(433, 312)
        Me.TabScripts.TabIndex = 3
        Me.TabScripts.Text = "Scripts"
        '
        'ListView1
        '
        Me.ListView1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.ListView1.FullRowSelect = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(10, 46)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(315, 236)
        Me.ListView1.TabIndex = 8
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Name"
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Path"
        Me.ColumnHeader2.Width = 176
        '
        'Button6
        '
        Me.Button6.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button6.Enabled = False
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button6.Location = New System.Drawing.Point(334, 93)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(84, 25)
        Me.Button6.TabIndex = 7
        Me.Button6.Text = "&Edit..."
        '
        'Label13
        '
        Me.Label13.AutoSize = True
        Me.Label13.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Label13.Location = New System.Drawing.Point(10, 18)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(121, 16)
        Me.Label13.TabIndex = 5
        Me.Label13.Text = "Auto-execute script list:"
        '
        'Button4
        '
        Me.Button4.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button4.Enabled = False
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button4.Location = New System.Drawing.Point(334, 237)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(84, 25)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "&Remove"
        '
        'Button3
        '
        Me.Button3.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button3.Enabled = False
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button3.Location = New System.Drawing.Point(334, 189)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(84, 25)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Move &Down"
        '
        'Button2
        '
        Me.Button2.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button2.Enabled = False
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button2.Location = New System.Drawing.Point(334, 141)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(84, 25)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Move &Up"
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button1.Location = New System.Drawing.Point(334, 45)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(84, 24)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "&Add..."
        '
        'ButtonOK
        '
        Me.ButtonOK.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonOK.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.ButtonOK.Location = New System.Drawing.Point(264, 347)
        Me.ButtonOK.Name = "ButtonOK"
        Me.ButtonOK.Size = New System.Drawing.Size(70, 26)
        Me.ButtonOK.TabIndex = 1
        Me.ButtonOK.Text = "&OK"
        '
        'ButtonCancel
        '
        Me.ButtonCancel.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ButtonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ButtonCancel.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.ButtonCancel.Location = New System.Drawing.Point(349, 347)
        Me.ButtonCancel.Name = "ButtonCancel"
        Me.ButtonCancel.Size = New System.Drawing.Size(70, 26)
        Me.ButtonCancel.TabIndex = 2
        Me.ButtonCancel.Text = "&Cancel"
        '
        'ErrorProvider1
        '
        Me.ErrorProvider1.ContainerControl = Me
        '
        'Validator1
        '
        Me.Validator1.DisplayControl = Nothing
        '
        'FormProperties
        '
        Me.AcceptButton = Me.ButtonOK
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.CancelButton = Me.ButtonCancel
        Me.ClientSize = New System.Drawing.Size(440, 378)
        Me.Controls.Add(Me.ButtonCancel)
        Me.Controls.Add(Me.ButtonOK)
        Me.Controls.Add(Me.TabControl1)
        Me.HelpButton = True
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormProperties"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Properties"
        Me.TabControl1.ResumeLayout(False)
        Me.TabPageGeneral.ResumeLayout(False)
        Me.GroupBox6.ResumeLayout(False)
        Me.GroupBox5.ResumeLayout(False)
        Me.TabPageServers.ResumeLayout(False)
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox8.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.TabPageAdvanced.ResumeLayout(False)
        Me.GroupBox7.ResumeLayout(False)
        CType(Me.NumericUpDown3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.NumericUpDown2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.NumericUpDown1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TabScripts.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub SeeCheckBox2()
        CheckBox2.Enabled = CheckBox1.Checked
        NumericUpDown1.Enabled = CheckBox1.Checked
        Label1.Enabled = CheckBox1.Checked
    End Sub

    Private Sub FormOption_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If oldAccountName <> "" Then
            Dim xDoc As New XmlDocument()
            Dim xNode As XmlNode
            xDoc.Load(mainStorage & "accounts.xml")
            xNode = xDoc.SelectSingleNode("descendant::Account[Name='" & oldAccountName & "']")
            If Not (xNode Is Nothing) Then
                TextBox7.Text = oldAccountName
                TextBox8.Text = xNode.SelectSingleNode("UserName").InnerText
                TextBox9.Text = xNode.SelectSingleNode("Organization").InnerText
                TextBox10.Text = xNode.SelectSingleNode("EmailAdr").InnerText
                TextBox11.Text = xNode.SelectSingleNode("ReplyAdr").InnerText
                CheckBox4.Checked = CType(xNode.SelectSingleNode("Included").InnerText, Boolean)
                TextBox2.Text = xNode.SelectSingleNode("POPPort").InnerText
                TextBox1.Text = xNode.SelectSingleNode("SMTPPort").InnerText
                CheckBox1.Checked = CType(xNode.SelectSingleNode("LeaveOnSvr").InnerText, Boolean)
                CheckBox2.Checked = CType(xNode.SelectSingleNode("RemoveFromSvr").InnerText, Boolean)
                NumericUpDown1.Value = CType(xNode.SelectSingleNode("HowManyDays").InnerText, Integer)
                TextBox3.Text = xNode.SelectSingleNode("POPSvr").InnerText
                TextBox4.Text = xNode.SelectSingleNode("SMTPSvr").InnerText
                If xNode.SelectSingleNode("Auth") Is Nothing Then
                    CheckBox3.Checked = False
                Else
                    Select Case xNode.SelectSingleNode("Auth").InnerText
                        Case "0"
                            CheckBox3.Checked = False
                        Case "1" 'basic auth
                            CheckBox3.Checked = True
                            RadioButton2.Checked = True
                            TextBox13.Text = xNode.SelectSingleNode("SmtpUserName").InnerText
                            TextBox12.Text = DecPW(xNode.SelectSingleNode("SmtpPassword").InnerText)
                        Case "2" 'use same POP username passowrd
                            CheckBox3.Checked = True
                            RadioButton1.Checked = True
                    End Select
                End If

                If Not xNode.SelectSingleNode("IfNoOfMsgPerConn") Is Nothing Then
                    If xNode.SelectSingleNode("IfNoOfMsgPerConn").InnerText = "True" Then
                        CheckBox5.Checked = True
                    End If
                    NumericUpDown2.Value = CDec(xNode.SelectSingleNode("NoOfMsgPerConn").InnerText)
                End If

                If Not xNode.SelectSingleNode("IfNoOfRecPerMsg") Is Nothing Then
                    If xNode.SelectSingleNode("IfNoOfRecPerMsg").InnerText = "True" Then
                        CheckBox6.Checked = True
                    End If
                    NumericUpDown3.Value = CDec(xNode.SelectSingleNode("NoOfRecPerMsg").InnerText)
                End If

            End If

            TextBox5.Text = xNode.SelectSingleNode("SvrAccountName").InnerText
            TextBox6.Text = DecPW(xNode.SelectSingleNode("SvrAccountPassword").InnerText)
            Dim xScript As XmlNode
            Dim xitem As ListViewItem
            For Each xScript In xNode.SelectNodes("Scripts/Script")
                xitem = ListView1.Items.Add(xScript.SelectSingleNode("Name").InnerText)
                xitem.SubItems.Add(xScript.SelectSingleNode("Path").InnerText)
            Next
        End If
        GroupBox8.Enabled = CheckBox3.Checked
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        SeeCheckBox2()
    End Sub

    Private Function InsertTextNode(ByVal xDoc As XmlDocument, ByVal xNode As XmlNode, ByVal strTag As String, ByVal strText As String) As XmlElement
        Dim xNodeTemp As XmlNode
        xNodeTemp = xDoc.CreateElement(strTag)
        xNodeTemp.AppendChild(xDoc.CreateTextNode(strText))
        xNode.AppendChild(xNodeTemp)
        Return CType(xNodeTemp, XmlElement)
    End Function

    Private Sub ButtonOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonOK.Click
        If Trim(TextBox7.Text) = "" Then
            TabControl1.SelectedIndex = 0
            TextBox7.Focus()
            ErrorProvider1.SetError(TextBox7, "Please enter a friendly name. It cannot be empty.")
        ElseIf Validator1.ValidateAll Then
            Dim xDoc As New XmlDocument
            xDoc.Load(mainStorage & "accounts.xml")
            Dim xNode As XmlNode
            Dim xRoot As XmlNode
            xRoot = xDoc.SelectSingleNode("Accounts")
            If Not (xRoot Is Nothing) Then
                If oldAccountName <> "" Then
                    xNode = xRoot.SelectSingleNode("descendant::Account[Name='" & oldAccountName & "']")
                    If Not (xNode Is Nothing) Then
                        If oldAccountName <> TextBox7.Text Then
                            Dim xCheckNode As XmlNode = xRoot.SelectSingleNode("descendant::Account[Name='" & TextBox7.Text & "']")
                            If Not (xCheckNode Is Nothing) Then
                                TabControl1.SelectedIndex = 0
                                TextBox7.Focus()
                                ErrorProvider1.SetError(TextBox7, "The name already in use by other account.")
                                Exit Sub
                            End If
                        End If
                        xRoot.RemoveChild(xNode)
                    End If
                Else
                    xNode = xRoot.SelectSingleNode("descendant::Account[Name='" & TextBox7.Text & "']")
                    If Not (xNode Is Nothing) Then
                        TabControl1.SelectedIndex = 0
                        TextBox7.Focus()
                        ErrorProvider1.SetError(TextBox7, "The name already in use by other account.")
                        Exit Sub
                    End If
                End If
            End If
            Dim xElmntAccount As XmlElement
            ' Search for a particular node
            xNode = xDoc.SelectSingleNode("Accounts")

            If Not (xNode Is Nothing) Then
                'Dim nAct As XmlNode = xElmntAccount.SelectSingleNode(TextBox7.Text)
                'If Not (nAct Is Nothing) Then xElmntAccount.RemoveChild(nAct)

                ' Insert node for each family member:

                xElmntAccount = xDoc.CreateElement("Account")
                Dim xNodeScripts As XmlNode
                xNodeScripts = xNode.AppendChild(xElmntAccount)
                InsertTextNode(xDoc, xElmntAccount, "Name", TextBox7.Text)
                InsertTextNode(xDoc, xElmntAccount, "UserName", TextBox8.Text)
                InsertTextNode(xDoc, xElmntAccount, "Organization", TextBox9.Text)
                InsertTextNode(xDoc, xElmntAccount, "EmailAdr", TextBox10.Text)
                InsertTextNode(xDoc, xElmntAccount, "ReplyAdr", TextBox11.Text)
                InsertTextNode(xDoc, xElmntAccount, "Included", CheckBox4.Checked.ToString)
                InsertTextNode(xDoc, xElmntAccount, "POPPort", TextBox2.Text)
                InsertTextNode(xDoc, xElmntAccount, "SMTPPort", TextBox1.Text)
                InsertTextNode(xDoc, xElmntAccount, "LeaveOnSvr", CheckBox1.Checked.ToString)
                InsertTextNode(xDoc, xElmntAccount, "RemoveFromSvr", CheckBox2.Checked.ToString)
                InsertTextNode(xDoc, xElmntAccount, "HowManyDays", NumericUpDown1.Value.ToString)
                InsertTextNode(xDoc, xElmntAccount, "POPSvr", TextBox3.Text)
                InsertTextNode(xDoc, xElmntAccount, "SMTPSvr", TextBox4.Text)
                InsertTextNode(xDoc, xElmntAccount, "SvrAccountName", TextBox5.Text)
                InsertTextNode(xDoc, xElmntAccount, "SvrAccountPassword", EncPW(TextBox6.Text))
                If Not CheckBox3.Checked Then
                    InsertTextNode(xDoc, xElmntAccount, "Auth", "0")
                Else
                    If RadioButton2.Checked Then
                        InsertTextNode(xDoc, xElmntAccount, "Auth", "1")
                    Else
                        InsertTextNode(xDoc, xElmntAccount, "Auth", "2")
                    End If
                End If
                InsertTextNode(xDoc, xElmntAccount, "SmtpUserName", TextBox13.Text)
                InsertTextNode(xDoc, xElmntAccount, "SmtpPassword", EncPW(TextBox12.Text))

                If CheckBox5.Checked Then
                    InsertTextNode(xDoc, xElmntAccount, "IfNoOfMsgPerConn", "True")
                Else
                    InsertTextNode(xDoc, xElmntAccount, "IfNoOfMsgPerConn", "False")
                End If
                InsertTextNode(xDoc, xElmntAccount, "NoOfMsgPerConn", NumericUpDown2.Value.ToString)

                If CheckBox6.Checked Then
                    InsertTextNode(xDoc, xElmntAccount, "IfNoOfRecPerMsg", "True")
                Else
                    InsertTextNode(xDoc, xElmntAccount, "IfNoOfRecPerMsg", "False")
                End If
                InsertTextNode(xDoc, xElmntAccount, "NoOfRecPerMsg", NumericUpDown3.Value.ToString)

                xElmntAccount = xDoc.CreateElement("Scripts")
                Dim nScripts As XmlNode
                nScripts = xNodeScripts.AppendChild(xElmntAccount)
                Dim xitem As ListViewItem
                Dim nScript As XmlNode
                For Each xitem In ListView1.Items
                    xElmntAccount = xDoc.CreateElement("Script")
                    nScript = nScripts.AppendChild(xElmntAccount)
                    InsertTextNode(xDoc, nScript, "Name", xitem.Text)
                    InsertTextNode(xDoc, nScript, "Path", xitem.SubItems(1).Text)
                Next
                xDoc.Save(mainStorage & "accounts.xml")
                Me.Close()
            Else
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OpenFileDialog1.DefaultExt = ".Script"
        OpenFileDialog1.Filter = "Faena Mail Script File (*.Script)|*.Script|All files (*.*)|*.*"
        If OpenFileDialog1.ShowDialog = DialogResult.OK Then
            Dim strFullName As String = OpenFileDialog1.FileName
            Dim lItem As ListViewItem = ListView1.Items.Add(strFullName.Substring(strFullName.LastIndexOf("\") + 1))
            lItem.SubItems.Add(strFullName)
        End If
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim i As Integer
        For i = ListView1.Items.Count - 1 To 0 Step -1
            If ListView1.Items(i).Selected Then
                ListView1.Items(i).Remove()
            End If
        Next
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim tmpItem As ListViewItem
        Dim curIdx As Integer
        For Each curIdx In ListView1.SelectedIndices
            ' check if it's not the first item, which cannot be moved further up
            If curIdx > 0 Then
                tmpItem = ListView1.Items.Item(curIdx - 1)
                ListView1.Items.Item(curIdx - 1) = CType(ListView1.Items.Item(curIdx).Clone, ListViewItem)
                ListView1.Items.Item(curIdx - 1).Selected = True
                ListView1.Items.Item(curIdx) = tmpItem
            End If
        Next
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim tmpItem As ListViewItem
        Dim curIdx As Integer
        For curIdx = ListView1.Items.Count - 2 To 0 Step -1
            ' check if it's not the last item, which cannot be moved further down
            If ListView1.Items.Item(curIdx).Selected Then
                tmpItem = ListView1.Items.Item(curIdx + 1)
                ListView1.Items.Item(curIdx + 1) = CType(ListView1.Items.Item(curIdx).Clone, ListViewItem)
                ListView1.Items.Item(curIdx + 1).Selected = True
                ListView1.Items.Item(curIdx) = tmpItem
            End If
        Next
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim frm As New FormScriptEditor(ListView1.SelectedItems(0).SubItems(1).Text)
        frm.Show()
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Try
            If ListView1.SelectedItems(0) Is Nothing Then
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button6.Enabled = False
            Else
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                Button6.Enabled = True
            End If
        Catch
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button6.Enabled = False
        End Try
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        Button6.PerformClick()
    End Sub

    Private Sub CheckBox3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox3.CheckedChanged
        GroupBox8.Enabled = CheckBox3.Checked
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        Panel1.Enabled = RadioButton2.Checked
    End Sub

    Private Sub CheckBox5_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox5.CheckedChanged
        NumericUpDown2.Enabled = CheckBox5.Checked
    End Sub

    Private Sub CheckBox6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox6.CheckedChanged
        NumericUpDown3.Enabled = CheckBox6.Checked
    End Sub
End Class
