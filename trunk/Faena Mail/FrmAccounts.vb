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
Public Class FormAccounts
    Inherits System.Windows.Forms.Form
    ' Module level variables for working with the DOM
    Private xElem As XmlElement
    Private mxDoc As XmlDocument
    Private mxNode As XmlNode
    Private mxNodeList As XmlNodeList
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
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents Button4 As System.Windows.Forms.Button
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents ThreeDLine1 As FaenaMail.ThreeDLine
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ListView1 = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader()
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.ThreeDLine1 = New FaenaMail.ThreeDLine()
        Me.SuspendLayout()
        '
        'ListView1
        '
        Me.ListView1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.ListView1.FullRowSelect = True
        Me.ListView1.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.ListView1.HideSelection = False
        Me.ListView1.Location = New System.Drawing.Point(10, 23)
        Me.ListView1.MultiSelect = False
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(406, 208)
        Me.ListView1.TabIndex = 0
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Account"
        Me.ColumnHeader1.Width = 150
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Status"
        '
        'Button1
        '
        Me.Button1.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button1.Location = New System.Drawing.Point(442, 28)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(107, 29)
        Me.Button1.TabIndex = 1
        Me.Button1.Text = "Add"
        '
        'Button2
        '
        Me.Button2.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button2.Enabled = False
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button2.Location = New System.Drawing.Point(442, 83)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(107, 28)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Remove"
        '
        'Button3
        '
        Me.Button3.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button3.Enabled = False
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button3.Location = New System.Drawing.Point(442, 137)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(107, 28)
        Me.Button3.TabIndex = 3
        Me.Button3.Text = "Properties"
        '
        'Button4
        '
        Me.Button4.Anchor = (System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button4.Enabled = False
        Me.Button4.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button4.Location = New System.Drawing.Point(442, 191)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(107, 29)
        Me.Button4.TabIndex = 4
        Me.Button4.Text = "Set as Default"
        Me.Button4.Visible = False
        '
        'Button5
        '
        Me.Button5.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right)
        Me.Button5.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.System
        Me.Button5.Location = New System.Drawing.Point(443, 254)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(107, 29)
        Me.Button5.TabIndex = 5
        Me.Button5.Text = "Close"
        '
        'ThreeDLine1
        '
        Me.ThreeDLine1.Anchor = ((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.ThreeDLine1.Location = New System.Drawing.Point(17, 241)
        Me.ThreeDLine1.Name = "ThreeDLine1"
        Me.ThreeDLine1.Orientation = FaenaMail.ThreeDLine.tOrientation.Horizontal
        Me.ThreeDLine1.Size = New System.Drawing.Size(532, 2)
        Me.ThreeDLine1.TabIndex = 6
        '
        'FormAccounts
        '
        Me.AcceptButton = Me.Button5
        Me.AutoScaleBaseSize = New System.Drawing.Size(6, 15)
        Me.CancelButton = Me.Button5
        Me.ClientSize = New System.Drawing.Size(568, 291)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.ThreeDLine1, Me.Button5, Me.Button4, Me.Button3, Me.Button2, Me.Button1, Me.ListView1})
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FormAccounts"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Accounts"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Private Sub LaunchProperties(ByVal actID As String)
        Dim frm As New FormProperties(actID)
        frm.ShowDialog()
        LoadListView()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            LaunchProperties(ListView1.SelectedItems(0).Text)
        Catch
        End Try
    End Sub

    Private Sub LoadXMLFromFile()
        If Not (mxDoc Is Nothing) Then mxDoc = Nothing
        Me.LoadXMLFromFile(False)
    End Sub

    Private Sub LoadXMLFromFile(ByVal ForceReload As Boolean)
        If ForceReload OrElse (mxDoc Is Nothing) Then
            mxDoc = New XmlDocument()
            mxDoc.Load(mainStorage & "accounts.xml")
        End If
    End Sub

    Private Sub LoadListView()
        LoadXMLFromFile()
        mxNodeList = mxDoc.SelectNodes("/Accounts/Account")
        Dim itmx As ListViewItem
        ListView1.Items.Clear()
        For Each mxNode In mxNodeList
            If Not (mxNode.SelectSingleNode("Name") Is Nothing) Then
                itmx = ListView1.Items.Add(mxNode.SelectSingleNode("Name").InnerText)
            End If
            If Not (mxNode.SelectSingleNode("Included") Is Nothing) Then
                itmx.SubItems.Add(mxNode.SelectSingleNode("Included").InnerText)
            End If
        Next mxNode
    End Sub

    Private Sub CreateXML()
        Dim xDoc As New XmlDocument()
        Dim xPI As XmlProcessingInstruction
        Dim xComment As XmlComment
        Dim xElmntRoot As XmlElement
        Dim xElmntFamily As XmlElement
        xPI = xDoc.CreateProcessingInstruction("xml", "version='1.0'")
        xDoc.AppendChild(xPI)

        xComment = xDoc.CreateComment("Faena Mail Account List")
        xDoc.AppendChild(xComment)

        xElmntRoot = xDoc.CreateElement("Accounts")
        xDoc.AppendChild(xElmntRoot)
        xDoc.Save(mainStorage & "accounts.xml")
    End Sub

    Private Sub FormAccounts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If Not File.Exists(mainStorage & "accounts.xml") Then
            CreateXML()
        End If
        LoadListView()
    End Sub

    Private Sub ListView1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView1.DoubleClick
        LaunchProperties(ListView1.SelectedItems(0).Text)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        LaunchProperties("")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If MsgBox("Are you sure want to delete the '" + ListView1.SelectedItems(0).Text + "' account?", CType(MsgBoxStyle.YesNo + MsgBoxStyle.Exclamation, MsgBoxStyle)) = MsgBoxResult.Yes Then
                Dim xDoc As New XmlDocument()
                xDoc.Load(mainStorage & "accounts.xml")
                Dim xNode As XmlNode
                Dim xRoot As XmlNode
                Dim xitem As ListViewItem
                xRoot = xDoc.SelectSingleNode("Accounts")
                If Not (xRoot Is Nothing) Then
                    For Each xitem In ListView1.SelectedItems
                        xNode = xRoot.SelectSingleNode("descendant::Account[Name='" & xitem.Text & "']")
                        If Not (xNode Is Nothing) Then
                            xRoot.RemoveChild(xNode)
                        End If
                    Next
                    xDoc.Save(mainStorage & "accounts.xml")
                End If
                LoadListView()
            End If
        Catch
        End Try
    End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListView1.SelectedIndexChanged
        Try
            If Not ListView1.SelectedItems(0) Is Nothing Then
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
            Else
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
            End If
        Catch
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        End Try
    End Sub
End Class
