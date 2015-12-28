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
Option Compare Text
Imports System.IO
'Imports System.Data.Odbc
Imports System.Text
'Imports System.Data.OleDb
Imports Microsoft.Win32
Imports System.Xml
Imports System.Drawing
Imports vb = Microsoft.VisualBasic
Imports System.ServiceProcess
Imports System.Runtime.InteropServices

Public Class FrmStart
    Inherits System.Windows.Forms.Form
    Private MyNode As TreeNode
    Private displayClock As Boolean
    Private mUnreadCnt As Integer
    Private mObjectCnt As Integer
    Private bCloseClicked As Boolean
    Public Const SC_CLOSE As Integer = 61536
    Public Const WM_SYSCOMMAND As Integer = 274
    Dim currentPath As String
    Private cls As FaenaMailLibrary.SedRecMail

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
    Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
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
    Friend WithEvents ListView1 As System.Windows.Forms.ListView
    Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
    Friend WithEvents MenuItem13 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem14 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem16 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem17 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem18 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem19 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem20 As System.Windows.Forms.MenuItem
    Friend WithEvents ColumnHeader4 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader5 As System.Windows.Forms.ColumnHeader
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents TreeView1 As System.Windows.Forms.TreeView
    Friend WithEvents Splitter2 As System.Windows.Forms.Splitter
    Friend WithEvents ListView2 As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader3 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader6 As System.Windows.Forms.ColumnHeader
    Friend WithEvents FSWFile As System.IO.FileSystemWatcher
    Friend WithEvents FSWDir As System.IO.FileSystemWatcher
    Friend WithEvents StatusBarPanel1 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents StatusBarPanel2 As System.Windows.Forms.StatusBarPanel
    Friend WithEvents NotifyIcon1 As System.Windows.Forms.NotifyIcon
    Friend WithEvents ContextMenu1 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem24 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem25 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem26 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem27 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem28 As System.Windows.Forms.MenuItem
    Friend WithEvents Timer2 As System.Windows.Forms.Timer
    Friend WithEvents MenuItem29 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem30 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem31 As System.Windows.Forms.MenuItem
    Friend WithEvents ImageList1 As System.Windows.Forms.ImageList
    Friend WithEvents MenuItem32 As System.Windows.Forms.MenuItem
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Timer3 As System.Windows.Forms.Timer
    Friend WithEvents MenuItem33 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem34 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem15 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem35 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem21 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem22 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem36 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem37 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem38 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem39 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem40 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem41 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem42 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem44 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem45 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem46 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu2 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem50 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem51 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem52 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem53 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem54 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem55 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem56 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem58 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem59 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem60 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem61 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem47 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem57 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem48 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem49 As System.Windows.Forms.MenuItem
    Friend WithEvents ContextMenu3 As System.Windows.Forms.ContextMenu
    Friend WithEvents MenuItem63 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem64 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem65 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem43 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem62 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem66 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem67 As System.Windows.Forms.MenuItem
    Friend WithEvents NotifyIcon2 As System.Windows.Forms.NotifyIcon
    Friend WithEvents MenuItem23 As System.Windows.Forms.MenuItem
    Friend WithEvents MenuItem68 As System.Windows.Forms.MenuItem
    Friend WithEvents ImageList2 As System.Windows.Forms.ImageList
    Friend WithEvents tmrStatus As System.Windows.Forms.Timer
    Friend WithEvents tmrFSWatch As System.Windows.Forms.Timer
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmStart))
        Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
        Me.MenuItem1 = New System.Windows.Forms.MenuItem
        Me.MenuItem20 = New System.Windows.Forms.MenuItem
        Me.MenuItem62 = New System.Windows.Forms.MenuItem
        Me.MenuItem66 = New System.Windows.Forms.MenuItem
        Me.MenuItem19 = New System.Windows.Forms.MenuItem
        Me.MenuItem31 = New System.Windows.Forms.MenuItem
        Me.MenuItem6 = New System.Windows.Forms.MenuItem
        Me.MenuItem2 = New System.Windows.Forms.MenuItem
        Me.MenuItem8 = New System.Windows.Forms.MenuItem
        Me.MenuItem18 = New System.Windows.Forms.MenuItem
        Me.MenuItem9 = New System.Windows.Forms.MenuItem
        Me.MenuItem17 = New System.Windows.Forms.MenuItem
        Me.MenuItem12 = New System.Windows.Forms.MenuItem
        Me.MenuItem11 = New System.Windows.Forms.MenuItem
        Me.MenuItem16 = New System.Windows.Forms.MenuItem
        Me.MenuItem68 = New System.Windows.Forms.MenuItem
        Me.MenuItem23 = New System.Windows.Forms.MenuItem
        Me.MenuItem7 = New System.Windows.Forms.MenuItem
        Me.MenuItem3 = New System.Windows.Forms.MenuItem
        Me.MenuItem10 = New System.Windows.Forms.MenuItem
        Me.MenuItem15 = New System.Windows.Forms.MenuItem
        Me.MenuItem35 = New System.Windows.Forms.MenuItem
        Me.MenuItem43 = New System.Windows.Forms.MenuItem
        Me.MenuItem21 = New System.Windows.Forms.MenuItem
        Me.MenuItem22 = New System.Windows.Forms.MenuItem
        Me.MenuItem36 = New System.Windows.Forms.MenuItem
        Me.MenuItem37 = New System.Windows.Forms.MenuItem
        Me.MenuItem38 = New System.Windows.Forms.MenuItem
        Me.MenuItem39 = New System.Windows.Forms.MenuItem
        Me.MenuItem41 = New System.Windows.Forms.MenuItem
        Me.MenuItem42 = New System.Windows.Forms.MenuItem
        Me.MenuItem40 = New System.Windows.Forms.MenuItem
        Me.MenuItem44 = New System.Windows.Forms.MenuItem
        Me.MenuItem46 = New System.Windows.Forms.MenuItem
        Me.MenuItem45 = New System.Windows.Forms.MenuItem
        Me.MenuItem4 = New System.Windows.Forms.MenuItem
        Me.MenuItem13 = New System.Windows.Forms.MenuItem
        Me.MenuItem14 = New System.Windows.Forms.MenuItem
        Me.MenuItem33 = New System.Windows.Forms.MenuItem
        Me.MenuItem34 = New System.Windows.Forms.MenuItem
        Me.MenuItem49 = New System.Windows.Forms.MenuItem
        Me.MenuItem48 = New System.Windows.Forms.MenuItem
        Me.MenuItem67 = New System.Windows.Forms.MenuItem
        Me.MenuItem5 = New System.Windows.Forms.MenuItem
        Me.ListView1 = New System.Windows.Forms.ListView
        Me.ColumnHeader4 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader5 = New System.Windows.Forms.ColumnHeader
        Me.ImageList2 = New System.Windows.Forms.ImageList(Me.components)
        Me.StatusBar1 = New System.Windows.Forms.StatusBar
        Me.StatusBarPanel1 = New System.Windows.Forms.StatusBarPanel
        Me.StatusBarPanel2 = New System.Windows.Forms.StatusBarPanel
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.TreeView1 = New System.Windows.Forms.TreeView
        Me.ContextMenu3 = New System.Windows.Forms.ContextMenu
        Me.MenuItem63 = New System.Windows.Forms.MenuItem
        Me.MenuItem64 = New System.Windows.Forms.MenuItem
        Me.MenuItem65 = New System.Windows.Forms.MenuItem
        Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
        Me.Splitter2 = New System.Windows.Forms.Splitter
        Me.ListView2 = New System.Windows.Forms.ListView
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader3 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader6 = New System.Windows.Forms.ColumnHeader
        Me.ContextMenu2 = New System.Windows.Forms.ContextMenu
        Me.MenuItem47 = New System.Windows.Forms.MenuItem
        Me.MenuItem57 = New System.Windows.Forms.MenuItem
        Me.MenuItem50 = New System.Windows.Forms.MenuItem
        Me.MenuItem51 = New System.Windows.Forms.MenuItem
        Me.MenuItem52 = New System.Windows.Forms.MenuItem
        Me.MenuItem53 = New System.Windows.Forms.MenuItem
        Me.MenuItem59 = New System.Windows.Forms.MenuItem
        Me.MenuItem60 = New System.Windows.Forms.MenuItem
        Me.MenuItem61 = New System.Windows.Forms.MenuItem
        Me.MenuItem54 = New System.Windows.Forms.MenuItem
        Me.MenuItem55 = New System.Windows.Forms.MenuItem
        Me.MenuItem56 = New System.Windows.Forms.MenuItem
        Me.MenuItem58 = New System.Windows.Forms.MenuItem
        Me.FSWFile = New System.IO.FileSystemWatcher
        Me.FSWDir = New System.IO.FileSystemWatcher
        Me.NotifyIcon1 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.ContextMenu1 = New System.Windows.Forms.ContextMenu
        Me.MenuItem32 = New System.Windows.Forms.MenuItem
        Me.MenuItem25 = New System.Windows.Forms.MenuItem
        Me.MenuItem24 = New System.Windows.Forms.MenuItem
        Me.MenuItem26 = New System.Windows.Forms.MenuItem
        Me.MenuItem29 = New System.Windows.Forms.MenuItem
        Me.MenuItem30 = New System.Windows.Forms.MenuItem
        Me.MenuItem27 = New System.Windows.Forms.MenuItem
        Me.MenuItem28 = New System.Windows.Forms.MenuItem
        Me.Timer2 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Timer3 = New System.Windows.Forms.Timer(Me.components)
        Me.NotifyIcon2 = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.tmrStatus = New System.Windows.Forms.Timer(Me.components)
        Me.tmrFSWatch = New System.Windows.Forms.Timer(Me.components)
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSWFile, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.FSWDir, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MainMenu1
        '
        Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem1, Me.MenuItem2, Me.MenuItem15, Me.MenuItem4})
        '
        'MenuItem1
        '
        Me.MenuItem1.Index = 0
        Me.MenuItem1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem20, Me.MenuItem19, Me.MenuItem31, Me.MenuItem6})
        Me.MenuItem1.Text = "&File"
        '
        'MenuItem20
        '
        Me.MenuItem20.Index = 0
        Me.MenuItem20.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem62, Me.MenuItem66})
        Me.MenuItem20.Text = "&New"
        '
        'MenuItem62
        '
        Me.MenuItem62.Index = 0
        Me.MenuItem62.Text = "Mail Message"
        '
        'MenuItem66
        '
        Me.MenuItem66.Index = 1
        Me.MenuItem66.Text = "Mail List Message"
        '
        'MenuItem19
        '
        Me.MenuItem19.Index = 1
        Me.MenuItem19.Text = "-"
        '
        'MenuItem31
        '
        Me.MenuItem31.Index = 2
        Me.MenuItem31.Text = "&Close"
        '
        'MenuItem6
        '
        Me.MenuItem6.Index = 3
        Me.MenuItem6.Text = "Close & E&xit"
        '
        'MenuItem2
        '
        Me.MenuItem2.Index = 1
        Me.MenuItem2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem8, Me.MenuItem12, Me.MenuItem11, Me.MenuItem16, Me.MenuItem68, Me.MenuItem23, Me.MenuItem7, Me.MenuItem3, Me.MenuItem10})
        Me.MenuItem2.Text = "&Tools"
        '
        'MenuItem8
        '
        Me.MenuItem8.Index = 0
        Me.MenuItem8.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem18, Me.MenuItem9, Me.MenuItem17})
        Me.MenuItem8.Text = "&Send and Receive"
        '
        'MenuItem18
        '
        Me.MenuItem18.Index = 0
        Me.MenuItem18.Text = "&Send and Receive All"
        '
        'MenuItem9
        '
        Me.MenuItem9.Index = 1
        Me.MenuItem9.Text = "&Receive All"
        '
        'MenuItem17
        '
        Me.MenuItem17.Index = 2
        Me.MenuItem17.Text = "&Send All"
        '
        'MenuItem12
        '
        Me.MenuItem12.Index = 1
        Me.MenuItem12.Text = "-"
        '
        'MenuItem11
        '
        Me.MenuItem11.Index = 2
        Me.MenuItem11.Text = "S&cript Editor"
        '
        'MenuItem16
        '
        Me.MenuItem16.Index = 3
        Me.MenuItem16.Text = "&Mailing List Manager"
        Me.MenuItem16.Visible = False
        '
        'MenuItem68
        '
        Me.MenuItem68.Index = 4
        Me.MenuItem68.Text = "&TXT2DB"
        '
        'MenuItem23
        '
        Me.MenuItem23.Index = 5
        Me.MenuItem23.Text = "&External Tools..."
        '
        'MenuItem7
        '
        Me.MenuItem7.Index = 6
        Me.MenuItem7.Text = "-"
        '
        'MenuItem3
        '
        Me.MenuItem3.Index = 7
        Me.MenuItem3.Text = "&Accounts..."
        '
        'MenuItem10
        '
        Me.MenuItem10.Index = 8
        Me.MenuItem10.Text = "&Options..."
        '
        'MenuItem15
        '
        Me.MenuItem15.Index = 2
        Me.MenuItem15.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem35, Me.MenuItem43, Me.MenuItem21, Me.MenuItem22, Me.MenuItem36, Me.MenuItem37, Me.MenuItem38, Me.MenuItem39, Me.MenuItem41, Me.MenuItem42, Me.MenuItem40, Me.MenuItem44, Me.MenuItem46, Me.MenuItem45})
        Me.MenuItem15.Text = "&Message"
        '
        'MenuItem35
        '
        Me.MenuItem35.Index = 0
        Me.MenuItem35.Text = "&New Message"
        '
        'MenuItem43
        '
        Me.MenuItem43.Index = 1
        Me.MenuItem43.Text = "New Mail &List Message"
        '
        'MenuItem21
        '
        Me.MenuItem21.Index = 2
        Me.MenuItem21.Text = "-"
        '
        'MenuItem22
        '
        Me.MenuItem22.Index = 3
        Me.MenuItem22.Text = "&Reply to Sender"
        '
        'MenuItem36
        '
        Me.MenuItem36.Index = 4
        Me.MenuItem36.Text = "Reply to &All"
        '
        'MenuItem37
        '
        Me.MenuItem37.Index = 5
        Me.MenuItem37.Text = "&Forward"
        '
        'MenuItem38
        '
        Me.MenuItem38.Index = 6
        Me.MenuItem38.Text = "Forward as Attachment"
        Me.MenuItem38.Visible = False
        '
        'MenuItem39
        '
        Me.MenuItem39.Index = 7
        Me.MenuItem39.Text = "-"
        '
        'MenuItem41
        '
        Me.MenuItem41.Index = 8
        Me.MenuItem41.Text = "&Move to Folder..."
        '
        'MenuItem42
        '
        Me.MenuItem42.Index = 9
        Me.MenuItem42.Text = "&Copy to Folder..."
        '
        'MenuItem40
        '
        Me.MenuItem40.Index = 10
        Me.MenuItem40.Text = "&Delete"
        '
        'MenuItem44
        '
        Me.MenuItem44.Index = 11
        Me.MenuItem44.Text = "-"
        '
        'MenuItem46
        '
        Me.MenuItem46.Index = 12
        Me.MenuItem46.Text = "Mark as Read"
        '
        'MenuItem45
        '
        Me.MenuItem45.Index = 13
        Me.MenuItem45.Text = "Mark as Unread"
        '
        'MenuItem4
        '
        Me.MenuItem4.Index = 3
        Me.MenuItem4.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem13, Me.MenuItem14, Me.MenuItem33, Me.MenuItem34, Me.MenuItem49, Me.MenuItem48, Me.MenuItem67, Me.MenuItem5})
        Me.MenuItem4.Text = "&Help"
        '
        'MenuItem13
        '
        Me.MenuItem13.Index = 0
        Me.MenuItem13.Text = "&Feana on the Web"
        '
        'MenuItem14
        '
        Me.MenuItem14.Index = 1
        Me.MenuItem14.Text = "&Send Feedback"
        '
        'MenuItem33
        '
        Me.MenuItem33.Index = 2
        Me.MenuItem33.Text = "&View Log File"
        '
        'MenuItem34
        '
        Me.MenuItem34.Index = 3
        Me.MenuItem34.Text = "-"
        '
        'MenuItem49
        '
        Me.MenuItem49.Index = 4
        Me.MenuItem49.Text = "&Check for Update..."
        '
        'MenuItem48
        '
        Me.MenuItem48.Index = 5
        Me.MenuItem48.Text = "-"
        '
        'MenuItem67
        '
        Me.MenuItem67.Index = 6
        Me.MenuItem67.Text = "&Enter Registration Code..."
        '
        'MenuItem5
        '
        Me.MenuItem5.Index = 7
        Me.MenuItem5.Text = "&About..."
        '
        'ListView1
        '
        Me.ListView1.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader4, Me.ColumnHeader5})
        Me.ListView1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.ListView1.FullRowSelect = True
        Me.ListView1.GridLines = True
        Me.ListView1.Location = New System.Drawing.Point(0, 244)
        Me.ListView1.Name = "ListView1"
        Me.ListView1.Size = New System.Drawing.Size(593, 82)
        Me.ListView1.StateImageList = Me.ImageList2
        Me.ListView1.TabIndex = 3
        Me.ListView1.UseCompatibleStateImageBehavior = False
        Me.ListView1.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader4
        '
        Me.ColumnHeader4.Text = "Tasks"
        Me.ColumnHeader4.Width = 200
        '
        'ColumnHeader5
        '
        Me.ColumnHeader5.Text = "Status"
        Me.ColumnHeader5.Width = 100
        '
        'ImageList2
        '
        Me.ImageList2.ImageStream = CType(resources.GetObject("ImageList2.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList2.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList2.Images.SetKeyName(0, "")
        Me.ImageList2.Images.SetKeyName(1, "")
        '
        'StatusBar1
        '
        Me.StatusBar1.Location = New System.Drawing.Point(0, 326)
        Me.StatusBar1.Name = "StatusBar1"
        Me.StatusBar1.Panels.AddRange(New System.Windows.Forms.StatusBarPanel() {Me.StatusBarPanel1, Me.StatusBarPanel2})
        Me.StatusBar1.ShowPanels = True
        Me.StatusBar1.Size = New System.Drawing.Size(593, 25)
        Me.StatusBar1.TabIndex = 4
        '
        'StatusBarPanel1
        '
        Me.StatusBarPanel1.AutoSize = System.Windows.Forms.StatusBarPanelAutoSize.Spring
        Me.StatusBarPanel1.Name = "StatusBarPanel1"
        Me.StatusBarPanel1.Width = 476
        '
        'StatusBarPanel2
        '
        Me.StatusBarPanel2.Name = "StatusBarPanel2"
        '
        'Splitter1
        '
        Me.Splitter1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Splitter1.Location = New System.Drawing.Point(0, 240)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(593, 4)
        Me.Splitter1.TabIndex = 5
        Me.Splitter1.TabStop = False
        '
        'TreeView1
        '
        Me.TreeView1.ContextMenu = Me.ContextMenu3
        Me.TreeView1.Dock = System.Windows.Forms.DockStyle.Left
        Me.TreeView1.FullRowSelect = True
        Me.TreeView1.HideSelection = False
        Me.TreeView1.ImageIndex = 1
        Me.TreeView1.ImageList = Me.ImageList1
        Me.TreeView1.Location = New System.Drawing.Point(0, 0)
        Me.TreeView1.Name = "TreeView1"
        Me.TreeView1.SelectedImageIndex = 2
        Me.TreeView1.Size = New System.Drawing.Size(142, 240)
        Me.TreeView1.TabIndex = 6
        '
        'ContextMenu3
        '
        Me.ContextMenu3.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem63, Me.MenuItem64, Me.MenuItem65})
        '
        'MenuItem63
        '
        Me.MenuItem63.Index = 0
        Me.MenuItem63.Text = "New Folder..."
        '
        'MenuItem64
        '
        Me.MenuItem64.Index = 1
        Me.MenuItem64.Text = "Rename..."
        '
        'MenuItem65
        '
        Me.MenuItem65.Index = 2
        Me.MenuItem65.Shortcut = System.Windows.Forms.Shortcut.Del
        Me.MenuItem65.ShowShortcut = False
        Me.MenuItem65.Text = "Delete"
        '
        'ImageList1
        '
        Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
        Me.ImageList1.Images.SetKeyName(0, "")
        Me.ImageList1.Images.SetKeyName(1, "")
        Me.ImageList1.Images.SetKeyName(2, "")
        Me.ImageList1.Images.SetKeyName(3, "")
        Me.ImageList1.Images.SetKeyName(4, "")
        '
        'Splitter2
        '
        Me.Splitter2.Location = New System.Drawing.Point(142, 0)
        Me.Splitter2.Name = "Splitter2"
        Me.Splitter2.Size = New System.Drawing.Size(3, 240)
        Me.Splitter2.TabIndex = 7
        Me.Splitter2.TabStop = False
        '
        'ListView2
        '
        Me.ListView2.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader6})
        Me.ListView2.ContextMenu = Me.ContextMenu2
        Me.ListView2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListView2.FullRowSelect = True
        Me.ListView2.HideSelection = False
        Me.ListView2.LargeImageList = Me.ImageList1
        Me.ListView2.Location = New System.Drawing.Point(145, 0)
        Me.ListView2.Name = "ListView2"
        Me.ListView2.Size = New System.Drawing.Size(448, 240)
        Me.ListView2.SmallImageList = Me.ImageList1
        Me.ListView2.StateImageList = Me.ImageList1
        Me.ListView2.TabIndex = 8
        Me.ListView2.UseCompatibleStateImageBehavior = False
        Me.ListView2.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "From"
        Me.ColumnHeader2.Width = 100
        '
        'ColumnHeader3
        '
        Me.ColumnHeader3.Text = "Subject"
        Me.ColumnHeader3.Width = 200
        '
        'ColumnHeader6
        '
        Me.ColumnHeader6.Text = "Date"
        Me.ColumnHeader6.Width = 200
        '
        'ContextMenu2
        '
        Me.ContextMenu2.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem47, Me.MenuItem57, Me.MenuItem50, Me.MenuItem51, Me.MenuItem52, Me.MenuItem53, Me.MenuItem59, Me.MenuItem60, Me.MenuItem61, Me.MenuItem54, Me.MenuItem55, Me.MenuItem56, Me.MenuItem58})
        '
        'MenuItem47
        '
        Me.MenuItem47.DefaultItem = True
        Me.MenuItem47.Index = 0
        Me.MenuItem47.Text = "&Open"
        '
        'MenuItem57
        '
        Me.MenuItem57.Index = 1
        Me.MenuItem57.Text = "-"
        '
        'MenuItem50
        '
        Me.MenuItem50.Index = 2
        Me.MenuItem50.Text = "&Reply to Sender"
        '
        'MenuItem51
        '
        Me.MenuItem51.Index = 3
        Me.MenuItem51.Text = "Reply to &All"
        '
        'MenuItem52
        '
        Me.MenuItem52.Index = 4
        Me.MenuItem52.Text = "&Forward"
        '
        'MenuItem53
        '
        Me.MenuItem53.Index = 5
        Me.MenuItem53.Text = "Forward as Attachment"
        Me.MenuItem53.Visible = False
        '
        'MenuItem59
        '
        Me.MenuItem59.Index = 6
        Me.MenuItem59.Text = "-"
        '
        'MenuItem60
        '
        Me.MenuItem60.Index = 7
        Me.MenuItem60.Text = "Mark as Read"
        '
        'MenuItem61
        '
        Me.MenuItem61.Index = 8
        Me.MenuItem61.Text = "Mark as Unread"
        '
        'MenuItem54
        '
        Me.MenuItem54.Index = 9
        Me.MenuItem54.Text = "-"
        '
        'MenuItem55
        '
        Me.MenuItem55.Index = 10
        Me.MenuItem55.Text = "&Move to Folder..."
        '
        'MenuItem56
        '
        Me.MenuItem56.Index = 11
        Me.MenuItem56.Text = "&Copy to Folder..."
        '
        'MenuItem58
        '
        Me.MenuItem58.Index = 12
        Me.MenuItem58.Shortcut = System.Windows.Forms.Shortcut.Del
        Me.MenuItem58.ShowShortcut = False
        Me.MenuItem58.Text = "&Delete"
        '
        'FSWFile
        '
        Me.FSWFile.EnableRaisingEvents = True
        Me.FSWFile.Filter = "*.mail"
        Me.FSWFile.NotifyFilter = CType((System.IO.NotifyFilters.FileName Or System.IO.NotifyFilters.LastWrite), System.IO.NotifyFilters)
        Me.FSWFile.SynchronizingObject = Me
        '
        'FSWDir
        '
        Me.FSWDir.EnableRaisingEvents = True
        Me.FSWDir.IncludeSubdirectories = True
        Me.FSWDir.NotifyFilter = CType((System.IO.NotifyFilters.DirectoryName Or System.IO.NotifyFilters.LastWrite), System.IO.NotifyFilters)
        Me.FSWDir.SynchronizingObject = Me
        '
        'NotifyIcon1
        '
        Me.NotifyIcon1.ContextMenu = Me.ContextMenu1
        Me.NotifyIcon1.Icon = CType(resources.GetObject("NotifyIcon1.Icon"), System.Drawing.Icon)
        Me.NotifyIcon1.Text = "Faena Mail"
        Me.NotifyIcon1.Visible = True
        '
        'ContextMenu1
        '
        Me.ContextMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItem32, Me.MenuItem25, Me.MenuItem24, Me.MenuItem26, Me.MenuItem29, Me.MenuItem30, Me.MenuItem27, Me.MenuItem28})
        '
        'MenuItem32
        '
        Me.MenuItem32.Index = 0
        Me.MenuItem32.Text = "&About Faena Mail"
        '
        'MenuItem25
        '
        Me.MenuItem25.Index = 1
        Me.MenuItem25.Text = "-"
        '
        'MenuItem24
        '
        Me.MenuItem24.Index = 2
        Me.MenuItem24.Text = "&Compose a new message"
        '
        'MenuItem26
        '
        Me.MenuItem26.Index = 3
        Me.MenuItem26.Text = "-"
        '
        'MenuItem29
        '
        Me.MenuItem29.Index = 4
        Me.MenuItem29.Text = "&Display Time..."
        '
        'MenuItem30
        '
        Me.MenuItem30.Index = 5
        Me.MenuItem30.Text = "-"
        '
        'MenuItem27
        '
        Me.MenuItem27.DefaultItem = True
        Me.MenuItem27.Index = 6
        Me.MenuItem27.Text = "&Open"
        '
        'MenuItem28
        '
        Me.MenuItem28.Index = 7
        Me.MenuItem28.Text = "E&xit"
        '
        'Timer2
        '
        Me.Timer2.Interval = 5000
        '
        'Timer1
        '
        Me.Timer1.Interval = 10000
        '
        'Timer3
        '
        '
        'NotifyIcon2
        '
        Me.NotifyIcon2.Icon = CType(resources.GetObject("NotifyIcon2.Icon"), System.Drawing.Icon)
        Me.NotifyIcon2.Text = "New email message"
        '
        'tmrStatus
        '
        Me.tmrStatus.Enabled = True
        Me.tmrStatus.Interval = 1000
        '
        'tmrFSWatch
        '
        Me.tmrFSWatch.Interval = 2000
        '
        'FrmStart
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(593, 351)
        Me.Controls.Add(Me.ListView2)
        Me.Controls.Add(Me.Splitter2)
        Me.Controls.Add(Me.TreeView1)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.ListView1)
        Me.Controls.Add(Me.StatusBar1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Menu = Me.MainMenu1
        Me.Name = "FrmStart"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Faena Mail"
        CType(Me.StatusBarPanel1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.StatusBarPanel2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSWFile, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.FSWDir, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

#End Region

    <System.STAThread()> _
    Public Shared Sub Main()
        'System.Windows.Forms.Application.EnableVisualStyles()
        LibExcept.UnhandledExceptionManager.AddHandler()

        System.Windows.Forms.Application.Run(New FrmStart)
    End Sub 'Main

    Enum enulevel
        critical
        information
    End Enum
    Dim acctName As String
    Dim uniqueID As Integer
    Dim tools As New FaenaMailLibrary.tools
    Private _oldSubDirs As New Hashtable
    Private _fsWatch As String = ""
    Structure stuRunScript
        Dim boxType As String
        Dim boxName As String
        Dim boxLeft As Integer
        Dim boxTop As Integer
        Dim boxText As String
        Dim boxSettingIf As Integer
        Dim boxSettingIfDes As String
        Dim boxSettingAction As Integer
        Dim boxSettingActionDes As String
        Dim boxSettingMsg As String
        Dim boxSettingAttachment As String
        Dim boxSettingPath As String
    End Structure
    Private aryRunScript() As stuRunScript

    Structure stuMail
        Dim size As Long
        Dim msg As String
        Dim msgTop As String
        Dim folder As String
        Dim forward As String
        Dim read As Boolean
    End Structure
    Private mail As stuMail

    Private Sub WriteMsglog(ByVal msg As String, Optional ByVal level As enulevel = enulevel.information)
        If (level And (enulevel.critical Or MsgBoxStyle.Exclamation)) <> 0 Then
            StatusBar1.Panels(0).Text = "Error:"
        Else
            StatusBar1.Panels(0).Text = ""
        End If
        StatusBar1.Panels(0).Text = StatusBar1.Panels(0).Text & msg
        Dim SW As StreamWriter
        Dim FS As FileStream
        Dim logFile As String = reg.GetSetting("Faena Mail\", "LogFile", AppDomain.CurrentDomain.BaseDirectory & "FaenaMailDebug.log")
        FS = New FileStream(logFile, FileMode.Append)
        SW = New StreamWriter(FS)
        SW.Write(Now)
        SW.Write(vbTab)
        SW.WriteLine(msg)
        SW.Close()
        FS.Close()
    End Sub

    'Private Function CheckDelOpt(ByVal age As Long, ByVal acctID As Integer) As Boolean
    '    If CType(myReader.Item("LeaveOnSvr"), Boolean) = False Then 'Del right away
    '        Return True
    '    Else  'Leave On Server
    '        If CType(myReader.Item("RemoveFromSvr"), Boolean) = True Then
    '            If (age > CInt(myReader.Item("HowManyDays"))) Then
    '                Return True
    '            End If
    '        Else
    '            Return False
    '        End If
    '    End If
    'End Function

    Private Function OkToDelete(ByVal folder As String, ByVal uid As String, ByVal acctID As Integer) As Boolean
        Dim age As Long
        Dim SR As StreamReader
        Dim FS As FileStream
        FS = New FileStream(mainStorage & folder & "\" & uid & ".mail", FileMode.Open)
        SR = New StreamReader(FS)
        SR.ReadLine()
        Dim downloadDate As Date = CType(SR.ReadLine().Substring(9), Date)

        age = DateDiff(DateInterval.Day, downloadDate, Now.UtcNow)
        WriteMsglog("Age " & age.ToString)
        SR.Close()
        FS.Close()
        Return True
    End Function

    Private Function AddToFolder(ByVal folder As String, ByVal acctName As String, ByVal uid As String, ByVal mail As stuMail) As Boolean
        WriteMsglog("[AddToFolder]adding email...")
        Dim SW As StreamWriter
        Dim FS As FileStream
        folder = mainStorage & folder
        If Not Directory.Exists(folder) Then
            Directory.CreateDirectory(folder)
        End If
        Dim logFile As String
        logFile = folder & "\" & uid & ".mail"
        FS = New FileStream(logFile, FileMode.Create)
        SW = New StreamWriter(FS)
        SW.WriteLine("Account: " & acctName)
        SW.WriteLine("Download: " & Now.UtcNow.ToLongDateString & " " & Now.UtcNow.ToLongTimeString)
        SW.Write(mail.msg)
        SW.Close()
        FS.Close()
        Return True
    End Function

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem3.Click
        Dim frm As New FormAccounts
        frm.Owner = Me
        frm.ShowDialog()
    End Sub


    Private Function TopLine(ByVal source As String, ByVal strFilter As String) As String
        Dim lines() As String
        Dim line As String
        'mail.msgTop.Replace(vbCrLf & " ", " ")
        lines = Split(source, vbCrLf)
        For Each line In lines
            If line.ToUpper.StartsWith(strFilter.ToUpper & ":") Then
                If line.Length > strFilter.Length + 1 Then
                    Return line.Substring(line.IndexOf(":") + 2)
                Else
                    Return " "
                End If
            End If
        Next
    End Function

    'Private Function InDBField(ByVal str1 As String, ByVal strSQL As String) As Boolean
    '    Dim myCommand As New OleDbCommand(strSQL, conAccess)
    '    Dim myReader As OleDbDataReader = myCommand.ExecuteReader
    '    Dim i As Integer
    '    While myReader.Read
    '        For i = 1 To myReader.FieldCount
    '            If InStr(str1, myReader.Item(i).ToString) > 0 Then
    '                Return True
    '            End If
    '        Next
    '    End While
    'End Function
    Private Sub MenuItem9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem9.Click
        Dim cls As New FaenaMailLibrary.SedRecMail
        cls.CheckEmail()
    End Sub

    Private Sub MenuItem10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem10.Click
        Dim frm As New FormOptions
        frm.Owner = Me
        frm.ShowDialog(Me)
        If Not CheckServiceInstallation() Then
            Timer1.Interval = reg.GetSetting("Faena Mail\", "CheckEmailTimer", 30) * 1000 * 60
            Timer1.Enabled = CType(reg.GetSetting("Faena Mail\", "CheckEmail", 1), Boolean)
        Else
            WriteMsglog("FaenaMail service found")
        End If
        LoadDirList()
    End Sub

    Private Sub MenuItem11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem11.Click
        Dim frm As New FormScriptEditor
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub MenuItem6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem6.Click
        Shutdown()
    End Sub

    Private Sub MenuItem17_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem17.Click
        Dim cls As New FaenaMailLibrary.SedRecMail
        cls.SendOutbox()
        cls = Nothing
    End Sub

    Private Sub CheckForUpdate(Optional ByVal ifPrompt As Boolean = True)
        Dim result As Integer
        Try
            StatusBar1.Panels(0).Text = "Checking for updates..."
            Dim strVer As String
            Dim myTxtReader As XmlTextReader = New XmlTextReader("http://www.faena.com/faenamail/appinfo.xml")
            While myTxtReader.Read
                If myTxtReader.NodeType = XmlNodeType.Element Then
                    If myTxtReader.Name = "FaenaMail" Then
                        If myTxtReader.MoveToNextAttribute() Then
                            strVer = myTxtReader.GetAttribute("version")
                            If strVer <> "" Then
                                If strVer <> Application.ProductVersion Then
                                    Dim frm As New FrmUDConfirm
                                    frm.Owner = Me
                                    frm.Label1.Text = "There is a newer version of Faena Mail. The version you have is " & Application.ProductVersion & ", the new version is " & strVer & ". Do you want to update it now?"
                                    If frm.ShowDialog = DialogResult.Yes Then
                                        Dim strURL As String = myTxtReader.GetAttribute("url")
                                        If strURL = "" Then
                                            ShellExecute(Handle.ToInt32, "open", "http://www.faena.com", vbNullString, vbNullString, 1)
                                        Else
                                            ShellExecute(Handle.ToInt32, "open", strURL, vbNullString, vbNullString, 1)
                                        End If
                                        Application.Exit()
                                    End If
                                    frm.Dispose()
                                    result = 1 'new version
                                Else
                                    result = 2 'uptodate
                                End If
                            Else
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End While
            myTxtReader.Close()
            StatusBar1.Panels(0).Text = "Checking for updates complete."
        Catch ex As Exception
            WriteMsglog("[CheckForUpdate]" & ex.Message, enulevel.critical)
        End Try

        Select Case result
            Case 1 'new version
            Case 2 'same version
                If ifPrompt Then
                    MsgBox("You have the latest version of Faena mail. No update is available at this time.", MsgBoxStyle.Information)
                End If
            Case Else 'error
                Dim frm1 As New FrmUDConfirm
                frm1.Owner = Me
                frm1.Label1.Text = "There is a problem when checking for update. Do you want to browse the website to see if there is an update?"
                If frm1.ShowDialog = DialogResult.Yes Then
                    ShellExecute(Handle.ToInt32, "open", "http://www.faena.com/faenamail", vbNullString, vbNullString, 1)
                    'Application.Exit()
                End If
        End Select
    End Sub

    Private Sub LoadDirList()
        TreeView1.BeginUpdate()
        TreeView1.Nodes.Clear()
        TreeView1.Nodes.Add(mainStorage)
        AddFolders(TreeView1.Nodes(0))
        TreeView1.Nodes(0).Text = "FaenaMail"
        TreeView1.Nodes(0).ImageIndex = 0
        TreeView1.ExpandAll()
        TreeView1.EndUpdate()
    End Sub

    Private Function UseCount() As Integer
        Try
            Dim use As Integer = reg.GetSetting("Faena Mail\Settings", "ProductCnt", 0)
            use = use + 1
            reg.SaveSetting("Faena Mail\Settings", "ProductCnt", use)
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical)
        End Try
    End Function

    Private Sub FormStart_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ClearStatus()
        If vb.Command().IndexOf("min") >= 0 Then
            Me.Visible = False
        End If
        Dim frm As New frmProductID
        frm.Owner = Me
        If Not CheckSN1(reg.GetSetting("Faena Mail\Settings", "ProductID", "")) Then
            'Not registered
            gblnProductID = False
            If UseCount() > 20 Then
                Select Case frm.ShowDialog
                    Case DialogResult.Ignore
                    Case Else
                        If Not gblnProductID Then End
                End Select
            End If
        Else
            If Not CheckSN2(reg.GetSetting("Faena Mail\Settings", "ProductID", "")) Then
                'Not registered
                gblnProductID = False
                If UseCount() > 20 Then
                    Select Case frm.ShowDialog
                        Case DialogResult.Ignore
                        Case Else
                            If Not gblnProductID Then End
                    End Select
                End If
            Else
                gblnProductID = True
            End If
        End If
        frm.Dispose()
        If Not gblnProductID Then
            Me.Text = Me.Text & " (Trial)"
            MenuItem41.Enabled = False
            MenuItem42.Enabled = False
            MenuItem37.Enabled = False
            MenuItem36.Enabled = False
        Else
            MenuItem67.Visible = False
        End If
        mainStorage = reg.GetSetting("Faena Mail\", "MsgStore", AppDomain.CurrentDomain.BaseDirectory)
        WriteMsglog(Application.ProductName & " " & Application.ProductVersion & " started", enulevel.information)
        LoadDirList()
        WriteMsglog("OS: " & tools.OS_Version)
        If tools.IsWin9x Then
            FSWDir.EnableRaisingEvents = False
            FSWFile.EnableRaisingEvents = False
            _fsWatch = mainStorage
            tmrFSWatch.Start()
        Else
            FSWDir.Path = mainStorage
            FSWDir.EnableRaisingEvents = True
        End If

        NotifyIcon1.Visible = True
        'reg.SaveSetting("Faena Mail\", "NewMailNotify", "False")

        If Not CheckServiceInstallation() Then
            Timer1.Interval = reg.GetSetting("Faena Mail\", "CheckEmailTimer", 30) * 1000 * 60
            Timer1.Enabled = CType(reg.GetSetting("Faena Mail\", "CheckEmail", 1), Boolean)
        Else
            WriteMsglog("FaenaMail service found")
        End If
        displayClock = CBool(reg.GetSetting("Faena Mail\DisplayTime", "DisplayClock", 0))
        MenuItem29.Checked = displayClock
        Timer2.Enabled = True
        Timer3.Enabled = True
    End Sub

    Private Sub AddFolders(ByVal nod As TreeNode)
        Try
            Dim strPath As String = nod.FullPath
            Dim strDir As String
            With nod
                For Each strDir In Directory.GetDirectories(strPath)
                    strDir = Path.GetFileName(strDir)
                    Dim newNod As TreeNode = nod.Nodes.Add(strDir)
                    newNod.Tag = strPath.Replace(mainStorage, "") & strDir & "\"
                    AddFolders(newNod)
                    'nod.Nodes(0).Expand()
                Next
            End With
        Catch ex As Exception
            WriteMsglog("[AddFolders]" & ex.Message, enulevel.critical)
        End Try
    End Sub

    Private Sub MenuItem21_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New FrmCompose("", "NewMail")
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub MenuItem22_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim frm As New FrmCompose("", "MailList")
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub TreeView1_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterSelect
        If Not e.Node Is Nothing Then
            mObjectCnt = 0
            mUnreadCnt = 0
            ListView2.BeginUpdate()
            ListView2.Items.Clear()
            Dim strFile As String
            If e.Node.Text <> "FaenaMail" Then
                Dim strNode As String = e.Node.Tag.ToString
                currentPath = strNode
                For Each strFile In Directory.GetFiles(mainStorage & strNode, "*.mail")
                    AddListView2(strFile)
                Next
                If tools.IsWin9x Then
                    _fsWatch = mainStorage & strNode
                Else
                    FSWFile.Path = mainStorage & strNode
                End If
            End If
            StatusBar1.Panels(0).Text = mObjectCnt & " message(s), " & mUnreadCnt & " unread"
            ListView2.EndUpdate()
        End If
    End Sub

    Private Sub AddListView2(ByVal strFile As String)
        Dim s As String
        Dim strMsg As String
        Dim topLines As String
        Dim FS As FileStream
        Dim SR As StreamReader
        FS = New FileStream(strFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        SR = New StreamReader(FS)
        While SR.Peek > 0
            strMsg = SR.ReadLine
            If strMsg = "" Then Exit While
            topLines = topLines & strMsg & vbNewLine
        End While
        Dim itmx As ListViewItem
        mObjectCnt = mObjectCnt + 1
        itmx = ListView2.Items.Add(TopLine(topLines, "From"))
        itmx.Tag = Path.GetFileName(strFile)
        itmx.SubItems.Add("")
        itmx.SubItems.Add("")
        itmx.SubItems(1).Text = TopLine(topLines, "Subject")

        Dim status As String = TopLine(topLines, "Status")
        Dim myfont As Font
        Select Case status
            Case "R"
                itmx.StateImageIndex = 4
                myfont = New Font(ListView2.Font, FontStyle.Regular)
            Case "U", "N", ""
                myfont = New Font(ListView2.Font, FontStyle.Bold)
                itmx.StateImageIndex = 3
                mUnreadCnt = mUnreadCnt + 1
        End Select
        itmx.Font = myfont

        s = TopLine(topLines, "Date")
        If Not s Is Nothing Then
            If s.IndexOf(",") > 0 Then s = s.Substring(s.IndexOf(",") + 1)
            If s.IndexOf("-") > 0 Then s = s.Substring(0, s.IndexOf("-"))
            If s.IndexOf("+") > 0 Then s = s.Substring(0, s.IndexOf("+"))
            If IsDate(s) Then
                itmx.SubItems(2).Text = Trim(s)
            Else
                itmx.SubItems(2).Text = TopLine(topLines, "Date")
            End If
        End If
        FS.Close()
    End Sub

    Private Sub FileChanged(ByVal fileName As String)
        Try
            Dim cntNew As Integer = 0
            Dim i As Integer
            If Path.GetExtension(fileName) = ".mail" Then
                For i = ListView2.Items.Count - 1 To 0 Step -1
                    If ListView2.Items(i).Tag.ToString = Path.GetFileName(fileName) Then
                        Dim strMsg As String
                        Dim topLines As String
                        Dim FS As FileStream
                        Dim SR As StreamReader
                        FS = New FileStream(fileName, FileMode.Open, FileAccess.Read)
                        SR = New StreamReader(FS)
                        While SR.Peek > 0
                            strMsg = SR.ReadLine
                            If strMsg = "" Then Exit While
                            topLines = topLines & strMsg & vbNewLine
                        End While
                        With ListView2.Items(i)
                            .Text = TopLine(topLines, "From")
                            .SubItems(1).Text = TopLine(topLines, "Subject")
                            .SubItems(2).Text = MailTimeToVBTime2(TopLine(topLines, "Date")).ToString
                            Dim status As String = TopLine(topLines, "Status")
                            Dim myfont As Font
                            Select Case status
                                Case "R"
                                    .StateImageIndex = 4
                                    myfont = New Font(ListView2.Font, FontStyle.Regular)
                                Case "N", ""
                                    myfont = New Font(ListView2.Font, FontStyle.Bold)
                                    .StateImageIndex = 3
                                    cntNew = cntNew + 1
                            End Select
                            .Font = myfont
                        End With
                        SR.Close()
                        FS.Close()
                    End If
                Next
                'CountNewMails()
            End If
        Catch
        End Try
    End Sub

    Private Sub FSW_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles FSWFile.Changed
        FileChanged(e.FullPath)
    End Sub

    Private Sub FSW_Created(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles FSWFile.Created
        If Path.GetExtension(e.Name) = ".mail" Then
            AddListView2(e.FullPath)
            StatusBar1.Panels(0).Text = "New message added"
        End If
    End Sub
    Private Sub FileDeleted(ByVal sfile As String)
        Dim i As Integer
        Dim cnt As Integer
        If Path.GetExtension(sfile) = ".mail" Then
            For i = ListView2.Items.Count - 1 To 0 Step -1
                If ListView2.Items(i).Tag.ToString = Path.GetFileName(sfile) Then
                    ListView2.Items(i).Remove()
                    cnt = cnt + 1
                End If
            Next
        End If
        StatusBar1.Panels(0).Text = cnt & " message(s) removed"
    End Sub
    Private Sub FSW_Deleted(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles FSWFile.Deleted
        FileDeleted(e.Name)
    End Sub

    Private Sub FSWFile_Renamed(ByVal sender As System.Object, ByVal e As System.IO.RenamedEventArgs) Handles FSWFile.Renamed
        Dim i As Integer
        If Path.GetExtension(e.OldName.ToString) = ".mail" Then
            For i = ListView2.Items.Count - 1 To 0 Step -1
                If ListView2.Items(i).Tag.ToString = Path.GetFileName(e.OldName.ToString) Then
                    ListView2.Items(i).Tag = Path.GetFileName(e.Name.ToString)
                End If
            Next
        End If
    End Sub

    Private Sub OpenMessage()
        Dim strPath As String = mainStorage
        strPath = strPath & currentPath
        strPath = strPath & ListView2.SelectedItems(0).Tag.ToString
        Dim frm As New FrmViewEmail(strPath)
        frm.Owner = Me
        frm.Show()
        reg.SaveSetting("Faena Mail\", "NewMailNotify", "False")
    End Sub

    Private Sub ListView2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListView2.DoubleClick
        OpenMessage()
    End Sub

    Private Sub MenuItem5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem5.Click
        Dim frm As New frmAbout
        frm.ShowDialog(Me)
    End Sub

    Private Sub MenuItem18_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem18.Click
        MenuItem9.PerformClick()
        MenuItem17.PerformClick()
    End Sub

    Private Sub MenuItem13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem13.Click
        Try
            System.Diagnostics.Process.Start("http://www.faena.com")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
        'ShellExecute(Handle.ToInt32, "open", "http://www.faena.com", vbNullString, vbNullString, 1)
    End Sub

    Private Sub FormStart_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If bCloseClicked Then 'user clicked close box
            Me.Hide()
            e.Cancel = True
            bCloseClicked = False
        Else 'user didn't click close box
        End If
    End Sub

    Protected Sub Shutdown()
        ' It's a good idea to make the system tray icon invisible before ending
        ' the application, otherwise, it can linger in the tray when the application
        ' is no longer running.
        NotifyIcon1.Visible = False
        Application.Exit()
    End Sub

    Private Sub NotifyIcon1_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon1.DoubleClick
        MenuItem27.PerformClick()
    End Sub

    Private Sub MenuItem27_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem27.Click
        Try
            Me.WindowState = FormWindowState.Normal
            Me.Show()
            Me.BringToFront()
        Catch
        End Try
    End Sub

    Private Sub MenuItem28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem28.Click
        Shutdown()
    End Sub

    Private Function GetTime() As DateTime
        Dim UtcNow As Date
        Dim targetZone As Integer
        UtcNow = Date.UtcNow
        targetZone = reg.GetSetting("Faena Mail\DisplayTime", "TimeZone", 0)
        If (targetZone = 0) Then
            Return Now
        Else
            Try
                Dim winTimeZones As Win32.Win32TimeZone()
                winTimeZones = Win32.TimeZones.GetTimeZones
                Array.Sort(winTimeZones, New Win32TimeZoneComparer)
                Return winTimeZones(targetZone - 1).ToLocalTime(UtcNow)
            Catch
                Return Now
            End Try
        End If
    End Function

    Private Function GetAbbrTime() As String
        Dim UtcNow As Date
        Dim targetZone As Integer
        UtcNow = Date.UtcNow
        targetZone = reg.GetSetting("Faena Mail\DisplayTime", "TimeZone", 0)
        If (targetZone = 0) Then
            Return ""
        Else
            Try
                Dim winTimeZones As Win32.Win32TimeZone()
                winTimeZones = Win32.TimeZones.GetTimeZones
                Array.Sort(winTimeZones, New Win32TimeZoneComparer)
                Return winTimeZones(targetZone - 1).GetAbbreviationUniversalTime(UtcNow)
            Catch
                Return ""
            End Try
        End If
    End Function
    Private Sub ClearStatus()
        Dim regkey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\Faena Mail\", True)
        If Not regkey Is Nothing Then
            regkey.DeleteSubKey("DisplayStatus", False)
        End If
        If Not regkey Is Nothing Then regkey.Close()
    End Sub
    Private Sub DisplayStatus()
        Dim xitm As ListViewItem
        For i As Integer = 0 To ListView1.Items.Count - 1
            ListView1.Items(i).Tag = ""
        Next
        Dim strName As String
        Dim regkey As RegistryKey = Registry.LocalMachine.OpenSubKey("Software\Faena\Faena Mail\DisplayStatus", False)
        If Not regkey Is Nothing Then
            For Each strName In regkey.GetValueNames
                Dim found As Boolean = False
                For Each xitm In ListView1.Items
                    If xitm.Text = strName Then
                        xitm.Tag = "touched"
                        found = True
                        Exit For
                    End If
                Next
                If found Then
                    xitm.SubItems(1).Text = reg.GetSetting("Faena Mail\DisplayStatus", strName, "")
                Else
                    xitm = ListView1.Items.Add(strName)
                    xitm.Tag = "touched"
                    xitm.SubItems.Add(reg.GetSetting("Faena Mail\DisplayStatus", strName, ""))
                    xitm.EnsureVisible()
                End If
                If xitm.SubItems(1).Text.StartsWith("Complete") Then
                    xitm.StateImageIndex = 0
                ElseIf xitm.SubItems(1).Text.StartsWith("Error") Then
                    xitm.StateImageIndex = 1
                Else
                    xitm.StateImageIndex = -1
                End If
            Next
        End If
        For i As Integer = ListView1.Items.Count - 1 To 0 Step -1
            If ListView1.Items(i).Tag.ToString = "" Then ListView1.Items(i).Remove()
        Next
        StatusBar1.Panels(1).Text = reg.GetSetting("Faena Mail\", "CurrentAction", "")
    End Sub

    Private Sub ShowNotify()
        If reg.GetSetting("Faena Mail\", "NotifyShowed", 0) = 0 Then
            If cntNewMailNotify < 1 Then
                Dim frm As New frmNewMailNotify(Me)
                frm.Owner = Me
                frm.Show()
                reg.SaveSetting("Faena Mail\", "NotifyShowed", 1)
            End If
        End If
    End Sub

    Private Sub Timer2_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer2.Tick
        Timer2.Enabled = False
        If displayClock Then
            Dim bmp As Bitmap
            bmp = New Bitmap(16, 16)
            Dim i As Integer
            Dim G As Graphics
            G = Graphics.FromImage(bmp)
            G.SmoothingMode = Drawing.Drawing2D.SmoothingMode.AntiAlias
            Dim R1 As Rectangle
            R1 = New Rectangle(0, 0, 15, 15)
            If reg.GetSetting("Faena Mail\", "NewMailNotify", "") = "True" Then
                NotifyIcon2.Visible = True
                If CType(reg.GetSetting("Faena Mail\", "ShowNotify", 1), Boolean) Then
                    ShowNotify()
                Else
                    Beep()
                End If
                G.FillEllipse(Brushes.AntiqueWhite, R1)
            Else
                NotifyIcon2.Visible = False
                G.FillEllipse(Brushes.WhiteSmoke, R1)
            End If
            G.DrawEllipse(New Pen(Color.Black, 1), R1)
            Dim AnHours As Single, AnMinutes As Single, AnSeconds As Single
            Dim TrueHours As Single
            AnHours = GetTime.Hour
            AnMinutes = GetTime.Minute
            TrueHours = AnHours + AnMinutes / 60
            Dim AHoursLength As Single = 4.5
            Dim AMinutesLength As Single = 5.3
            Dim MidX As Single = 7.5
            Dim MidY As Single = 7.5
            Dim AHoursX2, AHoursY2, AMinutesX2, AMinutesY2 As Integer
            AHoursX2 = CInt(AHoursLength * Math.Cos(Math.PI / 180 * (30 * TrueHours - 90)) + MidX)
            AHoursY2 = CInt(AHoursLength * Math.Sin(Math.PI / 180 * (30 * TrueHours - 90)) + MidY)
            AMinutesX2 = CInt(AMinutesLength * Math.Cos(Math.PI / 180 * (6 * AnMinutes - 90)) + MidX)
            AMinutesY2 = CInt(AMinutesLength * Math.Sin(Math.PI / 180 * (6 * AnMinutes - 90)) + MidY)
            G.DrawLine(New Pen(Color.Black, 1.8), 7.5, 7.5, AHoursX2, AHoursY2)
            G.DrawLine(New Pen(Color.Black, 1), 7.5, 7.5, AMinutesX2, AMinutesY2)
            NotifyIcon1.Text = "Faena Mail  " & GetTime.ToLongDateString & " " & GetTime.ToShortTimeString & " " & GetAbbrTime()
            Try
                NotifyIcon1.Icon = System.Drawing.Icon.FromHandle(CType(bmp, System.Drawing.Bitmap).GetHicon())
            Catch
            End Try
            G.Dispose()
            bmp.Dispose()
        End If
        Timer2.Enabled = True
    End Sub

    Private Sub FormStart_Closed(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Closed
        NotifyIcon1.Visible = False
    End Sub

    Private Sub MenuItem29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem29.Click
        Dim frm As New FrmDspTime
        frm.ShowDialog(Me)
        frm.Dispose()
        displayClock = CBool(reg.GetSetting("Faena Mail\DisplayTime", "DisplayClock", 0))
        MenuItem29.Checked = displayClock
        If Not displayClock Then
            NotifyIcon1.Text = "Faena Mail"
            NotifyIcon1.Icon = Me.Icon
        End If
    End Sub

    Private Sub MenuItem31_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem31.Click
        Me.Hide()
    End Sub

    'Private Sub DirectoryChanged()
    '    TreeView1.BeginUpdate()
    '    TreeView1.Nodes.Clear()
    '    TreeView1.Nodes.Add(mainStorage)
    '    AddFolders(TreeView1.Nodes(0))
    '    TreeView1.Nodes(0).Text = "FaenaMail"
    '    TreeView1.Nodes(0).Tag = ""
    '    TreeView1.Nodes(0).SelectedImageIndex = 0
    '    TreeView1.Nodes(0).ImageIndex = 0
    '    TreeView1.ExpandAll()
    '    TreeView1.EndUpdate()
    'End Sub

    Private Sub FSWDir_Changed(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FSWDir.Created, FSWDir.Changed, FSWDir.Deleted
        LoadDirList()
    End Sub

    Private Sub FSWDir_Renamed(ByVal sender As Object, ByVal e As System.IO.RenamedEventArgs) Handles FSWDir.Renamed
        LoadDirList()
    End Sub

    Private Sub MenuItem32_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem32.Click
        Dim frm As New frmAbout
        frm.ShowDialog(Me)
    End Sub
    Private Function CheckServiceInstallation() As Boolean
        ' Verify to see if the service is installed.
        Try
            Dim installedServices() As ServiceController
            Dim tmpService As ServiceController
            Dim i As Integer = 0

            installedServices = ServiceController.GetServices()

            For Each tmpService In installedServices
                If tmpService.DisplayName = "FaenaMailService" Then
                    CheckServiceInstallation = True
                    Exit For
                End If
            Next tmpService
        Catch ex As Exception
            Debug.WriteLine(ex.ToString)
        End Try
    End Function

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Timer1.Enabled = False
        MenuItem18.PerformClick()
        Timer1.Enabled = True
    End Sub

    Private Sub Timer3_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer3.Tick
        Timer3.Enabled = False
        If vb.Command().IndexOf("min") >= 0 Then
            Me.Hide()
        Else
            If reg.GetSetting("Faena Mail\", "DonotCheckUpdate", 0) = 0 Then
                CheckForUpdate(False)
            End If
        End If
        If reg.GetSetting("Faena Mail\", "CheckAtStartup", 1) <> 0 Then
            MenuItem18.PerformClick()
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As System.Windows.Forms.Message)
        If m.Msg = WM_SYSCOMMAND AndAlso m.WParam.ToInt32 = SC_CLOSE Then
            bCloseClicked = True
        End If
        MyBase.WndProc(m)
    End Sub

    Private Sub MenuItem33_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem33.Click
        Dim frm As New FrmView(reg.GetSetting("Faena Mail\", "LogFile", AppDomain.CurrentDomain.BaseDirectory & "FaenaMailDebug.log"))
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub MenuItem47_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem47.Click
        OpenMessage()
    End Sub

    Private Sub DelMsg()
        Dim msg As ListViewItem
        Dim strPath As String
        Dim tarPath As String
        Dim ok As Boolean = False
        If currentPath = "Deleted Items\" Then
            If MsgBox("Are you sure you want to permanently delete these message(s)?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                ok = True
            End If
        Else
            ok = True
        End If
        If ok Then
            Try
                reg.SaveSetting("Faena Mail\", "NewMailNotify", "False")
                If Not Directory.Exists(mainStorage & "Deleted Items\") Then
                    Directory.CreateDirectory(mainStorage & "Deleted Items\")
                End If
                For Each msg In ListView2.SelectedItems
                    strPath = mainStorage & currentPath & msg.Tag.ToString
                    tarPath = mainStorage & "Deleted Items\" & msg.Tag.ToString
                    If currentPath = "Deleted Items\" Then
                        Kill(strPath)
                    Else
                        Try
                            File.Move(strPath, tarPath)
                        Catch
                            Kill(strPath)
                        End Try
                    End If
                Next
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Delete Message Error")
            End Try
        End If
    End Sub

    Private Sub MenuItem58_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem58.Click
        DelMsg()
    End Sub

    Private Sub MenuItem49_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem49.Click
        Dim f As New FormWait
        f.Owner = Me
        f.Show()
        Application.DoEvents()
        CheckForUpdate(True)
        f.Dispose()
    End Sub

    Private Sub MenuItem40_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem40.Click
        DelMsg()
    End Sub

    Private Sub MenuItem65_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem65.Click
        Try
            If Not MyNode Is Nothing Then
                Dim strFullPath As String = MyNode.FullPath.Substring(MyNode.FullPath.IndexOf("\") + 1)
                Dim strName As String
                If strFullPath.IndexOf("\") >= 0 Then
                    strName = strFullPath.Substring(0, strFullPath.IndexOf("\"))
                Else
                    strName = strFullPath
                End If
                If strName = "Deleted Items" Then
                    If MsgBox("Are you sure you want to permanently delete the '" & strFullPath & "' folder?" & vbNewLine & _
                        vbNewLine & _
                        "This action cannot be undone.", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        Try
                            Directory.Delete(mainStorage & strFullPath, True)
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Delete Folder Error")
                        End Try
                    End If
                Else
                    If MsgBox("Are you sure you want to delete the '" & strFullPath & "' folder and move it to the 'Deleted Items' folder?", MsgBoxStyle.Exclamation Or MsgBoxStyle.YesNo) = MsgBoxResult.Yes Then
                        strName = mainStorage & "Deleted Items\" & MyNode.Text
                        Dim i As Integer
                        While Directory.Exists(strName)
                            i = i + 1
                            strName = mainStorage & "Deleted Items\" & MyNode.Text & "(" & i & ")"
                        End While
                        Try
                            Directory.Move(mainStorage & strFullPath, strName)
                        Catch ex As Exception
                            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Delete Folder Error")
                        End Try
                    End If
                End If
            End If
        Catch
            MyNode = Nothing
        End Try
    End Sub

    Private Sub TreeView1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles TreeView1.MouseDown
        If (e.Button = MouseButtons.Right) Then
            MyNode = TreeView1.GetNodeAt(e.X, e.Y)
        Else
            MyNode = Nothing
        End If
    End Sub

    Private Sub MenuItem63_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem63.Click
        Dim strName As String = InputBox("Folder name:", "New Folder")
        Dim strFullPath As String
        If strName <> "" Then
            If MyNode.FullPath.IndexOf("\") >= 0 Then
                strFullPath = MyNode.FullPath.Substring(MyNode.FullPath.IndexOf("\") + 1)
            Else
                strFullPath = ""
            End If
            Try
                Directory.CreateDirectory(mainStorage & strFullPath & "\" & strName)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "New Folder Error")
            End Try
        End If
    End Sub

    Private Sub MenuItem64_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem64.Click
        Dim strNewName As String = InputBox("Folder name:", "Rename Folder", MyNode.Text)
        Dim strFullPath As String
        If strNewName <> "" Then
            If MyNode.FullPath.IndexOf("\") >= 0 Then
                strFullPath = MyNode.FullPath.Substring(MyNode.FullPath.IndexOf("\") + 1)
                If strFullPath.LastIndexOf("\") >= 0 Then
                    strFullPath = strFullPath.Substring(0, strFullPath.LastIndexOf("\"))
                End If
            Else
                strFullPath = ""
            End If
            Try
                Directory.Move(mainStorage & strFullPath & "\" & MyNode.Text, mainStorage & strFullPath & "\" & strNewName)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Rename Folder Error")
            End Try
        End If
    End Sub

    Private Sub MenuItem46_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem46.Click
        Dim itemx As ListViewItem
        For Each itemx In ListView2.SelectedItems
            MarkMailStatus(mainStorage & currentPath & itemx.Tag.ToString, "R")
        Next
    End Sub

    Private Sub MenuItem45_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem45.Click
        Dim itemx As ListViewItem
        For Each itemx In ListView2.SelectedItems
            MarkMailStatus(mainStorage & currentPath & itemx.Tag.ToString, "N")
        Next
    End Sub

    Private Sub MenuItem60_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem60.Click
        MenuItem46.PerformClick()
    End Sub

    Private Sub MenuItem61_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem61.Click
        MenuItem45.PerformClick()
    End Sub

    Private Sub MenuItem22_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem22.Click
        Dim strPath As String = mainStorage & currentPath
        strPath = strPath & ListView2.SelectedItems(0).Tag.ToString
        Dim frm As New FrmViewEmail(strPath)
        frm.Owner = Me
        frm.Show()
        frm.MenuItem12.PerformClick()
    End Sub

    Private Sub MenuItem50_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem50.Click
        MenuItem22.PerformClick()
    End Sub

    Private Sub MenuItem36_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem36.Click
        Dim strPath As String = mainStorage & currentPath
        strPath = strPath & ListView2.SelectedItems(0).Tag.ToString
        Dim frm As New FrmViewEmail(strPath)
        frm.Owner = Me
        frm.Show()
        frm.MenuItem13.PerformClick()
    End Sub

    Private Sub MenuItem51_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem51.Click
        MenuItem36.PerformClick()
    End Sub

    Private Sub MenuItem37_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem37.Click
        Dim strPath As String = mainStorage & currentPath
        strPath = strPath & ListView2.SelectedItems(0).Tag.ToString
        Dim frm As New FrmViewEmail(strPath)
        frm.Owner = Me
        frm.Show()
        frm.MenuItem14.PerformClick()
    End Sub

    Private Sub MenuItem52_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem52.Click
        MenuItem37.PerformClick()
    End Sub

    Private Sub MenuItem35_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem35.Click
        Dim frm As New FrmCompose("", "")
        frm.Owner = Me
        frm.Show()
    End Sub

    Private Sub MenuItem43_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem43.Click
        Dim frm As New FrmCompose("", "MailList")
        frm.Owner = Me
        frm.Show()

    End Sub

    Private Sub MenuItem62_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem62.Click
        MenuItem35.PerformClick()
    End Sub

    Private Sub MenuItem66_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem66.Click
        MenuItem43.PerformClick()
    End Sub

    Private Sub MenuItem14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem14.Click
        Try
            System.Diagnostics.Process.Start("http://www.faena.com/user_support.htm")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub MenuItem67_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem67.Click
        Dim frm As New frmProductID
        frm.Owner = Me
        frm.ShowDialog()
    End Sub


    Private Sub ListView2_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles ListView2.ColumnClick
        If ListView2.View <> View.Details Then
            Debug.WriteLine("Info: The ColumnClick sort code only works if .View=View.Details.")
            Return
        End If

        Static oSorter As ListViewItemSorterClass
        If oSorter Is Nothing Then
            oSorter = New ListViewItemSorterClass
            ListView2.ListViewItemSorter = Nothing   ' Make sure no other ItemSorter (for example set by the listview Sorting property) disturbes.
        End If
        'oSorter.CultureInfoName = "iv"		 ' Override locale system setting.
        oSorter.Column = e.Column
        oSorter.ColumnType = oSorter.DetectColumnType(ListView2, e.Column)  ' Info: Optional - but preferred in most situations. Remove it to always use ColumnType=ListViewColumnType.Text. Replace it with other logic if you have a "special case".
        'oSorter.CompareMethod = Microsoft.VisualBasic.CompareMethod.Binary		 ' Specify case sensitivity (used if ColumnType=ListViewColumnType.Text).
        'oSorter.CompareFromLeftToDiff = False		 ' Compare as the dotNET framework does it (used if ColumnType=ListViewColumnType.Text).
        oSorter.InvertSortOrder()  ' Invert the SortOrder.
        If ListView2.ListViewItemSorter Is Nothing Then
            ListView2.ListViewItemSorter = oSorter   ' Info: Setting this property to a ItemSorter triggers a sort (the first time). Reassigning the property to the same ItemSorter does NOT trigger a sort.
        Else
            ListView2.Sort()   ' Do a sort.
        End If
    End Sub

    Private Sub NotifyIcon2_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles NotifyIcon2.DoubleClick
        MenuItem27.PerformClick()
    End Sub

    Private Sub MenuItem24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem24.Click
        MenuItem35.PerformClick()
    End Sub

    Private Sub MenuItem16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem16.Click

    End Sub

    Private Sub MenuItem68_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem68.Click
        Dim toolFileName As String = Application.StartupPath & "\txt2db.exe"
        If File.Exists(toolFileName) Then
            Process.Start(toolFileName)
        Else
            MessageBox.Show(toolFileName & " not found!", Me.Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
        End If
    End Sub

    Private Sub tmrStatus_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrStatus.Tick
        DisplayStatus()
    End Sub

    Private Sub FSDirWatch()
        Dim ifChanged As Boolean = False
        Dim value As Integer
        For Each value In _oldSubDirs
            value = 0
        Next
        For Each subdir As String In Directory.GetDirectories(mainStorage)
            If _oldSubDirs.Contains(subdir) Then
                _oldSubDirs(subdir) = 1
            Else
                ifChanged = True
                Exit For
            End If
        Next
        For Each value In _oldSubDirs
            If value <> 1 Then
                ifChanged = True
                Exit For
            End If
        Next
        If ifChanged Then
            LoadDirList()
        End If
    End Sub
    Private _oldFiles As New Hashtable
    Private Sub FSFileWatch()
        Dim sFile As String
        For Each sFile In Directory.GetFiles(_fsWatch, "*.mail")
            If _oldFiles.Contains(sFile) Then
                If _oldFiles(sFile).ToString = File.GetLastWriteTimeUtc(sFile).ToString Then
                    'no change
                Else
                    FileChanged(sFile)
                End If
            Else
                'new file
                _oldFiles.Add(sFile, File.GetLastWriteTimeUtc(sFile).ToString)
                AddListView2(sFile)
                StatusBar1.Panels(0).Text = "New message added"
            End If
        Next
        For Each sFile In _oldFiles.Keys
            If Not File.Exists(sFile) Then
                'file deleted
                FileDeleted(sFile)
            End If
        Next
    End Sub
    Private Sub tmrFSWatch_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrFSWatch.Tick
        FSDirWatch()
        FSFileWatch()
    End Sub

    Private Sub FrmStart_KeyPress(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyPressEventArgs) Handles MyBase.KeyPress

    End Sub
End Class
