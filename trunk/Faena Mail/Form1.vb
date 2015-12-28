'Copyright (C) 2003 Faena Technologies
'All rights reserved.
'THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER 
'EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE IMPLIED WARRANTIES OF 
'MERCHANTIBILITY AND/OR FITNESS FOR A PARTICULAR PURPOSE.

'Requires the Trial or Release version of Visual Studio .NET Professional (or greater).
Option Strict On
Imports System.Data
Imports System.Text
Imports System.Data.OleDb
Imports LumiSoft.Net.POP3.Client
Public Class Form1
    Inherits System.Windows.Forms.Form

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
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.ListBox1 = New System.Windows.Forms.ListBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ListBox1
        '
        Me.ListBox1.Anchor = (((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
                    Or System.Windows.Forms.AnchorStyles.Left) _
                    Or System.Windows.Forms.AnchorStyles.Right)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(376, 199)
        Me.ListBox1.TabIndex = 1
        '
        'Button2
        '
        Me.Button2.Anchor = (System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left)
        Me.Button2.Location = New System.Drawing.Point(232, 232)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(96, 27)
        Me.Button2.TabIndex = 2
        Me.Button2.Text = "Check Emails"
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(104, 232)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 24)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Properties"
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(376, 269)
        Me.Controls.AddRange(New System.Windows.Forms.Control() {Me.Button1, Me.Button2, Me.ListBox1})
        Me.Name = "Form1"
        Me.Text = "Faena LightWind"
        Me.ResumeLayout(False)

    End Sub

#End Region
    Dim conAccess As New OleDb.OleDbConnection()
    Dim cmdEmail As New OleDbCommand()
    Dim myCommand As OleDbDataAdapter
    Dim myDS As New DataSet()

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim c As New POP3_Client()
        Try
            WriteMsgLog("Connecting")
            c.Connect("192.168.100.52", 111)
            c.Authenticate("vincentm", "20021111")
            Dim mInf As POP3_MessagesInfo = c.GetMessagesInfo()
            WriteMsgLog("Count:" + mInf.Count.ToString)
            WriteMsgLog("TotalSize:" + mInf.TotalSize.ToString)
            Dim hTable As Hashtable = c.GetUidlList()
            Dim uid As String
            Dim msg As String
            conAccess.ConnectionString = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = C:\MyProjects\FaenaLightWind\LightWindDB.mdb;"
            conAccess.Open()
            For Each uid In hTable.Keys
                WriteMsgLog(uid & " : " & hTable.Item(uid).ToString)
                msg = Encoding.ASCII.GetString(c.GetMessage(CInt(hTable.Item(uid))))
                WriteMsgLog(msg)
                AddToDB(uid, msg)
                If OkToDelete(uid) Then
                    WriteMsgLog("Deleteting from server:" & uid)
                    c.DeleteMessage(CInt(hTable.Item(uid)))
                End If
            Next
            c.Disconnect()
            conAccess.Close()
        Catch Ex As Exception
            WriteMsgLog(Ex.Message)
        End Try
    End Sub

    Private Sub WriteMsgLog(ByVal msg As String)
        ListBox1.Items.Add(msg)
        ListBox1.Refresh()
    End Sub

    Private Function CheckDelOpt(ByVal age As Long) As Boolean
        If StrComp(GetSetting(Application.ProductName, "CheckEmail", "LeaveOnServer", "False"), "True", CompareMethod.Text) <> 0 Then 'Del right away
            Return True
        Else  'Leave On Server
            If StrComp(GetSetting(Application.ProductName, "CheckEmail", "RemoveFromServer", "False"), "True", CompareMethod.Text) = 0 Then
                If (age > CInt(GetSetting(Application.ProductName, "CheckEmail", "AfterDays", "5"))) Then
                    Return True
                End If
            Else
                Return False
            End If
        End If
    End Function

    Private Function OkToDelete(ByVal uid As String) As Boolean
        cmdEmail.Connection = conAccess
        cmdEmail.CommandText = "select downloaddate from emails where uid='" & uid & "'"

        Dim downloadDate As Date
        If Not IsDBNull(cmdEmail.ExecuteScalar) Then
            downloadDate = CType(cmdEmail.ExecuteScalar, Date)
        End If
        Dim age As Long
        age = DateDiff(DateInterval.Day, downloadDate, Now)
        WriteMsgLog("Age " & age.ToString)
        Return CheckDelOpt(age)
    End Function


    Private Sub AddToDB(ByVal uid As String, ByVal msg As String)
        Dim strSQL As String = "INSERT INTO emails (unread,folder,uid,source,downloaddate) VALUES (Yes,'Inbox','" & uid & "','" & msg & "',now())"
        myCommand = New OleDbDataAdapter(strSQL, conAccess)
        Try
            myCommand.Fill(myDS, "Catagory")
        Catch Ex As Exception
            WriteMsgLog(Ex.Message)
        End Try

    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim frm As New FormProperties()
        frm.ShowDialog()
    End Sub
End Class
