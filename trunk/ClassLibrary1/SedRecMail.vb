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
Imports System.Xml
Imports Microsoft.Win32
Imports System.Text
Imports System.Data.Odbc
Imports LumiSoft.Net.POP3.Client
Imports System.Net
Imports LumiSoft.Net
Imports LumiSoft.Net.SMTP.Client
Imports LumiSoft.Net.Dns
Imports LumiSoft.Net.Mime

Imports System.IO
Public Class SedRecMail
    'Dim xDoc As New XmlDocument
    Dim xNodelist As XmlNodeList
    Dim xNode As XmlNode
    Dim acctName As String
    Dim c As New POP3_Client
    Dim uniqueID As Integer
    Dim myReader As OdbcDataReader

    Enum enulevel
        critical
        information
    End Enum
    Dim aryVar As New Hashtable
    Structure stuMail
        Dim size As Long
        Dim msg As String
        Dim msgTop As String
        Dim folder As String
        Dim forward As String
        Dim status As String
    End Structure
    Private mail As stuMail
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
        Dim boxSettingBcc As String
        Dim boxSettingAttachment As String
        Dim boxSettingPath As String
    End Structure
    Private aryRunScript() As stuRunScript

    Private Function CheckProductID() As Boolean
        If Not CheckSN1(reg.GetSetting("Faena Mail\Settings", "ProductID", "")) Then
            Return False
        Else
            If Not CheckSN2(reg.GetSetting("Faena Mail\Settings", "ProductID", "")) Then
                Return False
            Else
                Return True
            End If
        End If
    End Function

    Private Sub SendOutboxThread()
        Try
            reg.SaveSetting("Faena Mail", "CurrentAction", "Sending mail...")
            Dim xDoc As New XmlDocument
            If File.Exists(mainStorage & "accounts.xml") Then
                xDoc.Load(mainStorage & "accounts.xml")
                Dim xNode As XmlNode
                Dim outMail As String
                Dim FS As FileStream
                Dim SR As StreamReader
                Dim actName As String
                Dim state As String
                Dim mailType As String
                Dim fromAdr As String
                Dim toMask As String
                Dim scheduleAt As DateTime
                Dim scheduleType As String
                Dim toStr As String
                Dim ok As Boolean
                Dim fromEmail As String
                Dim IfNoOfMsgPerConn As Boolean
                Dim noOfMsgPerConn As Decimal
                Dim IfNoOfRecPerMsg As Boolean
                Dim noOfRecPerMsg As Decimal
                Dim line As String
                Dim smtpMail As New SMTP_Client
                If Directory.Exists(mainStorage & "Outbox") Then
                    For Each outMail In Directory.GetFiles(mainStorage & "Outbox", "*.mail")
                        IfNoOfMsgPerConn = False
                        IfNoOfRecPerMsg = False
                        WriteMsglog("sending mail..." & outMail)
                        Dim toList As New ArrayList
                        Dim mailMsg As New MimeConstructor
                        FS = New FileStream(outMail, FileMode.Open, FileAccess.Read)
                        SR = New StreamReader(FS)
                        If SR.Peek > 0 Then
                            Do While SR.Peek > 0
                                line = SR.ReadLine
                                If line.StartsWith("Account:") Then
                                    actName = line.Substring(9)
                                    xNode = xDoc.SelectSingleNode("descendant::Account[Name='" & actName & "']")
                                    If Not (xNode Is Nothing) Then
                                        smtpMail.SmartHost = xNode.SelectSingleNode("SMTPSvr").InnerText
                                        smtpMail.UseSmartHost = CBool(smtpMail.SmartHost <> "")
                                        smtpMail.Username = ""
                                        smtpMail.Password = ""
                                        If Not xNode.SelectSingleNode("Auth") Is Nothing Then
                                            Select Case xNode.SelectSingleNode("Auth").InnerText
                                                Case "1" 'base auth
                                                    'mailMsg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'basic authentication
                                                    'mailMsg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", _
                                                    smtpMail.Username = xNode.SelectSingleNode("SmtpUserName").InnerText   'set your username here
                                                    'mailMsg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", _
                                                    smtpMail.Password = DecPW(xNode.SelectSingleNode("SmtpPassword").InnerText)      'set your password here
                                                Case "2" 'use pop user pass
                                                    'mailMsg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate", "1") 'basic authentication
                                                    ' mailMsg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendusername", _
                                                    smtpMail.Username = xNode.SelectSingleNode("SvrAccountName").InnerText    'set your username here
                                                    'mailMsg.Fields.Add("http://schemas.microsoft.com/cdo/configuration/sendpassword", _
                                                    smtpMail.Password = DecPW(xNode.SelectSingleNode("SvrAccountPassword").InnerText)   'set your password here
                                            End Select
                                        End If
                                        If Not xNode.SelectSingleNode("IfNoOfMsgPerConn") Is Nothing Then
                                            If xNode.SelectSingleNode("IfNoOfMsgPerConn").InnerText = "True" Then
                                                IfNoOfMsgPerConn = True
                                            End If
                                            noOfMsgPerConn = CDec(xNode.SelectSingleNode("NoOfMsgPerConn").InnerText)
                                        End If

                                        If Not xNode.SelectSingleNode("IfNoOfRecPerMsg") Is Nothing Then
                                            If xNode.SelectSingleNode("IfNoOfRecPerMsg").InnerText = "True" Then
                                                IfNoOfRecPerMsg = True
                                            End If
                                            noOfRecPerMsg = CDec(xNode.SelectSingleNode("NoOfRecPerMsg").InnerText)
                                        End If

                                    Else
                                        WriteMsglog("[SendOutBox]Smtp setting not found, use default", enulevel.critical)
                                        smtpMail.SmartHost = ""
                                        smtpMail.UseSmartHost = False
                                    End If
                                ElseIf line.StartsWith("MailType:") Then
                                    mailType = line.Substring(10)
                                ElseIf line.StartsWith("From:") Then
                                    'state = SR.ReadLine.Substring(8)
                                    mailMsg.From = line.Substring(6)
                                    fromEmail = mailMsg.From.Substring(mailMsg.From.IndexOf("<") + 1)
                                    fromEmail = fromEmail.Substring(0, fromEmail.Length - 1)
                                ElseIf line.StartsWith("To: ") Then
                                    toMask = line.Substring(4)
                                ElseIf line.StartsWith("Cc:") Then
                                    mailMsg.Cc = New String() {line.Substring(4)}
                                ElseIf line.StartsWith("Bcc:") Then
                                    mailMsg.Bcc = New String() {line.Substring(5)}
                                ElseIf line.StartsWith("Subject:") Then
                                    mailMsg.Subject = line.Substring(9)
                                ElseIf line.StartsWith("ScheduleType:") Then
                                    scheduleType = line.Substring(14)
                                ElseIf line.StartsWith("ScheduleAt:") Then
                                    If IsDate(line.Substring(12)) Then
                                        scheduleAt = CDate(line.Substring(12))
                                    Else
                                        scheduleAt = Now
                                    End If
                                ElseIf line.StartsWith("Attachment: ") Then
                                    mailMsg.Attachments.Add(New Attachment(line.Substring(Len("Attachment: "))))
                                ElseIf line = "" Then
                                    Exit Do
                                End If
                            Loop
                            mailMsg.Body = SR.ReadToEnd
                            If Not CheckProductID() Then
                                mailMsg.Body = mailMsg.Body & vbNewLine & _
                                "Thank you for evaluating Faena Mail!" & vbNewLine & _
                                "This is an unregister version. Please visit http://www.faena.com to register."
                            End If
                            'SR.Close()
                            'FS.Close()
                            If scheduleAt <= Now Then
                                If mailType = "Mail List" Then
                                    Dim db As New Functions
                                    Dim rdrSelEmail As OdbcDataReader
                                    rdrSelEmail = db.GetReaderObject(toMask)
                                    While rdrSelEmail.Read
                                        toList.Add(rdrSelEmail.Item("Name").ToString & "<" & rdrSelEmail.Item("Email").ToString & ">")
                                    End While
                                    rdrSelEmail.Close()
                                Else
                                    toList.Add(toMask)
                                End If
                                WriteMsglog("Sending mail to " & toList.Count & " email address(es)")
                                reg.SaveSetting("Faena Mail\DisplayStatus", "Outbox", "Sending mail to " & toList.Count & " email address(es)")
                                For Each toStr In toList
                                    mailMsg.To = New String() {toStr.Replace(",", "")}
                                    WriteMsglog("Sending to " & mailMsg.To(0))
                                    reg.SaveSetting("Faena Mail\DisplayStatus", "Outbox", "Sending to " & mailMsg.To(0))
                                    Try
                                        Try
                                            If smtpMail.Send(New String() {toStr.Replace(",", "")}, _
                                                fromEmail, _
                                                mailMsg.ConstructBinaryMime) Then
                                                SaveSent(actName, mailMsg)
                                                ok = True
                                            Else
                                                WriteMsglog("SMTP error: " + smtpMail.Errors(0).ErrorText)
                                            End If
                                        Catch e As Exception
                                            Dim errorMessage As String = ""
                                            Dim curException As Exception = e
                                            Do While (Not curException Is Nothing)
                                                errorMessage += curException.Message + " " 'vbNewLine
                                                curException = curException.InnerException
                                            Loop
                                            'throw a new exception with errorMessage as it's message...
                                            Throw New Exception(errorMessage)
                                        End Try
                                    Catch exp As Exception
                                        WriteMsglog("Error: " & exp.Message, enulevel.critical)
                                        'reg.SaveSetting("Faena Mail\DisplayStatus", "Send failed to " & mailMsg.To, "Error:" & exp.Message)
                                    End Try
                                Next
                            Else
                                WriteMsglog("No account for send email!", enulevel.critical)
                            End If
                        Else
                            WriteMsglog("wrong email size", enulevel.critical)
                        End If
                        SR.Close()
                        FS.Close()
                        If ok Then
                            If scheduleType <> "One time only" Then
                                RescheduleOut(outMail, scheduleType)
                            End If
                            reg.SaveSetting("Faena Mail\DisplayStatus", "Outbox", "Complete sending mail to " & toList.Count & " email address(es)")
                            WriteMsglog("Email processed.")
                            File.Delete(outMail)
                        Else
                            reg.SaveSetting("Faena Mail\DisplayStatus", "Outbox", "Error sending mail to " & toList.Count & " email address(es)")
                            WriteMsglog("Send email not OK.")
                        End If
                        toList = Nothing
                        mailMsg = Nothing
                    Next
                End If
                reg.SaveSetting("Faena Mail", "CurrentAction", "")
            End If
        Catch ex As Exception
            WriteMsglog(ex.ToString)
            reg.SaveSetting("Faena Mail", "CurrentAction", "Error")
        End Try
    End Sub

    Public Sub SendOutbox()
        Dim newThread As New Threading.Thread(AddressOf SendOutboxThread)
        newThread.Start()
    End Sub

    Private Sub CheckEmailThread()
        reg.SaveSetting("Faena Mail\", "CurrentAction", "Checking mail...")
        Dim xDoc As New XmlDocument
        mainStorage = reg.GetSetting("Faena Mail\", "MsgStore", AppDomain.CurrentDomain.BaseDirectory)
        If File.Exists(mainStorage & "accounts.xml") Then
            xDoc.Load(mainStorage & "accounts.xml")
            xNodelist = xDoc.SelectNodes("/Accounts/Account")
            For Each xNode In xNodelist
                If Not (xNode.SelectSingleNode("Name") Is Nothing) Then
                    If CType(xNode.SelectSingleNode("Included").InnerText, Boolean) = True Then
                        acctName = xNode.SelectSingleNode("Name").InnerText
                        Dim ok As Boolean
                        Try
                            c.Connect(xNode.SelectSingleNode("POPSvr").InnerText, _
                                CInt(xNode.SelectSingleNode("POPPort").InnerText))
                            ok = True
                        Catch ex As Exception
                            WriteMsglog("Error connecting: " & ex.Message)
                            ok = False
                        End Try
                        If ok Then
                            c.Authenticate(xNode.SelectSingleNode("SvrAccountName").InnerText, _
                                DecPW(xNode.SelectSingleNode("SvrAccountPassword").InnerText), False)
                            ok = False
                            Dim mInf As POP3_MessagesInfo
                            Try
                                mInf = c.GetMessagesInfo()
                                ok = True
                            Catch ex As Exception
                                WriteMsglog("GetMessagesInfo: " & ex.Message)
                            End Try
                            If ok Then
                                Dim hTable As Hashtable = c.GetUidlList()
                                Dim uid As String
                                For Each uid In hTable.Keys
                                    uniqueID = CInt(hTable.Item(uid))
                                    WriteMsglog("Message No.:" & uid & " UniqueID:" & uniqueID.ToString)
                                    If Not File.Exists(mainStorage & "Inbox\" & uid & ".mail") Then
                                        mail.msg = ""
                                        Dim ifMessageOnServer As Boolean
                                        ifMessageOnServer = False
                                        Try
                                            mail.size = mInf.GetMessageInfo(uniqueID).MessageSize
                                            ifMessageOnServer = True
                                        Catch ex As Exception
                                            WriteMsglog(ex.Message)
                                        End Try
                                        If ifMessageOnServer Then
                                            WriteMsglog("size:" & mail.size.ToString)
                                            mail.msgTop = Encoding.ASCII.GetString(c.GetTopOfMessage(uniqueID, 10))
                                            'WriteMsglog(mail.msgTop)
                                            If GetScripts(acctName) Then
                                                mail.status = "P" 'processed
                                            End If
                                            If mail.msg = "" Then
                                                mail.msg = Encoding.ASCII.GetString(c.GetMessage(uniqueID))
                                            End If
                                            If AddToFolder("Inbox", acctName, uid, mail) Then
                                                If OkToDelete("Inbox", uid, acctName) Then
                                                    WriteMsglog("Deleteting from server:" & uid)
                                                    c.DeleteMessage(uniqueID)
                                                End If
                                            End If
                                            reg.SaveSetting("Faena Mail\", "NewMailNotify", "True")
                                            reg.SaveSetting("Faena Mail\", "NotifyShowed", 0)
                                        End If
                                    Else
                                        If OkToDelete("Inbox", uid, acctName) Then
                                            WriteMsglog("Deleteting from server:" & uid)
                                            c.DeleteMessage(uniqueID)
                                        End If
                                    End If
                                Next
                            End If
                            Try
                                c.Disconnect()
                            Catch
                            End Try
                        End If
                    End If
                End If
            Next xNode
            reg.SaveSetting("Faena Mail\", "CurrentAction", "")
            xDoc = Nothing
        End If
    End Sub

    Public Sub CheckEmail()
        Dim newThread As New Threading.Thread(AddressOf CheckEmailThread)
        newThread.Start()
    End Sub

    Private Sub WriteMsglog(ByVal msg As String, Optional ByVal level As enulevel = enulevel.information)
        Dim SW As StreamWriter
        Dim FS As FileStream
        Dim logFile As String = reg.GetSetting("Faena Mail\", "LogFile", AppDomain.CurrentDomain.BaseDirectory & "FaenaMailDebug.log")
        Try
            FS = New FileStream(logFile, FileMode.Append)
        Catch ex As Exception
            logFile = AppDomain.CurrentDomain.BaseDirectory & "FaenaMailDebug.log"
            FS = New FileStream(logFile, FileMode.Append)
            MsgBox("Create log file error." & ex.Message & vbNewLine & _
            "Message:" & vbNewLine & msg, MsgBoxStyle.Critical)
        End Try
        SW = New StreamWriter(FS)
        SW.Write(Now)
        SW.Write(vbTab)
        SW.WriteLine(msg)
        SW.Close()
        FS.Close()
    End Sub

    Private Sub SaveSent(ByVal actName As String, ByVal m As MimeConstructor)
        If Not Directory.Exists(mainStorage & "Sent Items\") Then Directory.CreateDirectory(mainStorage & "Sent Items\")
        Dim fileName As String
        Do
            fileName = mainStorage & "Sent Items\" & Path.GetFileNameWithoutExtension(Path.GetTempFileName) & ".mail"
        Loop While File.Exists(fileName)
        Dim FS As New FileStream(fileName, FileMode.CreateNew, FileAccess.Write)
        Dim SW As New StreamWriter(FS)
        SW.WriteLine("Account: " & actName)
        SW.WriteLine("Date: " & Now.UtcNow.ToLongDateString & " " & Now.UtcNow.ToLongTimeString)
        SW.WriteLine("MailType: Sent")
        SW.WriteLine("From: " & DBStr(m.From))
        SW.WriteLine("To: " & DBStr(m.To(0)))
        Try
            SW.WriteLine("Cc: " & DBStr(m.Cc(0)))
        Catch
        End Try
        Try
            SW.WriteLine("Bcc: " & DBStr(m.Bcc(0)))
        Catch
        End Try
        SW.WriteLine("Subject: " & DBStr(m.Subject))
        SW.WriteLine("ScheduleType: ")
        SW.WriteLine("ScheduleAt: ")
        For Each at As Attachment In m.Attachments
            SW.WriteLine("Attachment: " & at.FileName)
        Next
        SW.WriteLine("Status: R")
        SW.WriteLine()
        SW.Write(DBStr(m.Body))
        SW.Close()
        FS.Close()
        reg.SaveSetting("Faena Mail\DisplayStatus", "SentItems", reg.GetSetting("Faena Mail\DisplayStatus", "SentItems", 0) + 1)
    End Sub

    Private Sub RescheduleOut(ByVal f As String, ByVal scheduleType As String)
        Dim fileName As String
        Dim line As String
        Do
            fileName = mainStorage & "Outbox\" & Path.GetFileNameWithoutExtension(Path.GetTempFileName) & ".mail"
        Loop While File.Exists(fileName)
        Dim FS As New FileStream(fileName, FileMode.CreateNew, FileAccess.Write)
        Dim SW As New StreamWriter(FS)
        Dim FR As New FileStream(f, FileMode.Open, FileAccess.Read)
        Dim SR As New StreamReader(FR)
        Do While SR.Peek > 0
            line = SR.ReadLine
            If line Like "ScheduleAt: *" Then
                Exit Do
            Else
                SW.WriteLine(line)
            End If
        Loop
        SW.Write("ScheduleAt: ")
        line = line.Substring(12)
        Select Case scheduleType
            Case "Everyday"
                SW.WriteLine(CType(line, DateTime).AddDays(1))
            Case "Every week"
                SW.WriteLine(CType(line, DateTime).AddDays(7))
            Case "Every month"
                SW.WriteLine(CType(line, DateTime).AddMonths(1))
            Case Else
                WriteMsglog("Error scheduletype:" & scheduleType, enulevel.critical)
                SW.WriteLine()
        End Select
        SW.Write(SR.ReadToEnd)
        SR.Close()
        SW.Close()
        FS.Close()
        FR.Close()
    End Sub

    Private Function GetScripts(ByVal acctName As String) As Boolean
        Dim ok As Boolean = True
        Dim xScript As XmlNode
        For Each xScript In xNode.SelectNodes("Scripts/Script")
            ok = GetScript(xScript.SelectSingleNode("Path").InnerText)
        Next
        Return ok
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
        SW.WriteLine("MailType: New")
        If mail.status = "" Then mail.status = "N"
        SW.Write(mail.msg)
        SW.Close()
        FS.Close()
        AddToFolder = True
        reg.SaveSetting("Faena Mail\DisplayStatus", "DownloadItems", _
            reg.GetSetting("Faena Mail\DisplayStatus", "DownloadItems", 0) + 1)
    End Function
    Private Function OkToDelete(ByVal folder As String, ByVal uid As String, ByVal acctID As String) As Boolean
        Dim age As Long
        'Try
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
        'Catch ex As Exception
        'Debuglog(ex)
        'End Try
        Return CheckDelOpt(age, acctID)
    End Function

    Private Function CheckDelOpt(ByVal age As Long, ByVal acctID As String) As Boolean
        'If CType(myReader.Item("LeaveOnSvr"), Boolean) = False Then 'Del right away
        WriteMsglog("LeaveOnSvr option not supported.")
        Return True
        'Else  'Leave On Server
        '    If CType(myReader.Item("RemoveFromSvr"), Boolean) = True Then
        '        If (age > CInt(myReader.Item("HowManyDays"))) Then
        '            Return True
        '        End If
        '    Else
        '        Return False
        '    End If
        'End If
    End Function

    Private Function LoadNodeText(ByVal xnode As XmlNode, ByVal strname As String) As String
        Dim oneNode As XmlNode
        oneNode = xnode.SelectSingleNode(strname)
        If Not oneNode Is Nothing Then
            Return oneNode.InnerText
        Else
            Return ""
        End If
    End Function

    Private Function GetScript(ByVal scriptFile As String) As Boolean
        WriteMsglog("[Runscript]Script file:" & scriptFile)
        Dim totalScript As Integer = -1
        Dim startUpBox As String
        Dim xDoc As New XmlDocument
        xDoc.Load(scriptFile)
        Dim xnodeList As XmlNodeList
        Dim xnode As XmlNode
        xnode = xDoc.SelectSingleNode("/scriptfile")
        If Not xnode.Attributes("startupbox") Is Nothing Then
            startUpBox = xnode.Attributes("startupbox").InnerText
        End If
        If startUpBox = "" Then
            WriteMsglog("[GetScript]startup script box not define", enulevel.critical)
            Return False
        Else
            xnodeList = xDoc.SelectNodes("/scriptfile/scriptbox")
            For Each xnode In xnodeList
                totalScript = totalScript + 1
                ReDim Preserve aryRunScript(totalScript)
                With aryRunScript(totalScript)
                    .boxType = xnode.Attributes("type").InnerText
                    .boxName = LoadNodeText(xnode, "name")
                    .boxLeft = XmlConvert.ToInt32(LoadNodeText(xnode, "left"))
                    .boxTop = XmlConvert.ToInt32(LoadNodeText(xnode, "top"))
                    .boxText = LoadNodeText(xnode, "textbox")
                    .boxSettingIf = XmlConvert.ToInt32(LoadNodeText(xnode, "settingif"))
                    .boxSettingIfDes = LoadNodeText(xnode, "settingifdes")
                    .boxSettingAction = XmlConvert.ToInt32(LoadNodeText(xnode, "settingaction"))
                    .boxSettingActionDes = LoadNodeText(xnode, "settingactiondes")
                    .boxSettingBcc = LoadNodeText(xnode, "settingbcc")
                    .boxSettingMsg = LoadNodeText(xnode, "settingmsg")
                    .boxSettingAttachment = LoadNodeText(xnode, "settingattachment")
                    .boxSettingPath = LoadNodeText(xnode, "settingpath")
                End With
            Next
            Return RunScript(startUpBox)
        End If
    End Function

    Private Function RunScript(ByVal scriptBoxName As String) As Boolean
        Dim passed As Boolean
        Dim i As Integer
        For i = 0 To aryRunScript.GetUpperBound(0)
            If aryRunScript(i).boxName = scriptBoxName Then
                Exit For
            End If
        Next

        Select Case aryRunScript(i).boxType
            Case "IfScript"
                passed = DoIfScript(i)
            Case "MailList"
                passed = DoMailList(i)
            Case "email"
                passed = DoEmail(i)
            Case "SQL"
                passed = DoSQL(i)
            Case Else
                WriteMsglog("Unknown boxtype:" & aryRunScript(i).boxType, enulevel.critical)
        End Select
        If passed Then
            If aryRunScript(i).boxSettingPath <> "" Then
                passed = RunScript(aryRunScript(i).boxSettingPath)
            End If
        End If
        Return passed
    End Function

    Private Function DoSQL(ByVal i As Integer) As Boolean
        Dim strEmail As String
        Dim getField As Boolean
        Dim aryAdr() As String
        Dim strTemp As String
        Dim aryEmails(,) As String
        Dim x As Integer = -1
        Dim y As Integer
        Dim ok As Boolean = True
        With aryRunScript(i)
            Dim cmds() As String = Split(.boxSettingIfDes, ";")
            Dim cmd As String
            For Each cmd In cmds
                cmd = Trim(cmd.Replace(vbCrLf, ""))
                If cmd <> "" Then
                    getField = False
                    If cmd.ToLower.StartsWith("select") Then getField = True
                    While cmd Like "*<=*>*"
                        strTemp = cmd.Substring(cmd.IndexOf("<="))
                        strTemp = strTemp.Substring(2, strTemp.IndexOf(">") - 2)
                        If strTemp = "Body" Then
                            strEmail = TextFromMime(BodyLine())
                        Else
                            strEmail = TopLine(mail.msgTop, strTemp)
                        End If
                        If strEmail = "" Then
                            Dim myEnumerator As IDictionaryEnumerator = aryVar.GetEnumerator()
                            While myEnumerator.MoveNext()
                                If myEnumerator.Key.ToString = strTemp Then
                                    strEmail = myEnumerator.Value.ToString
                                    Exit While
                                End If
                            End While
                        End If
                        If strEmail = "" Then ok = False
                        cmd = cmd.Replace("<=" & strTemp & ">", strEmail)
                    End While
                    While cmd Like "*LookForWord(""*"")*"
                        strTemp = cmd.Substring(cmd.IndexOf("LookForWord("""))
                        strTemp = strTemp.Substring(Len("LookForWord("""), strTemp.IndexOf(""")"))
                        strEmail = LookForWord(strTemp)
                        If strEmail = "" Then ok = False
                        cmd = cmd.Replace("<=" & strTemp & ">", strEmail)
                    End While
                    While cmd Like "*LookFor(""*"")*"
                        strTemp = cmd.Substring(cmd.IndexOf("LookFor("""))
                        strTemp = strTemp.Substring(Len("LookFor("""), strTemp.IndexOf(""")"))
                        strEmail = LookFor(strTemp)
                        If strEmail = "" Then ok = False
                        cmd = cmd.Replace("<=" & strTemp & ">", strEmail)
                    End While
                    If ok Then
                        Dim myRS As OdbcDataReader
                        Try
                            Dim db As New Functions
                            If getField Then
                                myRS = db.GetReaderObject(cmd)
                                If myRS.Read Then
                                    For i = 1 To myRS.FieldCount
                                        aryVar.Add(myRS.GetName(i - 1), myRS.Item(i - 1).ToString)
                                    Next
                                Else
                                    ok = False
                                    WriteMsglog("[doSQL]Record not found.", enulevel.critical)
                                End If
                            Else
                                db.GetCommandObject(cmd, True)
                            End If
                        Catch ex As OdbcException
                            ok = False
                            WriteMsglog("[DoSQL]" & ex.Message, enulevel.critical)
                        Catch ex As Exception
                            ok = False
                            WriteMsglog("[DoSQL]" & ex.Message)
                        End Try
                        myRS.Close()
                    End If
                End If
            Next
        End With
        Return ok
    End Function

    Private Function DoIfScript(ByVal i As Integer) As Boolean
        Dim ok As Boolean = False
        Dim bCondition As Boolean
        With aryRunScript(i)
            Select Case .boxSettingIf
                Case 0 '0.	If the From line contains specific words
                    bCondition = InStr(TopLine(mail.msgTop, "From"), .boxSettingIfDes) > 0
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 1 '1.	If the From line in database fields
                    bCondition = InDBField(TopLine(mail.msgTop, "From"), .boxSettingIfDes)
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 2 '2.	If the Subject line contains  specific words
                    bCondition = InStr(TopLine(mail.msgTop, "Subject"), .boxSettingIfDes) > 0
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 3 '3.	If the Subject line in database fields
                    bCondition = InDBField(TopLine(mail.msgTop, "Subject"), .boxSettingIfDes)
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 4 '4.	If the message body contains  specific words
                    bCondition = InStr(BodyLine(), .boxSettingIfDes) > 0
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 5 '5.	If the message body in database fields
                    bCondition = InDBField(BodyLine(), .boxSettingIfDes)
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 6 '6.	If the To line contains  specific words
                    bCondition = InStr(TopLine(mail.msgTop, "To"), .boxSettingIfDes) > 0
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 7 '7.	If the To line in database fields
                    bCondition = InDBField(TopLine(mail.msgTop, "To"), .boxSettingIfDes)
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 8 '8.	If the CC line contains  specific words
                    bCondition = InStr(TopLine(mail.msgTop, "Cc"), .boxSettingIfDes) > 0
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 9 '9.	If the CC line in database fields
                    bCondition = InDBField(TopLine(mail.msgTop, "Cc"), .boxSettingIfDes)
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 10 '10.	If the To or CC line contains  specific words
                    bCondition = InStr(TopLine(mail.msgTop, "To"), .boxSettingIfDes) > 0 Or InStr(TopLine(mail.msgTop, "Cc"), .boxSettingIfDes) > 0
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
                Case 11 '11.	If the To or CC line in database fields
                    bCondition = InDBField(TopLine(mail.msgTop, "To"), .boxSettingIfDes) Or InDBField(TopLine(mail.msgTop, "Cc"), .boxSettingIfDes)
                    ok = DoAction(.boxSettingAction, .boxSettingActionDes, bCondition)
            End Select
        End With
        Return ok
    End Function

    Private Function DoMailList(ByVal i As Integer) As Boolean
        Dim strEmail As String = ""
        Dim aryAdr() As String
        Dim strName As String
        Dim aryEmails(,) As String
        Dim x As Integer = -1
        Dim y As Integer
        Dim ok As Boolean = True
        'Try
        With aryRunScript(i)
            If .boxSettingIfDes Like "<=*>" Then
                strEmail = TopLine(mail.msgTop, .boxSettingIfDes.Substring(3, .boxSettingIfDes.Length - 3))
            ElseIf .boxSettingIfDes Like "LookFor(""*"")" Then
                strEmail = LookFor(.boxSettingIfDes.Substring(9, .boxSettingIfDes.Length - 11))
            Else
                WriteMsglog("Unknow boxSettingIfDes:" & .boxSettingIfDes, enulevel.critical)
                ok = False
            End If
            If strEmail.Length > 0 Then
                WriteMsglog("Found " & strEmail)
                aryAdr = Split(strEmail, ";")
                For Each strName In aryAdr
                    x = x + 1
                    ReDim Preserve aryEmails(x, 1)
                    If InStr(strEmail, "<") > 0 Then
                        aryEmails(x, 0) = Trim$(strName.Substring(0, strEmail.IndexOf("<")))
                        strName = strName.Substring(strName.IndexOf("<") + 1)
                        aryEmails(x, 1) = Trim$(strName.Substring(0, strEmail.IndexOf(">")))
                    Else
                        aryEmails(x, 1) = Trim$(strName)
                        aryEmails(x, 0) = Trim$(strName)
                    End If
                Next
                For y = 0 To x
                    Try
                        Dim db As New Functions
                        Select Case .boxSettingIf
                            Case 0 'Add list
                                db.GetCommandObject("insert into mailinglist (emailaddress,name) values ('" & aryEmails(y, 1) & "','" & aryEmails(y, 0) & "')", True)
                            Case 1 'Remove list
                                db.GetCommandObject("update mailinglist set removed=Yes where emailaddress ='" & aryEmails(y, 1) & "'", True)
                        End Select
                    Catch ex As Exception
                        WriteMsglog("[DomailList]" & ex.Message)
                    End Try
                Next
            End If
        End With


        Return ok
    End Function

    Private Function ReplaceField(ByVal source As String) As String
        Dim strTemp As String
        Dim strTemp2 As String
        While source Like "*<=*>*"
            strTemp = source.Substring(source.IndexOf("<="))
            strTemp = strTemp.Substring(2, strTemp.IndexOf(">") - 2)
            If strTemp = "Body" Then
                strTemp2 = TextFromMime(BodyLine())
            Else
                strTemp2 = TopLine(mail.msgTop, strTemp)
            End If
            If strTemp2 = "" Then
                Dim myEnumerator As IDictionaryEnumerator = aryVar.GetEnumerator()
                While myEnumerator.MoveNext()
                    If myEnumerator.Key.ToString = strTemp Then
                        strTemp2 = myEnumerator.Value.ToString
                        Exit While
                    End If
                End While
            End If
            source = source.Replace("<=" & strTemp & ">", strTemp2)
        End While
        Try
            While source Like "*LookFor(""*"")*"
                strTemp = source.Substring(source.IndexOf("LookFor("""))
                strTemp = strTemp.Substring(9, strTemp.IndexOf(""")") - 9)
                strTemp2 = LookFor(strTemp)
                source = source.Replace("LookFor(""" & strTemp & """)", strTemp2)
            End While
        Catch ex As Exception
            WriteMsglog("[ReplaceField]" & ex.Message & " check keyword LookFor()", enulevel.critical)
        End Try
        Return source
    End Function

    Private Function DoEmail(ByVal i As Integer) As Boolean
        Try
            Dim strTemp As String
            Dim toAddr As String = ReplaceField(aryRunScript(i).boxSettingIfDes)
            Dim ccAddr As String = ReplaceField(aryRunScript(i).boxSettingActionDes)
            Dim bccAddr As String = ReplaceField(aryRunScript(i).boxSettingBcc)
            Dim subject As String = ReplaceField(aryRunScript(i).boxText)
            Dim messageFile As String
            If (toAddr Like acctName) Or (ccAddr Like acctName) Or (bccAddr Like acctName) Then
                WriteMsglog("[DoEmail]From address is the same as To address!", enulevel.critical)
            Else
                If Not Directory.Exists(mainStorage & "Outbox\") Then Directory.CreateDirectory(mainStorage & "Outbox\")
                Do
                    messageFile = mainStorage & "Outbox\" & Path.GetFileNameWithoutExtension(Path.GetTempFileName) & ".mail"
                Loop While File.Exists(messageFile)
                Dim f As New FileStream(messageFile, FileMode.Create, FileAccess.Write)
                Dim s As New StreamWriter(f)
                s.WriteLine("Account: " & acctName)
                s.WriteLine("Created: " & Now.UtcNow.ToLongDateString & " " & Now.UtcNow.ToLongTimeString)
                s.WriteLine("MailType: AutoResponse")
                s.WriteLine("From: " & acctName)
                s.WriteLine("To: " & toAddr)
                s.WriteLine("Cc: " & ccAddr)
                s.WriteLine("Bcc: " & bccAddr)
                s.WriteLine("Subject: " & subject)
                s.WriteLine("ScheduleType: One Time Only")
                s.WriteLine("ScheduleAt: " & Now.ToString)
                s.WriteLine("Status: N")
                s.WriteLine()
                strTemp = ReplaceField(aryRunScript(i).boxSettingMsg)
                s.WriteLine(strTemp)
                s.WriteLine()
                s.Close()
                f.Close()
                DoEmail = True
            End If
        Catch ex As Exception
            WriteMsglog("[DoEmail] exception:" & ex.Message, enulevel.critical)
        End Try
    End Function

    Private Function BodyLine() As String
        If mail.msg = "" Then
            mail.msg = Encoding.ASCII.GetString(c.GetMessage(uniqueID))
        End If
        Return mail.msg.Substring(mail.msg.IndexOf(vbCrLf & vbCrLf) + 2)
    End Function



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
    'Private Function ExpandString(ByVal strOrg As String) As String
    '    Dim toAddr As String = strOrg
    '    Dim s As String

    '    If strOrg Like "*<=*>*" Then
    '        s = strOrg.Substring(strOrg.IndexOf("<="))
    '        s = s.Substring(0, s.IndexOf(">") + 1)
    '        toAddr = Replace(toAddr, s, TopLine(mail.msgTop, s.Substring(2, s.Length - 3)))
    '    End If
    '    If strOrg Like "*LookFor(""*"")*" Then
    '        s = strOrg.Substring(strOrg.IndexOf("LookFor("""))
    '        s = s.Substring(0, s.IndexOf(""")"))
    '        toAddr = Replace(toAddr, s, LookFor(s.Substring(9, s.Length - 11)))
    '    End If
    '    Return toAddr
    'End Function
    Private Function LookForWord(ByVal idString As String) As String
        If mail.msg = "" Then
            mail.msg = Encoding.ASCII.GetString(c.GetMessage(uniqueID))
        End If
        Dim pos As Integer = mail.msg.IndexOf(idString)
        Dim st As String = mail.msg.Substring(pos + idString.Length).TrimStart
        Dim ws As Integer = st.IndexOf(" ")
        Dim lf As Integer = st.IndexOf(vbNewLine)
        If ws < 0 And lf < 0 Then
            Return st
        ElseIf ws < 0 And lf > 0 Then
            Return st.Substring(0, lf)
        ElseIf ws > 0 And lf < 0 Then
            Return st.Substring(0, ws)
        ElseIf ws > 0 And lf > 0 And ws < lf Then
            Return st.Substring(0, ws)
        ElseIf ws > 0 And lf > 0 And ws > lf Then
            Return st.Substring(0, lf)
        Else
            Return ""
        End If
    End Function

    Private Function LookFor(ByVal idString As String) As String
        If mail.msg = "" Then
            mail.msg = Encoding.ASCII.GetString(c.GetMessage(uniqueID))
        End If
        Dim pos As Integer = mail.msg.IndexOf(idString)
        Dim st As String = mail.msg.Substring(pos + idString.Length).TrimStart
        Dim lf As Integer = st.IndexOf(vbNewLine)
        If lf < 0 Then
            Return st
        Else
            Return st.Substring(0, lf)
        End If
    End Function

    Private Function DoAction(ByVal strAction As Integer, ByVal strActionDes As String, ByVal bCondition As Boolean) As Boolean
        Select Case strAction
            Case 0 '0.	Delete it
                If bCondition Then
                    mail.folder = "Delete"
                    Return True
                End If
            Case 1 '1.	Mark as Junk mail
                If bCondition Then
                    mail.folder = "Junk"
                    Return True
                End If
            Case 2 '2.	Forward it to people
                If bCondition Then
                    mail.forward = strActionDes
                    Return True
                End If
            Case 3 '3.	Mark it as read
                If bCondition Then
                    mail.status = "R"
                    Return True
                End If
            Case 4 '4. Do not process
                If bCondition Then
                    Return False
                Else
                    Return True
                End If
            Case 5 '5. Go on
                If bCondition Then
                    Return True
                End If
            Case Else
                WriteMsglog("Unknown action:" & strAction & ":" & strActionDes & "(" & bCondition & ")", enulevel.critical)
        End Select
    End Function

    Private Function InDBField(ByVal str1 As String, ByVal strSQL As String) As Boolean
        Dim db As New Functions
        Dim myReader As OdbcDataReader = db.GetReaderObject(strSQL)
        Dim i As Integer
        While myReader.Read
            For i = 1 To myReader.FieldCount
                If InStr(str1, myReader.Item(i).ToString) > 0 Then
                    Return True
                End If
            Next
        End While
    End Function


End Class
