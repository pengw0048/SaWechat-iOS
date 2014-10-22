Imports System.IO
Imports System.Text
Imports System.Security.Cryptography
Imports System.Data.SQLite

Public Class Form1

    Dim rightpath As String

    Private Sub Label5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Label5.Click
        MsgBox("AMR是语音的原有储存格式，质量好，文件小，但在浏览器中可能无法播放。" + vbCrLf + _
               "请安装相应解码包，例如K-Lite Codec Pack。" + vbCrLf + _
               "MP3是通用格式，转换中可能损失部分音质，增加文件大小。" + vbCrLf + _
               "要不，您都试试？")
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        MsgBox("支付宝付款地址功能已被取消。" + vbCrLf + "麻烦您付款到支付宝账号tiancaihb@sina.com。")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        FolderBrowserDialog1.ShowDialog()
        TextBox1.Text = FolderBrowserDialog1.SelectedPath
        rightpath = TextBox1.Text
        If Not rightpath.EndsWith("\") Then
            rightpath = (rightpath & "\")
        End If
        If Directory.Exists(rightpath & "Documents") And Directory.Exists(rightpath & "Library") Then
            Label2.Text = "正确"
            Label2.ForeColor = Color.Green
            Button2.Enabled = True
            ComboBox1.Items.Clear()
            ComboBox1.Text = ""
            Dim DirInfo As New DirectoryInfo(rightpath & "Documents")
            Dim FilInfo As DirectoryInfo
            For Each FilInfo In DirInfo.GetDirectories
                If FilInfo.Name.Length = 32 And FilInfo.Name.Replace("0", "") <> "" Then
                    ComboBox1.Items.Add(FilInfo.Name)
                End If
            Next
            If ComboBox1.Items.Count = 0 Then
                Label2.Text = "未找到"
                Label2.ForeColor = Color.Red
                Button2.Enabled = False
                MsgBox("没有找到用户登录记录？")
            Else
                ComboBox1.SelectedIndex = 0
            End If
        Else
            Label2.Text = "未找到"
            Label2.ForeColor = Color.Red
            Button2.Enabled = False
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim ChatMD5Length As Integer
        Dim span As TimeSpan
        Button1.Enabled = False
        Button2.Enabled = False
        GroupBox1.Enabled = False
        GroupBox2.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        ComboBox1.Enabled = False
        ProgressBar1.Value = 0
        Label7.Text = "读取基本信息"
        Application.DoEvents()
        Dim DocumentsPath As String = (rightpath & "Documents\" & ComboBox1.Text & "\")
        Dim LibraryPath As String = (rightpath & "Library\")
        Dim NowDateString As String = Now.ToString("yyyyMMddHHmmss")
        Dim SavePath As String = Path.GetDirectoryName(Reflection.Assembly.GetExecutingAssembly.Location)
        If Not SavePath.EndsWith("\") Then
            SavePath = SavePath & "\"
        End If
        Dim ResBase As String = SavePath & "res\"
        SavePath = SavePath & NowDateString & "\"
        Directory.CreateDirectory(SavePath)
        Dim connection As New SQLiteConnection
        Dim command As New SQLiteCommand
        Dim dset As New DataSet
        connection.ConnectionString = ("data source=" & DocumentsPath & "DB\MM.sqlite")
        connection.Open()
        command.Connection = connection
        command.CommandText = "select count(*) from Friend"
        Dim FriendsLoopCount As Integer = CInt(command.ExecuteScalar()) + 10
        Dim Friend_UsrName(FriendsLoopCount) As String
        Dim Friend_NickName(FriendsLoopCount) As String
        Dim Friend_Sex(FriendsLoopCount) As Integer
        Dim FriendExt_ConRemark(FriendsLoopCount) As String
        Dim FriendExt_ConChatRoomMem(FriendsLoopCount) As String
        Dim Friend_UsrNameMD5(FriendsLoopCount) As String
        Dim Friend_UsrNameMD5_Alias(FriendsLoopCount) As String
        Dim PicUsr(FriendsLoopCount) As String
        Dim PicHD(FriendsLoopCount) As String
        Dim OwnerPotrait As String = ""
        Dim OwnerPotraitHD As String = ""
        If File.Exists(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0\" & Strings.Left(ComboBox1.Text, 2) & "\" & Strings.Mid(ComboBox1.Text, 3) & ".pic_usr") Then
            OwnerPotrait = ComboBox1.Text
        End If
        If File.Exists(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0\" & Strings.Left(ComboBox1.Text, 2) & "\" & Strings.Mid(ComboBox1.Text, 3) & ".pic_hd") Then
            OwnerPotraitHD = ComboBox1.Text
        End If
        command.CommandText = "select Friend.UsrName,Friend.NickName,Friend.Sex,Friend_Ext.ConRemark,Friend_Ext.ConChatRoomMem,Friend_Ext.ConStrRes2 from Friend join Friend_Ext on Friend.UsrName=Friend_Ext.UsrName"
        Dim reader As SQLiteDataReader = command.ExecuteReader
        Dim index As Integer = 0
        Do While reader.Read
            Friend_UsrName(index) = reader.GetString(0)
            Friend_UsrNameMD5(index) = MD5(Friend_UsrName(index), &H20)
            Friend_UsrNameMD5_Alias(index) = Friend_UsrNameMD5(index)
            If Not reader.IsDBNull(1) Then
                Friend_NickName(index) = reader.GetString(1)
            End If
            Friend_Sex(index) = reader.GetInt32(2)
            If Not reader.IsDBNull(3) Then
                FriendExt_ConRemark(index) = reader.GetString(3)
            End If
            If Not reader.IsDBNull(4) Then
                FriendExt_ConChatRoomMem(index) = reader.GetString(4)
            End If
            Dim AliasString As String = ""
            If Not reader.IsDBNull(5) Then
                AliasString = reader.GetString(5)
            End If
            If InStr(AliasString, "<alias>", CompareMethod.Text) <> 0 Then
                Friend_UsrNameMD5_Alias(index) = MD5(Split(Split(reader.GetString(5), "<alias>", -1, CompareMethod.Text)(1), "</alias>", -1, CompareMethod.Text)(0), &H20)
            End If
            If File.Exists(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0\" & Strings.Left(Friend_UsrNameMD5_Alias(index), 2) & "\" & Strings.Mid(Friend_UsrNameMD5_Alias(index), 3) & ".pic_usr") Then
                PicUsr(index) = Friend_UsrNameMD5_Alias(index)
            End If
            If File.Exists(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0\" & Strings.Left(Friend_UsrNameMD5(index), 2) & "\" & Strings.Mid(Friend_UsrNameMD5(index), 3) & ".pic_usr") Then
                PicUsr(index) = Friend_UsrNameMD5(index)
            End If
            If File.Exists(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0\" & Strings.Left(Friend_UsrNameMD5_Alias(index), 2) & "\" & Strings.Mid(Friend_UsrNameMD5_Alias(index), 3) & ".pic_hd") Then
                PicHD(index) = Friend_UsrNameMD5_Alias(index)
            End If
            If File.Exists(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0\" & Strings.Left(Friend_UsrNameMD5(index), 2) & "\" & Strings.Mid(Friend_UsrNameMD5(index), 3) & ".pic_hd") Then
                PicHD(index) = Friend_UsrNameMD5(index)
            End If
            index += 1
            Application.DoEvents()
        Loop
        Dim FriendsCount As Integer = index
        reader.Close()
        command.CommandText = "select name from sqlite_master where type='table' order by name;"
        reader = command.ExecuteReader
        Dim ChatMD5Index(FriendsLoopCount) As Integer
        Dim ChatMD5(FriendsLoopCount) As String
        Do While reader.Read
            Dim TableName As String = ""
            If Not reader.IsDBNull(0) Then
                TableName = reader.GetString(0)
            End If
            If (TableName.Length = &H25 And Strings.Left(TableName, 5).ToLower = "chat_") Then
                TableName = Strings.Mid(TableName, 6).ToLower
                For i = 0 To FriendsCount - 1
                    If TableName = Friend_UsrNameMD5(i) Or TableName = Friend_UsrNameMD5_Alias(i) Then
                        ChatMD5(ChatMD5Length) = TableName
                        ChatMD5Index(ChatMD5Length) = i
                        ChatMD5Length += 1
                        Exit For
                    End If
                Next
            End If
            Application.DoEvents()
        Loop
        reader.Close()
        If RadioButton3.Checked Then
            For index = 0 To ChatMD5Length - 1
                Dim NumInChatRoom As Integer
                Label7.Text = "处理聊天记录 " & Trim(index + 1) & "/" & Trim(ChatMD5Length)
                Application.DoEvents()
                Dim IndexInChatMD5 As Integer = ChatMD5Index(index)
                Dim ChatRoomUsrName() As String = Nothing
                Dim DisplayNickname() As String = Nothing
                Dim IndexInFriend_UsrName() As Integer = Nothing
                Dim isChatRoom As Boolean
                If Strings.Right(Friend_UsrName(IndexInChatMD5), 9).ToLower = "@chatroom" Then
                    isChatRoom = True
                    ChatRoomUsrName = Split(FriendExt_ConChatRoomMem(IndexInChatMD5), ";", -1, CompareMethod.Binary)
                    NumInChatRoom = UBound(ChatRoomUsrName) + 1
                    ReDim IndexInFriend_UsrName(NumInChatRoom)
                    ReDim DisplayNickname(NumInChatRoom)
                    For j = 0 To NumInChatRoom - 1
                        IndexInFriend_UsrName(j) = -1
                        Dim k As Integer = 0
                        For k = 0 To FriendsCount - 1
                            If (Friend_UsrName(k) = ChatRoomUsrName(j)) Then
                                IndexInFriend_UsrName(j) = k
                                Exit For
                            End If
                        Next
                        If (IndexInFriend_UsrName(j) <> -1) Then
                            DisplayNickname(j) = Friend_NickName(IndexInFriend_UsrName(j))
                            If (FriendExt_ConRemark(IndexInFriend_UsrName(j)) <> "") Then
                                DisplayNickname(j) = FriendExt_ConRemark(IndexInFriend_UsrName(j))
                            End If
                        Else
                            DisplayNickname(j) = ChatRoomUsrName(j)
                        End If
                    Next
                End If
                Dim ChatTitle As String = Friend_NickName(IndexInChatMD5)
                If FriendExt_ConRemark(IndexInChatMD5) <> "" Then
                    ChatTitle = FriendExt_ConRemark(IndexInChatMD5)
                End If
                command.CommandText = "select count(*) from Chat_" & ChatMD5(index)
                Dim ChatLinesCount As Integer = CInt(command.ExecuteScalar())
                Dim writer As New StreamWriter(SavePath & safeName(GoodName(ChatTitle)) & ".txt", False, Encoding.UTF8)
                For i = 0 To CInt((ChatLinesCount - 1) / 100 + 1)
                    ProgressBar1.Value = CInt(Math.Round((CDbl(i) / ((ChatLinesCount - 1) / 100 + 2)) * 100))
                    Application.DoEvents()
                    command.CommandText = "select MesLocalID,CreateTime,Message,Status,Des,Type from Chat_" & ChatMD5(index) & " order by CreateTime asc limit " & Trim(i * 100) & ",100"
                    reader = command.ExecuteReader
                    Do While reader.Read
                        Dim MesLocalID As Integer = reader.GetInt32(0)
                        Dim CreateTime As DateTime = New DateTime(1970, 1, 1)
                        span = New TimeSpan(reader.GetInt64(1) * &H989680)
                        CreateTime = TimeZone.CurrentTimeZone.ToLocalTime(CreateTime).Add(span)
                        Dim Message As String = ""
                        If Not reader.IsDBNull(2) Then
                            Message = reader.GetString(2)
                        End If
                        If Message Is Nothing Then
                            Message = ""
                        End If
                        Dim MesssageBackup As String = Message
                        Dim Status As Integer = reader.GetInt32(3)
                        Dim Des As Integer = reader.GetInt32(4)
                        Dim MsgType As Integer = reader.GetInt32(5)
                        If InStr(Message, "<voicemsg ", CompareMethod.Text) <> 0 Then
                            Message = "[语音]"
                        End If
                        If InStr(Message, "<emoji ", CompareMethod.Text) <> 0 Then
                            Message = "[表情]"
                        End If
                        If InStr(Message, "nickname=""", CompareMethod.Text) <> 0 Then
                            Message = "[名片]"
                        End If
                        If InStr(Message, "<emoticonmd5>", CompareMethod.Text) <> 0 Then
                            If Split(Split(Message, "<emoticonmd5>", -1, CompareMethod.Text)(1), "</emoticonmd5>", -1, CompareMethod.Text)(0).Trim <> "" Then
                                Message = "[表情]"
                            ElseIf Split(Split(Message, "<title>", -1, CompareMethod.Text)(1), "</title>", -1, CompareMethod.Text)(0).Trim <> "" Then
                                Message = "[链接]"
                            End If
                        End If
                        If InStr(Message, "<img ", CompareMethod.Text) <> 0 Then
                            Message = "[图片]"
                        End If
                        If InStr(Message, "<location ", CompareMethod.Text) <> 0 Then
                            Message = "[位置]"
                        End If
                        If MsgType = 10000 Then
                            writer.WriteLine("系统消息: " & Message)
                            Continue Do
                        End If
                        If Status = 2 Or Status = 3 Then
                            writer.WriteLine(TextBox2.Text & "(" & CreateTime.ToString & "): " & Message)
                            Continue Do
                        End If
                        If Status = 4 Then
                            If isChatRoom Then
                                Dim FoundUsrName As Boolean
                                Dim tmppos As Integer = InStr(Message, ":", CompareMethod.Binary)
                                If tmppos > 1 Then
                                    Dim tmpUsrName As String = Strings.Left(Message, tmppos - 1)
                                    For m = 0 To NumInChatRoom - 1
                                        If ChatRoomUsrName(m) = tmpUsrName And IndexInFriend_UsrName(m) <> -1 Then
                                            writer.WriteLine(DisplayNickname(m) & "(" & CreateTime.ToString & "): " & IIf(Message = MesssageBackup, MyTrim(Mid(Message, tmppos + 1)), Message))
                                            FoundUsrName = True
                                            Exit For
                                        End If
                                    Next
                                    If Not FoundUsrName Then
                                        writer.WriteLine(tmpUsrName & "(" & CreateTime.ToString & "): " & IIf(Message = MesssageBackup, MyTrim(Mid(Message, tmppos + 1)), Message))
                                        FoundUsrName = True
                                    End If
                                End If
                                If Not FoundUsrName Then
                                    writer.WriteLine("(" & CreateTime.ToString & "): " & Message)
                                End If
                                Continue Do
                            End If
                            writer.WriteLine(ChatTitle & "(" & CreateTime.ToString & "): " & Message)
                            Continue Do
                        End If
                    Loop
                    reader.Close()
                Next
                writer.Close()
            Next

        ElseIf RadioButton2.Checked Then
            Dim HasEmojiNum(&H1FBD0) As Boolean
            Dim DirInfo As New DirectoryInfo((ResBase & "Emoji"))
            Dim FilInfo As FileInfo
            For Each FilInfo In DirInfo.GetFiles
                If FilInfo.Name.ToLower.EndsWith(".png") Then
                    Try
                        If Val(Strings.Left(FilInfo.Name, FilInfo.Name.Length - 4)) > 0 And Val(Strings.Left(FilInfo.Name, FilInfo.Name.Length - 4)) < 130000 Then
                            HasEmojiNum(Val(Strings.Left(FilInfo.Name, FilInfo.Name.Length - 4))) = True
                        End If
                    Catch exception1 As Exception
                    End Try
                End If
            Next
            Directory.CreateDirectory(SavePath & "Potrait")
            Directory.CreateDirectory(SavePath & "Emotion")
            Directory.CreateDirectory(SavePath & "Emoji")
            My.Computer.FileSystem.CopyFile(ResBase & "DefaultProfileHead@2x.png", SavePath & "Potrait\DefaultProfileHead@2x.png", True)
            'changed in new version
            DirInfo = New DirectoryInfo(LibraryPath & "WechatPrivate\" & ComboBox1.Text & "\HeadImg\0")
            Label7.Text = "复制头像"
            Dim ProcessedCount As Integer = 0
            For Each FolderInfo As DirectoryInfo In DirInfo.GetDirectories
                ProcessedCount += 1
                For Each FilInfo In FolderInfo.GetFiles
                    ProgressBar1.Value = CInt(CDbl(ProcessedCount) / DirInfo.GetDirectories.Length * 100.0)
                    Application.DoEvents()
                    If FilInfo.Extension.ToLower = ".pic_usr" And FilInfo.Name.Length = 38 Then
                        My.Computer.FileSystem.CopyFile(FilInfo.FullName, SavePath & "Potrait\" & FolderInfo.Name & Strings.Left(FilInfo.Name, 30) & ".jpg", True)
                    ElseIf FilInfo.Extension.ToLower = ".pic_hd" And FilInfo.Name.Length = 37 Then
                        My.Computer.FileSystem.CopyFile(FilInfo.FullName, SavePath & "Potrait\" & FolderInfo.Name & Strings.Left(FilInfo.Name, 30) & "_hd.jpg", True)
                    End If
                Next
            Next
            Label7.Text = "复制表情"
            Application.DoEvents()
            My.Computer.FileSystem.CopyDirectory(ResBase & "Expression", SavePath & "Expression", True)
            Dim reader2 As New StreamReader(ResBase & "Expression\Expression.txt")
            ' ToDo: upload new expressions
            Dim ExpressionName(300) As String
            Dim ExpressionHTML(300) As String
            index = 1
            While reader2.EndOfStream = False
                Dim ts As String = reader2.ReadLine
                If Trim(ts) = "" Then Continue While
                ExpressionName(index) = ("[" & ts & "]")
                ExpressionHTML(index) = ("<img src=""Expression\Expression_" & Trim(index) & "@2x.png"" width=""18px"" height=""18px"" />")
                index += 1
            End While
            reader2.Close()
            For index = 0 To ChatMD5Length - 1
                Label7.Text = ("复制图片 " & Trim(index + 1) & "/" & Trim(ChatMD5Length))
                Application.DoEvents()
                Dim CurrentFriendIndex As Integer = ChatMD5Index(index)
                Dim path As String = (SavePath & Friend_UsrName(CurrentFriendIndex) & "_Img\")
                Directory.CreateDirectory(path)
                If Friend_UsrNameMD5(CurrentFriendIndex) <> "" AndAlso Directory.Exists(DocumentsPath & "Img\" & Friend_UsrNameMD5(CurrentFriendIndex)) Then
                    DirInfo = New DirectoryInfo(DocumentsPath & "Img\" & Friend_UsrNameMD5(CurrentFriendIndex))
                    ProcessedCount = 0
                    'Dim FilInfo As FileInfo
                    For Each FilInfo In DirInfo.GetFiles
                        ProgressBar1.Value = CInt(CDbl(ProcessedCount) / IIf(DirInfo.GetFiles.Length = 0, 1, DirInfo.GetFiles.Length) * 100.0)
                        ProcessedCount += 1
                        Application.DoEvents()
                        If FilInfo.Extension.ToLower = ".pic" Then
                            My.Computer.FileSystem.CopyFile(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".jpg", True)
                        ElseIf FilInfo.Extension.ToLower = ".pic_thum" Then
                            My.Computer.FileSystem.CopyFile(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & "_thum.jpg", True)
                        ElseIf FilInfo.Extension.ToLower = ".pic_hd" Then
                            My.Computer.FileSystem.CopyFile(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".jpg", True)
                        End If
                    Next
                End If
                If (Friend_UsrNameMD5_Alias(CurrentFriendIndex) <> "" And Friend_UsrName(CurrentFriendIndex) <> Friend_UsrNameMD5_Alias(CurrentFriendIndex)) AndAlso Directory.Exists(DocumentsPath & "Img\" & Friend_UsrNameMD5_Alias(CurrentFriendIndex)) Then
                    DirInfo = New DirectoryInfo(DocumentsPath & "Img\" & Friend_UsrNameMD5_Alias(CurrentFriendIndex))
                    ProcessedCount = 0
                    'Dim FilInfo As FileInfo
                    For Each FilInfo In DirInfo.GetFiles
                        ProgressBar1.Value = CInt(CDbl(ProcessedCount) / IIf(DirInfo.GetFiles.Length = 0, 1, DirInfo.GetFiles.Length) * 100.0)
                        ProcessedCount += 1
                        Application.DoEvents()
                        If FilInfo.Extension.ToLower = ".pic" Then
                            My.Computer.FileSystem.CopyFile(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".jpg", True)
                        ElseIf FilInfo.Extension.ToLower = ".pic_thum" Then
                            My.Computer.FileSystem.CopyFile(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & "_thum.jpg", True)
                        ElseIf FilInfo.Extension.ToLower = ".pic_hd" Then
                            My.Computer.FileSystem.CopyFile(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".jpg", True)
                        End If
                    Next
                End If
            Next
            For index = 0 To ChatMD5Length - 1
                Label7.Text = IIf(RadioButton5.Checked, "复制和转换语音 ", "复制语音 ") & Trim(index + 1) & "/" & Trim(ChatMD5Length)
                Application.DoEvents()
                Dim CurrentFriendIndex As Integer = ChatMD5Index(index)
                Dim path As String = (SavePath & Friend_UsrName(CurrentFriendIndex) & "_Aud\")
                Directory.CreateDirectory(path)
                If Friend_UsrNameMD5(CurrentFriendIndex) <> "" AndAlso Directory.Exists(DocumentsPath & "Audio\" & Friend_UsrNameMD5(CurrentFriendIndex)) Then
                    DirInfo = New DirectoryInfo(DocumentsPath & "Audio\" & Friend_UsrNameMD5(CurrentFriendIndex))
                    ProcessedCount = 0
                    'Dim FilInfo As FileInfo
                    For Each FilInfo In DirInfo.GetFiles
                        ProgressBar1.Value = CInt(CDbl(ProcessedCount) / IIf(DirInfo.GetFiles.Length = 0, 1, DirInfo.GetFiles.Length) * 100.0)
                        ProcessedCount += 1
                        Application.DoEvents()
                        If (FilInfo.Extension.ToLower = ".aud") Then
                            makeAmr(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".amr", True)
                            If RadioButton5.Checked Then
                                convertAmr(path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0))
                                Try
                                    My.Computer.FileSystem.DeleteFile(path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".amr")
                                Catch exception7 As Exception
                                End Try
                            End If
                        End If
                    Next
                End If
                If (Friend_UsrNameMD5_Alias(CurrentFriendIndex) <> "" And Friend_UsrName(CurrentFriendIndex) <> Friend_UsrNameMD5_Alias(CurrentFriendIndex)) AndAlso Directory.Exists(DocumentsPath & "Audio\" & Friend_UsrNameMD5_Alias(CurrentFriendIndex)) Then
                    DirInfo = New DirectoryInfo(DocumentsPath & "Audio\" & Friend_UsrNameMD5_Alias(CurrentFriendIndex))
                    ProcessedCount = 0
                    'Dim FilInfo As FileInfo
                    For Each FilInfo In DirInfo.GetFiles
                        ProgressBar1.Value = CInt(CDbl(ProcessedCount) / IIf(DirInfo.GetFiles.Length = 0, 1, DirInfo.GetFiles.Length) * 100.0)
                        ProcessedCount += 1
                        Application.DoEvents()
                        If (FilInfo.Extension.ToLower = ".aud") Then
                            makeAmr(FilInfo.FullName, path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".amr", True)
                            If RadioButton5.Checked Then
                                convertAmr(path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0))
                                Try
                                    My.Computer.FileSystem.DeleteFile(path & Split(FilInfo.Name, ".", -1, CompareMethod.Binary)(0) & ".amr")
                                Catch exception7 As Exception
                                End Try
                            End If
                        End If
                    Next
                End If
            Next
            Dim writer2 As New StreamWriter(SavePath & "index.htm", False, Encoding.UTF8)
            writer2.WriteLine("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
            writer2.WriteLine("<html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title>微信聊天记录</title></head>")
            writer2.WriteLine("<body><table width=""210"" border=""0"" style=""font-size:12px;border-collapse:separate;border-spacing:0px 20px;word-break:break-all;table-layout:fixed;word-wrap:break-word;"" align=""center"">")
            For index = 0 To ChatMD5Length - 1
                Dim NumInChatRoom As Integer
                Label7.Text = "处理聊天记录 " & Trim(index + 1) & "/" & Trim(ChatMD5Length)
                Application.DoEvents()
                Dim IndexInChatMD5 As Integer = ChatMD5Index(index)
                Dim ChatRoomUsrName() As String = Nothing
                Dim DisplayNickname() As String = Nothing
                Dim IndexInFriend_UsrName() As Integer = Nothing
                Dim IsChatRoom As Boolean
                If Strings.Right(Friend_UsrName(IndexInChatMD5), 9).ToLower = "@chatroom" Then
                    IsChatRoom = True
                    ChatRoomUsrName = Split(FriendExt_ConChatRoomMem(IndexInChatMD5), ";", -1, CompareMethod.Binary)
                    NumInChatRoom = UBound(ChatRoomUsrName) + 1
                    ReDim IndexInFriend_UsrName(NumInChatRoom)
                    ReDim DisplayNickname(NumInChatRoom)
                    For j = 0 To NumInChatRoom - 1
                        IndexInFriend_UsrName(j) = -1
                        Dim k As Integer = 0
                        For k = 0 To FriendsCount - 1
                            If (Friend_UsrName(k) = ChatRoomUsrName(j)) Then
                                IndexInFriend_UsrName(j) = k
                                Exit For
                            End If
                        Next
                        If (IndexInFriend_UsrName(j) <> -1) Then
                            DisplayNickname(j) = Friend_NickName(IndexInFriend_UsrName(j))
                            If (FriendExt_ConRemark(IndexInFriend_UsrName(j)) <> "") Then
                                DisplayNickname(j) = FriendExt_ConRemark(IndexInFriend_UsrName(j))
                            End If
                        Else
                            DisplayNickname(j) = ChatRoomUsrName(j)
                        End If
                    Next
                End If
                Dim ChatTitle As String = Friend_NickName(IndexInChatMD5)
                If FriendExt_ConRemark(IndexInChatMD5) <> "" Then
                    ChatTitle = FriendExt_ConRemark(IndexInChatMD5)
                End If
                writer2.WriteLine("<tr><td width=""80"" align=""center""><a href=""" & Friend_UsrName(IndexInChatMD5) & "_1.htm""><img src=""Potrait\" & IIf(PicUsr(IndexInChatMD5) = "", "DefaultProfileHead@2x.png", PicUsr(IndexInChatMD5) & ".jpg") & """ width=""50"" height=""50""/></a></td><td width=""140"" align=""center""><a href=""" & Friend_UsrName(IndexInChatMD5) & "_1.htm"">" & ChatTitle & "</a></td></tr>")
                command.CommandText = "select count(*) from Chat_" & ChatMD5(index)
                Dim ChatLinesCount As Integer = CInt(command.ExecuteScalar())
                For n = 0 To (ChatLinesCount - 1) / 100
                    ProgressBar1.Value = CInt(Math.Round((CDbl(n) / ((ChatLinesCount - 1) / 100 + 2)) * 100))
                    Application.DoEvents()
                    Dim writer3 As New StreamWriter(SavePath & Friend_UsrName(IndexInChatMD5) & "_" & Trim(n + 1) & ".htm")
                    writer3.WriteLine("<!DOCTYPE html PUBLIC ""-//W3C//DTD XHTML 1.0 Transitional//EN"" ""http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"">")
                    writer3.WriteLine("<html xmlns=""http://www.w3.org/1999/xhtml""><head><meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" /><title>" & ChatTitle & " - 微信聊天记录 " & Trim(n + 1) & "/" & Trim(CInt((ChatLinesCount - 1) / 100 + 1)) & "</title></head>")
                    Dim NaviString As String = "<p style=""font-size:12px""><a href=""index.htm"">返回目录</a>&nbsp;&nbsp;"
                    If (n > 0) Then
                        NaviString = NaviString & "<a href=""" & Friend_UsrName(IndexInChatMD5) & "_1.htm"">第一页</a>&nbsp;<a href=""" & Friend_UsrName(IndexInChatMD5) & "_" & Trim(n) & ".htm"">上一页</a>&nbsp;"
                    End If
                    For i = 0 To CInt((ChatLinesCount - 1) / 100)
                        If (i <> n) Then
                            NaviString = NaviString & "<a href=""" & Friend_UsrName(IndexInChatMD5) & "_" & Trim(i + 1) & ".htm"">" & Trim(i + 1) & "</a>&nbsp;"
                        Else
                            NaviString = NaviString & "<strong>" & Trim(i + 1) & "</strong>&nbsp;"
                        End If
                    Next
                    If n < (ChatLinesCount - 1) / 100 Then
                        NaviString = NaviString & "<a href=""" & Friend_UsrName(IndexInChatMD5) & "_" & Trim(n + 2) & ".htm"">下一页</a>&nbsp;<a href=""" & Friend_UsrName(IndexInChatMD5) & "_" & Trim(CInt((ChatLinesCount - 1) / 100) + 1) & ".htm"">最后一页</a>&nbsp;"
                    End If
                    NaviString = NaviString & "</p>"
                    writer3.WriteLine(NaviString)
                    writer3.WriteLine("<table width=""600"" border=""0"" style=""font-size:12px;border-collapse:separate;border-spacing:0px 20px;word-break:break-all;table-layout:fixed;word-wrap:break-word;"" align=""center"">")
                    command.CommandText = "select MesLocalID,CreateTime,Message,Status,Des,Type from Chat_" & ChatMD5(index) & " order by CreateTime asc limit " & Trim(n * 100) & ",100"
                    reader = command.ExecuteReader
                    Do While reader.Read
                        Dim MesLocalID As Integer = reader.GetInt32(0)
                        Dim CreateTime As New DateTime(1970, 1, 1)
                        span = New TimeSpan(reader.GetInt64(1) * &H989680)
                        CreateTime = TimeZone.CurrentTimeZone.ToLocalTime(CreateTime).Add(span)
                        Dim Message As String = ""
                        If Not reader.IsDBNull(2) Then
                            Message = reader.GetString(2)
                        End If
                        If Message Is Nothing Then
                            Message = ""
                        End If
                        Dim MesssageBackup As String = Message
                        Dim Status As Integer = reader.GetInt32(3)
                        Dim Des As Integer = reader.GetInt32(4)
                        Dim MsgType As Integer = reader.GetInt32(5)
                        If InStr(Message, "<voicemsg ", CompareMethod.Text) <> 0 Then
                            If File.Exists(SavePath & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".amr") Then
                                Dim VoiceLength As Integer
                                If InStr(Message, "voicelength=""", CompareMethod.Text) <> 0 Then
                                    VoiceLength = CInt(Split(Split(Message, "voicelength=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Binary)(0))
                                End If
                                Message = "<object width=""100px"" height=""40px"" classid=""clsid:CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701""><param name=""FileName"" value=""" & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".amr"" /><param NAME=""AutoStart"" VALUE=""0""><embed type=""application/x-mplayer2"" width=""100px"" height=""40px"" autoplay=""false"" autostart=""false"" pluginspage=""http://www.microsoft.com/Windows/MediaPlayer/"" src=""" & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".amr"" /></object>"
                                If VoiceLength > 0 Then
                                    Message = Message & "&nbsp;" & Trim(CInt(VoiceLength / 1000) + 1) & "'"
                                End If
                            ElseIf File.Exists(SavePath & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".mp3") Then
                                Dim VoiceLength As Integer
                                If InStr(Message, "voicelength=""", CompareMethod.Text) <> 0 Then
                                    VoiceLength = CInt(Split(Split(Message, "voicelength=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Binary)(0))
                                End If
                                Message = "<audio controls preload=""none""><source src=""" & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".mp3"" type=""audio/mpeg""><object width=""100px"" height=""40px"" classid=""clsid:CLSID:22D6F312-B0F6-11D0-94AB-0080C74C7E95"" codebase=""http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701""><param name=""FileName"" value=""" & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".mp3"" /><param NAME=""AutoStart"" VALUE=""0""><embed type=""application/x-mplayer2"" width=""100px"" height=""40px"" autoplay=""false"" autostart=""false"" pluginspage=""http://www.microsoft.com/Windows/MediaPlayer/"" src=""" & Friend_UsrName(IndexInChatMD5) & "_Aud\" & Trim(MesLocalID) & ".mp3"" /></object></audio>"
                                If VoiceLength > 0 Then
                                    Message = Message & "&nbsp;" & Trim(CInt(VoiceLength / 1000) + 1) & "'"
                                End If
                            Else
                                Message = "[语音]"
                            End If
                        ElseIf InStr(Message, "<emoji ", CompareMethod.Text) <> 0 And InStr(Message, "md5=""", CompareMethod.Text) <> 0 Then
                            Dim tmpstring As String = Split(Split(Message, "md5=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Text)(0).Trim
                            If tmpstring <> "" Then
                                If File.Exists(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic") Then
                                    Message = "<img src=""Emotion\" & tmpstring & "." & copyImage(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic", SavePath & "Emotion\" & tmpstring) & """ width=""75px""/>"
                                ElseIf File.Exists((LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic..thumb")) Then
                                    Message = "<img src=""Emotion\" & tmpstring & "." & copyImage(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic..thumb", SavePath & "Emotion\" & tmpstring) & """ width=""75px""/>"
                                Else
                                    Message = "[表情]"
                                End If
                            End If
                        ElseIf InStr(Message, "<emoticonmd5>", CompareMethod.Text) <> 0 Then
                            Dim tmpstring As String = Split(Split(Message, "<emoticonmd5>", -1, CompareMethod.Text)(1), "</emoticonmd5>", -1, CompareMethod.Text)(0).Trim
                            If tmpstring <> "" Then
                                If File.Exists(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic") Then
                                    Message = "<img src=""Emotion\" & tmpstring & "." & copyImage(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic", SavePath & "Emotion\" & tmpstring) & """ width=""75px""/>"
                                ElseIf File.Exists(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic..thumb") Then
                                    Message = "<img src=""Emotion\" & tmpstring & "." & copyImage(LibraryPath & "WechatPrivate\emoticon1\" & tmpstring & ".pic..thumb", SavePath & "Emotion\" & tmpstring) & """ width=""75px""/>"
                                Else
                                    Message = "[表情]"
                                End If
                            ElseIf InStr(Message, "<title>", CompareMethod.Text) <> 0 Then
                                Try
                                    Dim tmptitle As String = Split(Split(Message, "<title>", -1, CompareMethod.Text)(1), "</title>", -1, CompareMethod.Text)(0)
                                    Dim tmpdes As String = Split(Split(Message, "<des>", -1, CompareMethod.Text)(1), "</des>", -1, CompareMethod.Text)(0)
                                    Dim tmpurl As String = Split(Split(Message, "<url>", -1, CompareMethod.Text)(1), "</url>", -1, CompareMethod.Text)(0)
                                    Message = "<a href=""" & tmpurl.Replace("&amp;", "&") & """><strong>" & tmptitle & "</strong><br/>" & tmpdes & "</a>"
                                Catch exception9 As Exception
                                    Message = "[链接]"
                                End Try
                            End If
                        ElseIf InStr(Message, "<img ", CompareMethod.Text) <> 0 Then
                            If File.Exists(SavePath & Friend_UsrName(IndexInChatMD5) & "_Img\" & Trim(MesLocalID) & "_thum.jpg") Then
                                If File.Exists(SavePath & Friend_UsrName(IndexInChatMD5) & "_Img\" & Trim(MesLocalID) & ".jpg") Then
                                    Message = "<a href=""" & Friend_UsrName(IndexInChatMD5) & "_Img\" & Trim(MesLocalID) & ".jpg""><img src=""" & Friend_UsrName(IndexInChatMD5) & "_Img\" & Trim(MesLocalID) & "_thum.jpg""/></a>"
                                Else
                                    Message = "<img src=""" & Friend_UsrName(IndexInChatMD5) & "_Img\" & Trim(MesLocalID) & "_thum.jpg""/>"
                                End If
                            Else
                                Message = "[图片]"
                            End If
                        ElseIf InStr(Message, "<location ", CompareMethod.Text) <> 0 Then
                            Try
                                Dim tmpx As String = Split(Split(Message, "x=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Text)(0)
                                Dim tmpy As String = Split(Split(Message, "y=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Text)(0)
                                Dim tmplabel As String = Split(Split(Message, "label=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Text)(0)
                                Message = "[位置] <strong>" & tmplabel & "</strong><br/>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;纬度" & tmpx & "&nbsp;经度" & tmpy
                            Catch exception10 As Exception
                                Message = "[位置]"
                            End Try
                        ElseIf InStr(Message, "nickname=""", CompareMethod.Text) <> 0 Then
                            Try
                                Message = "[名片] " & Split(Split(Message, "nickname=""", -1, CompareMethod.Text)(1), """", -1, CompareMethod.Binary)(0)
                            Catch exception11 As Exception
                                Message = "[名片]"
                            End Try
                        Else
                            Message = SafeHTML(Message)
                            For i = 1 To 105
                                Message = Message.Replace(ExpressionName(i), ExpressionHTML(i))
                            Next
                            Dim builder As New StringBuilder
                            For i = 0 To Message.Length - 1
                                If HasEmojiNum(AscW(Strings.Mid(Message, (i + 1), 1))) Then
                                    builder.Append(HandleEmoji(AscW(Strings.Mid(Message, i + 1, 1)), SavePath, ResBase))
                                Else
                                    builder.Append(Strings.Mid(Message, (i + 1), 1))
                                End If
                            Next
                            Message = builder.ToString
                            If (Message.Length > 1) Then
                                builder = New StringBuilder
                                Dim tasc As Integer = 0
                                For i = 1 To Message.Length - 1
                                    If (AscW(Strings.Mid(Message, i, 1)) And &HFC00) = &HD800 Then
                                        tasc = (AscW(Strings.Mid(Message, i, 1)) And &H3FF) << 10
                                        tasc = tasc + (Strings.AscW(Strings.Mid(Message, (i + 1), 1)) And &H3FF)
                                        tasc = tasc + &H10000
                                        If tasc > &H1FBD0 Or tasc < 0 Then
                                            tasc = 0
                                        End If
                                        If HasEmojiNum(tasc) Then
                                            builder.Append(HandleEmoji(tasc, SavePath, ResBase))
                                            i += 1
                                        Else
                                            builder.Append(Strings.Mid(Message, i, 1))
                                            If i = Message.Length - 1 Then
                                                builder.Append(Strings.Mid(Message, i + 1, 1))
                                            End If
                                        End If
                                    Else
                                        builder.Append(Strings.Mid(Message, i, 1))
                                        If i = Message.Length - 1 Then
                                            builder.Append(Strings.Mid(Message, i + 1, 1))
                                        End If
                                    End If
                                Next
                                Message = builder.ToString
                            End If
                        End If
                        If MsgType = 10000 Then
                            writer3.WriteLine("<tr><td width=""80"">&nbsp;</td><td width=""140"">&nbsp;</td><td>系统消息: " & Message & "</td></tr>")
                            Continue Do
                        End If
                        If Status = 2 Or Status = 3 Then
                            If OwnerPotrait = "" Then
                                writer3.WriteLine("<tr><td width=""80"" align=""center""><img src=""Potrait\DefaultProfileHead@2x.png"" width=""50"" height=""50"" /><br />" & TextBox2.Text & "</td>")
                            Else
                                writer3.WriteLine("<tr><td width=""80"" align=""center""><a href=""Potrait\" & IIf(OwnerPotraitHD = "", OwnerPotrait, OwnerPotraitHD & "_hd") & ".jpg""><img src=""Potrait\" & OwnerPotrait & ".jpg"" width=""50"" height=""50"" /></a><br />" & TextBox2.Text & "</td>")
                            End If
                            writer3.WriteLine("<td width=""140"" align=""center"">" & CreateTime.ToString & "</td>")
                            writer3.WriteLine("<td>" & Message & "</td></tr>")
                            Continue Do
                        End If
                        If Status = 4 Then
                            If IsChatRoom Then
                                Dim FoundUsrName As Boolean
                                Dim tmppos As Integer = InStr(MesssageBackup, ":", CompareMethod.Binary)
                                If tmppos > 1 Then
                                    Dim tmpUsrName As String = Strings.Left(MesssageBackup, tmppos - 1)
                                    For m = 0 To NumInChatRoom - 1
                                        If ChatRoomUsrName(m) = tmpUsrName And IndexInFriend_UsrName(m) <> -1 Then
                                            If PicUsr(IndexInFriend_UsrName(m)) = "" Then
                                                writer3.WriteLine("<tr><td width=""80"" align=""center""><img src=""Potrait\DefaultProfileHead@2x.png"" width=""50"" height=""50"" /><br />" & DisplayNickname(m) & "</td>")
                                            Else
                                                writer3.WriteLine("<tr><td width=""80"" align=""center""><a href=""Potrait\" & IIf(PicHD(IndexInFriend_UsrName(m)) = "", PicUsr(IndexInFriend_UsrName(m)), PicHD(IndexInFriend_UsrName(m)) & "_hd") & ".jpg""><img src=""Potrait\" & PicUsr(IndexInFriend_UsrName(m)) & ".jpg"" width=""50"" height=""50"" /></a><br />" & DisplayNickname(m) & "</td>")
                                            End If
                                            writer3.WriteLine("<td width=""140"" align=""center"">" & CreateTime.ToString & "</td>")
                                            writer3.WriteLine("<td>" & IIf(Not Message.StartsWith("<"), MyTrim(Strings.Mid(Message, tmppos + 1)), Message) & "</td></tr>")
                                            FoundUsrName = True
                                            Exit Do
                                        End If
                                    Next
                                    If Not FoundUsrName Then
                                        writer3.WriteLine("<tr><td width=""80"" align=""center""><img src=""Potrait\DefaultProfileHead@2x.png"" width=""50"" height=""50"" /><br />" & tmpUsrName & "</td>")
                                        writer3.WriteLine("<td width=""140"" align=""center"">" & CreateTime.ToString & "</td>")
                                        writer3.WriteLine("<td>", IIf(Not Message.StartsWith("<"), MyTrim(Strings.Mid(Message, tmppos + 1)), Message) & "</td></tr>")
                                        FoundUsrName = True
                                    End If
                                End If
                                If Not FoundUsrName Then
                                    writer3.WriteLine("<tr><td width=""80"">&nbsp;</td>")
                                    writer3.WriteLine("<td width=""140"" align=""center"">" & CreateTime.ToString & "</td>")
                                    writer3.WriteLine("<td>" & Message & "</td></tr>")
                                End If
                                Continue Do
                            End If
                            If PicUsr(IndexInChatMD5) = "" Then
                                writer3.WriteLine("<tr><td width=""80"" align=""center""><img src=""Potrait\DefaultProfileHead@2x.png"" width=""50"" height=""50"" /><br />" & ChatTitle & "</td>")
                            Else
                                writer3.WriteLine("<tr><td width=""80"" align=""center""><a href=""Potrait\" & IIf(PicHD(IndexInChatMD5) = "", PicUsr(IndexInChatMD5), PicHD(IndexInChatMD5) & "_hd") & ".jpg""><img src=""Potrait\" & PicUsr(IndexInChatMD5) & ".jpg"" width=""50"" height=""50"" /></a><br />" & ChatTitle & "</td>")
                            End If
                            writer3.WriteLine("<td width=""140"" align=""center"">" & CreateTime.ToString & "</td>")
                            writer3.WriteLine("<td>" & Message & "</td></tr>")
                            Continue Do
                        End If
                    Loop
                    reader.Close()
                    writer3.WriteLine("</table>")
                    writer3.WriteLine(NaviString)
                    writer3.WriteLine("</body></html>")
                    writer3.Close()
                Next
            Next
            writer2.WriteLine("</table></body></html>")
            writer2.Close()
        End If
        connection.Close()
        MsgBox("保存完成，文件在程序目录下的" & NowDateString & "文件夹中")
        Button1.Enabled = True
        Button2.Enabled = True
        GroupBox1.Enabled = True
        If Not RadioButton3.Checked Then
            GroupBox2.Enabled = True
        End If
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        ComboBox1.Enabled = True
        Label7.Text = "这里是显示进度的"
        ProgressBar1.Value = 100
        If Not connection Is Nothing Then
            connection.Dispose()
        End If
        connection = Nothing
    End Sub


    Public Sub makeAmr(ByVal source As String, ByVal dest As String, ByVal overwrite As Boolean)
        Dim stream As New FileStream(source, FileMode.Open)
        Dim tarray(6) As Byte
        stream.Read(tarray, 0, 6)
        stream.Close()
        If tarray(0) = &H23 And tarray(1) = &H21 And tarray(2) = &H41 And tarray(3) = &H4D And tarray(4) = &H52 And tarray(5) = 10 Then
            My.Computer.FileSystem.CopyFile(source, dest, True)
        Else
            Dim stream2 As New FileStream(dest, FileMode.OpenOrCreate)
            stream = New FileStream(source, FileMode.Open)
            stream2.WriteByte(&H23)
            stream2.WriteByte(&H21)
            stream2.WriteByte(&H41)
            stream2.WriteByte(&H4D)
            stream2.WriteByte(&H52)
            stream2.WriteByte(10)
            ReDim tarray(stream.Length)
            stream.Read(tarray, 0, stream.Length)
            stream2.Write(tarray, 0, stream.Length)
            stream.Close()
            stream2.Close()
        End If
    End Sub

    Public Sub convertAmr(ByVal name As String)
        Dim startInfo As New ProcessStartInfo("ffmpeg.exe", "-y -i """ & name & ".amr"" -acodec libmp3lame -ab 32k """ & name & ".mp3""")
        startInfo.WindowStyle = ProcessWindowStyle.Hidden
        Process.Start(startInfo).WaitForExit()
    End Sub

    Public Function copyImage(ByVal source As String, ByVal dest As String) As String
        Dim stream As New FileStream(source, FileMode.Open)
        Dim num As Byte = CByte(stream.ReadByte)
        stream.Close()
        Select Case num
            Case &H47
                If Not File.Exists(dest & ".gif") Then
                    My.Computer.FileSystem.CopyFile(source, dest & ".gif")
                End If
                Return "gif"
            Case &H89
                If Not File.Exists(dest & ".png") Then
                    My.Computer.FileSystem.CopyFile(source, dest & ".png")
                End If
                Return "png"
            Case &HFF
                If Not File.Exists(dest & ".jpg") Then
                    My.Computer.FileSystem.CopyFile(source, dest & ".jpg")
                End If
                Return "gif"
            Case Else
                If Not File.Exists(dest & ".pic") Then
                    My.Computer.FileSystem.CopyFile(source, dest & ".pic")
                End If
                Return "pic"
        End Select
    End Function

    Public Function GoodName(ByVal s As String) As String
        s = s.Replace("\", "")
        s = s.Replace("/", "")
        s = s.Replace(":", "")
        s = s.Replace("*", "")
        s = s.Replace("?", "")
        s = s.Replace("""", "")
        s = s.Replace("<", "")
        s = s.Replace(">", "")
        s = s.Replace("|", "")
        Return s
    End Function

    Public Function SafeHTML(ByVal s As String) As String
        s = s.Replace("&", "&amp;")
        s = s.Replace(" ", "&nbsp;")
        s = s.Replace("<", "&lt;")
        s = s.Replace(">", "&gt;")
        s = s.Replace(ChrW(13) & ChrW(10), "<br/>")
        s = s.Replace(ChrW(13), "<br/>")
        s = s.Replace(ChrW(10), "<br/>")
        Return s
    End Function

    Public Function HandleEmoji(ByVal tc As Integer, ByRef saveBase As String, ByRef ResBase As String) As String
        Dim EmojiNum As String = Trim(tc)
        If Not File.Exists(saveBase & "Emoji\" & EmojiNum & ".png") Then
            My.Computer.FileSystem.CopyFile(ResBase & "Emoji\" & EmojiNum & ".png", saveBase & "Emoji\" & EmojiNum & ".png")
        End If
        Return "<img src=""Emoji\" & EmojiNum & ".png"" width=""18px"" height=""18px"" />"
    End Function

    Public Function MD5(ByVal strSource As String, Optional ByVal Code As Short = 32) As String
        Dim bytes As Byte() = New ASCIIEncoding().GetBytes(strSource)
        Dim buffer2 As Byte() = DirectCast(CryptoConfig.CreateFromName("MD5"), HashAlgorithm).ComputeHash(bytes)
        Dim str As String = ""
        Select Case Code
            Case 16
                For num = 4 To 11
                    str = str & Conversion.Hex(buffer2(num)).PadLeft(2, "0"c).ToLower
                Next
                Return str
            Case Else
                For num = 0 To 15
                    str = str & Conversion.Hex(buffer2(num)).PadLeft(2, "0"c).ToLower
                Next
                Return str
        End Select
    End Function

    Public Function MyTrim(ByVal ChatTitle As String) As String
        If ChatTitle.Length >= 2 AndAlso (Strings.Left(ChatTitle, 1) = ChrW(13) Or Strings.Left(ChatTitle, 1) = ChrW(10)) Then
            Return Strings.Mid(ChatTitle, 2)
        End If
        Return ChatTitle
    End Function

    Public Function safeName(ByVal name As String) As String
        Return Strings.Replace(name, "*", "", 1, -1, CompareMethod.Binary)
    End Function

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        GroupBox2.Enabled = True
    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        GroupBox2.Enabled = True
    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged
        GroupBox2.Enabled = False
    End Sub
End Class
