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
            Dim info As New DirectoryInfo(rightpath & "Documents")
            Dim info2 As DirectoryInfo
            For Each info2 In info.GetDirectories
                If info2.Name.Length = 32 And info2.Name.Replace("0", "") <> "" Then
                    ComboBox1.Items.Add(info2.Name)
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
        Dim PicHd(FriendsLoopCount) As String
        Dim OwnerPotrait As String = ""
        Dim text As String = ""
        '这里在做什么，还需调查
        If File.Exists(DocumentsPath & "Usr\" & ComboBox1.Text & ".pic_usr") Then
            OwnerPotrait = ComboBox1.Text
        End If
        If File.Exists(DocumentsPath & "Usr\" & ComboBox1.Text & ".pic_hd") Then
            OwnerPotrait = ComboBox1.Text
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
            If File.Exists((DocumentsPath & "Usr\" & Friend_UsrNameMD5_Alias(index) & ".pic_usr")) Then
                PicUsr(index) = Friend_UsrNameMD5_Alias(index)
            End If
            If File.Exists((DocumentsPath & "Usr\" & Friend_UsrNameMD5(index) & ".pic_usr")) Then
                PicUsr(index) = Friend_UsrNameMD5(index)
            End If
            If File.Exists((DocumentsPath & "Usr\" & Friend_UsrNameMD5_Alias(index) & ".pic_hd")) Then
                PicHd(index) = Friend_UsrNameMD5_Alias(index)
            End If
            If File.Exists((DocumentsPath & "Usr\" & Friend_UsrNameMD5(index) & ".pic_hd")) Then
                PicHd(index) = Friend_UsrNameMD5(index)
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
                Dim ChatRoomUsrName() As String
                Dim DisplayNickname() As String
                Dim IndexInFriend_UsrName() As Integer
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
                Dim num44 As Integer = (((ChatLinesCount - 1) / 100) + 1)
                For i = 0 To (ChatLinesCount - 1) / 100 + 1
                    ProgressBar1.Value = CInt(Math.Round((CDbl(i) / (ChatLinesCount - 1) / 100 + 2) * 100))
                    Application.DoEvents()
                    command.CommandText = "select MesLocalID,CreateTime,Message,Status,Des,Type from Chat_" & ChatMD5(index) & " order by CreateTime asc limit " & Trim(CInt(i * 100)) & ",100"
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
            'radiobutton2?
        End If
        connection.Close()
        Interaction.MsgBox("保存完成，文件在程序目录下的" & NowDateString & "文件夹中")
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

    Public Function GoodName(ByVal ChatTitle As String) As String
        ChatTitle = ChatTitle.Replace("\", "")
        ChatTitle = ChatTitle.Replace("/", "")
        ChatTitle = ChatTitle.Replace(":", "")
        ChatTitle = ChatTitle.Replace("*", "")
        ChatTitle = ChatTitle.Replace("?", "")
        ChatTitle = ChatTitle.Replace("""", "")
        ChatTitle = ChatTitle.Replace("<", "")
        ChatTitle = ChatTitle.Replace(">", "")
        ChatTitle = ChatTitle.Replace("|", "")
        Return ChatTitle
    End Function

    Public Function handleEmoji(ByVal tc As Integer, ByRef saveBase As String, ByRef ResBase As String) As String
        Dim LibraryPath As String = Trim(tc)
        If Not File.Exists(saveBase & "Emoji\" & LibraryPath & ".png") Then
            My.Computer.FileSystem.CopyFile(ResBase & "Emoji\" & LibraryPath & ".png", saveBase & "Emoji\" & LibraryPath & ".png")
        End If
        Return "<img src=""Emoji\" & LibraryPath & ".png"" width=""18px"" height=""18px"" />"
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
