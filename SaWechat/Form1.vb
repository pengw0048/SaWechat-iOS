Imports System.IO
Imports System.Text
Imports System.Security.Cryptography

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

    Public Function handleEmoji(ByVal tc As Integer, ByRef saveBase As String, ByRef resBase As String) As String
        Dim str2 As String = Trim(tc)
        If Not File.Exists(saveBase & "Emoji\" & str2 & ".png") Then
            My.Computer.FileSystem.CopyFile(resBase & "Emoji\" & str2 & ".png", saveBase & "Emoji\" & str2 & ".png")
        End If
        Return "<img src=""Emoji\" & str2 & ".png"" width=""18px"" height=""18px"" />"
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

    Public Function MyTrim(ByVal s As String) As String
        If s.Length >= 2 AndAlso (Strings.Left(s, 1) = ChrW(13) Or Strings.Left(s, 1) = ChrW(10)) Then
            Return Strings.Mid(s, 2)
        End If
        Return s
    End Function

End Class
