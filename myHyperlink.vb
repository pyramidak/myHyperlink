#Region " Hyperlink "

Class myLink

    Public Shared Sub Start(link As String)
        If My.Computer.Network.IsAvailable = False Then
            MessageBox.Show("No internet connection.", "open link")
            Exit Sub
        End If
        Try
            Process.Start(link)
        Catch ex As Exception
            MessageBox.Show(ex.Message, "open link")
        End Try
    End Sub

    Public Shared Sub StartHidden(link As String)
        Try
            If Not link.ToLower.StartsWith("http") Then link = "http://" & link
            Dim wResponse = CreateConnection(link, 1000)
            If wResponse IsNot Nothing Then wResponse.Close()
        Catch
        End Try
    End Sub

    Private Shared Function CreateConnection(link As String, Optional timeout As Integer = 500) As Net.HttpWebResponse
        'protocol activation Tls 1.2 needed from Framework 4.5
        Try
            Net.ServicePointManager.SecurityProtocol = CType(3072, Net.SecurityProtocolType)
            Dim wRequest As Net.HttpWebRequest = CType(Net.WebRequest.Create(link), Net.HttpWebRequest)
            wRequest.UserAgent = Application.ExeName
            wRequest.Timeout = timeout
            Dim wResponse As Net.HttpWebResponse = CType(wRequest.GetResponse(), Net.HttpWebResponse)
            Return wResponse
        Catch
            Return Nothing
        End Try
    End Function

    Public Shared Function AbsoluteUri(link As String) As String
        If link Is Nothing OrElse link = "" Then Return ""
        If My.Computer.Network.IsAvailable = False Then Return ""

        If Not link.ToLower.StartsWith("http") Then link = "http://" & link
        AbsoluteUri = ""
        Dim wResponse = CreateConnection(link, 250)
        If wResponse IsNot Nothing Then
            AbsoluteUri = wResponse.ResponseUri.AbsoluteUri().ToString
            wResponse.Close()
        End If
    End Function

    Public Shared Function Exist(link As String) As Boolean
        If link Is Nothing OrElse link = "" Then Return False
        If My.Computer.Network.IsAvailable = False Then Return False

        If Not link.ToLower.StartsWith("http") Then link = "http://" & link
        Exist = False
        Dim wResponse = CreateConnection(link, 250)
        If wResponse IsNot Nothing Then
            Exist = True
            wResponse.Close()
        End If
    End Function

    Public Shared Function Address(link As String) As Boolean
        If link Is Nothing OrElse link = "" Then Return False
        If link.ToLower.StartsWith("https://") Or link.ToLower.StartsWith("http://") Or link.ToLower.StartsWith("www.") Then Return True
        Return False
    End Function

    Public Shared Function WebIcon(link As String) As System.Drawing.Icon
        If link Is Nothing OrElse link = "" Then Return Nothing
        If My.Computer.Network.IsAvailable = False Then Return Nothing

        If Not link.ToLower.StartsWith("http") Then link = "https://" & link
        Dim url As Uri = New Uri(link)
        If url.HostNameType = UriHostNameType.Dns Then
            Dim iconURL = If(link.StartsWith("https"), "https://", "http://") & url.Host & "/favicon.ico"
            Try
                Dim wResponse = CreateConnection(iconURL, 250)
                If wResponse IsNot Nothing Then
                    Dim stream As IO.Stream = wResponse.GetResponseStream()
                    Dim favicon As System.Drawing.Image = System.Drawing.Image.FromStream(stream)
                    wResponse.Close()
                    Dim iconBitmap As System.Drawing.Bitmap = New System.Drawing.Bitmap(favicon)
                    Return System.Drawing.Icon.FromHandle(iconBitmap.GetHicon)
                End If
            Catch
            End Try
        End If

        Return Nothing
    End Function

    Public Shared Function WebName(link As String) As String
        If link Is Nothing OrElse link = "" Then Return ""

        If My.Computer.Network.IsAvailable Then
            If Not link.ToLower.StartsWith("http") Then link = "https://" & link
            Dim url As Uri = New Uri(link)
            If url.HostNameType = UriHostNameType.Dns Then link = url.Host
        End If

        Dim a As Integer = link.IndexOf(".")
        If a = -1 Then
            WebName = link
        Else
            Dim b As Integer = link.IndexOf(".", a + 1)
            If b = -1 Then
                WebName = link.Substring(0, a)
            Else
                WebName = link.Substring(a + 1, b - a - 1)
            End If
        End If
        Return UCase(WebName.Substring(0, 1)) & WebName.Substring(1, WebName.Length - 1)
    End Function

    Private Shared Function RemoveLastSlash(ByVal Link As String) As String
        If Link Is Nothing OrElse Link = "" Then Return ""
        If Link.EndsWith("/") Then
            Return Link.Substring(0, Link.Length - 1)
        Else
            Return Link
        End If
    End Function

End Class

#End Region