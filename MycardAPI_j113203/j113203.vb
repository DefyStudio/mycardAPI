Public Class j113203

    Private Sub DocumentCompleted(sender As Object, e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs)
        Try
            If e.Url.AbsoluteUri.ToLower Like "https://member.mycard520.com.tw/MemberLoginService/*".ToLower Then ''儲值帳號確認 3
                sender.Document.GetElementById("ctl00_ContentPlaceHolder1_btnLogin").InvokeMember("click")
            ElseIf e.Url.AbsoluteUri.ToLower Like "https://www.mycard520.com.tw/web5/Redeem/Choice/Game*".ToLower Then ''儲值方法 , 2 | 2
                sender.Document.GetElementById("gamechoiced").SetAttribute("value", "MYPT|7")
                For Each CurElement As System.Windows.Forms.HtmlElement In sender.Document.Forms(0).GetElementsByTagName("input")
                    If CurElement.GetAttribute("className") = "normal_blue_btn" Then CurElement.InvokeMember("click")
                Next
            ElseIf e.Url.AbsoluteUri.ToLower Like "https://www.mycard520.com.tw/web5/Redeem/RMessage*".ToLower Then ''儲值信息 , 2 | 1
                If sender.Document.GetElementById("s4_p") IsNot Nothing Then
                    回傳 = sender.Document.GetElementById("s4_p").InnerText + vbNewLine + 儲值額
                Else
                    For Each CurElement As System.Windows.Forms.HtmlElement In sender.Document.GetElementsByTagName("ul")
                        If CurElement.GetAttribute("className") = "mcp_result" Then 回傳 = CurElement.InnerText
                    Next
                End If
            ElseIf e.Url.AbsoluteUri.ToLower Like "https://www.mycard520.com.tw/web5/Redeem*".ToLower Then ''儲值 , 1
                sender.Document.GetElementById("MyCardId").SetAttribute("value", MyCard卡號)
                sender.Document.GetElementById("MyCardPwd").SetAttribute("value", MyCard密碼)
                sender.Document.GetElementById("freemcardId").SetAttribute("value", 免費MyCard帳號)
                For Each CurElement As System.Windows.Forms.HtmlElement In sender.Document.Forms(0).GetElementsByTagName("input")
                    If CurElement.GetAttribute("className") = "normal_blue_btn" Then CurElement.InvokeMember("click")
                Next
            ElseIf e.Url.AbsoluteUri.ToLower Like "https://member.mycard520.com.tw/Login/MyCardMemberLogin.aspx*".ToLower Then ''登入
                sender.Document.GetElementById("TextBox1").SetAttribute("value", MyCard會員帳號)
                sender.Document.GetElementById("TextBox2").SetAttribute("value", MyCard會員密碼)
                sender.Document.GetElementById("UpdatePanel1").InnerHtml = ""
                sender.Document.GetElementById("Button1").InvokeMember("click")
            End If
        Catch ex As Exception
        End Try
    End Sub
    Public Function 智冠儲值_儲值(MyCard卡號_ As String, MyCard密碼_ As String, 帳號 As String, 密碼 As String) As String
        Dim 點數 As Newtonsoft.Json.Linq.JObject = Newtonsoft.Json.Linq.JObject.Parse(智冠儲值_快速驗證(MyCard卡號_, MyCard密碼_))
        If Not String.IsNullOrEmpty(點數("ReturnMsg")) Then
            MyCard卡號 = MyCard卡號_
            MyCard密碼 = MyCard密碼_
            MyCard會員帳號 = 帳號
            MyCard會員密碼 = 密碼
            回傳 = ""
            儲值額 = 點數("ReturnMsg").ToString.Substring(7, 點數("ReturnMsg").ToString.LastIndexOf("點") - 7)
            RealBrowser.ScriptErrorsSuppressed = True
            AddHandler RealBrowser.DocumentCompleted, AddressOf DocumentCompleted
            RealBrowser.Navigate("https://www.mycard520.com.tw/web5/Redeem")
            Do While String.IsNullOrEmpty(回傳)
                System.Windows.Forms.Application.DoEvents()
            Loop
            Return 回傳
        End If
        Return "親愛的客戶您好，由於以下原因，請您重新再試。" + vbNewLine + "卡號或密碼錯誤,請確認後重新輸入!" + vbNewLine + "若您還有任何問題，請您連聯絡MyCard客服專員。" + vbNewLine + "MyCard客服電話：(02)2651-0754"
    End Function
    Public Function 智冠儲值_快速驗證(MyCard卡號 As String, MyCard密碼 As String) As String
        Dim 智冠儲值 As New Specialized.NameValueCollection
        智冠儲值.Add("MyID", MyCard卡號)
        智冠儲值.Add("MyPW", MyCard密碼)
        智冠儲值.Add("FactoryId", "free-mycard")
        智冠儲值.Add("FreeMyCardId", "91738490")
        Using 快速驗證 As New System.Net.WebClient()
            快速驗證.Encoding = System.Text.Encoding.UTF8
            Return System.Text.Encoding.UTF8.GetString(快速驗證.UploadValues("https://www.mycard520.com.tw/MyCardBilllingCheckCoupon/default/FreeMyCardQuery", 智冠儲值))
        End Using
    End Function
    Dim MyCard卡號, MyCard密碼, 免費MyCard帳號, 回傳, 儲值額, MyCard會員帳號, MyCard會員密碼 As String, RealBrowser As System.Windows.Forms.WebBrowser
End Class
