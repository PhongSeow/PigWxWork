Imports PigWxWorkLib

Public Class ConsoleDemo
    Public WorkApp As WorkApp
    Public CorpId As String = "ww18330c6627358859"
    Public CorpSecret As String = "gS-NlfNPj_qf1IMBrKrDtcAuZMw9hoZzf9lk-oLrD8E"

    Public Sub Main()
        Dim strRet As String = ""
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to New WorkApp")
            Console.WriteLine("Press B to WorkApp.GetOauth2Url")
            Console.WriteLine("Press C to WorkApp.GetApiIpList")
            Console.WriteLine("Press D to Test Regex")
            Console.WriteLine("Press E to WorkApp.SendWorkMessage")
            Console.WriteLine("*******************")
            Console.CursorVisible = False
            Select Case Console.ReadKey(True).Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("New WorkApp")
                    Console.WriteLine("*******************")
                    Console.CursorVisible = True
                    Console.WriteLine("Input corpid:" & Me.CorpId)
                    Dim strCorpid As String = Console.ReadLine()
                    If strCorpid = "" Then strCorpid = Me.CorpId
                    Console.WriteLine("Input corpsecret:" & Me.CorpSecret)
                    Dim strCorpSecret As String = Console.ReadLine()
                    If strCorpSecret = "" Then strCorpSecret = Me.CorpSecret
                    Console.Write("New WorkApp...")
                    Me.WorkApp = New WorkApp(strCorpid, strCorpSecret)
                    With Me.WorkApp
                        If .LastErr <> "" Then
                            Console.WriteLine(.LastErr)
                        Else
                            Console.WriteLine("OK")
                        End If
                    End With
                Case ConsoleKey.B
                    Console.WriteLine("*******************")
                    Console.WriteLine("WorkApp.GetOauth2Url")
                    Console.WriteLine("*******************")
                    If Me.WorkApp Is Nothing Then
                        Console.Write("Me.WorkApp Is Nothing")
                    Else
                        With Me.WorkApp
                            Console.Write("GetOauth2Url...")
                            Dim strUrl As String = .GetOauth2Url("http://aaa/", "ABC")
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                Console.WriteLine("Oauth2Url=" & strUrl)
                            End If
                        End With
                    End If
                Case ConsoleKey.C
                    Console.WriteLine("*******************")
                    Console.WriteLine("WorkApp.GetApiIpList")
                    Console.WriteLine("*******************")
                    If Me.WorkApp Is Nothing Then
                        Console.Write("Me.WorkApp Is Nothing")
                    Else
                        With Me.WorkApp
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                            End If
                            Console.Write("GetApiIpList...")
                            Dim astrIpList As String() = .GetApiIpList()
                            If .LastErr <> "" Then
                                Console.WriteLine(.LastErr)
                            Else
                                Console.WriteLine("OK")
                                For i = 0 To astrIpList.Length - 1
                                    Console.WriteLine(astrIpList(i))
                                Next
                            End If
                        End With
                    End If
 '               Case ConsoleKey.D
                    'Console.WriteLine("Url:")
                    'Dim strUrl As String = Console.ReadLine
                    'Dim oPigFunc As New PigToolsLib.PigFunc
                    'Console.WriteLine(oPigFunc.UrlEncode(strUrl))
                Case ConsoleKey.E
                    Console.WriteLine("*******************")
                    Console.WriteLine("WorkApp.SendWorkMessage")
                    Console.WriteLine("*******************")
                    Dim oWorkMessage As New WorkMessage("1000037")  '广东龙信
                    Dim strContent As String = "广东龙信企业微信测试" & vbCrLf & "中华人民共和国\n广东省"
                    With oWorkMessage
                        .MsgType = WorkMessage.MsgTypeEnum.Text
                        .ToUser = "18022310718"
                        .Content = strContent
                    End With
                    With Me.WorkApp
                        strRet = .SendWorkMessage(oWorkMessage)
                        Console.WriteLine(strRet)
                    End With
                    ''Dim oPigJSon As New PigJSon("{""errcode"":0,""errmsg"":""ok"",""access_token"":""ta_KKc1U0cOlLWkX-KlHd8Nw9J5-fMChJI4ILMykX7btfOYzwO0Mkx2XOrS1zLSFdiqSc3GIRYfaJcDLSt0SXsBiT7ufiZutqnsykfHmvINs7yaRKhO5nU68qyUw87donK9FrV7l9WO6OD9jyCPrQmC6d6fAeF3dW3J-nx2Pl14feUAI9vYqiXe0rgJPHxCppJWhEmvUJ41bAq2g2dBNvw"",""expires_in"":7200}")
                    'Dim oPigJSon As New fPigJSon("{    ""ip_list"":[        ""182.254.11.176"",        ""182.254.78.66""    ],    ""errcode"":0,    ""errmsg"":""ok""}")
                    'Console.WriteLine(oPigJSon.LastErr)
                    'Console.WriteLine(oPigJSon.GetStrValue("ip_list[0]"))
                    'Console.WriteLine(oPigJSon.GetStrValue("ip_list.length"))
                    '                Case ConsoleKey.F
                    ''Dim oPigJSon As New PigJSon("{""errcode"":0,""errmsg"":""ok"",""access_token"":""ta_KKc1U0cOlLWkX-KlHd8Nw9J5-fMChJI4ILMykX7btfOYzwO0Mkx2XOrS1zLSFdiqSc3GIRYfaJcDLSt0SXsBiT7ufiZutqnsykfHmvINs7yaRKhO5nU68qyUw87donK9FrV7l9WO6OD9jyCPrQmC6d6fAeF3dW3J-nx2Pl14feUAI9vYqiXe0rgJPHxCppJWhEmvUJ41bAq2g2dBNvw"",""expires_in"":7200}")
                    'Dim oPigJSon As New fPigJSon("{    ""ip_list"":[        ""182.254.11.176"",        ""182.254.78.66""    ],    ""errcode"":0,    ""errmsg"":""ok""}")
                    'Console.WriteLine(oPigJSon.LastErr)
                    'Console.WriteLine(oPigJSon.GetStrValue("ip_list[0]"))
                    'Console.WriteLine(oPigJSon.GetStrValue("ip_list.length"))
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
