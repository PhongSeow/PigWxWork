'**********************************
'* Name: WorkApp
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 企业微信应用类
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 23/2/2021
'* 1.0.2  25/2/2021   Add GetErrMsg, Modify mNew
'* 1.0.3  1/3/2021   Add RefAccessToken,Oauth2
'* 1.0.4  5/3/2021   Add GetApiIpList
'* 1.0.5  6/3/2021   Add GetUserIdentity
'* 1.0.6  7/3/2021   Add MailList_ReadMember,GetWorkMemberFromOauth2Redirect
'* 1.0.7  8/3/2021   Add AppPath,AppTitle
'* 1.0.8  8/3/2021   Modify GetApiIpList
'* 1.0.9  10/3/2021   Modify AccessToken
'* 1.0.10  15/4/2021  Use PigToolsWinLib
'* 1.0.11  17/5/2021  Modify _New
'* 1.0.12  14/7/2021  Modify GetWorkMemberFromOauth2Redirect,GetUserIdentity
'* 1.0.13  15/7/2021  Modify GetUserIdentity
'* 1.0.15  25/8/2021  Remove Imports PigToolsLib
'* 1.1  29/8/2021	Use PigToolsLiteLib
'* 1.2  13/9/2021	Add  mWebReqGetText,
'**********************************

Imports System.Web
Imports System.Web.Services
Imports PigToolsWinLib
Imports PigToolsLiteLib


Public Class WorkApp
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.2.2"
	Private Const QYAPI_URL As String = "https://qyapi.weixin.qq.com"
	Private Const QYAPI_CGIBIN_URL As String = QYAPI_URL & "/cgi-bin"
	Private Const OEPN_WX_URL As String = "https://open.weixin.qq.com"
	Private Const OEPN_WX_OAUTH2_URL As String = OEPN_WX_URL & "/connect/oauth2/authorize"
	Private moPigFunc As New PigFunc

	''' <summary>
	''' 
	''' </summary>
	''' <param name="CorpId">企业ID</param>
	''' <param name="CorpSecret">应用的凭证密钥</param>
	Public Sub New(CorpId As String, CorpSecret As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(CorpId, CorpSecret)
	End Sub

	''' <summary>
	''' 
	''' </summary>
	''' <param name="CorpId">企业ID</param>
	''' <param name="CorpSecret">应用的凭证密钥</param>
	''' <param name="AccessToken">缓存的access_token</param>
	''' <param name="ExpiresTime">缓存的access_token过期时间</param>
	Public Sub New(CorpId As String, CorpSecret As String, AccessToken As String, ExpiresTime As Date)
		MyBase.New(CLS_VERSION)
		Me.mNew(CorpId, CorpSecret, AccessToken, ExpiresTime)
	End Sub

	Public Shadows ReadOnly Property IsDebug() As Boolean
		Get
			Return MyBase.IsDebug
		End Get
	End Property

	Public Shadows ReadOnly Property AppPath() As String
		Get
			Return MyBase.AppPath
		End Get
	End Property

	Public Shadows ReadOnly Property AppTitle() As String
		Get
			Return MyBase.AppTitle
		End Get
	End Property

	'Public Shadows Sub SetDebug(DebugFilePath As String)
	'	MyBase.SetDebug(DebugFilePath)
	'End Sub

	'Public Shadows Sub SetDebug(DebugFilePath As String, IsHardDebug As Boolean)
	'	MyBase.SetDebug(DebugFilePath, IsHardDebug)
	'End Sub

	Private Sub mNew(CorpId As String, CorpSecret As String, Optional AccessToken As String = "", Optional ExpiresTime As DateTime = Nothing)
		Try
			mstrCorpId = CorpId
			mstrCorpSecret = CorpSecret
			mstrAccessToken = AccessToken
			If IsNothing(ExpiresTime) Then
				mdteAccessTokenExpiresTime = New DateTime(1970, 1, 1)
			Else
				mdteAccessTokenExpiresTime = ExpiresTime
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("mNew", ex)
		End Try
	End Sub

	''' <summary>
	''' ACCESS_TOKEN是否就绪
	''' </summary>
	Public ReadOnly Property IsAccessTokenReady() As Boolean
		Get
			If mstrAccessToken = "" Or Me.AccessTokenExpiresTime < Now Then
				Return False
			Else
				Return True
			End If
		End Get
	End Property


	''' <summary>
	''' 凭证的过期时间
	''' </summary>
	Private mdteAccessTokenExpiresTime As DateTime
	Public ReadOnly Property AccessTokenExpiresTime() As DateTime
		Get
			Return mdteAccessTokenExpiresTime
		End Get
	End Property

	''' <summary>
	''' 企业ID
	''' </summary>
	Private mstrCorpId As String
	Public ReadOnly Property CorpId() As String
		Get
			Return mstrCorpId
		End Get
	End Property

	''' <summary>
	''' 应用的凭证密钥
	''' </summary>
	Private mstrCorpSecret As String
	Friend ReadOnly Property CorpSecret() As String
		Get
			Return mstrCorpSecret
		End Get
	End Property

	''' <summary>
	''' 用于后续接口的调用的access_token
	''' </summary>
	Private mstrAccessToken As String
	Public Property AccessToken() As String
		Get
			Return mstrAccessToken
		End Get
		Friend Set(ByVal value As String)
			mstrAccessToken = value
		End Set
	End Property

	''' <summary>
	''' 刷新access_token
	''' </summary>
	''' <param name="IsForce">是否强制刷新，企业微信可能会出于运营需要，提前使access_token失效，开发者应实现access_token失效时重新获取的逻辑。</param>
	Public Sub RefAccessToken(Optional IsForce As Boolean = False)
		Const SUB_NAME As String = "RefAccessToken"
		Dim strStepName As String = ""
		Try
			If mstrAccessToken = "" Or mdteAccessTokenExpiresTime < Now.AddMinutes(-1) Or IsForce = True Then
				'提前1分钟过期
				Dim strUrl As String = QYAPI_CGIBIN_URL & "/gettoken?corpid=" & Me.CorpId & "&corpsecret=" & Me.CorpSecret
				strStepName = "New PigWebReq"
				If Me.IsDebug = True Then strStepName &= "(" & strUrl & ")"
				Dim oPigWebReq As New PigWebReq(strUrl)
				With oPigWebReq
					strStepName = "PigWebReq.GetText"
					.GetText()
					If .LastErr <> "" Then
						Throw New Exception(.LastErr)
					Else
						strStepName = "New PigJSon"
						If Me.IsDebug = True Then strStepName &= "(" & .ResString & ")"
						Dim oPigJSon As New PigJSon(.ResString)
						If oPigJSon.LastErr <> "" Then Throw New Exception(oPigJSon.LastErr)
						strStepName = "检查 errcode"
						Select Case oPigJSon.GetStrValue("errcode")
							Case "", "0"
								strStepName = "检查 AccessToken"
								mstrAccessToken = oPigJSon.GetStrValue("access_token")
								Dim intExpiresIn As Integer = oPigJSon.GetIntValue("expires_in")
								If mstrAccessToken.Length = 0 Or intExpiresIn <= 0 Then
									Throw New Exception("返回数据不合理")
								End If
								mdteAccessTokenExpiresTime = Now.AddSeconds(intExpiresIn)
							Case Else
								Throw New Exception(oPigJSon.GetStrValue("errmsg"))
						End Select
					End If
				End With
				oPigWebReq = Nothing
			End If
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			If Me.IsDebug = True Then Me.PrintDebugLog(SUB_NAME, "Catch ex As Exception", Me.LastErr)
		End Try
	End Sub

	''' <summary>
	''' 处理OAuth2登录后从企业微信跳转回来并获取一个 WorkMember 对象
	''' </summary>
	''' <param name="HttpContext"></param>
	''' <param name="IsGetDetInf">是否获取详细信息，如果可以从数据库获取，可以使用默认值</param>
	Public Function GetWorkMemberFromOauth2Redirect(HttpContext As HttpContext, Optional IsGetDetInf As Boolean = False) As WorkMember
		Const SUB_NAME As String = "GetWorkMemberFromOauth2Redirect"
		Dim strStepName As String = ""
		Try
			strStepName = "New PigHttpContext"
			Dim oPigHttpContext As New PigHttpContext(HttpContext)
			Dim oWorkMember As New WorkMember
			strStepName = "GetUserIdentity"
			Me.GetUserIdentity(oPigHttpContext.RequestItem("code"), oWorkMember)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			oWorkMember.Oauth2State = oPigHttpContext.RequestItem("state")
			If IsGetDetInf = True Then
				strStepName = "MailList_ReadMember"
				Me.MailList_ReadMember(oWorkMember)
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			End If
			GetWorkMemberFromOauth2Redirect = oWorkMember
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Return Nothing
		End Try
	End Function

	''' <summary>
	''' 回调配置
	''' </summary>
	''' <param name="HttpContext"></param>
	Public Sub CallbackConf(HttpContext As HttpContext)
		Try
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("CallbackConf", ex)
		End Try
	End Sub

	''' <summary>
	''' 通讯录管理-读取成员
	''' </summary>
	''' <param name="oWorkMember">企业微信成员</param>
	Public Sub MailList_ReadMember(ByRef oWorkMember As WorkMember)
		Const SUB_NAME As String = "MailList_ReadMember"
		Dim strStepName As String = ""
		Try
			If Me.IsAccessTokenReady = False Then
				strStepName = "RefAccessToken"
				Me.RefAccessToken(True)
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			End If
			strStepName = "检查WorkMember.UserId"
			If oWorkMember.UserId = "" Then
				Throw New Exception("未初始化")
			End If
			Dim strUrl As String = QYAPI_CGIBIN_URL & "/user/get?access_token=" & Me.AccessToken & "&userid=" & oWorkMember.UserId
			strStepName = "New PigWebReq"
			Dim oPigWebReq As New PigWebReq(strUrl)
			With oPigWebReq
				strStepName = "PigWebReq.GetText"
				Me.PrintDebugLog(SUB_NAME, strStepName, strUrl, True)
				.GetText()
				If .LastErr <> "" Then
					Throw New Exception(.LastErr)
				Else
					Me.PrintDebugLog(SUB_NAME, strStepName, .ResString, True)
					strStepName = "New PigJSon"
					Dim oPigJSon As New PigJSon(.ResString)
					If oPigJSon.LastErr <> "" Then Throw New Exception(oPigJSon.LastErr)
					strStepName = "检查 errcode"
					Select Case oPigJSon.GetStrValue("errcode")
						Case "0"
							strStepName = "FillPropertiesByJSon"
							oWorkMember.FillPropertiesByJSon(oPigJSon)
							If oWorkMember.LastErr <> "" Then Throw New Exception(oWorkMember.LastErr)
						Case Else
							Throw New Exception(oPigJSon.GetStrValue("errmsg"))
					End Select
				End If
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Sub

	Private Function mWebReqGetText(ByRef PigWebReq As PigWebReq, ByRef ResPigJSon As PigJSon, Url As String) As String
		Const SUB_NAME As String = "mWebReqGetText"
		Dim strStepName As String = ""
		Try
			strStepName = "New PigWebReq"
			PigWebReq = New PigWebReq(Url)
			If PigWebReq.LastErr <> "" Then
				If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & ")"
				Throw New Exception()
			End If
			With PigWebReq
				strStepName = "GetText"
				.GetText()
				If .LastErr <> "" Then
					If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & ")"
					Throw New Exception(.LastErr)
				Else
					ResPigJSon = Nothing
					strStepName = "New PigJSon"
					ResPigJSon = New PigJSon(.ResString)
					If ResPigJSon.LastErr <> "" Then
						If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & ")"
						Throw New Exception(ResPigJSon.LastErr)
					End If
					strStepName = "检查 errcode"
					Select Case ResPigJSon.GetStrValue("errcode")
						Case "0"
						Case Else
							If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & ")"
							Throw New Exception(ResPigJSon.GetStrValue("errmsg"))
					End Select
				End If
			End With
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

	Private Function mWebReqPostText(ByRef PigWebReq As PigWebReq, ByRef ResPigJSon As PigJSon, Url As String, Para As String) As String
		Const SUB_NAME As String = "mWebReqPostText"
		Dim strStepName As String = ""
		Try
			strStepName = "New PigWebReq"
			PigWebReq = New PigWebReq(Url)
			If PigWebReq.LastErr <> "" Then
				If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & ")"
				Throw New Exception()
			End If
			With PigWebReq
				strStepName = "PostText"
				.PostText(Para)
				If .LastErr <> "" Then
					If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & "?" & Para & ")"
					Throw New Exception(.LastErr)
				Else
					ResPigJSon = Nothing
					strStepName = "New PigJSon"
					ResPigJSon = New PigJSon(.ResString)
					If ResPigJSon.LastErr <> "" Then
						If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & "?" & Para & ")"
						Throw New Exception(ResPigJSon.LastErr)
					End If
					strStepName = "检查 errcode"
					Select Case ResPigJSon.GetStrValue("errcode")
						Case "0"
						Case Else
							If Me.IsDebug Or Me.IsHardDebug Then strStepName &= "(" & Url & "?" & Para & ")"
							Throw New Exception(ResPigJSon.GetStrValue("errmsg"))
					End Select
				End If
			End With
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function


	''' <summary>
	''' 获取访问用户身份
	''' </summary>
	''' <param name="Code">通过成员授权获取到的code，最大为512字节。每次成员授权带上的code将不一样，code只能使用一次，5分钟未被使用自动过期。</param>
	''' <param name="oWorkMember">企业微信成员</param>
	Friend Sub GetUserIdentity(Code As String, ByRef oWorkMember As WorkMember)
		Const SUB_NAME As String = "GetUserIdentity"
		Dim strStepName As String = ""
		Try
			If Me.IsAccessTokenReady = False Then
				strStepName = "RefAccessToken"
				Me.RefAccessToken(True)
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			End If
			Dim strUrl As String = QYAPI_CGIBIN_URL & "/user/getuserinfo?access_token=" & Me.AccessToken & "&code=" & Code
			strStepName = "New PigWebReq"
			Dim oPigWebReq As New PigWebReq(strUrl)
			With oPigWebReq
				strStepName = "PigWebReq.GetText"
				Me.PrintDebugLog(SUB_NAME, strStepName, strUrl, True)
				.GetText()
				If .LastErr <> "" Then
					Throw New Exception(.LastErr)
				Else
					Me.PrintDebugLog(SUB_NAME, strStepName, .ResString, True)
					strStepName = "New PigJSon"
					Dim oPigJSon As New PigJSon(.ResString)
					If oPigJSon.LastErr <> "" Then Throw New Exception(oPigJSon.LastErr)
					strStepName = "检查 errcode"
					Select Case oPigJSon.GetStrValue("errcode")
						Case "0"
							strStepName = "设置oWorkMember属性"
							oWorkMember.UserId = oPigJSon.GetStrValue("UserId")
							oWorkMember.DeviceId = oPigJSon.GetStrValue("DeviceId")
							oWorkMember.OpenID = oPigJSon.GetStrValue("OpenId")
							Me.PrintDebugLog(SUB_NAME, strStepName, "UserId=" & oWorkMember.UserId, True)
							Me.PrintDebugLog(SUB_NAME, strStepName, "DeviceId=" & oWorkMember.DeviceId, True)
							Me.PrintDebugLog(SUB_NAME, strStepName, "OpenID=" & oWorkMember.OpenID, True)
						Case Else
							Throw New Exception(oPigJSon.GetStrValue("errmsg"))
					End Select
				End If
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Sub


	''' <summary>
	''' 获取企业微信API域名IP段
	''' </summary>
	''' <returns>地址数组</returns>
	Public Function GetApiIpList() As String()
		Const SUB_NAME As String = "GetApiIpList"
		Dim strStepName As String = ""
		Try
			If Me.IsAccessTokenReady = False Then
				strStepName = "RefAccessToken"
				Me.RefAccessToken(True)
				If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			End If
			Dim asApiIpList(-1) As String
			Dim strUrl As String = QYAPI_CGIBIN_URL & "/get_api_domain_ip?access_token=" & Me.AccessToken
			strStepName = "New PigWebReq"
			Dim oPigWebReq As New PigWebReq(strUrl)
			With oPigWebReq
				strStepName = "PigWebReq.GetText"
				.GetText()
				If .LastErr <> "" Then
					Throw New Exception(.LastErr)
				Else
					strStepName = "New PigJSon"
					Dim oPigJSon As New PigJSon(.ResString)
					If oPigJSon.LastErr <> "" Then Throw New Exception(oPigJSon.LastErr)
					strStepName = "检查 errcode"
					Select Case oPigJSon.GetStrValue("errcode")
						Case "0"
							strStepName = "检查 AccessToken"
							Dim intLen As Integer = oPigJSon.GetIntValue("ip_list.length")
							ReDim asApiIpList(intLen - 1)
							For i = 0 To intLen - 1
								asApiIpList(i) = oPigJSon.GetStrValue("ip_list[" & i.ToString & "]")
							Next
							GetApiIpList = asApiIpList
						Case Else
							Throw New Exception(oPigJSon.GetStrValue("errmsg"))
					End Select
				End If
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(SUB_NAME, strStepName, ex)
			Return Nothing
		End Try
	End Function

	''' <summary>
	''' 获取构造网页授权链接
	''' </summary>
	''' <param name="RedirectUri">授权后重定向的回调链接地址，内部会做使用urlencode对链接进行处理</param>
	''' <param name="State">重定向后会带上state参数，企业可以填写a-zA-Z0-9的参数值，长度不可超过128个字节</param>
	''' <returns></returns>
	Public Function GetOauth2Url(RedirectUri As String, State As String) As String
		Dim strStepName As String = ""
		Try
			strStepName = "检查 State"
			If moPigFunc.IsRegexMatch(State, "^[A-Za-z0-9]+$") = False Or State.Length > 128 Then
				Throw New Exception("数据不合法，企业可以填写a-zA-Z0-9的参数值，长度不可超过128个字节")
			End If
			GetOauth2Url = OEPN_WX_OAUTH2_URL & "?appid=" & Me.CorpId & "&redirect_uri=" & moPigFunc.UrlEncode(RedirectUri) & "&response_type=code&scope=snsapi_base"
			If State.Length > 0 Then
				GetOauth2Url &= "&state=" & State
			End If
			GetOauth2Url &= "#wechat_redirect"
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("RefAccessToken", strStepName, ex)
			Return ""
		End Try
	End Function

	Public Function GetErrMsg(ErrCode As Integer) As String
		Try
			Select Case ErrCode
				Case -1
					GetErrMsg = "系统繁忙"
				Case 0
					GetErrMsg = "请求成功"
				Case 6000
					GetErrMsg = "数据版本冲突"
				Case 40001
					GetErrMsg = "不合法的secret参数"
				Case 40003
					GetErrMsg = "无效的UserID"
				Case 40004
					GetErrMsg = "不合法的媒体文件类型"
				Case 40005
					GetErrMsg = "不合法的type参数"
				Case 40006
					GetErrMsg = "不合法的文件大小"
				Case 40007
					GetErrMsg = "不合法的media_id参数"
				Case 40008
					GetErrMsg = "不合法的msgtype参数"
				Case 40009
					GetErrMsg = "上传图片大小不是有效值"
				Case 40011
					GetErrMsg = "上传视频大小不是有效值"
				Case 40013
					GetErrMsg = "不合法的CorpID"
				Case 40014
					GetErrMsg = "不合法的access_token"
				Case 40016
					GetErrMsg = "不合法的按钮个数"
				Case 40017
					GetErrMsg = "不合法的按钮类型"
				Case 40018
					GetErrMsg = "不合法的按钮名字长度"
				Case 40019
					GetErrMsg = "不合法的按钮KEY长度"
				Case 40020
					GetErrMsg = "不合法的按钮URL长度"
				Case 40022
					GetErrMsg = "不合法的子菜单级数"
				Case 40023
					GetErrMsg = "不合法的子菜单按钮个数"
				Case 40024
					GetErrMsg = "不合法的子菜单按钮类型"
				Case 40025
					GetErrMsg = "不合法的子菜单按钮名字长度"
				Case 40026
					GetErrMsg = "不合法的子菜单按钮KEY长度"
				Case 40027
					GetErrMsg = "不合法的子菜单按钮URL长度"
				Case 40029
					GetErrMsg = "不合法的oauth_code"
				Case 40031
					GetErrMsg = "不合法的UserID列表"
				Case 40032
					GetErrMsg = "不合法的UserID列表长度"
				Case 40033
					GetErrMsg = "不合法的请求字符"
				Case 40035
					GetErrMsg = "不合法的参数"
				Case 40039
					GetErrMsg = "不合法的url长度"
				Case 40050
					GetErrMsg = "chatid不存在"
				Case 40054
					GetErrMsg = "不合法的子菜单url域名"
				Case 40055
					GetErrMsg = "不合法的菜单url域名"
				Case 40056
					GetErrMsg = "不合法的agentid"
				Case 40057
					GetErrMsg = "不合法的callbackurl或者callbackurl验证失败"
				Case 40058
					GetErrMsg = "不合法的参数"
				Case 40059
					GetErrMsg = "不合法的上报地理位置标志位"
				Case 40063
					GetErrMsg = "参数为空"
				Case 40066
					GetErrMsg = "不合法的部门列表"
				Case 40068
					GetErrMsg = "不合法的标签ID"
				Case 40070
					GetErrMsg = "指定的标签范围结点全部无效"
				Case 40071
					GetErrMsg = "不合法的标签名字"
				Case 40072
					GetErrMsg = "不合法的标签名字长度"
				Case 40073
					GetErrMsg = "不合法的openid"
				Case 40074
					GetErrMsg = "news消息不支持保密消息类型"
				Case 40077
					GetErrMsg = "不合法的pre_auth_code参数"
				Case 40078
					GetErrMsg = "不合法的auth_code参数"
				Case 40080
					GetErrMsg = "不合法的suite_secret"
				Case 40082
					GetErrMsg = "不合法的suite_token"
				Case 40083
					GetErrMsg = "不合法的suite_id"
				Case 40084
					GetErrMsg = "不合法的permanent_code参数"
				Case 40085
					GetErrMsg = "不合法的的suite_ticket参数"
				Case 40086
					GetErrMsg = "不合法的第三方应用appid"
				Case 40088
					GetErrMsg = "jobid不存在"
				Case 40089
					GetErrMsg = "批量任务的结果已清理"
				Case 40091
					GetErrMsg = "secret不合法"
				Case 40092
					GetErrMsg = "导入文件存在不合法的内容"
				Case 40093
					GetErrMsg = "jsapi签名错误"
				Case 40094
					GetErrMsg = "不合法的URL"
				Case 40096
					GetErrMsg = "不合法的外部联系人userid"
				Case 40097
					GetErrMsg = "该成员尚未离职"
				Case 40098
					GetErrMsg = "成员尚未实名认证"
				Case 40099
					GetErrMsg = "外部联系人的数量已达上限"
				Case 40100
					GetErrMsg = "此用户的外部联系人已经在转移流程中"
				Case 40102
					GetErrMsg = "域名或IP不可与应用市场上架应用重复"
				Case 40123
					GetErrMsg = "上传临时图片素材，图片格式非法"
				Case 40124
					GetErrMsg = "推广活动里的sn禁止绑定"
				Case 40125
					GetErrMsg = "无效的openuserid参数"
				Case 40126
					GetErrMsg = "企业标签个数达到上限，最多为3000个"
				Case 40127
					GetErrMsg = "不支持的uri schema"
				Case 40128
					GetErrMsg = "客户转接过于频繁（90天内只允许转接一次，同一个客户最多只能转接两次）"
				Case 40129
					GetErrMsg = "当前客户正在转接中"
				Case 40130
					GetErrMsg = "原跟进人与接手人一样，不可继承"
				Case 40131
					GetErrMsg = "handover_userid 并不是外部联系人的跟进人"
				Case 41001
					GetErrMsg = "缺少access_token参数"
				Case 41002
					GetErrMsg = "缺少corpid参数"
				Case 41004
					GetErrMsg = "缺少secret参数"
				Case 41006
					GetErrMsg = "缺少media_id参数"
				Case 41008
					GetErrMsg = "缺少auth code参数"
				Case 41009
					GetErrMsg = "缺少userid参数"
				Case 41010
					GetErrMsg = "缺少url参数"
				Case 41011
					GetErrMsg = "缺少agentid参数"
				Case 41016
					GetErrMsg = "缺少title参数"
				Case 41019
					GetErrMsg = "缺少 department 参数"
				Case 41017
					GetErrMsg = "缺少tagid参数"
				Case 41021
					GetErrMsg = "缺少suite_id参数"
				Case 41022
					GetErrMsg = "缺少suite_access_token参数"
				Case 41023
					GetErrMsg = "缺少suite_ticket参数"
				Case 41024
					GetErrMsg = "缺少secret参数"
				Case 41025
					GetErrMsg = "缺少permanent_code参数"
				Case 41033
					GetErrMsg = "缺少 description 参数"
				Case 41035
					GetErrMsg = "缺少外部联系人userid参数"
				Case 41036
					GetErrMsg = "不合法的企业对外简称"
				Case 41037
					GetErrMsg = "缺少「联系我」type参数"
				Case 41038
					GetErrMsg = "缺少「联系我」scene参数"
				Case 41039
					GetErrMsg = "无效的「联系我」type参数"
				Case 41040
					GetErrMsg = "无效的「联系我」scene参数"
				Case 41041
					GetErrMsg = "「联系我」使用人数超过限制"
				Case 41042
					GetErrMsg = "无效的「联系我」style参数"
				Case 41043
					GetErrMsg = "缺少「联系我」config_id参数"
				Case 41044
					GetErrMsg = "无效的「联系我」config_id参数"
				Case 41045
					GetErrMsg = "API添加「联系我」达到数量上限"
				Case 41046
					GetErrMsg = "缺少企业群发消息id"
				Case 41047
					GetErrMsg = "无效的企业群发消息id"
				Case 41048
					GetErrMsg = "无可发送的客户"
				Case 41049
					GetErrMsg = "缺少欢迎语code参数"
				Case 41050
					GetErrMsg = "无效的欢迎语code"
				Case 41051
					GetErrMsg = "客户和服务人员已经开始聊天了"
				Case 41052
					GetErrMsg = "无效的发送时间"
				Case 41053
					GetErrMsg = "客户未同意聊天存档"
				Case 41054
					GetErrMsg = "该用户尚未激活"
				Case 41055
					GetErrMsg = "群欢迎语模板数量达到上限"
				Case 41056
					GetErrMsg = "外部联系人id类型不正确"
				Case 41057
					GetErrMsg = "企业或服务商未绑定微信开发者账号"
				Case 41059
					GetErrMsg = "缺少moment_id参数"
				Case 41060
					GetErrMsg = "不合法的moment_id参数"
				Case 41061
					GetErrMsg = "不合法朋友圈发送成员userid，当前朋友圈并非此用户发表"
				Case 41062
					GetErrMsg = "企业创建的朋友圈尚未被成员userid发表"
				Case 41063
					GetErrMsg = "群发消息正在被派发中，请稍后再试"
				Case 41102
					GetErrMsg = "缺少菜单名"
				Case 42001
					GetErrMsg = "access_token已过期"
				Case 42007
					GetErrMsg = "pre_auth_code已过期"
				Case 42009
					GetErrMsg = "suite_access_token已过期"
				Case 42012
					GetErrMsg = "jsapi_ticket不可用，一般是没有正确调用接口来创建jsapi_ticket"
				Case 42013
					GetErrMsg = "小程序未登陆或登录态已经过期"
				Case 42014
					GetErrMsg = "任务卡片消息的task_id不合法"
				Case 42015
					GetErrMsg = "更新的消息的应用与发送消息的应用不匹配"
				Case 42016
					GetErrMsg = "更新的task_id不存在"
				Case 42017
					GetErrMsg = "按钮key值不存在"
				Case 42018
					GetErrMsg = "按钮key值不合法"
				Case 42019
					GetErrMsg = "缺少按钮key值不合法"
				Case 42020
					GetErrMsg = "缺少按钮名称"
				Case 42021
					GetErrMsg = "device_access_token 过期"
				Case 42022
					GetErrMsg = "code已经被使用过。只能使用一次"
				Case 43004
					GetErrMsg = "指定的userid未绑定微信或未关注微工作台（原企业号）"
				Case 43009
					GetErrMsg = "企业未验证主体"
				Case 44001
					GetErrMsg = "多媒体文件为空"
				Case 44004
					GetErrMsg = "文本消息content参数为空"
				Case 45001
					GetErrMsg = "多媒体文件大小超过限制"
				Case 45002
					GetErrMsg = "消息内容大小超过限制"
				Case 45004
					GetErrMsg = "应用description参数长度不符合系统限制"
				Case 45007
					GetErrMsg = "语音播放时间超过限制"
				Case 45008
					GetErrMsg = "图文消息的文章数量不符合系统限制"
				Case 45009
					GetErrMsg = "接口调用超过限制"
				Case 45022
					GetErrMsg = "应用name参数长度不符合系统限制"
				Case 45024
					GetErrMsg = "帐号数量超过上限"
				Case 45026
					GetErrMsg = "触发删除用户数的保护"
				Case 45029
					GetErrMsg = "回包大小超过上限"
				Case 45032
					GetErrMsg = "图文消息author参数长度超过限制"
				Case 45033
					GetErrMsg = "接口并发调用超过限制"
				Case 45034
					GetErrMsg = "url必须有协议头"
				Case 46003
					GetErrMsg = "菜单未设置"
				Case 46004
					GetErrMsg = "指定的用户不存在"
				Case 48002
					GetErrMsg = "API接口无权限调用"
				Case 48003
					GetErrMsg = "不合法的suite_id"
				Case 48004
					GetErrMsg = "授权关系无效"
				Case 48005
					GetErrMsg = "API接口已废弃"
				Case 48006
					GetErrMsg = "接口权限被收回"
				Case 49004
					GetErrMsg = "签名不匹配"
				Case 49008
					GetErrMsg = "群已经解散"
				Case 50001
					GetErrMsg = "redirect_url未登记可信域名"
				Case 50002
					GetErrMsg = "成员不在权限范围"
				Case 50003
					GetErrMsg = "应用已禁用"
				Case 50100
					GetErrMsg = "分页查询的游标无效"
				Case 60001
					GetErrMsg = "部门长度不符合限制"
				Case 60003
					GetErrMsg = "部门ID不存在"
				Case 60004
					GetErrMsg = "父部门不存在"
				Case 60005
					GetErrMsg = "部门下存在成员"
				Case 60006
					GetErrMsg = "部门下存在子部门"
				Case 60007
					GetErrMsg = "不允许删除根部门"
				Case 60008
					GetErrMsg = "部门已存在"
				Case 60009
					GetErrMsg = "部门名称含有非法字符"
				Case 60010
					GetErrMsg = "部门存在循环关系"
				Case 60011
					GetErrMsg = "指定的成员/部门/标签参数无权限"
				Case 60012
					GetErrMsg = "不允许删除默认应用"
				Case 60020
					GetErrMsg = "访问ip不在白名单之中"
				Case 60021
					GetErrMsg = "userid不在应用可见范围内"
				Case 60028
					GetErrMsg = "不允许修改第三方应用的主页 URL"
				Case 60102
					GetErrMsg = "UserID已存在"
				Case 60103
					GetErrMsg = "手机号码不合法"
				Case 60104
					GetErrMsg = "手机号码已存在"
				Case 60105
					GetErrMsg = "邮箱不合法"
				Case 60106
					GetErrMsg = "邮箱已存在"
				Case 60107
					GetErrMsg = "微信号不合法"
				Case 60110
					GetErrMsg = "用户所属部门数量超过限制"
				Case 60111
					GetErrMsg = "UserID不存在"
				Case 60112
					GetErrMsg = "成员name参数不合法"
				Case 60123
					GetErrMsg = "无效的部门id"
				Case 60124
					GetErrMsg = "无效的父部门id"
				Case 60125
					GetErrMsg = "非法部门名字"
				Case 60127
					GetErrMsg = "缺少department参数"
				Case 60129
					GetErrMsg = "成员手机和邮箱都为空"
				Case 60132
					GetErrMsg = "is_leader_in_dept和department的元素个数不一致"
				Case 60136
					GetErrMsg = "记录不存在"
				Case 60137
					GetErrMsg = "家长手机号重复"
				Case 60203
					GetErrMsg = "不合法的模版ID"
				Case 60204
					GetErrMsg = "模版状态不可用"
				Case 60205
					GetErrMsg = "模版关键词不匹配"
				Case 60206
					GetErrMsg = "该种类型的消息只支持第三方独立应用使用"
				Case 60207
					GetErrMsg = "第三方独立应用只允许发送模板消息"
				Case 60208
					GetErrMsg = "第三方独立应用不支持指定@all，不支持参数toparty和totag"
				Case 65000
					GetErrMsg = "学校已经迁移"
				Case 65001
					GetErrMsg = "无效的关注模式"
				Case 65002
					GetErrMsg = "导入家长信息数量过多"
				Case 65003
					GetErrMsg = "学校尚未迁移"
				Case 65004
					GetErrMsg = "组织架构不存在"
				Case 65005
					GetErrMsg = "无效的同步模式"
				Case 65006
					GetErrMsg = "无效的管理员类型"
				Case 65007
					GetErrMsg = "无效的家校部门类型"
				Case 65008
					GetErrMsg = "无效的入学年份"
				Case 65009
					GetErrMsg = "无效的标准年级类型"
				Case 65010
					GetErrMsg = "此userid并不是学生"
				Case 65011
					GetErrMsg = "家长userid数量超过限制"
				Case 65012
					GetErrMsg = "学生userid数量超过限制"
				Case 65013
					GetErrMsg = "学生已有家长"
				Case 65014
					GetErrMsg = "非学校企业"
				Case 65015
					GetErrMsg = "父部门类型不匹配"
				Case 65018
					GetErrMsg = "家长人数达到上限"
				Case 660001
					GetErrMsg = "无效的商户号"
				Case 660002
					GetErrMsg = "无效的企业收款人id"
				Case 660003
					GetErrMsg = "userid不在应用的可见范围"
				Case 660004
					GetErrMsg = "partyid不在应用的可见范围"
				Case 660005
					GetErrMsg = "tagid不在应用的可见范围"
				Case 660006
					GetErrMsg = "找不到该商户号"
				Case 660007
					GetErrMsg = "申请已经存在"
				Case 660008
					GetErrMsg = "商户号已经绑定"
				Case 660009
					GetErrMsg = "商户号主体和商户主体不一致"
				Case 660010
					GetErrMsg = "超过商户号绑定数量限制"
				Case 660011
					GetErrMsg = "商户号未绑定"
				Case 670001
					GetErrMsg = "应用不在共享范围"
				Case 72023
					GetErrMsg = "发票已被其他公众号锁定"
				Case 72024
					GetErrMsg = "发票状态错误"
				Case 72037
					GetErrMsg = "存在发票不属于该用户"
				Case 80001
					GetErrMsg = "可信域名不正确，或者无ICP备案"
				Case 81001
					GetErrMsg = "部门下的结点数超过限制（3W）"
				Case 81002
					GetErrMsg = "部门最多15层"
				Case 81003
					GetErrMsg = "标签下节点个数超过30000个"
				Case 81011
					GetErrMsg = "无权限操作标签"
				Case 81012
					GetErrMsg = "缺失可见范围"
				Case 81013
					GetErrMsg = "UserID、部门ID、标签ID全部非法或无权限"
				Case 81014
					GetErrMsg = "标签添加成员，单次添加user或party过多"
				Case 81015
					GetErrMsg = "邮箱域名需要跟企业邮箱域名一致"
				Case 81016
					GetErrMsg = "logined_userid字段缺失"
				Case 81017
					GetErrMsg = "items字段大小超过限制（20）"
				Case 81018
					GetErrMsg = "该服务商可获取名字数量配额不足"
				Case 81019
					GetErrMsg = "items数组成员缺少id字段"
				Case 81020
					GetErrMsg = "items数组成员缺少type字段"
				Case 81021
					GetErrMsg = "items数组成员的type字段不合法"
				Case 82001
					GetErrMsg = "指定的成员/部门/标签全部为空"
				Case 82002
					GetErrMsg = "不合法的PartyID列表长度"
				Case 82003
					GetErrMsg = "不合法的TagID列表长度"
				Case 84014
					GetErrMsg = "成员票据过期"
				Case 84015
					GetErrMsg = "成员票据无效"
				Case 84019
					GetErrMsg = "缺少templateid参数"
				Case 84020
					GetErrMsg = "templateid不存在"
				Case 84021
					GetErrMsg = "缺少register_code参数"
				Case 84022
					GetErrMsg = "无效的register_code参数"
				Case 84023
					GetErrMsg = "不允许调用设置通讯录同步完成接口"
				Case 84024
					GetErrMsg = "无注册信息"
				Case 84025
					GetErrMsg = "不符合的state参数"
				Case 84052
					GetErrMsg = "缺少caller参数"
				Case 84053
					GetErrMsg = "缺少callee参数"
				Case 84054
					GetErrMsg = "缺少auth_corpid参数"
				Case 84055
					GetErrMsg = "超过拨打公费电话频率"
				Case 84056
					GetErrMsg = "被拨打用户安装应用时未授权拨打公费电话权限"
				Case 84057
					GetErrMsg = "公费电话余额不足"
				Case 84058
					GetErrMsg = "caller 呼叫号码不支持"
				Case 84059
					GetErrMsg = "号码非法"
				Case 84060
					GetErrMsg = "callee 呼叫号码不支持"
				Case 84061
					GetErrMsg = "不存在外部联系人的关系"
				Case 84062
					GetErrMsg = "未开启公费电话应用"
				Case 84063
					GetErrMsg = "caller不存在"
				Case 84064
					GetErrMsg = "callee不存在"
				Case 84065
					GetErrMsg = "caller跟callee电话号码一致"
				Case 84066
					GetErrMsg = "服务商拨打次数超过限制"
				Case 84067
					GetErrMsg = "管理员收到的服务商公费电话个数超过限制"
				Case 84069
					GetErrMsg = "拨打方被限制拨打公费电话"
				Case 84070
					GetErrMsg = "不支持的电话号码"
				Case 84071
					GetErrMsg = "不合法的外部联系人授权码"
				Case 84072
					GetErrMsg = "应用未配置客服"
				Case 84073
					GetErrMsg = "客服userid不在应用配置的客服列表中"
				Case 84074
					GetErrMsg = "没有外部联系人权限"
				Case 84075
					GetErrMsg = "不合法或过期的authcode"
				Case 84076
					GetErrMsg = "缺失authcode"
				Case 84077
					GetErrMsg = "订单价格过高，无法受理"
				Case 84078
					GetErrMsg = "购买人数不正确"
				Case 84079
					GetErrMsg = "价格策略不存在"
				Case 84080
					GetErrMsg = "订单不存在"
				Case 84081
					GetErrMsg = "存在未支付订单"
				Case 84082
					GetErrMsg = "存在申请退款中的订单"
				Case 84083
					GetErrMsg = "非服务人员"
				Case 84084
					GetErrMsg = "非跟进用户"
				Case 84085
					GetErrMsg = "应用已下架"
				Case 84086
					GetErrMsg = "订单人数超过可购买最大人数"
				Case 84087
					GetErrMsg = "打开订单支付前禁止关闭订单"
				Case 84088
					GetErrMsg = "禁止关闭已支付的订单"
				Case 84089
					GetErrMsg = "订单已支付"
				Case 84090
					GetErrMsg = "缺失user_ticket"
				Case 84091
					GetErrMsg = "订单价格不可低于下限"
				Case 84092
					GetErrMsg = "无法发起代下单操作"
				Case 84093
					GetErrMsg = "代理关系已占用，无法代下单"
				Case 84094
					GetErrMsg = "该应用未配置代理分润规则，请先联系应用服务商处理"
				Case 84095
					GetErrMsg = "免费试用版，无法扩容"
				Case 84096
					GetErrMsg = "免费试用版，无法续期"
				Case 84097
					GetErrMsg = "当前企业有未处理订单"
				Case 84098
					GetErrMsg = "固定总量，无法扩容"
				Case 84099
					GetErrMsg = "非购买状态，无法扩容"
				Case 84100
					GetErrMsg = "未购买过此应用，无法续期"
				Case 84101
					GetErrMsg = "企业已试用付费版本，无法全新购买"
				Case 84102
					GetErrMsg = "企业当前应用状态已过期，无法扩容"
				Case 84103
					GetErrMsg = "仅可修改未支付订单"
				Case 84104
					GetErrMsg = "订单已支付，无法修改"
				Case 84105
					GetErrMsg = "订单已被取消，无法修改"
				Case 84106
					GetErrMsg = "企业含有该应用的待支付订单，无法代下单"
				Case 84107
					GetErrMsg = "企业含有该应用的退款中订单，无法代下单"
				Case 84108
					GetErrMsg = "企业含有该应用的待生效订单，无法代下单"
				Case 84109
					GetErrMsg = "订单定价不能未0"
				Case 84110
					GetErrMsg = "新安装应用不在试用状态，无法升级为付费版"
				Case 84111
					GetErrMsg = "无足够可用优惠券"
				Case 84112
					GetErrMsg = "无法关闭未支付订单"
				Case 84113
					GetErrMsg = "无付费信息"
				Case 84114
					GetErrMsg = "虚拟版本不支持下单"
				Case 84115
					GetErrMsg = "虚拟版本不支持扩容"
				Case 84116
					GetErrMsg = "虚拟版本不支持续期"
				Case 84117
					GetErrMsg = "在虚拟正式版期内不能扩容"
				Case 84118
					GetErrMsg = "虚拟正式版期内不能变更版本"
				Case 84119
					GetErrMsg = "当前企业未报备，无法进行代下单"
				Case 84120
					GetErrMsg = "当前应用版本已删除"
				Case 84121
					GetErrMsg = "应用版本已删除，无法扩容"
				Case 84122
					GetErrMsg = "应用版本已删除，无法续期"
				Case 84123
					GetErrMsg = "非虚拟版本，无法升级"
				Case 84124
					GetErrMsg = "非行业方案订单，不能添加部分应用版本的订单"
				Case 84125
					GetErrMsg = "购买人数不能少于最少购买人数"
				Case 84126
					GetErrMsg = "购买人数不能多于最大购买人数"
				Case 84127
					GetErrMsg = "无应用管理权限"
				Case 84128
					GetErrMsg = "无该行业方案下全部应用的管理权限"
				Case 84129
					GetErrMsg = "付费策略已被删除，无法下单"
				Case 84130
					GetErrMsg = "订单生效时间不合法"
				Case 84200
					GetErrMsg = "文件转译解析错误"
				Case 85002
					GetErrMsg = "包含不合法的词语"
				Case 85004
					GetErrMsg = "每企业每个月设置的可信域名不可超过20个"
				Case 85005
					GetErrMsg = "可信域名未通过所有权校验"
				Case 86001
					GetErrMsg = "参数 chatid 不合法"
				Case 86003
					GetErrMsg = "参数 chatid 不存在"
				Case 86004
					GetErrMsg = "参数 群名不合法"
				Case 86005
					GetErrMsg = "参数 群主不合法"
				Case 86006
					GetErrMsg = "群成员数过多或过少"
				Case 86007
					GetErrMsg = "不合法的群成员"
				Case 86008
					GetErrMsg = "非法操作非自己创建的群"
				Case 86101
					GetErrMsg = "仅群主才有操作权限"
				Case 86201
					GetErrMsg = "参数 需要chatid"
				Case 86202
					GetErrMsg = "参数 需要群名"
				Case 86203
					GetErrMsg = "参数 需要群主"
				Case 86204
					GetErrMsg = "参数 需要群成员"
				Case 86205
					GetErrMsg = "参数 字符串chatid过长"
				Case 86206
					GetErrMsg = "参数 数字chatid过大"
				Case 86207
					GetErrMsg = "群主不在群成员列表"
				Case 86214
					GetErrMsg = "群发类型不合法"
				Case 86215
					GetErrMsg = "会话ID已经存在"
				Case 86216
					GetErrMsg = "存在非法会话成员ID"
				Case 86217
					GetErrMsg = "会话发送者不在会话成员列表中"
				Case 86220
					GetErrMsg = "指定的会话参数不合法"
				Case 86224
					GetErrMsg = "不是受限群，不允许使用该接口"
				Case 90001
					GetErrMsg = "未认证摇一摇周边"
				Case 90002
					GetErrMsg = "缺少摇一摇周边ticket参数"
				Case 90003
					GetErrMsg = "摇一摇周边ticket参数不合法"
				Case 90100
					GetErrMsg = "非法的对外属性类型"
				Case 90101
					GetErrMsg = "对外属性：文本类型长度不合法"
				Case 90102
					GetErrMsg = "对外属性：网页类型标题长度不合法"
				Case 90103
					GetErrMsg = "对外属性：网页url不合法"
				Case 90104
					GetErrMsg = "对外属性：小程序类型标题长度不合法"
				Case 90105
					GetErrMsg = "对外属性：小程序类型pagepath不合法"
				Case 90106
					GetErrMsg = "对外属性：请求参数不合法"
				Case 90200
					GetErrMsg = "缺少小程序appid参数"
				Case 90201
					GetErrMsg = "小程序通知的content_item个数超过限制"
				Case 90202
					GetErrMsg = "小程序通知中的key长度不合法"
				Case 90203
					GetErrMsg = "小程序通知中的value长度不合法"
				Case 90204
					GetErrMsg = "小程序通知中的page参数不合法"
				Case 90206
					GetErrMsg = "小程序未关联到企业中"
				Case 90207
					GetErrMsg = "不合法的小程序appid"
				Case 90208
					GetErrMsg = "小程序appid不匹配"
				Case 90300
					GetErrMsg = "orderid 不合法"
				Case 90302
					GetErrMsg = "付费应用已过期"
				Case 90303
					GetErrMsg = "付费应用超过最大使用人数"
				Case 90304
					GetErrMsg = "订单中心服务异常，请稍后重试"
				Case 90305
					GetErrMsg = "参数错误，errmsg中有提示具体哪个参数有问题"
				Case 90306
					GetErrMsg = "商户设置不合法，详情请见errmsg"
				Case 90307
					GetErrMsg = "登录态过期"
				Case 90308
					GetErrMsg = "在开启IP鉴权的前提下，识别为无效的请求IP"
				Case 90309
					GetErrMsg = "订单已经存在，请勿重复下单"
				Case 90310
					GetErrMsg = "找不到订单"
				Case 90311
					GetErrMsg = "关单失败, 可能原因：该单并没被拉起支付页面; 已经关单；已经支付；渠道失败；单处于保护状态；等等"
				Case 90312
					GetErrMsg = "退款请求失败, 详情请看errmsg"
				Case 90313
					GetErrMsg = "退款调用频率限制，超过规定的阈值"
				Case 90314
					GetErrMsg = "订单状态错误，可能未支付，或者当前状态操作受限"
				Case 90315
					GetErrMsg = "退款请求失败，主键冲突，请核实退款refund_id是否已使用"
				Case 90316
					GetErrMsg = "退款原因编号不对"
				Case 90317
					GetErrMsg = "尚未注册成为供应商"
				Case 90318
					GetErrMsg = "参数nonce_str 为空或者重复，判定为重放攻击"
				Case 90319
					GetErrMsg = "时间戳为空或者与系统时间间隔太大"
				Case 90320
					GetErrMsg = "订单token无效"
				Case 90321
					GetErrMsg = "订单token已过有效时间"
				Case 90322
					GetErrMsg = "旧套件（包含多个应用的套件）不支持支付系统"
				Case 90323
					GetErrMsg = "单价超过限额"
				Case 90324
					GetErrMsg = "商品数量超过限额"
				Case 90325
					GetErrMsg = "预支单已经存在"
				Case 90326
					GetErrMsg = "预支单单号非法"
				Case 90327
					GetErrMsg = "该预支单已经结算下单"
				Case 90328
					GetErrMsg = "结算下单失败，详情请看errmsg"
				Case 90329
					GetErrMsg = "该订单号已经被预支单占用"
				Case 90330
					GetErrMsg = "创建供应商失败"
				Case 90331
					GetErrMsg = "更新供应商失败"
				Case 90332
					GetErrMsg = "还没签署合同"
				Case 90333
					GetErrMsg = "创建合同失败"
				Case 90338
					GetErrMsg = "已经过了可退款期限"
				Case 90339
					GetErrMsg = "供应商主体名包含非法字符"
				Case 90340
					GetErrMsg = "创建客户失败，可能信息真实性校验失败"
				Case 90341
					GetErrMsg = "退款金额大于付款金额"
				Case 90342
					GetErrMsg = "退款金额超过账户余额"
				Case 90343
					GetErrMsg = "退款单号已经存在"
				Case 90344
					GetErrMsg = "指定的付款渠道无效"
				Case 90345
					GetErrMsg = "超过5w人民币不可指定微信支付渠道"
				Case 90346
					GetErrMsg = "同一单的退款次数超过限制"
				Case 90347
					GetErrMsg = "退款金额不可为0"
				Case 90348
					GetErrMsg = "管理端没配置支付密钥"
				Case 90349
					GetErrMsg = "记录数量太大"
				Case 90350
					GetErrMsg = "银行信息真实性校验失败"
				Case 90351
					GetErrMsg = "应用状态异常"
				Case 90352
					GetErrMsg = "延迟试用期天数超过限制"
				Case 90353
					GetErrMsg = "预支单列表不可为空"
				Case 90354
					GetErrMsg = "预支单列表数量超过限制"
				Case 90355
					GetErrMsg = "关联有退款预支单，不可删除"
				Case 90356
					GetErrMsg = "不能0金额下单"
				Case 90357
					GetErrMsg = "代下单必须指定支付渠道"
				Case 90358
					GetErrMsg = "预支单或代下单，不支持部分退款"
				Case 90359
					GetErrMsg = "预支单与下单者企业不匹配"
				Case 90381
					GetErrMsg = "参数 refunded_credit_orderid 不合法"
				Case 90456
					GetErrMsg = "必须指定组织者"
				Case 90457
					GetErrMsg = "日历ID异常"
				Case 90458
					GetErrMsg = "日历ID列表不能为空"
				Case 90459
					GetErrMsg = "日历已删除"
				Case 90460
					GetErrMsg = "日程已删除"
				Case 90461
					GetErrMsg = "日程ID异常"
				Case 90462
					GetErrMsg = "日程ID列表不能为空"
				Case 90463
					GetErrMsg = "不能变更组织者"
				Case 90464
					GetErrMsg = "参与者数量超过限制"
				Case 90465
					GetErrMsg = "不支持的重复类型"
				Case 90466
					GetErrMsg = "不能操作别的应用创建的日历/日程"
				Case 90467
					GetErrMsg = "星期参数异常"
				Case 90468
					GetErrMsg = "不能变更组织者"
				Case 90469
					GetErrMsg = "每页大小超过限制"
				Case 90470
					GetErrMsg = "页数异常"
				Case 90471
					GetErrMsg = "提醒时间异常"
				Case 90472
					GetErrMsg = "没有日历/日程操作权限"
				Case 90473
					GetErrMsg = "颜色参数异常"
				Case 90474
					GetErrMsg = "组织者不能与参与者重叠"
				Case 90475
					GetErrMsg = "不是组织者的日历"
				Case 90479
					GetErrMsg = "不允许操作用户创建的日程"
				Case 90500
					GetErrMsg = "群主并未离职"
				Case 90501
					GetErrMsg = "该群不是客户群"
				Case 90502
					GetErrMsg = "群主已经离职"
				Case 90503
					GetErrMsg = "满人 & 99个微信成员，没办法踢，要客户端确认"
				Case 90504
					GetErrMsg = "群主没变"
				Case 90507
					GetErrMsg = "离职群正在继承处理中"
				Case 90508
					GetErrMsg = "离职群已经继承"
				Case 91040
					GetErrMsg = "获取ticket的类型无效"
				Case 92000
					GetErrMsg = "成员不在应用可见范围之内"
				Case 92001
					GetErrMsg = "应用没有敏感信息权限"
				Case 92002
					GetErrMsg = "不允许跨企业调用"
				Case 92006
					GetErrMsg = "该直播已经开始或取消"
				Case 92007
					GetErrMsg = "该直播回放不能被删除"
				Case 92008
					GetErrMsg = "当前应用没权限操作这个直播"
				Case 93000
					GetErrMsg = "机器人webhookurl不合法或者机器人已经被移除出群"
				Case 93004
					GetErrMsg = "机器人被停用"
				Case 93008
					GetErrMsg = "不在群里"
				Case 94000
					GetErrMsg = "应用未开启工作台自定义模式"
				Case 94001
					GetErrMsg = "不合法的type类型"
				Case 94002
					GetErrMsg = "缺少keydata字段"
				Case 94003
					GetErrMsg = "keydata的items列表长度超出限制"
				Case 94005
					GetErrMsg = "缺少list字段"
				Case 94006
					GetErrMsg = "list的items列表长度超出限制"
				Case 94007
					GetErrMsg = "缺少webview字段"
				Case 94008
					GetErrMsg = "应用未设置自定义工作台模版类型"
				Case 301002
					GetErrMsg = "无权限操作指定的应用"
				Case 301005
					GetErrMsg = "不允许删除创建者"
				Case 301012
					GetErrMsg = "参数 position 不合法"
				Case 301013
					GetErrMsg = "参数 telephone 不合法"
				Case 301014
					GetErrMsg = "参数 english_name 不合法"
				Case 301015
					GetErrMsg = "参数 mediaid 不合法"
				Case 301016
					GetErrMsg = "上传语音文件不符合系统要求"
				Case 301017
					GetErrMsg = "上传语音文件仅支持AMR格式"
				Case 301021
					GetErrMsg = "参数 userid 无效"
				Case 301022
					GetErrMsg = "获取打卡数据失败"
				Case 301023
					GetErrMsg = "useridlist非法或超过限额"
				Case 301024
					GetErrMsg = "获取打卡记录时间间隔超限"
				Case 301025
					GetErrMsg = "审批开放接口参数错误"
				Case 301036
					GetErrMsg = "不允许更新该用户的userid"
				Case 301039
					GetErrMsg = "请求参数错误，请检查输入参数"
				Case 301042
					GetErrMsg = "ip白名单限制，请求ip不在设置白名单范围"
				Case 301048
					GetErrMsg = "sdkfileid对应的文件不存在或已过期"
				Case 301052
					GetErrMsg = "会话存档服务已过期"
				Case 301053
					GetErrMsg = "会话存档服务未开启"
				Case 301058
					GetErrMsg = "拉取会话数据请求超过大小限制，可减少limit参数"
				Case 301059
					GetErrMsg = "非内部群，不提供数据"
				Case 301060
					GetErrMsg = "拉取同意情况请求量过大，请减少到100个参数以下"
				Case 301061
					GetErrMsg = "userid或者exteropenid用户不存在"
				Case 302003
					GetErrMsg = "批量导入任务的文件中userid有重复"
				Case 302004
					GetErrMsg = "组织架构不合法（1不是一棵树，2 多个一样的partyid，3 partyid空，4 partyid name 空，5 同一个父节点下有两个子节点 部门名字一样 可能是以上情况，请一一排查）"
				Case 302005
					GetErrMsg = "批量导入系统失败，请重新尝试导入"
				Case 302006
					GetErrMsg = "批量导入任务的文件中partyid有重复"
				Case 302007
					GetErrMsg = "批量导入任务的文件中，同一个部门下有两个子部门名字一样"
				Case 2000002
					GetErrMsg = "CorpId参数无效"
				Case 600001
					GetErrMsg = "不合法的sn"
				Case 600002
					GetErrMsg = "设备已注册"
				Case 600003
					GetErrMsg = "不合法的硬件activecode"
				Case 600004
					GetErrMsg = "该硬件尚未授权任何企业"
				Case 600005
					GetErrMsg = "硬件Secret无效"
				Case 600007
					GetErrMsg = "缺少硬件sn"
				Case 600008
					GetErrMsg = "缺少nonce参数"
				Case 600009
					GetErrMsg = "缺少timestamp参数"
				Case 600010
					GetErrMsg = "缺少signature参数"
				Case 600011
					GetErrMsg = "签名校验失败"
				Case 600012
					GetErrMsg = "长连接已经注册过设备"
				Case 600013
					GetErrMsg = "缺少activecode参数"
				Case 600014
					GetErrMsg = "设备未网络注册"
				Case 600015
					GetErrMsg = "缺少secret参数"
				Case 600016
					GetErrMsg = "设备未激活"
				Case 600018
					GetErrMsg = "无效的起始结束时间"
				Case 600020
					GetErrMsg = "设备未登录"
				Case 600021
					GetErrMsg = "设备sn已存在"
				Case 600023
					GetErrMsg = "时间戳已失效"
				Case 600024
					GetErrMsg = "固件大小超过5M"
				Case 600025
					GetErrMsg = "固件名为空或者超过20字节"
				Case 600026
					GetErrMsg = "固件信息不存在"
				Case 600027
					GetErrMsg = "非法的固件参数"
				Case 600028
					GetErrMsg = "固件版本已存在"
				Case 600029
					GetErrMsg = "非法的固件版本"
				Case 600030
					GetErrMsg = "缺少固件版本参数"
				Case 600031
					GetErrMsg = "硬件固件不允许升级"
				Case 600032
					GetErrMsg = "无法解析硬件二维码"
				Case 600033
					GetErrMsg = "设备型号id冲突"
				Case 600034
					GetErrMsg = "指纹数据大小超过限制"
				Case 600035
					GetErrMsg = "人脸数据大小超过限制"
				Case 600036
					GetErrMsg = "设备sn冲突"
				Case 600037
					GetErrMsg = "缺失设备型号id"
				Case 600038
					GetErrMsg = "设备型号不存在"
				Case 600039
					GetErrMsg = "不支持的设备类型"
				Case 600040
					GetErrMsg = "打印任务id不存在"
				Case 600041
					GetErrMsg = "无效的offset或limit参数值"
				Case 600042
					GetErrMsg = "无效的设备型号id"
				Case 600043
					GetErrMsg = "门禁规则未设置"
				Case 600044
					GetErrMsg = "门禁规则不合法"
				Case 600045
					GetErrMsg = "设备已订阅企业信息"
				Case 600046
					GetErrMsg = "操作id和用户userid不匹配"
				Case 600047
					GetErrMsg = "secretno的status非法"
				Case 600048
					GetErrMsg = "无效的指纹算法"
				Case 600049
					GetErrMsg = "无效的人脸识别算法"
				Case 600050
					GetErrMsg = "无效的算法长度"
				Case 600051
					GetErrMsg = "设备过期"
				Case 600052
					GetErrMsg = "无效的文件分块"
				Case 600053
					GetErrMsg = "该链接已经激活"
				Case 600054
					GetErrMsg = "该链接已经订阅"
				Case 600055
					GetErrMsg = "无效的用户类型"
				Case 600056
					GetErrMsg = "无效的健康状态"
				Case 600057
					GetErrMsg = "缺少体温参数"
				Case 610001
					GetErrMsg = "永久二维码超过每个员工5000的限制"
				Case 610003
					GetErrMsg = "scene参数不合法"
				Case 610004
					GetErrMsg = "userid不在客户联系配置的使用范围内"
				Case 610014
					GetErrMsg = "无效的unionid"
				Case 610015
					GetErrMsg = "小程序对应的开放平台账号未认证"
				Case 610016
					GetErrMsg = "企业未认证"
				Case 610017
					GetErrMsg = "小程序和企业主体不一致"
				Case 640001
					GetErrMsg = "微盘不存在当前空间"
				Case 640002
					GetErrMsg = "文件不存在"
				Case 640003
					GetErrMsg = "文件已删除"
				Case 640004
					GetErrMsg = "无权限访问"
				Case 640005
					GetErrMsg = "成员不在空间内"
				Case 640006
					GetErrMsg = "超出当前成员拥有的容量"
				Case 640007
					GetErrMsg = "超出微盘的容量"
				Case 640008
					GetErrMsg = "没有空间权限"
				Case 640009
					GetErrMsg = "非法文件名"
				Case 640010
					GetErrMsg = "超出空间的最大成员数"
				Case 640011
					GetErrMsg = "json格式不匹配"
				Case 640012
					GetErrMsg = "非法的userid"
				Case 640013
					GetErrMsg = "非法的departmentid"
				Case 640014
					GetErrMsg = "空间没有有效的管理员"
				Case 640015
					GetErrMsg = "不支持设置预览权限"
				Case 640016
					GetErrMsg = "不支持设置文件水印"
				Case 640017
					GetErrMsg = "微盘管理端未开通API 权限"
				Case 640018
					GetErrMsg = "微盘管理端未设置编辑权限"
				Case 640019
					GetErrMsg = "API 调用次数超出限制"
				Case 640020
					GetErrMsg = "非法的权限类型"
				Case 640021
					GetErrMsg = "非法的fatherid"
				Case 640022
					GetErrMsg = "非法的文件内容的base64"
				Case 640023
					GetErrMsg = "非法的权限范围"
				Case 640024
					GetErrMsg = "非法的fileid"
				Case 640025
					GetErrMsg = "非法的space_name"
				Case 640026
					GetErrMsg = "非法的spaceid"
				Case 640027
					GetErrMsg = "参数错误"
				Case 640028
					GetErrMsg = "空间设置了关闭成员邀请链接"
				Case 640029
					GetErrMsg = "只支持下载普通文件，不支持下载文件夹等其他非文件实体类型"
				Case 844001
					GetErrMsg = "非法的output_file_format"
				Case Else
					GetErrMsg = "未知错误码" & ErrCode.ToString
			End Select
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("", ex)
			Return ""
		End Try
	End Function

	Public Function SendWorkMessage(ByRef InWorkMessage As WorkMessage) As String
		Const SUB_NAME As String = "SendWorkMessage"
		Dim strStepName As String = "", strRet As String = ""
		Try
			Dim oPigWebReq As PigWebReq = Nothing, strUrl As String = "", pjRes As PigJSon = Nothing
			With InWorkMessage
				strStepName = "检查消息类型"
				Select Case .MsgType
					Case WorkMessage.MsgTypeEnum.Text
						strStepName = "检查消息内容"
						If .Content = "" Then Throw New Exception("未指定内容")
						If .Content.Length > 2048 Then Me.PrintDebugLog(SUB_NAME, strStepName, "内容超过2048，将截断")
						If Me.IsAccessTokenReady = False Then
							strStepName = "RefAccessToken"
							Me.RefAccessToken(True)
							If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
						End If
						strUrl = QYAPI_CGIBIN_URL & "/message/send?access_token=" & Me.AccessToken
						strStepName = "mWebReqGetText"
						strRet = Me.mWebReqPostText(oPigWebReq, pjRes, strUrl, .PostJSon)
						If strRet <> "OK" Then Throw New Exception(strRet)
						Dim strValue As String = ""
						strValue = pjRes.GetStrValue("invaliduser")
						If strValue <> "" Then .InvalidParty = strValue
						strValue = pjRes.GetStrValue("invalidparty")
						If strValue <> "" Then .InvalidParty = strValue
						strValue = pjRes.GetStrValue("invalidtag")
						If strValue <> "" Then .InvalidTag = strValue
					Case Else
						Throw New Exception("目前尚不支持" & .MsgType.ToString)
				End Select
			End With
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(SUB_NAME, strStepName, ex)
		End Try
	End Function

End Class
