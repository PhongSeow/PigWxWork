'**********************************
'* Name: WorkMember
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 企业微信消息
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 13/9/2021
'* 1.1  14/9/2021	Modify PostJSon
'**********************************
Imports PigToolsWinLib

Public Class WorkMessage
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.1.1"
	Public Sub New(AgentId As String)
		MyBase.New(CLS_VERSION)
		Me.AgentId = AgentId
	End Sub

	Public Enum MsgTypeEnum
		Text = 0
		Image = 10
		Voice = 20
		Video = 30
		File = 40
		TextCard = 50
		News = 60
		MpNews = 70
		Markdown = 80
		MiniProgram_Notice = 90
		Template_Card = 100
	End Enum

	''' <summary>
	''' 指定接收消息的成员，成员ID列表（多个接收者用‘|’分隔，最多支持1000个）。
	''' 特殊情况：指定为”@all”，则向该企业应用的全部成员发送
	''' </summary>
	''' <returns></returns>
	Public Property ToUser As String = ""

	''' <summary>
	''' 指定接收消息的部门，部门ID列表，多个接收者用‘|’分隔，最多支持100个。
	''' 当touser为”@all”时忽略本参数
	''' </summary>
	''' <returns></returns>
	Public Property ToParty As String = ""

	''' <summary>
	''' 指定接收消息的标签，标签ID列表，多个接收者用‘|’分隔，最多支持100个。
	''' 当touser为”@all”时忽略本参数
	''' </summary>
	''' <returns></returns>
	Public Property ToTag As String = ""

	''' <summary>
	''' 消息类型
	''' </summary>
	''' <returns></returns>
	Public Property MsgType As MsgTypeEnum
	Private ReadOnly Property mMsgTypeStr As String
		Get
			Return LCase(Me.MsgType.ToString)
		End Get
	End Property

	''' <summary>
	''' 企业应用的id，整型。企业内部开发，可在应用的设置页面查看；第三方服务商，可通过接口 获取企业授权信息 获取该参数值
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property AgentId As String

	''' <summary>
	''' 消息内容，最长不超过2048个字节，超过将截断（支持id转译）
	''' </summary>
	''' <returns></returns>
	Public Property Content As String

	''' <summary>
	''' 表示是否是保密消息，0表示可对外分享，1表示不能分享且内容显示水印，默认为0
	''' </summary>
	''' <returns></returns>
	Public Property Safe As Boolean = False

	''' <summary>
	''' 表示是否开启id转译，0表示否，1表示是，默认0。仅第三方应用需要用到，企业自建应用可以忽略。
	''' </summary>
	''' <returns></returns>
	Public Property Enable_Id_Trans As Boolean = False

	''' <summary>
	''' 表示是否开启重复消息检查，0表示否，1表示是，默认0
	''' </summary>
	''' <returns></returns>
	Public Property Enable_Duplicate_Check As Boolean = False

	''' <summary>
	''' 表示是否重复消息检查的时间间隔，默认1800s，最大不超过4小时
	''' </summary>
	''' <returns></returns>
	Public Property Duplicate_Check_Interval As Integer = 0


	''' <summary>
	''' 是否 touser、toparty、totag 全为空
	''' </summary>
	''' <returns></returns>
	Public ReadOnly Property IsToAllEmpty As Boolean
		Get
			If Me.ToUser = "" And Me.ToParty = "" And Me.ToTag = "" Then
				Return True
			Else
				Return False
			End If
		End Get
	End Property

	''' <summary>
	''' 不合法的userid，不区分大小写，统一转为小写
	''' </summary>
	''' <returns></returns>
	Private mstrInvalidUser As String
	Public Property InvalidUser As String
		Get
			Return mstrInvalidUser
		End Get
		Friend Set(value As String)
			mstrInvalidUser = value
		End Set
	End Property

	''' <summary>
	''' 不合法的partyid
	''' </summary>
	''' <returns></returns>
	Private mstrInvalidParty As String
	Public Property InvalidParty As String
		Get
			Return mstrInvalidParty
		End Get
		Friend Set(value As String)
			mstrInvalidParty = value
		End Set
	End Property

	''' <summary>
	''' 不合法的partyid
	''' </summary>
	''' <returns></returns>
	Private mstrInvalidTag As String
	Public Property InvalidTag As String
		Get
			Return mstrInvalidTag
		End Get
		Friend Set(value As String)
			mstrInvalidTag = value
		End Set
	End Property

	Public ReadOnly Property PostJSon As String
		Get
			Try
				Dim pjMain As New PigJSon
				Dim pjContent As New PigJSon
				With pjMain
					.AddEle("msgtype", Me.mMsgTypeStr, True)
					If Me.ToUser <> "" Then .AddEle("touser", Me.ToUser)
					If Me.ToParty <> "" Then .AddEle("toparty", Me.ToParty)
					If Me.ToTag <> "" Then .AddEle("totag", Me.ToTag)
					.AddEle("agentid", Me.AgentId)
					Select Case Me.MsgType
						Case MsgTypeEnum.Text
							pjContent.AddEle("content", Me.Content, True)
							pjContent.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
							.AddOneObjectEle("text", pjContent.MainJSonStr)
					End Select
					If Me.Safe Then .AddEle("safe", "1")
					If Me.Enable_Id_Trans Then .AddEle("enable_id_trans", "1")
					If Me.Duplicate_Check_Interval > 0 Then .AddEle("duplicate_check_interval", Me.Duplicate_Check_Interval)
					.AddSymbol(PigJSon.xpSymbolType.EleEndFlag)
				End With
				PostJSon = pjMain.MainJSonStr
				pjContent = Nothing
				pjMain = Nothing
			Catch ex As Exception
				Me.SetSubErrInf("PostJSon.Get", ex)
				Return ""
			End Try
		End Get
	End Property

End Class
