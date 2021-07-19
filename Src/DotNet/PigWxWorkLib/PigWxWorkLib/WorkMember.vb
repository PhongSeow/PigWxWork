'**********************************
'* Name: WorkMember
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 企业微信成员
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 5/3/2021
'* 1.0.2  6/3/2021   Add FillPropertiesByJSon
'* 1.0.3  15/4/2021   Use PigToolsWinLib
'* 1.0.4  15/7/2021   Add OpenID
'**********************************

Imports PigToolsWinLib

Public Class WorkMember
	Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.0.4"

	Public Sub New()
		MyBase.New(CLS_VERSION)
	End Sub

	Public Sub New(JSonStr As String)
		MyBase.New(CLS_VERSION)
		Dim strStepName As String = ""
		Try
			strStepName = "New PigJSon"
			Dim oPigJSon As New PigJSon(JSonStr)
			If oPigJSon.LastErr <> "" Then Throw New Exception(oPigJSon.LastErr)
			strStepName = "FillPropertiesByJSon"
			Me.FillPropertiesByJSon(oPigJSon)
			If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("New", ex)
		End Try
	End Sub

	''' <summary>
	''' Oauth2登录后跳转回来的 state 值
	''' </summary>
	Private mstrOauth2State As String
	Public Property Oauth2State() As String
		Get
			Return mstrOauth2State
		End Get
		Friend Set(ByVal value As String)
			mstrOauth2State = value
		End Set
	End Property

	''' <summary>
	''' 是否已获取详细信息
	''' </summary>
	Private mbolIsGetDetInf As Boolean = False
	Public Property IsGetDetInf() As Boolean
		Get
			Return mbolIsGetDetInf
		End Get
		Friend Set(ByVal value As Boolean)
			mbolIsGetDetInf = value
		End Set
	End Property

	''' <summary>
	''' 成员UserID。对应管理端的帐号，企业内必须唯一。不区分大小写，长度为1~64个字节
	''' </summary>
	Private mstrUserId As String
	Public Property UserId() As String
		Get
			If mstrUserId Is Nothing Then mstrUserId = ""
			Return mstrUserId
		End Get
		Friend Set(ByVal value As String)
			mstrUserId = value
		End Set
	End Property

	''' <summary>
	''' 成员名称；第三方不可获取，调用时返回userid以代替name；对于非第三方创建的成员，第三方通讯录应用也不可获取；第三方页面需要通过通讯录展示组件来展示名字
	''' </summary>
	Private mstrName As String
	Public Property Name() As String
		Get
			Return mstrName
		End Get
		Friend Set(ByVal value As String)
			mstrName = value
		End Set
	End Property

	''' <summary>
	''' 手机号码，第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrMobile As String
	Public Property Mobile() As String
		Get
			Return mstrMobile
		End Get
		Friend Set(ByVal value As String)
			mstrMobile = value
		End Set
	End Property

	''' <summary>
	''' 成员所属部门id列表，仅返回该应用有查看权限的部门id
	''' </summary>
	Private mstrDepartment As String
	Public Property Department() As String
		Get
			Return mstrDepartment
		End Get
		Friend Set(ByVal value As String)
			mstrDepartment = value
		End Set
	End Property

	''' <summary>
	''' 部门内的排序值，默认为0。数量必须和department一致，数值越大排序越前面。值范围是[0, 2^32)
	''' </summary>
	Private mstrOrder As String
	Public Property Order() As String
		Get
			Return mstrOrder
		End Get
		Friend Set(ByVal value As String)
			mstrOrder = value
		End Set
	End Property

	''' <summary>
	''' 职务信息；第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrPosition As String
	Public Property Position() As String
		Get
			Return mstrPosition
		End Get
		Friend Set(ByVal value As String)
			mstrPosition = value
		End Set
	End Property

	''' <summary>
	''' 性别。0表示未定义，1表示男性，2表示女性
	''' </summary>
	Private mstrGender As String
	Public Property Gender() As String
		Get
			Return mstrGender
		End Get
		Friend Set(ByVal value As String)
			mstrGender = value
		End Set
	End Property

	''' <summary>
	''' 邮箱，第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrEmail As String
	Public Property Email() As String
		Get
			Return mstrEmail
		End Get
		Friend Set(ByVal value As String)
			mstrEmail = value
		End Set
	End Property

	''' <summary>
	''' 表示在所在的部门内是否为上级。；第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrIs_Leader_In_Dept As String
	Public Property Is_Leader_In_Dept() As String
		Get
			Return mstrIs_Leader_In_Dept
		End Get
		Friend Set(ByVal value As String)
			mstrIs_Leader_In_Dept = value
		End Set
	End Property

	''' <summary>
	''' 头像url。 第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrAvatar As String
	Public Property Avatar() As String
		Get
			Return mstrAvatar
		End Get
		Friend Set(ByVal value As String)
			mstrAvatar = value
		End Set
	End Property

	''' <summary>
	''' 头像缩略图url。第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrThumb_Avatar As String
	Public Property Thumb_Avatar() As String
		Get
			Return mstrThumb_Avatar
		End Get
		Friend Set(ByVal value As String)
			mstrThumb_Avatar = value
		End Set
	End Property

	''' <summary>
	''' 座机。第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrTelephone As String
	Public Property Telephone() As String
		Get
			Return mstrTelephone
		End Get
		Friend Set(ByVal value As String)
			mstrTelephone = value
		End Set
	End Property

	''' <summary>
	''' 别名；第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrAliasName As String
	Public Property AliasName() As String
		Get
			Return mstrAliasName
		End Get
		Friend Set(ByVal value As String)
			mstrAliasName = value
		End Set
	End Property

	''' <summary>
	''' 扩展属性，第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrExtattr As String
	Public Property Extattr() As String
		Get
			Return mstrExtattr
		End Get
		Friend Set(ByVal value As String)
			mstrExtattr = value
		End Set
	End Property

	''' <summary>
	''' 激活状态: 1=已激活，2=已禁用，4=未激活，5=退出企业。
	''' 已激活代表已激活企业微信或已关注微工作台（原企业号）。未激活代表既未激活企业微信又未关注微工作台（原企业号）。
	''' </summary>
	Private mstrStatus As String
	Public Property Status() As String
		Get
			Return mstrStatus
		End Get
		Friend Set(ByVal value As String)
			mstrStatus = value
		End Set
	End Property

	''' <summary>
	''' 员工个人二维码，扫描可添加为外部联系人(注意返回的是一个url，可在浏览器上打开该url以展示二维码)；第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrQr_Code As String
	Public Property Qr_Code() As String
		Get
			Return mstrQr_Code
		End Get
		Friend Set(ByVal value As String)
			mstrQr_Code = value
		End Set
	End Property

	''' <summary>
	''' 成员对外属性，字段详情见对外属性；第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrExternal_Profile As String
	Public Property External_Profile() As String
		Get
			Return mstrExternal_Profile
		End Get
		Friend Set(ByVal value As String)
			mstrExternal_Profile = value
		End Set
	End Property

	''' <summary>
	''' 对外职务，如果设置了该值，则以此作为对外展示的职务，否则以position来展示。第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrExternal_Position As String
	Public Property External_Position() As String
		Get
			Return mstrExternal_Position
		End Get
		Friend Set(ByVal value As String)
			mstrExternal_Position = value
		End Set
	End Property

	''' <summary>
	''' 地址。第三方仅通讯录应用可获取；对于非第三方创建的成员，第三方通讯录应用也不可获取
	''' </summary>
	Private mstrAddress As String
	Public Property Address() As String
		Get
			Return mstrAddress
		End Get
		Friend Set(ByVal value As String)
			mstrAddress = value
		End Set
	End Property

	''' <summary>
	''' 全局唯一。对于同一个服务商，不同应用获取到企业内同一个成员的open_userid是相同的，最多64个字节。仅第三方应用可获取
	''' </summary>
	Private mstrOpen_UserId As String
	Public Property Open_UserId() As String
		Get
			Return mstrOpen_UserId
		End Get
		Friend Set(ByVal value As String)
			mstrOpen_UserId = value
		End Set
	End Property

	''' <summary>
	''' 主部门
	''' </summary>
	Private mstrMain_Department As String
	Public Property Main_Department() As String
		Get
			Return mstrMain_Department
		End Get
		Friend Set(ByVal value As String)
			mstrMain_Department = value
		End Set
	End Property

	Friend Sub FillPropertiesByJSon(PigJSon As PigJSon)
		Try
			With Me
				If .UserId.Length = 0 Then
					.UserId = PigJSon.GetStrValue("userid")
					If .UserId.Length = 0 Then Throw New Exception("UserId未初始化")
				ElseIf .UserId <> PigJSon.GetStrValue("userid") Then
					Throw New Exception("UserId与JSon中的userid不匹配")
				End If
				.Name = PigJSon.GetStrValue("name")
				.Mobile = PigJSon.GetStrValue("mobile")
				.Department = PigJSon.GetStrValue("department")
				.Order = PigJSon.GetStrValue("order")
				.Position = PigJSon.GetStrValue("position")
				.Gender = PigJSon.GetStrValue("gender")
				.Email = PigJSon.GetStrValue("email")
				.Is_Leader_In_Dept = PigJSon.GetStrValue("is_leader_in_dept")
				.Avatar = PigJSon.GetStrValue("avatar")
				.Thumb_Avatar = PigJSon.GetStrValue("thumb_avatar")
				.Telephone = PigJSon.GetStrValue("telephone")
				.AliasName = PigJSon.GetStrValue("alias")
				.Extattr = PigJSon.GetStrValue("extattr")
				.Status = PigJSon.GetStrValue("status")
				.Qr_Code = PigJSon.GetStrValue("qr_code")
				.External_Profile = PigJSon.GetStrValue("external_profile")
				.External_Position = PigJSon.GetStrValue("external_position")
				.Address = PigJSon.GetStrValue("address")
				.Open_UserId = PigJSon.GetStrValue("open_userid")
				.Main_Department = PigJSon.GetStrValue("main_department")
				.IsGetDetInf = True
			End With
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf("FillPropertiesByJSon", ex)
		End Try
	End Sub

	''' <summary>
	''' 手机设备号(由企业微信在安装时随机生成，删除重装会改变，升级不受影响)
	''' </summary>
	Private mstrDeviceId As String
	Public Property DeviceId() As String
		Get
			Return mstrDeviceId
		End Get
		Friend Set(ByVal value As String)
			mstrDeviceId = value
		End Set
	End Property

	''' <summary>
	''' 
	''' </summary>
	Private mstrOpenID As String
	Public Property OpenID() As String
		Get
			Return mstrOpenID
		End Get
		Friend Set(ByVal value As String)
			mstrOpenID = value
		End Set
	End Property
End Class
