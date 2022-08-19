'**********************************
'* Name: 豚豚配置应用|PigConfigApp
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 配置应用类|Configure application classes
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.13
'* Create Time: 18/12/2021
'* 1.1    20/12/2020   Add mNew,MkEncKey,mLoadConfig,GetEncStr
'* 1.2    21/12/2020   Modify mLoadConfig,EnmSaveType,mNew, add LoadConfig,LoadConfigFile,SaveConfigFile,SaveConfig,PigConfigSessions,AddNewConfigSession
'* 1.3    22/12/2020   Modify mNew,SaveConfig,mSaveConfig,SaveConfigFile, add TextType,fGetUnEncStr,LoadConfigFile,mLoadConfig
'* 1.4    23/12/2020   Modify mNew,mLoadConfig,GetPigConfigSession,MkEncKey
'* 1.5    24/12/2020   Modify GetPigConfig,mNew,mLoadConfig,MkEncKey, Add ConfStrValue,ConfIntValue,ConfBoolValue,ConfDateValue,ConfDecValue
'* 1.6    25/12/2020   Modify mLoadConfig
'* 1.7    26/12/2020   Add IsClearFrist, modify mLoadConfig
'* 1.8    1/1/2021     Modify mSaveConfig,mLoadConfig
'* 1.9    3/1/2021     Modify mSaveConfig
'* 1.10   2/2/2022     Add IsChange
'* 1.11   21/2/2022    Modify mLoadConfig,IsChange,SaveConfigFile,SaveConfig, add fSetIsChangeFalse
'* 1.12   9/5/2022     Remove fSetIsChangeFalse, modify fSaveCurrMD5
'* 1.13   10/5/2022    Modify mLoadConfig,IsLoadConfClearFrist
'**********************************
Imports Microsoft.VisualBasic
Public Class PigConfigApp
    Inherits PigBaseMini
	Private Const CLS_VERSION As String = "1.13.6"
	Public Enum EnmSaveType
		''' <summary>
		''' XML text
		''' </summary>
		Xml = 0
		''' <summary>
		''' Encrypted save
		''' </summary>
		EncData = 1
		''' <summary>
		''' Ini text
		''' </summary>
		Ini = 2
		''' <summary>
		''' Registry, windows platform only
		''' </summary>
		Registry = 3
	End Enum

	Private mprEnc As PigRsa
	Private mpaEnc As PigAes
	Private moPigFunc As New PigFunc
	Private moPigConfigSessions As PigConfigSessions
	Public Property PigConfigSessions As PigConfigSessions
		Get
			Return moPigConfigSessions
		End Get
		Friend Set(value As PigConfigSessions)
			moPigConfigSessions = value
		End Set
	End Property


	Public Sub New()
		MyBase.New(CLS_VERSION)
		Me.mNew("")
	End Sub

	Public Sub New(TextType As PigText.enmTextType)
		MyBase.New(CLS_VERSION)
		Me.mNew("", TextType)
	End Sub

	Public Sub New(EncKey As String, TextType As PigText.enmTextType)
		MyBase.New(CLS_VERSION)
		Me.mNew(EncKey, TextType)
	End Sub

	Public Sub New(EncKey As String)
		MyBase.New(CLS_VERSION)
		Me.mNew(EncKey)
	End Sub

	Private Sub mNew(EncKey As String, Optional TextType As PigText.enmTextType = PigText.enmTextType.UTF8)
		Dim LOG As New PigStepLog("mNew")
		Try
			mpaEnc = New PigAes
			LOG.StepName = "PigAes.LoadEncKey"
			Dim strEncKey As String = "zLZ+47SGy/2pp2glqxlXnUKYUmbmq5+B28w8qCE+K2R0vScuXI1zBFGuEW9TH7vo6uL7sY6yR4eD0+XzykWPax3DzzTB1xxOtvi5dvxqZO6T6aO3N7jw0iwdjflySzhxxQ54O2nexFVe/2J8pHAMOTszTL7fA5jY1CQ2RBuW4LxuFCCFEihtWwdJxsdNd7U/oDqwCIgJQSN9KppKf5yJwhZfhYy6LxWmr1BvzfXBXUZIQFkPMFSlKSV1Q1EoozHJv2UtkmM1eqxYVhcYnsgGkPGLAVmVWk4winvrm9Dt2hNnbNbdC4AiswChwB5GEq6XmZGqYD1h9noygpSiefSCGg=="
			LOG.Ret = mpaEnc.LoadEncKey(strEncKey)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			mprEnc = New PigRsa
			If EncKey = "" Then
				EncKey = "Uo0ltXATzVUmxfeOzTndcp9Xt9FcHSq62dqHhJyWSPedC1csnw5z4AExmOywhc3XCGM6iiQBxznelkGwbVHlx5n7N8Y5ZzY3PI3NsqJFHVtmRjkRc7eFwF5ADQKZ5t+lVTMinIK2mccGJvxHaeu/ocmn91OdzbtDpaItpvUzS49iwq9GWoePCZXl1CTiZ2eVEL1tuZTji4pXd7WH7Q+Ti5vlOKYMvmkxcdf4nkiif8gZyRR2pXb80JlkJFor14UuFgJcIMHnOIZEwtkU3mhcgD7DELdFo9BLwA/ohzI8kwlC5slZ1TJNd43Xgjjz6pqP3DqL5z2xUsZxV2736IRxQCLuv1BYA/90QUEyrGDdiQNf+tiL5e+nFWNC1ZRFETwbYS/CB6JwqUKdyPYyY44KNmvIModb+IMY4SLz2HropLAGVCqroCrM7xxQrCCgeAuIa4SYnzSiB+9YG31sFQFXpuPQ54Slva/ppS5I0Xko3cxoPOv7jNPSKyK8RRtvYvcGVI09dJMQkYjZHGkuJKkKLdAEub0Tmrnm8h4vUlbe/dc2c3yKzIUx9Mhd8icaD3Ucxn6/3BIfKPuDZ4S8rTdhoTb2IenOZ23e0h1xgk5hU5S8u9A9+By9W/3ctevOFOCfN4ZLcenuSpAqEgGefv2wS05IMidhc3c92RdITXc/ZUjMTc3mZe8qw45eq6NdH6Q8y0S8OoB4/3lFdQeQcL/+WeX1FSGLFm2v+pEIe4285g+1XcqBmJS9E6A9SlzTAwXEFcohrNht4VJl6uZ4CPsmLpbb4G0FlZO340GEfvtEBz5D9anW6Ny3tSApl9jStlhuX5RZmB1micLBRuAncBOs656oI1/DSBfxbi/Yw3UB2zYUqTO7V9CmtQq0W0jm8jBdgAAsnNjMehqE9QFnohRaDNIK0JUpA/3VXVS7vmkMC2rBLp11n5SkkAQk5OS4G61UBEWTMRwtf908VOh1JlfzvUKSUdSDDp7IjAyij/yOpSWFayLb5lGtT3acGoCLdUyZecoYVtTg9grE5Sn4NbunOFTLtfb/dfaKw35hvWDK8RWPKSVZ5zqp+WzyTttwUCtBgTdETkQ0tp4FGQTzoYuEb/0Vkn0JCBtLKBYkWemNkhdpiwW/7Yfcp5R5URIMB1BgymM2og5nn2bq9alkeVFhdlldP5PhxzVsnDUT4UB/cQGQ0kP46bfhPFVBqWLS0ISwuZwq+oBDQSvXKI+MiPGtmA=="
				LOG.StepName = "PigAes.Decrypt"
				LOG.Ret = mpaEnc.Decrypt(EncKey, strEncKey, PigText.enmTextType.UTF8)
				If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Else
				LOG.StepName = "New PigBytes(EncKey)"
				Dim oPigBytes As New PigBytes(EncKey)
				If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
				LOG.StepName = "PigAes.Decrypt"
				LOG.Ret = mpaEnc.Decrypt(oPigBytes.Main, strEncKey, Me.TextType)
				If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
				oPigBytes = Nothing
			End If
			LOG.StepName = "PigRsa.LoadPubKey"
			LOG.Ret = mprEnc.LoadPubKey(strEncKey)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			LOG.StepName = "New PigConfigSessions"
			Me.PigConfigSessions = New PigConfigSessions(Me)
			If Me.PigConfigSessions.LastErr <> "" Then Throw New Exception(Me.PigConfigSessions.LastErr)
			LOG.StepName = "PigConfigSessions.Add(Main)"
			Me.PigConfigSessions.Add("Main")
			If Me.PigConfigSessions.LastErr <> "" Then Throw New Exception(Me.PigConfigSessions.LastErr)
			Me.TextType = TextType
			Me.ClearErr()
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Sub

	''' <summary>
	''' 生成用于加密内容的密钥，格式为Base64|Generate a key for encrypting content in Base64 format
	''' </summary>
	''' <returns></returns>
	Public Function MkEncKey(ByRef Base64EncKey As String) As String
		Dim LOG As New PigStepLog("MkEncKey")
		Try
			LOG.StepName = "PigRsa.MkPubKey"
			Dim strRsaXmlKey As String = ""
			Dim oPigRsa As New PigRsa
			LOG.Ret = oPigRsa.MkPubKey(True, strRsaXmlKey)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			oPigRsa = Nothing
			LOG.StepName = "New PigText"
			Dim oPigText As New PigText(strRsaXmlKey, Me.TextType)
			If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
			LOG.StepName = "New PigBytes"
			Dim oPigBytes As New PigBytes(oPigText.TextBytes)
			If oPigBytes.LastErr <> "" Then Throw New Exception(oPigBytes.LastErr)
			LOG.StepName = "PigAes.Encrypt"
			LOG.Ret = mpaEnc.Encrypt(oPigBytes.Main, Base64EncKey)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			oPigBytes = Nothing
			oPigText = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function


	Public Function LoadConfigFile(FilePath As String, SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("LoadConfig")
		Try
			LOG.StepName = "New PigFile"
			Dim oPigFile As New PigFile(FilePath)
			If oPigFile.LastErr <> "" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(oPigFile.LastErr)
			End If
			LOG.StepName = "PigFile.LoadFile"
			LOG.Ret = oPigFile.LoadFile()
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(LOG.Ret)
			End If
			If oPigFile.GbMain Is Nothing Then oPigFile.GbMain = New PigBytes
			LOG.StepName = "mLoadConfig"
			LOG.Ret = Me.mLoadConfig(oPigFile.GbMain.Main, SaveType)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(LOG.Ret)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function LoadConfig(ConfData As String, SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("LoadConfig.Base64")
		Try
			LOG.StepName = "New PigText"
			Dim oPigText As New PigText(ConfData, Me.TextType)
			If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
			LOG.StepName = "mLoadConfig"
			LOG.Ret = Me.mLoadConfig(oPigText.TextBytes, SaveType)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function LoadConfig(ConfData As Byte(), SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("LoadConfig.Bytes")
		Try
			LOG.StepName = "mLoadConfig"
			LOG.Ret = Me.mLoadConfig(ConfData, SaveType)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mLoadConfig(AbConfData As Byte(), SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("mLoadConfig")
		Try
			LOG.StepName = "Check SaveType"
			Select Case SaveType
				Case EnmSaveType.Registry
					If Me.IsWindows = False Then
						Throw New Exception(SaveType.ToString & ", which can only be used on Windows platforms")
					Else
						Throw New Exception(SaveType.ToString & " support coming soon")
					End If
				Case EnmSaveType.EncData
					Throw New Exception(SaveType.ToString & " support coming soon")
				Case EnmSaveType.Xml, EnmSaveType.Ini
				Case Else
					Throw New Exception(SaveType.ToString & " is invalid")
			End Select
			If Me.PigConfigSessions Is Nothing Then
				LOG.StepName = "New PigConfigSessions"
				Me.PigConfigSessions = New PigConfigSessions(Me)
				If Me.PigConfigSessions.LastErr <> "" Then Throw New Exception(Me.PigConfigSessions.LastErr)
			ElseIf Me.IsLoadConfClearFrist = True Then
				LOG.StepName = "PigConfigSessions.Clear"
				Me.PigConfigSessions.Clear()
				If Me.PigConfigSessions.LastErr <> "" Then Throw New Exception(Me.PigConfigSessions.LastErr)
			End If
			LOG.StepName = "PigConfigSessions.Add(Main)"
			Me.PigConfigSessions.AddOrGet("Main")
			If Me.PigConfigSessions.LastErr <> "" Then Throw New Exception(Me.PigConfigSessions.LastErr)
			LOG.StepName = "New PigText"
			Dim ptConfData As New PigText(AbConfData, Me.TextType)
			If ptConfData.LastErr <> "" Then Throw New Exception(ptConfData.LastErr)
			Select Case SaveType
				Case EnmSaveType.Registry
				Case EnmSaveType.Ini
					Dim strData As String = ptConfData.Text
					Dim strLastLine As String = "", strCurrLine As String, strSessionName As String = "", strTmpLine As String = "", strTmpSession As String = ""
					Dim strLeft As String, strRight As String, strDataTmp As String
					'数据清理
					Do While True
						If InStr(strData, vbCrLf) = 0 Then Exit Do
						strData = Replace(strData, vbCrLf, vbLf)
					Loop
					Do While True
						If InStr(strData, vbCr) = 0 Then Exit Do
						strData = Replace(strData, vbCr, vbLf)
					Loop
					Do While True
						If InStr(strData, vbLf & vbLf) = 0 Then Exit Do
						strData = Replace(strData, vbLf & vbLf, vbLf)
					Loop
					strDataTmp = strData
					If Right(strDataTmp, 1) <> vbLf Then strDataTmp &= vbLf
					strData = ""
					Dim dteBegin As DateTime = Now
					LOG.StepName = "Re merge strData"
					Do While True
						strTmpLine = Trim(moPigFunc.GetStr(strDataTmp, "", vbLf))
						Select Case strTmpLine
							Case "", vbTab
							Case Else
								strData &= strTmpLine & vbLf
						End Select
						If strDataTmp = "" Then Exit Do
						Dim tsTimeDiff As TimeSpan = Now - dteBegin
						If tsTimeDiff.Milliseconds > 500 Then
							Me.PrintDebugLog(LOG.StepName, "Run timeout")
							Exit Do
						End If
					Loop
					If Right(strData, 1) <> vbLf Then strData &= vbLf
					Dim oPigConfigSession As PigConfigSession = Nothing
					Do While True
						strCurrLine = Trim(moPigFunc.GetStr(strData, "", vbLf))
						If strCurrLine = "" Then Exit Do
						strTmpLine = Trim(strCurrLine) & vbLf
						strLeft = "["
						strRight = "]" & vbLf
						If Left(strTmpLine, 1) = strLeft And Right(strTmpLine, Len(strRight)) = strRight Then
							strTmpSession = moPigFunc.GetStr(strTmpLine, strLeft, strRight)
							If strTmpSession <> strSessionName Then
								strSessionName = strTmpSession
								LOG.StepName = "PigConfigSessions.AddOrGet"
								oPigConfigSession = Me.PigConfigSessions.AddOrGet(strSessionName)
								If Me.PigConfigSessions.LastErr <> "" Then
									LOG.AddStepNameInf(strSessionName)
									Throw New Exception(Me.PigConfigSessions.LastErr)
								ElseIf oPigConfigSession Is Nothing Then
									LOG.AddStepNameInf(strSessionName)
									Throw New Exception("oPigConfigSession Is Nothing")
								End If
								If strLastLine <> "" Then
									Select Case Left(strLastLine, 1)
										Case ";", "#"
											oPigConfigSession.SessionDesc = Mid(strLastLine, 2)
									End Select
								End If
							End If
						Else
							Select Case Left(strTmpLine, 1)
								Case ";", "#"
								Case Else
									If oPigConfigSession IsNot Nothing Then
										If InStr(strTmpLine, "=") > 0 Then
											Dim strConfName As String, strConfValue As String
											strConfName = moPigFunc.GetStr(strTmpLine, "", "=")
											strConfValue = moPigFunc.GetStr(strTmpLine, "", vbLf)
											LOG.StepName = "PigConfigs.AddOrGet"
											Dim oPigConfig As PigConfig = oPigConfigSession.PigConfigs.AddOrGet(strConfName, strConfValue)
											If oPigConfigSession.PigConfigs.LastErr <> "" Then
												LOG.AddStepNameInf(strSessionName)
												LOG.AddStepNameInf(strConfName)
												Throw New Exception(oPigConfigSession.PigConfigs.LastErr)
											ElseIf oPigConfig Is Nothing Then
												LOG.AddStepNameInf(strSessionName)
												LOG.AddStepNameInf(strConfName)
												Throw New Exception("oPigConfig Is Nothing")
											End If
											oPigConfig.ConfValue = strConfValue
											Select Case Left(strLastLine, 1)
												Case ";", "#"
													oPigConfig.ConfDesc = Mid(strLastLine, 2)
											End Select
										End If
									End If
							End Select
						End If
						strLastLine = strCurrLine
					Loop
				Case EnmSaveType.EncData
				Case EnmSaveType.Xml
					Dim pxMain As New PigXml(False)
					LOG.StepName = "SetMainXml"
					pxMain.SetMainXml(ptConfData.Text)
					If pxMain.LastErr <> "" Then
						LOG.AddStepNameInf("Main")
						Me.PrintDebugLog(LOG.SubName, LOG.StepName, ptConfData.Text)
						Throw New Exception(pxMain.LastErr)
					End If
					Do While True
						Dim strSession As String = ""
						strSession = pxMain.XmlGetStr("Session")
						If strSession = "" Then Exit Do
						LOG.StepName = "SetMainXml"
						Dim pxSession As New PigXml(False)
						pxSession.SetMainXml(strSession)
						If pxSession.LastErr <> "" Then
							LOG.AddStepNameInf("Session")
							Me.PrintDebugLog(LOG.SubName, LOG.StepName, strSession)
							Throw New Exception(pxSession.LastErr)
						End If
						Dim strSessionName As String = pxSession.XmlGetStr("SessionName")
						LOG.StepName = "PigConfigSessions.AddOrGet"
						Dim oPigConfigSession As PigConfigSession = Me.PigConfigSessions.AddOrGet(strSessionName)
						If Me.PigConfigSessions.LastErr <> "" Then
							LOG.AddStepNameInf(strSessionName)
							Throw New Exception(Me.PigConfigSessions.LastErr)
						ElseIf oPigConfigSession Is Nothing Then
							LOG.AddStepNameInf(strSessionName)
							Throw New Exception("oPigConfigSession Is Nothing")
						End If
						If oPigConfigSession.PigConfigs Is Nothing Then
							LOG.StepName = "New PigConfigs"
							oPigConfigSession.PigConfigs = New PigConfigs(oPigConfigSession)
							If oPigConfigSession.PigConfigs.LastErr <> "" Then Throw New Exception(oPigConfigSession.PigConfigs.LastErr)
						End If
						Do While True
							Dim strConfItem As String = ""
							strConfItem = pxSession.XmlGetStr("ConfItem")
							If strConfItem = "" Then Exit Do
							LOG.StepName = "SetMainXml"
							Dim pxConfItem As New PigXml(False)
							pxConfItem.SetMainXml(strConfItem)
							If pxConfItem.LastErr <> "" Then
								LOG.AddStepNameInf("ConfItem")
								Me.PrintDebugLog(LOG.SubName, LOG.StepName, ptConfData.Text)
								Throw New Exception(pxConfItem.LastErr)
							End If
							Dim strConfName As String = pxConfItem.XmlGetStr("ConfName")
							LOG.StepName = "PigConfigs.AddOrGet"
							Dim oPigConfig As PigConfig = oPigConfigSession.PigConfigs.AddOrGet(strConfName)
							If oPigConfigSession.PigConfigs.LastErr <> "" Then
								LOG.AddStepNameInf(strConfName)
								Throw New Exception(oPigConfigSession.PigConfigs.LastErr)
							ElseIf oPigConfig Is Nothing Then
								LOG.AddStepNameInf(strConfName)
								Throw New Exception("oPigConfig Is Nothing")
							End If
							With oPigConfig
								.ConfValue = pxConfItem.XmlGetStr("ConfValue")
								.ConfDesc = pxConfItem.XmlGetStr("ConfDesc")
							End With
						Loop
					Loop
			End Select
			LOG.StepName = "fSaveCurrMD5"
			LOG.Ret = Me.fSaveCurrMD5
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			ptConfData = Nothing
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Friend Function fGetUnEncStr(ByRef OutUnEncStr As String, EncStrBase64 As String) As String
		Dim LOG As New PigStepLog("fGetUnEncStr")
		Try
			If Mid(EncStrBase64, 1, 5) <> "{Enc}" Then Throw New Exception("Invalid EncStrBase64")
			EncStrBase64 = Mid(EncStrBase64, 6)
			LOG.StepName = "Decrypt"
			LOG.Ret = Me.mprEnc.Decrypt(EncStrBase64, OutUnEncStr, Me.TextType)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			OutUnEncStr = ""
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function GetEncStr(SrcStr As String) As String
		Dim LOG As New PigStepLog("GetEncStr")
		Try
			Dim strEncStr As String = ""
			LOG.StepName = "New PigText"
			Dim oPigText As New PigText(SrcStr, Me.TextType)
			If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
			LOG.StepName = "Encrypt"
			LOG.Ret = Me.mprEnc.Encrypt(oPigText.TextBytes, strEncStr)
			strEncStr = "{Enc}" & strEncStr
			oPigText = Nothing
			Return strEncStr
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Private Function mSaveConfig(ByRef OutConfData As Byte(), SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("mSaveConfig")
		Try
			LOG.StepName = "Check SaveType"
			Select Case SaveType
				Case EnmSaveType.Registry
					If Me.IsWindows = False Then
						Throw New Exception(SaveType.ToString & ", which can only be used on Windows platforms")
					Else
						Throw New Exception(SaveType.ToString & " support coming soon")
					End If
				Case EnmSaveType.EncData
					Throw New Exception(SaveType.ToString & " support coming soon")
				Case EnmSaveType.Xml, EnmSaveType.Ini
				Case Else
					Throw New Exception(SaveType.ToString & " is invalid")
			End Select
			Dim strConfData As String = ""
			Select Case SaveType
				Case EnmSaveType.Registry
				Case EnmSaveType.Ini
					LOG.StepName = "For Each PigConfigSessions"
					Dim strRemFrist As String, strOsCrLf As String = Me.OsCrLf
					If Me.IsWindows = True Then
						strRemFrist = ";"
					Else
						strRemFrist = "#"
					End If
					For Each oPigConfigSession As PigConfigSession In Me.PigConfigSessions
						If oPigConfigSession.SessionDesc <> "" Then strConfData &= strRemFrist & oPigConfigSession.SessionDesc & strOsCrLf
						strConfData &= "[" & oPigConfigSession.SessionName & "]" & strOsCrLf
						For Each oPigConfig As PigConfig In oPigConfigSession.PigConfigs
							If oPigConfig.ConfDesc <> "" Then strConfData &= strRemFrist & oPigConfig.ConfDesc & strOsCrLf
							strConfData &= oPigConfig.ConfName & "=" & oPigConfig.fConfValue & strOsCrLf
						Next
						strConfData &= strOsCrLf
					Next
					LOG.StepName = "New PigText(Ini)"
					Dim oPigText As New PigText(strConfData, Me.TextType)
					OutConfData = oPigText.TextBytes
					oPigText = Nothing
				Case EnmSaveType.EncData
				Case EnmSaveType.Xml
					LOG.StepName = "For Each PigConfigSessions"
					Dim pxMain As New PigXml(True)
					For Each oPigConfigSession As PigConfigSession In Me.PigConfigSessions
						pxMain.AddEleLeftSign("Session")
						pxMain.AddEle("SessionName", oPigConfigSession.SessionName, 1)
						pxMain.AddEle("SessionDesc", oPigConfigSession.SessionDesc, 1)
						For Each oPigConfig As PigConfig In oPigConfigSession.PigConfigs
							pxMain.AddEleLeftSign("ConfItem", 1)
							pxMain.AddEle("ConfName", oPigConfig.ConfName, 2)
							pxMain.AddEle("ConfValue", oPigConfig.fConfValue, 2)
							pxMain.AddEle("ConfDesc", oPigConfig.ConfDesc, 2)
							pxMain.AddEleRightSign("ConfItem", 1)
						Next
						pxMain.AddEleRightSign("Session")
					Next
					strConfData = pxMain.MainXmlStr
					If pxMain.LastErr <> "" Then Throw New Exception(pxMain.LastErr)
					LOG.StepName = "New PigText(Xml)"
					Dim oPigText As New PigText(strConfData, Me.TextType)
					OutConfData = oPigText.TextBytes
					oPigText = Nothing
			End Select
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function SaveConfigFile(FilePath As String, SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("SaveConfigFile")
		Try
			LOG.StepName = "New PigFile"
			Dim oPigFile As New PigFile(FilePath)
			If oPigFile.LastErr <> "" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(oPigFile.LastErr)
			End If
			If oPigFile.GbMain Is Nothing Then oPigFile.GbMain = New PigBytes
			LOG.StepName = "mSaveConfig"
			LOG.Ret = Me.mSaveConfig(oPigFile.GbMain.Main, SaveType)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(LOG.Ret)
			End If
			LOG.StepName = "SaveFile"
			LOG.Ret = oPigFile.SaveFile(False)
			If LOG.Ret <> "OK" Then
				LOG.AddStepNameInf(FilePath)
				Throw New Exception(LOG.Ret)
			End If
			LOG.StepName = "fSaveCurrMD5"
			LOG.Ret = Me.fSaveCurrMD5()
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function SaveConfig(ByRef OutConfData As String, SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("SaveConfig.Base64")
		Try
			Dim abConfData As Byte()
			ReDim abConfData(0)
			LOG.StepName = "mSaveConfig"
			LOG.Ret = Me.mSaveConfig(abConfData, SaveType)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			LOG.StepName = "New PigText"
			Dim oPigText As New PigText(abConfData, Me.TextType)
			If oPigText.LastErr <> "" Then Throw New Exception(oPigText.LastErr)
			OutConfData = oPigText.Text
			oPigText = Nothing
			LOG.StepName = "fSaveCurrMD5"
			LOG.Ret = Me.fSaveCurrMD5()
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function SaveConfig(ByRef OutConfData As Byte(), SaveType As EnmSaveType) As String
		Dim LOG As New PigStepLog("SaveConfig.Base64")
		Try
			Dim oPigBytes As New PigBytes
			LOG.StepName = "mSaveConfig"
			LOG.Ret = Me.mSaveConfig(OutConfData, SaveType)
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			OutConfData = oPigBytes.Main
			oPigBytes = Nothing
			LOG.StepName = "fSaveCurrMD5"
			LOG.Ret = Me.fSaveCurrMD5()
			If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	Public Function GetPigConfigSession(SessionName As String) As PigConfigSession
		Dim LOG As New PigStepLog("GetPigConfigSession")
		Try
			If Me.PigConfigSessions Is Nothing Then
				LOG.StepName = "New PigConfigSessions"
				Me.PigConfigSessions = New PigConfigSessions(Me)
				If Me.PigConfigSessions.LastErr <> "" Then Throw New Exception(Me.PigConfigSessions.LastErr)
			End If
			LOG.StepName = "Return PigConfigSession"
			Return Me.PigConfigSessions.Item(SessionName)
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function

	Public Function GetPigConfig(SessionName As String, ConfName As String) As PigConfig
		Return Me.mGetPigConfig(SessionName, ConfName)
	End Function

	Public Function GetPigConfig(ConfName As String) As PigConfig
		Return Me.mGetPigConfig("", ConfName)
	End Function

	Private Function mGetPigConfig(SessionName As String, ConfName As String) As PigConfig
		Dim LOG As New PigStepLog("mGetPigConfig")
		Try
			If SessionName = "" Then SessionName = "Main"
			LOG.StepName = "GetPigConfigSession"
			Dim oPigConfigSession As PigConfigSession = Me.GetPigConfigSession(SessionName)
			If oPigConfigSession Is Nothing Then
				LOG.AddStepNameInf(SessionName)
				Throw New Exception("oPigConfigSession Is Nothing")
			End If
			LOG.StepName = "Return PigConfig"
			Return oPigConfigSession.PigConfigs.Item(ConfName)
		Catch ex As Exception
			Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
			Return Nothing
		End Try
	End Function

	''' <summary>
	''' 文本类型
	''' </summary>
	Private mintTextType As PigText.enmTextType = PigText.enmTextType.UTF8
	Public Property TextType As PigText.enmTextType
		Get
			Return mintTextType
		End Get
		Set(value As PigText.enmTextType)
			mintTextType = value
		End Set
	End Property

	Public ReadOnly Property ConfStrValue(ConfName As String) As String
		Get
			If Me.PigConfigSessions.IsItemExists("Main") = False Then
				Return ""
			Else
				Return Me.PigConfigSessions.Item("Main").ConfStrValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfStrValue(SessionName As String, ConfName As String) As String
		Get
			If Me.PigConfigSessions.IsItemExists(SessionName) = False Then
				Return ""
			Else
				Return Me.PigConfigSessions.Item(SessionName).ConfStrValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfIntValue(ConfName As String) As Long
		Get
			If Me.PigConfigSessions.IsItemExists("Main") = False Then
				Return 0
			Else
				Return Me.PigConfigSessions.Item("Main").ConfIntValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfIntValue(SessionName As String, ConfName As String) As Long
		Get
			If Me.PigConfigSessions.IsItemExists(SessionName) = False Then
				Return 0
			Else
				Return Me.PigConfigSessions.Item(SessionName).ConfIntValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfBoolValue(ConfName As String) As Boolean
		Get
			If Me.PigConfigSessions.IsItemExists("Main") = False Then
				Return False
			Else
				Return Me.PigConfigSessions.Item("Main").ConfBoolValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfBoolValue(SessionName As String, ConfName As String) As Boolean
		Get
			If Me.PigConfigSessions.IsItemExists(SessionName) = False Then
				Return False
			Else
				Return Me.PigConfigSessions.Item(SessionName).ConfBoolValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfDateValue(ConfName As String) As DateTime
		Get
			If Me.PigConfigSessions.IsItemExists("Main") = False Then
				Return DateTime.MinValue
			Else
				Return Me.PigConfigSessions.Item("Main").ConfDateValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfDateValue(SessionName As String, ConfName As String) As DateTime
		Get
			If Me.PigConfigSessions.IsItemExists(SessionName) = False Then
				Return DateTime.MinValue
			Else
				Return Me.PigConfigSessions.Item(SessionName).ConfDateValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfDecValue(ConfName As String) As Decimal
		Get
			If Me.PigConfigSessions.IsItemExists("Main") = False Then
				Return 0
			Else
				Return Me.PigConfigSessions.Item("Main").ConfDecValue(ConfName)
			End If
		End Get
	End Property

	Public ReadOnly Property ConfDecValue(SessionName As String, ConfName As String) As Decimal
		Get
			If Me.PigConfigSessions.IsItemExists(SessionName) = False Then
				Return 0
			Else
				Return Me.PigConfigSessions.Item(SessionName).ConfDecValue(ConfName)
			End If
		End Get
	End Property


	''' <summary>
	''' 是否在导入配置前清空原有配置|Clear the original configuration before importing the configuration
	''' 默认为False，好处是配置文件中没有新增的参数，也不会影响新增加参数的使用|The default value is false. The advantage is that there are no new parameters in the configuration file, and the use of new parameters will not be affected
	''' </summary>
	''' <returns></returns>
	Private mbolIsLoadConfClearFrist As Boolean = False
	Public Property IsLoadConfClearFrist() As Boolean
		Get
			Return mbolIsLoadConfClearFrist
		End Get
		Set(value As Boolean)
			mbolIsLoadConfClearFrist = value
		End Set
	End Property



	Public Overloads Function SetDebug(Optional DebugFilePath As String = "") As String
		Dim LOG As New PigStepLog("SetDebug")
		Try
			If DebugFilePath = "" Then DebugFilePath = Me.AppPath & Me.OsPathSep & Me.AppTitle & ".log"
			LOG.StepName = "SetDebug"
			MyBase.SetDebug(DebugFilePath)
			If Me.LastErr <> "" Then
				LOG.AddStepNameInf(DebugFilePath)
				Throw New Exception(Me.LastErr)
			End If
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function

	'Friend Function fSetIsChangeFalse() As String
	'	Try
	'		For Each oPigConfigSession As PigConfigSession In Me.PigConfigSessions
	'			For Each oPigConfig As PigConfig In oPigConfigSession.PigConfigs
	'				If oPigConfig.IsChange = True Then
	'					oPigConfig.IsChange = False
	'				End If
	'			Next
	'			If oPigConfigSession.IsChange = True Then
	'				oPigConfigSession.IsChange = False
	'			End If
	'		Next
	'		If Me.IsChange = True Then
	'			Me.IsChange = False
	'		End If
	'		Return "OK"
	'	Catch ex As Exception
	'		Return Me.GetSubErrInf("fSetIsChangeFalse", ex)
	'	End Try
	'End Function
	Public ReadOnly Property IsChange As Boolean
		Get
			IsChange = False
			For Each oPigConfigSession As PigConfigSession In Me.PigConfigSessions
				If oPigConfigSession.IsChange = True Then
					IsChange = True
					Exit For
				End If
			Next
		End Get
	End Property

	''' <summary>
	''' 保存当前值的PigMD5
	''' </summary>
	''' <returns></returns>
	Friend Function fSaveCurrMD5() As String
		Dim LOG As New PigStepLog("fSaveCurrMD5")
		Try
			Dim strErr As String = ""
			LOG.StepName = "For Each oPigConfigSession"
			For Each oPigConfigSession As PigConfigSession In Me.PigConfigSessions
				LOG.Ret = oPigConfigSession.fSaveCurrMD5()
				If LOG.Ret <> "OK" Then
					strErr &= "fSaveCurrMD5(" & oPigConfigSession.SessionName & ")" & LOG.Ret & Me.OsCrLf
				End If
			Next
			If strErr <> "" Then Throw New Exception(strErr)
			Return "OK"
		Catch ex As Exception
			Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
		End Try
	End Function


End Class
