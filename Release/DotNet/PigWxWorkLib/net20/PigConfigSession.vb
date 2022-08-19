'**********************************
'* Name: 豚豚配置段落|PigConfigSession
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 配置项段落|Configuration session
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8
'* Create Time: 21/12/2021
'* 1.1    22/12/2020   Modify mNew 
'* 1.2    23/12/2020   Modify mNew 
'* 1.3    24/12/2020   Add ConfStrValue,ConfIntValue,ConfBoolValue,ConfDateValue,ConfDecValue, modify mNew,SetDebug
'* 1.4    25/12/2020   Add SessionDesc, modify New,mNew
'* 1.5    2/2/2022     Add IsChange
'* 1.6    21/2/2022    Modify IsChange
'* 1.7    8/5/2022     Modify IsChange,mNew, add LastMD5,fCurrPigMD5
'* 1.8    9/5/2022     Add fSaveCurrMD5
'**********************************
Public Class PigConfigSession
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.8.1"

    Private moPigConfigs As PigConfigs
    Public Property PigConfigs As PigConfigs
        Get
            Return moPigConfigs
        End Get
        Friend Set(value As PigConfigs)
            moPigConfigs = value
        End Set
    End Property

    Public Sub New(SessionName As String, Parent As PigConfigApp)
        MyBase.New(CLS_VERSION)
        Me.mNew(SessionName, "", Parent)
    End Sub

    Public Sub New(SessionName As String, SessionDesc As String, Parent As PigConfigApp)
        MyBase.New(CLS_VERSION)
        Me.mNew(SessionName, SessionDesc, Parent)
    End Sub

    Private Sub mNew(SessionName As String, SessionDesc As String, Parent As PigConfigApp)
        Try
            If SessionName = "" Then Throw New Exception("SessionName is a space")
            If Parent Is Nothing Then Throw New Exception("PigConfigApp Is Nothing")
            With Me
                .SessionName = SessionName
                .SessionDesc = SessionDesc
                .Parent = Parent
                .PigConfigs = New PigConfigs(Me)
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 段落名|Section name
    ''' </summary>
    Private mstrSessionName As String
    Public Property SessionName As String
        Get
            Return mstrSessionName
        End Get
        Friend Set(value As String)
            mstrSessionName = value
        End Set
    End Property

    ''' <summary>
    ''' 段落描述|Section describe
    ''' </summary>
    Private mstrSessionDesc As String
    Public Property SessionDesc As String
        Get
            Return mstrSessionDesc
        End Get
        Friend Set(value As String)
            mstrSessionDesc = value
        End Set
    End Property

    ''' <summary>
    ''' 父对象|Parent object
    ''' </summary>
    Private mpcaParent As PigConfigApp
    Public Property Parent As PigConfigApp
        Get
            Return mpcaParent
        End Get
        Friend Set(value As PigConfigApp)
            mpcaParent = value
        End Set
    End Property

    Public ReadOnly Property ConfStrValue(ConfName As String) As String
        Get
            If Me.PigConfigs.IsItemExists(ConfName) = True Then
                Return Me.PigConfigs.Item(ConfName).ConfValue
            Else
                Return ""
            End If
        End Get
    End Property

    Public ReadOnly Property ConfIntValue(ConfName As String) As Long
        Get
            If Me.PigConfigs.IsItemExists(ConfName) = True Then
                With Me.PigConfigs.Item(ConfName)
                    If .ContentType = PigConfig.EnmContentType.Numerical Then
                        Return CLng(Me.PigConfigs.Item(ConfName).ConfValue)
                    Else
                        Return 0
                    End If
                End With
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property ConfBoolValue(ConfName As String) As Boolean
        Get
            If Me.PigConfigs.IsItemExists(ConfName) = True Then
                With Me.PigConfigs.Item(ConfName)
                    Select Case .ContentType
                        Case PigConfig.EnmContentType.BooleanValue, PigConfig.EnmContentType.Numerical
                            Return CBool(Me.PigConfigs.Item(ConfName).ConfValue)
                        Case Else
                            Return False
                    End Select
                End With
            Else
                Return False
            End If
        End Get
    End Property

    Public ReadOnly Property ConfDateValue(ConfName As String) As DateTime
        Get
            If Me.PigConfigs.IsItemExists(ConfName) = True Then
                With Me.PigConfigs.Item(ConfName)
                    If .ContentType = PigConfig.EnmContentType.DateTime Then
                        Return CDate(Me.PigConfigs.Item(ConfName).ConfValue)
                    Else
                        Return DateTime.MinValue
                    End If
                End With
            Else
                Return DateTime.MinValue
            End If
        End Get
    End Property

    Public ReadOnly Property ConfDecValue(ConfName As String) As Decimal
        Get
            If Me.PigConfigs.IsItemExists(ConfName) = True Then
                With Me.PigConfigs.Item(ConfName)
                    If .ContentType = PigConfig.EnmContentType.Numerical Then
                        Return CDec(Me.PigConfigs.Item(ConfName).ConfValue)
                    Else
                        Return 0
                    End If
                End With
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property IsChange As Boolean
        Get
            If Me.fLastPigMD5 = Me.fCurrPigMD5 Then
                IsChange = False
            Else
                IsChange = True
            End If
            If IsChange = False Then
                For Each oPigConfig As PigConfig In Me.PigConfigs
                    If oPigConfig.IsChange = True Then
                        IsChange = True
                        Exit For
                    End If
                Next
            End If
        End Get
    End Property

    ''' <summary>
    ''' 当前值的PigMD5
    ''' </summary>
    Friend ReadOnly Property fCurrPigMD5 As String
        Get
            Try
                Dim oPigXml As New PigXml(False)
                With Me
                    oPigXml.AddEle("SessionName", .SessionName)
                    oPigXml.AddEle("SessionDesc", .SessionDesc)
                End With
                Dim oPigMD5 As New PigMD5(oPigXml.MainXmlStr, PigMD5.enmTextType.UTF8)
                fCurrPigMD5 = oPigMD5.PigMD5
                oPigXml = Nothing
                oPigMD5 = Nothing
            Catch ex As Exception
                Me.SetSubErrInf("fCurrPigMD5", ex)
                Return ""
            End Try
        End Get
    End Property

    ''' <summary>
    ''' 上个值的PigMD5
    ''' </summary>
    Private mstrLastMD5 As String
    Friend ReadOnly Property fLastPigMD5 As String
        Get
            Return mstrLastMD5
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
            mstrLastMD5 = Me.fCurrPigMD5
            LOG.StepName = "For Each oPigConfig"
            For Each oPigConfig As PigConfig In Me.PigConfigs
                LOG.Ret = oPigConfig.fSaveCurrMD5()
                If LOG.Ret <> "OK" Then
                    strErr &= "fSaveCurrMD5(" & oPigConfig.ConfName & ")" & LOG.Ret & Me.OsCrLf
                End If
            Next
            If strErr <> "" Then Throw New Exception(strErr)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

End Class
