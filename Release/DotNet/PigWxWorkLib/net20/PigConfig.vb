'**********************************
'* Name: 豚豚配置|PigConfig
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 配置项|Configuration item
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.7
'* Create Time: 18/12/2021
'* 1.1    20/12/2021   Add ConfValue,ContentType
'* 1.2    22/12/2021   Modify ConfValue,mNew
'* 1.3    2/2/2022   Add IsChange
'* 1.4    2/5/2022   Modify ConfValue,IsChange
'* 1.5    3/5/2022   Modify ConfValue,IsChange,mNew, add LastMD5,CurrMD5
'* 1.6    8/5/2022   Modify fCurrPigMD5
'* 1.7    9/5/2022   Modify IsChange
'**********************************
Public Class PigConfig
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.7.1"

    Public Enum EnmContentType
        ''' <summary>
        ''' Text
        ''' </summary>
        Text = 0
        ''' <summary>
        ''' Encrypted text
        ''' </summary>
        EncText = 1
        ''' <summary>
        ''' Numerical type
        ''' </summary>
        Numerical = 2
        ''' <summary>
        ''' Date and time type
        ''' </summary>
        DateTime = 3
        ''' <summary>
        ''' Boolean value
        ''' </summary>
        BooleanValue = 4

    End Enum


    Public Sub New(ConfName As String, Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConfName, Parent)
    End Sub

    Public Sub New(ConfName As String, ConfValue As String, Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConfName, Parent, ConfValue)
    End Sub

    Public Sub New(ConfName As String, ConfValue As String, ConfDesc As String, Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.mNew(ConfName, Parent, ConfValue, ConfDesc)
    End Sub

    Private Sub mNew(ConfName As String, Parent As PigConfigSession, Optional ConfValue As String = "", Optional ConfDesc As String = "")
        Try
            If ConfName = "" Then Throw New Exception("ConfName is a space")
            If Parent Is Nothing Then Throw New Exception("PigConfigSession Is Nothing")
            With Me
                .ConfName = ConfName
                .Parent = Parent
                .ConfValue = ConfValue
                .ConfDesc = ConfDesc
                .fSaveCurrMD5()
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mNew", ex)
        End Try
    End Sub

    ''' <summary>
    ''' 父对象|Parent object
    ''' </summary>
    Private moParent As PigConfigSession
    Public Property Parent As PigConfigSession
        Get
            Return moParent
        End Get
        Friend Set(value As PigConfigSession)
            moParent = value
        End Set
    End Property


    ''' <summary>
    ''' 当前值的PigMD5
    ''' </summary>
    Friend ReadOnly Property fCurrPigMD5 As String
        Get
            Try
                Dim oPigXml As New PigXml(False)
                With Me
                    oPigXml.AddEle("ConfName", .ConfName)
                    oPigXml.AddEle("ConfValue", .ConfValue)
                    oPigXml.AddEle("ConfDesc", .ConfDesc)
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
        Try
            mstrLastMD5 = Me.fCurrPigMD5
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("fSaveCurrMD5", ex)
        End Try
    End Function

    ''' <summary>
    ''' 配置名
    ''' </summary>
    Private mstrConfName As String
    Public Property ConfName As String
        Get
            Return mstrConfName
        End Get
        Friend Set(value As String)
            If mstrConfName <> value Then
                mstrConfName = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' 配置描述
    ''' </summary>
    Private mstrConfDesc As String
    Public Property ConfDesc As String
        Get
            Return mstrConfDesc
        End Get
        Set(value As String)
            If mstrConfDesc <> value Then
                mstrConfDesc = value
            End If
        End Set
    End Property

    ''' <summary>
    ''' 配置值
    ''' </summary>
    Private mstrConfValue As String
    Private mstrUnEncConfValue As String = ""
    Public Property ConfValue As String
        Get
            Dim LOG As New PigStepLog("ConfValue")
            Try
                If Me.ContentType = EnmContentType.EncText Then
                    If mstrUnEncConfValue = "" Then
                        LOG.StepName = "fGetUnEncStr"
                        LOG.Ret = Me.Parent.Parent.fGetUnEncStr(mstrUnEncConfValue, mstrConfValue)
                        If LOG.Ret <> "OK" Then
                            mstrUnEncConfValue = ""
                            If Me.Parent IsNot Nothing Then LOG.AddStepNameInf(Me.Parent.SessionName)
                            LOG.AddStepNameInf(Me.ConfName)
                            Throw New Exception(LOG.Ret)
                        End If
                    End If
                    Return mstrUnEncConfValue
                Else
                    Return mstrConfValue
                End If
            Catch ex As Exception
                If Me.Parent IsNot Nothing Then
                    If Me.Parent.Parent IsNot Nothing Then
                        Me.Parent.Parent.PrintDebugLog(Me.MyClassName, Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex))
                    End If
                End If
                Return ""
            End Try
        End Get
        Set(value As String)
            If mstrConfValue <> value Then
                mstrConfValue = value
            End If
            If Mid(mstrConfValue, 1, 5) = "{Enc}" Then
                mstrUnEncConfValue = ""
                Me.ContentType = EnmContentType.EncText
            ElseIf IsDate(mstrConfValue) = True Then
                Me.ContentType = EnmContentType.DateTime
            ElseIf IsNumeric(mstrConfValue) Then
                Me.ContentType = EnmContentType.Numerical
            Else
                Select Case UCase(mstrConfValue)
                    Case "TRUE", "FALSE"
                        Me.ContentType = EnmContentType.BooleanValue
                    Case Else
                        Me.ContentType = EnmContentType.Text
                End Select
            End If
        End Set
    End Property

    Friend ReadOnly Property fConfValue As String
        Get
            Return mstrConfValue
        End Get
    End Property


    Private mintContentType As EnmContentType
    Public Property ContentType As EnmContentType
        Get
            Return mintContentType
        End Get
        Friend Set(value As EnmContentType)
            If mintContentType <> value Then
                mintContentType = value
            End If
        End Set
    End Property

    Public ReadOnly Property IsChange As Boolean
        Get
            If Me.fLastPigMD5 = Me.fCurrPigMD5 Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property


End Class

