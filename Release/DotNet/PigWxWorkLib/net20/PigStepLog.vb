'**********************************
'* Name: PigStepLog
'* Author: Seow Phong
'* License: Copyright (c) 2020-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigStepLog is for logging and error handling in the process.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.3.2
'* Create Time: 8/12/2019
'1.1    18/12/2021  Add TrackID,ErrInf2User, modify mNew,StepName
'1.2    21/12/2021  Modify TrackID
'1.3    10/2/2022  Add StepLogInf
'************************************
Public Class PigStepLog
    Public ReadOnly Property SubName As String

    Private mbolIsLogUseTime As Boolean
    Public Property IsLogUseTime As Boolean
        Get
            Return mbolIsLogUseTime
        End Get
        Friend Set(value As Boolean)
            mbolIsLogUseTime = value
        End Set
    End Property

    Public Ret As String = ""

    Private moUseTime As UseTime

    ''' <summary>
    ''' 显示给用户看的错误信息，不能显示内部错误内容。
    ''' </summary>
    Private mstrErrInf2User As String
    Public Property ErrInf2User As String
        Get
            If Me.TrackID = "" Then
                Return mstrErrInf2User
            Else
                Return mstrErrInf2User & "(" & Me.TrackID & ")"
            End If
        End Get
        Set(value As String)
            mstrErrInf2User = value
        End Set
    End Property

    Private Sub mNew(Optional IsTrack As Boolean = False, Optional IsLogUseTime As Boolean = False)
        Me.IsLogUseTime = IsLogUseTime
        If IsTrack = True Then
            Dim oPigFunc As New PigFunc
            Me.TrackID = oPigFunc.GetPKeyValue("PigStepLog", False)
            oPigFunc = Nothing
        End If
        If Me.IsLogUseTime = True Then
            moUseTime = New UseTime
            moUseTime.GoBegin()
        End If
    End Sub
    Public Sub New(SubName As String)
        Me.SubName = SubName
        Me.mNew()
    End Sub

    Public Sub New(SubName As String, IsLogUseTime As Boolean)
        Me.SubName = SubName
        Me.mNew(, IsLogUseTime)
    End Sub


    Private mstrTrackID As String = ""
    Public Property TrackID As String
        Get
            Return mstrTrackID
        End Get
        Friend Set(value As String)
            mstrTrackID = value
        End Set
    End Property

    Private mstrStepName As String = ""
    Public Property StepName As String
        Get
            Return mstrStepName
        End Get
        Set(value As String)
            mstrStepName = value
            If Me.TrackID <> "" Then
                Me.AddStepNameInf(Me.TrackID)
            End If
        End Set
    End Property

    Public Sub AddStepNameInf(AddInf As String)
        Me.StepName &= "(" & AddInf & ")"
    End Sub

    Public ReadOnly Property BeginTime As DateTime
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.BeginTime
            Else
                Return DateTime.MinValue
            End If
        End Get
    End Property

    Public ReadOnly Property EndTime As DateTime
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.EndTime
            Else
                Return DateTime.MaxValue
            End If
        End Get
    End Property

    Public ReadOnly Property AllDiffSeconds As Decimal
        Get
            If Me.IsLogUseTime Then
                Return moUseTime.AllDiffSeconds
            Else
                Return -1
            End If
        End Get
    End Property


    Public Sub ToEnd()
        If Me.IsLogUseTime = True Then
            moUseTime.ToEnd()
        End If
    End Sub

    ''' <summary>
    ''' 当前步骤的日志信息|Log information of the current step
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property StepLogInf As String
        Get
            With Me
                StepLogInf = "[" & Me.SubName & "]"
                StepLogInf &= "[" & Me.StepName & "]"
                If Me.Ret <> "" Then StepLogInf &= "[" & Me.Ret & "]"
                If Me.TrackID <> "" Then StepLogInf &= "[TrackID:" & Me.TrackID & "]"
                If Me.IsLogUseTime = True Then StepLogInf &= "[" & Me.AllDiffSeconds.ToString & "]"
            End With
        End Get
    End Property

End Class
