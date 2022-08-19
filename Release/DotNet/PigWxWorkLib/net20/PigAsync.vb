'**********************************
'* Name: PigAsync
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigAsync is for Asynchronous processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1.3
'* Create Time: 18/5/2022
'*1.1  5/6/2022 Add AsyncCmdPID
'**********************************
Imports System.Threading
Public Class PigAsync

    ''' <summary>
    ''' 开始异步处理|Start asynchronous processing
    ''' </summary>
    ''' <returns></returns>
    Public Function AsyncBegin() As String
        Try
            With Me
                .AsyncBeginTime = Now
                .AsyncEndTime = DateTime.MinValue
                .AsyncThreadID = Thread.CurrentThread.ManagedThreadId
                .AsyncRet = ""
                .AsyncCmdPID = -1
            End With
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 异步处理成功|Asynchronous processing succeeded
    ''' </summary>
    ''' <returns></returns>
    Public Function AsyncSucc() As String
        Try
            With Me
                .AsyncEndTime = Now
                .AsyncRet = "OK"
            End With
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 异步处理成功|Asynchronous processing succeeded
    ''' </summary>
    ''' <returns></returns>
    Public Function AsyncError(ErrString As String) As String
        Try
            With Me
                .AsyncEndTime = Now
                .AsyncRet = ErrString
            End With
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 异步处理返回结果|Asynchronous processing returns results
    ''' </summary>
    ''' <returns></returns>
    Private mstrAsyncRet As String
    Public Property AsyncRet As String
        Get
            Return mstrAsyncRet
        End Get
        Friend Set(value As String)
            mstrAsyncRet = value
        End Set
    End Property

    ''' <summary>
    ''' 异步处理返回cmd.exe的进程号|Asynchronous processing returns cmd.exe process number
    ''' </summary>
    ''' <returns></returns>
    Private mintAsyncCmdPID As Integer
    Public Property AsyncCmdPID As Integer
        Get
            Return mintAsyncCmdPID
        End Get
        Set(value As Integer)
            mintAsyncCmdPID = value
        End Set
    End Property

    ''' <summary>
    ''' 异步处理开始时间|Asynchronous processing start time
    ''' </summary>
    ''' <returns></returns>
    Private mdteAsyncBeginTime As DateTime
    Public Property AsyncBeginTime As DateTime
        Get
            Return mdteAsyncBeginTime
        End Get
        Friend Set(value As DateTime)
            mdteAsyncBeginTime = value
        End Set
    End Property

    ''' <summary>
    ''' 异步处理结束时间|Asynchronous processing end time
    ''' </summary>
    ''' <returns></returns>
    Private mdteAsyncEndTime As DateTime
    Public Property AsyncEndTime As DateTime
        Get
            Return mdteAsyncEndTime
        End Get
        Friend Set(value As DateTime)
            mdteAsyncEndTime = value
        End Set
    End Property

    ''' <summary>
    ''' 异步处理线程号|Asynchronous processing thread number
    ''' </summary>
    ''' <returns></returns>
    Private mintAsyncThreadID As Integer
    Public Property AsyncThreadID As Integer
        Get
            Return mintAsyncThreadID
        End Get
        Friend Set(value As Integer)
            mintAsyncThreadID = value
        End Set
    End Property

End Class
