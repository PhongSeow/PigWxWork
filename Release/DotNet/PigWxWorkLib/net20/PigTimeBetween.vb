'**********************************
'* Name: 豚豚时间间隔|PigTimeBetween
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 判断一个日期是否落在时间间隔之中|Judge whether a date falls within the time interval
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0
'* Create Time: 20/3/2022
'**********************************

Friend Class PigTimeBetween
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.1"

    ''' <summary>间隔类型|Interval type</summary>
    Public Enum EnmBetweenType
        CurrentYear = 0
        CurrentQuarter = 1
        CurrentMonth = 2
        ''' <summary>
        ''' 当旬
        ''' </summary>
        CurrentTendays = 3
        CurrentWeek = 4
        PrevYear = 10
        PrevQuarter = 11
        PrevMonth = 12
        ''' <summary>
        ''' 上旬
        ''' </summary>
        PrevTendays = 13
        PrevWeek = 14
        NextYear = 20
        NextQuarter = 21
        NextMonth = 22
        ''' <summary>
        ''' 下旬
        ''' </summary>
        NextTendays = 23
        NextWeek = 24
    End Enum


    Private mintBetweenType As EnmBetweenType
    Public Property BetweenType As EnmBetweenType
        Get
            Return mintBetweenType
        End Get
        Friend Set(value As EnmBetweenType)
            Dim strNow As String = Format(Now, "yyyy-MM-dd HH:mm:ss.fff")
            Dim strBeginTime As String, strEndTime As String
            Select Case value
                Case EnmBetweenType.CurrentYear
                    strBeginTime = Left(strNow, 4) & "-01-01"
                    strEndTime = Format(Now.AddYears(1), "yyyy-MM-dd HH:mm:ss.fff")
                    strEndTime = Left(strEndTime, 4) & "-01-01"
                Case EnmBetweenType.CurrentQuarter
                Case EnmBetweenType.CurrentMonth
                Case EnmBetweenType.CurrentTendays
                Case EnmBetweenType.CurrentWeek
                Case EnmBetweenType.PrevYear
                    strBeginTime = Format(Now.AddYears(-1), "yyyy-MM-dd HH:mm:ss.fff")
                    strBeginTime = Left(strBeginTime, 4) & "-01-01"
                    strEndTime = Left(strNow, 4) & "-01-01"
                Case EnmBetweenType.PrevQuarter
                Case EnmBetweenType.PrevMonth
                Case EnmBetweenType.PrevTendays
                Case EnmBetweenType.PrevWeek
                Case EnmBetweenType.NextYear
                    strBeginTime = Format(Now.AddYears(1), "yyyy-MM-dd HH:mm:ss.fff")
                    strBeginTime = Left(strBeginTime, 4) & "-01-01"
                    strEndTime = Format(Now.AddYears(2), "yyyy-MM-dd HH:mm:ss.fff")
                    strEndTime = Left(strEndTime, 4) & "-01-01"
                Case EnmBetweenType.NextQuarter
                Case EnmBetweenType.NextMonth
                Case EnmBetweenType.NextTendays
                Case EnmBetweenType.NextWeek
                Case Else
                    mintBetweenType = EnmBetweenType.CurrentMonth
            End Select
        End Set
    End Property




    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Private mdteBeginTime As DateTime
    Public Property BeginTime As DateTime
        Get
            Return mdteBeginTime
        End Get
        Friend Set(value As DateTime)
            mdteBeginTime = value
        End Set
    End Property

    Private mdteEndTime As DateTime
    Public Property EndTime As DateTime
        Get
            Return mdteEndTime
        End Get
        Friend Set(value As DateTime)
            mdteEndTime = value
        End Set
    End Property

    ''' <summary>
    ''' 判断时间是否在间隔中|Determine whether the time is in the interval
    ''' </summary>
    ''' <param name="InTime">输入时间|Input time</param>
    ''' <returns></returns>
    Public ReadOnly Property IsTimeBetween(InTime As DateTime) As Boolean
        Get
            If InTime >= Me.BeginTime And InTime < Me.EndTime Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property


End Class
