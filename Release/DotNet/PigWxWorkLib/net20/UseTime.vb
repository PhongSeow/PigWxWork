'**********************************
'* Name: ShareMem
'* Author: Part of the source code found on the Internet, I do not know who is the author, organized by Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Use time processing 使用时间处理
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1.2
'* Create Time: 10/8/2019
'* 1.0.2  2019-10-12  GoBegin,Step2NowSeconds,Step2NowSeconds优化,新增 AllDiffChnMain
'* 1.0.3  2019-10-12  增加 FreeSeconds 
'* 1.1.2  8/12/2021  Add BeginTime,EndTime
'**********************************
Public Class UseTime
    Private mdteBegin As DateTime
    Private mdteCurrBegin As DateTime
    Private mtsTimeDiff As TimeSpan

    ''' <summary>空闲的秒数，返回0表示未空闲或未初始化</summary>
    Public Function FreeSeconds() As Decimal
        If mtsTimeDiff.TotalMilliseconds = 0 Then
            FreeSeconds = 0
        Else
            Dim tsAny As TimeSpan = Now - mdteCurrBegin
            FreeSeconds = (tsAny.TotalMilliseconds - mtsTimeDiff.TotalMilliseconds) / 1000
        End If
    End Function


    ''' <summary>总时间开始计时</summary>
    Public Sub GoBegin()
        mdteBegin = Now
        mdteCurrBegin = mdteBegin
        mtsTimeDiff = Nothing
        mtsTimeDiff = New TimeSpan
    End Sub

    ''' <summary>当前步骤开始计时</summary>
    Public Sub GoNext()
        mdteCurrBegin = Now
    End Sub


    ''' <summary>结束计时</summary>
    Public Sub ToEnd()
        Me.EndTime = Now
        mtsTimeDiff = Now - mdteBegin
    End Sub

    ''' <summary>本步到现在相差秒数</summary>
    Public Overloads Function Step2NowSeconds() As Decimal
        If mtsTimeDiff.TotalMilliseconds = 0 Then
            Dim tsAny As TimeSpan = Now - mdteCurrBegin
            Step2NowSeconds = tsAny.TotalMilliseconds / 1000
        Else
            Step2NowSeconds = 0
        End If
    End Function

    ''' <summary>本步到现在相差秒数</summary>
    Public Overloads Function Step2NowSeconds(IsGoNextNow As Boolean) As Decimal
        If mtsTimeDiff.TotalMilliseconds = 0 Then
            Dim tsAny As TimeSpan = Now - mdteCurrBegin
            Step2NowSeconds = tsAny.TotalMilliseconds / 1000
            If IsGoNextNow = True Then Me.GoNext()
        Else
            Step2NowSeconds = 0
        End If
    End Function


    ''' <summary>全部过程相差秒数</summary>
    Public Function AllDiffSeconds() As Decimal
        AllDiffSeconds = mtsTimeDiff.TotalMilliseconds / 1000
    End Function

    ''' <summary>全部过程相差分钟数</summary>
    Public Function AllDiffMinutes() As Decimal
        AllDiffMinutes = mtsTimeDiff.TotalMilliseconds / 1000 / 60
    End Function

    Public Function AllDiffChn() As String
        AllDiffChn = Me.AllDiffChnMain(mtsTimeDiff)
    End Function

    Public Function StepDiffChn() As String
        If mtsTimeDiff.TotalMilliseconds = 0 Then
            StepDiffChn = ""
        Else
            Dim tsAny As TimeSpan = Now - mdteCurrBegin
            StepDiffChn = Me.AllDiffChnMain(tsAny)
        End If

    End Function

    Private Function AllDiffChnMain(tsTimeDiff As TimeSpan) As String
        '全部过程相差中文时间
        Dim decSs As Long, lngMm As Long, lngHh As Long, cuyDd As Decimal, lngYy As Long
        Try
            decSs = tsTimeDiff.TotalSeconds
            '取出天

            cuyDd = Int(decSs / 86400)
            decSs = decSs - cuyDd * 86400
            '取出小时
            lngHh = Int(decSs / 3600)
            decSs = decSs - lngHh * 3600
            '取出分钟
            lngMm = Int(decSs / 60)
            decSs = decSs - lngMm * 60
            AllDiffChnMain = ""
            If cuyDd > 365 Then
                lngYy = Int(cuyDd / 365)
                cuyDd = cuyDd - lngYy * 365
            End If
            If lngYy > 0 Then AllDiffChnMain &= lngYy.ToString & "年"
            If cuyDd > 0 Then AllDiffChnMain &= cuyDd.ToString & "天"
            If lngHh > 0 Then AllDiffChnMain &= lngHh.ToString & "小时"
            If lngMm > 0 Then AllDiffChnMain &= lngMm.ToString & "分钟"
            If decSs > 0 Then AllDiffChnMain &= decSs.ToString & "秒"
        Catch ex As Exception
            AllDiffChnMain = ""
        End Try
    End Function

    Public ReadOnly Property BeginTime As DateTime
        Get
            Return mdteBegin
        End Get
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

End Class
