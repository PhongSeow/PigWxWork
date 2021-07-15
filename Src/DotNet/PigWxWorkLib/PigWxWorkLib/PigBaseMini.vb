'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Basic lightweight Edition
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.25
'* Create Time: 31/8/2019
'*1.0.2  1/10/2019   Add mGetSubErrInf 
'*1.0.3  4/11/2019   Add LastErr
'*1.0.4  5/11/2019-11-5   Add SetSubErrInf
'*1.0.5  6/2/2020    Add GetSubStepDebugInf
'*1.0.6  3/3/2020    Add Debug function, in debug mode, you can also print log
'*1.0.7  3/4/2020    Add Hard debug function, modify GetSubStepDebugInf
'*1.0.8  3/6/2020    Add KeyInf
'*1.0.9  6/17/2020   modify mPrintDebugLog Add StepName
'*1.0.10 6/25/2020   Not used My.Application , better compatibility, Add MyClassName
'*1.0.11 30/11/2020  Update some summary, modify AppTitle, add AppPath
'*1.0.12 6/12/2020   Modify mPrintDebugLog
'*1.0.13 8/12/2020   Modify ClearErr
'*1.0.14 27/12/2020  Add IsWindows,
'*1.0.15 4/1/2021    Modify New
'*1.0.16 15/1/2021   Modify New
'*1.0.17 15/1/2021   Err.Raise change to Throw New Exception
'*1.0.18 26/1/2021   Change some sub or function Public to Friend, modify 
'*1.0.19 27/1/2021   Change KeyInf,ClearErr Public to Friend, modify 
'*1.0.20 20/2/2021   Fix bug mstrKeyInf is nothing
'*1.0.21 25/2/2021   Fix bug mstrLastErr is nothing
'*1.0.22 6/7/2021    Modify New 
'*1.0.23 9/7/2021    Modify New for fix bugs that identify Windows and Linux operating system types.
'*1.0.24 14/7/2021   Modify mPrintDebugLog
'*1.0.25 15/7/2021   Modify mPrintDebugLog,PrintDebugLog,mGetSubErrInf
'************************************
Imports System.Runtime.InteropServices
Public Class PigBaseMini
    ''' <summary>
    ''' 类名
    ''' </summary>
    Private mstrClsName As String

    ''' <summary>
    ''' 类版本
    ''' </summary>
    Private mstrClsVersion As String
    Private mstrLastErr As String = ""
    Private mbolIsDebug As Boolean
    Private mbolIsHardDebug As Boolean
    Private mstrDebugFilePath As String
    Private mstrKeyInf As String = ""


    Public Sub New(Version As String)
        mstrClsName = Me.GetType.Name.ToString()
        mstrClsVersion = Version

#If NETCOREAPP Or NET5_0_OR_GREATER Then
        mbolIsWindows = RuntimeInformation.IsOSPlatform(OSPlatform.Windows)
#Else
        mbolIsWindows = True
#End If
        If mbolIsWindows = True Then
            mstrOsCrLf = vbCrLf
            mstrOsPathSep = "\"
        Else
            mstrOsCrLf = vbLf
            mstrOsPathSep = "/"
        End If
    End Sub

    ''' <remarks>返回成功</remarks>
    Friend Sub ClearErr()
        mstrLastErr = ""
    End Sub

    ''' <summary>设置调试</summary>
    Friend Overloads Sub SetDebug(DebugFilePath As String)
        mbolIsDebug = True
        mbolIsHardDebug = False
        mstrDebugFilePath = DebugFilePath
    End Sub

    ''' <summary>设置调试</summary>
    Friend Overloads Sub SetDebug(DebugFilePath As String, IsHardDebug As Boolean)
        mbolIsDebug = True
        mbolIsHardDebug = IsHardDebug
        mstrDebugFilePath = DebugFilePath
    End Sub

    ''' <summary>调试</summary>
    Friend Overloads Function PrintDebugLog(SubName As String, LogInf As String, IsHardDebug As Boolean) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, "", LogInf, IsHardDebug)
    End Function

    ''' <summary>调试</summary>
    Friend Overloads Function PrintDebugLog(SubName As String, StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, StepName, LogInf, IsHardDebug)
    End Function

    ''' <summary>调试</summary>
    Friend Overloads Function PrintDebugLog(SubName As String, StepName As String, LogInf As String) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, StepName, LogInf, False)
    End Function

    ''' <summary>调试</summary>
    Friend Overloads Function PrintDebugLog(SubName As String, LogInf As String) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, "", LogInf, False)
    End Function

    ''' <summary>调试</summary>
    Private Function mPrintDebugLog(SubName As String, StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        Try
            If IsHardDebug = True And mbolIsHardDebug = False Then Throw New Exception("Only hard debug mode can print logs")
            If mbolIsDebug = False Then Throw New Exception("Only debug mode can print logs")
            Dim sfAny As New System.IO.FileStream(Me.mstrDebugFilePath, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.Write, 10240, False)
            Dim swAny = New System.IO.StreamWriter(sfAny)
            Dim dtNow As System.DateTime = System.DateTime.Now
            Dim sbAny As New System.Text.StringBuilder("")
            Dim strLogInf As String
            If IsHardDebug = True Then
                sbAny.Append("[HardDebug]")
            Else
                sbAny.Append("[Debug]")
            End If
            sbAny.Append("[" & Me.AppTitle & "][" & Me.MyClassName & "." & SubName)
            If StepName <> "" Then
                sbAny.Append("." & StepName)
            End If
            sbAny.Append("]")
            sbAny.Append(LogInf)
            strLogInf = "[" & dtNow.ToString("yyyy-MM-dd HH:mm:ss.fff") & "][" & System.Diagnostics.Process.GetCurrentProcess.Id.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString & "]" & sbAny.ToString
            swAny.WriteLine(strLogInf)
            swAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>我的类名</summary>
    Public Overloads ReadOnly Property MyClassName() As String
        Get
            MyClassName = mstrClsName
        End Get
    End Property

    ''' <summary>我的类名</summary>
    Public Overloads ReadOnly Property MyClassName(IsIncAppTitle As Boolean) As String
        Get
            If IsIncAppTitle = True Then
                MyClassName = System.Reflection.Assembly.GetExecutingAssembly().GetName.Name & "."
            End If
            MyClassName = mstrClsName
        End Get
    End Property

    ''' <summary>是否硬调试</summary>
    Friend ReadOnly Property IsHardDebug() As Boolean
        Get
            IsHardDebug = mbolIsHardDebug
        End Get
    End Property

    ''' <summary>是否调试</summary>
    Friend ReadOnly Property IsDebug() As Boolean
        Get
            IsDebug = mbolIsDebug
        End Get
    End Property

    ''' <summary>应用标题</summary>
    Friend ReadOnly Property AppTitle() As String
        Get
            Return System.Reflection.Assembly.GetExecutingAssembly().GetName.Name
        End Get
    End Property

    ''' <summary>应用路径</summary>
    Friend ReadOnly Property AppPath() As String
        Get
            Return System.AppDomain.CurrentDomain.BaseDirectory
        End Get
    End Property


    ''' <remarks>最近的错误信息</remarks>
    Public ReadOnly Property LastErr() As String
        Get
            LastErr = mstrLastErr
        End Get
    End Property

    ''' <remarks>完整过程名</remarks>
    Friend ReadOnly Property FullSubName(SubName As String) As String
        Get
            FullSubName = mstrClsName & "." & SubName
        End Get
    End Property

    Private Function mGetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False, Optional IsSetLastErr As Boolean = False) As String
        Try
            Dim sbAny As New System.Text.StringBuilder("[Error]")
            sbAny.Append("[" & Me.AppTitle & "][")
            sbAny.Append(Me.FullSubName(SubName))
            If StepName.Length > 0 Then
                sbAny.Append("." & StepName)
            End If
            sbAny.Append("]")
            If mstrKeyInf.Length > 0 Then sbAny.Append("[Key:" & mstrKeyInf & "]")
            sbAny.Append("[ErrInf:")
            sbAny.Append(exIn.Message & "]")
            If IsStackTrace = True Then
                Dim strExStackTrace As String = exIn.StackTrace
                With strExStackTrace
                    If .Length > 0 Then
                        If .LastIndexOf(vbCrLf) >= 0 Then .Replace(vbCrLf, "")
                        If .LastIndexOf(vbTab) >= 0 Then .Replace(vbTab, " ")
                        .Trim()
                    End If
                End With
                sbAny.Append("[Trace:")
                sbAny.Append(strExStackTrace & "]")
            End If
            If IsSetLastErr = True Then mstrLastErr = sbAny.ToString
            mGetSubErrInf = sbAny.ToString
            sbAny = Nothing
        Catch ex As Exception
            If IsSetLastErr = True Then mstrLastErr = ex.Message.ToString
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>设置过程错误信息,用于不返回结果的调用，结果会保存在LastErr，但如果执行成功，要调用ClearErr</summary>
    ''' <param name="SubName">过程名</param>
    ''' <param name="exIn">错误对象</param>
    ''' <param name="IsStackTrace">是否跟踪</param>
    Friend Overloads Sub SetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        Me.mGetSubErrInf(SubName, "", exIn, IsStackTrace, True)
    End Sub

    ''' <summary>设置过程错误信息,带步骤名称,用于不返回结果的调用，结果会保存在LastErr，但如果执行成功，要调用ClearErr</summary>
    ''' <param name="SubName">过程名</param>
    ''' <param name="StepName">步骤名</param>
    ''' <param name="exIn">错误对象</param>
    ''' <param name="IsStackTrace">是否跟踪</param>
    Friend Overloads Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        Me.mGetSubErrInf(SubName, StepName, exIn, IsStackTrace, True)
    End Sub


    ''' <summary>获取过程错误信息</summary>
    ''' <param name="SubName">过程名</param>
    ''' <param name="exIn">错误对象</param>
    ''' <param name="IsStackTrace">是否跟踪</param>
    Friend Overloads Function GetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        GetSubErrInf = Me.mGetSubErrInf(SubName, "", exIn, IsStackTrace)
    End Function

    ''' <remarks>获取过程错误信息，带步骤名称</remarks>
    Friend Overloads Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        GetSubErrInf = Me.mGetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Function

    ''' <remarks>获取未出错时过程调试信息</remarks>
    Friend Function GetSubStepDebugInf(SubName As String, StepName As String, DebugInf As String) As String
        Try
            Dim sbAny As New System.Text.StringBuilder("")
            sbAny.Append(Me.FullSubName(SubName))
            If StepName.Length > 0 Then
                sbAny.Append("(")
                sbAny.Append(StepName)
                sbAny.Append(")")
            End If
            If mstrKeyInf.Length > 0 Then sbAny.Append(";Key:" & mstrKeyInf)
            sbAny.Append(";Debug:")
            sbAny.Append(DebugInf)
            Return sbAny.ToString
            sbAny = Nothing
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 是否 Windows
    ''' </summary>
    Private mbolIsWindows As Boolean
    Friend ReadOnly Property IsWindows() As Boolean
        Get
            Return mbolIsWindows
        End Get
    End Property

    ''' <summary>
    ''' 跨平台换行符，自动识别 Windows 和 Linux
    ''' </summary>
    Private mstrOsCrLf As String
    Friend ReadOnly Property OsCrLf() As String
        Get
            Return mstrOsCrLf
        End Get
    End Property

    ''' <summary>
    ''' 跨平台路径分隔符，自动识别 Windows 和 Linux
    ''' </summary>
    Private mstrOsPathSep As String
    Friend ReadOnly Property OsPathSep() As String
        Get
            Return mstrOsPathSep
        End Get
    End Property

End Class
