'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2020-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Basic lightweight Edition
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.6.1
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
'*1.0.26 23/7/2021   Modify mPrintDebugLog,mGetSubErrInf,FullSubName add AppVersion
'*1.0.27 25/7/2021   Add mOpenDebug,OpenDebug, remove GetSubStepDebugInf, Modify mstrAppTitle
'*1.1.1 31/8/2021    Modify MyClassName
'*1.2 7/12/2021      Add MyPID, modify mGetSubErrInf,MyClassName,mGetSubErrInf, remove mstrKeyInf
'*1.3 8/12/2021      Add StruSubLog
'*1.5 17/12/2021     Modify StruSubLog
'*1.6 10/2/2022     Modify fMyPID
'************************************
Imports System.Runtime.InteropServices
Public Class PigBaseMini
    Private mstrClsName As String
    Private ReadOnly mstrClsVersion As String
    Private mstrLastErr As String = ""
    Private mbolIsDebug As Boolean
    Private mbolIsHardDebug As Boolean
    Private mstrDebugFilePath As String

    Friend Structure StruSubLog
        Dim SubName As String
        Dim StepName As String
        Dim Ret As String
    End Structure

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

    Friend Sub ClearErr()
        mstrLastErr = ""
    End Sub

    Private mlngMyPID As Long = 0
    Friend ReadOnly Property fMyPID() As Long
        Get
            Try
                If mlngMyPID <= 0 Then
#If NET5_0_OR_GREATER Then
                    mlngMyPID = System.Environment.ProcessId
#Else
                    mlngMyPID = System.Diagnostics.Process.GetCurrentProcess.Id
#End If
                End If
                Return mlngMyPID
            Catch ex As Exception
                Me.SetSubErrInf("fMyPID", ex)
                Return 0
            End Try
        End Get
    End Property



    Friend Overloads Sub SetDebug(DebugFilePath As String)
        mbolIsDebug = True
        mbolIsHardDebug = False
        mstrDebugFilePath = DebugFilePath
    End Sub

    Friend Overloads Sub SetDebug(DebugFilePath As String, IsHardDebug As Boolean)
        mbolIsDebug = True
        mbolIsHardDebug = IsHardDebug
        mstrDebugFilePath = DebugFilePath
    End Sub

    Friend Overloads Function PrintDebugLog(SubName As String, LogInf As String, IsHardDebug As Boolean) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, "", LogInf, IsHardDebug)
    End Function

    Friend Overloads Function PrintDebugLog(SubName As String, StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, StepName, LogInf, IsHardDebug)
    End Function

    Friend Overloads Function PrintDebugLog(SubName As String, StepName As String, LogInf As String) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, StepName, LogInf, False)
    End Function

    Friend Overloads Function PrintDebugLog(SubName As String, LogInf As String) As String
        PrintDebugLog = Me.mPrintDebugLog(SubName, "", LogInf, False)
    End Function

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
            sbAny.Append("[" & Me.AppTitle & "(" & Me.AppVersion & ")][" & Me.MyClassName & "(" & Me.mstrClsVersion & ")." & SubName)
            If StepName <> "" Then
                sbAny.Append("{" & StepName & "}")
            End If
            sbAny.Append("]" & LogInf)
            strLogInf = "[" & dtNow.ToString("yyyy-MM-dd HH:mm:ss.fff") & "][" & Me.fMyPID.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString & "]" & sbAny.ToString
            swAny.WriteLine(strLogInf)
            swAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    Public Overloads Property MyClassName() As String
        Get
            MyClassName = mstrClsName
        End Get
        Friend Set(value As String)
            mstrClsName = value
        End Set
    End Property

    Public Overloads ReadOnly Property MyClassName(IsIncAppTitle As Boolean) As String
        Get
            If IsIncAppTitle = True Then
                MyClassName = System.Reflection.Assembly.GetExecutingAssembly().GetName.Name & "." & mstrClsName
            Else
                MyClassName = mstrClsName
            End If
        End Get
    End Property

    Friend ReadOnly Property IsHardDebug() As Boolean
        Get
            IsHardDebug = mbolIsHardDebug
        End Get
    End Property

    Friend ReadOnly Property IsDebug() As Boolean
        Get
            IsDebug = mbolIsDebug
        End Get
    End Property

    Private mstrAppTitle As String = ""
    Friend ReadOnly Property AppTitle() As String
        Get
            If mstrAppTitle = "" Then mstrAppTitle = System.Reflection.Assembly.GetExecutingAssembly().GetName.Name
            Return mstrAppTitle
        End Get
    End Property

    Private mstrAppVersion As String = ""
    Friend ReadOnly Property AppVersion() As String
        Get
            If mstrAppVersion = "" Then mstrAppVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString
            Return mstrAppVersion
        End Get
    End Property

    Private mstrAppPath As String = ""
    Friend ReadOnly Property AppPath() As String
        Get
            If mstrAppPath = "" Then
                mstrAppPath = System.AppDomain.CurrentDomain.BaseDirectory
                If Right(mstrAppPath, 1) = Me.OsPathSep Then
                    mstrAppPath = Left(mstrAppPath, mstrAppPath.Length - 1)
                End If
            End If
            Return mstrAppPath
        End Get
    End Property


    Public ReadOnly Property LastErr() As String
        Get
            LastErr = mstrLastErr
        End Get
    End Property

    Friend ReadOnly Property FullSubName(SubName As String) As String
        Get
            FullSubName = mstrClsName & "(" & Me.mstrClsVersion & ")." & SubName
        End Get
    End Property

    Private Function mGetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False, Optional IsSetLastErr As Boolean = False) As String
        Try
            Dim sbAny As New System.Text.StringBuilder("[Error]")
            sbAny.Append("[" & Me.AppTitle & "][" & Me.AppVersion & "][")
            sbAny.Append(Me.FullSubName(SubName))
            If StepName.Length > 0 Then
                sbAny.Append("{" & StepName & "}")
            End If
            sbAny.Append("][ErrInf:")
            sbAny.Append(exIn.Message & "]")
            If IsStackTrace = True Then
                Dim strExStackTrace As String = exIn.StackTrace.Trim
                With strExStackTrace
                    If .Length > 0 Then
                        If .LastIndexOf(vbCrLf) >= 0 Then strExStackTrace = .Replace(vbCrLf, "")
                        If .LastIndexOf(vbTab) >= 0 Then strExStackTrace = .Replace(vbTab, " ")
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

    Friend Overloads Sub SetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        Me.mGetSubErrInf(SubName, "", exIn, IsStackTrace, True)
    End Sub

    Friend Overloads Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        Me.mGetSubErrInf(SubName, StepName, exIn, IsStackTrace, True)
    End Sub


    Friend Overloads Function GetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        GetSubErrInf = Me.mGetSubErrInf(SubName, "", exIn, IsStackTrace)
    End Function

    ''' <summary>Get process error information with step name</summary>
    Friend Overloads Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        GetSubErrInf = Me.mGetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Function


    Private ReadOnly mbolIsWindows As Boolean
    Friend ReadOnly Property IsWindows() As Boolean
        Get
            Return mbolIsWindows
        End Get
    End Property

    ''' <summary>
    ''' Cross platform line feed, automatically identifying windows and Linux
    ''' </summary>
    Private ReadOnly mstrOsCrLf As String
    Friend ReadOnly Property OsCrLf() As String
        Get
            Return mstrOsCrLf
        End Get
    End Property

    ''' <summary>
    ''' Cross platform path separator, automatically identifying Windows and Linux
    ''' </summary>
    Private ReadOnly mstrOsPathSep As String
    Friend ReadOnly Property OsPathSep() As String
        Get
            Return mstrOsPathSep
        End Get
    End Property

    Public Sub OpenDebug()
        Me.mOpenDebug()
    End Sub

    Public Sub OpenDebug(LogFileDir As String)
        Me.mOpenDebug(LogFileDir)
    End Sub

    Public Sub OpenDebug(LogFileDir As String, IsIncDate As Boolean)
        Me.mOpenDebug(LogFileDir, IsIncDate)
    End Sub

    Public Sub OpenDebug(IsIncDate As Boolean)
        Me.mOpenDebug(, IsIncDate)
    End Sub

    Private Sub mOpenDebug(Optional LogFileDir As String = "", Optional IsIncDate As Boolean = False)
        Try
            Dim strLogFilePath As String = Me.AppTitle
            If LogFileDir = "" Then LogFileDir = Me.AppPath
            Dim dteNow As DateTime = DateTime.Now
            If IsIncDate = True Then strLogFilePath &= dteNow.ToString("yyyyMMdd")
            strLogFilePath = LogFileDir & Me.OsPathSep & strLogFilePath & ".log"
            Me.SetDebug(strLogFilePath)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mOpenDebug", ex)
        End Try
    End Sub

End Class
