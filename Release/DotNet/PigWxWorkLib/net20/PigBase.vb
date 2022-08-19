'**********************************
'* Name: PigBaseMini
'* Author: Seow Phong
'* License: Copyright (c) 2020-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: External interface class of PigBaseMini
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8.18
'* Create Time: 31/8/2019
'* The following is the code for creating PigBaseLocal
'* Public Class PigBaseLocal
'*     Inherits PigToolsLiteLib.PigBase
'*     Public Sub New(Version As String)
'*         MyBase.New(Version, System.Reflection.Assembly.GetExecutingAssembly().GetName.Version.ToString, System.Reflection.Assembly.GetExecutingAssembly().GetName.Name)
'*     End Sub
'* End Class
'************************************
Public Class PigBase
    Inherits PigBaseMini

    Public Sub New(Version As String, AppVersion As String, AppTitle As String)
        MyBase.New(Version, AppVersion, AppTitle)
    End Sub

    Public Shadows Function PrintDebugLog(SubName As String, LogInf As String) As String
        Return MyBase.PrintDebugLog(SubName, LogInf)
    End Function

    Public Shadows Function PrintDebugLog(SubName As String, StepName As String, LogInf As String) As String
        Return MyBase.PrintDebugLog(SubName, StepName, LogInf)
    End Function

    Public Shadows Function PrintDebugLog(SubName As String, StepName As String, LogInf As String, IsHardDebug As Boolean) As String
        Return MyBase.PrintDebugLog(SubName, StepName, LogInf, IsHardDebug)
    End Function

    Public Shadows Function PrintDebugLog(SubName As String, LogInf As String, IsHardDebug As Boolean) As String
        Return MyBase.PrintDebugLog(SubName, LogInf, IsHardDebug)
    End Function

    Public Shadows Sub ClearErr()
        MyBase.ClearErr()
    End Sub

    Public Shadows Sub SetDebug(DebugFilePath As String)
        MyBase.SetDebug(DebugFilePath)
    End Sub

    Public Shadows Sub SetDebug(DebugFilePath As String, IsHardDebug As Boolean)
        MyBase.SetDebug(DebugFilePath, IsHardDebug)
    End Sub

    Public Shadows Sub SetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        MyBase.SetSubErrInf(SubName, exIn, IsStackTrace)
    End Sub

    Public Shadows Sub SetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False)
        MyBase.SetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Sub


    Public Shadows Function GetSubErrInf(SubName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        Return MyBase.GetSubErrInf(SubName, exIn, IsStackTrace)
    End Function

    Public Shadows Function GetSubErrInf(SubName As String, StepName As String, ByRef exIn As System.Exception, Optional IsStackTrace As Boolean = False) As String
        Return MyBase.GetSubErrInf(SubName, StepName, exIn, IsStackTrace)
    End Function

    Public Shadows ReadOnly Property AppVersion() As String
        Get
            Return MyBase.AppVersion
        End Get
    End Property

    Public Shadows ReadOnly Property AppTitle() As String
        Get
            Return MyBase.AppTitle
        End Get
    End Property

    Public Shadows ReadOnly Property AppPath() As String
        Get
            Return MyBase.AppPath
        End Get
    End Property

    Public Shadows ReadOnly Property OsCrLf() As String
        Get
            Return MyBase.OsCrLf
        End Get
    End Property

    Public Shadows ReadOnly Property OsPathSep() As String
        Get
            Return MyBase.OsPathSep
        End Get
    End Property

    Public Shadows ReadOnly Property IsWindows() As Boolean
        Get
            Return MyBase.IsWindows
        End Get
    End Property

    Public Shadows ReadOnly Property IsDebug() As Boolean
        Get
            Return MyBase.IsDebug
        End Get
    End Property

    Public Shadows ReadOnly Property IsHardDebug() As Boolean
        Get
            Return MyBase.IsHardDebug
        End Get
    End Property

    Public Shadows Sub OpenDebug()
        MyBase.OpenDebug()
    End Sub

    Public Shadows Sub OpenDebug(LogFileDir As String)
        MyBase.OpenDebug(LogFileDir)
    End Sub

    Public Shadows Sub OpenDebug(LogFileDir As String, IsIncDate As Boolean)
        MyBase.OpenDebug(LogFileDir, IsIncDate)
    End Sub

    Public Shadows Sub OpenDebug(IsIncDate As Boolean)
        MyBase.OpenDebug(IsIncDate)
    End Sub

End Class
