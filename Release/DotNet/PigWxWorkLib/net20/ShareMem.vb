'**********************************
'* Name: ShareMem
'* Author: Part of the source code found on the Internet, I do not know who is the author, organized by Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Shared memory processing|共享内存处理
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 23/12/2019
'* 1.0.2  2019-12-24  联调
'* 1.0.3  2019-12-26  增加 IsInit 和修订BUG
'* 1.0.5  2020-1-5  增加 SMName As String, SMSize As Long
'* 1.0.6  2020-3-3    增加 write 和 read  的重载
'* 1.0.7  2020-3-19   增加 MaxSize
'* 1.0.8  2020-6-14   取消 Finalize ，不要执行Close
'* 1.0.9  1/2/2021   Err.Raise change to Throw New Exception|Err.Raise改为Throw New Exception
'* 1.1    2/2/2022   Modify Init
'**********************************

Imports System
Imports System.Runtime.InteropServices

Public Class ShareMem
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"
    Private mbolIsInit As Boolean
    Private mstrSMName As String
    Private mlngSMSize As Long
    Private mlngMaxSize As Long = 20971520

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Sub New(MaxSize As Long)
        MyBase.New(CLS_VERSION)
        If MaxSize <= 0 Then
            mlngSMSize = 20971520
        Else
            mlngSMSize = MaxSize
        End If
    End Sub


    ''' <summary>共享内存名</summary>
    Public ReadOnly Property SMName As String
        Get
            SMName = mstrSMName
        End Get
    End Property

    ''' <summary>共享内存大小</summary>
    Public ReadOnly Property SMSize As Long
        Get
            SMSize = mlngSMSize
        End Get
    End Property

    ''' <summary>是否初始化</summary>
    Public ReadOnly Property IsInit As Boolean
        Get
            IsInit = mbolIsInit
        End Get
    End Property

    Public Function Close() As String
        Try
            If Me.m_bInit Then
                ShareMem.UnmapViewOfFile(Me.m_pwData)
                ShareMem.CloseHandle(Me.m_hSharedMemoryFile)
            End If
            mbolIsInit = False
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Close", ex)
        End Try
    End Function

    <DllImport("Kernel32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function CloseHandle(ByVal handle As IntPtr) As Boolean
    End Function

    <DllImport("Kernel32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function CreateFileMapping(ByVal hFile As Integer, ByVal lpAttributes As IntPtr, ByVal flProtect As UInt32, ByVal dwMaxSizeHi As UInt32, ByVal dwMaxSizeLow As UInt32, ByVal lpName As String) As IntPtr
    End Function


    <DllImport("kernel32")>
    Private Shared Function GetLastError() As Integer
    End Function

    Public Function Init(SMName As String, SMSize As Long) As String
        Dim strStepName As String = ""
        Try
            If Me.IsWindows = False Then
                Throw New Exception("This class can only run on windows")
            End If
            If ((SMSize <= 0) OrElse (SMSize > mlngMaxSize)) Then
                SMSize = mlngMaxSize
            End If
            Me.m_MemSize = SMSize
            If (SMName.Length <= 0) Then
                Throw New Exception("Shared memory name not specified")
            Else
                strStepName = "CreateFileMapping"
                '                Me.m_hSharedMemoryFile = ShareMem.CreateFileMapping(-1, IntPtr.Zero, 4, 0, DirectCast(CUInt(SMSize), UInt32), SMName)
                Dim intSMSize As UInt32 = SMSize
                Me.m_hSharedMemoryFile = ShareMem.CreateFileMapping(-1, IntPtr.Zero, 4, 0, intSMSize, SMName)
                If (Me.m_hSharedMemoryFile = IntPtr.Zero) Then
                    Me.m_bAlreadyExist = False
                    Me.m_bInit = False
                    Throw New Exception("Zero2")
                Else
                    Me.m_bAlreadyExist = (ShareMem.GetLastError = &HB7)
                    strStepName = "MapViewOfFile"
                    Me.m_pwData = ShareMem.MapViewOfFile(Me.m_hSharedMemoryFile, 2, 0, 0, intSMSize)
                    If (Me.m_pwData = IntPtr.Zero) Then
                        Me.m_bInit = False
                        ShareMem.CloseHandle(Me.m_hSharedMemoryFile)
                        Throw New Exception("Zero3")
                    Else
                        Me.m_bInit = True
                        If Not Me.m_bAlreadyExist Then
                        End If
                    End If
                End If
            End If
            mstrSMName = SMName
            mlngSMSize = SMSize
            mbolIsInit = True
            Return "OK"
        Catch ex As Exception
            mbolIsInit = False
            Return Me.GetSubErrInf("Init", ex, True)
        End Try
    End Function

    <DllImport("Kernel32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function MapViewOfFile(ByVal hFileMapping As IntPtr, ByVal dwDesiredAccess As UInt32, ByVal dwFileOffsetHigh As UInt32, ByVal dwFileOffsetLow As UInt32, ByVal dwNumberOfBytesToMap As UInt32) As IntPtr
    End Function

    <DllImport("Kernel32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function OpenFileMapping(ByVal dwDesiredAccess As Integer, <MarshalAs(UnmanagedType.Bool)> ByVal bInheritHandle As Boolean, ByVal lpName As String) As IntPtr
    End Function

    Public Overloads Function Read(ByRef bytData As Byte()) As String
        Try
            If mbolIsInit = False Then Throw New Exception("Not Init")
            If bytData.Length <> mlngSMSize Then ReDim bytData(mlngSMSize - 1)
            Marshal.Copy(Me.m_pwData, bytData, 0, mlngSMSize)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Read", ex)
        End Try
    End Function

    Public Overloads Function Read(ByRef bytData As Byte(), Addr As Integer, DatSize As Integer) As String
        Try
            If mbolIsInit = False Then Throw New Exception("Not Init")
            If ((Addr + DatSize) > Me.m_MemSize) Then
                Throw New Exception("Data too long")
            ElseIf Not Me.m_bInit Then
                Throw New Exception("Data too long")
            Else
                Marshal.Copy(Me.m_pwData, bytData, Addr, DatSize)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Read", ex)
        End Try
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function SendMessage(ByVal hWnd As IntPtr, ByVal Msg As Integer, ByVal wParam As Integer, ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("Kernel32.dll", CharSet:=CharSet.Auto)>
    Private Shared Function UnmapViewOfFile(ByVal pvBaseAddress As IntPtr) As Boolean
    End Function

    Public Overloads Function Write(ByVal bytData As Byte()) As String
        Try
            If mbolIsInit = False Then Throw New Exception("Not Init")
            If bytData.Length <> mlngSMSize Then Throw New Exception("Data length must be " & mlngSMSize.ToString)
            Marshal.Copy(bytData, 0, Me.m_pwData, mlngSMSize)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Write", ex)
        End Try
    End Function

    Public Overloads Function Write(ByVal bytData As Byte(), ByVal Addr As Integer, ByVal DatSize As Integer) As String
        Try
            If mbolIsInit = False Then Throw New Exception("Not Init")
            If ((Addr + DatSize) > Me.m_MemSize) Then
                Throw New Exception("Data too long")
            ElseIf Not Me.m_bInit Then
                Throw New Exception("Not Init")
            Else
                Marshal.Copy(bytData, Addr, Me.m_pwData, DatSize)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Write", ex)
        End Try
    End Function


    ' Fields
    Private Const ERROR_ALREADY_EXISTS As Integer = &HB7
    Private Const FILE_MAP_COPY As Integer = 1
    Private Const FILE_MAP_WRITE As Integer = 2
    Private Const FILE_MAP_READ As Integer = 4
    Private Const FILE_MAP_ALL_ACCESS As Integer = 6
    Private Const PAGE_READONLY As Integer = 2
    Private Const PAGE_READWRITE As Integer = 4
    Private Const PAGE_WRITECOPY As Integer = 8
    Private Const PAGE_EXECUTE As Integer = &H10
    Private Const PAGE_EXECUTE_READ As Integer = &H20
    Private Const PAGE_EXECUTE_READWRITE As Integer = &H40
    Private Const SEC_COMMIT As Integer = &H8000000
    Private Const SEC_IMAGE As Integer = &H1000000
    Private Const SEC_NOCACHE As Integer = &H10000000
    Private Const SEC_RESERVE As Integer = &H4000000
    Private Const INVALID_HANDLE_VALUE As Integer = -1
    Private m_hSharedMemoryFile As IntPtr = IntPtr.Zero
    Private m_pwData As IntPtr = IntPtr.Zero
    Private m_bAlreadyExist As Boolean = False
    Private m_bInit As Boolean = False
    Private m_MemSize As Long = 0
End Class
