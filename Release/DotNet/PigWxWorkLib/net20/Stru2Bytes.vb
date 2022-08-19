'**********************************
'* Name: Stru2Bytes
'* Author: Part of the source code found on the Internet, I do not know who is the author, organized by Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Structure and byte array conversion 结构与字节数组转换

'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 25/10/2019
'1.0.2  2019-11-2   修改 Stru2Bytes
'1.0.3  2019-11-27   Bytes2Stru
'**********************************

Imports System.Runtime.InteropServices
Imports System.IO
Public Class Stru2Bytes
    '256 bytes
    Private Structure sSftDataHeader
        '96 bytes
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim bDate() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim bTime() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim bAID() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim bEID() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim bWeight() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=8)>
        Dim bSexAge() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
        Dim bComment() As Byte
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=16)>
        Dim bIMPID() As Byte
        '8 bytes
        Dim nSamplingRate As Short
        Dim nChnums As Byte
        Dim nDataMode As Byte
        Dim type1 As Byte
        Dim type2 As Byte
        Dim type3 As Byte
        Dim type4 As Byte

        '48 bytes
        Dim Cal1 As Short
        Dim Upper1 As Short
        Dim Lower1 As Short
        Dim Base1 As Short
        Dim wave_color_1 As UInteger


        Dim Cal2 As Short
        Dim Upper2 As Short
        Dim Lower2 As Short
        Dim Base2 As Short
        Dim wave_color_2 As UInteger


        Dim Cal3 As Short
        Dim Upper3 As Short
        Dim Lower3 As Short
        Dim Base3 As Short
        Dim wave_color_3 As UInteger

        Dim Cal4 As Short
        Dim Upper4 As Short
        Dim Lower4 As Short
        Dim Base4 As Short
        Dim wave_color_4 As UInteger

        '104bytes
        <MarshalAs(UnmanagedType.ByValArray, SizeConst:=104)>
        Dim dummy() As Byte

        Private Sub New(ByRef size As Integer)
            ReDim dummy(103)
            ReDim bDate(8)
            ReDim bTime(8)
            ReDim bAID(8)
            ReDim bEID(8)
            ReDim bWeight(8)
            ReDim bSexAge(8)
            ReDim bComment(32)
            ReDim bIMPID(16)
        End Sub
    End Structure


    Private g_SH As sSftDataHeader
    'Private m_fn As FileStream
    'Private m_fr As BinaryReader
    'Private m_fw As BinaryWriter
    ''' <summary>结构体转换成字节数组</summary>
    ''' <param name="TarBytes">要转化的结构体</param>
    Public Function Stru2Bytes(SrcStru As Object, ByRef TarBytes As Byte()) As String
        Try
            Dim oSize As Integer = Marshal.SizeOf(SrcStru)
            Dim tByte(oSize - 1) As Byte
            Dim tPtr As IntPtr = Marshal.AllocHGlobal(oSize)
            Marshal.StructureToPtr(SrcStru, tPtr, False)
            Marshal.Copy(tPtr, tByte, 0, oSize)
            Marshal.FreeHGlobal(tPtr)
            TarBytes = tByte
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>字节数组转换成结构体</summary>
    ''' <param name="SrcBytes">源要转化的字节数组</param>
    ''' <param name="TarType">源要转化成的结构体类型，格式：目标结构体.GetType</param>
    ''' <param name="TarStru">目标结构体</param>
    Public Function Bytes2Stru(SrcBytes() As Byte, TarType As Type, ByRef TarStru As Object) As String
        'TarType，例：struAny.GetType
        Try
            Dim oSize As Integer = Marshal.SizeOf(TarType)
            If oSize > SrcBytes.Length Then
                Err.Raise(-1, , "源数据长度不足")
            End If

            Dim tPtr As IntPtr = Marshal.AllocHGlobal(oSize)
            Marshal.Copy(SrcBytes, 0, tPtr, oSize)
            Dim oStructure As Object = Marshal.PtrToStructure(tPtr, TarType)
            Marshal.FreeHGlobal(tPtr)
            TarStru = oStructure
            Return "OK"
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function
End Class
