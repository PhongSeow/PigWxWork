'**********************************
'* Name: PigCompressor
'* Author: Seow Phong
'* License: Copyright (c) 2019-2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Compression processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 11/12/2019
'* 1.0.2  17/10/2020  optimization
'* 1.1    16/10/2021  Modify Depress
'* 1.2    3/2/2022  Modify Compress
'**********************************

Imports System
Imports System.IO
Imports System.IO.Compression
Imports System.Text

Public Class PigCompressor
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.6"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function Compress(SrcStr As String) As Byte()
        Dim abAny As Byte() = Encoding.UTF8.GetBytes(SrcStr)
        Return Me.Compress(abAny)
    End Function

    Public Function Compress(SrcStr As String, TextType As PigText.enmTextType) As Byte()
        Dim abAny As Byte()
        Select Case TextType
            Case PigText.enmTextType.Ascii
                abAny = Encoding.ASCII.GetBytes(SrcStr)
            Case PigText.enmTextType.Unicode
                abAny = Encoding.Unicode.GetBytes(SrcStr)
            Case PigText.enmTextType.UTF8
                abAny = Encoding.UTF8.GetBytes(SrcStr)
            Case Else
                abAny = Encoding.UTF8.GetBytes(SrcStr)
        End Select
        Return Me.Compress(abAny)
    End Function

    Public Function Compress(SrcBytes As Byte()) As Byte()
        Try
            Using msAny As MemoryStream = New MemoryStream
                Using gzsAny As GZipStream = New GZipStream(msAny, CompressionMode.Compress)
                    gzsAny.Write(SrcBytes, 0, SrcBytes.Length)
                    gzsAny.Flush()
                End Using
                Return msAny.ToArray
            End Using
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Compress", ex)
            Return Nothing
        End Try
    End Function

    Public Function Depress(CompressBytes As Byte()) As Byte()
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
        Try
            Dim buffer As Byte()
            Using msAny As MemoryStream = New MemoryStream(CompressBytes)
                Using msAny2 As MemoryStream = New MemoryStream
                    Using gzsAny As GZipStream = New GZipStream(msAny, CompressionMode.Decompress)
                        gzsAny.CopyTo(msAny2)
                        gzsAny.Flush()
                    End Using
                    buffer = msAny2.ToArray
                End Using
            End Using
            Me.ClearErr()
            Return buffer
        Catch ex As Exception
            Me.SetSubErrInf("Depress", ex)
            Return Nothing
        End Try
#Else
        Try
            Throw New Exception("Need to run in .net 4.0 or higher framework")
        Catch ex As Exception
            Me.SetSubErrInf("Depress", ex)
            Return Nothing
        End Try
#End If
    End Function

End Class
