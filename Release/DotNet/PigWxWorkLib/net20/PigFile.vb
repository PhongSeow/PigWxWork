'**********************************
'* Name: PigFile
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: File processing,Handle file reading, writing, information, etc
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 4/11/2019
'*1.0.2  2019-11-5   增加mSaveFile
'*1.0.3  2019-11-20  增加 CopyFileTo
'*1.0.5  2020-2-13   增加 IsPigMD5
'*1.0.6  2020-2-15   部分功能迁到 fGEFileSystem
'*1.0.7  2020-3-5    增加 SetData
'*1.0.8  2020-10-27  修改 LoadFile
'*1.0.9  1/2/2021 Err.Raise change to Throw New Exception|Err.Raise改为Throw New Exception
'*1.0.10  26/7/2021 Modify LoadFile
'*1.1  10/5/2022    Add PigMD5, modify LoadFile
'**********************************
Imports System.IO
Public Class PigFile
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.2"
    Private mstrFilePath As String '文件路径
    Private moFileInfo As FileInfo '文件信息
    Public GbMain As PigBytes '主数据数组

    Private mintKeepFileVerCnt As Integer = 10 '保留文件的版本数
    Private mintKeepFileDays As Integer = 365 '保留文件的时间，以天为单位。

    Public Sub New(FilePath As String)
        MyBase.New(CLS_VERSION)
        mstrFilePath = FilePath
    End Sub

    ''' <summary>保留文件的时间，以天为单位。</summary>
    Public Property KeepFileDays As Integer
        Get
            KeepFileDays = mintKeepFileDays
        End Get
        Set(value As Integer)
            If value <= 0 Then value = 1
            mintKeepFileDays = value
        End Set
    End Property


    ''' <summary>保留文件的版本数</summary>
    Public Property KeepFileVerCnt As Integer
        Get
            KeepFileVerCnt = mintKeepFileVerCnt
        End Get
        Set(value As Integer)
            If value <= 0 Then value = 1
            mintKeepFileVerCnt = value
        End Set
    End Property

    ''' <summary>复制文件到</summary>
    ''' <param name="TarFile">目标文件路径</param>
    Public Overloads Function CopyFileTo(TarFile As String) As String
        Try
            System.IO.File.Copy(Me.FilePath, TarFile, True)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CopyFileTo", ex, False)
        End Try
    End Function

    ''' <summary>复制文件到</summary>
    ''' <param name="TarFile">目标文件路径</param>
    ''' <param name="IsOverWrite">是否强制覆盖</param>
    Public Overloads Function CopyFileTo(TarFile As String, IsOverWrite As Boolean) As String
        Try
            System.IO.File.Copy(Me.FilePath, TarFile, IsOverWrite)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CopyFileTo", ex, False)
        End Try
    End Function

    ''' <summary>设置数据</summary>
    Public Overloads Function SetData(DataBytes As Byte()) As String
        Try
            If Not Me.GbMain Is Nothing Then Me.GbMain = Nothing
            Me.GbMain = New PigBytes(DataBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SetData", ex, False)
        End Try
    End Function

    ''' <summary>设置数据</summary>
    Public Overloads Function SetData(GETextData As PigText) As String
        Try
            If Not Me.GbMain Is Nothing Then Me.GbMain = Nothing
            Me.GbMain = New PigBytes(GETextData.TextBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SetData", ex, False)
        End Try
    End Function


    ''' <summary>导入数据</summary>
    Public Function LoadFile() As String
        Dim LOG As New PigStepLog("LoadFile")
        Try
            If Me.GbMain IsNot Nothing Then Me.GbMain = Nothing
            LOG.StepName = "New FileStream"
            Dim sfAny As New FileStream(mstrFilePath, FileMode.Open, FileAccess.Read, FileShare.Read)
            LOG.StepName = "New BinaryReader"
            Dim brAny = New BinaryReader(sfAny)
            LOG.StepName = "New PigBytes"
            Me.GbMain = New PigBytes
            LOG.StepName = "GbMain ReadBytes"
            GbMain.Main = brAny.ReadBytes(Me.Size)
            LOG.StepName = "Close"
            brAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(mstrFilePath)
            If Me.IsDebug Or Me.IsHardDebug Then
                Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex, True)
            Else
                Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            End If
        End Try
    End Function

    ''' <summary>保存数据</summary>
    ''' <param name="IsKeepVerFile">是否保留版本文件</param>
    Public Overloads Function SaveFile(IsKeepVerFile As Boolean) As String
        SaveFile = Me.mSaveFile(IsKeepVerFile)
    End Function

    ''' <summary>保存数据</summary>
    Public Overloads Function SaveFile() As String
        SaveFile = Me.mSaveFile(True)
    End Function

    ''' <summary>保存数据</summary>
    ''' <param name="IsKeepVerFile">是否保留版本文件</param>
    Private Function mSaveFile(IsKeepVerFile As Boolean) As String
        Dim strStepName As String = ""
        Try
            strStepName = "Check the data"    '检查数据
            If Me.GbMain Is Nothing Then Throw New Exception("No data")
            If IsKeepVerFile = True Then
                strStepName = "Locate the current backup file"  '定位当前备份文件
                Dim i As Integer, strBakFile As String = "", bolIsFind As Boolean = False
                For i = 1 To Me.KeepFileVerCnt
                    strBakFile = Me.FilePath & "." & i.ToString
                    If Me.IsExists(strBakFile) = False Then
                        bolIsFind = True
                        Exit For
                    End If
                Next
                strStepName = "Locating backup file"   '定位备份文件
                If bolIsFind = False Then
                    strBakFile = Me.FilePath & "." & Me.KeepFileVerCnt.ToString
                    Dim oGEFile As New PigFile(strBakFile)
                    If oGEFile.IsExists = False Then
                        bolIsFind = True
                    Else
                        For i = 1 To Me.KeepFileVerCnt
                            strBakFile = Me.FilePath & "." & i.ToString
                            Dim oGEFile2 As New PigFile(strBakFile)
                            If oGEFile2.IsExists = False Then
                                bolIsFind = True
                                Exit For
                            End If
                            If oGEFile2.UpdateTime < oGEFile.UpdateTime Then
                                bolIsFind = True
                                Exit For
                            End If
                            oGEFile2 = Nothing
                        Next
                        If bolIsFind = False Then
                            bolIsFind = True
                            strBakFile = Me.FilePath & ".1"
                        End If
                    End If
                End If
                If bolIsFind = False Then Throw New Exception("Unable to determine backup file")
                strStepName = "Backup" & Me.FilePath & " to " & strBakFile

            End If
            strStepName = "New FileStream(" & mstrFilePath & ")"
            Dim sfAny As New FileStream(mstrFilePath, FileMode.Create, FileAccess.Write, FileShare.None, Me.GbMain.Main.Length, False)
            'strStepName = "New StreamWriter"
            'Dim swAny = New StreamWriter(sfAny)
            strStepName = "New BinaryWriter"
            Dim bwAny = New BinaryWriter(sfAny)
            strStepName = "Write Bytes"
            bwAny.Write(Me.GbMain.Main)
            strStepName = "Close"
            bwAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mSaveFile", ex, False)
        End Try
    End Function


    ''' <summary>本文件MD5</summary>
    Public ReadOnly Property MD5 As String
        Get
            Return Me.mGetMyMD5(False)
        End Get
    End Property

    ''' <summary>本文件豚豚MD5</summary>
    Public ReadOnly Property PigMD5 As String
        Get
            Return Me.mGetMyMD5(True)
        End Get
    End Property

    ''' <summary>获取本文件MD5</summary>
    ''' <param name="IsPigMD5">是否豚豚MD5</param>
    Private Function mGetMyMD5(IsPigMD5 As Boolean) As String
        Try
            If Me.GbMain Is Nothing Then
                Dim strRet As String = Me.LoadFile
                If strRet <> "OK" Then Throw New Exception(strRet)
            End If
            Dim oMD5 As New PigMD5(Me.GbMain.Main)
            If IsPigMD5 = True Then
                mGetMyMD5 = oMD5.PigMD5
            Else
                mGetMyMD5 = oMD5.MD5
            End If
            oMD5 = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mGetMyMD5", ex, False)
            Return ""
        End Try
    End Function


    ''' <summary>文件完整路径</summary>
    Public ReadOnly Property FilePath As String
        Get
            FilePath = mstrFilePath
        End Get
    End Property

    Private Sub mInitFileInf()
        Try
            If moFileInfo Is Nothing Then
                moFileInfo = New FileInfo(mstrFilePath)
            Else
                moFileInfo.Refresh()
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mInitFileInf", ex, False)
        End Try
    End Sub

    ''' <summary>文件的目录路径</summary>
    Public ReadOnly Property DirPath() As String
        Get
            Try
                Me.mInitFileInf()
                DirPath = moFileInfo.DirectoryName
            Catch ex As Exception
                DirPath = ""
            End Try
        End Get
    End Property

    ''' <summary>文件名，不</summary>
    Public ReadOnly Property FileTitle() As String
        Get
            Try
                Me.mInitFileInf()
                FileTitle = moFileInfo.Name
            Catch ex As Exception
                FileTitle = ""
            End Try
        End Get
    End Property

    ''' <summary>文件的扩展名</summary>
    Public ReadOnly Property ExtName() As String
        Get
            Try
                Me.mInitFileInf()
                ExtName = moFileInfo.Extension
            Catch ex As Exception
                ExtName = ""
            End Try
        End Get
    End Property

    ''' <summary>文件大小</summary>
    Public ReadOnly Property Size() As Long
        Get
            Try
                Me.mInitFileInf()
                Size = moFileInfo.Length
            Catch ex As Exception
                Size = -1
            End Try
        End Get
    End Property

    ''' <summary>文件的创建时间</summary>
    Public ReadOnly Property CreationTime() As DateTime
        Get
            Try
                Me.mInitFileInf()
                CreationTime = moFileInfo.CreationTime
            Catch ex As Exception
                CreationTime = #1/1/1900#
            End Try
        End Get
    End Property

    ''' <summary>文件的更新时间</summary>
    Public ReadOnly Property UpdateTime() As DateTime
        Get
            Try
                Me.mInitFileInf()
                UpdateTime = moFileInfo.LastWriteTime
            Catch ex As Exception
                UpdateTime = #1/1/1900#
            End Try
        End Get
    End Property

    ''' <summary>文件是否存在</summary>
    Public Overloads ReadOnly Property IsExists As Boolean
        Get
            Try
                Me.mInitFileInf()
                IsExists = moFileInfo.Exists
            Catch ex As Exception
                IsExists = False
            End Try
        End Get
    End Property

    ''' <summary>文件是否存在</summary>
    ''' <param name="FilePath">指定文件路径</param>
    Public Overloads ReadOnly Property IsExists(FilePath As String) As Boolean
        Get
            Try
                Dim oFileInfo = New FileInfo(FilePath)
                IsExists = oFileInfo.Exists
                oFileInfo = Nothing
            Catch ex As Exception
                IsExists = False
            End Try
        End Get
    End Property



End Class
