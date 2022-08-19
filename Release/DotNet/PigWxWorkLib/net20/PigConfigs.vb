'**********************************
'* Name: PigConfigs
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigConfig 的集合类|Collection class of PigConfig
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.8
'* Create Time: 18/12/2021
'* 1.1    22/12/2020   Modify Add 
'* 1.2    23/12/2020   Add Parent, modify New,Add
'* 1.3    24/12/2020   Add Clear
'* 1.4    25/12/2021   Add AddOrGet
'* 1.5    26/12/2021   Modify AddOrGet
'* 1.6    3/1/2022   Modify AddOrGet
'* 1.7    21/2/2022   Modify Modify mAdd,Remove,Clear
'* 1.8    9/5/2022   Remove IsChange
'************************************
Public Class PigConfigs
    Inherits PigBaseMini
    Implements IEnumerable(Of PigConfig)
    Private Const CLS_VERSION As String = "1.8.2"
    Friend Property Parent As PigConfigSession
    Private ReadOnly moList As New List(Of PigConfig)

    Public Sub New(Parent As PigConfigSession)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of PigConfig) Implements IEnumerable(Of PigConfig).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigConfig
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(ConfName As String) As PigConfig
        Get
            Try
                Item = Nothing
                For Each oPigConfig As PigConfig In moList
                    If oPigConfig.ConfName = ConfName Then
                        Item = oPigConfig
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.ConfName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(ConfName) As Boolean
        Try
            IsItemExists = False
            For Each oPigConfig As PigConfig In moList
                If oPigConfig.ConfName = ConfName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As PigConfig) As String
        Try
            If Me.IsItemExists(NewItem.ConfName) = True Then Throw New Exception(NewItem.ConfName & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function


    Public Function AddOrGet(ConfName As String, ConfValue As String) As PigConfig
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(ConfName) = True Then
                AddOrGet = Me.Item(ConfName)
            Else
                AddOrGet = Me.Add(ConfName, ConfValue)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(ConfName As String, ConfValue As String) As PigConfig
        Dim LOG As New PigStepLog("Remove.ConfName.ConfValue")
        Try
            LOG.StepName = "New PigConfig"
            Add = New PigConfig(ConfName, ConfValue, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(ConfName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(ConfName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function AddOrGet(ConfName As String, ConfValue As String, ConfDesc As String) As PigConfig
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(ConfName) = True Then
                AddOrGet = Me.Item(ConfName)
            Else
                AddOrGet = Me.Add(ConfName, ConfValue, ConfDesc)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Add(ConfName As String, ConfValue As String, ConfDesc As String) As PigConfig
        Dim LOG As New PigStepLog("Remove.ConfName.ConfValue.ConfDesc")
        Try
            LOG.StepName = "New PigConfig"
            Add = New PigConfig(ConfName, ConfValue, ConfDesc, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(ConfName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(ConfName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(ConfName As String) As PigConfig
        Dim LOG As New PigStepLog("Remove.ConfName")
        Try
            LOG.StepName = "New PigConfig"
            Add = New PigConfig(ConfName, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(ConfName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(ConfName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(ConfName As String) As String
        Dim LOG As New PigStepLog("Remove.ConfName")
        Try
            LOG.StepName = "For Each"
            For Each oPigConfig As PigConfig In moList
                If oPigConfig.ConfName = ConfName Then
                    LOG.AddStepNameInf(ConfName)
                    moList.Remove(oPigConfig)
                    Exit For
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Remove(Index As Integer) As String
        Dim LOG As New PigStepLog("Remove.Index")
        Try
            LOG.StepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AddOrGet(ConfName As String) As PigConfig
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(ConfName) = True Then
                AddOrGet = Me.Item(ConfName)
            Else
                AddOrGet = Me.Add(ConfName)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Clear() As String
        Try
            moList.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function

End Class

