'**********************************
'* Name: PigVBCode
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: And generate VB code|且于生成VB的代码
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.2
'* Create Time: 16/6/2021
'* 1.1  1/7/2022    Modify MkCollectionClass
'* 1.2  6/7/2022    Modify MkCollectionClass
'**********************************
Public Class PigVBCode
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.2.8"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <summary>
    ''' 生成集合类代码|Generate collection class code
    ''' </summary>
    ''' <param name="MemberClassName">成员类名称|Member class name</param>
    ''' <param name="MemberClassKeyName">成员类键值，类型默认字符型|Member class key value, default character type</param>
    ''' <returns></returns>
    Public Function MkCollectionClass(ByRef OutVBCode As String, MemberClassName As String, MemberClassKeyName As String) As String
        Try
            OutVBCode = "Imports PigToolsLiteLib" & vbCrLf
            OutVBCode &= "Public Class " & MemberClassName & "s" & vbCrLf
            OutVBCode &= vbTab & "Inherits PigBaseLocal" & vbCrLf
            OutVBCode &= vbTab & "Implements IEnumerable(Of " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & "Private Const CLS_VERSION As String = ""1.0.0""" & vbCrLf
            OutVBCode &= vbTab & "Private ReadOnly moList As New List(Of " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & "Public Sub New()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "MyBase.New(CLS_VERSION)" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf
            OutVBCode &= vbTab & "Public ReadOnly Property Count() As Integer" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Get" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return moList.Count" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Count"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return -1" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Get" & vbCrLf
            OutVBCode &= vbTab & "End Property" & vbCrLf
            OutVBCode &= vbTab & "Public Function GetEnumerator() As IEnumerator(Of " & MemberClassName & ") Implements IEnumerable(Of " & MemberClassName & ").GetEnumerator" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Return moList.GetEnumerator()" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf
            OutVBCode &= vbTab & "Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Return Me.GetEnumerator()" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf
            OutVBCode &= vbTab & "Public ReadOnly Property Item(Index As Integer) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Get" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return moList.Item(Index)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Item.Index"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Get" & vbCrLf
            OutVBCode &= vbTab & "End Property" & vbCrLf

            OutVBCode &= vbTab & "Public ReadOnly Property Item(" & MemberClassKeyName & " As String) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Get" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Item = Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "For Each o" & MemberClassName & " As " & MemberClassName & " In moList" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "If o" & MemberClassName & "." & MemberClassKeyName & " = " & MemberClassKeyName & " Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Item = o" & MemberClassName & "" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "Exit For" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Next" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Item.Key"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Get" & vbCrLf
            OutVBCode &= vbTab & "End Property" & vbCrLf

            OutVBCode &= vbTab & "Public Function IsItemExists(" & MemberClassKeyName & ") As Boolean" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "IsItemExists = False" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "For Each o" & MemberClassName & " As " & MemberClassName & " In moList" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "If o" & MemberClassName & "." & MemberClassKeyName & " = " & MemberClassKeyName & " Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "IsItemExists = True" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "Exit For" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Next" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""IsItemExists"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Return False" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf

            OutVBCode &= vbTab & "Private Sub mAdd(NewItem As " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If Me.IsItemExists(NewItem." & MemberClassKeyName & ") = True Then Throw New Exception(NewItem." & MemberClassKeyName & " & ""Already exists"")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "moList.Add(NewItem)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""mAdd"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Sub Add(NewItem As " & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Me.mAdd(NewItem)" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Function AddOrGet(" & MemberClassKeyName & " As String) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim LOG As New PigStepLog(""AddOrGet"")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If Me.IsItemExists(" & MemberClassKeyName & ") = True Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Me.Item(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Else" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "Return Me.Add(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf

            OutVBCode &= vbTab & "Public Function Add(" & MemberClassKeyName & " As String) As " & MemberClassName & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim LOG As New PigStepLog(""Add"")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "LOG.StepName = ""New " & MemberClassName & """" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Dim o" & MemberClassName & " As New " & MemberClassName & "(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If o" & MemberClassName & ".LastErr <> """" Then Throw New Exception(o" & MemberClassName & ".LastErr)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "LOG.StepName = ""mAdd""" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.mAdd(o" & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "If Me.LastErr <> """" Then Throw New Exception(Me.LastErr)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Add = o" & MemberClassName & "" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Return Nothing" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Function" & vbCrLf

            OutVBCode &= vbTab & "Private Sub Remove(" & MemberClassKeyName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim strStepName As String = """"" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "strStepName = ""For Each""" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "For Each o" & MemberClassName & " As " & MemberClassName & " In moList" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "If o" & MemberClassName & "." & MemberClassKeyName & " = " & MemberClassKeyName & " Then" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "strStepName = ""Remove "" & " & MemberClassKeyName & "" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "moList.Remove(o" & MemberClassName & ")" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & vbTab & "Exit For" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & vbTab & "End If" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Next" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Remove.Key"", strStepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Sub Remove(Index As Integer)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Dim strStepName As String = """"" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "strStepName = ""Index="" & Index.ToString" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "moList.RemoveAt(Index)" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Remove.Index"", strStepName, ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= vbTab & "Public Sub Clear()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Try" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "moList.Clear()" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.ClearErr()" & vbCrLf
            OutVBCode &= vbTab & vbTab & "Catch ex As Exception" & vbCrLf
            OutVBCode &= vbTab & vbTab & vbTab & "Me.SetSubErrInf(""Clear"", ex)" & vbCrLf
            OutVBCode &= vbTab & vbTab & "End Try" & vbCrLf
            OutVBCode &= vbTab & "End Sub" & vbCrLf

            OutVBCode &= "End Class"

            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("", ex)
        End Try
    End Function

End Class
