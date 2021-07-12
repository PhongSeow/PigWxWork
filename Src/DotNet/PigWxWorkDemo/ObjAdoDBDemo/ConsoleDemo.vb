Imports ObjAdoDBLib

Public Class ConsoleDemo
    Public Conn As New Connection
    Public ConnStr As String
    Public SQL As String
    Public RS As Recordset

    Public Sub Main()
        Do While True
            Console.WriteLine("*******************")
            Console.WriteLine("Main menu")
            Console.WriteLine("*******************")
            Console.WriteLine("Press Q to Exit")
            Console.WriteLine("Press A to Set Connection String")
            Console.WriteLine("Press B to Open Connection")
            Console.WriteLine("Press C to Show Connection Information")
            Console.WriteLine("Press D to Create Recordset with Execute")
            Console.WriteLine("Press E to Show Recordset Information")
            Console.WriteLine("Press F to Recordset.MoveNext")
            Console.WriteLine("*******************")
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Q
                    Exit Do
                Case ConsoleKey.A
                    Console.WriteLine("*******************")
                    Console.WriteLine("Set Connection String")
                    Console.WriteLine("*******************")
                    Console.WriteLine("Input SQL Server:")
                    Dim strDBSrv As String = Console.ReadLine()
                    Console.WriteLine("Input DB User:")
                    Dim strDBUser As String = Console.ReadLine()
                    Console.WriteLine("Input DB Password:")
                    Dim strDBPwd As String = Console.ReadLine()
                    Console.WriteLine("Input Default DB:")
                    Dim strDefaDB As String = Console.ReadLine()
                    Me.ConnStr = "Provider=Sqloledb;Data Source=" & strDBSrv & ";Database=" & strDefaDB & ";User ID=" & strDBUser & ";Password=" & strDBPwd
                    'Me.ConnStr = "Provider=SQLNCLI10;Data Source=" & strDBSrv & ";Database=" & strDefaDB & ";User ID=" & strDBUser & ";Password=" & strDBPwd
                    With Me.Conn
                        .ConnectionTimeout = 5
                        .ConnectionString = Me.ConnStr
                    End With
                Case ConsoleKey.B
                    Console.WriteLine("#################")
                    Console.WriteLine("Open Connection")
                    Console.WriteLine("#################")
                    With Me.Conn
                        .Open()
                        If .LastErr <> "" Then
                            Console.WriteLine("Connect Error:" & .LastErr)
                        Else
                            Console.WriteLine("Connect OK")
                        End If
                    End With
                Case ConsoleKey.C
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Connection Information")
                    Console.WriteLine("#################")
                    Console.WriteLine("ConnectionString=" & Me.Conn.ConnectionString)
                    Console.WriteLine("State=" & Me.Conn.State)
                Case ConsoleKey.D
                    Console.WriteLine("#################")
                    Console.WriteLine("Create Recordset with Execute")
                    Console.WriteLine("#################")
                    Console.WriteLine("Input SQL:")
                    Me.SQL = Console.ReadLine()
                    With Me.Conn
                        Me.RS = .Execute(SQL)
                        If .LastErr <> "" Then
                            Console.WriteLine("Execute Error:" & .LastErr)
                        Else
                            Console.WriteLine("Execute OK")
                        End If
                    End With
                Case ConsoleKey.E
                    Console.WriteLine("#################")
                    Console.WriteLine("Show Recordset Information")
                    Console.WriteLine("#################")
                    With Me.RS
                        Console.WriteLine("Fields.Count=" & .Fields.Count)
                        If .Fields.Count > 0 Then
                            Dim i As Integer
                            For i = 0 To .Fields.Count - 1
                                Console.WriteLine(".Fields.Item(" & i & ").Name=" & .Fields.Item(i).Name & "[" & .Fields.Item(i).Value.ToString & "]")
                            Next
                        End If
                        Console.WriteLine("PageCount=" & .PageCount)
                        Console.WriteLine("EOF=" & .EOF)
                    End With
                Case ConsoleKey.F
                    Console.WriteLine("#################")
                    Console.WriteLine("Recordset.MoveNext")
                    Console.WriteLine("#################")
                    With Me.RS
                        .MoveNext()
                        If .LastErr <> "" Then
                            Console.WriteLine("MoveNext Error:" & .LastErr)
                        Else
                            Console.WriteLine("MoveNext OK")
                        End If
                    End With
            End Select
        Loop
    End Sub

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
