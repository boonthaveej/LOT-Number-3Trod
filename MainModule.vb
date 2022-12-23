
Imports System.Data
Imports System.Data.SqlClient
Module MainModule

End Module
Module DB_SQL_Connection
    Public str As String
    Public cmd As New SqlCommand
    Public rd As SqlDataReader
    Public sqlConn As New SqlConnection
    Public varIP As String = "185.78.164.192"
    Public varDB As String = "LotDB"
    Public varUSR As String = "lot_number"
    Public varPWD As String = "lotnumber1234"
    '--Dim connection As New SqlConnection("Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234")
    Sub ConDB()
        Try
            sqlConn = New SqlConnection("data source=" & varIP & ";initial catalog =" & varDB & ";user id =sa; password=" & varPWD)
            If sqlConn.State = ConnectionState.Closed Then
                sqlConn.Open()
            Else
                sqlConn.Close()
                sqlConn.Open()
            End If
        Catch ex As Exception
            MessageBox.Show(" ไม่สามารถติดต่อฐานข้อมูลได้ " & vbCrLf & Err.Description, "ERROR", MessageBoxButtons.OK)
            Exit Sub
        End Try
    End Sub

End Module
