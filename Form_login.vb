Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.IO.TextReader
Imports System.IO.TextWriter
Imports System.Runtime.CompilerServices
Public Class Form_login
    Private strConn As String = "Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234"
    Private sqlCon As SqlConnection
    Dim userid As Char
    Dim userpas As Char
    Public Sub ClearLogin()
        On Error Resume Next
        Me.TextUsername.Text = ""
        Me.TextPassword.Text = ""
        userid = ""
        userpas = ""
    End Sub
    Public Sub CallUpdateLastLogin(LoginID As String)
        On Error Resume Next
        sqlCon = New SqlConnection(strConn)
        Using (sqlCon)
            Dim sqlComm As New SqlCommand
            sqlComm.Connection = sqlCon
            sqlComm.CommandText = "UpdateLastLogin"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("LoginID", Val(LoginID))
            sqlComm.Parameters.AddWithValue("LastLogin", Now.ToString("G"))
            'Console.WriteLine(Now.ToString("yyyyMMddHHmmss"))
            sqlCon.Open()
            sqlComm.ExecuteNonQuery()
        End Using
        sqlCon.Close()
    End Sub
    Public Sub CallLogin()
        On Error Resume Next
        If TextUsername.Text <> "" And TextPassword.Text <> "" Then
            Dim connection As New SqlConnection(strConn)
            Dim command As New SqlCommand("select LoginID,LoginUser from [LotDB].[dbo].[TblLOG] where  LoginUser = @username and LoginPass = @password", connection)
            command.Parameters.Add("@username", SqlDbType.VarChar).Value = Trim(Me.TextUsername.Text)
            command.Parameters.Add("@password", SqlDbType.VarChar).Value = Trim(Me.TextPassword.Text)
            Dim adapter As New SqlDataAdapter(command)
            Dim table As New DataTable()
            adapter.Fill(table)
            '---- First row.
            Dim FirstRow As DataRow = table.Rows(0)
            If table.Rows.Count() <= 0 Then
                MessageBox.Show("Username or Password is Invalid!")
                Call ClearLogin()
                TextUsername.Focus()
            Else
                CallUpdateLastLogin(FirstRow("LoginID"))
                'MessageBox.Show("Welcome Khun.[ " & FirstRow("LoginUser") & " ] to your Job")
                '== Console.WriteLine(row1("Breed"))
                connection.Close()
                Me.Hide()
                Form_Input.Text = "WELCOME BACK <<<<< K." & FirstRow("LoginUser") & ">>>>> Please Let's to do your job role! | Today is " & Now.ToString("D") & " --Enjoy--"
                Form_Input.LabelLoginID.Text = FirstRow("LoginID")
                Form_Input.LabelLoginName.Text = FirstRow("LoginUser")
                Form_Input.Show()
            End If
        Else
            MessageBox.Show("Please insert your Username and Password again!")
            Call ClearLogin()
            TextUsername.Focus()
        End If
    End Sub
    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        If MessageBox.Show("Do you want to exit this program?", "", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Application.Exit()
        Else
            Call ClearLogin()
            Me.TextUsername.Focus()
        End If
    End Sub

    Private Sub ButtonLogin_Click(sender As Object, e As EventArgs) Handles ButtonLogin.Click
        Call CallLogin()
    End Sub

    Private Sub TextUsername_KeyDown(sender As Object, e As KeyEventArgs) Handles TextUsername.KeyDown
        'On Error Resume Next
        '--vbKeyTab 9 TAB key
        '--vbKeyClear 12 CLEAR key
        '--vbKeyReturn 13 ENTER key
        '--vbKeyShift 16 SHIFT key
        '--vbKeyMultiply 106 MULTIPLICATION SIGN key ( * ) ปุ่มเครื่องหมายคูณ
        '--vbKeyAdd 107 PLUS SIGN key ( + ) ปุ่มเครื่องหมายบวก
        '--vbKeySeparator 108 ENTER (keypad) key ปุ่ม Enter ทาง Keypad
        '--vbKeySubtract 109 MINUS SIGN key ( - ) ปุ่มเครื่องหมายลบ
        '--vbKeyDecimal 110 DECIMAL POINT key ( . ) ปุ่มเครื่องหมายจุด
        '--vbKeyDivide 111 DIVISION SIGN key ( / ) ปุ่มเครื่องหมายหาร
        '--vbKeyUp 38 UP ARROW key
        '--vbKeyRight 39 RIGHT ARROW key
        '--vbKeyDown 40 DOWN ARROW key
        On Error Resume Next
        If e.KeyCode = 13 Then '-- Key Enter
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextUsername_TextChanged(sender As Object, e As EventArgs) Handles TextUsername.TextChanged

    End Sub

    Private Sub TextPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles TextPassword.KeyDown
        'On Error Resume Next
        '--vbKeyTab 9 TAB key
        '--vbKeyClear 12 CLEAR key
        '--vbKeyReturn 13 ENTER key
        '--vbKeyShift 16 SHIFT key
        '--vbKeyMultiply 106 MULTIPLICATION SIGN key ( * ) ปุ่มเครื่องหมายคูณ
        '--vbKeyAdd 107 PLUS SIGN key ( + ) ปุ่มเครื่องหมายบวก
        '--vbKeySeparator 108 ENTER (keypad) key ปุ่ม Enter ทาง Keypad
        '--vbKeySubtract 109 MINUS SIGN key ( - ) ปุ่มเครื่องหมายลบ
        '--vbKeyDecimal 110 DECIMAL POINT key ( . ) ปุ่มเครื่องหมายจุด
        '--vbKeyDivide 111 DIVISION SIGN key ( / ) ปุ่มเครื่องหมายหาร
        '--vbKeyUp 38 UP ARROW key
        '--vbKeyRight 39 RIGHT ARROW key
        '--vbKeyDown 40 DOWN ARROW key
        On Error Resume Next
        If e.KeyCode = 13 Then '-- Key Enter
            e.SuppressKeyPress = True
            Call CallLogin()
        End If
    End Sub

    Private Sub TextPassword_TextChanged(sender As Object, e As EventArgs) Handles TextPassword.TextChanged

    End Sub
End Class
