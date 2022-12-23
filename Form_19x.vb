Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Public Class Form_19x
    Private strConn As String = "Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234"
    Private sqlCon As SqlConnection
    Private Sub ButtonEnter_Click(sender As Object, e As EventArgs) Handles ButtonEnter.Click
        On Error GoTo ErrorHandler   ' Enable error-handling routine.
        sqlCon = New SqlConnection(strConn)
        Using (sqlCon)
            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon
            sqlComm.CommandText = "InsertDataLoop"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("Lot_ID", Now.ToString("yyyyMMdd"))
            sqlComm.Parameters.AddWithValue("Lot_Number", Trim(TextBoxNumber.Text))
            sqlComm.Parameters.AddWithValue("Lot_Val_1", Trim(TextBoxTop.Text))
            sqlComm.Parameters.AddWithValue("Lot_Val_2", Trim(TextBoxDown.Text))
            sqlComm.Parameters.AddWithValue("Lot_Date", Now.ToString("yyyyMMddhhmmss"))
            sqlComm.Parameters.AddWithValue("Login_ID", "1")
            sqlCon.Open()
            sqlComm.ExecuteNonQuery()
            sqlCon.Close()
            Call ClearText()
            Me.TopMost = False
            Me.Hide()
            Form_Input.ButtonRefresh.PerformClick()
        End Using
        Exit Sub      ' Exit to avoid handler.
ErrorHandler:  ' Error-handling routine.
        MessageBox.Show(Err.Number & "//" & Err.Description)
        Resume Next  ' Resume execution at the statement immediately 
    End Sub
    Private Sub ClearText()
        On Error Resume Next
        Me.TextBoxNumber.Text = "0"
        Me.TextBoxTop.Text = "0"
        Me.TextBoxDown.Text = "0"
    End Sub
    Private Sub Form_19x_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.TopMost = True
        Call ClearText()
        Me.TextBoxNumber.Focus()
    End Sub

    Private Sub TextBoxNumber_GotFocus(sender As Object, e As EventArgs) Handles TextBoxNumber.GotFocus
        On Error Resume Next
        TextBoxNumber.SelectAll()
    End Sub

    Private Sub TextBoxNumber_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxNumber.KeyDown
        On Error Resume Next
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxNumber_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxNumber.KeyPress
        On Error Resume Next
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลขนะ")
        End Select
    End Sub

    Private Sub TextBoxNumber_TextChanged(sender As Object, e As EventArgs) Handles TextBoxNumber.TextChanged

    End Sub

    Private Sub TextBoxDown_GotFocus(sender As Object, e As EventArgs) Handles TextBoxDown.GotFocus
        On Error Resume Next
        TextBoxDown.SelectAll()
    End Sub

    Private Sub TextBoxTop_GotFocus(sender As Object, e As EventArgs) Handles TextBoxTop.GotFocus
        On Error Resume Next
        TextBoxTop.SelectAll()
    End Sub

    Private Sub TextBoxDown_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxDown.KeyDown
        On Error Resume Next
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxDown_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxDown.KeyPress
        On Error Resume Next
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลขนะ")
        End Select
    End Sub

    Private Sub TextBoxTop_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxTop.KeyDown
        On Error Resume Next
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxTop_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxTop.KeyPress
        On Error Resume Next
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลขนะ")
        End Select
    End Sub

    Private Sub TextBoxDown_TextChanged(sender As Object, e As EventArgs) Handles TextBoxDown.TextChanged

    End Sub

    Private Sub TextBoxTop_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTop.TextChanged

    End Sub
End Class