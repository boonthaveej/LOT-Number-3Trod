Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Public Class Form_6x
    Private strConn As String = "Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234"
    Private sqlCon As SqlConnection
    Private Sub ButtonEnter_Click(sender As Object, e As EventArgs) Handles ButtonEnter.Click
        On Error GoTo ErrorHandler   ' Enable error-handling routine.
        sqlCon = New SqlConnection(strConn)
        Dim st As String = Trim(Me.TextBoxNumber.Text)
        Dim st1 As String = st.Substring(0, 1)
        Dim st2 As String = st.Substring(1, 1)
        Dim st3 As String = st.Substring(2, 1)
        Using (sqlCon)
            Dim sqlComm As New SqlCommand()
            sqlComm.Connection = sqlCon
            sqlComm.CommandText = "Combinations3Val"
            sqlComm.CommandType = CommandType.StoredProcedure
            sqlComm.Parameters.AddWithValue("Lot_N_1", st1)
            sqlComm.Parameters.AddWithValue("Lot_N_2", st2)
            sqlComm.Parameters.AddWithValue("Lot_N_3", st3)
            sqlComm.Parameters.AddWithValue("Val1", Trim(TextBoxTop.Text))
            sqlComm.Parameters.AddWithValue("Val2", Trim(TextBoxTod.Text))
            sqlComm.Parameters.AddWithValue("Val3", Trim(TextBoxDown.Text))
            sqlComm.Parameters.AddWithValue("LotDate", Now.ToString("yyyyMMddhhmmss"))
            sqlComm.Parameters.AddWithValue("LoginID", "1")
            '--MessageBox.Show(st1 & "," & st2 & "," & st3)
            sqlCon.Open()
            sqlComm.ExecuteNonQuery()
            sqlCon.Close()
            Call ClearText()
            Me.Hide()
            Form_Input.Button3Refresh.PerformClick()
        End Using
        Exit Sub      ' Exit to avoid handler.
ErrorHandler:  ' Error-handling routine.
        MessageBox.Show(Err.Number & " : " & Err.Description)
        Form_Input.Button3Refresh.PerformClick()
        Resume Next  ' Resume execution at the statement immediately 
    End Sub

    Private Sub ClearText()
        On Error Resume Next
        Me.TextBoxNumber.Text = "000"
        Me.TextBoxTop.Text = "0"
        Me.TextBoxTod.Text = "0"
        Me.TextBoxDown.Text = "0"
    End Sub

    Private Function Left(p1 As String, p2 As Integer) As Object
        Throw New NotImplementedException
    End Function

    Private Sub Form_6x_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        Call ClearText()
        TextBoxNumber.Focus()
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

    Private Sub TextBoxTop_QueryAccessibilityHelp(sender As Object, e As QueryAccessibilityHelpEventArgs) Handles TextBoxTop.QueryAccessibilityHelp
        Me.TextBoxTop.SelectAll()
    End Sub

    Private Sub TextBoxTop_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTop.TextChanged

    End Sub

    Private Sub TextBoxTod_GotFocus(sender As Object, e As EventArgs) Handles TextBoxTod.GotFocus
        Me.TextBoxTod.SelectAll()
    End Sub

    Private Sub TextBoxTod_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxTod.KeyDown
        On Error Resume Next
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxTod_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxTod.KeyPress
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

    Private Sub TextBoxTod_TextChanged(sender As Object, e As EventArgs) Handles TextBoxTod.TextChanged

    End Sub

    Private Sub TextBoxDown_GotFocus(sender As Object, e As EventArgs) Handles TextBoxDown.GotFocus
        Me.TextBoxDown.SelectAll()
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

    Private Sub TextBoxDown_TextChanged(sender As Object, e As EventArgs) Handles TextBoxDown.TextChanged

    End Sub

    Private Sub TextBoxNumber_GotFocus(sender As Object, e As EventArgs) Handles TextBoxNumber.GotFocus
        Me.TextBoxNumber.SelectAll()
    End Sub

    Private Sub TextBoxNumber_TextChanged(sender As Object, e As EventArgs) Handles TextBoxNumber.TextChanged

    End Sub

    Private Sub Button3Refresh_Click()
        Throw New NotImplementedException
    End Sub

End Class
