Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Public Class FormConfig
    Private strConn As String = "Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234"
    Private sqlCon As SqlConnection
    Private Sub LoadData()
        On Error Resume Next
        Dim strQuery As String
        strQuery = "SELECT ConfigID, ConfigTop2Digit, ConfigDown2Digit, ConfigDiv2Digit, ConfigTop3Digit, ConfigDown3Digit, ConfigDiv3Digit, ConfigDate, LoginID FROM TblConfig WHERE ConfigID =1"
        sqlCon = New SqlConnection(strConn)
        Using (sqlCon)
            Dim sqlComm As SqlCommand = New SqlCommand(strQuery, sqlCon)
            sqlCon.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()
            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    Me.TextBoxPercent2Top.Text = sqlReader.GetDecimal(1).ToString  '--ConfigTop2Digit
                    Me.TextBoxPercent2Down.Text = sqlReader.GetDecimal(2).ToString '--ConfigDown2Digit
                    Me.TextBox2Div.Text = sqlReader.GetDecimal(3).ToString '--ConfigDiv2Digit
                    '----------------------------------------------------------
                    Me.TextBoxPercent3Top.Text = sqlReader.GetDecimal(4).ToString '--ConfigTop3Digit
                    Me.TextBoxPercent3Down.Text = sqlReader.GetDecimal(5).ToString ' --ConfigDown3Digit
                    Me.TextBox3Div.Text = sqlReader.GetDecimal(6).ToString '--ConfigDiv3Digit
                End While
            End If
            sqlReader.Close()
        End Using
    End Sub
    Private Sub UpdateRecord()
        '==========================================
        'use [LotDB]
        'CREATE PROCEDURE UpdateConfig
        '@ConfigTop2Digit decimal,
        '@ConfigDown2Digit decimal,
        '@ConfigDiv2Digit decimal,
        '@ConfigTop3Digit decimal,
        '@ConfigDown3Digit decimal,
        '@ConfigDiv3Digit decimal,
        '@LoginId bigint
        'AS
        '        BEGIN()
        '        Update([LotDB].[dbo].[TblConfig])

        'SET ConfigTop2Digit  = @ConfigTop2Digit,
        'ConfigDown2Digit = @ConfigDown2Digit,'
        'ConfigDiv2Digit  = @ConfigDiv2Digit,
        'ConfigTop3Digit  = @ConfigTop3Digit,
        'ConfigDown3Digit = @ConfigDown3Digit,
        'ConfigDiv3Digit  = @ConfigDiv3Digit,
        'ConfigDate       = GETDATE(),
        'LoginID = @LoginId
        '
        'WHERE ConfigID = 1
        '
        'End
        '==========================================
        sqlCon = New SqlConnection(strConn)

        Using (sqlCon)

            Dim sqlComm As New SqlCommand

            sqlComm.Connection = sqlCon


            sqlComm.CommandText = "UpdateConfig"
            sqlComm.CommandType = CommandType.StoredProcedure

            sqlComm.Parameters.AddWithValue("ConfigTop2Digit", Val(Me.TextBoxPercent2Top.Text))
            sqlComm.Parameters.AddWithValue("ConfigDown2Digit", Val(Me.TextBoxPercent2Down.Text))
            sqlComm.Parameters.AddWithValue("ConfigDiv2Digit", Val(Me.TextBox2Div.Text))
            '===================================================================================
            sqlComm.Parameters.AddWithValue("ConfigTop3Digit", Val(Me.TextBoxPercent3Top.Text))
            sqlComm.Parameters.AddWithValue("ConfigDown3Digit", Val(Me.TextBoxPercent3Down.Text))
            sqlComm.Parameters.AddWithValue("ConfigDiv3Digit", Val(Me.TextBox3Div.Text))
            sqlComm.Parameters.AddWithValue("LoginId", "1")
            sqlCon.Open()
            sqlComm.ExecuteNonQuery()
        End Using
        Call LoadData()
        sqlCon.Close()
    End Sub
    Private Sub ButtonSaveConfig_Click(sender As Object, e As EventArgs) Handles ButtonSaveConfig.Click
        On Error Resume Next
        '------------ Validate -----------------
        If Me.TextBoxPercent2Top.Text <> "" And Me.TextBoxPercent2Down.Text <> "" And Me.TextBox2Div.Text <> "" and _
         Me.TextBoxPercent3Top.Text <> "" And Me.TextBoxPercent3Down.Text <> "" And Me.TextBox3Div.Text <> "" Then
            Call UpdateRecord()
            Me.TopMost = False
            Me.Close()
        Else
            MessageBox.Show("กรุณาใส่ข้อมูลให้ถูกต้อง")
            Me.TextBoxPercent2Top.Focus()
        End If
    End Sub

    Private Sub TextBoxPercent2Top_GotFocus(sender As Object, e As EventArgs) Handles TextBoxPercent2Top.GotFocus
        On Error Resume Next
        TextBoxPercent2Top.SelectAll()
    End Sub

    Private Sub TextBoxPercent2Top_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxPercent2Top.KeyDown
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
        If e.KeyCode = 9 Then '-- Key Enter
            Me.ProcessTabKey(+1)
        ElseIf e.KeyCode = 13 Then
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxPercent2Top_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxPercent2Top.KeyPress
        On Error Resume Next
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBoxPercent2Top_TextChanged(sender As Object, e As EventArgs) Handles TextBoxPercent2Top.TextChanged

    End Sub

    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs) Handles ButtonCancel.Click
        Me.TopMost = False
        Me.Close()
    End Sub

    Private Sub FormConfig_Load(sender As Object, e As EventArgs) Handles Me.Load
        On Error Resume Next
        Me.TopMost = True
        Call LoadData()
        Me.TextBox2Div.Focus()
    End Sub

    Private Sub TextBoxPercent2Down_ClientSizeChanged(sender As Object, e As EventArgs) Handles TextBoxPercent2Down.ClientSizeChanged

    End Sub

    Private Sub TextBox2Div_GotFocus(sender As Object, e As EventArgs) Handles TextBox2Div.GotFocus
        On Error Resume Next
        TextBox2Div.SelectAll()
    End Sub

    Private Sub TextBox2Div_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2Div.KeyDown
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
        If e.KeyCode = 9 Then '-- Key Enter
            Me.ProcessTabKey(+1)
        ElseIf e.KeyCode = 13 Then
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBox2Div_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2Div.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBoxPercent2Down_GotFocus(sender As Object, e As EventArgs) Handles TextBoxPercent2Down.GotFocus
        On Error Resume Next
        TextBoxPercent2Down.SelectAll()
    End Sub

    Private Sub TextBoxPercent2Down_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxPercent2Down.KeyDown
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
        If e.KeyCode = 9 Then '-- Key Enter
            Me.ProcessTabKey(+1)
        ElseIf e.KeyCode = 13 Then
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxPercent2Down_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxPercent2Down.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBoxPercent2Down_TextChanged(sender As Object, e As EventArgs) Handles TextBoxPercent2Down.TextChanged

    End Sub

    Private Sub TextBox2Div_TextChanged(sender As Object, e As EventArgs) Handles TextBox2Div.TextChanged

    End Sub

    Private Sub TextBoxPercent3Top_GotFocus(sender As Object, e As EventArgs) Handles TextBoxPercent3Top.GotFocus
        On Error Resume Next
        TextBoxPercent3Top.SelectAll()
    End Sub

    Private Sub TextBoxPercent3Top_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxPercent3Top.KeyDown
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
        If e.KeyCode = 9 Then '-- Key Enter
            Me.ProcessTabKey(+1)
        ElseIf e.KeyCode = 13 Then
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxPercent3Top_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxPercent3Top.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBoxPercent3Top_TextChanged(sender As Object, e As EventArgs) Handles TextBoxPercent3Top.TextChanged

    End Sub

    Private Sub TextBoxPercent3Down_GotFocus(sender As Object, e As EventArgs) Handles TextBoxPercent3Down.GotFocus
        On Error Resume Next
        TextBoxPercent3Down.SelectAll()
    End Sub

    Private Sub TextBoxPercent3Down_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBoxPercent3Down.KeyDown
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
        If e.KeyCode = 9 Then '-- Key Enter
            Me.ProcessTabKey(+1)
        ElseIf e.KeyCode = 13 Then
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBoxPercent3Down_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBoxPercent3Down.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBoxPercent3Down_TextChanged(sender As Object, e As EventArgs) Handles TextBoxPercent3Down.TextChanged

    End Sub

    Private Sub TextBox3Div_GotFocus(sender As Object, e As EventArgs) Handles TextBox3Div.GotFocus
        On Error Resume Next
        TextBox3Div.SelectAll()
    End Sub

    Private Sub TextBox3Div_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3Div.KeyDown
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
        If e.KeyCode = 9 Then '-- Key Enter
            Me.ProcessTabKey(+1)
        ElseIf e.KeyCode = 13 Then
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBox3Div_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3Div.KeyPress
        Select Case Asc(e.KeyChar)
            Case 48 To 57 ' key โค๊ด ของตัวเลขจะอยู่ระหว่าง48-57ครับ 48คือเลข0 57คือเลข9ตามลำดับ
                e.Handled = False
            Case 8, 13, 46 ' ปุ่ม Backspace = 8,ปุ่ม Enter = 13, ปุ่มDelete = 46
                e.Handled = False

            Case Else
                e.Handled = True
                MessageBox.Show("สามารถกดได้แค่ตัวเลข")
        End Select
    End Sub

    Private Sub TextBox3Div_TextChanged(sender As Object, e As EventArgs) Handles TextBox3Div.TextChanged

    End Sub
End Class