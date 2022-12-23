Imports System.Data.SqlClient
Imports System.IO
Imports System.Text
Imports System.IO.TextReader
Imports System.IO.TextWriter
Imports System.Runtime.CompilerServices

Public Class Form_Input
    'Dim connection As New SqlConnection("Server=.; Database=LotDB; Integrated Security = true")
    Dim StringSQL As String = "SELECT TOP 30 LOTTRN as รายการที่ , LOTNUMBER as หมายเลข2หลัก , LOTVALUES1 as จำนวนเงิน_บน , LOTVALUES2 as จำนวนเงิน_ล่าง ,LOTTIME as วันและเวลา FROM TblLOTDetail WHERE LOTID= " & Now.ToString("yyyyMMdd") & " Order By LOTTRN DESC"
    Dim StringSQL2 As String = "SELECT TOP 30 LOTTRN as รายการที่,LOTNUMBER as หมายเลข3หลัก, LOTVALUES1 as จำนวนเงิน_บน , LOTVALUES2 as จำนวนเงิน_โต้ด ,LOTVALUES3 as จำนวนเงิน_ล่าง ,LOTTIME as วันและเวลา FROM TblLOT3Detail WHERE LOTID= " & Now.ToString("yyyyMMdd") & " Order By LOTTRN DESC"
    Dim connection As New SqlConnection("Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234")
    'Dim connection As New SqlConnection("Data Source=local;Initial Catalog=LotDB;User ID=sa;Password=")
    'Dim connection As New SqlConnection("DSN=lot_number;Uid=lot_number;Pwd=;lotnumber1234")
    Dim Stringfile As String
    Dim Top_Val As Decimal
    Dim Down_Val As Decimal
    Dim Dplus As Boolean
    Dim Percent2Top As Decimal
    Dim Percent2Down As Decimal
    Dim Percent3Top As Decimal
    Dim Percent3Down As Decimal
    Dim Rate2Top As Decimal
    Dim Rate2Down As Decimal
    Dim Text3Number As String = "000"
    Dim Text3Amount As String = "0"
    Private DataLotNumber As DataTable
    Private strConn As String = "Data Source=185.78.164.192;Initial Catalog=LotDB;Persist Security Info=True;User ID=lot_number;Password=lotnumber1234"
    Private sqlCon As SqlConnection

    Private Sub LoadConfig()
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
                    Percent2Top = sqlReader.GetDecimal(1).ToString  '--ConfigTop2Digit
                    Percent2Down = sqlReader.GetDecimal(2).ToString '--ConfigDown2Digit
                    'Me.TextBox2Div.Text = sqlReader.GetDecimal(3).ToString '--ConfigDiv2Digit
                    '----------------------------------------------------------
                    Percent3Top = sqlReader.GetDecimal(4).ToString '--ConfigTop3Digit
                    Percent3Down = sqlReader.GetDecimal(5).ToString ' --ConfigDown3Digit
                    'Me.TextBox3Div.Text = sqlReader.GetDecimal(6).ToString '--ConfigDiv3Digit
                End While
            End If
            sqlReader.Close()
        End Using
        sqlCon.Close()
    End Sub
    Public Sub Insert3DB()
        On Error Resume Next
        Dim searchString As String
        Dim searchStar As String = "*"
        Dim searchPlus As String = "+"
        Dim Pos1 As Integer
        Dim Pos2 As Integer
        Dim LotNumber As String
        Dim LotFlag As String
        Dim Amount1 As String = "0" '-- บน
        Dim Amount2 As String = "0" '-- โต้ด
        Dim Amount3 As String = "0" '-- ล่าง
        Dim insertQuery As String
        If Len(Me.TextBox3_Number.Text) <> 3 Then
            MessageBox.Show("Please insert 3 Number again")
            Call Clear3Text()
            Me.TextBox3_Number.SelectAll()
            Me.TextBox3_Number.Focus()
        ElseIf Len(Me.TextBox3_Amount.Text) = 0 Then
            MessageBox.Show("Please insert data")
            Call Clear3Text()
            Me.TextBox3_Number.SelectAll()
            Me.TextBox3_Number.Focus()
        ElseIf Len(Me.Text3Tod.Text) = 0 Then
            MessageBox.Show("Please insert data")
            Call Clear3Text()
            Me.TextBox3_Number.SelectAll()
            Me.TextBox3_Number.Focus()
        ElseIf Len(Me.Text3Down.Text) = 0 Then
            MessageBox.Show("Please insert data")
            Call Clear3Text()
            Me.TextBox3_Number.SelectAll()
            Me.TextBox3_Number.Focus()
        Else
            LotNumber = Trim(TextBox3_Number.Text)
            Amount1 = Trim(TextBox3_Amount.Text) '-- บน
            Amount2 = Trim(Text3Tod.Text) '-- โต้ด
            Amount3 = Trim(Text3Down.Text) '-- ล่าง
            LotFlag = "1"
            '-------------- Prepare for insert data into DATABASE -------------------
            insertQuery = "INSERT INTO TblLOT3Detail (LOTID,LOTNUMBER,LOTVALUES1,LOTVALUES2,LOTVALUES3,LOTFLAG,LOTTIME,LOTSTATUS,LOTDATE,LoginID) VALUES("
            insertQuery = insertQuery & Now.ToString("yyyyMMdd") '-- LOTID 20160821
            insertQuery = insertQuery & ",'" & LotNumber & "'" '-- LOTNUMBER
            insertQuery = insertQuery & "," & Amount1 '-- บน
            insertQuery = insertQuery & "," & Amount2 '-- โต้ด
            insertQuery = insertQuery & "," & Amount3 '-- ล่าง
            insertQuery = insertQuery & ",'" & LotFlag & "'"   '-- LOTFLAG 1
            insertQuery = insertQuery & ", Getdate()" '-- LOTTIME
            insertQuery = insertQuery & ",'1'" '-- LOTSTATUS
            insertQuery = insertQuery & "," & Now.ToString("yyyyMMddhhmmss")  '-- LOTDATE
            insertQuery = insertQuery & "," & Me.LabelLoginID.Text & ")" '-- LOTSTATUS
            '--MessageBox.Show(insertQuery)
            ExecuteQuery(insertQuery)
            '------- Copy text -----------
            'Text3Number = Trim(Me.TextBox3_Number.Text)
            'Text3Amount = Trim(Me.TextBox3_Amount.Text)
            '-------- Clear text ---------
            Call Clear3Text()
        End If
    End Sub
    Public Sub InsertDB()
        On Error Resume Next
        Dim insertQuery As String
        '----------- Copy Value -------------
        Me.LabelTop.Text = Me.TextBox_Top.Text
        Me.LabelDown.Text = Me.TextBox_Down.Text
        '-----------------------------------
        If Len(TextBox_Number.Text) <> 2 Then
            MessageBox.Show("Please insert 2Number again")
            Me.TextBox_Number.Text = "0"
            Me.TextBox_Number.SelectAll()
            Me.TextBox_Number.Focus()
        ElseIf Len(TextBox_Top.Text) = 0 Then
            MessageBox.Show("Please insert data")
            Me.TextBox_Top.Text = "0"
            Me.TextBox_Top.SelectAll()
            Me.TextBox_Top.Focus()
        ElseIf Len(TextBox_Down.Text) = 0 Then
            MessageBox.Show("Please insert data")
            Me.TextBox_Down.Text = "0"
            Me.TextBox_Down.SelectAll()
            Me.TextBox_Down.Focus()
        Else
            '-------------- Prepare for insert data into DATABASE -------------------
            insertQuery = "INSERT INTO TblLOTDetail (LOTID, LOTNUMBER, LOTVALUES1, LOTVALUES2, LOTFLAG, LOTTIME, LOTSTATUS, LOTDATE,LoginID) VALUES("
            insertQuery = insertQuery & Now.ToString("yyyyMMdd") '-- LOTID 20160821
            insertQuery = insertQuery & ",'" & Trim(TextBox_Number.Text) & "'" '-- LOTNUMBER
            insertQuery = insertQuery & "," & Val(TextBox_Top.Text) '-- LOTVALUES1
            insertQuery = insertQuery & "," & Val(TextBox_Down.Text) '-- LOTVALUES2
            insertQuery = insertQuery & ",'1'" '-- LOTFLAG
            insertQuery = insertQuery & ", Getdate()" '-- LOTTIME
            insertQuery = insertQuery & ",'1'" '-- LOTSTATUS
            insertQuery = insertQuery & "," & Now.ToString("yyyyMMddhhmmss")  '-- LOTDATE
            insertQuery = insertQuery & "," & Me.LabelLoginID.Text & ")" '-- LOTSTATUS
            'MessageBox.Show(insertQuery)
            ExecuteQuery(insertQuery)
            Call LoadData(StringSQL)
            Call CopyText()
            Call ClearText()
            Call DisplayPlus(False)
            Call DisplayTotalline()
        End If
    End Sub
    Public Sub DisplayTotal()
        On Error Resume Next
        Dim strQuery As String
        strQuery = "SELECT sum(LOTVALUES1) as A , sum(LOTVALUES2) as B FROM [LotDB].[dbo].[TblLOTDetail] WHERE LOTID = " & Now.ToString("yyyyMMdd") '-- LOTID 20160821
        sqlCon = New SqlConnection(strConn)
        Using (sqlCon)
            Dim sqlComm As SqlCommand = New SqlCommand(strQuery, sqlCon)
            sqlCon.Open()
            Dim sqlReader As SqlDataReader = sqlComm.ExecuteReader()
            If sqlReader.HasRows Then
                While (sqlReader.Read())
                    TextBoxTotal_Top.Text = sqlReader.GetDecimal(0).ToString  '--sum(LOTVALUES1)
                    TextBoxTotal_Down.Text = sqlReader.GetDecimal(1).ToString '--sum(LOTVALUES2)
                End While
            End If
            sqlReader.Close()
        End Using
        sqlCon.Close()
    End Sub
    Public Sub DisplayTotalline()
        On Error Resume Next
        Dim connection As New SqlConnection(strConn)
        Dim command As New SqlCommand("SELECT sum(LOTVALUES1) as A , sum(LOTVALUES2) as B FROM [LotDB].[dbo].[TblLOTDetail] WHERE LOTID = " & Now.ToString("yyyyMMdd"), connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        Dim FirstRow As DataRow = table.Rows(0)
        If table.Rows.Count() <= 0 Then
            TextBoxTotal_Top.Text = "0.00" '--sum(LOTVALUES1)
            TextBoxTotal_Down.Text = "0.00" '--sum(LOTVALUES2)
        Else
            TextBoxTotal_Top.Text = FirstRow("A") '--sum(LOTVALUES1)
            TextBoxTotal_Down.Text = FirstRow("B") '--sum(LOTVALUES2)
            connection.Close()
        End If
    End Sub
    Public Sub Display3Totalline()
        On Error Resume Next
        Dim connection As New SqlConnection(strConn)
        Dim command As New SqlCommand("SELECT sum(LOTVALUES1) as A , sum(LOTVALUES2) as B , sum(LOTVALUES3) as C FROM [LotDB].[dbo].[TblLOT3Detail] WHERE LOTID = " & Now.ToString("yyyyMMdd"), connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        Dim FirstRow As DataRow = table.Rows(0)
        If table.Rows.Count() <= 0 Then
            TextBox3Top.Text = "0.00" '--sum(LOTVALUES1) บน
            TextBox3Tod.Text = "0.00" '--sum(LOTVALUES2) โต้ด
            TextBox3Down.Text = "0.00" '--sum(LOTVALUES3)  ล่าง
        Else
            TextBox3Top.Text = FirstRow("A") '--sum(LOTVALUES1) บน
            TextBox3Tod.Text = FirstRow("B") '--sum(LOTVALUES2) โต้ด
            TextBox3Down.Text = FirstRow("C") '--sum(LOTVALUES3)  ล่าง
            connection.Close()
        End If
    End Sub
    Public Sub DisplayPlus(Plus As Boolean)
        On Error Resume Next
        If Plus = True Then
            Me.LabelPlus.ForeColor = Color.Red
            Dplus = True
        Else
            Me.LabelPlus.ForeColor = Color.Black
            Dplus = False
        End If
    End Sub
    Public Sub ExecuteQuery(query As String)
        Dim Command As New SqlCommand(query, connection)
        connection.Open()
        Command.ExecuteNonQuery()
        connection.Close()
    End Sub
    Private Sub Form_Input_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Timer_0.Enabled = True
        GroupBoxReport.Visible = False
        Call ClearText()
        Call Clear3Text()
        Call LoadData(StringSQL)
        Call DisplayTotalline()
        Call Load3Data(StringSQL2)
        Call Display3Totalline()
        Me.TextBox_Number.Focus()
    End Sub
    Private Sub Timer_0_Tick(sender As Object, e As EventArgs) Handles Timer_0.Tick
        Label_time.Text = Date.Now.ToString("dd-MMM-yyyy hh:mm:ss")
    End Sub
    Public Sub Clear3Text()
        On Error Resume Next
        Me.TextBox3_Number.Text = "000"
        TextBox3_Amount.Text = "0"
        Text3Tod.Text = "0"
        Text3Down.Text = "0"
        Me.Label3Remark.Text = "Tips:"
    End Sub
    Private Sub ClearText()
        Dim CurrentDate As DateTime = DateTime.Today
        Dim Yesterday As DateTime = DateTime.Today.AddDays(-1)
        Me.TextBox_Number.Text = "0"
        Me.TextBox_Top.Text = "0"
        Me.TextBox_Down.Text = "0"
        Me.TextBox_Number.Focus()
        Me.GroupBoxReport.Visible = False
        'Me.MaskedTextBoxStart.Text = Now.ToString("yyyyMMdd")
        Me.MaskedTextBoxStart.Text = Yesterday.ToString("yyyyMMdd")
        Me.MaskedTextBoxStop.Text = CurrentDate.ToString("yyyyMMdd")
        Dplus = False
    End Sub
    Private Sub CopyText()
        On Error Resume Next
        If Dplus = True Then
            Me.TextBox_Top.Text = Val(LabelTop.Text)
            Me.TextBox_Down.Text = Val(LabelDown.Text)
        Else
            Me.TextBox_Top.Text = Val(LabelTop.Text)
            Me.TextBox_Down.Text = Val(LabelDown.Text)
        End If
    End Sub
    Public Sub LoadData(SQLCommand As String)
        Dim command As New SqlCommand(SQLCommand, connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        DataGridView1.DataSource = table
    End Sub
    Public Sub Load3Data(SQLCommand As String)
        Dim command As New SqlCommand(SQLCommand, connection)
        Dim adapter As New SqlDataAdapter(command)
        Dim table As New DataTable()
        adapter.Fill(table)
        DataGridView2.DataSource = table
    End Sub
    Private Sub TextBox_Number_Focus()
        Me.TextBox_Number.BackColor = Color.SkyBlue
    End Sub

    Private Sub GroupBox1_Enter(sender As Object, e As EventArgs) Handles GroupBox1.Enter

    End Sub

    Private Sub Button_Insert_Click(sender As Object, e As EventArgs) Handles Button_Insert.Click
        Call InsertDB()
        Me.TextBox_Number.Focus()
    End Sub
    Private Sub ButtonRefresh_Click(sender As Object, e As EventArgs) Handles ButtonRefresh.Click
        Call ClearText()
        Call LoadData(StringSQL)
        Call DisplayTotalline()
        Me.TextBox_Number.Focus()
    End Sub

    Private Sub ButtonReport_Click(sender As Object, e As EventArgs) Handles ButtonReport.Click
        On Error Resume Next
        If Me.GroupBoxReport.Visible = False Then
            '----------------------------------------
            Me.LabelFlag.Text = "เลือกช่วงวันที่ต้องการดึงรายงาน"
            Me.ButtonExportData.Visible = True
            Me.ButtonDelete.Visible = False
            '-----------------------------------------
            Call ClearText()
            Call LoadConfig()
            Me.GroupBoxReport.Visible = True
            Me.MaskedTextBoxStart.Focus()
        Else
            Me.GroupBoxReport.Visible = False
        End If
    End Sub
    Private Sub ButtonCancel_Click(sender As Object, e As EventArgs)
        Me.GroupBoxReport.Visible = False
        Me.TextBox_Number.Focus()
    End Sub

    Private Sub ButtonClose_Click(sender As Object, e As EventArgs) Handles ButtonClose.Click
        Me.GroupBoxReport.Visible = False
        Me.ButtonReport.Enabled = True
    End Sub

    Private Sub ButtonExportData_Click(sender As Object, e As EventArgs) Handles ButtonExportData.Click
        On Error Resume Next
        'For Text  file this is ok
        '== check blank 
        If Len(Me.MaskedTextBoxStart.Text) <> 8 Or Len(MaskedTextBoxStop.Text) <> 8 Then
            MessageBox.Show("Please insert date in =>>yyymmdd")
            Me.MaskedTextBoxStart.Text = Now.ToString("yyyyMMdd")
            Me.MaskedTextBoxStop.Text = Now.ToString("yyyyMMdd")
            Me.MaskedTextBoxStart.Focus()
        Else
            Dim SQLCommand As String
            SQLCommand = "SELECT LOTNUMBER as NUMBER , sum(LOTVALUES1) as TOP_VALUE ,sum(LOTVALUES2) as DOWN_VALUE  FROM TblLOTDetail"
            SQLCommand = SQLCommand & " WHERE LOTSTATUS =1"
            SQLCommand = SQLCommand & " AND LOTID BETWEEN " & Trim(MaskedTextBoxStart.Text) & " AND " & Trim(MaskedTextBoxStop.Text)
            SQLCommand = SQLCommand & " group by LOTNUMBER"
            SQLCommand = SQLCommand & " order by LOTNUMBER ASC"
            'MessageBox.Show(SQLCommand)
            Top_Val = 0
            Down_Val = 0
            Rate2Top = 0
            Rate2Down = 0
            Call ExportCSV(SQLCommand) '-------------- Way#1 
            MessageBox.Show("Export data finish at " & Stringfile)
            SQLCommand = "SELECT LOTNUMBER as NUMBER , sum(LOTVALUES1) - " & Rate2Top & " as TOP_VALUE "
            SQLCommand = SQLCommand & ",sum(LOTVALUES2) - " & Rate2Down & " as DOWN_VALUE  FROM TblLOTDetail "
            SQLCommand = SQLCommand & " WHERE LOTSTATUS =1"
            SQLCommand = SQLCommand & " AND LOTID BETWEEN " & Trim(MaskedTextBoxStart.Text) & " AND " & Trim(MaskedTextBoxStop.Text)
            SQLCommand = SQLCommand & " group by LOTNUMBER"
            SQLCommand = SQLCommand & " order by LOTNUMBER ASC"
            Call ExportSQL(SQLCommand)  '-------------- Way#2
            MessageBox.Show("Export data finish at " & Stringfile)
            Call LoadData(StringSQL)
            '------------------ Display File Name ----------------------
            'MessageBox.Show("Export data finish at " & Stringfile)
            'Process.Start("explorer.exe", "/select," & Stringfile) '-- Not Open Window Explorer
        End If
    End Sub
    Private Sub ExportCSV(sqlcommand As String) '---------------- Way#1
        On Error Resume Next
        Call LoadData(sqlcommand)
        Dim filePath As String
        filePath = System.IO.Path.Combine(
        My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOT(2M)-" & Now.ToString("yyyy-MM-dd") & ".csv")
        Dim writer As TextWriter = New StreamWriter(filePath)
        writer.WriteLine("Export file date: " & Now.ToString("yyyy-MM-dd") & "and time: " & Now.ToString("hh:mm:ss"))
        writer.WriteLine("NUMBER,TOP-VALUE,DOWN-VALUE,")
        For i As Integer = 0 To DataGridView1.Rows.Count - 1 Step +1

            For j As Integer = 0 To DataGridView1.Columns.Count - 1 Step +1

                writer.Write("'" & DataGridView1.Rows(i).Cells(j).Value.ToString() & ",")

            Next
            writer.WriteLine("")
            Top_Val = Top_Val + Val(DataGridView1.Rows(i).Cells(1).Value.ToString)
            Down_Val = Down_Val + Val(DataGridView1.Rows(i).Cells(2).Value.ToString)
            'MessageBox.Show(" Top Value is :: " & Top_Val & " and Down Value is :: " & Down_Val)
        Next
        writer.WriteLine("SUM_TOTAL,'" & Top_Val & ".0000," & "'" & Down_Val & ".0000,")
        Rate2Top = Top_Val - (Top_Val * (Percent2Top / 100))
        Rate2Down = Down_Val - (Down_Val * (Percent2Down / 100))
        writer.WriteLine("SUM_DISCOUNT_PERCENT," & "DEDUCT(" & (Percent2Top / 100) & "%)=" & Rate2Top & ",DEDUCT(" & (Percent2Down / 100) & "%)=" & Rate2Down & ",")
        writer.Close()
        'MessageBox.Show("Data Export file name:" & filePath)
        Stringfile = filePath

    End Sub
    Private Sub ExportSQL(sqlcommand As String) '------------------ Way#2
        'For Text  file this is ok
        Dim Table As New DataTable
        Dim fileloc As String
        Dim Adapter As New SqlDataAdapter(sqlcommand, connection)
        Adapter.Fill(Table)
        Dim txt As String
        fileloc = System.IO.Path.Combine(
        My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOT(2N)-" & Now.ToString("yyyy-MM-dd") & ".csv")
        '======================= Loop to write data records from SQL Command ===============================  
        Dim line As String = "NUMBER,TOP-VALUE,DOWN-VALUE,"
        txt = txt + line & vbCrLf
        For Each row As DataRow In Table.Rows
            line = ""
            For Each column As DataColumn In Table.Columns
                'Add the Data rows.
                line = line + "," & row(column.ColumnName).ToString()
                'line += vbTab & row(column.ColumnName).ToString()
                'line = line + "," + row(column.ColumnName(0)).ToString()
            Next
            'Add new line
            txt = txt + line.Substring(1) & vbCrLf
        Next

        'If File.Exists(fileloc) Then
        Using sw As StreamWriter = New StreamWriter(fileloc)
            sw.WriteLine(txt)
            sw.WriteLine("Export file (LOT2N)  Date: " & Now.ToString("yyyy-MM-dd") & "/ time is: " & Now.ToString("hh:mm:ss"))
        End Using
        'End If
        Stringfile = fileloc
    End Sub

    Private Sub Button3Export_Click(sender As Object, e As EventArgs) Handles Button3Export.Click
        On Error Resume Next
        If Len(Me.MaskedText3Start.Text) <> 8 Or Len(Me.MaskedText3Stop.Text) <> 8 Then
            MessageBox.Show("Please insert date in =>>yyymmdd")
            Me.MaskedText3Start.Text = Now.ToString("yyyyMMdd")
            Me.MaskedText3Stop.Text = Now.ToString("yyyyMMdd")
            Me.MaskedText3Start.Focus()
        Else
            Dim SQLCommand As String
            SQLCommand = "SELECT LOTNUMBER as NUMBER , sum(LOTVALUES1) as TOP_VALUE "
            SQLCommand = SQLCommand & ",sum(LOTVALUES2)  as TOD_VALUE "
            SQLCommand = SQLCommand & ",sum(LOTVALUES3)  as DOWN_VALUE  FROM TblLOT3Detail "
            SQLCommand = SQLCommand & " WHERE LOTSTATUS =1"
            SQLCommand = SQLCommand & " AND LOTID BETWEEN " & Trim(Me.MaskedText3Start.Text) & " AND " & Trim(Me.MaskedText3Stop.Text)
            SQLCommand = SQLCommand & " group by LOTNUMBER"
            SQLCommand = SQLCommand & " order by LOTNUMBER ASC"
            Call Export3SQL(SQLCommand)  '-------------- Way#2
            MessageBox.Show("Export data finish at " & Stringfile)
            Call LoadData(StringSQL2)
        End If
    End Sub
    Private Sub Export3SQL(sqlcommand As String) '-------------------- Way#2 สำหรับ 3 ตัว
        'For Text  file this is ok
        Dim Table As New DataTable
        Dim fileloc As String
        Dim Adapter As New SqlDataAdapter(sqlcommand, connection)
        Dim TotalNumber As Integer = 0
        Dim TotalTop As Integer = 0
        Dim TotalDown As Integer = 0
        Dim TotalTod As Integer = 0
        Adapter.Fill(Table)
        Dim txt As String
        fileloc = System.IO.Path.Combine(
        My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\LOT(3N)-" & Now.ToString("yyyyMMdd-hhmmss") & ".csv")
        '======================= Loop to write data records from SQL Command ===============================  
        Dim line As String = "NUMBER,TOP_VALUE,TOD_VALUE,DOWN_VALUE"
        txt = txt + line & vbCrLf
        For Each row As DataRow In Table.Rows
            line = ""
            For Each column As DataColumn In Table.Columns
                'Add the Data rows.
                line = line + "," & row(column.ColumnName).ToString()
                'MessageBox.Show(row(column.ColumnName).ToString())
                'line += vbTab & row(column.ColumnName).ToString()
                'line = line + "," + row(column.ColumnName(0)).ToString()
                Select column.ColumnName
                    Case "NUMBER" : TotalNumber = TotalNumber + 1
                    Case "TOP_VALUE" : TotalTop = TotalTop + Val(row(column.ColumnName).ToString())
                    Case "TOD_VALUE" : TotalTod = TotalTod + Val(row(column.ColumnName).ToString())
                    Case "DOWN_VALUE" : TotalDown = TotalDown + Val(row(column.ColumnName).ToString())
                End Select
            Next
            'Add new line
            txt = txt + line.Substring(1) & vbCrLf
        Next
        Using sw As StreamWriter = New StreamWriter(fileloc)
            sw.WriteLine(txt)
            sw.WriteLine("==========,==========,==========,==========")
            sw.WriteLine("Total Count,Total Top,Total Tod,Total Down")
            sw.WriteLine(TotalNumber & "," & TotalTop & "," & TotalTod & "," & TotalDown)  '--TotalTod
            sw.WriteLine("Export file (LOT3N) date: " & Now.ToString("yyyy-MM-dd") & "/ time is: " & Now.ToString("hh:mm:ss"))
        End Using
        Stringfile = fileloc
    End Sub
    Private Sub ExportSQL_Recheck(sqlcommand As String)
        'For Text  file this is ok
        Dim Table As New DataTable
        Dim fileloc As String
        Dim Adapter As New SqlDataAdapter(sqlcommand, connection)
        Adapter.Fill(Table)
        Dim txt As String
        fileloc = System.IO.Path.Combine(
        My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\file_recheck_(2N)_" & Now.ToString("yyyyMMdd-hhmmss") & ").csv")
        '======================= Loop to write data records from SQL Command ===============================  
        Dim line As String = "TRANS_ID,KEY_NUMBER,TOP_VALUE,DOWN_VALUE,LOT_DATE_TIME"
        txt = txt + line & vbCrLf
        For Each row As DataRow In Table.Rows
            line = ""
            For Each column As DataColumn In Table.Columns
                'Add the Data rows.
                line = line + "," & row(column.ColumnName).ToString()
                'line += vbTab & row(column.ColumnName).ToString()
                'line = line + "," + row(column.ColumnName(0)).ToString()
            Next
            'Add new line
            txt = txt + line.Substring(1) & vbCrLf
        Next

        'If File.Exists(fileloc) Then
        Using sw As StreamWriter = New StreamWriter(fileloc)
            sw.WriteLine(txt)
            sw.WriteLine("Recheck file Date: " & Now.ToString("yyyy-MM-dd") & "/ time is: " & Now.ToString("hh:mm:ss"))
        End Using
        'End If
        Stringfile = fileloc

    End Sub

    Public Sub ExportData_by_SQLCommand(file_name As String, header_column As String, sqlcommand As String)
        '=================== This Sub Module for Export data to CSV file by SQL command and Get Colummn Header ===================
        Dim Table As New DataTable
        Dim fileloc As String
        Dim Adapter As New SqlDataAdapter(sqlcommand, connection)
        Adapter.Fill(Table)
        Dim txt As String
        fileloc = System.IO.Path.Combine(
        My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\" & file_name & "-" & Now.ToString("yyyy-MM-dd") & ".csv")
        '======================= Loop to write data records from SQL Command ===============================  
        Dim line As String = header_column
        txt = txt + line & vbCrLf
        For Each row As DataRow In Table.Rows
            line = ""
            For Each column As DataColumn In Table.Columns
                'Add the Data rows.
                line = line + "," & row(column.ColumnName).ToString()
                'line += vbTab & row(column.ColumnName).ToString()
                'line = line + "," + row(column.ColumnName(0)).ToString()
            Next
            'Add new line
            txt = txt + line.Substring(1) & vbCrLf
        Next

        'If File.Exists(fileloc) Then
        Using sw As StreamWriter = New StreamWriter(fileloc)
            sw.WriteLine(txt)
        End Using
        'End If
        Stringfile = fileloc
    End Sub
    Private Sub TextBox_Number_GotFocus(sender As Object, e As EventArgs) Handles TextBox_Number.GotFocus
        On Error Resume Next
        TextBox_Number.SelectAll()
    End Sub

    Private Sub TextBox_Number_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Number.KeyDown
        On Error Resume Next
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Call CopyText()
            Me.ProcessTabKey(+1)
        End If
        If e.KeyCode = Keys.Down Then
            e.SuppressKeyPress = True
            Call CopyText()
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBox_Top_GotFocus(sender As Object, e As EventArgs) Handles TextBox_Top.GotFocus
        On Error Resume Next
        TextBox_Top.SelectAll()
    End Sub


    Private Sub TextBox_Top_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Top.KeyDown
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
            '-------------- Save Data as Click ButtonInsert -----------------------
            e.SuppressKeyPress = True
            Call InsertDB()
            Me.TextBox_Number.Focus()
        ElseIf e.KeyCode = 107 Then '-- Key (+)
            e.SuppressKeyPress = True
            Call DisplayPlus(True)
            Me.TextBox_Down.Focus()
        End If
    End Sub

    Private Sub TextBox_Down_GotFocus(sender As Object, e As EventArgs) Handles TextBox_Down.GotFocus
        On Error Resume Next
        TextBox_Down.SelectAll()
    End Sub

    Private Sub TextBox_Down_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox_Down.KeyDown
        On Error Resume Next
        If e.KeyCode = 13 Then
            '-------------- Save Data as Click ButtonInsert -----------------------
            e.SuppressKeyPress = True
            Call InsertDB()
            Me.TextBox_Number.Focus()
        End If
    End Sub

    Private Sub ButtonDelete_Click(sender As Object, e As EventArgs) Handles ButtonDelete.Click
        On Error Resume Next
        Dim StartDate As String = ""
        Dim StopDate As String = ""
        StartDate = Trim(Me.MaskedTextBoxStart.Text)
        StopDate = Trim(Me.MaskedTextBoxStop.Text)
        If Len(StartDate) <> 8 Or Len(StopDate) <> 8 Then
            MessageBox.Show("Please insert start / stop date is format YYYYMMDD")
            Me.MaskedTextBoxStart.Focus()
        Else
            If MessageBox.Show("Do you want to delete data between date:: " & StartDate & " and " & StopDate & " Yes or No?", _
                                                      "Important Question", _
                                                      MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim insertQuery As String
                insertQuery = "DELETE FROM TblLOTDetail WHERE LOTID BETWEEN " & StartDate
                insertQuery = insertQuery & " AND " & StopDate
                'MessageBox.Show(insertQuery)
                ExecuteQuery(insertQuery)
                MessageBox.Show("Delete data done!!")
                Call LoadData(StringSQL)
                Call DisplayTotalline()
                Me.MaskedTextBoxStart.Focus()
            Else
                MessageBox.Show("Cancel delete data")
                Me.MaskedTextBoxStart.Focus()
            End If
        End If
    End Sub

    Private Sub MaskedTextBoxStart_GotFocus(sender As Object, e As EventArgs) Handles MaskedTextBoxStart.GotFocus
        On Error Resume Next
        Me.MaskedTextBoxStart.SelectAll()
    End Sub

    Private Sub MaskedTextBoxStart_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBoxStart.KeyDown
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
        If e.KeyCode = 13 Or e.KeyCode = 9 Then '-- Key Enter
            '-------------- Save Data as Click ButtonInsert -----------------------
            e.SuppressKeyPress = True
            MaskedTextBoxStop.Focus()
        End If
    End Sub

    Private Sub MaskedTextBoxStart_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MaskedTextBoxStart.KeyPress
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

    Private Sub MaskedTextBoxStart_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles MaskedTextBoxStart.MaskInputRejected

    End Sub

    Private Sub TextBox_Top_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_Top.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' ตรวจสอบการ Return ค่ากลับว่า True หรือ False
        e.Handled = CheckDigitOnly(KeyAscii)
    End Sub

    Private Sub TextBox_Top_LostFocus(sender As Object, e As EventArgs) Handles TextBox_Top.LostFocus

    End Sub

    Private Sub TextBox_Top_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Top.TextChanged
        On Error Resume Next
        If Me.TextBox_Top.Text = "+" Then
            Me.TextBox_Top.Text = "0"
        End If
        If Val(Me.TextBox_Top.Text) <> Val(Me.LabelTop.Text) Then
            Me.TextBox_Down.Text = "0"
        End If
    End Sub

    Private Sub MaskedTextBoxStop_GotFocus(sender As Object, e As EventArgs) Handles MaskedTextBoxStop.GotFocus
        On Error Resume Next
        Me.MaskedTextBoxStop.SelectAll()
    End Sub

    Private Sub MaskedTextBoxStop_KeyDown(sender As Object, e As KeyEventArgs) Handles MaskedTextBoxStop.KeyDown
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
        If e.KeyCode = 13 Or e.KeyCode = 9 Then '-- Key Enter
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub MaskedTextBoxStop_KeyPress(sender As Object, e As KeyPressEventArgs) Handles MaskedTextBoxStop.KeyPress
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

    Private Sub MaskedTextBoxStop_MaskInputRejected(sender As Object, e As MaskInputRejectedEventArgs) Handles MaskedTextBoxStop.MaskInputRejected

    End Sub

    Private Sub TextBox_Number_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_Number.KeyPress
        On Error Resume Next
        If (Not Char.IsControl(e.KeyChar) _
                   AndAlso (Not Char.IsDigit(e.KeyChar) _
                   AndAlso (e.KeyChar <> Microsoft.VisualBasic.ChrW(46)))) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox_Number_RegionChanged(sender As Object, e As EventArgs) Handles TextBox_Number.RegionChanged

    End Sub
    Private Sub ButtonCalculate_Click(sender As Object, e As EventArgs) Handles ButtonCalculate.Click
        On Error Resume Next
        'System.Diagnostics.Process.Start("calc.exe")
        Form_19x.Show()
    End Sub

    Private Sub Label7_Click(sender As Object, e As EventArgs) Handles LabelFlag.Click

    End Sub

    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        On Error Resume Next
        If Me.GroupBoxReport.Visible = False Then
            '----------------------------------------
            Me.LabelFlag.Text = "เลือกช่วงวันที่ต้องการลบข้อมูล"
            Me.ButtonExportData.Visible = False
            Me.ButtonDelete.Visible = True
            '-----------------------------------------
            Call ClearText()
            Me.GroupBoxReport.Visible = True
            Me.MaskedTextBoxStart.Focus()
        Else
            Me.GroupBoxReport.Visible = False
        End If
    End Sub

    Private Sub TextBox_Down_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox_Down.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' ตรวจสอบการ Return ค่ากลับว่า True หรือ False
        e.Handled = CheckDigitOnly(KeyAscii)
    End Sub
    Function CheckDigitOnly(ByVal index As Integer) As Boolean
        '-- How to call function 
        'Private Sub TextBox1_KeyPress(ByVal sender As Object, _
        ' ByVal e As System.Windows.Forms.KeyPressEventArgs _
        ' ) _
        'Handles TextBox1.KeyPress
        'Dim KeyAscii As Short = Asc(e.KeyChar)
        ' ตรวจสอบการ Return ค่ากลับว่า True หรือ False
        'e.Handled = CheckDigitOnly(KeyAscii)
        Select Case index
            Case 43, 45, 46 ' + 43 ,- 45 , . 46
                CheckDigitOnly = False
            Case 48 To 57 ' เลข 0 - 9
                CheckDigitOnly = False
            Case 8, 13 ' Backspace = 8, Enter = 13
                CheckDigitOnly = False
            Case Else
                CheckDigitOnly = True
        End Select
    End Function

    Private Sub TextBox_Down_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles TextBox_Down.PreviewKeyDown

    End Sub
    Private Sub TextBox_Down_TextChanged(sender As Object, e As EventArgs) Handles TextBox_Down.TextChanged
        On Error Resume Next
        If Me.TextBox_Down.Text = "+" Then
            Me.TextBox_Down.Text = Val(Me.LabelDown.Text)
        End If
    End Sub
    Private Sub Button_Exit_Click(sender As Object, e As EventArgs) Handles Button_Exit.Click
        On Error Resume Next
        Dim result = MsgBox("Do you want to exit program?", vbYesNo)
        If result = DialogResult.Yes Then
            Me.Close()
            Application.Exit()
        End If
    End Sub
    Private Sub ButtonCrystal_Click(sender As Object, e As EventArgs) Handles ButtonCrystal.Click
        'Re-check file Log after key-in
        On Error Resume Next
        If Len(Me.MaskedTextBoxStart.Text) <> 8 Or Len(MaskedTextBoxStop.Text) <> 8 Then
            MessageBox.Show("Please insert date in =>>yyymmdd")
            Me.MaskedTextBoxStart.Text = Now.ToString("yyyyMMdd")
            Me.MaskedTextBoxStop.Text = Now.ToString("yyyyMMdd")
            Me.MaskedTextBoxStart.Focus()
        Else
            Dim SQLCommand As String
            SQLCommand = "SELECT [LOTTRN] AS TRANS_ID ,[LOTNUMBER] AS KEY_NUMBER ,[LOTVALUES1] AS VAL_TOP ,[LOTVALUES2] AS VAL_DOWN ,[LOTTIME] AS LOT_DATE_TIME"
            SQLCommand = SQLCommand & " FROM [LotDB].[dbo].[TblLOTDetail]"
            SQLCommand = SQLCommand & " WHERE LOTID BETWEEN " & Trim(MaskedTextBoxStart.Text) & " AND " & Trim(MaskedTextBoxStop.Text)
            SQLCommand = SQLCommand & " ORDER BY [LOTTRN] DESC"
            Call ExportSQL_Recheck(SQLCommand)  '-------------- Way#2
            MessageBox.Show("Export data finish at " & Stringfile)
            Call LoadData(StringSQL)
        End If
    End Sub
    Private Sub Button3Report_Click(sender As Object, e As EventArgs) Handles Button3Report.Click
        On Error Resume Next
        If Me.GroupBox3.Visible = False Then
            '----------------------------------------
            Me.Label3Report.Text = "กรุณาเลือกวันที่ต้องการดึงข้อมูล"
            Me.MaskedText3Start.Text = Now.ToString("yyyyMMdd")
            Me.MaskedText3Stop.Text = Now.ToString("yyyyMMdd")
            Me.Button3Export.Visible = True
            Me.Button3Delete.Visible = False
            '-----------------------------------------
            Call Clear3Text()
            Call LoadConfig()
            Me.GroupBox3.Visible = True
            Me.MaskedText3Start.Focus()
        Else
            Me.GroupBox3.Visible = False
        End If
    End Sub
    Private Sub Label19_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Label20_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub TextBox3_Amount_GotFocus(sender As Object, e As EventArgs) Handles TextBox3_Amount.GotFocus
        On Error Resume Next
        Me.Label3Remark.Text = "Tips:กรุณาใส่ค่าเงิน [บน] และกด [Enter] เพื่อบันทึกค่า หรือใส่ [+] เพื่อใส่ค่าเงิน [โต้ด] ต่อไป>>"
        Me.TextBox3_Amount.SelectAll()
    End Sub

    Private Sub TextBox3_Amount_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3_Amount.KeyDown
        On Error Resume Next
        If e.KeyCode = 13 Then '-- Key Enter
            '-------------- Save Data as Click ButtonInsert -----------------------
            e.SuppressKeyPress = True
            Call Insert3DB()
            Call Load3Data(StringSQL2)
            Call Display3Totalline()
            TextBox3_Number.SelectAll()
            TextBox3_Number.Focus()
        ElseIf e.KeyCode = 107 Then '-- Key (+)
            e.SuppressKeyPress = True
            Text3Tod.SelectAll()
            Text3Tod.Focus()
        End If
    End Sub

    Private Sub TextBox3_Amount_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3_Amount.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' ตรวจสอบการ Return ค่ากลับว่า True หรือ False
        e.Handled = CheckDigitOnly(KeyAscii)
    End Sub

    Private Sub TextBox3_Amount_QueryContinueDrag(sender As Object, e As QueryContinueDragEventArgs) Handles TextBox3_Amount.QueryContinueDrag

    End Sub

    Private Sub TextBox3_top_val_TextChanged(sender As Object, e As EventArgs) Handles TextBox3_Amount.TextChanged

    End Sub

    Private Sub Button3Exit_Click(sender As Object, e As EventArgs) Handles Button3Exit.Click
        On Error Resume Next
        Dim result = MsgBox("Do you want to exit program?", vbYesNo)
        If result = DialogResult.Yes Then
            Me.Close()
            Application.Exit()
        End If
    End Sub

    Private Sub Button3Enter_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button3Clear_Click(sender As Object, e As EventArgs) Handles Button3Clear.Click
        On Error Resume Next
        If Me.GroupBox3.Visible = False Then
            '----------------------------------------
            Me.Label3Report.Text = "เลือกช่วงวันที่ต้องการลบข้อมูล"
            Me.MaskedText3Start.Text = Now.ToString("yyyyMMdd")
            Me.MaskedText3Stop.Text = Now.ToString("yyyyMMdd")
            Me.Button3Export.Visible = False
            Me.Button3Delete.Visible = True
            '-----------------------------------------
            Call Clear3Text()
            Me.GroupBox3.Visible = True
            Me.MaskedText3Start.Focus()
        Else
            Me.GroupBox3.Visible = False
        End If
    End Sub

    Private Sub TextBox3_Number_GotFocus(sender As Object, e As EventArgs) Handles TextBox3_Number.GotFocus
        On Error Resume Next
        Me.Label3Remark.Text = "Tips:กรุณาใส่ตัวเลข 3 ตัวและกด [Enter] เพื่อใส่ค่าเงิน [บน] ต่อไป>>"
        Me.TextBox3_Number.SelectAll()
    End Sub

    Private Sub TextBox3_Number_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3_Number.KeyDown
        On Error Resume Next
        If e.KeyCode = Keys.Enter Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
        If e.KeyCode = Keys.Down Then
            e.SuppressKeyPress = True
            Me.ProcessTabKey(+1)
        End If
    End Sub

    Private Sub TextBox3_Number_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3_Number.KeyPress
        On Error Resume Next
        If (Not Char.IsControl(e.KeyChar) _
                   AndAlso (Not Char.IsDigit(e.KeyChar) _
                   AndAlso (e.KeyChar <> Microsoft.VisualBasic.ChrW(46)))) Then
            e.Handled = True
        End If
    End Sub

    Private Sub TextBox3_Number_TextChanged(sender As Object, e As EventArgs) Handles TextBox3_Number.TextChanged

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
        On Error Resume Next
        Call Clear3Text()
        Me.TextBox3_Number.Focus()
    End Sub

    Private Sub Button3Refresh_Click(sender As Object, e As EventArgs) Handles Button3Refresh.Click
        On Error Resume Next
        Call Clear3Text()
        Call Load3Data(StringSQL2)
        Call Display3Totalline()
        Me.TextBox3_Number.Focus()
    End Sub


    Private Sub Button3Close_Click(sender As Object, e As EventArgs) Handles Button3Close.Click
        Me.GroupBox3.Visible = False

    End Sub

    Private Sub Button3Delete_Click(sender As Object, e As EventArgs) Handles Button3Delete.Click
        On Error Resume Next
        Dim StartDate As String = ""
        Dim StopDate As String = ""
        StartDate = Trim(Me.MaskedText3Start.Text)
        StopDate = Trim(Me.MaskedText3Stop.Text)
        If Len(StartDate) <> 8 Or Len(StopDate) <> 8 Then
            MessageBox.Show("Please insert start / stop date is format YYYYMMDD")
            Me.MaskedTextBoxStart.Focus()
        Else
            If MessageBox.Show("Do you want to delete data between date:: " & StartDate & " and " & StopDate & " Yes or No?", _
                                                      "Important Question", _
                                                      MessageBoxButtons.YesNo) = Windows.Forms.DialogResult.Yes Then
                Dim insertQuery As String
                insertQuery = "DELETE FROM TblLOT3Detail WHERE LOTID BETWEEN " & StartDate
                insertQuery = insertQuery & " AND " & StopDate
                'MessageBox.Show(insertQuery)
                ExecuteQuery(insertQuery)
                MessageBox.Show("Delete data done!!")
                Call Load3Data(StringSQL2)
                Call Display3Totalline()
                Me.MaskedText3Start.Focus()
            Else
                MessageBox.Show("Cancel delete data")
                Me.MaskedText3Start.Focus()
            End If
        End If
    End Sub

    Private Sub Text3Tod_GotFocus(sender As Object, e As EventArgs) Handles Text3Tod.GotFocus
        On Error Resume Next
        Me.Label3Remark.Text = "Tips:กรุณาใส่ค่าเงิน [โต้ด] และกด [Enter] เพื่อบันทึกค่า หรือใส่ [+] เพื่อใส่ค่าเงิน [ล่าง] ต่อไป>>"
        Me.Text3Tod.SelectAll()
    End Sub

    Private Sub Text3Tod_KeyDown(sender As Object, e As KeyEventArgs) Handles Text3Tod.KeyDown
        On Error Resume Next
        If e.KeyCode = 13 Then '-- Key Enter
            '-------------- Save Data as Click ButtonInsert -----------------------
            e.SuppressKeyPress = True
            Call Insert3DB()
            Call Load3Data(StringSQL2)
            Call Display3Totalline()
            TextBox3_Number.SelectAll()
            TextBox3_Number.Focus()
        ElseIf e.KeyCode = 107 Then '-- Key (+)
            e.SuppressKeyPress = True
            Text3Down.Focus()
        End If
    End Sub

    Private Sub Text3Tod_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text3Tod.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' ตรวจสอบการ Return ค่ากลับว่า True หรือ False
        e.Handled = CheckDigitOnly(KeyAscii)
    End Sub

    Private Sub Text3Tod_TextChanged(sender As Object, e As EventArgs) Handles Text3Tod.TextChanged

    End Sub

    Private Sub Text3Down_GotFocus(sender As Object, e As EventArgs) Handles Text3Down.GotFocus
        On Error Resume Next
        Me.Label3Remark.Text = "Tips:กรุณาใส่ค่าเงิน [ล่าง] และกด [Enter] เพื่อบันทึกค่า>>"
        Me.Text3Down.SelectAll()
    End Sub

    Private Sub Text3Down_KeyDown(sender As Object, e As KeyEventArgs) Handles Text3Down.KeyDown
        On Error Resume Next
        If e.KeyCode = 13 Then '-- Key Enter
            '-------------- Save Data as Click ButtonInsert -----------------------
            e.SuppressKeyPress = True
            Call Insert3DB()
            Call Load3Data(StringSQL2)
            Call Display3Totalline()
            TextBox3_Number.SelectAll()
            TextBox3_Number.Focus()
        End If
    End Sub

    Private Sub Text3Down_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Text3Down.KeyPress
        Dim KeyAscii As Short = Asc(e.KeyChar)
        ' ตรวจสอบการ Return ค่ากลับว่า True หรือ False
        e.Handled = CheckDigitOnly(KeyAscii)
    End Sub

    Private Sub Text3Down_TextChanged(sender As Object, e As EventArgs) Handles Text3Down.TextChanged

    End Sub
    Private Sub Export3SQL_Recheck(sqlcommand As String)
        'For Text  file this is ok
        Dim Table As New DataTable
        Dim fileloc As String
        Dim Adapter As New SqlDataAdapter(sqlcommand, connection)
        Adapter.Fill(Table)
        Dim txt As String
        fileloc = System.IO.Path.Combine(
        My.Computer.FileSystem.SpecialDirectories.MyDocuments & "\file_recheck_(3N)_" & Now.ToString("yyyyMMdd-hhmmss") & ").csv")
        '======================= Loop to write data records from SQL Command ===============================  
        Dim line As String = "TRANS_ID,KEY_NUMBER,TOP_VALUE,TOD_VALUE,DOWN_VALUE,LOT_DATE_TIME"
        txt = txt + line & vbCrLf
        For Each row As DataRow In Table.Rows
            line = ""
            For Each column As DataColumn In Table.Columns
                'Add the Data rows.
                line = line + "," & row(column.ColumnName).ToString()
                'line += vbTab & row(column.ColumnName).ToString()
                'line = line + "," + row(column.ColumnName(0)).ToString()
            Next
            'Add new line
            txt = txt + line.Substring(1) & vbCrLf
        Next

        'If File.Exists(fileloc) Then
        Using sw As StreamWriter = New StreamWriter(fileloc)
            sw.WriteLine(txt)
            sw.WriteLine("Recheck file Date: " & Now.ToString("yyyy-MM-dd") & "/ time is: " & Now.ToString("hh:mm:ss"))
        End Using
        'End If
        Stringfile = fileloc
    End Sub
    Private Sub Button3Recheck_Click(sender As Object, e As EventArgs) Handles Button3Recheck.Click
        'Re-check file 3 digit
        On Error Resume Next
        Dim StartDate As String = Trim(Me.MaskedText3Start.Text)
        Dim StopDate As String = Trim(Me.MaskedText3Stop.Text)
        '------------- Check (Start date and Stop date)
        If Len(StartDate) <> 8 Or Len(StopDate) <> 8 Then
            MessageBox.Show("Please insert date in =>>yyymmdd")
            MaskedText3Start.Text = Now.ToString("yyyyMMdd")
            MaskedText3Stop.Text = Now.ToString("yyyyMMdd")
            MaskedText3Start.Focus()
        Else
            Dim SQLCommand As String
            SQLCommand = "SELECT LOTTRN , LOTNUMBER,LOTVALUES1 , LOTVALUES2 , LOTVALUES3 , LOTTIME "
            SQLCommand = SQLCommand & " FROM [LotDB].[dbo].[TblLOT3Detail]"
            SQLCommand = SQLCommand & " WHERE LOTID BETWEEN " & StartDate & " AND " & StopDate
            SQLCommand = SQLCommand & " ORDER BY [LOTTRN] DESC"
            Call Export3SQL_Recheck(SQLCommand)  '-------------- Way#2
            MessageBox.Show("Export data finish at " & Stringfile)
            Call LoadData(StringSQL)
        End If
    End Sub

    Private Sub Button3Tord_Click(sender As Object, e As EventArgs) Handles Button3Tord.Click
        'MessageBox.Show("Commng soon!!")
        On Error Resume Next
        Form_6x.Show()
    End Sub

    Private Sub ButtonConfig_Click(sender As Object, e As EventArgs) Handles ButtonConfig.Click
        On Error Resume Next
        FormConfig.Show()
    End Sub

End Class