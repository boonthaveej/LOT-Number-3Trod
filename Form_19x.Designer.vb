<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form_19x
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.TextBoxNumber = New System.Windows.Forms.TextBox()
        Me.TextBoxTop = New System.Windows.Forms.TextBox()
        Me.TextBoxDown = New System.Windows.Forms.TextBox()
        Me.ButtonEnter = New System.Windows.Forms.Button()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TextBoxNumber
        '
        Me.TextBoxNumber.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextBoxNumber.ForeColor = System.Drawing.Color.Navy
        Me.TextBoxNumber.Location = New System.Drawing.Point(11, 72)
        Me.TextBoxNumber.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBoxNumber.MaxLength = 1
        Me.TextBoxNumber.Name = "TextBoxNumber"
        Me.TextBoxNumber.Size = New System.Drawing.Size(32, 62)
        Me.TextBoxNumber.TabIndex = 1
        Me.TextBoxNumber.Text = "0"
        Me.TextBoxNumber.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxTop
        '
        Me.TextBoxTop.BackColor = System.Drawing.SystemColors.Window
        Me.TextBoxTop.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextBoxTop.ForeColor = System.Drawing.Color.Navy
        Me.TextBoxTop.Location = New System.Drawing.Point(88, 72)
        Me.TextBoxTop.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBoxTop.MaxLength = 5
        Me.TextBoxTop.Name = "TextBoxTop"
        Me.TextBoxTop.Size = New System.Drawing.Size(142, 62)
        Me.TextBoxTop.TabIndex = 2
        Me.TextBoxTop.Text = "99999"
        Me.TextBoxTop.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'TextBoxDown
        '
        Me.TextBoxDown.BackColor = System.Drawing.SystemColors.Window
        Me.TextBoxDown.Font = New System.Drawing.Font("Microsoft Sans Serif", 36.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.TextBoxDown.ForeColor = System.Drawing.Color.Navy
        Me.TextBoxDown.Location = New System.Drawing.Point(275, 72)
        Me.TextBoxDown.Margin = New System.Windows.Forms.Padding(2)
        Me.TextBoxDown.MaxLength = 5
        Me.TextBoxDown.Name = "TextBoxDown"
        Me.TextBoxDown.Size = New System.Drawing.Size(142, 62)
        Me.TextBoxDown.TabIndex = 3
        Me.TextBoxDown.Text = "99999"
        Me.TextBoxDown.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'ButtonEnter
        '
        Me.ButtonEnter.BackColor = System.Drawing.SystemColors.MenuHighlight
        Me.ButtonEnter.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.ButtonEnter.ForeColor = System.Drawing.SystemColors.InfoText
        Me.ButtonEnter.Location = New System.Drawing.Point(421, 72)
        Me.ButtonEnter.Margin = New System.Windows.Forms.Padding(2)
        Me.ButtonEnter.Name = "ButtonEnter"
        Me.ButtonEnter.Size = New System.Drawing.Size(113, 62)
        Me.ButtonEnter.TabIndex = 4
        Me.ButtonEnter.Text = "ENTER"
        Me.ButtonEnter.UseVisualStyleBackColor = False
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.BackColor = System.Drawing.SystemColors.Window
        Me.Label6.Font = New System.Drawing.Font("Microsoft Sans Serif", 25.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label6.Location = New System.Drawing.Point(47, 80)
        Me.Label6.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(37, 39)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "="
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.BackColor = System.Drawing.SystemColors.Window
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 25.8!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label1.Location = New System.Drawing.Point(234, 80)
        Me.Label1.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(37, 39)
        Me.Label1.TabIndex = 16
        Me.Label1.Text = "+"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.BackColor = System.Drawing.SystemColors.Window
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label2.ForeColor = System.Drawing.Color.Red
        Me.Label2.Location = New System.Drawing.Point(7, 9)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 0, 2, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(425, 20)
        Me.Label2.TabIndex = 17
        Me.Label2.Text = "กรุณากรอกข้อมูลเลขวิ่ง แค่ตัวเดียว / โปรแกรมจะคำนวณ 19 ประตูเอง"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label3.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label3.Location = New System.Drawing.Point(7, 38)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 20)
        Me.Label3.TabIndex = 18
        Me.Label3.Text = "เลขวิ่ง"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label4.Location = New System.Drawing.Point(123, 38)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(50, 20)
        Me.Label4.TabIndex = 19
        Me.Label4.Text = "วิ่ง บน"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(222, Byte))
        Me.Label5.ForeColor = System.Drawing.Color.DarkBlue
        Me.Label5.Location = New System.Drawing.Point(321, 38)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(55, 20)
        Me.Label5.TabIndex = 20
        Me.Label5.Text = "วิ่ง ล่าง"
        '
        'Form_19x
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.BackColor = System.Drawing.SystemColors.Window
        Me.ClientSize = New System.Drawing.Size(541, 154)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.ButtonEnter)
        Me.Controls.Add(Me.TextBoxDown)
        Me.Controls.Add(Me.TextBoxTop)
        Me.Controls.Add(Me.TextBoxNumber)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form_19x"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Lot-Number x 19 "
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TextBoxNumber As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxTop As System.Windows.Forms.TextBox
    Friend WithEvents TextBoxDown As System.Windows.Forms.TextBox
    Friend WithEvents ButtonEnter As System.Windows.Forms.Button
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
End Class
