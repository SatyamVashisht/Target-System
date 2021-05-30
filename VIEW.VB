Imports System.Data.SqlClient

Public Class VIEW
    Dim coninfo As New Connectioninfo
    Dim con As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Private Sub RadioButton6_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton6.CheckedChanged


        If RadioButton6.Checked = True Then
            DateTimePicker1.Visible = True
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Button2.Visible = False
            Label8.Visible = False
            Label9.Visible = False
            Label10.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            ComboBox1.Visible = False
            DateTimePicker4.Visible = False
            DateTimePicker5.Visible = False
            CheckBox1.Visible = True
            Label1.Visible = False
            Label2.Visible = False
            DateTimePicker2.Visible = False
            DateTimePicker3.Visible = False

        End If
    End Sub

    Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged
        If RadioButton1.Checked = True Then
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Label8.Visible = False
            Label9.Visible = False
            Button2.Visible = False
            Label10.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            CheckBox1.Visible = False
            ComboBox1.Visible = False
            DateTimePicker4.Visible = False
            DateTimePicker5.Visible = False
            DateTimePicker1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            DateTimePicker2.Visible = False
            DateTimePicker3.Visible = False
        End If

    End Sub

    Private Sub RadioButton2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton2.CheckedChanged
        If RadioButton2.Checked = True Then
            Label5.Visible = False
            Label6.Visible = False
            Label7.Visible = False
            Button2.Visible = False
            Label8.Visible = False
            Label9.Visible = False
            Label10.Visible = False
            Label3.Visible = False
            Label4.Visible = False
            ComboBox1.Visible = False
            CheckBox1.Visible = False
            DateTimePicker4.Visible = False
            DateTimePicker5.Visible = False
            DateTimePicker1.Visible = False
            Label1.Visible = False
            Label2.Visible = False
            DateTimePicker2.Visible = False
            DateTimePicker3.Visible = False
        End If


    End Sub

    Private Sub RadioButton3_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton3.CheckedChanged

        If RadioButton3.Checked = True Then
            Label5.Visible = False
            Label6.Visible = False
            Button2.Visible = False
            Label7.Visible = False
            Label8.Visible = False
            Label9.Visible = False
            Label10.Visible = False
            Label1.Visible = True
            Label2.Visible = True
            DateTimePicker2.Visible = True
            DateTimePicker3.Visible = True
            CheckBox1.Visible = True
            Label3.Visible = False
            Label4.Visible = False
            ComboBox1.Visible = False
            DateTimePicker4.Visible = False
            DateTimePicker5.Visible = False
            DateTimePicker1.Visible = False
        End If
    End Sub

    Private Sub RadioButton4_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton4.CheckedChanged


        If RadioButton4.Checked = True Then
            Label5.Visible = True
            Label6.Visible = True
            Label7.Visible = False
            Label8.Visible = True
            Label9.Visible = False
            Label10.Visible = False
            Label3.Visible = True
            Label4.Visible = True
            ComboBox1.Visible = True
            DateTimePicker4.Visible = True
            DateTimePicker5.Visible = True
            CheckBox1.Visible = True
            Label1.Visible = False
            Label2.Visible = False
            DateTimePicker2.Visible = False
            DateTimePicker3.Visible = False
            DateTimePicker1.Visible = False
            coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
        End If

    End Sub

    Private Sub VIEW_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        RadioButton1.Checked = True
        coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
        coninfo.BindCombo(ComboBox2, "Select MACHINE_NO from TARGET", "MACHINE_NO")
        coninfo.BindCombo(ComboBox3, "Select FUEL_NOZZLE from TARGET", "FUEL_NOZZLE")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If RadioButton1.Checked = True Then
            Button2.Visible = False
            coninfo.BindGrid(DataGridView1, "Select MACHINE_NO,FUEL_NOZZLE,TARGET_12,TARGET_24 from TARGET ")
        ElseIf RadioButton2.Checked = True Then
            Button2.Visible = False
            coninfo.BindGrid(DataGridView1, "Select MACHINE_NO,FUEL_NOZZLE,AVARAGE,TARGET_12,TARGET_24 from TARGET ")

        ElseIf RadioButton6.Checked = True And CheckBox1.Checked = True Then
            Button2.Visible = False
            coninfo.BindGrid(DataGridView1, "Select EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION from SALES WHERE DATE ='" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' ")
        ElseIf RadioButton6.Checked = True And CheckBox1.Checked = False Then
            Button2.Visible = False
            coninfo.BindGrid(DataGridView1, "Select EMPLOYEE,FUEL_NOZZLE,TARGET,SALES,POINTS from SALES WHERE DATE ='" & DateTimePicker1.Value.ToString("yyyy-MM-dd") & "' ")
        ElseIf RadioButton3.Checked = True And CheckBox1.Checked = True Then
            coninfo.BindGrid(DataGridView1, "Select DATE , EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION from SALES WHERE DATE >= '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker3.Value.ToString("yyyy-MM-dd") & "' ORDER BY DATE ")
            Button2.Visible = False
        ElseIf RadioButton3.Checked = True And CheckBox1.Checked = False Then
            Button2.Visible = False
            coninfo.BindGrid(DataGridView1, "Select DATE , EMPLOYEE,FUEL_NOZZLE,TARGET,SALES,POINTS from SALES WHERE DATE >= '" & DateTimePicker2.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker3.Value.ToString("yyyy-MM-dd") & "' ORDER BY DATE")
        ElseIf RadioButton4.Checked = True And CheckBox1.Checked = True Then
            Dim VAL1, VAL2, VAL3 As String
            Dim QRY1, QRY2, QRY3 As String

            QRY1 = "SELECT SUM(POINTS) FROM SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "'"
            VAL1 = coninfo.ExecuteScaler(QRY1)
            If VAL1 <> "" Then
                Label7.Visible = True
                Label7.Text = VAL1
            End If
            QRY2 = "SELECT SUM(SALES) FROM SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "'"
            VAL2 = coninfo.ExecuteScaler(QRY2)
            If VAL2 <> "" Then
                Label9.Visible = True
                Label9.Text = VAL2
            End If
            QRY3 = "SELECT SUM(TARGET) FROM SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "'"
            VAL3 = coninfo.ExecuteScaler(QRY3)
            If VAL3 <> "" Then
                Label10.Visible = True
                Label10.Text = VAL3
            End If

            coninfo.BindGrid(DataGridView1, "Select DATE , EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION from SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "' ORDER BY DATE")
            Button2.Visible = True
        ElseIf RadioButton4.Checked = True And CheckBox1.Checked = False Then
            Dim VAL1, VAL2, VAL3 As String
            Dim QRY1, QRY2, QRY3 As String
            Button2.Visible = False
            QRY1 = "SELECT SUM(POINTS) FROM SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "'"
            VAL1 = coninfo.ExecuteScaler(QRY1)
            If VAL1 <> "" Then
                Label7.Visible = True
                Label7.Text = VAL1
            End If
            QRY2 = "SELECT SUM(SALES) FROM SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "'"
            VAL2 = coninfo.ExecuteScaler(QRY2)
            If VAL2 <> "" Then
                Label9.Visible = True
                Label9.Text = VAL2
            End If
            QRY3 = "SELECT SUM(TARGET) FROM SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "'"
            VAL3 = coninfo.ExecuteScaler(QRY3)
            If VAL3 <> "" Then
                Label10.Visible = True
                Label10.Text = VAL3
            End If


            coninfo.BindGrid(DataGridView1, "Select DATE , EMPLOYEE,FUEL_NOZZLE,TARGET,SALES,POINTS from SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "' ORDER BY DATE")

        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)



    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        GroupBox2.Visible = True
        GroupBox1.Visible = False
    End Sub

    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        Dim DATE_E As Date
        Dim EMP, MAC, FUEL, DES, QRY As String
        Dim DUTY, TAR, SALES, POINTS As Integer
        Dim ROWS As Integer = 0

        DATE_E = DateTimePicker10.Value
        EMP = Label13.Text
        MAC = ComboBox2.Text
        FUEL = ComboBox3.Text
        DES = TextBox5.Text
        DUTY = Convert.ToInt32(TextBox1.Text)
        TAR = Convert.ToInt32(TextBox2.Text)
        SALES = Convert.ToInt32(TextBox3.Text)

        POINTS = SALES - TAR
        Me.ErrorProvider1.SetError(Me.DateTimePicker10, "")
        Me.ErrorProvider1.SetError(Me.TextBox1, "")
        Me.ErrorProvider1.SetError(Me.TextBox4, "")
        Me.ErrorProvider1.SetError(Me.TextBox2, "")
        Me.ErrorProvider1.SetError(Me.TextBox3, "")

        
            If IsNumeric(TextBox1.Text) = False Or TextBox1.Text = "" Then
                Me.ErrorProvider1.SetError(Me.TextBox1, "INVALID DUTY HOURS")
            ElseIf TextBox4.Text = "" Then
                Me.ErrorProvider1.SetError(Me.TextBox4, "INVALID DESCRIPTION")
            ElseIf IsNumeric(TextBox2.Text) = False Or TextBox2.Text = "" Then
                Me.ErrorProvider1.SetError(Me.TextBox2, "INVALID TARGET")
            ElseIf IsNumeric(TextBox3.Text) = False Or TextBox3.Text = "" Then
                Me.ErrorProvider1.SetError(Me.TextBox3, "INVALID SALES")
            Else
            QRY = "UPDATE SALES SET TARGET = @TARGET ,SALES = @SALES , POINTS = @POINTS, DESCRIPTION = @DESCRIPTION WHERE DATE = '" & DateTimePicker10.Value.ToString("yyyy-MM-dd") & "' AND MACHINE_NO = '" & MAC & "' AND EMPLOYEE = '" & EMP & "' AND DUTY = '" & DUTY & "' "
                Dim cmd As New SqlCommand(QRY, con)




                cmd.Parameters.AddWithValue("@TARGET", TAR)
                cmd.Parameters.AddWithValue("@SALES", SALES)
                cmd.Parameters.AddWithValue("@POINTS", POINTS)
                cmd.Parameters.AddWithValue("@DESCRIPTION", DES)

                con.Open()

                ROWS = cmd.ExecuteNonQuery()

                If ROWS > 0 Then
                MsgBox("SUCCESSFULLY UPDATED")
                coninfo.BindGrid(DataGridView1, "Select DATE , EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION from SALES WHERE DATE >= '" & DateTimePicker5.Value.ToString("yyyy-MM-dd") & "'AND DATE <= '" & DateTimePicker4.Value.ToString("yyyy-MM-dd") & "'AND EMPLOYEE='" & ComboBox1.Text & "' ORDER BY DATE")
                Else
                    MsgBox("NOT SUCCESSFULLY UPDATED")
                End If
                con.Close()






            End If



















    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        If GroupBox2.Visible = True Then
            If e.RowIndex >= 0 Then
                Dim row As DataGridViewRow = DataGridView1.Rows(e.RowIndex)

                DateTimePicker10.Value = row.Cells(0).Value
                Label13.Text = row.Cells(1).Value
                TextBox1.Text = row.Cells(2).Value.ToString()
                ComboBox2.Text = row.Cells(3).Value.ToString()
                ComboBox3.Text = row.Cells(4).Value.ToString()
                TextBox2.Text = row.Cells(5).Value.ToString()
                TextBox3.Text = row.Cells(6).Value.ToString()
                TextBox4.Text = row.Cells(7).Value.ToString()
                TextBox5.Text = row.Cells(8).Value.ToString()

            End If
        End If

    End Sub


    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        GroupBox2.Visible = False
        GroupBox1.Visible = True
        Label13.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
        TextBox1.Text = ""
        ComboBox2.Text = ""
        ComboBox3.Text = ""
    End Sub
End Class