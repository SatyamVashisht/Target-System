Imports System.Text.RegularExpressions
Imports System.Data.SqlClient

Public Class EMPLOYEE
    Dim coninfo As New Connectioninfo
    Dim con As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Private Sub ADD_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ADD.CheckedChanged
        If ADD.Checked = True Then
            Panel1.Visible = True
            Panel2.Visible = False
        ElseIf ADD.Checked = False Then
            Panel1.Visible = False
            Panel2.Visible = True

        End If
    End Sub

    Private Sub DELETE_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETE.CheckedChanged
        If DELETE.Checked = True Then
            Panel1.Visible = False
            Panel2.Visible = True
        ElseIf DELETE.Checked = False Then
            Panel1.Visible = True
            Panel2.Visible = False
        End If
        coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim NAME_E, QRY As String
        Dim ROWS As Integer

        NAME_E = TextBox1.Text
        Me.ErrorProvider1.SetError(Me.TextBox1, "")

        If Not Regex.Match(TextBox1.Text, "^[a-zA-Z ]*$").Success Or (TextBox1.Text = "") Then
            Me.ErrorProvider1.SetError(Me.TextBox1, "INVALID NAME")
        Else
            QRY = "INSERT INTO EMPLOYEES (EMPLOYEE) VALUES (@NAME)"
            Dim cmd As New SqlCommand(QRY, con)

            cmd.Parameters.AddWithValue("@NAME", NAME_E)

            con.Open()

            ROWS = cmd.ExecuteNonQuery()

            If ROWS > 0 Then

                MessageBox.Show("EMPLOYEE ADDED")
                TextBox1.Text = ""
                coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
            Else
                MessageBox.Show("EMPLOYEE NOT ADDED")
                TextBox1.Text = ""
            End If
            con.Close()
        End If
    End Sub

    Private Sub EMPLOYEE_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim NAME_E, QRY As String
        Dim ROWS As Integer

        NAME_E = ComboBox1.Text
        QRY = "DELETE FROM EMPLOYEES WHERE EMPLOYEE = @NAME"
        Dim cmd As New SqlCommand(QRY, con)

        cmd.Parameters.AddWithValue("@NAME", NAME_E)

        con.Open()

        ROWS = cmd.ExecuteNonQuery()

        If ROWS > 0 Then

            MessageBox.Show("EMPLOYEE DELETED")
            TextBox1.Text = ""
            coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
        Else
            MessageBox.Show("EMPLOYEE NOT DELETED")
            TextBox1.Text = ""
        End If
        con.Close()

    End Sub
End Class