Imports System.Data.SqlClient

Public Class ENTRY
    Dim coninfo As New Connectioninfo
    Dim con As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con1 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con2 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con3 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con4 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con5 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con6 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Private Sub ENTRY_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        coninfo.BindCombo(ComboBox1, "Select EMPLOYEE from EMPLOYEES", "EMPLOYEE")
        coninfo.BindCombo(ComboBox2, "Select MACHINE_NO from TARGET", "MACHINE_NO")
        coninfo.BindCombo(ComboBox3, "Select MACHINE_NO from TARGET", "MACHINE_NO")
        
        'coninfo.BindCombo(ComboBox2, "Select MACHINE_NO from TARGET", "MACHINE_NO")
        
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Button2.Visible = True
        ComboBox3.Visible = True
        TextBox8.Visible = True
    End Sub

   

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim EMP As String
        Dim DATEE As Date
        Dim MACHINE_1, MACHINE_2 As String
        Dim DUTY_1, DUTY, DUTY_2 As Integer


        Dim target_12, target_24 As Integer

        Dim arr(), arr2() As String
        Dim Target, Target2 As Integer
        Me.ErrorProvider1.SetError(Me.DUTY, "")
        Me.ErrorProvider1.SetError(Me.DateTimePicker1, "")
        target_12 = 0
        target_24 = 0
        Me.ErrorProvider1.SetError(Me.DUTY, "")
             DATEE = DateTimePicker1.Value
        EMP = ComboBox1.Text



        If DateTimePicker1.Value.Date > DateTime.Today Then
            Me.ErrorProvider1.SetError(Me.DateTimePicker1, "INCORRECT DATE")
            Label24.Visible = False
            RichTextBox1.Visible = False
            Panel1.Visible = False
            Label25.Visible = False
            Label26.Visible = False
            Label27.Visible = False
        ElseIf IsNumeric(TextBox1.Text) = False Or TextBox1.Text = "" Then
            Me.ErrorProvider1.SetError(Me.DUTY, "INVALID DUTY HOURS")
            If TextBox8.Visible = True Then
            ElseIf IsNumeric(TextBox2.Text) = False Or TextBox2.Text = "" Then
                Me.ErrorProvider1.SetError(Me.DUTY, "INVALID DUTY HOURS")
            End If

        Else

            If Button1.Visible = True And Button2.Visible = False Then
                MACHINE_1 = ComboBox2.Text
                DUTY_1 = TextBox1.Text.Trim()
                '''DUTY
                DUTY = DUTY_1
                If DUTY > 24 Or DUTY = 0 Or TextBox1.Text.Trim() = "" Then
                    Me.ErrorProvider1.SetError(Me.DUTY, "INVALID DUTY HOURS")
                    Label24.Visible = False
                    RichTextBox1.Visible = False
                    Panel1.Visible = False
                    Label25.Visible = False
                    Label26.Visible = False
                    Label27.Visible = False
                ElseIf DUTY >= 1 And DUTY < 12 Then
                    Label7.Text = MACHINE_1
                    arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                    Target = Convert.ToInt32(arr(1)) / 12 * DUTY
                    Label6.Text = arr(0)
                    Label23.Text = Target
                    Panel1.Visible = True
                    Label24.Visible = True
                    RichTextBox1.Visible = True
                ElseIf DUTY >= 13 And DUTY < 24 Then
                    Label7.Text = MACHINE_1
                    arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_24 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                    Target = Convert.ToInt32(arr(1)) / 24 * DUTY
                    Label6.Text = arr(0)
                    Label23.Text = Target
                    Panel1.Visible = True
                    Label24.Visible = True
                    RichTextBox1.Visible = True
                ElseIf DUTY = 12 Then
                    Label7.Text = MACHINE_1
                    arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                    Target = Convert.ToInt32(arr(1))
                    Label6.Text = arr(0)
                    Label23.Text = Target
                    Panel1.Visible = True
                    Label24.Visible = False
                    RichTextBox1.Visible = False
                ElseIf DUTY = 24 Then
                    Label7.Text = MACHINE_1
                    arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_24 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                    Target = Convert.ToInt32(arr(1))
                    Label6.Text = arr(0)
                    Label23.Text = Target
                    Panel1.Visible = True
                    Label24.Visible = False
                    RichTextBox1.Visible = False
                End If
                '''''




            End If


            '===================================================================



            '===================================================================

            If Button2.Visible = True Then



                MACHINE_1 = ComboBox2.Text
                MACHINE_2 = ComboBox3.Text
                DUTY_1 = Convert.ToInt32(TextBox1.Text)
                DUTY_2 = Convert.ToInt32(TextBox8.Text)
                '''DUTY
                DUTY = DUTY_1 + DUTY_2
                If DUTY > 24 Or DUTY = 0 Then
                    Me.ErrorProvider1.SetError(Me.DUTY, "INVALID DUTY HOURS")
                    Label24.Visible = False
                    RichTextBox1.Visible = False
                    Panel1.Visible = False
                    Label25.Visible = False
                    Label26.Visible = False
                    Label27.Visible = False
                ElseIf DUTY >= 1 And DUTY < 12 Then
                    Label7.Text = MACHINE_1
                    Label9.Text = MACHINE_2

                    arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                    arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")

                    Target = Convert.ToInt32(arr(1)) / 12 * DUTY_1
                    Target2 = Convert.ToInt32(arr2(1)) / 12 * DUTY_2
                    Label6.Text = arr(0)
                    Label8.Text = arr2(0)
                    Label23.Text = Target
                    Label22.Text = Target2

                    Panel1.Visible = True
                    Label24.Visible = True
                    Label8.Visible = True
                    Label9.Visible = True
                    Label22.Visible = True
                    TextBox3.Visible = True
                    Label24.Visible = True
                    RichTextBox1.Visible = True
                ElseIf DUTY >= 13 And DUTY <= 24 Then


                    If DUTY_1 < 12 And DUTY_2 >= 12 Then

                        Label7.Text = MACHINE_1
                        Label9.Text = MACHINE_2

                        arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                        If DUTY_2 = 12 Then
                            arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")
                            Target2 = Convert.ToInt32(arr2(1))
                            Label8.Text = arr2(0)
                        Else
                            arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_24 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")
                            Target2 = Convert.ToInt32(arr2(1)) / 24 * DUTY_2
                            Label8.Text = arr2(0)
                        End If


                        Target = Convert.ToInt32(arr(1)) / 12 * DUTY_1

                        Label6.Text = arr(0)

                        Label23.Text = Target
                        Label22.Text = Target2

                        Panel1.Visible = True
                        Label24.Visible = True
                        Label8.Visible = True
                        Label9.Visible = True
                        Label22.Visible = True
                        TextBox3.Visible = True

                        If DUTY = 24 Then
                            Label24.Visible = False
                            RichTextBox1.Visible = False
                        Else
                            Label24.Visible = True
                            RichTextBox1.Visible = True
                        End If
                    ElseIf DUTY_2 < 12 And DUTY_1 >= 12 Then
                        Label7.Text = MACHINE_1
                        Label9.Text = MACHINE_2
                        If DUTY_1 = 12 Then
                            arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                            Target = Convert.ToInt32(arr(1))
                            Label6.Text = arr(0)
                        Else
                            arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_24 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                            Target = Convert.ToInt32(arr(1)) / 24 * DUTY_1
                            Label6.Text = arr(0)
                        End If

                        arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")


                        Target2 = Convert.ToInt32(arr2(1)) / 12 * DUTY_2

                        Label8.Text = arr2(0)
                        Label23.Text = Target
                        Label22.Text = Target2

                        Panel1.Visible = True
                        Label24.Visible = True
                        Label8.Visible = True
                        Label9.Visible = True
                        Label22.Visible = True
                        TextBox3.Visible = True
                        If DUTY = 24 Then
                            Label24.Visible = False
                            RichTextBox1.Visible = False
                        Else
                            Label24.Visible = True
                            RichTextBox1.Visible = True
                        End If
                    ElseIf DUTY_2 = 12 And DUTY_1 = 12 Then
                        Label7.Text = MACHINE_1
                        Label9.Text = MACHINE_2

                        arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                        arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")

                        Target = Convert.ToInt32(arr(1)) / 12 * DUTY_1
                        Target2 = Convert.ToInt32(arr2(1)) / 12 * DUTY_2
                        Label6.Text = arr(0)
                        Label8.Text = arr2(0)
                        Label23.Text = Target
                        Label22.Text = Target2

                        Panel1.Visible = True
                        Label24.Visible = True
                        Label8.Visible = True
                        Label9.Visible = True
                        Label22.Visible = True
                        TextBox3.Visible = True
                        If DUTY = 24 Then
                            Label24.Visible = False
                            RichTextBox1.Visible = False
                        Else
                            Label24.Visible = True
                            RichTextBox1.Visible = True
                        End If
                    ElseIf DUTY_2 < 12 And DUTY_1 < 12 Then
                        Label7.Text = MACHINE_1
                        Label9.Text = MACHINE_2

                        arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                        arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")

                        Target = Convert.ToInt32(arr(1)) / 12 * DUTY_1
                        Target2 = Convert.ToInt32(arr2(1)) / 12 * DUTY_2
                        Label6.Text = arr(0)
                        Label8.Text = arr2(0)
                        Label23.Text = Target
                        Label22.Text = Target2

                        Panel1.Visible = True
                        Label24.Visible = True
                        Label8.Visible = True
                        Label9.Visible = True
                        Label22.Visible = True
                        TextBox3.Visible = True

                        If DUTY = 24 Then
                            Label24.Visible = False
                            RichTextBox1.Visible = False
                        Else
                            Label24.Visible = True
                            RichTextBox1.Visible = True
                        End If

                    ElseIf DUTY = 12 Then
                        Label7.Text = MACHINE_1
                        Label9.Text = MACHINE_2

                        arr = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_1 & "'")
                        arr2 = coninfo.GetTarget("SELECT FUEL_NOZZLE,TARGET_12 FROM TARGET WHERE MACHINE_NO='" & MACHINE_2 & "'")

                        Target = Convert.ToInt32(arr(1)) / 12 * DUTY_1
                        Target2 = Convert.ToInt32(arr2(1)) / 12 * DUTY_2
                        Label6.Text = arr(0)
                        Label8.Text = arr2(0)
                        Label23.Text = Target
                        Label22.Text = Target2

                        Panel1.Visible = True
                        Label24.Visible = True
                        Label8.Visible = True
                        Label9.Visible = True
                        Label22.Visible = True
                        TextBox3.Visible = True
                        Label24.Visible = False
                        RichTextBox1.Visible = False

                    End If

                End If
            End If



        End If










            







    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        Dim SALE1, SALE2 As Integer
        Dim TARGET1, TARGET2 As Integer
        Dim POINT, POINT1, POINT2 As Integer
        Dim date_e As Date
        Dim DUTY_1, DUTY_2 As String
        Dim emp, des As String
        Dim MACHINE_1, MACHINE_2 As String
        Dim FUEL_1, FUEL_2 As String
        Dim ROWS, ROWS2, ROWS3 As Integer
        Dim QRY, QRY2, QRY3 As String
        date_e = DateTimePicker1.Value


        Me.ErrorProvider1.SetError(Me.RichTextBox1, "")
        Me.ErrorProvider1.SetError(Me.TextBox2, "")
        Me.ErrorProvider1.SetError(Me.TextBox3, "")
        Me.ErrorProvider1.SetError(Me.RichTextBox1, "")







        If IsNumeric(TextBox2.Text) = False Or TextBox2.Text = "" Then
            Me.ErrorProvider1.SetError(Me.TextBox2, "INVALID SALE")

        ElseIf Button2.Visible = True And (IsNumeric(TextBox3.Text) = False Or TextBox3.Text = "") Then
            Me.ErrorProvider1.SetError(Me.TextBox3, "INVALID SALE")

        ElseIf RichTextBox1.Visible = True And RichTextBox1.Text = "" Then
            Me.ErrorProvider1.SetError(Me.RichTextBox1, "INVALID INPUT")

        ElseIf Button2.Visible = True Then



            emp = ComboBox1.Text
            DUTY_1 = TextBox1.Text
            DUTY_2 = TextBox8.Text
            MACHINE_1 = Label7.Text
            MACHINE_2 = Label9.Text
            FUEL_1 = Label6.Text
            FUEL_2 = Label8.Text
            TARGET1 = Label23.Text
            TARGET2 = Label22.Text
            SALE1 = TextBox2.Text
            SALE2 = TextBox3.Text
            POINT1 = Convert.ToInt32(SALE1) - Convert.ToInt32(TARGET1)
            POINT2 = Convert.ToInt32(SALE2) - Convert.ToInt32(TARGET2)
            POINT = POINT1 + POINT2

            If Button2.Visible = True Then
                des = RichTextBox1.Text
            Else
                des = "NULL"
            End If

            QRY = "INSERT INTO SALES (DATE,EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION) VALUES (@DATE,@EMP,@DUTY,@MAC,@FUEL,@TAR,@SALES,@POINT,@DES)"
            Dim cmd As New SqlCommand(QRY, con)

            cmd.Parameters.AddWithValue("@DATE", date_e)
            cmd.Parameters.AddWithValue("@EMP", emp)
            cmd.Parameters.AddWithValue("@DUTY", DUTY_1)
            cmd.Parameters.AddWithValue("@MAC", MACHINE_1)
            cmd.Parameters.AddWithValue("@FUEL", FUEL_1)
            cmd.Parameters.AddWithValue("@TAR", TARGET1)
            cmd.Parameters.AddWithValue("@SALES", SALE1)
            cmd.Parameters.AddWithValue("@POINT", POINT1)
            cmd.Parameters.AddWithValue("@DES", des)
            con.Open()

            ROWS = cmd.ExecuteNonQuery()
            If ROWS > 0 Then

                If Button2.Visible = True Then
                    des = RichTextBox1.Text
                Else
                    des = "NULL"
                End If

                QRY2 = "INSERT INTO SALES (DATE,EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION) VALUES (@DATE,@EMP,@DUTY,@MAC,@FUEL,@TAR,@SALES,@POINT,@DES)"
                Dim cmd2 As New SqlCommand(QRY2, con2)

                cmd2.Parameters.AddWithValue("@DATE", date_e)
                cmd2.Parameters.AddWithValue("@EMP", emp)
                cmd2.Parameters.AddWithValue("@DUTY", DUTY_2)
                cmd2.Parameters.AddWithValue("@MAC", MACHINE_2)
                cmd2.Parameters.AddWithValue("@FUEL", FUEL_2)
                cmd2.Parameters.AddWithValue("@TAR", TARGET2)
                cmd2.Parameters.AddWithValue("@SALES", SALE2)
                cmd2.Parameters.AddWithValue("@POINT", POINT2)
                cmd2.Parameters.AddWithValue("@DES", des)

                con2.Open()

                ROWS2 = cmd2.ExecuteNonQuery()


                If ROWS2 > 0 Then
                    MsgBox("SALES SUCCESSFULLY UPDATED")
                    Panel1.Visible = False
                    TextBox1.Text = ""
                    TextBox8.Text = ""
                    TextBox2.Text = ""
                    TextBox3.Text = ""
                    RichTextBox1.Text = ""
                    Button2.Visible = False
                    ComboBox3.Visible = False
                    TextBox8.Visible = False
                    ComboBox1.Text = ""
                    ComboBox2.Text = ""
                    ComboBox3.Text = ""
                End If

            End If
            con.Close()
            con2.Close()

        ElseIf Button2.Visible = False Then

            SALE1 = TextBox2.Text
            emp = ComboBox1.Text
            TARGET1 = Label23.Text
            DUTY_1 = TextBox1.Text
            MACHINE_1 = Label7.Text
            POINT1 = Convert.ToInt32(SALE1) - Convert.ToInt32(TARGET1)
            POINT = POINT1
            FUEL_1 = Label6.Text
            If Label24.Visible = True Then
                des = RichTextBox1.Text
            Else
                des = "NULL"
            End If
            QRY = "INSERT INTO SALES (DATE,EMPLOYEE,DUTY,MACHINE_NO,FUEL_NOZZLE,TARGET,SALES,POINTS,DESCRIPTION) VALUES (@DATE,@EMP,@DUTY,@MAC,@FUEL,@TAR,@SALES,@POINT,@DES)"
            Dim cmd As New SqlCommand(QRY, con)

            cmd.Parameters.AddWithValue("@DATE", date_e)
            cmd.Parameters.AddWithValue("@EMP", emp)
            cmd.Parameters.AddWithValue("@DUTY", DUTY_1)
            cmd.Parameters.AddWithValue("@MAC", MACHINE_1)
            cmd.Parameters.AddWithValue("@FUEL", FUEL_1)
            cmd.Parameters.AddWithValue("@TAR", TARGET1)
            cmd.Parameters.AddWithValue("@SALES", SALE1)
            cmd.Parameters.AddWithValue("@POINT", POINT1)
            cmd.Parameters.AddWithValue("@DES", des)
            con.Open()

            ROWS = cmd.ExecuteNonQuery()
            If ROWS > 0 Then
                MsgBox("SALES SUCCESSFULLY UPDATED")
                Panel1.Visible = False
                TextBox1.Text = ""
                TextBox8.Text = ""
                TextBox2.Text = ""
                TextBox3.Text = ""
                ComboBox3.Visible = False
                Button2.Visible = False
                TextBox8.Visible = False
                RichTextBox1.Text = ""
                ComboBox1.Text = ""
                ComboBox2.Text = ""
                ComboBox3.Text = ""
            End If
            con.Close()


        End If


















    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim SALE1, SALE2 As Integer
        Dim TARGET1, TARGET2 As Integer
        Dim POINT, POINT1, POINT2 As Integer
        Dim QRY1, EMP As String
        Me.ErrorProvider1.SetError(Me.TextBox2, "")
        Me.ErrorProvider1.SetError(Me.TextBox3, "")





        If IsNumeric(TextBox2.Text) = False Or TextBox2.Text = "" Then
            Me.ErrorProvider1.SetError(Me.TextBox2, "INVALID SALE")
            If Button2.Visible = True Then
                If IsNumeric(TextBox3.Text) = False Or TextBox3.Text = "" Then
                    Me.ErrorProvider1.SetError(Me.TextBox3, "INVALID SALE")
                End If
            End If
        Else
            If Button2.Visible = True Then
                TARGET1 = Convert.ToInt32(Label23.Text)
                TARGET2 = Convert.ToInt32(Label22.Text)
                SALE1 = Convert.ToInt32(TextBox2.Text)
                SALE2 = Convert.ToInt32(TextBox3.Text)
                POINT1 = Convert.ToInt32(SALE1) - Convert.ToInt32(TARGET1)
                POINT2 = Convert.ToInt32(SALE2) - Convert.ToInt32(TARGET2)
                POINT = POINT1 + POINT2

                EMP = ComboBox1.Text

                Dim current, today As Date
                Dim val As String

                current = today.Year & "-" & today.Month & "-" & "01"
                QRY1 = "SELECT SUM(POINTS) FROM SALES WHERE DATE > '" & current & "'AND DATE <= '" & Date.Today.ToString("yyyy-MM-dd") & "' AND EMPLOYEE='" & EMP & "'"




                val = coninfo.ExecuteScaler(QRY1)
                If val <> "" Then
                    Label25.Visible = True
                    Label26.Visible = True
                    Label27.Visible = True
                    Label27.Text = val
                    Label26.Text = POINT
                End If


            ElseIf Button2.Visible = False Then

                TARGET1 = Convert.ToInt32(Label23.Text)

                SALE1 = Convert.ToInt32(TextBox2.Text)

                POINT1 = Convert.ToInt32(SALE1) - Convert.ToInt32(TARGET1)

                POINT = POINT1

                EMP = ComboBox1.Text

                Dim current, today As Date
                Dim val As String

                current = today.Year & "-" & today.Month & "-" & "01"
                QRY1 = "SELECT SUM(POINTS) FROM SALES WHERE DATE > '" & current & "'AND DATE <= '" & Date.Today.ToString("yyyy-MM-dd") & "' AND EMPLOYEE='" & EMP & "'"

                


                val = coninfo.ExecuteScaler(QRY1)
                If val <> "" Then
                    Label25.Visible = True
                    Label26.Visible = True
                    Label27.Visible = True
                    Label27.Text = val
                    Label26.Text = POINT
                End If


            End If



        End If









        
    End Sub
End Class