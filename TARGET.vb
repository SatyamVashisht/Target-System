Imports System.Data.SqlClient

Public Class TARGET
    Dim coninfo As New Connectioninfo
    Dim con1 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con2 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con3 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con4 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con5 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con6 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con7 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con8 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con9 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con10 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con11 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con12 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con13 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con14 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con15 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con16 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")


    Dim con_1 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_2 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_3 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_4 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_5 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_6 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_7 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_8 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_9 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_10 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_11 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_12 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_13 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_14 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_15 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")
    Dim con_16 As New SqlConnection("Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;")



    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim PWD As String
        PWD = MyInputBox("Enter password", "Enter password", TextBox1.Text, True)

        If PWD = "UPTARGET" Or PWD = "uptarget" Then
            Button2.Visible = True
            TextBox1.Enabled = True
            TextBox2.Enabled = True
            TextBox3.Enabled = True
            TextBox4.Enabled = True
            TextBox5.Enabled = True
            TextBox6.Enabled = True
            TextBox7.Enabled = True
            TextBox8.Enabled = True
            TextBox9.Enabled = True
            TextBox10.Enabled = True
            TextBox11.Enabled = True
            TextBox12.Enabled = True
            TextBox13.Enabled = True
            TextBox14.Enabled = True
            TextBox15.Enabled = True
            TextBox16.Enabled = True
            TextBox17.Enabled = True
            TextBox18.Enabled = True
            TextBox19.Enabled = True
            TextBox20.Enabled = True
            TextBox21.Enabled = True
            TextBox22.Enabled = True
            TextBox23.Enabled = True
            TextBox24.Enabled = True
            TextBox25.Enabled = True
            TextBox26.Enabled = True
            TextBox27.Enabled = True
            TextBox28.Enabled = True
            TextBox29.Enabled = True
            TextBox30.Enabled = True
            TextBox31.Enabled = True
            TextBox32.Enabled = True
            TextBox33.Enabled = True
            TextBox34.Enabled = True
            TextBox35.Enabled = True
            TextBox36.Enabled = True
            TextBox37.Enabled = True
            TextBox38.Enabled = True
            TextBox39.Enabled = True
            TextBox40.Enabled = True
            TextBox41.Enabled = True
            TextBox42.Enabled = True
            TextBox43.Enabled = True
            TextBox44.Enabled = True
            TextBox45.Enabled = True
            TextBox46.Enabled = True
            TextBox47.Enabled = True
            TextBox48.Enabled = True
        Else
            MsgBox("INCORRECT PASSWORD")
        End If






    End Sub

    Private Sub TARGET_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ROW1, ROW2, ROW3, ROW4, ROW5, ROW6, ROW7, ROW8, ROW9, ROW10, ROW11, ROW12, ROW13, ROW14, ROW15, ROW16 As SqlDataReader
        Dim QRY1, QRY2, QRY3, QRY4, QRY5, QRY6, QRY7, QRY8, QRY9, QRY10, QRY11, QRY12, QRY13, QRY14, QRY15, QRY16 As String

        QRY1 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '4' AND FUEL_NOZZLE = 'ULP-1'"
        ''''CHANGE QRY
        QRY2 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '3' AND FUEL_NOZZLE = 'ULP-2'"
        QRY3 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '6' AND FUEL_NOZZLE = 'ULP-3'"
        QRY4 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '5' AND FUEL_NOZZLE = 'ULP-4'"
        QRY5 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '9' AND FUEL_NOZZLE = 'ULP-5,ULP-7'"
        QRY6 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '10' AND FUEL_NOZZLE = 'ULP-8,ULP-6'"

        QRY7 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '7' AND FUEL_NOZZLE = 'XP-4'"
        QRY8 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '8' AND FUEL_NOZZLE = 'XP-3'"
        QRY9 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '11,12' AND FUEL_NOZZLE = 'XP-2'"
        QRY10 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = '13,14' AND FUEL_NOZZLE = 'XP-1'"

        QRY11 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = 'CNG-1' AND FUEL_NOZZLE = 'CNG-1'"
        QRY12 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = 'CNG_1' AND FUEL_NOZZLE = 'CNG-2'"
        QRY13 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = 'CNG-2' AND FUEL_NOZZLE = 'CNG-3'"
        QRY14 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = 'CNG_2' AND FUEL_NOZZLE = 'CNG-4'"
        QRY15 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = 'CNG-3' AND FUEL_NOZZLE = 'CNG-5'"
        QRY16 = "SELECT AVARAGE , TARGET_12 , TARGET_24 FROM TARGET WHERE MACHINE_NO = 'CNG_3' AND FUEL_NOZZLE = 'CNG-6'"


        ROW1 = coninfo.ExecuteDQL(QRY1, con_1)
        If ROW1.Read() Then
            TextBox12.Text = ROW1("AVARAGE").ToString()
            TextBox1.Text = ROW1("TARGET_12").ToString()
            TextBox18.Text = ROW1("TARGET_24").ToString()
        End If
        ROW2 = coninfo.ExecuteDQL(QRY2, con_2)
        If ROW2.Read() Then
            TextBox11.Text = ROW2("AVARAGE").ToString()
            TextBox2.Text = ROW2("TARGET_12").ToString()
            TextBox17.Text = ROW2("TARGET_24").ToString()
        End If
        ROW3 = coninfo.ExecuteDQL(QRY3, con_3)
        If ROW3.Read() Then
            TextBox10.Text = ROW3("AVARAGE").ToString()
            TextBox4.Text = ROW3("TARGET_12").ToString()
            TextBox16.Text = ROW3("TARGET_24").ToString()
        End If
        ROW4 = coninfo.ExecuteDQL(QRY4, con_4)
        If ROW4.Read() Then
            TextBox9.Text = ROW4("AVARAGE").ToString()
            TextBox3.Text = ROW4("TARGET_12").ToString()
            TextBox15.Text = ROW4("TARGET_24").ToString()
        End If
        ROW5 = coninfo.ExecuteDQL(QRY5, con_5)
        If ROW5.Read() Then
            TextBox6.Text = ROW5("AVARAGE").ToString()
            TextBox8.Text = ROW5("TARGET_12").ToString()
            TextBox14.Text = ROW5("TARGET_24").ToString()
        End If
        ROW6 = coninfo.ExecuteDQL(QRY6, con_6)
        If ROW6.Read() Then
            TextBox5.Text = ROW6("AVARAGE").ToString()
            TextBox7.Text = ROW6("TARGET_12").ToString()
            TextBox13.Text = ROW6("TARGET_24").ToString()
        End If
        ROW7 = coninfo.ExecuteDQL(QRY7, con_7)
        If ROW7.Read() Then
            TextBox28.Text = ROW7("AVARAGE").ToString()
            TextBox34.Text = ROW7("TARGET_12").ToString()
            TextBox22.Text = ROW7("TARGET_24").ToString()
        End If
        ROW8 = coninfo.ExecuteDQL(QRY8, con_8)
        If ROW8.Read() Then
            TextBox27.Text = ROW8("AVARAGE").ToString()
            TextBox33.Text = ROW8("TARGET_12").ToString()
            TextBox21.Text = ROW8("TARGET_24").ToString()
        End If
        ROW9 = coninfo.ExecuteDQL(QRY9, con_9)
        If ROW9.Read() Then
            TextBox26.Text = ROW9("AVARAGE").ToString()
            TextBox32.Text = ROW9("TARGET_12").ToString()
            TextBox20.Text = ROW9("TARGET_24").ToString()
        End If
        ROW10 = coninfo.ExecuteDQL(QRY10, con_10)
        If ROW10.Read() Then
            TextBox25.Text = ROW10("AVARAGE").ToString()
            TextBox31.Text = ROW10("TARGET_12").ToString()
            TextBox19.Text = ROW10("TARGET_24").ToString()
        End If
        ROW11 = coninfo.ExecuteDQL(QRY11, con_11)
        If ROW11.Read() Then
            TextBox38.Text = ROW11("AVARAGE").ToString()
            TextBox42.Text = ROW11("TARGET_12").ToString()
            TextBox30.Text = ROW11("TARGET_24").ToString()
        End If
        ROW12 = coninfo.ExecuteDQL(QRY12, con_12)
        If ROW12.Read() Then
            TextBox37.Text = ROW12("AVARAGE").ToString()
            TextBox41.Text = ROW12("TARGET_12").ToString()
            TextBox29.Text = ROW12("TARGET_24").ToString()
        End If
        ROW13 = coninfo.ExecuteDQL(QRY13, con_13)
        If ROW13.Read() Then
            TextBox36.Text = ROW13("AVARAGE").ToString()
            TextBox40.Text = ROW13("TARGET_12").ToString()
            TextBox24.Text = ROW13("TARGET_24").ToString()
        End If
        ROW14 = coninfo.ExecuteDQL(QRY14, con_14)
        If ROW14.Read() Then
            TextBox35.Text = ROW14("AVARAGE").ToString()
            TextBox39.Text = ROW14("TARGET_12").ToString()
            TextBox23.Text = ROW14("TARGET_24").ToString()
        End If
        ROW15 = coninfo.ExecuteDQL(QRY15, con_15)
        If ROW15.Read() Then
            TextBox46.Text = ROW15("AVARAGE").ToString()
            TextBox48.Text = ROW15("TARGET_12").ToString()
            TextBox44.Text = ROW15("TARGET_24").ToString()
        End If
        ROW16 = coninfo.ExecuteDQL(QRY16, con_16)
        If ROW16.Read() Then
            TextBox45.Text = ROW16("AVARAGE").ToString()
            TextBox47.Text = ROW16("TARGET_12").ToString()
            TextBox43.Text = ROW16("TARGET_24").ToString()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim QRY1, QRY2, QRY3, QRY4, QRY5, QRY6, QRY7, QRY8, QRY9, QRY10, QRY11, QRY12, QRY13, QRY14, QRY15, QRY16 As String
        Dim ROW1, ROW2, ROW3, ROW4, ROW5, ROW6, ROW7, ROW8, ROW9, ROW10, ROW11, ROW12, ROW13, ROW14, ROW15, ROW16 As Integer
        QRY1 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '4' AND FUEL_NOZZLE = 'ULP-1'"
        QRY2 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '3' AND FUEL_NOZZLE = 'ULP-2'"
        QRY3 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '6' AND FUEL_NOZZLE = 'ULP-3'"
        QRY4 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '5' AND FUEL_NOZZLE = 'ULP-4'"
        QRY5 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '9' AND FUEL_NOZZLE = 'ULP-5,ULP-7'"
        QRY6 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '10' AND FUEL_NOZZLE = 'ULP-8,ULP-6'"

        QRY7 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '7' AND FUEL_NOZZLE = 'XP-4'"
        QRY8 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '8' AND FUEL_NOZZLE = 'XP-3'"
        QRY9 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '11,12' AND FUEL_NOZZLE = 'XP-2'"
        QRY10 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = '13,14' AND FUEL_NOZZLE = 'XP-1'"

        QRY11 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = 'CNG-1' AND FUEL_NOZZLE = 'CNG-1'"
        QRY12 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = 'CNG_1' AND FUEL_NOZZLE = 'CNG-2'"
        QRY13 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = 'CNG-2' AND FUEL_NOZZLE = 'CNG-3'"
        QRY14 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = 'CNG_2' AND FUEL_NOZZLE = 'CNG-4'"
        QRY15 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = 'CNG-3' AND FUEL_NOZZLE = 'CNG-5'"
        QRY16 = "UPDATE TARGET SET AVARAGE = @AVG  ,TARGET_12 = @12 ,TARGET_24 = @24 WHERE MACHINE_NO = 'CNG_3' AND FUEL_NOZZLE = 'CNG-6'"
        Dim CMD1 As New SqlCommand(QRY1, con1)
        CMD1.Parameters.AddWithValue("@AVG", TextBox12.Text)
        CMD1.Parameters.AddWithValue("@12", TextBox1.Text)
        CMD1.Parameters.AddWithValue("@24", TextBox18.Text)
        con1.Open()
        ROW1 = CMD1.ExecuteNonQuery()
        con1.Close()

        If ROW1 > 0 Then
            Dim CMD2 As New SqlCommand(QRY2, con2)
            CMD2.Parameters.AddWithValue("@AVG", TextBox11.Text)
            CMD2.Parameters.AddWithValue("@12", TextBox2.Text)
            CMD2.Parameters.AddWithValue("@24", TextBox17.Text)
            con2.Open()
            ROW2 = CMD2.ExecuteNonQuery()
            If ROW2 > 0 Then
                Dim CMD3 As New SqlCommand(QRY3, con3)
                CMD3.Parameters.AddWithValue("@AVG", TextBox10.Text)
                CMD3.Parameters.AddWithValue("@12", TextBox4.Text)
                CMD3.Parameters.AddWithValue("@24", TextBox16.Text)
                con3.Open()
                ROW3 = CMD3.ExecuteNonQuery()
                If ROW3 > 0 Then
                    Dim CMD4 As New SqlCommand(QRY4, con4)
                    CMD4.Parameters.AddWithValue("@AVG", TextBox9.Text)
                    CMD4.Parameters.AddWithValue("@12", TextBox3.Text)
                    CMD4.Parameters.AddWithValue("@24", TextBox15.Text)
                    con4.Open()
                    ROW4 = CMD4.ExecuteNonQuery()
                    If ROW4 > 0 Then
                        Dim CMD5 As New SqlCommand(QRY5, con5)
                        CMD5.Parameters.AddWithValue("@AVG", TextBox6.Text)
                        CMD5.Parameters.AddWithValue("@12", TextBox8.Text)
                        CMD5.Parameters.AddWithValue("@24", TextBox14.Text)
                        con5.Open()
                        ROW5 = CMD5.ExecuteNonQuery()
                        If ROW5 > 0 Then
                            Dim CMD6 As New SqlCommand(QRY6, con6)
                            CMD6.Parameters.AddWithValue("@AVG", TextBox5.Text)
                            CMD6.Parameters.AddWithValue("@12", TextBox7.Text)
                            CMD6.Parameters.AddWithValue("@24", TextBox13.Text)
                            con6.Open()
                            ROW6 = CMD6.ExecuteNonQuery()
                            If ROW6 > 0 Then
                                Dim CMD7 As New SqlCommand(QRY7, con7)
                                CMD7.Parameters.AddWithValue("@AVG", TextBox28.Text)
                                CMD7.Parameters.AddWithValue("@12", TextBox34.Text)
                                CMD7.Parameters.AddWithValue("@24", TextBox22.Text)
                                con7.Open()
                                ROW7 = CMD7.ExecuteNonQuery()
                                If ROW7 > 0 Then
                                    Dim CMD8 As New SqlCommand(QRY8, con8)
                                    CMD8.Parameters.AddWithValue("@AVG", TextBox27.Text)
                                    CMD8.Parameters.AddWithValue("@12", TextBox33.Text)
                                    CMD8.Parameters.AddWithValue("@24", TextBox21.Text)
                                    con8.Open()
                                    ROW8 = CMD8.ExecuteNonQuery()
                                    If ROW8 > 0 Then
                                        Dim CMD9 As New SqlCommand(QRY9, con9)
                                        CMD9.Parameters.AddWithValue("@AVG", TextBox26.Text)
                                        CMD9.Parameters.AddWithValue("@12", TextBox32.Text)
                                        CMD9.Parameters.AddWithValue("@24", TextBox20.Text)
                                        con9.Open()
                                        ROW9 = CMD9.ExecuteNonQuery()
                                        If ROW9 > 0 Then
                                            Dim CMD10 As New SqlCommand(QRY10, con10)
                                            CMD10.Parameters.AddWithValue("@AVG", TextBox25.Text)
                                            CMD10.Parameters.AddWithValue("@12", TextBox31.Text)
                                            CMD10.Parameters.AddWithValue("@24", TextBox19.Text)
                                            con10.Open()
                                            ROW10 = CMD10.ExecuteNonQuery()
                                            If ROW10 > 0 Then
                                                Dim CMD11 As New SqlCommand(QRY11, con11)
                                                CMD11.Parameters.AddWithValue("@AVG", TextBox38.Text)
                                                CMD11.Parameters.AddWithValue("@12", TextBox42.Text)
                                                CMD11.Parameters.AddWithValue("@24", TextBox30.Text)
                                                con11.Open()
                                                ROW11 = CMD11.ExecuteNonQuery()
                                                If ROW11 > 0 Then
                                                    Dim CMD12 As New SqlCommand(QRY12, con12)
                                                    CMD12.Parameters.AddWithValue("@AVG", TextBox37.Text)
                                                    CMD12.Parameters.AddWithValue("@12", TextBox41.Text)
                                                    CMD12.Parameters.AddWithValue("@24", TextBox29.Text)
                                                    con12.Open()
                                                    ROW12 = CMD12.ExecuteNonQuery()
                                                    If ROW12 > 0 Then
                                                        Dim CMD13 As New SqlCommand(QRY13, con13)
                                                        CMD13.Parameters.AddWithValue("@AVG", TextBox36.Text)
                                                        CMD13.Parameters.AddWithValue("@12", TextBox40.Text)
                                                        CMD13.Parameters.AddWithValue("@24", TextBox24.Text)
                                                        con13.Open()
                                                        ROW13 = CMD13.ExecuteNonQuery()
                                                        If ROW13 > 0 Then
                                                            Dim CMD14 As New SqlCommand(QRY14, con14)
                                                            CMD14.Parameters.AddWithValue("@AVG", TextBox35.Text)
                                                            CMD14.Parameters.AddWithValue("@12", TextBox39.Text)
                                                            CMD14.Parameters.AddWithValue("@24", TextBox23.Text)
                                                            con14.Open()
                                                            ROW14 = CMD14.ExecuteNonQuery()
                                                            If ROW14 > 0 Then
                                                                Dim CMD15 As New SqlCommand(QRY15, con15)
                                                                CMD15.Parameters.AddWithValue("@AVG", TextBox46.Text)
                                                                CMD15.Parameters.AddWithValue("@12", TextBox48.Text)
                                                                CMD15.Parameters.AddWithValue("@24", TextBox44.Text)
                                                                con15.Open()
                                                                ROW15 = CMD15.ExecuteNonQuery()
                                                                If ROW15 > 0 Then
                                                                    Dim CMD16 As New SqlCommand(QRY16, con16)
                                                                    CMD16.Parameters.AddWithValue("@AVG", TextBox45.Text)
                                                                    CMD16.Parameters.AddWithValue("@12", TextBox47.Text)
                                                                    CMD16.Parameters.AddWithValue("@24", TextBox43.Text)
                                                                    con16.Open()
                                                                    ROW16 = CMD16.ExecuteNonQuery()
                                                                    If ROW16 > 0 Then
                                                                        MsgBox("UPDATED SUCCESSFULLY")
                                                                        Me.Close()
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Else
            MsgBox("NOT UPDATED")
        End If
        con1.Close()
        con2.Close()
        con3.Close()
        con4.Close()
        con5.Close()
        con6.Close()
        con7.Close()
        con8.Close()
        con9.Close()
        con10.Close()
        con11.Close()
        con12.Close()
        con13.Close()
        con14.Close()
        con15.Close()
        con16.Close()
    End Sub
End Class