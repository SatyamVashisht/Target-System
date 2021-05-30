Imports System.Data.SqlClient
Public Class Connectioninfo
    Public ConnectionString As String = "Data Source=.\SQLExpress;Integrated Security=true;AttachDbFilename=C:\Users\HP\Desktop\TARGET_SYSTEM\TARGET_SYSTEM\TARGET.mdf;User Instance=true;"
    Sub BindCombo(ByRef combo As ComboBox, ByVal qry As String, ByVal column As String)

        Dim con As New SqlConnection(ConnectionString)
        Dim da As New SqlDataAdapter(qry, con)
        Dim ds As New DataSet
        da.Fill(ds)

        combo.DataSource = ds.Tables(0)
        combo.ValueMember = column
        combo.DisplayMember = column
    End Sub

    Function GetTarget(ByVal qry As String) As String()

        Dim con As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand(qry, con)
        Dim dr As SqlDataReader
        Dim target As Integer = 0
        Dim fuel As String = ""

        Dim arr(1) As String

        If con.State = ConnectionState.Open Then
            con.Close()
        End If
        con.Open()

        dr = cmd.ExecuteReader()
        If dr.Read() Then
            arr(0) = dr(0).ToString()
            arr(1) = dr(1).ToString()

        End If

        con.Close()

        Return arr

    End Function



    Function ExecuteDQL(ByVal qry As String, ByRef con As SqlConnection) As SqlDataReader

        Dim cmd As New SqlCommand(qry, con)
        Dim dr As SqlDataReader

        If con.State = ConnectionState.Open Then
            con.Close()
        End If
        con.Open()

        dr = cmd.ExecuteReader()

        Return dr
        con.Close()
    End Function

    Function ExecuteScaler(ByVal qry As String) As String

        Dim con As New SqlConnection(ConnectionString)
        Dim cmd As New SqlCommand(qry, con)
        Dim val As String = ""

        If con.State = ConnectionState.Open Then
            con.Close()
        End If
        con.Open()
        Try
            val = cmd.ExecuteScalar()
        Catch ex As Exception
            MsgBox("NO DATA FOUND")

        End Try

        con.Close()


        Return val

    End Function
     Sub BindGrid(ByRef grid As DataGridView, ByVal qry As String)

        Dim con As New SqlConnection(ConnectionString)
        Dim da As New SqlDataAdapter(qry, con)
        Dim ds As New DataSet
        da.Fill(ds)

        grid.DataSource = ds.Tables(0)

    End Sub
End Class
