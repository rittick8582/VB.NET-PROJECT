Imports System.Data.OleDb

Public Class Form4
    Dim con As OleDbConnection
    Dim com As OleDbCommand
    Dim adpt As OleDbDataAdapter
    Dim ds As DataSet
    Dim sql As String

    Sub Connect()
        con = New OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0; Data Source=stock.mdb")
        con.Open()
    End Sub

    Private Sub Form4_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Add initialization code here if needed
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Connect()
        sql = "SELECT * FROM product WHERE pid=@pid"
        adpt = New OleDbDataAdapter(sql, con)
        adpt.SelectCommand.Parameters.AddWithValue("@pid", TextBox1.Text)
        ds = New DataSet
        adpt.Fill(ds)
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            TextBox2.Text = ds.Tables(0).Rows(0)(1).ToString()

            ' Parse the date value into DateTimePicker
            Dim dateValue As DateTime
            If DateTime.TryParse(ds.Tables(0).Rows(0)(2).ToString(), dateValue) Then
                DateTimePicker1.Value = dateValue
            Else
                ' Handle parsing error, if necessary
            End If

            TextBox3.Text = ds.Tables(0).Rows(0)(3).ToString()
            TextBox4.Text = ds.Tables(0).Rows(0)(4).ToString()
            TextBox5.Text = ds.Tables(0).Rows(0)(5).ToString()
            TextBox1.Enabled = False
            Button3.Enabled = True
            Button1.Enabled = False
        Else
            MsgBox("Record Not Found")
        End If
        con.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Connect()
        sql = "UPDATE product SET pname=@pname, phsdate=@phsdate, quantity=@quantity, rate=@rate, total=@total WHERE pid=@pid"
        com = New OleDbCommand(sql, con)
        com.Parameters.AddWithValue("@pname", TextBox2.Text)
        com.Parameters.AddWithValue("@phsdate", DateTimePicker1.Value.ToString("yyyy-MM-dd"))
        com.Parameters.AddWithValue("@quantity", TextBox3.Text)
        com.Parameters.AddWithValue("@rate", TextBox4.Text)
        com.Parameters.AddWithValue("@total", TextBox5.Text)
        com.Parameters.AddWithValue("@pid", TextBox1.Text)

        If com.ExecuteNonQuery() > 0 Then
            MsgBox("Record Updated")
            TextBox1.Enabled = True
        Else
            MsgBox("Record Not Updated")
        End If

        ' Dispose of resources using Using statement
        com.Dispose()
        con.Close()
    End Sub

    Private Sub TextBox3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox3.TextChanged
        CalculateTotal()
    End Sub

    Private Sub TextBox4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox4.TextChanged
        CalculateTotal()
    End Sub

    Private Sub CalculateTotal()
        If Not String.IsNullOrEmpty(TextBox3.Text) AndAlso Not String.IsNullOrEmpty(TextBox4.Text) Then
            Dim num21 As Double = Double.Parse(TextBox3.Text)
            Dim num22 As Double = Double.Parse(TextBox4.Text)
            Dim ans As Double = num21 * num22
            TextBox5.Text = FormatNumber(ans)
        End If
    End Sub

    Sub ClearFields()
        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Value = DateTime.Now ' Reset DateTimePicker to current date and time
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        Button1.Enabled = True
        Button3.Enabled = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ClearFields()
    End Sub

End Class
