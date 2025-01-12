Imports System.Data.OleDb

Public Class Form2
    Dim con As OleDbConnection
    Dim com As OleDbCommand
    Dim sql As String

    Sub connect()
        con = New OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;Data Source=stock.mdb")
        con.Open()
    End Sub

    Private Sub Form2_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' Add any necessary initialization code here
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        connect()
        sql = "INSERT INTO product (pid, pname, phsdate, quantity, rate, total) VALUES (?, ?, ?, ?, ?, ?)"
        com = New OleDbCommand(sql, con)

        ' Use parameters to prevent SQL injection and handle date format conversion
        com.Parameters.AddWithValue("@pid", TextBox1.Text)
        com.Parameters.AddWithValue("@pname", TextBox2.Text)
        com.Parameters.AddWithValue("@phsdate", DateTimePicker1.Value.ToString("MM/dd/yyyy"))
        com.Parameters.AddWithValue("@quantity", TextBox3.Text)
        com.Parameters.AddWithValue("@rate", TextBox4.Text)
        com.Parameters.AddWithValue("@total", TextBox5.Text)

        If com.ExecuteNonQuery() > 0 Then
            MsgBox("Record inserted successfully")
        Else
            MsgBox("Record not inserted")
        End If

        cls()
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

    Sub cls()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        DateTimePicker1.Value = DateTime.Now
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cls()
    End Sub
End Class
