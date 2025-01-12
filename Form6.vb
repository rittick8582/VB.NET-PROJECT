Imports System.Data
Imports System.Data.OleDb

Public Class Form6
    Dim con As OleDbConnection
    Dim com As OleDbCommand
    Dim adpt As OleDbDataAdapter
    Dim ds As DataSet
    Dim sql As String

    Sub connect()
        con = New OleDbConnection("Provider=Microsoft.JET.OLEDB.4.0;Data Source=stock.mdb")
        con.Open()
    End Sub

    Private Sub Form6_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' You can add any initialization code here if needed
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        connect()
        sql = "SELECT * FROM product WHERE pid='" & TextBox1.Text & "'"
        adpt = New OleDbDataAdapter(sql, con)
        ds = New DataSet
        adpt.Fill(ds)

        If ds.Tables(0).Rows.Count > 0 Then
            TextBox2.Text = ds.Tables(0).Rows(0)(1).ToString()
        Else
            MsgBox("Record Not Found")
        End If
        'TextBox3.Text = ds.Tables(0).Rows(0)(4).ToString
        'TextBox4.Text = ds.Tables(0).Rows(0)(3).ToString
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        Button3.Enabled = False
        Button1.Enabled = True

        con.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        connect()
        sql = "insert into sale values('" & TextBox6.Text & "', '" & TextBox1.Text & "','" & TextBox2.Text & "','" & DateTimePicker1.Value & "','" & TextBox3.Text & "','" & TextBox4.Text & "','" & TextBox5.Text & "')"
        com = New OleDbCommand(sql, con)
        If com.ExecuteNonQuery Then
            MsgBox("Record Inserted Succesfully")
        Else
            MsgBox("Record Not Inserted")
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
        TextBox6.Clear()
        DateTimePicker1.Value = DateTime.Now
        TextBox1.Enabled = True
        Button1.Enabled = False
        Button3.Enabled = True
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cls()
        Button3.Enabled = True
        TextBox1.Enabled = True
    End Sub
   
End Class
