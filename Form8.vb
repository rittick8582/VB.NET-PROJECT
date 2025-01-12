Imports System.Data.OleDb
Imports System.Data
Public Class Form8
    Dim con As OleDbConnection
    Dim com As OleDbCommand
    Dim adpt As OleDbDataAdapter
    Dim ds As DataSet
    Dim sql As String
    Sub connect()
        con = New OleDbConnection("provider=Microsoft.JET.OLEDB.4.0; data source=stock.mdb")
        con.Open()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        connect()
        sql = "select * from sale where pid='" & TextBox1.Text & "'"
        adpt = New OleDbDataAdapter(sql, con)
        ds = New DataSet
        adpt.Fill(ds)
        If ds.Tables(0).Rows.Count > 0 Then
            TextBox6.Text = ds.Tables(0).Rows(0)(0).ToString()
            TextBox2.Text = ds.Tables(0).Rows(0)(2).ToString()
        Else
            MsgBox("Record Not Inserted")
        End If
        DateTimePicker1.Value = ds.Tables(0).Rows(0)(3).ToString()
        TextBox3.Text = ds.Tables(0).Rows(0)(4).ToString()
        TextBox4.Text = ds.Tables(0).Rows(0)(5).ToString()
        TextBox5.Text = ds.Tables(0).Rows(0)(6).ToString()
        TextBox1.Enabled = False
        Button3.Enabled = True
        Button1.Enabled = False    

        con.Close()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        connect()
        sql = "update sale set sid='" & TextBox6.Text & "',  pname='" & TextBox2.Text & "', sellingdate='" & DateTimePicker1.Value & "', sellingprice='" & TextBox3.Text & "', quantity='" & TextBox4.Text & "', total='" & TextBox5.Text & "' where pid='" & TextBox1.Text & "'"
        com = New OleDbCommand(sql, con)
        If com.ExecuteNonQuery Then
            MsgBox("Record Succesfully Updated")
            TextBox1.Enabled = True
        Else
            MsgBox("Record Not Updated")
        End If
        cls()
        con.Close()
    End Sub

    Sub cls()
        TextBox1.Clear()
        TextBox2.Clear()
        DateTimePicker1.Refresh()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox6.Clear()
        Button1.Enabled = True
        Button3.Enabled = False
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cls()
        TextBox1.Enabled = True
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
End Class