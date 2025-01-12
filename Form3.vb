Imports System.Data
Imports System.Data.OleDb


Public Class Form3
    Dim con As OleDbConnection
    Dim adpt As OleDbDataAdapter
    Dim ds As DataSet
    Dim sql As String
    Sub connect()
        con = New OleDbConnection("provider=Microsoft.JET.OLEDB.4.0; data source=stock.mdb")
        con.Open()
    End Sub
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        connect()
        sql = "select* from product where pid='" & TextBox1.Text & "'"
        adpt = New OleDbDataAdapter(sql, con)
        ds = New DataSet
        adpt.Fill(ds, "DM")
        If ds.Tables("DM").Rows.Count > 0 Then
            DataGridView1.DataSource = ds.Tables("DM")
        Else
            MsgBox("Record Not Found")
        End If
        con.Close()
    End Sub
    Private Sub TextBox1_MouseLeave(sender As System.Object, e As System.EventArgs) Handles TextBox1.MouseLeave
        connect()
        sql = "select* from product where pid='" & TextBox1.Text & "'"
        adpt = New OleDbDataAdapter(sql, con)
        ds = New DataSet
        adpt.Fill(ds, "DM")
        If ds.Tables("DM").Rows.Count > 0 Then
            DataGridView1.DataSource = ds.Tables("DM")
        Else
            MsgBox("Record Not Found")
        End If
        con.Close()
    End Sub


    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click
        DataGridView1.DataSource = Nothing
        TextBox1.Clear()
    End Sub

   
End Class