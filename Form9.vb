Imports System.Data
Imports System.Data.OleDb


Public Class Form9
    Dim con As OleDbConnection
    Dim com As OleDbCommand
    Dim sql As String
    Sub connect()
        con = New OleDbConnection("provider=Microsoft.JET.OLEDB.4.0; data source=stock.mdb")
        con.Open()
    End Sub
  
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        connect()
        sql = " delete from sale where pid='" & TextBox1.Text & "'"
        com = New OleDbCommand(sql, con)
        If com.ExecuteNonQuery Then
            MsgBox("Record Deleted")
        Else
            MsgBox("Record Not Deleted")
        End If
        cls()
        con.Close()
    End Sub
    Sub cls()
        TextBox1.Clear()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cls()
    End Sub
End Class