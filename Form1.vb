Imports System.Data
Imports System.Data.OleDb

Public Class Form1
    Dim con As OleDbConnection
    Dim adpt As OleDbDataAdapter
    Dim es As DataSet
    Dim sql As String
    Dim c As Integer = 0

    Sub conn()
        con = New OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=cre.mdb")
        con.Open()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        conn()
        If c < 2 Then
            sql = "select * from log where uid='" & TextBox1.Text & "' and upas='" & TextBox2.Text & "' and utype='" & ComboBox1.Text & "'"

            adpt = New OleDbDataAdapter(sql, con)
            es = New DataSet()

            Try
                adpt.Fill(es)

                If es.Tables(0).Rows.Count > 0 Then
                    Timer1.Start()
                    ProgressBar1.Visible = True
                Else
                    c += 1
                    MsgBox("Invalid Login Credential")
                End If
            Catch ex As Exception
                MsgBox("Error: " & ex.Message)
            End Try
        Else
            MsgBox("Too Many Attempts")
            Me.Close()
        End If
        con.Close()
        cls()
    End Sub
    Sub cls()
        TextBox1.Clear()
        TextBox2.Clear()
        ComboBox1.ResetText()
    End Sub
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ProgressBar1.Visible = False
        Label5.Text = ""
        With ProgressBar1
            .Minimum = 0
            .Maximum = 105
        End With
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        ProgressBar1.Value += 5
        Label5.Text = ProgressBar1.Value & "% Loading...."
        If ProgressBar1.Value = 100 Then
            Timer1.Stop()
            MsgBox("Welcome")
            MDIParent1.Show()
            MDIParent2.Show()
            Me.Hide()
            ProgressBar1.Value = 0
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        cls()
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            TextBox2.UseSystemPasswordChar = True
        Else
            TextBox2.UseSystemPasswordChar = False
        End If
    End Sub
End Class
