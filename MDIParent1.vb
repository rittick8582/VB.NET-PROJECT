Imports System.Windows.Forms

Public Class MDIParent1
    Dim i, f, a, j As Integer



    

 

    Private Sub INSERTToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles INSERTToolStripMenuItem.Click

        If f = 0 Then
            INSERTToolStripMenuItem.BackColor = Color.ForestGreen
            f = 1
        Else
            INSERTToolStripMenuItem.BackColor = Color.GreenYellow
            f = 0
        End If

        Form2.Show()
        Form2.MdiParent = Me
    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SEARCHToolStripMenuItem.Click

        If i = 0 Then
            SEARCHToolStripMenuItem.BackColor = Color.PowderBlue
            i = 1
        Else
            SEARCHToolStripMenuItem.BackColor = Color.BlueViolet
            i = 0
        End If
        Form3.Show()
        Form3.MdiParent = Me
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click

        If a = 0 Then
            UPDATEToolStripMenuItem.BackColor = Color.BlanchedAlmond
            a = 1
        Else
            UPDATEToolStripMenuItem.BackColor = Color.Brown
            a = 0
        End If
        Form4.Show()
        Form4.MdiParent = Me
    End Sub

    Private Sub DELETEDToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DELETEDToolStripMenuItem.Click

        If j = 0 Then
            DELETEDToolStripMenuItem.BackColor = Color.Red
            j = 1
        Else
            DELETEDToolStripMenuItem.BackColor = Color.Purple
            j = 0
        End If

        Form5.Show()
        Form5.MdiParent = Me
    End Sub
End Class
