Imports System.Windows.Forms

Public Class MDIParent2
    Dim f, i, j, a As Integer

    Private Sub INSERTToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles INSERTToolStripMenuItem.Click

        If f = 0 Then
            INSERTToolStripMenuItem.BackColor = Color.ForestGreen
            f = 1
        Else
            INSERTToolStripMenuItem.BackColor = Color.LightGreen
            f = 0
        End If

        Form6.Show()
        Form6.MdiParent = Me
    End Sub

    Private Sub SEARCHToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SEARCHToolStripMenuItem.Click
        If i = 0 Then
            SEARCHToolStripMenuItem.BackColor = Color.BlueViolet
            i = 1
        Else
            SEARCHToolStripMenuItem.BackColor = Color.SkyBlue
            i = 0

        End If
        Form7.Show()
        Form7.MdiParent = Me
    End Sub

    Private Sub UPDATEToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles UPDATEToolStripMenuItem.Click

        If a = 0 Then
            UPDATEToolStripMenuItem.BackColor = Color.RosyBrown
            a = 1
        Else
            UPDATEToolStripMenuItem.BackColor = Color.BurlyWood
            a = 0

        End If
        Form8.Show()
        Form8.MdiParent = Me
    End Sub

    Private Sub DELETEDToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DELETEDToolStripMenuItem.Click
        If j = 0 Then
            DELETEDToolStripMenuItem.BackColor = Color.Red
            j = 1
        Else
            DELETEDToolStripMenuItem.BackColor = Color.MistyRose
            j = 0

        End If
        Form9.Show()
        Form9.MdiParent = Me
    End Sub
End Class
