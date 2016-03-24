Imports System.Windows.Forms
Imports System.Data.SQLite
Imports System.Data


Public Class SearchDialog

    Public xmbh As String

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        xmbh = ""
        Try
            Dim conn As New SQLiteConnection("Data Source=c:\items.sqlite;Pooling=true;FailIfMissing=false")
            Dim cmd As New SQLiteCommand("select xmbh,kcbh,kcmc,xmmc from items where xmmc='" & TextBox1.Text.Trim() & "'", conn)
            Dim da As New SQLiteDataAdapter(cmd)
            Dim dt As New DataTable()
            da.Fill(dt)
            DataGridView1.DataSource = dt
            conn.Close()
        Catch ex As Exception
            MessageBox.Show(ex.ToString())
        End Try
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        xmbh = DataGridView1.CurrentCell.Value.ToString()
    End Sub
End Class
