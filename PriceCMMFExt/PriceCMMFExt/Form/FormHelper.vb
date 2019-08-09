Public Class FormHelper

    Dim bs As BindingSource


    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.bs = bs 'New BindingSource(bs.DataSource, bs.DataMember)
        Me.bs.Filter = ""
    End Sub

    'Private Sub FormHelper_FormClosed(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    '    Dim drv As DataRowView = bs.Current
    '    MessageBox.Show("hello")
    'End Sub

    Private Sub FormGetValidator_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs
    End Sub


    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Button1.PerformClick()
    End Sub

    Private Sub TextBox1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox1.TextChanged
        Try
            bs.Filter = String.Format("[" & DataGridView1.Columns(0).DataPropertyName & "] like '*{0}*'", TextBox1.Text.Replace("'", "''"))
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        bs.Filter = ""
    End Sub
End Class