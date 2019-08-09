Public Class FormGetValidator
    Dim bs As BindingSource


    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.bs = bs

    End Sub

    Private Sub FormGetValidator_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        DataGridView1.AutoGenerateColumns = False
        DataGridView1.DataSource = bs
    End Sub


    Private Sub DataGridView1_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentDoubleClick
        Button1.PerformClick()
    End Sub
End Class