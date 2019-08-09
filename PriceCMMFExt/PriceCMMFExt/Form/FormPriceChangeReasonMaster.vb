Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormPriceChangeReasonMaster
    Private Property sqlstr As String
    Private DS As DataSet
    Private BS As BindingSource
    Private CM As CurrencyManager

    Private Sub FormCutoff_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If Not IsNothing(getChanges) Then
            Select Case MessageBox.Show("Save unsaved records?", "Unsaved records", MessageBoxButtons.YesNoCancel)
                Case Windows.Forms.DialogResult.Yes
                    ToolStripButton4.PerformClick()
                Case Windows.Forms.DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Function getChanges() As DataSet
        If IsNothing(BS) Then
            Return Nothing
        End If
        Me.Validate()
        BS.EndEdit()
        DS.EnforceConstraints = False
        Return DS.GetChanges()
    End Function

    Private Sub FormCutoff_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Public Sub LoadData()      
        DS = New DataSet
        Dim mymessage As String = ""
        sqlstr = " select * from pricechangereason U order by lineno;"


        If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            DS.Tables(0).TableName = "Reason"

            Dim idx(0) As DataColumn
            idx(0) = DS.Tables(0).Columns("id")
            DS.Tables(0).PrimaryKey = idx

            DS.Tables(0).Columns("id").AutoIncrement = True
            DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            DS.Tables(0).Columns("id").AutoIncrementStep = -1

            'Binding Object

            BS = New BindingSource



            BS.DataSource = DS.Tables(0)

            'BS.Sort = "ordernum asc"
            DataGridView1.AutoGenerateColumns = False
            DataGridView1.DataSource = BS
            CM = CType(Me.BindingContext(BS), CurrencyManager)
        End If
    End Sub
    'Add
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not IsNothing(BS) Then
            BS.Sort = ""
        End If
        Dim drv As DataRowView = BS.AddNew()
        Dim dr = drv.Row
        dr.Item("isactive") = True
        DS.Tables(0).Rows.Add(dr)
        Dim myform = New FormInputReason(BS)
        If Not myform.ShowDialog() = Windows.Forms.DialogResult.OK Then
            DS.Tables(0).Rows.Remove(dr)
        End If
        Me.DataGridView1.Invalidate()
    End Sub
    'Update
    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        If Not IsNothing(BS.Current) Then
            Dim myform = New FormInputReason(BS)
            If Not myform.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                BS.CancelEdit()
            End If
            myform.Dispose()
        Else
            MessageBox.Show("No record to update.")
        End If

        Me.DataGridView1.Invalidate()
    End Sub
    'Delete
    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(BS.Current) Then
            If MessageBox.Show("Delete this record(s)", "Delete Record", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.OK Then
                For Each dsrow In DataGridView1.SelectedRows
                    BS.RemoveAt(CType(dsrow, DataGridViewRow).Index)
                Next
            End If
        Else
            MessageBox.Show("No record to delete.")
        End If
    End Sub
    'save
    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click

        Dim myposition = CM.Position
        Me.Validate()
        BS.EndEdit()
        DS.EnforceConstraints = False
        Dim ds2 = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer

            Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
            If DbAdapter1.PriceChangeReasonTx(Me, mye) Then
                'Dim myquery = From row As DataRow In DS.Tables(0).Rows
                '              Where row.RowState = DataRowState.Added

                'For Each rec In myquery.ToArray
                '    rec.Delete()
                'Next
                DS.Merge(ds2)
                DS.AcceptChanges()

                BS.Position = myposition
                MessageBox.Show("Saved!")

                'LoadData()

            Else
                MessageBox.Show(mye.message)
            End If
        Else
            MessageBox.Show("Nothing to save.")
        End If
        Me.DataGridView1.Invalidate()
    End Sub

    Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
        ToolStripButton5.PerformClick()
    End Sub

    Private Sub ToolStripButton6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        LoadData()
    End Sub
End Class