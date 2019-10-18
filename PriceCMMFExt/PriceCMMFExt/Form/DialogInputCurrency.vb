Imports System.Windows.Forms

Public Class DialogInputCurrency
    Dim DRV As DataRowView

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        DRV.EndEdit()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        DRV.CancelEdit()
        Me.Close()
    End Sub

    Private Sub DialogInputCurrency_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        BindingObject()
    End Sub


    Private Sub BindingObject()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()

        TextBox1.DataBindings.Add("Text", DRV, "myyear", True, DataSourceUpdateMode.OnPropertyChanged)
        TextBox2.DataBindings.Add("Text", DRV, "crcy", True, DataSourceUpdateMode.OnPropertyChanged)
        TextBox3.DataBindings.Add("Text", DRV, "currency", True, DataSourceUpdateMode.OnPropertyChanged)
        TextBox4.DataBindings.Add("Text", DRV, "budgetcurrency", True, DataSourceUpdateMode.OnPropertyChanged)

    End Sub

    Public Sub New(ByVal DRV As DataRowView)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = DRV
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class
