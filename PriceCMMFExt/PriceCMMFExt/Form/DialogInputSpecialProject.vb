Imports System.Windows.Forms

Public Class DialogInputSpecialProject
    Private bs As New BindingSource
    Dim myrow As DataRowView

    Public Sub New(ByRef bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.bs = bs
        Me.myrow = bs.Current

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click        
        If Not Me.Validate() Then
            Me.DialogResult = DialogResult.None
            bs.CancelEdit()
            Exit Sub
        End If
        bs.EndEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        bs.CancelEdit()
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()    
    End Sub

    Public Overloads Function validate() As Boolean
        MyBase.Validate()
        If TextBox2.Text = "" Then
            ErrorProvider1.SetError(TextBox2, "Value cannot be blank.")
            Return False
        End If
        Return True
    End Function

    Private Sub DialogInputSpecialProject_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        TextBox2.DataBindings.Clear()
        CheckBox1.DataBindings.Clear()
        TextBox1.DataBindings.Clear()

        TextBox1.DataBindings.Add("Text", myrow, "lineno")
        TextBox2.DataBindings.Add("Text", myrow, "specialproject")
        CheckBox1.DataBindings.Add(New Binding("checked", myrow, "isactive", False, DataSourceUpdateMode.OnPropertyChanged))

    End Sub
End Class
