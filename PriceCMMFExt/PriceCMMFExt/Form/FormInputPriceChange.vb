Public Class FormInputPriceChange
    Private myrow As DataRowView
    Private BS As BindingSource
    Dim WithEvents oBindingNumeric1 As Binding


    Public Sub New(ByVal bs As BindingSource)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.BS = bs

        oBindingNumeric1 = New Binding("Text", bs, "integer")

        'myrow = bs.Current
    End Sub
    Private Sub FormInputPriceChange_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub obinding_Format(ByVal sender As Object, ByVal e As System.Windows.Forms.ConvertEventArgs) Handles oBindingNumeric1.Format
        If Not IsDBNull(e.Value) Then
            Select Case CType(sender, System.Windows.Forms.Binding).BindingMemberInfo.BindingField
                Case "integer"
                    e.Value = Format(e.Value, "#,##0")
                Case "numeric"
                    e.Value = Format(e.Value, "#,##0.00")
            End Select
        End If
    End Sub
    Private Sub TextBox1_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            Dim obj = CType(sender, TextBox)

            If obj.Tag = "String" Then
                If obj.Text = "" Then
                    ErrorProvider1.SetError(obj, "Value cannot be empty.")
                    Button1.DialogResult = Windows.Forms.DialogResult.None
                    e.Cancel = True
                End If
            ElseIf obj.Tag = "Number" Then
                If Not IsNumeric(obj.Text) Then
                    ErrorProvider1.SetError(obj, "Please enter numeric value.")
                    Button1.DialogResult = Windows.Forms.DialogResult.None
                    e.Cancel = True
                End If
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub
    Private Sub TextBox_validate(ByVal sender As Object, ByVal e As System.EventArgs)
        ErrorProvider1.SetError(sender, "")
        Button1.DialogResult = Windows.Forms.DialogResult.OK
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim e1 = New System.ComponentModel.CancelEventArgs
        Dim e2 = New System.ComponentModel.CancelEventArgs
        Dim e3 = New System.ComponentModel.CancelEventArgs
        Dim e4 = New System.ComponentModel.CancelEventArgs
        Dim e5 = New System.ComponentModel.CancelEventArgs
        Dim e6 = New System.ComponentModel.CancelEventArgs
        Dim e7 = New System.ComponentModel.CancelEventArgs
        TextBox1_Validating(TextBox1, e1)
        TextBox1_Validating(TextBox2, e2)
        TextBox1_Validating(TextBox3, e3)
        TextBox1_Validating(TextBox4, e4)
        TextBox1_Validating(TextBox5, e5)
        TextBox1_Validating(TextBox6, e6)
        TextBox1_Validating(TextBox7, e7)
        If e1.Cancel Or e2.Cancel Or e3.Cancel Or e4.Cancel Or e5.Cancel Or e6.Cancel Or e7.Cancel Then
            DialogResult = Windows.Forms.DialogResult.None
            Exit Sub
        End If

        Me.Validate()
        BS.EndEdit()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

    End Sub
End Class