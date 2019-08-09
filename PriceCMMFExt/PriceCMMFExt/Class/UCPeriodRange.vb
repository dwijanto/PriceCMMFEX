Public Class UCPeriodRange
    Public Property Year1 As Integer
    Public Property Year2 As Integer

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        SetDateTimePicker()
    End Sub

    Private Sub UCPeriodRange_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        SetDateTimePicker()
        setYear()
    End Sub

    Private Sub SetDateTimePicker()
        DateTimePicker1.Enabled = CheckBox1.Checked
        DateTimePicker2.Enabled = CheckBox1.Checked
    End Sub

    Private Sub setYear()
        Year1 = DateTimePicker1.Value.Year
        Year2 = DateTimePicker2.Value.Year
    End Sub

    Private Sub DateTimePicker1_ValueChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged, DateTimePicker2.ValueChanged
        setYear()
    End Sub
End Class
