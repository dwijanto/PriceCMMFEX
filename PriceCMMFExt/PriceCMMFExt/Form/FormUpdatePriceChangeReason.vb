Imports System.Threading
Imports System.Text

Public Class FormUpdatePriceChangeReason
    Dim myThreadDelegate As New ThreadStart(AddressOf dowork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim openfiledialog1 As New OpenFileDialog
    Dim FileName As String
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            OpenFileDialog1.FileName = ""
            'Get file
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                FileName = openfiledialog1.FileName
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub dowork()
        Dim myProcess As New UpdatePriceChangeReason(FileName, Me)
        If Not myProcess.Run() Then
            MessageBox.Show(myProcess.ErrorMessage)
        End If
    End Sub

End Class