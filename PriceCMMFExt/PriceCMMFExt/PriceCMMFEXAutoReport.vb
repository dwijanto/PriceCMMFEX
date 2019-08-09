Imports System.Threading
Public Class PriceCMMFEXAutoReport
    Dim myThread As New System.Threading.Thread(AddressOf doWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Private Sub PriceCMMFExAuto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Me.Load

        Me.WindowState = FormWindowState.Minimized
        LoadMe()
    End Sub


    Private Sub PriceCMMFExAuto_Resize(ByVal sender As Object, ByVal e As System.EventArgs)
        If Me.WindowState = FormWindowState.Minimized Then
            Me.ShowInTaskbar = False
            Me.Hide()
            NotifyIcon1.Visible = True
        End If
    End Sub


    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

    End Sub

    Private Sub LoadMe()

        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(AddressOf doWork)
                myThread.TrySetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Sub doWork()
        Logger.log("--------Start----------")

        Logger.log("Send Auto PriceChange Summary")
        Logger.log("Send email to Assigned User")
        Dim MyReport = New AutoReport(AutoReport.ReportType.AutoPriceChangeSummary)
        If Not MyReport.Execute Then
            Logger.log(MyReport.errormessage)
        End If

        Logger.log("--------End------------")
        ProgressReport(1, "Close Apps")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.Close()
                Case 2

                Case 3

                Case 4

                Case 5

                Case 6

                Case 7

            End Select
        End If
    End Sub
End Class