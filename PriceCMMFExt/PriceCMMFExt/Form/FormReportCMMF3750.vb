Public Class FormReportCMMF3750
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Me.ToolStripStatusLabel1.Text = ""
        Me.ToolStripStatusLabel2.Text = ""


        runreport(sender, e)
    End Sub
    Private Sub runreport(ByVal sender As System.Object, ByVal e As System.EventArgs)

        Dim mymessage As String = String.Empty

        Dim sqlstr As String = "select pp.plant,p.vendorcode,p.cmmf,p.validfrom,p.amount,p.currency,p.perunit,p.uom from doc.cmmf3750 c" &
                               " left join pricelist p on p.cmmf = c.cmmf" &
                               " left join priceplant pp on pp.pricelistid = p.pricelistid" &
                               " where pp.plant = 3750 order by vendorcode,cmmf,validfrom desc;"

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("CMMF3750-{0:yyyyMMdd}.xlsx", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

            Dim myreport As New ExportToExcelFile(Me, sqlstr, filename, reportname, mycallback, PivotCallback, datasheet, "\templates\ExcelTemplate.xltx")
            myreport.Run(Me, e)
        End If
    End Sub

    Private Sub FormattingReport(ByRef sender As Object, ByRef e As EventArgs)

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
    End Sub


    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try
                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel2.Text = message
                    Case 4
                        runreport(Me, New System.EventArgs)
                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub
End Class