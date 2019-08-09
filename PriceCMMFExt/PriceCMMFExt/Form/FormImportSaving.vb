Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Imports System.IO

Public Class FormImportSaving
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    'Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    'Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim validperiod As Integer
    Dim SavingLookupSeq As Long
    Dim updatesequencesqlstr As String
    Dim myseqid As Integer = 1
    Dim startdate As Date
    Dim enddate As Date
    Dim DataReaderCallback As FormatReportDelegate = AddressOf doReaderCallBack
    Dim FileName As String = String.Empty
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'Get file
            validperiod = TextBox1.Text

            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Ask confirmation
        startdate = DateTimePicker1.Value.Date
        enddate = DateTimePicker2.Value.Date
        If MessageBox.Show(String.Format("Do you want to delete this period From {0:dd-MMM-yyyy} To {1:dd-MMM-yyyy}", startdate, enddate), "Delete Records", System.Windows.Forms.MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.Cancel Then
            Exit Sub
        End If
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""

            myThread = New Thread(AddressOf DoDelete)
            myThread.Start()
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim SaveFileDialog1 = New SaveFileDialog
        SaveFileDialog1.FileName = "Saving.csv"
        If SaveFileDialog1.ShowDialog = DialogResult.OK Then
            FileName = SaveFileDialog1.FileName
            myThread = New Thread(AddressOf DoSaveFile)
            myThread.Start()
        End If
    End Sub

    Private Sub DoDelete()
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Deleting... Please wait.")
        ProgressReport(6, "Marquee")
        Dim mymessage As String = String.Empty
        Dim sqlstr As String = String.Format("delete from saving where startdate >= '{0:yyyy-MM-dd}' and startdate <= '{1:yyyy-MM-dd}';", startdate, enddate)
        If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
            sw.Stop()
            ProgressReport(1, mymessage)
        Else
            sw.Stop()
            ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        End If
        ProgressReport(5, "Continuous")

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 2
                    Me.ToolStripStatusLabel2.Text = message
                Case 4
                    'Me.Label4.Text = message
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select

        End If

    End Sub



    Sub DoWork()
        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        Dim enddate As Date
        Dim SavingLookupSB As New System.Text.StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())
        Dim typeid As Long
        Dim actionid As Long
        Dim sqlstr As String = String.Empty

        Dim mymessage As String = String.Empty
        sw.Start()
        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'FillData
                ProgressReport(2, "Initialize Table..")
                sqlstr = "select savingid,actionid,cmmf,mytotal,startdate,enddate from saving;" &
                         " select savinglookupid,savinglookupname,parentid from savinglookup order by savinglookupid desc;"

                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "Saving"
                Dim idx0(2) As DataColumn
                idx0(0) = DS.Tables(0).Columns(1)
                idx0(1) = DS.Tables(0).Columns(2)
                idx0(2) = DS.Tables(0).Columns(4)

                DS.Tables(0).PrimaryKey = idx0

                DS.Tables(1).TableName = "savinglookup"
                Dim idx1(0) As DataColumn
                idx1(0) = DS.Tables(1).Columns(1)
                DS.Tables(1).CaseSensitive = True
                DS.Tables(1).PrimaryKey = idx1


                If DS.Tables(1).Rows.Count > 0 Then
                    SavingLookupSeq = DS.Tables(1).Rows(0).Item(0)
                    myseqid = SavingLookupSeq + 1

                End If
                updatesequencesqlstr = "select setval('savinglookup_savinglookupid_seq'," & myseqid & ",false);"

                ProgressReport(2, "Read Text File...")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")
                Try
                    For i = 0 To mylist.Count - 1
                        'find the record in existing table.
                        ProgressReport(7, i + 1 & "," & mylist.Count)
                        myrecord = mylist(i)
                        If i >= 0 Then

                            enddate = CDate(myrecord(4)).AddYears(validperiod)
                            Dim result As DataRow

                            Dim pkey1(0) As Object
                            pkey1(0) = myrecord(0)
                            result = DS.Tables(1).Rows.Find(pkey1)
                            If IsNothing(result) Then
                                SavingLookupSeq += 1
                                Dim dr As DataRow = DS.Tables(1).NewRow
                                dr.Item(0) = SavingLookupSeq
                                dr.Item(1) = myrecord(0)
                                dr.Item(2) = 0
                                DS.Tables(1).Rows.Add(dr)
                                typeid = SavingLookupSeq
                            Else
                                typeid = result.Item(0)
                            End If

                            'check actionid
                            Dim pkey2(0) As Object
                            pkey2(0) = myrecord(1)
                            result = DS.Tables(1).Rows.Find(pkey2)
                            If IsNothing(result) Then
                                SavingLookupSeq += 1
                                Dim dr As DataRow = DS.Tables(1).NewRow
                                dr.Item(0) = SavingLookupSeq
                                dr.Item(1) = myrecord(1)
                                dr.Item(2) = typeid
                                DS.Tables(1).Rows.Add(dr)
                                actionid = SavingLookupSeq
                            Else
                                actionid = result.Item(0)
                            End If

                            Dim pkey(2) As Object
                            pkey(0) = actionid
                            pkey(1) = myrecord(2)
                            pkey(2) = myrecord(4)

                            result = DS.Tables(0).Rows.Find(pkey)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(0).NewRow
                                dr.Item(1) = actionid
                                dr.Item(2) = myrecord(2)
                                dr.Item(3) = myrecord(3)
                                dr.Item(4) = myrecord(4)
                                dr.Item(5) = enddate
                                DS.Tables(0).Rows.Add(dr)
                            Else
                                If result.Item("mytotal") <> myrecord(3) Then
                                    result.Item("mytotal") = myrecord(3)
                                End If
                                If result.Item("enddate") <> enddate Then
                                    result.Item("enddate") = enddate
                                End If
                            End If
                        End If
                    Next
                Catch ex As Exception
                    ProgressReport(2, ex.Message)
                    Exit Sub
                End Try
            End With
        End Using
        'update record
        Try
            Dim errmsg As String = String.Empty
            ProgressReport(6, "Marque")

            Dim ds2 As New DataSet
            ds2 = DS.GetChanges
            If Not IsNothing(ds2) Then
                'Dim mymessage As String = String.Empty
                Dim ra As Integer
                'reset sequence number
                If DbAdapter1.ExecuteNonQuery(updatesequencesqlstr, ra, mymessage) Then
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.ImportTx(Me, mye) Then
                        ProgressReport(1, mye.message)
                    End If
                Else
                    ProgressReport(1, mymessage)
                End If
            End If

        Catch ex As Exception
            ProgressReport(1, ex.Message)

        End Try
        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub

    Sub doReaderCallBack(ByRef sender As Object, ByRef e As EventArgs)
        Dim dr As Npgsql.NpgsqlDataReader = DirectCast(sender, Npgsql.NpgsqlDataReader)
        Using sw As StreamWriter = New StreamWriter(FileName)
            sw.WriteLine("""savingid"",""actionid"",""cmmf"",""mytotal"",""startdate"",""enddate""")
            While (dr.Read())
                sw.WriteLine(String.Format("{0},{1},{2},{3},""{4:yyyy-MM-dd}"",""{5:yyyy-MM-dd}""", dr(0), dr(1), dr(2), dr(3), dr(4), dr(5)))
            End While
        End Using
    End Sub

    Private Sub DoSaveFile()
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Generating Raw Data.. Please wait.")
        ProgressReport(6, "Marquee")
        Dim mymessage As String = String.Empty
        Dim sqlstr As String = String.Format("select savingid,actionid,cmmf,mytotal,startdate,enddate from saving order by startdate;")
        DbAdapter1.DataReaderCallback = DataReaderCallback
        If Not DbAdapter1.ExecuteReader(sqlstr, message:=mymessage) Then
            sw.Stop()
            ProgressReport(1, mymessage)
        Else
            sw.Stop()
            ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        End If
        ProgressReport(5, "Continuous")


    End Sub






End Class