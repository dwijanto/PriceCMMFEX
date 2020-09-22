Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormImportStdCost
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Private Property FileName As String

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)

    Private DS As DataSet
    Dim AddSTDPriceSB As StringBuilder
    Dim UpdSTDPriceSB As StringBuilder
    Dim UpdValidToSB As StringBuilder
    Dim startdate As Date
    Dim enddate = CDate("9999-01-01")

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            OpenFileDialog1.FileName = ""
            'Get file
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then
                startdate = String.Format("{0:yyyy-MM-1}", DateTimePicker1.Value.Date)
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
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
        Dim sw2 As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of STDCostModel)
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

                Dim sb As New StringBuilder
                sb.Append("with std as (select cmmf,first_value(planprice1) over (partition by cmmf order by validfrom desc) as planprice1" &
                          " ,first_value(per) over (partition by cmmf order by validfrom desc) as per" &
                          ",first_value(validfrom) over (partition by cmmf order by validfrom desc) as validfrom from standardcostad" &
                          " where not planprice1 isnull)" &
                          " select distinct * from std order by cmmf")
                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "STDCost"
                Dim idx0(0) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                DS.Tables(0).PrimaryKey = idx0

                ProgressReport(2, "Read Text File...")
                Try
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count >= 1 Then
                            Dim mydata As New STDCostModel With {.cmmf = myrecord(0),
                                                                 .amount = myrecord(1)}
                            mylist.Add(mydata)
                        End If
                        count += 1
                    Loop
                Catch ex As Exception
                    ProgressReport(1, ex.Message)
                    Exit Sub
                End Try

                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                AddSTDPriceSB = New StringBuilder
                UpdSTDPriceSB = New StringBuilder
                UpdValidToSB = New StringBuilder
                Dim i As Long
                sw2.Start()

                Try
                    For i = 0 To mylist.Count - 1
                 
                        Dim updateFlag As Boolean = False
                        'find the record in existing table.
                        ProgressReport(7, i + 1 & "," & mylist.Count)
                        Dim mydata = mylist(i)
                        'Find Data if not avail then create else update
                        Dim result As DataRow
                        Dim pkey1(0) As Object
                        pkey1(0) = CLng(mydata.cmmf)
                        result = DS.Tables(0).Rows.Find(pkey1)
                        If IsNothing(result) Then
                            'Create
                            AddSTDPriceSB.Append(mydata.cmmf & vbTab &
                                                 mydata.amount & vbTab &
                                                 "1" & vbTab &
                                                 "Null" & vbTab &
                                                 String.Format("'{0:yyyy-MM-dd}'", startdate) & vbTab &
                                                 String.Format("'{0:yyyy-MM-dd}'", enddate) & vbCrLf)
                        Else
                            'cmmf,planprice1,per,ppdate,validfrom,validto

                            If result.Item("validfrom") = startdate Then
                                If result.Item("planprice1") / result.Item("per") <> Val(mydata.amount) Or IsNothing(result.Item("planprice1")) Then
                                    If UpdSTDPriceSB.Length > 0 Then
                                        UpdSTDPriceSB.Append(",")
                                    End If
                                    UpdSTDPriceSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{3:yyyy-MM-dd}'::character varying]", mydata.cmmf, mydata.amount, startdate))
                                End If
                               

                            ElseIf result.Item("planprice1") / result.Item("per") <> Val(mydata.amount) Or IsNothing(result.Item("planprice1")) Then
                                If UpdValidToSB.Length > 0 Then
                                    UpdValidToSB.Append(",")
                                End If
                                UpdValidToSB.Append(String.Format("['{0}'::character varying,'{1:yyyy-MM-dd}'::character varying,'{2:yyyy-MM-dd}'::character varying]", mydata.cmmf, result.Item("validfrom"), startdate.AddDays(-1)))
                                AddSTDPriceSB.Append(mydata.cmmf & vbTab &
                                                 mydata.amount & vbTab &
                                                 "1" & vbTab &
                                                 "Null" & vbTab &
                                                 String.Format("'{0:yyyy-MM-dd}'", startdate) & vbTab &
                                                 String.Format("'{0:yyyy-MM-dd}'", enddate) & vbCrLf)
                            End If
                            'If updateFlag Then
                            '    If UpdSTDPriceSB.Length > 0 Then
                            '        UpdSTDPriceSB.Append(",")
                            '    End If
                            '    UpdSTDPriceSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3:yyyy-MM-dd}'::character varying]", mydata.cmmf, mydata.amount, startdate))
                            'End If

                        End If
                    Next
                Catch ex As Exception
                    ProgressReport(2, ex.Message & "  record :" & i)
                    Exit Sub
                End Try
            End With
        End Using
        sw2.Stop()
        ProgressReport(1, String.Format("Build Record Done. Elapsed Time: {0}:{1}.{2}", Format(sw2.Elapsed.Minutes, "00"), Format(sw2.Elapsed.Seconds, "00"), sw2.Elapsed.Milliseconds.ToString))
        sw2.Start()

        'update record
        Dim myret As Boolean = False
        Try
            Dim errmsg As String = String.Empty
            ProgressReport(6, "Marque")

            If UpdSTDPriceSB.Length > 0 Then
                ProgressReport(1, "Update Data...")
                'myyear	cmmf	vendorcode	qtyshipped	purchasevalue	averprice	averpricefixcurr
                sqlstr = "update standardcostad set planprice1= foo.planprice::numeric from (select * from array_to_set3(Array[" & UpdSTDPriceSB.ToString &
                         "]) as tb (cmmf1 character varying,planprice character varying,validfrom1 character varying))foo where cmmf = foo.cmmf1::bigint and validfrom = foo.validfrom1::date;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(1, errmsg)
                    myret = False
                    Err.Raise(513, Description:=errmsg & " ::Update Data")
                End If
            End If

            If UpdValidToSB.Length > 0 Then
                ProgressReport(1, "Update Data...")
                'myyear	cmmf	vendorcode	qtyshipped	purchasevalue	averprice	averpricefixcurr
                sqlstr = "update standardcostad set validto= foo.validto1::date from (select * from array_to_set3(Array[" & UpdValidToSB.ToString &
                         "]) as tb (cmmf1 character varying,validfrom1 character varying,validto1 character varying))foo where cmmf = foo.cmmf1::bigint and validfrom = foo.validfrom1::date;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(1, errmsg)
                    myret = False
                    Err.Raise(513, Description:=errmsg & " ::Update Data")
                End If
            End If

            If AddSTDPriceSB.Length > 0 Then
                ProgressReport(2, "Copy standardcostad")

                sqlstr = "copy standardcostad(cmmf,planprice1,per,ppdate,validfrom,validto) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                myret = False
                errmessage = DbAdapter1.copy(sqlstr, AddSTDPriceSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy standardcostad" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            myret = True
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        Finally
            ProgressReport(5, "Continue")
        End Try
        sw2.Stop()
        If myret Then
            ProgressReport(1, String.Format("Copy Done. Elapsed Time: {0}:{1}.{2}", Format(sw2.Elapsed.Minutes, "00"), Format(sw2.Elapsed.Seconds, "00"), sw2.Elapsed.Milliseconds.ToString))
        Else

        End If
        sw.Stop()

    End Sub

End Class

Public Class STDCostModel
    Public Property cmmf As String
    Public Property amount As String
End Class