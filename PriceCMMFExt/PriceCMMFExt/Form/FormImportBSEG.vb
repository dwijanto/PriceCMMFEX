Imports System.Threading
Imports PriceCMMFExt.PublicClass
Imports System.Text
Imports PriceCMMFExt.SharedClass

Public Class FormImportBSEG
    Dim mythread As New Thread(AddressOf doWork)
    Dim openfiledialog1 As New OpenFileDialog
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim startdate As Date
    Dim enddate As Date
    Dim appendfile As Boolean
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not mythread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            'Get file
            startdate = getfirstdate(DateTimePicker1.Value)
            enddate = getlastdate(DateTimePicker2.Value)
            'appendfile = RadioButton1.Checked

            If openfiledialog1.ShowDialog = DialogResult.OK Then
                mythread = New Thread(AddressOf dowork)
                mythread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub
    
    Private Sub dowork()
        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        Dim DC As Integer = 1
        sw.Start()
        'Try
        Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'Get Existing Record
                'Dim sqlstr = "select documentno,item from pocurr;"
                'Dim mymessage As String = String.Empty
                'If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                '    ProgressReport(1, mymessage)
                '    Exit Sub
                'End If

                'DS.Tables(0).TableName = "POCurr"
                'Dim idx0(1) As DataColumn
                'idx0(0) = DS.Tables(0).Columns(0)
                'idx0(1) = DS.Tables(0).Columns(1)
                'DS.Tables(0).PrimaryKey = idx0
                ProgressReport(1, "Read Text File. Please wait...")
                ProgressReport(6, "Marquee")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(5, "continuous")

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)
                    If DbAdapter1.dateformatdotdate(myrecord(19)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(19)) <= enddate.Date And
                        myrecord(12) <> "" Then


                        'Dim pkey(1) As Object
                        'pkey(0) = myrecord(1)
                        'pkey(1) = myrecord(3)

                        'Dim result As DataRow = DS.Tables(0).Rows.Find(pkey)
                        'If IsNothing(result) Then
                        'Dim dr As DataRow = DS.Tables(0).NewRow
                        'dr.Item(0) = myrecord(1)
                        'dr.Item(1) = myrecord(3)
                        'DS.Tables(0).Rows.Add(dr)

                        'pohd bigint, poitem bigint, cmmf bigint, qty numeric, curr character varying,  amount numeric, amountlc numeric,
                        If i = 776 Then
                            Debug.Print("debug")
                        End If
                        DC = 1
                        If myrecord(25) = "H" Then
                            DC = -1
                        End If
                        myInsert.Append(myrecord(1) & vbTab &
                                    myrecord(3) & vbTab &
                                    myrecord(12) & vbTab &
                                    myrecord(13) & vbTab &
                                    DbAdapter1.validlongNull(myrecord(8)) & vbTab &
                                    DbAdapter1.validdec(myrecord(10)) & vbTab &
                                    DbAdapter1.validchar(myrecord(17)) & vbTab &
                                    DbAdapter1.validdec(myrecord(15)) & vbTab &
                                    DbAdapter1.validdec(myrecord(23)) & vbTab &
                                    DbAdapter1.dateformatdot(myrecord(19)) & vbTab &
                                    DC & vbCrLf)
                    End If
                    'End If
                Next


            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            Dim ra As Long = 0
            Dim sqlstr As String = "delete from pocurr where valuedate >= " & DbAdapter1.dateformatYYYYMMdd(startdate) & " and valuedate <= " & DbAdapter1.dateformatYYYYMMdd(enddate)

            'If Not appendfile Then
            ra = DbAdapter1.ExNonQuery(sqlstr)
            'End If

            sqlstr = "copy pocurr(documentno,item,pohd , poitem , cmmf , qty, curr,  amount, amountlc,valuedate,dc) from stdin with null as 'Null';"


            ProgressReport(1, "Start Add New Records")
            ProgressReport(6, "Marquee")
            'mystr.Append(sqlstr)

            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            '    MessageBox.Show(errmessage)
            'Else
            '    ProgressReport(1, "Update Done.")
            'End If
            Try

                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                sw.Stop()
                If myret Then

                    ProgressReport(1, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                Else
                    ProgressReport(1, errmessage)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        Else
            sw.Stop()
            ProgressReport(1, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        End If
        ProgressReport(5, "continuous")
        'Catch ex As Exception
        '    'MessageBox.Show(ex.Message)
        '    ProgressReport(1, ex.Message)
        '    ProgressReport(5, "continuous")
        'End Try


    End Sub

    Private Sub doWork1()
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String

        Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'Get Existing Record
                Dim sqlstr = "select pohd,poitem from pocurr;"
                Dim mymessage As String = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(1, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "POCurr"
                Dim idx0(1) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                idx0(1) = DS.Tables(0).Columns(1)
                DS.Tables(0).PrimaryKey = idx0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        'find the record in existing table.
                        Dim pkey(1) As Object
                        pkey(0) = myrecord(12)
                        pkey(1) = myrecord(13)

                        Dim result As DataRow = DS.Tables(0).Rows.Find(pkey)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(12)
                            dr.Item(1) = myrecord(13)
                            DS.Tables(0).Rows.Add(dr)

                            'pohd bigint, poitem bigint, cmmf bigint, qty numeric, curr character varying,  amount numeric, amountlc numeric,

                            myInsert.Append(myrecord(12) & vbTab &
                                        myrecord(13) & vbTab &
                                        DbAdapter1.validlong(myrecord(8)) & vbTab &
                                        DbAdapter1.validdec(myrecord(10)) & vbTab &
                                        DbAdapter1.validchar(myrecord(17)) & vbTab &
                                        DbAdapter1.validdec(myrecord(15)) & vbTab &
                                        DbAdapter1.validdec(myrecord(23)) & vbCrLf)
                        End If

                    End If
                    count += 1

                Loop




            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            ProgressReport(1, "Start Add New Records")
            Dim sqlstr As String = "copy pocurr(pohd , poitem , cmmf , qty, curr,  amount, amountlc) from stdin with null as 'Null';"
            'mystr.Append(sqlstr)
            Dim ra As Long = 0
            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            '    MessageBox.Show(errmessage)
            'Else
            '    ProgressReport(1, "Update Done.")
            'End If
            Try
                'ra = DbAdapter1.ExNonQuery(mystr.ToString)
                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                If myret Then
                    ProgressReport(1, "Add Records Done.")
                Else
                    ProgressReport(1, errmessage)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        End If
    End Sub
    Private Sub dowork2()
        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        sw.Start()
        Using objTFParser = New FileIO.TextFieldParser(openfiledialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'Get Existing Record
                Dim sqlstr = "select pohd,poitem from pocurr;"
                Dim mymessage As String = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(1, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "POCurr"
                Dim idx0(1) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                idx0(1) = DS.Tables(0).Columns(1)
                DS.Tables(0).PrimaryKey = idx0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)
                    If DbAdapter1.dateformatdotdate(myrecord(19)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(19)) <= enddate.Date And
                        myrecord(12) <> "" Then


                        Dim pkey(1) As Object
                        pkey(0) = myrecord(12)
                        pkey(1) = myrecord(13)

                        Dim result As DataRow = DS.Tables(0).Rows.Find(pkey)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(12)
                            dr.Item(1) = myrecord(13)
                            DS.Tables(0).Rows.Add(dr)

                            'pohd bigint, poitem bigint, cmmf bigint, qty numeric, curr character varying,  amount numeric, amountlc numeric,

                            myInsert.Append(myrecord(12) & vbTab &
                                        myrecord(13) & vbTab &
                                        DbAdapter1.validlong(myrecord(8)) & vbTab &
                                        DbAdapter1.validdec(myrecord(10)) & vbTab &
                                        DbAdapter1.validchar(myrecord(17)) & vbTab &
                                        DbAdapter1.validdec(myrecord(15)) & vbTab &
                                        DbAdapter1.validdec(myrecord(23)) & vbTab &
                                        DbAdapter1.dateformatdot(myrecord(19)) & vbCrLf)
                        End If
                    End If
                Next


            End With
        End Using
        'update record
        If myInsert.Length > 0 Then
            Dim ra As Long = 0
            Dim sqlstr As String = "delete from pocurr where valuedate >= " & DbAdapter1.dateformatYYYYMMdd(startdate) & " and valuedate <= " & DbAdapter1.dateformatYYYYMMdd(enddate)

            If Not appendfile Then
                ra = DbAdapter1.ExNonQuery(sqlstr)
            End If

            sqlstr = "copy pocurr(pohd , poitem , cmmf , qty, curr,  amount, amountlc,valuedate) from stdin with null as 'Null';"


            ProgressReport(1, "Start Add New Records")
            'mystr.Append(sqlstr)

            Dim errmessage As String = String.Empty
            Dim myret As Boolean = False
            'If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmessage) Then
            '    MessageBox.Show(errmessage)
            'Else
            '    ProgressReport(1, "Update Done.")
            'End If
            Try

                errmessage = DbAdapter1.copy(sqlstr, myInsert.ToString, myret)
                sw.Stop()
                If myret Then

                    ProgressReport(1, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
                Else
                    ProgressReport(1, errmessage)
                End If
            Catch ex As Exception
                ProgressReport(1, ex.Message)
            End Try
        Else
            sw.Stop()
            ProgressReport(1, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        End If

    End Sub

    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.ToolStripStatusLabel1.Text = message
                Case 5
                    ToolStripProgressBar1.ProgressBar.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.ProgressBar.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    ToolStripProgressBar1.Minimum = 1
                    ToolStripProgressBar1.Value = myvalue(0)
                    ToolStripProgressBar1.Maximum = myvalue(1)
            End Select

        End If

    End Sub
    'Private Function validstr(ByVal data As Object) As Object
    '    If IsDBNull(data) Then
    '        Return "Null"
    '    ElseIf data = "" Then
    '        Return "Null"
    '    End If
    '    Return data
    'End Function




   

End Class