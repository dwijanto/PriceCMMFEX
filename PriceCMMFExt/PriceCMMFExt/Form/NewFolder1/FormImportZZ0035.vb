Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormImportZZ0035
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)


    Dim myQueryThread As New System.Threading.Thread(QueryDelegate)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date

    Dim miroSeq As Long
    Dim podtlseq As Long
    Dim cmmfpriceseq As Long
    Dim cmmfvendorpriceseq As Long

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            'Get file
            startdate = DateTimePicker1.Value 'getfirstdate(DateTimePicker1.Value)
            enddate = DateTimePicker2.Value 'getlastdate(DateTimePicker2.Value)
            'appendfile = RadioButton1.Checked
            If Year(startdate) <> Year(enddate) Then
                MessageBox.Show("Period should be the same year!")
                Exit Sub
            End If
            If Year(startdate) <> Year(Today) Then
                If MessageBox.Show("The year is not the same as current year. Proceed?", "Question", MessageBoxButtons.OKCancel) = Windows.Forms.DialogResult.Cancel Then
                    Exit Sub
                End If
            End If
            If openfiledialog1.ShowDialog = DialogResult.OK Then
                myThread = New Thread(AddressOf DoWork)
                myThread.Start()
            End If
        Else
            MessageBox.Show("Process still running. Please Wait!")
        End If
    End Sub

    Sub DoQuery()
        'Get last MiroPostingDate
        Dim sqlstr = "select miropostingdate from miro m order by miropostingdate desc limit 1;"
        Dim DS As New DataSet
        Dim mymessage As String = String.Empty
        If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
            ProgressReport(2, mymessage)
        Else
            If DS.Tables(0).Rows.Count > 0 Then
                ProgressReport(4, String.Format("{0:dd-MMM-yyyy}", DS.Tables(0).Rows(0).Item(0)))
            End If

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
                    Me.Label4.Text = message
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

    Private Sub FormImportZZ0035_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        myQueryThread.Start()
    End Sub

    Sub DoWork()
        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        Dim MiroSB As New System.Text.StringBuilder
        Dim POHDSB As New System.Text.StringBuilder
        Dim PODtlSB As New System.Text.StringBuilder
        Dim POMiroSB As New System.Text.StringBuilder
        Dim cmmfSB As New System.Text.StringBuilder
        Dim cmmfpriceSB As New System.Text.StringBuilder
        Dim cmmfvendorpriceSB As New System.Text.StringBuilder
        Dim updatecmmfpricesb As New System.Text.StringBuilder
        Dim updateCMMFvendorpriceLastsb As New System.Text.StringBuilder
        Dim updateCMMFvendorpriceInitsb As New System.Text.StringBuilder
        Dim vendorSB As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        Dim miroid As Long
        Dim podtlid As Long
        Dim sqlstr As String = String.Empty
        Dim postingdate As Date

        sw.Start()
        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                'Delete Existing Record
                ProgressReport(2, "Delete ..")
                ProgressReport(6, "Marque")
                sqlstr = "delete from pomiro where miroid in (select miroid from miro where miropostingdate >= " & DateFormatyyyyMMdd(startdate) & " and miropostingdate <= " & DateFormatyyyyMMdd(enddate) & "); " &
                         "delete from miro m where m.miropostingdate >= " & DateFormatyyyyMMdd(startdate) & " and m.miropostingdate <= " & DateFormatyyyyMMdd(enddate) & ";" &
                         "select setval('miro_miroid_seq',(select miroid from miro order by miroid desc limit 1) + 1,false);" &
                         "select setval('pomiro_pomiroid_seq',(select pomiroid from pomiro order by pomiroid desc limit 1) + 1,false);"
                Dim mymessage As String = String.Empty
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                'FillData
                ProgressReport(2, "Initialize Table..")
                sqlstr = "select mironumber,miroid from miro where mironumber = 0;" &
                         " select pohd from pohd;" &
                         " select pohd,polineno,podtlid from podtl;" &
                         " select miroid from miro order by miroid desc limit 1;" &
                         " select podtlid from podtl order by podtlid desc limit 1;" &
                         " select cmmf from cmmf;" &
                         " select cmmf,myyear,lasttx,cpid from cmmfprice where myyear = " & Year(startdate) & ";" &
                         " select cmmf,vendorcode,myyear,lasttx,initialtx,cpid from  cmmfvendorprice where myyear = " & Year(startdate) & ";" &
                         " select cpid from cmmfprice order by cpid desc limit 1;" &
                         " select cpid from cmmfvendorprice order by cpid desc limit 1;" &
                         " select vendorcode from vendor"
                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "Miro"
                Dim idx0(0) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                DS.Tables(0).PrimaryKey = idx0

                DS.Tables(1).TableName = "pohd"
                Dim idx1(0) As DataColumn
                idx1(0) = DS.Tables(1).Columns(0)
                DS.Tables(1).PrimaryKey = idx1

                DS.Tables(2).TableName = "podtl"
                Dim idx2(1) As DataColumn
                idx2(0) = DS.Tables(2).Columns(0)
                idx2(1) = DS.Tables(2).Columns(1)
                DS.Tables(2).PrimaryKey = idx2

                DS.Tables(3).TableName = "miroseq"
                If DS.Tables(3).Rows.Count > 0 Then
                    miroSeq = DS.Tables(3).Rows(0).Item(0)
                End If

                DS.Tables(4).TableName = "podtlseq"
                If DS.Tables(4).Rows.Count > 0 Then
                    podtlseq = DS.Tables(4).Rows(0).Item(0)
                End If

                DS.Tables(5).TableName = "cmmf"
                Dim idx5(0) As DataColumn
                idx5(0) = DS.Tables(5).Columns(0)
                DS.Tables(5).PrimaryKey = idx5

                DS.Tables(6).TableName = "cmmfprice"
                Dim idx6(1) As DataColumn
                idx6(0) = DS.Tables(6).Columns(0)
                idx6(1) = DS.Tables(6).Columns(1)
                DS.Tables(6).PrimaryKey = idx6

                DS.Tables(7).TableName = "cmmfvendorprice"
                Dim idx7(2) As DataColumn
                idx7(0) = DS.Tables(7).Columns(0)
                idx7(1) = DS.Tables(7).Columns(1)
                idx7(2) = DS.Tables(7).Columns(2)
                DS.Tables(7).PrimaryKey = idx7

                DS.Tables(8).TableName = "cmmfpriceseq"
                If DS.Tables(8).Rows.Count > 0 Then
                    cmmfpriceseq = DS.Tables(8).Rows(0).Item(0)
                End If

                DS.Tables(9).TableName = "cmmfvendorpriceseq"
                If DS.Tables(9).Rows.Count > 0 Then
                    cmmfvendorpriceseq = DS.Tables(9).Rows(0).Item(0)
                End If

                DS.Tables(10).TableName = "vendor"
                Dim idx10(0) As DataColumn
                idx10(0) = DS.Tables(10).Columns(0)
                DS.Tables(10).PrimaryKey = idx10

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

                For i = 0 To mylist.Count - 1
                    'find the record in existing table.
                    ProgressReport(7, i + 1 & "," & mylist.Count)
                    myrecord = mylist(i)
                    If i >= 1 Then
                        postingdate = DbAdapter1.dateformatdotdate(myrecord(11))
                        'If DbAdapter1.dateformatdotdate(myrecord(11)) >= startdate.Date AndAlso DbAdapter1.dateformatdotdate(myrecord(11)) <= enddate.Date Then
                        If postingdate >= startdate.Date AndAlso postingdate <= enddate.Date Then
                            'If myrecord(4) = "1830003754" Then
                            '    Debug.Print("debug")
                            'End If
                            Dim result As DataRow
                            'check cmmf if not avail then create
                            If Not myrecord(4) = "" Then

                                Dim pkey5(0) As Object
                                pkey5(0) = myrecord(4)
                                result = DS.Tables(5).Rows.Find(pkey5)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables(5).NewRow
                                    dr.Item(0) = myrecord(4)
                                    DS.Tables(5).Rows.Add(dr)
                                    cmmfSB.Append(myrecord(4) & vbCrLf)
                                End If

                                'check cmmfprice
                                Dim pkey6(1) As Object
                                pkey6(0) = myrecord(4)
                                pkey6(1) = Year(startdate)
                                result = DS.Tables(6).Rows.Find(pkey6)
                                If IsNothing(result) Then
                                    cmmfpriceseq += 1
                                    Dim dr As DataRow = DS.Tables(6).NewRow
                                    dr.Item(0) = myrecord(4)
                                    dr.Item(1) = Year(startdate)
                                    dr.Item(2) = postingdate
                                    dr.Item(3) = cmmfpriceseq

                                    DS.Tables(6).Rows.Add(dr)
                                    'cmmf,myyear,initailtx,initialprice,incoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2
                                    cmmfpriceSB.Append(myrecord(4) & vbTab &
                                                       Year(startdate) & vbTab &
                                                       DbAdapter1.dateformatdot(myrecord(11)) & vbTab &
                                                       DbAdapter1.validdec(myrecord(16)) & vbTab &
                                                       DbAdapter1.validlong(myrecord(10)) & vbTab &
                                                       DbAdapter1.dateformatdot(myrecord(11)) & vbTab &
                                                       DbAdapter1.validdec(myrecord(16)) & vbTab &
                                                       DbAdapter1.validlong(myrecord(10)) & vbCrLf)
                                Else
                                    'update
                                    If result.Item(2) < postingdate Then
                                        If updatecmmfpricesb.Length > 0 Then
                                            updatecmmfpricesb.Append(",")
                                        End If
                                        'lasttx,lastprice,invoiceverificationnumber2
                                        updatecmmfpricesb.Append(String.Format("['{0}'::character varying,{1}::character varying,'{2}'::character varying,'{3}'::character varying]",
                                                             result.Item(3), DbAdapter1.dateformatdot(myrecord(11)), DbAdapter1.validdec(myrecord(16)), DbAdapter1.validlong(myrecord(10))))

                                    End If
                                End If


                                'check cmmfpricevendor
                                Dim pkey7(2) As Object
                                pkey7(0) = myrecord(4)
                                pkey7(1) = myrecord(7)
                                pkey7(2) = Year(startdate)
                                result = DS.Tables(7).Rows.Find(pkey7)
                                If IsNothing(result) Then
                                    cmmfvendorpriceseq += 1
                                    Dim dr As DataRow = DS.Tables(7).NewRow
                                    dr.Item(0) = myrecord(4)
                                    dr.Item(1) = myrecord(7)
                                    dr.Item(2) = Year(startdate)
                                    dr.Item(3) = postingdate
                                    dr.Item(4) = postingdate
                                    dr.Item(5) = cmmfvendorpriceseq
                                    DS.Tables(7).Rows.Add(dr)
                                    'cmmf,vendorcode,myyear,initailtx,initialprice,incoiceverificationnumber,lastprice,invoiceverificationnumber2
                                    cmmfvendorpriceSB.Append(myrecord(4) & vbTab &
                                                       myrecord(7) & vbTab &
                                                       Year(startdate) & vbTab &
                                                       DbAdapter1.dateformatdot(myrecord(11)) & vbTab &
                                                       DbAdapter1.validdec(myrecord(16)) & vbTab &
                                                       DbAdapter1.validlong(myrecord(10)) & vbTab &
                                                       DbAdapter1.dateformatdot(myrecord(11)) & vbTab &
                                                       DbAdapter1.validdec(myrecord(16)) & vbTab &
                                                       DbAdapter1.validlong(myrecord(10)) & vbCrLf)
                                Else
                                    'update
                                    If result.Item(3) < postingdate Then
                                        If updateCMMFvendorpriceLastsb.Length > 0 Then
                                            updateCMMFvendorpriceLastsb.Append(",")
                                        End If
                                        'lasttx,lastprice,invoiceverificationnumber2
                                        result.Item("lasttx") = postingdate
                                        updateCMMFvendorpriceLastsb.Append(String.Format("['{0}'::character varying,{1}::character varying,'{2}'::character varying,'{3}'::character varying]",
                                                      result.Item(5), DbAdapter1.dateformatdot(myrecord(11)), DbAdapter1.validdec(myrecord(16)), DbAdapter1.validlong(myrecord(10))))

                                    ElseIf result.Item(4) > postingdate Then

                                        'initialtx,initialprice,invoiceverificationnumber
                                        If updateCMMFvendorpriceInitsb.Length > 0 Then
                                            updateCMMFvendorpriceInitsb.Append(",")
                                        End If
                                        result.Item(4) = postingdate
                                        updateCMMFvendorpriceInitsb.Append(String.Format("['{0}'::character varying,{1}::character varying,'{2}'::character varying,'{3}'::character varying]",
                                                           result.Item(5), DbAdapter1.dateformatdot(myrecord(11)), DbAdapter1.validdec(myrecord(16)), DbAdapter1.validlong(myrecord(10))))

                                    End If
                                End If
                            End If

                            'Check Vendorcode
                            Dim pkey10(0) As Object
                            pkey10(0) = myrecord(7)
                            result = DS.Tables(10).Rows.Find(pkey10)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(10).NewRow
                                dr.Item(0) = myrecord(7)
                                DS.Tables(10).Rows.Add(dr)
                                'vendorcode,vendorname
                                vendorSB.Append(myrecord(7) & vbTab &
                                            (myrecord(8)) & vbCrLf)
                            End If

                            'Miro

                            'Check Vendorcode


                            Dim pkey(0) As Object
                            pkey(0) = myrecord(10)
                            result = DS.Tables(0).Rows.Find(pkey)
                            If IsNothing(result) Then
                                miroSeq += 1

                                Dim dr As DataRow = DS.Tables(0).NewRow
                                dr.Item(0) = myrecord(10)
                                dr.Item(1) = miroSeq
                                DS.Tables(0).Rows.Add(dr)
                                miroid = miroSeq
                                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                                MiroSB.Append(myrecord(10) & vbTab &
                                            DbAdapter1.dateformatdot(myrecord(11)) & vbTab &
                                            DbAdapter1.validchar(myrecord(9)) & vbTab &
                                            DbAdapter1.validlong(myrecord(7)) & vbCrLf)
                            Else
                                miroid = result.Item(1)
                            End If

                            'pohd
                            Dim pkey1(0) As Object
                            pkey1(0) = myrecord(1)
                            result = DS.Tables(1).Rows.Find(pkey1)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(1).NewRow
                                dr.Item(0) = myrecord(1)
                                DS.Tables(1).Rows.Add(dr)
                                'pohd bigint, pono character varying,purchasinggroup character varying, payt character varying
                                POHDSB.Append(myrecord(1) & vbTab &
                                                DbAdapter1.validchar(myrecord(3)) & vbTab &
                                                 DbAdapter1.validchar(myrecord(6)) & vbCrLf)
                            End If

                            'podtl
                            Dim pkey2(1) As Object
                            pkey2(0) = myrecord(1)
                            pkey2(1) = myrecord(2)
                            result = DS.Tables(2).Rows.Find(pkey2)

                            If IsNothing(result) Then

                                podtlseq += 1
                                Dim dr As DataRow = DS.Tables(2).NewRow
                                dr.Item(0) = myrecord(1)
                                dr.Item(1) = myrecord(2)
                                dr.Item(2) = podtlseq
                                DS.Tables(2).Rows.Add(dr)
                                podtlid = podtlseq
                                'pohd bigint, polineno character varying,cmmf bigint,oun character varying
                                PODtlSB.Append(myrecord(1) & vbTab &
                                               DbAdapter1.validint(myrecord(2)) & vbTab &
                                               DbAdapter1.validlongNull(myrecord(4)) & vbTab &
                                               DbAdapter1.validchar(myrecord(15)) & vbCrLf)
                            Else
                                podtlid = result.Item(2)
                            End If

                            'pomiro
                            'podtlid bigint,miroid bigint,amount numeric,qty numeric,crcy charcter varying,unitprice

                            POMiroSB.Append(podtlid & vbTab &
                                                miroid & vbTab &
                                                DbAdapter1.validdec(myrecord(13)) & vbTab &
                                                DbAdapter1.validdec(myrecord(14)) & vbTab &
                                                DbAdapter1.validchar(myrecord(12)) & vbTab &
                                                DbAdapter1.validdec(myrecord(16)) & vbCrLf)
                        End If
                    End If
                Next


            End With
        End Using
        'update record
        Try
            Dim errmsg As String = String.Empty
            ProgressReport(6, "Marque")
            If vendorSB.Length > 0 Then
                ProgressReport(2, "Copy Vendor")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy vendor(vendorcode,vendorname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, vendorSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Vendor" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If cmmfpriceSB.Length > 0 Then
                ProgressReport(2, "Copy CMMFPrice")
                'cmmf,myyear,initailtx,initialprice,incoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2
                sqlstr = "copy cmmfprice(cmmf,myyear,initialtx,initialprice,invoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, cmmfpriceSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy CMMFPrice" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If updatecmmfpricesb.Length > 0 Then
                ProgressReport(2, "Update CMMFPrice")
                'lasttx,lastprice,invoiceverificationnumber2
                sqlstr = "update cmmfprice set lasttx= foo.lasttx::date,lastprice = foo.lastprice::numeric,invoiceverificationnumber2 = foo.invoiceverificationnumber2::bigint from (select * from array_to_set4(Array[" & updatecmmfpricesb.ToString &
                         "]) as tb (id character varying,lasttx character varying,lastprice character varying,invoiceverificationnumber2 character varying))foo where cpid = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(2, "Copy CMMFVendorPrice" & "::" & errmsg)
                    Exit Sub
                End If
            End If

            If cmmfvendorpriceSB.Length > 0 Then
                ProgressReport(2, "Copy CMMFVendorPrice")
                'cmmf,myyear,initailtx,initialprice,incoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2
                sqlstr = "copy cmmfvendorprice(cmmf,vendorcode,myyear,initialtx,initialprice,invoiceverificationnumber,lasttx,lastprice,invoiceverificationnumber2) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, cmmfvendorpriceSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy CMMFVendorPrice" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If updateCMMFvendorpriceLastsb.Length > 0 Then
                ProgressReport(2, "Update CMMFVendorPrice LastTx")
                'lasttx,lastprice,invoiceverificationnumber2
                sqlstr = "update cmmfvendorprice set lasttx= foo.lasttx::date,lastprice = foo.lastprice::numeric,invoiceverificationnumber2 = foo.invoiceverificationnumber2::bigint,agv2 = 0 from (select * from array_to_set4(Array[" & updateCMMFvendorpriceLastsb.ToString &
                         "]) as tb (id character varying,lasttx character varying,lastprice character varying,invoiceverificationnumber2 character varying) order by id,lasttx )foo where cpid = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(2, "Copy CMMFVendorPrice LastTx" & "::" & errmsg)
                    Exit Sub
                End If
            End If

            If updateCMMFvendorpriceInitsb.Length > 0 Then
                ProgressReport(2, "Update CMMFVendorPrice InitTx")
                'lasttx,lastprice,invoiceverificationnumber2
                sqlstr = "update cmmfvendorprice set inittx= foo.inittx::date,initialprice = foo.initialprice::numeric,invoiceverificationnumber = foo.invoiceverificationnumber::bigint from (select * from array_to_set4(Array[" & updateCMMFvendorpriceInitsb.ToString &
                         "]) as tb (id character varying,inittx character varying,initialprice character varying,invoiceverificationnumber character varying))foo where cpid = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(2, "Copy CMMFVendorPrice InitTx" & "::" & errmsg)
                    Exit Sub
                End If
            End If


            If MiroSB.Length > 0 Then
                ProgressReport(2, "Copy Miro")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy miro(mironumber,miropostingdate,supplierinvoicenum ,vendorcode) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, MiroSB.ToString.Replace("\", "\\"), myret)
                If Not myret Then
                    ProgressReport(2, "Copy Miro" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If POHDSB.Length > 0 Then
                ProgressReport(2, "Copy POHD")
                'pohd bigint, pono character varying,purchasinggroup character varying, payt character varying
                sqlstr = "copy pohd(pohd,pono,purchasinggroup) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, POHDSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy POHD" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If PODtlSB.Length > 0 Then
                ProgressReport(2, "Copy PODTL")
                'pohd bigint, polineno character varying,cmmf bigint,oun character varying
                sqlstr = "copy podtl(pohd,polineno,cmmf,oun) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PODtlSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy PODTL" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If POMiroSB.Length > 0 Then
                ProgressReport(2, "Copy POMIRO")
                'podtlid bigint,miroid bigint,amount numeric,qty numeric,crcy charcter varying,unitprice
                sqlstr = "copy pomiro(podtlid,miroid,amount,qty,crcy,unitprice) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, POMiroSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy POMiro" & "::" & errmessage)
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            ProgressReport(1, ex.Message)

        End Try
        ProgressReport(5, "Continue")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub

End Class