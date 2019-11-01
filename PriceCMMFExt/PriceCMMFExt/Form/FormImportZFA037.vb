Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass

Public Class FormImportZFA037

    Dim PricelistSB As StringBuilder
    Dim VendorSB As StringBuilder
    Dim VendorUpdateSB As StringBuilder
    Dim CMMFSB As StringBuilder
    Dim ProjectSB As StringBuilder
    Dim PCRangeSB As StringBuilder
    Dim priceplntSB As StringBuilder
    Dim priceplntScaleSB As StringBuilder
    Dim PCCMMFSB As Object
    Dim PriceListIdTemp As Long

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Private Property FileName As String

    Dim myThreadDelegate As New ThreadStart(AddressOf dowork)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date

    'Dim miroSeq As Long
    'Dim podtlseq As Long
    'Dim cmmfpriceseq As Long
    'Dim cmmfvendorpriceseq As Long

    Dim pcprojectSeq As Long
    Dim pcprojectId As Long
    Dim pcrangeSeq As Long
    Dim pcrangeId As Long
    Dim pricelistid As Long

    Private DS As DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
            OpenFileDialog1.FileName = ""
            'Get file
            If OpenFileDialog1.ShowDialog = DialogResult.OK Then

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
        Dim sw2 As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        'Dim enddate As Date
        Dim SavingLookupSB As New System.Text.StringBuilder

        Dim myrecord() As String
        Dim mylist As New List(Of String())
        'Dim typeid As Long
        'Dim actionid As Long
        Dim sqlstr As String = String.Empty

        Dim mymessage As String = String.Empty
        sw.Start()
        Try
            If DbAdapter1.getproglock("FCMMFPL", HelperClass1.UserInfo.DisplayName, 1) Then
                ProgressReport(2, "This Program is being used by other person")
                Exit Sub
            End If
        Catch ex As Exception
            ProgressReport(2, ex.Message)
            Exit Sub
        End Try
        
        'delete existing data

        ProgressReport(2, String.Format("Delete Pricelist ..........."))

        sqlstr = "Delete from pricelist;" & _
                 "Delete from priceplant;" & _
                 "Delete from priceplantscale;" & _
                 "select setval('pricelist_pricelistid_seq'::regclass,1,false);" &
                 "select setval('priceplant_priceplantid_seq'::regclass,1,false);" &
                 "select setval('priceplantscale_id_seq'::regclass,1,false);"
        sw2.Start()
        If Not DbAdapter1.ExecuteNonQuery(sqlstr, message:=mymessage) Then
            ProgressReport(2, mymessage)
            Exit Sub
        End If
        pricelistid = 0
        ProgressReport(2, String.Format("Delete Pricelist ..........."))
        sw2.Stop()
        ProgressReport(1, String.Format("Delete Pricelist Done. Elapsed Time: {0}:{1}.{2}", Format(sw2.Elapsed.Minutes, "00"), Format(sw2.Elapsed.Seconds, "00"), sw2.Elapsed.Milliseconds.ToString))

        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True

                Dim count As Long = 0

                'FillData
                ProgressReport(2, "Initialize Table..")

                Dim sb As New StringBuilder
                'sb.Append("select cmmf,vendorcode,validfrom::character varying,validto::character varying from pricelist;")
                sb.Append("select cmmf,vendorcode,to_char(validfrom,'dd.MM.YYYY'),to_char(validto,'dd.MM.YYYY'),pricelistid from pricelist;")
                sb.Append("select vendorcode,vendorname::character varying,officerid::character varying,ssmidpl,pmid from vendor;")
                sb.Append("select cmmf from cmmf;")
                sb.Append("select cmmf,pcrangeid from pccmmf;")
                sb.Append(" with pb as (select min(pcp.pcprojectid) as pcprojectid,ssmid,spmid from pcproject pcp " &
                          " where projectname isnull and familyid isnull" &
                          " group by ssmid,spmid" &
                          " order by ssmid,spmid)" &
                          " select pb.ssmid,pb.spmid,pb.pcprojectid,pc.pcrangeid from pb" &
                          " left join pcrange pc on pc.pcprojectid = pb.pcprojectid" &
                          " where rangename isnull and imagepath isnull" &
                          " order by ssmid,spmid,pb.pcprojectid;")
                sb.Append("select nextval('pcproject_pcprojectid_seq');")
                sb.Append("select nextval('pcrange_rangeid_seq');")
                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "Pricelist"
                Dim idx0(3) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                idx0(1) = DS.Tables(0).Columns(1)
                idx0(2) = DS.Tables(0).Columns(2)
                idx0(3) = DS.Tables(0).Columns(3)

                DS.Tables(0).PrimaryKey = idx0

                DS.Tables(1).TableName = "Vendor"
                Dim idx1(0) As DataColumn
                idx1(0) = DS.Tables(1).Columns(0)
                DS.Tables(1).PrimaryKey = idx1


                DS.Tables(2).TableName = "CMMF"
                Dim idx2(0) As DataColumn
                idx2(0) = DS.Tables(2).Columns(0)
                DS.Tables(2).PrimaryKey = idx2

                DS.Tables(3).TableName = "PCCMMF"
                Dim idx3(0) As DataColumn
                idx3(0) = DS.Tables(3).Columns(0)
                DS.Tables(3).PrimaryKey = idx3

                'DS.Tables(4).TableName = "PCRange"
                'Dim idx4(1) As DataColumn
                'idx4(0) = DS.Tables(4).Columns(0)
                'idx4(1) = DS.Tables(4).Columns(1)
                'DS.Tables(4).PrimaryKey = idx4


                DS.Tables(4).TableName = "Project"
                Dim idx4(1) As DataColumn
                idx4(0) = DS.Tables(4).Columns(0)
                idx4(1) = DS.Tables(4).Columns(1)
                DS.Tables(4).PrimaryKey = idx4

                pcprojectSeq = DS.Tables(5).Rows(0).Item(0)
                pcrangeSeq = DS.Tables(6).Rows(0).Item(0)

                ProgressReport(2, "Read Text File...")
                Try
                    Do Until .EndOfData
                        myrecord = .ReadFields
                        If count = 298580 Then
                            Debug.Print("debug")
                        End If
                        If count > 1 Then
                            mylist.Add(myrecord)
                        End If
                        count += 1
                    Loop
                Catch ex As Exception
                    ProgressReport(1, ex.Message)
                    Exit Sub
                End Try

                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                PricelistSB = New StringBuilder
                VendorSB = New StringBuilder
                VendorUpdateSB = New StringBuilder
                CMMFSB = New StringBuilder
                ProjectSB = New StringBuilder
                PCRangeSB = New StringBuilder
                priceplntSB = New StringBuilder
                priceplntScaleSB = New StringBuilder
                PCCMMFSB = New StringBuilder
                Dim i As Long
                sw2.Start()

                Try
                    For i = 0 To mylist.Count - 1
                        'find the record in existing table.
                        ProgressReport(7, i + 1 & "," & mylist.Count)
                        myrecord = mylist(i)
                        If i = 352461 Then
                            Debug.Print("hello")
                        End If
                        If i >= 0 Then
                            If Not IsNumeric(myrecord(4)) Then
                                'skip cmmf with characters
                                Debug.Print("skip")
                            Else
                                Dim result As DataRow
                                'Find Vendor if not avail then create

                                Dim pkey1(0) As Object
                                pkey1(0) = myrecord(2)
                                Dim spmid = ""
                                Dim ssmidpl = ""
                                result = DS.Tables(1).Rows.Find(pkey1)
                                If IsNothing(result) Then
                                    'create
                                    'vendorcode,vendorname,officerid
                                    Dim dr As DataRow = DS.Tables(1).NewRow
                                    dr.Item(0) = myrecord(2)
                                    dr.Item(1) = myrecord(3)
                                    dr.Item(2) = myrecord(6)
                                    DS.Tables(1).Rows.Add(dr)

                                    VendorSB.Append(myrecord(2) & vbTab &
                                                    validstr(myrecord(3)) & vbTab &
                                                    validlong(myrecord(6)) & vbCrLf)
                                Else
                                    'update
                                    If Not IsDBNull(result.Item("pmid")) Then
                                        spmid = result.Item("pmid")
                                    End If
                                    If Not IsDBNull(result.Item("ssmidpl")) Then
                                        ssmidpl = result.Item("ssmidpl")
                                    End If

                                    Dim update As Boolean = False
                                    If result.Item("vendorname").ToString <> myrecord(3) Then
                                        If myrecord(3) <> "Ya Horng Electronic Co., Ltd." Then
                                            ''Debug.Print("ya horng")
                                            ''Else
                                            'result.Item("vendorname") = myrecord(3)
                                            'update = True
                                        End If

                                    End If
                                    If IsDBNull(result.Item("officerid")) Then
                                        result.Item("officerid") = myrecord(6)
                                        update = True
                                    Else
                                        If result.Item("officerid") <> myrecord(6) Then
                                            result.Item("officerid") = myrecord(6)
                                            update = True
                                        End If
                                    End If


                                    If update Then
                                        If VendorUpdateSB.Length > 0 Then
                                            VendorUpdateSB.Append(",")
                                        End If
                                        VendorUpdateSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying]", myrecord(2), validstr(myrecord(3)), validlong(myrecord(6))))
                                    End If

                                End If


                                'Find CMMF 2
                                Dim pkey2(0) As Object
                                pkey2(0) = myrecord(4)

                                result = DS.Tables(2).Rows.Find(pkey2)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables(2).NewRow
                                    dr.Item(0) = myrecord(4)
                                    DS.Tables(2).Rows.Add(dr)
                                    'cmmf,materialdesc,vendorcode,plnt 
                                    '** this is not true anymore
                                    'cmmf -> vendorcode -> plnt  not one to one relations
                                    'only for initial value purpose
                                    'the correct relation for cmmf -> vendorcode is table pricelist
                                    'cmmf->plant priceplant
                                    CMMFSB.Append(myrecord(4) & vbTab &
                                                        validstr(myrecord(5)) & vbTab &
                                                        validreal(myrecord(2)) & vbTab &
                                                        validint(myrecord(1)) & vbCrLf)
                                End If

                                'Find PCCMMF 3
                                Dim pkey3(0) As Object
                                pkey3(0) = myrecord(4)

                                result = DS.Tables(3).Rows.Find(pkey3)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables(3).NewRow
                                    dr.Item(0) = myrecord(4)

                                    Dim myrangeid = "Null"
                                    If myrecord(1) = "3720" Then
                                        'Find Project with ssmid and pmid
                                        'if not avail then create pcrange,project
                                        If Not (ssmidpl = "" Or spmid = "") Then 'New ss
                                            Dim pkey4(1) As Object
                                            pkey4(0) = ssmidpl
                                            pkey4(1) = spmid
                                            result = DS.Tables(4).Rows.Find(pkey4)
                                            If IsNothing(result) Then
                                                pcprojectSeq = pcprojectSeq + 1
                                                pcrangeSeq = pcrangeSeq + 1
                                                pcprojectId = pcprojectSeq
                                                pcrangeId = pcrangeSeq
                                                'create pcproject for ssm an spm
                                                ProjectSB.Append(pcprojectId & vbTab &
                                                                 ssmidpl & vbTab &
                                                                 spmid & vbCrLf)



                                                'create pcrange for new pcproject
                                                PCRangeSB.Append(pcrangeId & vbTab &
                                                                 pcprojectId & vbCrLf)
                                                'create table(5)
                                                Dim dr1 As DataRow = DS.Tables(4).NewRow
                                                dr1.Item("ssmid") = ssmidpl
                                                dr1.Item("spmid") = spmid
                                                dr1.Item("pcprojectid") = pcprojectId
                                                dr1.Item("pcrangeid") = pcrangeId
                                                myrangeid = pcrangeId
                                                'pcprojectSeq = pcprojectSeq + 1
                                                'pcrangeSeq = pcrangeSeq + 1
                                                DS.Tables(4).Rows.Add(dr1)
                                            Else
                                                myrangeid = result.Item("pcrangeid")
                                            End If
                                        End If
                                    End If


                                    DS.Tables(3).Rows.Add(dr)
                                    PCCMMFSB.Append(myrecord(4) & vbTab &
                                                  myrangeid & vbCrLf)

                                End If

                                'Find Price List if not avail then add to PricelistSB
                                If pricelistid = 210690 Then
                                    Debug.Print("hello")
                                End If

                                Dim pkey0(3) As Object
                                pkey0(0) = myrecord(4) 'cmmf
                                pkey0(1) = myrecord(2) 'vendor
                                pkey0(2) = myrecord(14) 'valid from
                                pkey0(3) = myrecord(15) 'valid to

                                result = DS.Tables(0).Rows.Find(pkey0)
                                If IsNothing(result) Then
                                    Dim dr As DataRow = DS.Tables(0).NewRow
                                    pricelistid = pricelistid + 1
                                    PriceListIdTemp = pricelistid
                                    dr.Item(0) = myrecord(4)
                                    dr.Item(1) = myrecord(2)
                                    dr.Item(2) = myrecord(14)
                                    dr.Item(3) = myrecord(15)
                                    dr.Item(4) = pricelistid
                                    DS.Tables(0).Rows.Add(dr)

                                    'cmmf,scaleqty,amount,perunit,validfrom,validto,vendorcode,currency)    
                                    PricelistSB.Append(myrecord(4) & vbTab &
                                                        validreal(myrecord(9)) & vbTab &
                                                        validreal(myrecord(10)) & vbTab &
                                                        validint(myrecord(12)) & vbTab &
                                                        dateformatdotyyyymmdd(myrecord(14)) & vbTab &
                                                        dateformatdotyyyymmdd(myrecord(15)) & vbTab &
                                                        validint(myrecord(2)) & vbTab &
                                                        validstr(myrecord(11)) & vbTab &
                                                        validstr(myrecord(13)) & vbCrLf)
                                    priceplntSB.Append(pricelistid & vbTab &
                                                       myrecord(1) & vbCrLf)
                                Else
                                    PriceListIdTemp = result.Item("pricelistid")
                                    'priceplntSB.Append(PriceListIdTemp & vbTab &
                                    '                  myrecord(1) & vbCrLf)
                                End If

                                priceplntScaleSB.Append(PriceListIdTemp & vbTab &
                                                            myrecord(1) & vbTab &
                                                            validreal(myrecord(12)) & vbTab &
                                                            validreal(myrecord(10)) & vbTab &
                                                            validreal(myrecord(9)) & vbCrLf)
                            End If
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

            If VendorSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy Vendor"))
                sqlstr = "copy vendor(vendorcode,vendorname,officerid)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, VendorSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy Vendor")


                End If

            End If

            If VendorUpdateSB.Length > 0 Then
                ProgressReport(2, "Update Vendor")
                'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
                sqlstr = "update vendor set vendorname= foo.vendorname,officerid = foo.officerid::bigint from (select * from array_to_set3(Array[" & VendorUpdateSB.ToString &
                         "]) as tb (id character varying,vendorname character varying,officerid character varying))foo where vendorcode = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(1, errmsg)

                    myret = False
                    Err.Raise(513, Description:=errmsg & " ::Update Vendor")
                End If
            End If

            If CMMFSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy CMMF"))
                sqlstr = "copy cmmf(cmmf,materialdesc,vendorcode,plnt)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, CMMFSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy CMMF")
                End If

            End If

            If ProjectSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy Project"))
                sqlstr = "copy pcproject(pcprojectid,ssmid,spmid)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, ProjectSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy Project")
                End If

            End If

            If PCRangeSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy ProjectRange"))
                sqlstr = "copy pcrange(pcrangeid,pcprojectid)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, PCRangeSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy PCRange")
                End If

            End If
            If PCCMMFSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy PCCMMF"))
                sqlstr = "copy pccmmf(cmmf,pcrangeid)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, PCCMMFSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy PCCMMF")
                End If

            End If
            If PricelistSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy Pricelist"))
                sqlstr = "copy pricelist(cmmf,scaleqty,amount,perunit,validfrom,validto,vendorcode,currency,uom)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, PricelistSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy Pricelist")
                End If

            End If
            If priceplntSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy PricePlant"))
                sqlstr = "copy priceplant(pricelistid,plant)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, priceplntSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy PricePlant")
                End If

            End If
            If priceplntScaleSB.ToString <> "" Then
                ProgressReport(2, String.Format("Copy PricePlantScale"))
                sqlstr = "copy priceplantscale(pricelistid,plant,perunit,amount,scale)  from stdin with null as 'Null';"
                errmsg = DbAdapter1.copy(sqlstr, priceplntScaleSB.ToString, myret)
                If Not myret Then
                    ProgressReport(1, errmsg)
                    Err.Raise(513, Description:=errmsg & " ::Copy PricePlantScale")
                End If

            End If
            myret = True
        Catch ex As Exception

            ProgressReport(1, ex.Message)

        End Try
        sw2.Stop()
        ProgressReport(1, String.Format("Copy Done. Elapsed Time: {0}:{1}.{2}", Format(sw2.Elapsed.Minutes, "00"), Format(sw2.Elapsed.Seconds, "00"), sw2.Elapsed.Milliseconds.ToString))
        ProgressReport(5, "Continue")
        sw.Stop()
        If myret Then
            DbAdapter1.getproglock("FCMMFPL", HelperClass1.UserInfo.DisplayName, 0)
            ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        Else
            ProgressReport(2, String.Format("Error. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
        End If


    End Sub




End Class







