Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormImportPO40Plus

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim startdate As Date
    Dim enddate As Date

    Dim miroSeq As Long
    Dim podtlseq As Long
    Dim cmmfpriceseq As Long
    Dim cmmfvendorpriceseq As Long

    Private DS As DataSet

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Start Thread
        If Not myThread.IsAlive Then
            Me.ToolStripStatusLabel1.Text = ""
            Me.ToolStripStatusLabel2.Text = ""
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
    Sub dowork()
        Dim sw As New Stopwatch
        Dim mylist As New List(Of String())
        Dim sqlstr As String = String.Empty
        Dim myrecord() As String
        Dim obj As Object
        Dim errormsg As String = String.Empty
        Using objTFParser = New FileIO.TextFieldParser(OpenFileDialog1.FileName)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0
                Dim mymessage As String

                sw.Start()

                'FillData
                ProgressReport(2, "Initialize Table..")
                sqlstr = "select vendorcode from vendor;" &
                         "select brandid from brand;" &
                         "select familyid from family;" &
                         "select loadingcode from loading;" &
                         "select cmmf from cmmf;" &
                         "select range,rangeid from range;" &
                         "select agreement,material,postingdate from agreementtx;" &
                         "select agreement from agvalue;" &
                         "select miropostingdate,mironumber,pohd,polineno from poreversed;" &
                         "select range,rangeid from range order by rangeid desc limit 1;" &
                         "select vendorcode,vendorname,shortname2 from vendor where vendorcode = 0;" &
                         "select cmmf,plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid from cmmf;" &
                         "select pohd,payt from pohd;" &
                         "select salesdoc,shiptoparty from aasdhd;" &
                         "select salesdoc,salesdocitem,pohd,poitem from aasdpo;"

                mymessage = String.Empty
                DS = New DataSet
                DS.CaseSensitive = True
                If Not DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "Vendor"
                Dim idx0(0) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                DS.Tables(0).PrimaryKey = idx0

                DS.Tables(1).TableName = "Brand"
                Dim idx1(0) As DataColumn
                idx1(0) = DS.Tables(1).Columns(0)
                DS.Tables(1).PrimaryKey = idx1

                DS.Tables(2).TableName = "Family"
                Dim idx2(0) As DataColumn
                idx2(0) = DS.Tables(2).Columns(0)
                DS.Tables(2).PrimaryKey = idx2

                DS.Tables(3).TableName = "LoadingCode"
                Dim idx3(0) As DataColumn
                idx3(0) = DS.Tables(3).Columns(0)
                DS.Tables(3).PrimaryKey = idx3
                
                DS.Tables(4).TableName = "cmmf"
                Dim idx4(0) As DataColumn
                idx4(0) = DS.Tables(4).Columns(0)
                DS.Tables(4).PrimaryKey = idx4

                DS.Tables(5).TableName = "Range"
                Dim idx5(0) As DataColumn
                idx5(0) = DS.Tables(5).Columns(0)
                DS.Tables(5).PrimaryKey = idx5

                DS.Tables(6).TableName = "AgreementTx"
                Dim idx6(2) As DataColumn
                idx6(0) = DS.Tables(6).Columns(0)
                idx6(1) = DS.Tables(6).Columns(1)
                idx6(2) = DS.Tables(6).Columns(2)
                DS.Tables(6).PrimaryKey = idx6

                DS.Tables(7).TableName = "AgValue"
                Dim idx7(0) As DataColumn
                idx7(0) = DS.Tables(7).Columns(0)
                DS.Tables(7).PrimaryKey = idx7

                DS.Tables(8).TableName = "PoReserved"
                Dim idx8(3) As DataColumn
                idx8(0) = DS.Tables(8).Columns(0)
                idx8(1) = DS.Tables(8).Columns(1)
                idx8(2) = DS.Tables(8).Columns(2)
                idx8(3) = DS.Tables(8).Columns(3)
                DS.Tables(8).PrimaryKey = idx8

                DS.Tables(9).TableName = "UpdateVendor"
                Dim idx9(0) As DataColumn
                idx9(0) = DS.Tables(9).Columns(0)
                DS.Tables(9).PrimaryKey = idx9

                DS.Tables(10).TableName = "UpdateCMMF"
                Dim idx10(0) As DataColumn
                idx10(0) = DS.Tables(10).Columns(0)
                DS.Tables(10).PrimaryKey = idx10

                DS.Tables(11).TableName = "UpdatePOHD"
                Dim idx11(0) As DataColumn
                idx11(0) = DS.Tables(11).Columns(0)
                DS.Tables(11).PrimaryKey = idx11

                DS.Tables(13).TableName = "ShipToParty"
                Dim idx13(0) As DataColumn
                idx13(0) = DS.Tables(13).Columns(0)
                DS.Tables(13).PrimaryKey = idx13

                DS.Tables(14).TableName = "RelSalesDocPO"
                Dim idx14(3) As DataColumn
                idx14(0) = DS.Tables(14).Columns(0)
                idx14(1) = DS.Tables(14).Columns(1)
                idx14(2) = DS.Tables(14).Columns(2)
                idx14(3) = DS.Tables(14).Columns(3)
                DS.Tables(14).PrimaryKey = idx14


                ProgressReport(2, "Read Text File...")
                ProgressReport(6, "Continuous")
                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop
                ProgressReport(2, "Build Record...")
                ProgressReport(5, "Continuous")

                If mylist(0).Length = 44 Then
                    'Build POFG
                    obj = New POFG(DS, mylist)
                Else
                    'Build POComp
                    obj = New POComp(DS, mylist)
                End If
                Dim myevent = New POEventArgs

                If Not obj.BuildSb(Me, myevent) Then
                    ProgressReport(1, myevent.mymessage)
                    Exit Sub
                End If

            End With
        End Using
        'Update Record        
        Try
            ProgressReport(6, "Marque")
            If obj.vendorSB.Length > 0 Then
                ProgressReport(2, "Copy Vendor")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy vendor(vendorcode,vendorname,shortname2) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.vendorSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Vendor" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.updatevendorsb.Length > 0 Then
                ProgressReport(2, "Update Vendor")
                'vendorname,shortname2
                sqlstr = "update vendor set vendorname= foo.vendorname,shortname2 = foo.shortname2 from (select * from array_to_set3(Array[" & obj.updatevendorsb.ToString &
                         "]) as tb (id character varying,vendorname character varying,shortname2 character varying))foo where vendorcode = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errormsg) Then
                    ProgressReport(2, "Update Vendor" & "::" & errormsg)
                    Exit Sub
                End If
            End If
            If obj.brandsb.Length > 0 Then
                ProgressReport(2, "Copy Brand")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy brand(brandid,brandname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.brandsb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Brand" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If obj.familysb.Length > 0 Then
                ProgressReport(2, "Copy Family")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy family(familyid,familyname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.familysb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Family" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.loadingsb.Length > 0 Then
                ProgressReport(2, "Copy Loading")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy loading(loadingcode,loadingname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.loadingsb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Loading" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.rangesb.Length > 0 Then
                ProgressReport(2, "Copy Range")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy range(range,rangedesc) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.rangesb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Range" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.cmmfsb.Length > 0 Then
                ProgressReport(2, "Copy CMMF")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy cmmf(cmmf,plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.cmmfsb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy CMMF" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If obj.updatecmmfsb.Length > 0 Then
                ProgressReport(2, "Update CMMF")
                'plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                sqlstr = "update cmmf set plnt= foo.plnt::integer,brandid=foo.brandid::integer,cmmftype=foo.cmmftype,comfam=foo.comfam::integer,rir=foo.rir::character(2),activitycode=foo.activitycode::character(2),modelcode=foo.modelcode,rangeid=foo.rangeid::bigint from (select * from array_to_set9(Array[" & obj.updatecmmfsb.ToString &
                         "]) as tb (id character varying,plnt character varying,brandid character varying,cmmftype character varying,comfam character varying,rir character varying,activitycode character varying,modelcode character varying,rangeid character varying))foo where cmmf = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errormsg) Then
                    ProgressReport(2, "Update CMMF" & "::" & errormsg)
                    Exit Sub
                End If
            End If

            If obj.agreementsb.Length > 0 Then
                ProgressReport(2, "Copy Agreement")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy agreementtx(agreement,material,postingdate) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.agreementsb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Agreement" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.agvaluesb.Length > 0 Then
                ProgressReport(2, "Copy AgreementValue")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy agvalue(agreement) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.agvaluesb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy AgreementValue" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            'No need this information can be take from EKKO
            'If obj.updatepohdsb.Length > 0 Then
            '    ProgressReport(2, "Update POHD")
            '    'plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
            '    sqlstr = "update pohd set payt = foo.payt from (select * from array_to_set2(Array[" & obj.updatepohdsb.ToString &
            '             "]) as tb (id character varying,payt character varying))foo where pohd = foo.id::bigint;"
            '    Dim ra As Long
            '    If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errormsg) Then
            '        ProgressReport(2, "Update POHD" & "::" & errormsg)
            '        Exit Sub
            '    End If
            'End If
            If obj.salesdocsb.Length > 0 Then
                ProgressReport(2, "Copy SalesDoc")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy aasdhd(salesdoc,custpo,creationdate,soldtoparty,shiptoparty) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.salesdocsb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy SalesDoc" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.updatesalesdocsb.Length > 0 Then
                ProgressReport(2, "Update SalesDoc")
                'plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                sqlstr = "update aasdhd set shiptoparty = foo.shiptoparty::bigint from (select * from array_to_set2(Array[" & obj.updatesalesdocsb.ToString &
                         "]) as tb (id character varying,shiptoparty character varying))foo where salesdoc = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errormsg) Then
                    ProgressReport(2, "Update SalesDoc" & "::" & errormsg)
                    Exit Sub
                End If
            End If


            If obj.relsalesdocposb.Length > 0 Then
                ProgressReport(2, "Copy RelSalesDocPOSB")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy aasdpo(salesdoc,salesdocitem,pohd,poitem) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.relsalesdocposb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy RelSalesDocPOSB" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If obj.PoReversedSB.Length > 0 Then
                ProgressReport(2, "Copy POReversed")
                'miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reservedby,cmmf,plnt,amount,crcy,dc,
                'qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir)

                sqlstr = "copy poreversed(miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reversedby,cmmf,plnt,amount,crcy,dc,qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, obj.poreversedsb.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy POReversed" & "::" & errmessage)
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            ProgressReport(1, ex.Message)
        End Try
        ProgressReport(5, "Continuous")
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub



    Sub DoWork1()
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
                         "delete from miro m " &
                         "where m.miropostingdate >= " & DateFormatyyyyMMdd(startdate) & " and m.miropostingdate <= " & DateFormatyyyyMMdd(enddate) & ";" &
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
                                        updateCMMFvendorpriceLastsb.Append(String.Format("['{0}'::character varying,{1}::character varying,'{2}'::character varying,'{3}'::character varying]",
                                                      result.Item(5), DbAdapter1.dateformatdot(myrecord(11)), DbAdapter1.validdec(myrecord(16)), DbAdapter1.validlong(myrecord(10))))

                                    ElseIf result.Item(4) > postingdate Then
                                        'initialtx,initialprice,invoiceverificationnumber
                                        If updateCMMFvendorpriceInitsb.Length > 0 Then
                                            updateCMMFvendorpriceInitsb.Append(",")
                                        End If
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

            If vendorSB.Length > 0 Then
                ProgressReport(2, "Copy Miro")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy vendor(vendorcode,vendorname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, vendorSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Miro" & "::" & errmessage)
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
                         "]) as tb (id character varying,lasttx character varying,lastprice character varying,invoiceverificationnumber2 character varying))foo where cpid = foo.id::bigint;"
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
                errmessage = DbAdapter1.copy(sqlstr, MiroSB.ToString, myret)
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

        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

    End Sub
End Class

Interface iPO
    Function BuildSb(ByVal sender As Object, ByVal e As POEventArgs)
End Interface

Class POFG
    Inherits BasePO
    Implements iPO

    Public Sub New(ByVal DS As DataSet, ByVal myList As List(Of String()))
        MyBase.New(DS, myList)
    End Sub

    Public Function BuildSb(ByVal sender As Object, ByVal e As POEventArgs) As Object Implements iPO.BuildSb
        Dim myret As Boolean = False
        Parent = sender
        Dim myrecord As String()
        Dim result As DataRow
        Dim rangeIdSeq As Long = DS.Tables(9).Rows(0).Item(1)
        Dim recid As Long
        Dim lastTable As String = String.Empty
        Try
            For i = 0 To MyList.Count - 1
                'find the record in existing table.
                ProgressReport(7, i + 1 & "," & MyList.Count)
                If i >= 1 Then
                    myrecord = MyList(i)
                    recid = i
                    If myrecord(26) <> "" Then
                        'agreement
                        lastTable = "Agreement"
                        Dim pkey6(2) As Object
                        pkey6(0) = myrecord(26)
                        pkey6(1) = myrecord(13)
                        pkey6(2) = DbAdapter1.dateformatdotdate(myrecord(1))

                        result = DS.Tables(6).Rows.Find(pkey6)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(6).NewRow
                            dr.Item(0) = myrecord(26)
                            dr.Item(1) = myrecord(13)
                            dr.Item(2) = pkey6(2)
                            DS.Tables(6).Rows.Add(dr)
                            'agreement,material,postingdate
                            AgreementSB.Append(validlong(myrecord(26)) & vbTab &
                                            validlong(myrecord(13)) & vbTab & dateformatdotyyyymmdd(myrecord(1)) & vbCrLf)

                            'add agv
                            Dim pkey7(0) As Object
                            pkey7(0) = myrecord(26)
                            result = DS.Tables(7).Rows.Find(pkey7)
                            If IsNothing(result) Then
                                dr = DS.Tables(7).NewRow
                                dr.Item(0) = myrecord(26)
                                DS.Tables(7).Rows.Add(dr)
                                AgValueSB.Append(validlong(myrecord(26)) & vbCrLf)
                            End If
                        End If
                    End If

                    If myrecord(11) = "" Then 'not reversal tx
                        'Vendor Table
                        'Check Vendorcode
                        lastTable = "Vendor"
                        Dim pkey0(0) As Object
                        pkey0(0) = myrecord(21)
                        result = DS.Tables(0).Rows.Find(pkey0)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(21)
                            DS.Tables(0).Rows.Add(dr)
                            'vendorcode,vendorname,shortname2
                            VendorSB.Append(validlong(myrecord(21)) & vbTab &
                                            validstr(myrecord(22)) & vbTab &
                                           validstr(myrecord(23)) & vbCrLf)
                        Else
                            Dim pkey10(0) As Object
                            pkey10(0) = myrecord(21)
                            result = DS.Tables(10).Rows.Find(pkey10)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(10).NewRow
                                dr.Item(0) = myrecord(21)
                                dr.Item(1) = myrecord(22)
                                dr.Item(2) = myrecord(23)
                                DS.Tables(10).Rows.Add(dr)
                            Else
                                result.Item(1) = myrecord(22)
                                result.Item(2) = myrecord(23)
                            End If
                        End If

                        'Check Brand
                        lastTable = "Brand"
                        Dim brandid = "Null"
                        If myrecord(31) <> "" Then
                            brandid = myrecord(31)
                            Dim pkey1(0) As Object
                            pkey1(0) = myrecord(31)
                            result = DS.Tables(1).Rows.Find(pkey1)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(1).NewRow
                                dr.Item(0) = myrecord(31)
                                DS.Tables(1).Rows.Add(dr)
                                'vendorcode,vendorname
                                BrandSB.Append(validint(myrecord(31)) & vbTab &
                                               validstr(myrecord(31)) & vbCrLf)
                            End If
                        End If

                        'Check Family
                        If myrecord(33) <> "" Then

                            lastTable = "Family"
                            Dim pkey2(0) As Object
                            pkey2(0) = myrecord(33)
                            result = DS.Tables(2).Rows.Find(pkey2)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(2).NewRow
                                dr.Item(0) = myrecord(33)
                                DS.Tables(2).Rows.Add(dr)
                                'vendorcode,vendorname
                                BrandSB.Append(validint(myrecord(33)) & vbTab &
                                               validstr(myrecord(33)) & vbCrLf)
                            End If
                        End If
                        'Check Loading
                        lastTable = "Loading"
                        Dim pkey3(0) As Object
                        pkey3(0) = myrecord(28)
                        result = DS.Tables(3).Rows.Find(pkey3)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(3).NewRow
                            dr.Item(0) = myrecord(28)
                            DS.Tables(3).Rows.Add(dr)
                            '
                            LoadingSB.Append(validint(myrecord(28)) & vbTab &
                                            validstr(myrecord(28)) & vbCrLf)
                        End If
                        Dim myrangeid As String = ""

                        If myrecord(41) <> "" Then
                            'Check Range
                            lastTable = "Range"
                            Dim pkey5(0) As Object
                            pkey5(0) = myrecord(41)
                            result = DS.Tables(5).Rows.Find(pkey5)
                            If IsNothing(result) Then
                                rangeIdSeq += 1
                                Dim dr As DataRow = DS.Tables(5).NewRow
                                dr.Item(0) = myrecord(41)
                                dr.Item(1) = rangeIdSeq
                                myrangeid = rangeIdSeq
                                DS.Tables(5).Rows.Add(dr)
                                myrangeid = rangeIdSeq
                                'range
                                RangeSB.Append(validstr(myrecord(41)) & vbTab &
                                               validstr(myrecord(41)) & vbCrLf)
                            Else
                                myrangeid = result.Item(1)
                            End If
                        End If

                        If myrecord(13) <> "" Then
                            'Check CMMF
                            lastTable = "CMMF"
                            'Dim pkey4(0) As Object
                            'pkey4(0) = myrecord(13)
                            'result = DS.Tables(4).Rows.Find(pkey4)

                            'If IsNothing(result) Then
                            '    Dim dr As DataRow = DS.Tables(4).NewRow
                            '    dr.Item(0) = myrecord(13)
                            '    DS.Tables(4).Rows.Add(dr)
                            '    'cmmf,plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                            '    CMMFSB.Append(validlong(myrecord(13)) & vbTab &
                            '                     validint(myrecord(14)) & vbTab &
                            '                     validlong(myrecord(21)) & vbTab &
                            '                     validstr(myrecord(28)) & vbTab &
                            '                     validint(myrecord(31)) & vbTab &
                            '                     validstr(myrecord(32)) & vbTab &
                            '                     validint(myrecord(33)) & vbTab &
                            '                     validstr(myrecord(34)) & vbTab &
                            '                     validstr(myrecord(34)) & vbTab &
                            '                     validstr(myrecord(40)) & vbTab &
                            '                     myrangeid & vbCrLf)
                            'Else
                            'Check for update
                            'plnt character varying,vendorcode character varying,loadingcode character varying,brandid character varying,cmmftype character varying,comfam character varying,rir character varying,activitycode character varying,modelcode character varying,rangeid character varying
                            Dim pkey11(0) As Object
                            pkey11(0) = myrecord(13)
                            result = DS.Tables(11).Rows.Find(pkey11)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(11).NewRow
                                dr.Item(0) = myrecord(13)
                                dr.Item(1) = myrecord(14)
                                dr.Item(2) = myrecord(21)
                                dr.Item(3) = myrecord(28)
                                dr.Item(4) = DbAdapter1.validint(myrecord(31))
                                dr.Item(5) = myrecord(32)
                                dr.Item(6) = DbAdapter1.validint(myrecord(33))
                                dr.Item(7) = DbAdapter1.validchar(myrecord(34))
                                dr.Item(8) = DbAdapter1.validchar(myrecord(34))
                                dr.Item(9) = DbAdapter1.validchar(myrecord(40))
                                dr.Item(10) = DbAdapter1.validint(myrangeid)
                                DS.Tables(11).Rows.Add(dr)
                                CMMFSB.Append(validlong(myrecord(13)) & vbTab &
                                                 validint(myrecord(14)) & vbTab &
                                                 validlong(myrecord(21)) & vbTab &
                                                 validstr(myrecord(28)) & vbTab &
                                                 validint(myrecord(31)) & vbTab &
                                                 validstr(myrecord(32)) & vbTab &
                                                 validint(myrecord(33)) & vbTab &
                                                 validstr(myrecord(34)) & vbTab &
                                                 validstr(myrecord(34)) & vbTab &
                                                 validstr(myrecord(40)) & vbTab &
                                                 myrangeid & vbCrLf)
                            Else
                                Dim flag As Boolean = False
                                If IsDBNull(result.Item(1)) Or IsDBNull(result.Item(4)) Or _
                                    IsDBNull(result.Item(5)) Or IsDBNull(result.Item(6)) Or IsDBNull(result.Item(7)) _
                                    Or IsDBNull(result.Item(8)) Or IsDBNull(result.Item(9)) Or IsDBNull(result.Item(10)) Then
                                    flag = True
                                ElseIf Not (result.Item(1) = myrecord(14) AndAlso result.Item(1) = myrecord(14) AndAlso
                                    result.Item(4) = DbAdapter1.validint(myrecord(31)) AndAlso
                                    result.Item(5) = myrecord(32) AndAlso
                                    result.Item(6) = DbAdapter1.validint(myrecord(33)) AndAlso
                                    result.Item(7) = myrecord(34) AndAlso
                                    result.Item(8) = myrecord(34) AndAlso
                                    result.Item(9) = myrecord(40) AndAlso
                                    result.Item(10) = DbAdapter1.validint(myrangeid)) Then
                                    flag = True

                                End If
                                If flag Then
                                    Dim mybrandid As String = "Null"
                                    If Not myrecord(31) = "" Then
                                        mybrandid = myrecord(31)
                                    End If
                                    Dim mycomfam As String = "Null"
                                    If Not myrecord(33) = "" Then
                                        mycomfam = myrecord(33)
                                    End If
                                    Dim strrangeid As String = "Null"
                                    If Not myrangeid = "" Then
                                        strrangeid = myrangeid
                                    End If
                                    Dim rri As String = "Null"
                                    If Not myrecord(34) = "" Then
                                        rri = "'" & myrecord(34) & "'"
                                    End If
                                    Dim modelcode = "Null"
                                    If Not myrecord(40) = "" Then
                                        modelcode = "'" & myrecord(40) & "'"
                                    End If
                                    If UpdateCMMFSB.Length > 0 Then
                                        UpdateCMMFSB.Append(",")
                                    End If
                                    'cmmf,plnt,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                                    UpdateCMMFSB.Append(String.Format("[{0}::character varying,{1}::character varying,{2}::character varying,'{3}'::character varying,{4}::character varying,{5}::character varying,{6}::character varying,{7}::character varying,{8}::character varying]",
                                                   myrecord(13), myrecord(14), mybrandid,
                                                  validstr(myrecord(32)), mycomfam,
                                                  rri, rri, modelcode, strrangeid))

                                    result.Item(1) = myrecord(14)
                                    'result.Item(2) = myrecord(21)
                                    'result.Item(3) = myrecord(28)
                                    result.Item(4) = DbAdapter1.validint(myrecord(31))
                                    result.Item(5) = myrecord(32)
                                    result.Item(6) = DbAdapter1.validint(myrecord(33))
                                    result.Item(7) = myrecord(34)
                                    result.Item(8) = myrecord(34)
                                    result.Item(9) = myrecord(40)
                                    result.Item(10) = DbAdapter1.validint(myrangeid)
                                Else
                                    'Debug.Print("no update")
                                End If

                                
                            End If
                            'End If
                        End If


                        'ShipToParty
                        If myrecord(29) <> "" Then
                            'Check PayT
                            lastTable = "ShipToParty"

                            Dim pkey13(0) As Object
                            pkey13(0) = myrecord(29)
                            result = DS.Tables(13).Rows.Find(pkey13)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(13).NewRow
                                dr.Item(0) = myrecord(29)
                                dr.Item(1) = myrecord(38)
                                DS.Tables(13).Rows.Add(dr)
                                SalesDocSB.Append(myrecord(29) & vbTab &
                                                       myrecord(7) & vbTab &
                                                       dateformatdotyyyymmdd(myrecord(42)) & vbTab &
                                                       myrecord(35) & vbTab &
                                                        myrecord(38) & vbCrLf)
                            Else
                                Dim flag As Boolean = False
                                If IsDBNull(result.Item(1)) Then
                                    flag = True
                                ElseIf result.Item(1) <> myrecord(38) Then
                                    flag = True

                                End If
                                If flag Then
                                    If UpdateSalesDocSB.Length > 0 Then
                                        UpdateSalesDocSB.Append(",")
                                    End If

                                    UpdateSalesDocSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]",
                                           myrecord(29), validlong(myrecord(38))))

                                    result.Item(1) = myrecord(38)
                                End If

                            End If

                            lastTable = "RelSalesDocPo"
                            Dim pkey14(3) As Object
                            pkey14(0) = myrecord(29)
                            pkey14(1) = myrecord(30)
                            pkey14(2) = myrecord(8)
                            pkey14(3) = myrecord(9)
                            result = DS.Tables(14).Rows.Find(pkey14)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(14).NewRow
                                dr.Item(0) = myrecord(29)
                                dr.Item(1) = myrecord(30)
                                dr.Item(2) = myrecord(8)
                                dr.Item(3) = myrecord(9)
                                DS.Tables(14).Rows.Add(dr)
                                RelSalesDocPoSB.Append(myrecord(29) & vbTab &
                                                       myrecord(30) & vbTab &
                                                       myrecord(8) & vbTab &
                                                        myrecord(9) & vbCrLf)

                            End If

                        End If

                            'If myrecord(24) <> "" Then
                            '    'Check PayT
                            '    lastTable = "PayT"
                            '    If UpdatePOHDSB.Length > 0 Then
                            '        UpdatePOHDSB.Append(",")
                            '    End If
                            '    UpdatePOHDSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]",
                            '                       myrecord(8), validstr(myrecord(24))))


                            'End If
                        Else
                            'check poreversed
                            'Check Brand
                            lastTable = "POREVERSED"
                            Dim pkey8(3) As Object
                            pkey8(0) = DbAdapter1.dateformatdotdate(myrecord(1))
                            pkey8(1) = myrecord(3)
                            pkey8(2) = myrecord(8)
                            pkey8(3) = myrecord(9)
                            result = DS.Tables(8).Rows.Find(pkey8)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(8).NewRow
                                dr.Item(0) = DbAdapter1.dateformatdotdate(myrecord(1))
                                dr.Item(1) = myrecord(3)
                                dr.Item(2) = myrecord(8)
                                dr.Item(3) = myrecord(9)

                                DS.Tables(8).Rows.Add(dr)
                                'vendorcode,vendorname
                                Dim mydc = IIf(myrecord(18) = "H", -1, 1)
                                'miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reservedby,cmmf,plnt,amount,crcy,dc,
                                'qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir)
                                Try
                                    PoReversedSB.Append(dateformatdotyyyymmdd(myrecord(1)) & vbTab &
                                                        validlong(myrecord(3)) & vbTab &
                                                        validstr(myrecord(6)) & vbTab &
                                                        validstr(myrecord(7)) & vbTab &
                                                        validlong(myrecord(8)) & vbTab &
                                                        validint(myrecord(9)) & vbTab &
                                                        validlong(myrecord(11)) & vbTab &
                                                        validlong(myrecord(13)) & vbTab &
                                                        validint(myrecord(14)) & vbTab &
                                                        validreal(myrecord(15)) & vbTab &
                                                        validstr(myrecord(16)) & vbTab &
                                                        mydc & vbTab &
                                                        validreal(myrecord(19)) & vbTab &
                                                        validstr(myrecord(20)) & vbTab &
                                                        validlong(myrecord(21)) & vbTab &
                                                        validstr(myrecord(24)) & vbTab &
                                                        validstr(myrecord(25)) & vbTab &
                                                        validlong(myrecord(26)) & vbTab &
                                                        validlong(myrecord(29)) & vbTab &
                                                        validint(myrecord(30)) & vbTab &
                                                        validint(myrecord(31)) & vbTab &
                                                        validstr(myrecord(32)) & vbTab &
                                                        validint(myrecord(33)) & vbTab &
                                                        validstr(myrecord(34)) & vbCrLf)
                                Catch ex As Exception
                                    Debug.Print("hello")
                                End Try

                            End If
                        End If
                End If
            Next

            'get updatevendor
            For Each row As DataRow In DS.Tables(10).Rows
                If UpdateVendorSB.Length > 0 Then
                    UpdateVendorSB.Append(",")
                End If
                Dim shortname As String = "Null"
                If row.Item("shortname2") <> "" Then
                    shortname = row.Item("shortname2")
                End If
                UpdateVendorSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying]",
                                   row.Item(0), validstr(row.Item(1)), validstr(shortname)))
            Next

            'getupdatecmmf
            'For Each row As DataRow In DS.Tables(11).Rows
            '    lastTable = "UpdateCMMFSB"
            '    If UpdateCMMFSB.Length > 0 Then
            '        UpdateCMMFSB.Append(",")
            '    End If

            '    'plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
            '    Dim myrangeid As String = "Null"
            '    If Not IsDBNull(row.Item(10)) Then
            '        myrangeid = row.Item(10)
            '    End If
            '    Dim mybrandid As String = "Null"
            '    If Not IsDBNull(row.Item(4)) Then
            '        mybrandid = row.Item(4)
            '    End If
            '    Dim mycomfam As String = "Null"
            '    If Not IsDBNull(row.Item(6)) Then
            '        mycomfam = row.Item(6)
            '    End If
            '    Dim rri As String = "Null"
            '    If Not row.Item(7) = "" Then
            '        rri = "'" & row.Item(7) & "'"
            '    End If
            '    Dim modelcode = "Null"
            '    If Not row.Item(9) = "" Then
            '        modelcode = "'" & row.Item(9) & "'"
            '    End If
            '    UpdateCMMFSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,{4}::character varying,'{5}'::character varying,{6}::character varying,{7}::character varying,{8}::character varying,{9}::character varying,{10}::character varying]",
            '                                      row.Item(0), row.Item(1), row.Item(2),
            '                                      validstr(row.Item(3)), mybrandid,
            '                                      validstr(row.Item(5)), mycomfam,
            '                                      rri, rri, modelcode, myrangeid))


            'Next
            myret = True
        Catch ex As Exception
            e.mymessage = ex.Message & " Recid:" & recid & " LastTable:" & lastTable
        End Try
        Return myret
    End Function
End Class


Class POComp
    Inherits BasePO
    Implements iPO

    Public Sub New(ByVal ds As DataSet, ByVal MyList As List(Of String()))
        MyBase.New(ds, MyList)
    End Sub

    Public Function BuildSb(ByVal sender As Object, ByVal e As POEventArgs) As Object Implements iPO.BuildSb
        Dim myret As Boolean = False
        Parent = sender
        Dim myrecord As String()
        Dim result As DataRow
        Dim rangeIdSeq As Long = DS.Tables(9).Rows(0).Item(1)
        Dim recid As Long
        Dim lastTable As String = String.Empty
        Try
            For i = 0 To MyList.Count - 1
                'find the record in existing table.
                ProgressReport(7, i + 1 & "," & MyList.Count)
                If i >= 1 Then
                    myrecord = MyList(i)
                    recid = i
                    'No Agreement because doesn't have CMMF

                    If myrecord(10) = "" Then 'not reversal tx
                        'Vendor Table
                        'Check Vendorcode
                        lastTable = "Vendor"
                        Dim pkey0(0) As Object
                        pkey0(0) = myrecord(20)
                        result = DS.Tables(0).Rows.Find(pkey0)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(20)
                            DS.Tables(0).Rows.Add(dr)
                            'vendorcode,vendorname,shortname2
                            VendorSB.Append(validlong(myrecord(20)) & vbTab &
                                            validstr(myrecord(21)) & vbTab &
                                           validstr(myrecord(22)) & vbCrLf)
                        Else
                            'Prepare Table for update
                            Dim pkey10(0) As Object
                            pkey10(0) = myrecord(20)
                            result = DS.Tables(10).Rows.Find(pkey10)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(10).NewRow
                                dr.Item(0) = myrecord(20)
                                dr.Item(1) = myrecord(21)
                                dr.Item(2) = myrecord(22)
                                DS.Tables(10).Rows.Add(dr)
                            Else
                                result.Item(1) = myrecord(21)
                                result.Item(2) = myrecord(22)
                            End If
                        End If
                    Else
                        'check poreversed
                        'Check Brand
                        lastTable = "POREVERSED"
                        Dim pkey8(3) As Object
                        pkey8(0) = DbAdapter1.dateformatdotdate(myrecord(1))
                        pkey8(1) = myrecord(3)
                        pkey8(2) = myrecord(7)
                        pkey8(3) = myrecord(8)
                        result = DS.Tables(8).Rows.Find(pkey8)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(8).NewRow
                            dr.Item(0) = DbAdapter1.dateformatdotdate(myrecord(1))
                            dr.Item(1) = myrecord(3)
                            dr.Item(2) = myrecord(7)
                            dr.Item(3) = myrecord(8)

                            DS.Tables(8).Rows.Add(dr)
                            'vendorcode,vendorname
                            Dim mydc = IIf(myrecord(17) = "H", -1, 1)
                            'miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reservedby,cmmf,plnt,amount,crcy,dc,
                            'qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir)
                            'Try
                            PoReversedSB.Append(dateformatdotyyyymmdd(myrecord(1)) & vbTab &
                                                validlong(myrecord(3)) & vbTab &
                                                validstr(myrecord(6)) & vbTab &
                                                "Null" & vbTab &
                                                validlong(myrecord(7)) & vbTab &
                                                validint(myrecord(8)) & vbTab &
                                                validlong(myrecord(10)) & vbTab &
                                                validlong(myrecord(12)) & vbTab &
                                                validint(myrecord(13)) & vbTab &
                                                validreal(myrecord(14)) & vbTab &
                                                validstr(myrecord(15)) & vbTab &
                                                mydc & vbTab &
                                                validreal(myrecord(18)) & vbTab &
                                                validstr(myrecord(19)) & vbTab &
                                                validlong(myrecord(20)) & vbTab &
                                                validstr(myrecord(23)) & vbTab &
                                                validstr(myrecord(24)) & vbTab &
                                                validstr(myrecord(25)) & vbTab &
                                                "Null" & vbTab &
                                                "Null" & vbTab &
                                                "Null" & vbTab &
                                                "Null" & vbTab &
                                                "Null" & vbTab &
                                                "Null" & vbCrLf)
                            
                            'Catch ex As Exception
                            '    Debug.Print("hello")
                            'End Try

                        End If
                    End If
                End If
            Next

            'get updatevendor
            For Each row As DataRow In DS.Tables(10).Rows
                If UpdateVendorSB.Length > 0 Then
                    UpdateVendorSB.Append(",")
                End If
                Dim shortname As String = "Null"
                If row.Item("shortname2") <> "" Then
                    shortname = row.Item("shortname2")
                End If
                UpdateVendorSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying]",
                                   row.Item(0), validstr(row.Item(1)), validstr(shortname)))
            Next
            myret = True
        Catch ex As Exception
            e.mymessage = ex.Message & " Recid:" & recid & " LastTable:" & lastTable
        End Try
        Return myret
    End Function

    Public Function BuildSbold(ByVal sender As Object, ByVal e As POEventArgs) As Object 'Implements iPO.BuildSb
        Dim myret As Boolean = False
        Parent = sender
        Dim myrecord As String()
        Dim result As DataRow
        Dim rangeIdSeq As Long = DS.Tables(9).Rows(0).Item(1)
        Dim recid As Long
        Dim lastTable As String = String.Empty
        Try

            For i = 0 To MyList.Count - 1
                'find the record in existing table.
                ProgressReport(7, i + 1 & "," & MyList.Count)
                If i >= 1 Then
                    myrecord = MyList(i)
                    recid = i

                    If myrecord(10) = "" Then 'not reversal tx
                        'Vendor Table
                        'Check Vendorcode
                        lastTable = "Vendor"
                        Dim pkey0(0) As Object
                        pkey0(0) = myrecord(20)
                        result = DS.Tables(0).Rows.Find(pkey0)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(0).NewRow
                            dr.Item(0) = myrecord(20)
                            DS.Tables(0).Rows.Add(dr)
                            'vendorcode,vendorname
                            VendorSB.Append(validlong(myrecord(20)) & vbTab &
                                            validstr(myrecord(21)) & vbTab &
                                           validstr(myrecord(22)) & vbCrLf)
                        Else
                            If UpdateVendorSB.Length > 0 Then
                                UpdateVendorSB.Append(",")
                            End If
                            Dim shortname As String = "Null"
                            If myrecord(22) <> "" Then
                                shortname = myrecord(22)
                            End If
                            UpdateVendorSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying]",
                                               myrecord(20), validstr(myrecord(21)), validstr(shortname)))

                        End If
                    Else
                        'check poreversed
                        'Check Brand
                        lastTable = "POREVERSED"
                        Dim pkey8(3) As Object
                        pkey8(0) = DbAdapter1.dateformatdotdate(myrecord(1))
                        pkey8(1) = myrecord(3)
                        pkey8(2) = myrecord(7)
                        pkey8(3) = myrecord(8)
                        result = DS.Tables(8).Rows.Find(pkey8)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(8).NewRow
                            dr.Item(0) = DbAdapter1.dateformatdotdate(myrecord(1))
                            dr.Item(1) = myrecord(3)
                            dr.Item(2) = myrecord(7)
                            dr.Item(3) = myrecord(8)

                            DS.Tables(8).Rows.Add(dr)
                            'vendorcode,vendorname
                            Dim mydc = IIf(myrecord(17) = "H", -1, 1)
                            'miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reservedby,cmmf,plnt,amount,crcy,dc,
                            'qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir)
                            PoReversedSB.Append(dateformatdotyyyymmdd(myrecord(1)) & vbTab &
                                                validlong(myrecord(3)) & vbTab &
                                                validstr(myrecord(6)) & vbTab &
                                                validlong(myrecord(7)) & vbTab &
                                                validint(myrecord(8)) & vbTab &
                                                validlong(myrecord(12)) & vbTab &
                                                validint(myrecord(13)) & vbTab &
                                                validreal(myrecord(14)) & vbTab &
                                                validstr(myrecord(15)) & vbTab &
                                                validint(mydc) & vbTab &
                                                validreal(myrecord(18)) & vbTab &
                                                validstr(myrecord(19)) & vbTab &
                                                validlong(myrecord(20)) & vbTab &
                                                validstr(myrecord(23)) & vbTab &
                                                validstr(myrecord(24)) & vbTab &
                                                validlong(myrecord(25)) & vbCrLf)
                        End If
                    End If
                End If



            Next
            myret = True
        Catch ex As Exception
            e.mymessage = ex.Message & " Recid:" & recid & " LastTable:" & lastTable
        End Try


        Return myret
    End Function
End Class


MustInherit Class BasePO
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Public Property DS As DataSet
    Public Property MyList As List(Of String())
    Public Property VendorSB As New StringBuilder
    Public Property UpdateVendorSB As New StringBuilder
    Public Property BrandSB As New StringBuilder
    Public Property FamilySB As New StringBuilder
    Public Property LoadingSB As New StringBuilder
    Public Property CMMFSB As New StringBuilder
    Public Property RangeSB As New StringBuilder
    Public Property UpdateCMMFSB As New StringBuilder
    Public Property AgreementSB As New StringBuilder
    Public Property AgValueSB As New StringBuilder
    Public Property UpdatePOHDSB As New StringBuilder
    Public Property PoReversedSB As New StringBuilder
    Public Property SalesDocSB As New StringBuilder
    Public Property UpdateSalesDocSB As New StringBuilder
    Public Property RelSalesDocPoSB As New StringBuilder
    Public Parent As Object

    Public Sub New(ByVal DS As DataSet, ByVal MyList As List(Of String()))
        Me.DS = DS
        Me.MyList = MyList
    End Sub

    Public Overridable Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Parent.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Parent.ToolStripStatusLabel1.Text = message
                Case 2
                    Parent.ToolStripStatusLabel2.Text = message
                Case 4
                    'Me.Label4.Text = message
                Case 5
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7
                    Dim myvalue = message.ToString.Split(",")
                    Parent.ToolStripProgressBar1.Minimum = 1
                    Parent.ToolStripProgressBar1.Value = myvalue(0)
                    Parent.ToolStripProgressBar1.Maximum = myvalue(1)
            End Select

        End If
    End Sub

End Class

Public Class POEventArgs
    Inherits EventArgs
    Public mymessage As String

    Public Sub New()
        MyBase.New()
    End Sub

End Class