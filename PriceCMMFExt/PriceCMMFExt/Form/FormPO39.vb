Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormPO39

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
    Public Property MyList As List(Of String())
    Public Property VendorSB As New StringBuilder
    Public Property BrandSB As New StringBuilder
    Public Property FamilySB As New StringBuilder
    Public Property LoadingSB As New StringBuilder
    Public Property CMMFSB As New StringBuilder
    Public Property RangeSB As New StringBuilder
    Public Property AgreementSB As New StringBuilder
    Public Property AgValueSB As New StringBuilder
    Public Property PoReversedSB As New StringBuilder
    Public Property PoPlantSB As New StringBuilder

    Public Property UpdateCMMFSB As New StringBuilder
    
    Public Property UpdatePOHDSB As New StringBuilder

    Public Property UpdateVendorSB As New StringBuilder

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
                         "select po from poplant;"
                               

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

                DS.Tables(9).TableName = "POPlant"
                Dim idx9(0) As DataColumn
                idx9(0) = DS.Tables(9).Columns(0)
                DS.Tables(9).PrimaryKey = idx9

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

                Dim message As String = String.Empty

                If Not BuildSb(mylist, message) Then
                    ProgressReport(1, message)
                    Exit Sub
                End If
                


            End With
        End Using
        'Update Record        
        Try
            ProgressReport(6, "Marque")
            If VendorSB.Length > 0 Then
                ProgressReport(2, "Copy Vendor")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy vendor(vendorcode,vendorname,shortname2) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, VendorSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Vendor" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If UpdateVendorSB.Length > 0 Then
                ProgressReport(2, "Update Vendor")
                'vendorname,shortname2
                sqlstr = "update vendor set vendorname= foo.vendorname,shortname2 = foo.shortname2 from (select * from array_to_set3(Array[" & UpdateVendorSB.ToString &
                         "]) as tb (id character varying,vendorname character varying,shortname2 character varying))foo where vendorcode = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errormsg) Then
                    ProgressReport(2, "Update Vendor" & "::" & errormsg)
                    Exit Sub
                End If
            End If
            If BrandSB.Length > 0 Then
                ProgressReport(2, "Copy Brand")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy brand(brandid,brandname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, BrandSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Brand" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If FamilySB.Length > 0 Then
                ProgressReport(2, "Copy Family")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy family(familyid,familyname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, FamilySB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Family" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If LoadingSB.Length > 0 Then
                ProgressReport(2, "Copy Loading")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy loading(loadingcode,loadingname) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, LoadingSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Loading" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If RangeSB.Length > 0 Then
                ProgressReport(2, "Copy Range")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy range(range,rangedesc) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, RangeSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Range" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If CMMFSB.Length > 0 Then
                ProgressReport(2, "Copy CMMF")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy cmmf(cmmf,plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, CMMFSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy CMMF" & "::" & errmessage)
                    Exit Sub
                End If
            End If
            If UpdateCMMFSB.Length > 0 Then
                ProgressReport(2, "Update CMMF")
                'plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                sqlstr = "update cmmf set plnt= foo.plnt::integer,brandid=foo.brandid::integer,cmmftype=foo.cmmftype,comfam=foo.comfam::integer,rir=foo.rir::character(2),activitycode=foo.activitycode::character(2),modelcode=foo.modelcode,rangeid=foo.rangeid::bigint from (select * from array_to_set9(Array[" & UpdateCMMFSB.ToString &
                         "]) as tb (id character varying,plnt character varying,brandid character varying,cmmftype character varying,comfam character varying,rir character varying,activitycode character varying,modelcode character varying,rangeid character varying))foo where cmmf = foo.id::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errormsg) Then
                    ProgressReport(2, "Update CMMF" & "::" & errormsg)
                    Exit Sub
                End If
            End If

            If AgreementSB.Length > 0 Then
                ProgressReport(2, "Copy Agreement")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy agreementtx(agreement,material,postingdate) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, AgreementSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy Agreement" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If AgValueSB.Length > 0 Then
                ProgressReport(2, "Copy AgreementValue")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy agvalue(agreement) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, AgValueSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy AgreementValue" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If PoPlantSB.Length > 0 Then
                ProgressReport(2, "Copy POPlant")
                'mironumber bigint ,miropostingdate date, supplierinvoicenum character varying, vendorcode bigint
                sqlstr = "copy poplant(po,plant) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PoPlantSB.ToString, myret)
                If Not myret Then
                    ProgressReport(2, "Copy POPlant" & "::" & errmessage)
                    Exit Sub
                End If
            End If

            If PoReversedSB.Length > 0 Then
                ProgressReport(2, "Copy POReversed")
                sqlstr = "copy poreversed(miropostingdate,mironumber,supplierinvoicenum,pohd,polineno,reversedby,cmmf,plnt,amount,crcy,dc,qty,oun,vendorcode,payt,purchasinggroup,agreement,brandid,cmmftype,comfam,rir) from stdin with null as 'Null';"
                Dim errmessage As String = String.Empty
                Dim myret As Boolean = False
                errmessage = DbAdapter1.copy(sqlstr, PoReversedSB.ToString, myret)
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

    Private Function BuildSb(ByVal myList As List(Of String()), ByRef message As String) As Boolean
        Dim myret As Boolean = False

        'Check vendor
        'Check Brand
        'Check Family
        'Check Loading Code
        'Check CMMF
        'Check Range
        'Check Agreement
        'Check Agreement Value
        'Check PO Reversed
        Dim myrecord As String()
        Dim lastTable As String
        Dim result As Object
        AgreementSB.Clear()
        VendorSB.Clear()
        BrandSB.Clear()
        LoadingSB.Clear()
        FamilySB.Clear()
        RangeSB.Clear()
        PoReversedSB.Clear()
        CMMFSB.Clear()
        AgValueSB.Clear()
        PoPlantSB.Clear()
        Try
            For i = 0 To myList.Count - 1
                'find the record in existing table.
                ProgressReport(7, i + 1 & "," & myList.Count)
                If i >= 0 Then
                    myrecord = myList(i)
                    'Checking Agreement Table
                    If myrecord(23) <> "" Then
                        'agreement
                        lastTable = "Agreement"
                        Dim pkey6(2) As Object
                        pkey6(0) = myrecord(23) 'Agreement
                        pkey6(1) = myrecord(11) 'CMMF
                        pkey6(2) = DbAdapter1.dateformatdotdate(myrecord(1))

                        result = DS.Tables(6).Rows.Find(pkey6)
                        If IsNothing(result) Then
                            Dim dr As DataRow = DS.Tables(6).NewRow
                            dr.Item(0) = myrecord(23)
                            dr.Item(1) = myrecord(11)
                            dr.Item(2) = pkey6(2)
                            DS.Tables(6).Rows.Add(dr)
                            'agreement,material,postingdate
                            AgreementSB.Append(validlong(myrecord(23)) & vbTab &
                                            validlong(myrecord(11)) & vbTab & dateformatdotyyyymmdd(myrecord(1)) & vbCrLf)
                        End If
                    End If

                    If myrecord(10) = "" Then 'not reversal tx
                        'check POPlant


                        If myrecord(7) <> "" Then

                            lastTable = "POPlant"
                            Dim pkey1(0) As Object
                            pkey1(0) = myrecord(7)
                            result = DS.Tables(9).Rows.Find(pkey1)
                            If IsNothing(result) Then
                                Dim dr As DataRow = DS.Tables(9).NewRow
                                dr.Item(0) = myrecord(7)
                                DS.Tables(9).Rows.Add(dr)
                                'vendorcode,vendorname
                                PoPlantSB.Append(validlong(myrecord(7)) & vbTab &
                                               validint(myrecord(12)) & vbCrLf)
                            End If
                        End If

                        'Vendor Table
                        'Check Vendorcode
                        'lastTable = "Vendor"
                        'Dim pkey0(0) As Object
                        'pkey0(0) = myrecord(18)
                        'result = DS.Tables(0).Rows.Find(pkey0)
                        'If IsNothing(result) Then
                        '    Dim dr As DataRow = DS.Tables(0).NewRow
                        '    dr.Item(0) = myrecord(18)
                        '    DS.Tables(0).Rows.Add(dr)
                        '    'vendorcode,vendorname,shortname2
                        '    VendorSB.Append(validlong(myrecord(18)) & vbTab &
                        '                    validstr(myrecord(19)) & vbTab &
                        '                   validstr(myrecord(20)) & vbCrLf)
                        'Else
                        '    Dim pkey10(0) As Object
                        '    pkey10(0) = myrecord(18)
                        '    result = DS.Tables(10).Rows.Find(pkey10)
                        '    If IsNothing(result) Then
                        '        Dim dr As DataRow = DS.Tables(10).NewRow
                        '        dr.Item(0) = myrecord(18)
                        '        dr.Item(1) = myrecord(19)
                        '        dr.Item(2) = myrecord(20)
                        '        DS.Tables(10).Rows.Add(dr)
                        '    Else
                        '        result.Item(1) = myrecord(19)
                        '        result.Item(2) = myrecord(20)
                        '    End If
                        'End If

                        ''Check Brand
                        'lastTable = "Brand"
                        'Dim brandid = "Null"
                        'If myrecord(31) <> "" Then
                        '    brandid = myrecord(31)
                        '    Dim pkey1(0) As Object
                        '    pkey1(0) = myrecord(31)
                        '    result = DS.Tables(1).Rows.Find(pkey1)
                        '    If IsNothing(result) Then
                        '        Dim dr As DataRow = DS.Tables(1).NewRow
                        '        dr.Item(0) = myrecord(31)
                        '        DS.Tables(1).Rows.Add(dr)
                        '        'vendorcode,vendorname
                        '        BrandSB.Append(validint(myrecord(31)) & vbTab &
                        '                       validstr(myrecord(31)) & vbCrLf)
                        '    End If
                        'End If

                        ''Check Family
                        'If myrecord(33) <> "" Then

                        '    lastTable = "Family"
                        '    Dim pkey2(0) As Object
                        '    pkey2(0) = myrecord(33)
                        '    result = DS.Tables(2).Rows.Find(pkey2)
                        '    If IsNothing(result) Then
                        '        Dim dr As DataRow = DS.Tables(2).NewRow
                        '        dr.Item(0) = myrecord(33)
                        '        DS.Tables(2).Rows.Add(dr)
                        '        'vendorcode,vendorname
                        '        BrandSB.Append(validint(myrecord(33)) & vbTab &
                        '                       validstr(myrecord(33)) & vbCrLf)
                        '    End If
                        'End If
                        ''Check Loading
                        'lastTable = "Loading"
                        'Dim pkey3(0) As Object
                        'pkey3(0) = myrecord(28)
                        'result = DS.Tables(3).Rows.Find(pkey3)
                        'If IsNothing(result) Then
                        '    Dim dr As DataRow = DS.Tables(3).NewRow
                        '    dr.Item(0) = myrecord(28)
                        '    DS.Tables(3).Rows.Add(dr)
                        '    '
                        '    LoadingSB.Append(validint(myrecord(28)) & vbTab &
                        '                    validstr(myrecord(28)) & vbCrLf)
                        'End If
                        'Dim myrangeid As String = ""

                        ''If myrecord(41) <> "" Then
                        ''    'Check Range
                        ''    lastTable = "Range"
                        ''    Dim pkey5(0) As Object
                        ''    pkey5(0) = myrecord(41)
                        ''    result = DS.Tables(5).Rows.Find(pkey5)
                        ''    If IsNothing(result) Then
                        ''        rangeIdSeq += 1
                        ''        Dim dr As DataRow = DS.Tables(5).NewRow
                        ''        dr.Item(0) = myrecord(41)
                        ''        dr.Item(1) = rangeIdSeq
                        ''        myrangeid = rangeIdSeq
                        ''        DS.Tables(5).Rows.Add(dr)
                        ''        myrangeid = rangeIdSeq
                        ''        'range
                        ''        RangeSB.Append(validstr(myrecord(41)) & vbTab &
                        ''                       validstr(myrecord(41)) & vbCrLf)
                        ''    Else
                        ''        myrangeid = result.Item(1)
                        ''    End If
                        ''End If

                        'If myrecord(13) <> "" Then
                        '    'Check CMMF
                        '    lastTable = "CMMF"
                        '    'Dim pkey4(0) As Object
                        '    'pkey4(0) = myrecord(13)
                        '    'result = DS.Tables(4).Rows.Find(pkey4)

                        '    'If IsNothing(result) Then
                        '    '    Dim dr As DataRow = DS.Tables(4).NewRow
                        '    '    dr.Item(0) = myrecord(13)
                        '    '    DS.Tables(4).Rows.Add(dr)
                        '    '    'cmmf,plnt,vendorcode,loadingcode,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                        '    '    CMMFSB.Append(validlong(myrecord(13)) & vbTab &
                        '    '                     validint(myrecord(14)) & vbTab &
                        '    '                     validlong(myrecord(21)) & vbTab &
                        '    '                     validstr(myrecord(28)) & vbTab &
                        '    '                     validint(myrecord(31)) & vbTab &
                        '    '                     validstr(myrecord(32)) & vbTab &
                        '    '                     validint(myrecord(33)) & vbTab &
                        '    '                     validstr(myrecord(34)) & vbTab &
                        '    '                     validstr(myrecord(34)) & vbTab &
                        '    '                     validstr(myrecord(40)) & vbTab &
                        '    '                     myrangeid & vbCrLf)
                        '    'Else
                        '    'Check for update
                        '    'plnt character varying,vendorcode character varying,loadingcode character varying,brandid character varying,cmmftype character varying,comfam character varying,rir character varying,activitycode character varying,modelcode character varying,rangeid character varying
                        '    Dim pkey11(0) As Object
                        '    pkey11(0) = myrecord(13)
                        '    result = DS.Tables(11).Rows.Find(pkey11)
                        '    If IsNothing(result) Then
                        '        Dim dr As DataRow = DS.Tables(11).NewRow
                        '        dr.Item(0) = myrecord(13)
                        '        dr.Item(1) = myrecord(14)
                        '        dr.Item(2) = myrecord(21)
                        '        dr.Item(3) = myrecord(28)
                        '        dr.Item(4) = DbAdapter1.validint(myrecord(31))
                        '        dr.Item(5) = myrecord(32)
                        '        dr.Item(6) = DbAdapter1.validint(myrecord(33))
                        '        dr.Item(7) = DbAdapter1.validchar(myrecord(34))
                        '        dr.Item(8) = DbAdapter1.validchar(myrecord(34))
                        '        dr.Item(9) = DbAdapter1.validchar(myrecord(40))
                        '        dr.Item(10) = DbAdapter1.validint(myrangeid)
                        '        DS.Tables(11).Rows.Add(dr)
                        '        CMMFSB.Append(validlong(myrecord(13)) & vbTab &
                        '                         validint(myrecord(14)) & vbTab &
                        '                         validlong(myrecord(21)) & vbTab &
                        '                         validstr(myrecord(28)) & vbTab &
                        '                         validint(myrecord(31)) & vbTab &
                        '                         validstr(myrecord(32)) & vbTab &
                        '                         validint(myrecord(33)) & vbTab &
                        '                         validstr(myrecord(34)) & vbTab &
                        '                         validstr(myrecord(34)) & vbTab &
                        '                         validstr(myrecord(40)) & vbTab &
                        '                         myrangeid & vbCrLf)
                        '    Else
                        '        Dim flag As Boolean = False
                        '        If IsDBNull(result.Item(1)) Or IsDBNull(result.Item(4)) Or _
                        '            IsDBNull(result.Item(5)) Or IsDBNull(result.Item(6)) Or IsDBNull(result.Item(7)) _
                        '            Or IsDBNull(result.Item(8)) Or IsDBNull(result.Item(9)) Or IsDBNull(result.Item(10)) Then
                        '            flag = True
                        '        ElseIf Not (result.Item(1) = myrecord(14) AndAlso result.Item(1) = myrecord(14) AndAlso
                        '            result.Item(4) = DbAdapter1.validint(myrecord(31)) AndAlso
                        '            result.Item(5) = myrecord(32) AndAlso
                        '            result.Item(6) = DbAdapter1.validint(myrecord(33)) AndAlso
                        '            result.Item(7) = myrecord(34) AndAlso
                        '            result.Item(8) = myrecord(34) AndAlso
                        '            result.Item(9) = myrecord(40) AndAlso
                        '            result.Item(10) = DbAdapter1.validint(myrangeid)) Then
                        '            flag = True

                        '        End If
                        '        If flag Then
                        '            Dim mybrandid As String = "Null"
                        '            If Not myrecord(31) = "" Then
                        '                mybrandid = myrecord(31)
                        '            End If
                        '            Dim mycomfam As String = "Null"
                        '            If Not myrecord(33) = "" Then
                        '                mycomfam = myrecord(33)
                        '            End If
                        '            Dim strrangeid As String = "Null"
                        '            If Not myrangeid = "" Then
                        '                strrangeid = myrangeid
                        '            End If
                        '            Dim rri As String = "Null"
                        '            If Not myrecord(34) = "" Then
                        '                rri = "'" & myrecord(34) & "'"
                        '            End If
                        '            Dim modelcode = "Null"
                        '            If Not myrecord(40) = "" Then
                        '                modelcode = "'" & myrecord(40) & "'"
                        '            End If
                        '            If UpdateCMMFSB.Length > 0 Then
                        '                UpdateCMMFSB.Append(",")
                        '            End If
                        '            'cmmf,plnt,brandid,cmmftype,comfam,rir,activitycode,modelcode,rangeid
                        '            UpdateCMMFSB.Append(String.Format("[{0}::character varying,{1}::character varying,{2}::character varying,'{3}'::character varying,{4}::character varying,{5}::character varying,{6}::character varying,{7}::character varying,{8}::character varying]",
                        '                           myrecord(13), myrecord(14), mybrandid,
                        '                          validstr(myrecord(32)), mycomfam,
                        '                          rri, rri, modelcode, strrangeid))

                        '            result.Item(1) = myrecord(14)
                        '            'result.Item(2) = myrecord(21)
                        '            'result.Item(3) = myrecord(28)
                        '            result.Item(4) = DbAdapter1.validint(myrecord(31))
                        '            result.Item(5) = myrecord(32)
                        '            result.Item(6) = DbAdapter1.validint(myrecord(33))
                        '            result.Item(7) = myrecord(34)
                        '            result.Item(8) = myrecord(34)
                        '            result.Item(9) = myrecord(40)
                        '            result.Item(10) = DbAdapter1.validint(myrangeid)
                        '        Else
                        '            'Debug.Print("no update")
                        '        End If


                        '    End If
                        '    'End If
                        'End If


                        ''ShipToParty
                        'If myrecord(29) <> "" Then
                        '    'Check PayT
                        '    lastTable = "ShipToParty"

                        '    Dim pkey13(0) As Object
                        '    pkey13(0) = myrecord(29)
                        '    result = DS.Tables(13).Rows.Find(pkey13)
                        '    If IsNothing(result) Then
                        '        Dim dr As DataRow = DS.Tables(13).NewRow
                        '        dr.Item(0) = myrecord(29)
                        '        dr.Item(1) = myrecord(38)
                        '        DS.Tables(13).Rows.Add(dr)
                        '        SalesDocSB.Append(myrecord(29) & vbTab &
                        '                               myrecord(7) & vbTab &
                        '                               dateformatdotyyyymmdd(myrecord(42)) & vbTab &
                        '                               myrecord(35) & vbTab &
                        '                                myrecord(38) & vbCrLf)
                        '    Else
                        '        Dim flag As Boolean = False
                        '        If IsDBNull(result.Item(1)) Then
                        '            flag = True
                        '        ElseIf result.Item(1) <> myrecord(38) Then
                        '            flag = True

                        '        End If
                        '        If flag Then
                        '            If UpdateSalesDocSB.Length > 0 Then
                        '                UpdateSalesDocSB.Append(",")
                        '            End If

                        '            UpdateSalesDocSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]",
                        '                   myrecord(29), validlong(myrecord(38))))

                        '            result.Item(1) = myrecord(38)
                        '        End If

                        '    End If

                        '    lastTable = "RelSalesDocPo"
                        '    Dim pkey14(3) As Object
                        '    pkey14(0) = myrecord(29)
                        '    pkey14(1) = myrecord(30)
                        '    pkey14(2) = myrecord(8)
                        '    pkey14(3) = myrecord(9)
                        '    result = DS.Tables(14).Rows.Find(pkey14)
                        '    If IsNothing(result) Then
                        '        Dim dr As DataRow = DS.Tables(14).NewRow
                        '        dr.Item(0) = myrecord(29)
                        '        dr.Item(1) = myrecord(30)
                        '        dr.Item(2) = myrecord(8)
                        '        dr.Item(3) = myrecord(9)
                        '        DS.Tables(14).Rows.Add(dr)
                        '        RelSalesDocPoSB.Append(myrecord(29) & vbTab &
                        '                               myrecord(30) & vbTab &
                        '                               myrecord(8) & vbTab &
                        '                                myrecord(9) & vbCrLf)

                        '    End If

                        'End If

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
                            Dim mydc = IIf(myrecord(15) = "H", -1, 1)
                            '                                     miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reservedby,cmmf,plnt,amount,crcy,dc,
                            '                                     qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir)
                            'Try
                            'Sqlstr = "insert into poreversed(
                            'miropostingdate,mironumber,supplierinvoicenum,pono,pohd,polineno,reversedby,cmmf,plnt,amount,crcy,dc,
                            'qty,oun,vendorcode,payt,purchasinggroup,agreement,salesdoc,salesdocno,brandid,cmmftype,comfam,rir) Values(" & _
                            ' DateFormatDot(myrecord(1)) & "," & validNum(myrecord(3)) & "," & escapeString(myrecord(6)) & "," & escapeString(myrecord(7)) & "," & validNum(myrecord(8)) & "," & validNum(myrecord(9)) & "," & validNum(myrecord(11)) & "," & validNum(myrecord(12)) & "," & validNum(myrecord(13)) & "," & validNum(Replace(myrecord(14), ",", ".")) & "," & escapeString(myrecord(15)) & "," & mydc & "," & validNum(Replace(myrecord(17), ",", ".")) & "," & escapeString(myrecord(18)) & "," & validNum(myrecord(19)) & "," & escapeString(myrecord(22)) & "," & escapeString(myrecord(23)) & "," & validNum(myrecord(24)) & "," & validNum(myrecord(27)) & "," & validNum(myrecord(28)) & "," & validNum(myrecord(29)) & "," & escapeString(myrecord(30)) & "," & validNum(myrecord(31)) & "," & escapeString(myrecord(32)) & ")"

                            PoReversedSB.Append(dateformatdotyyyymmdd(myrecord(1)) & vbTab &
                                                validlong(myrecord(3)) & vbTab &
                                                validstr(myrecord(6)) & vbTab &
                                                validlong(myrecord(7)) & vbTab &
                                                validint(myrecord(8)) & vbTab &
                                                validlong(myrecord(10)) & vbTab &
                                                validlong(myrecord(11)) & vbTab &
                                                validint(myrecord(12)) & vbTab &
                                                validreal(myrecord(13)) & vbTab &
                                                validstr(myrecord(14)) & vbTab &
                                                mydc & vbTab &
                                                validreal(myrecord(16)) & vbTab &
                                                validstr(myrecord(17)) & vbTab &
                                                validlong(myrecord(18)) & vbTab &
                                                validstr(myrecord(21)) & vbTab &
                                                validstr(myrecord(22)) & vbTab &
                                                validlong(myrecord(23)) & vbTab &
                                                validint(myrecord(26)) & vbTab &
                                                validstr(myrecord(27)) & vbTab &
                                                validint(myrecord(28)) & vbTab &
                                                validstr(myrecord(29)) & vbCrLf)
                            'Catch ex As Exception
                            '    Debug.Print("hello")
                            'End Try

                        End If
                    End If
                End If
            Next            
        Catch ex As Exception
            message = ex.Message
            Return myret
        End Try
        myret = True
        Return myret

    End Function



End Class