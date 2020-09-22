Imports PriceCMMFExt.PublicClass
Imports System.IO
Imports System.Text

Public Class ExportSAPClass
    Dim DS As DataSet
    Public mymessage As String
    Dim HDBS As New BindingSource
    Dim seq As Integer
    Public Sub New()
        If IsNothing(DbAdapter1) Then
            DbAdapter1 = New DbAdapter
        End If
    End Sub
    Private Function getdataset() As Boolean
        Dim myret As Boolean = False
        'Dim sqlstr = "select * from pricechangedtl pcdt" &
        '             " left join pricechangehd pchd on pchd.pricechangehdid = pcdt.pricechangehdid" &
        '             " where exportfileid isnull and status = 5 and pricetype = 'FOB'" &
        '             " order by plant,vendorcode,cmmf,validon;" &
        '             " select cvalue from paramhd where paramname = 'ExportToSAPPriceChange';" &
        '             " select nextval('exportfileid_seq') as sequence;"
        'Dim sqlstr = "with oum as (select distinct * from (select cmmf,vendorcode,first_value(uom) OVER (partition by cmmf,vendorcode ORDER BY validfrom DESC) as uom" &
        '             " from pricelist order by cmmf,vendorcode,validfrom desc) as foo)" &
        '             " select pcdt.*,pchd.*,doc.getvendorcurr(pcdt.vendorcode,pcdt.validon) as curr ,oum.uom from pricechangedtl pcdt" &
        '             " left join pricechangehd pchd on pchd.pricechangehdid = pcdt.pricechangehdid" &
        '             " left join oum on oum.vendorcode = pcdt.vendorcode and oum.cmmf = pcdt.cmmf" &
        '             " where   exportfileid isnull   and   status = 5 and   pricetype = 'FOB' order by plant,pcdt.vendorcode,pcdt.cmmf,validon;" &
        '             " select cvalue from paramhd where paramname = 'ExportToSAPPriceChange';" &
        '             " select nextval('exportfileid_seq') as sequence;"
        Dim sqlstr = "with oum as (select distinct * from (select cmmf,vendorcode,first_value(uom) OVER (partition by cmmf,vendorcode ORDER BY validfrom DESC) as uom" &
                     " from pricelist order by cmmf,vendorcode,validfrom desc) as foo)," &
                     " dtl as (select pcdt.pricechangedtlid,pcdt.pricechangehdid,pcdt.vendorcode,pcdt.cmmf,pcdt.purchorg,plant,pcdt.validon,pcdt.price,pcdt.pricingunit,pcdt.comment,pcdt.sap," &
                     " pchd.pricechangehdid,pchd.creator,pchd.validator1,pchd.validator2,pchd.pricetype,pchd.description,pchd.submitdate,pchd.negotiateddate,pchd.attachment,pchd.creationdate,pchd.status,pchd.actiondate,       pchd.actionby,pchd.sendstdvalidatedtocreator,pchd.sendtocc,pchd.reasonid,pchd.sendcompletedtocreator,pchd.exportfileid,pchd.exportfiledate,pchd.validator3,pchd.sendtowmf,pchd.specialprojectid," &
                     " doc.getvendorcurr(pcdt.vendorcode,pcdt.validon) as curr ,oum.uom from pricechangedtl pcdt" &
                     " left join pricechangehd pchd on pchd.pricechangehdid = pcdt.pricechangehdid" &
                     " left join oum on oum.vendorcode = pcdt.vendorcode and oum.cmmf = pcdt.cmmf" &
                     " where   exportfileid isnull   and   status = 5 and   pricetype = 'FOB' order by plant,pcdt.vendorcode,pcdt.cmmf,validon)" &
                     " select * from dtl" &
                     " union all " &
                     " (select pcdt.pricechangedtlid,pcdt.pricechangehdid,pcdt.vendorcode,pcdt.cmmf,pcdt.purchorg,3750,pcdt.validon,pcdt.price,pcdt.pricingunit,pcdt.comment,pcdt.sap," &
                     " pchd.pricechangehdid, pchd.creator, pchd.validator1, pchd.validator2, pchd.pricetype, pchd.description, pchd.submitdate, pchd.negotiateddate, pchd.attachment, pchd.creationdate, pchd.status, pchd.actiondate," &
                     " pchd.actionby, pchd.sendstdvalidatedtocreator, pchd.sendtocc, pchd.reasonid, pchd.sendcompletedtocreator, pchd.exportfileid, pchd.exportfiledate, pchd.validator3, pchd.sendtowmf, pchd.specialprojectid" &
                     " ,doc.getvendorcurr(pcdt.vendorcode,pcdt.validon) as curr ,oum.uom from pricechangedtl pcdt " &
                     " inner join doc.cmmf3750 c on c.cmmf = pcdt.cmmf" &
                     " left join pricechangehd pchd on pchd.pricechangehdid = pcdt.pricechangehdid" &
                     " left join oum on oum.vendorcode = pcdt.vendorcode and oum.cmmf = pcdt.cmmf" &
                     " where exportfileid isnull and status = 5 and pricetype = 'FOB' and pcdt.plant <> 3750 order by plant,pcdt.vendorcode,pcdt.cmmf,validon);" &
                     " select cvalue from paramhd where paramname = 'ExportToSAPPriceChange';" &
                     " select nextval('exportfileid_seq') as sequence;"
        DS = New DataSet
        Try
            If DbAdapter1.TbgetDataSet(sqlstr, DS, mymessage) Then
                DS.Tables(0).TableName = "PriceChangeDTL"
                DS.Tables(1).TableName = "Path"
                DS.Tables(2).TableName = "Sequence"
                seq = DS.Tables(2).Rows(0).Item("sequence")
                HDBS.DataSource = DS.Tables(0)
                myret = True
            End If
        Catch ex As Exception
            mymessage = ex.Message
        End Try


        Return myret
    End Function
    Public Function ExportFile() As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder

        If getdataset() Then
            If DS.Tables(0).Rows.Count > 0 Then
                'sb.Append("[Variant ID]" & vbTab & "[Variant Text]" & vbTab & "LIFNR" & vbTab & "MATNR" & vbTab & "PURC_ORG" & vbTab & "WERKS" & vbTab & "DATE" & vbTab & "PRICE" & vbTab & "UNIT" & vbCrLf)
                'sb.Append("-->" & vbTab & "Parameter texts" & vbTab & "Vendor" & vbTab & "Material" & vbTab & "Purch. Organization" & vbTab & "Plant" & vbTab & "Valid on" & vbTab & "Rate" & vbTab & "Pricing unit" & vbCrLf)
                'sb.Append("-->" & vbTab & "Default Values" & vbTab & "#99003700" & vbTab & "#1500631606" & vbTab & "#6101" & vbTab & "#6110" & vbTab & "02.01.2006" & vbTab & "#           10000" & vbTab & "# 1000" & vbCrLf)
                'sb.Append("*** Changes to the default values displayed above not effective" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbCrLf)
                sb.Append("Purchasing Organization" & vbTab & "Plant" & vbTab & "Info Record Type" & vbTab & "Vendor" & vbTab & "Vendor name" & vbTab & "Material" & vbTab & "Material Description" & vbTab & "Beginning Date" & vbTab & "End date" & vbTab & "Condition Type" & vbTab & "Scale Quantity" & vbTab & "Unit Quantity" & vbTab & "Amount" & vbTab & "Currency" & vbTab & "Pricing Unit" & vbTab & "Unit Quantity" & vbTab & "Purch.Info Record" & vbTab & "Scale" & vbTab & "Type of metal" & vbTab & "Vendor Material" & vbCrLf)
                For Each dr As DataRow In DS.Tables(0).Rows
                    sb.Append(dr.Item("purchorg") & vbTab & dr.Item("plant") & vbTab & vbTab & dr.Item("vendorcode") & vbTab & vbTab & dr.Item("cmmf") & vbTab & vbTab & sapdate(dr.Item("validon")) & vbTab & sapdate(CDate("9999-12-31")) & vbTab & "PB00" & vbTab & vbTab & vbTab & dr.Item("price") & vbTab & dr.Item("curr") & vbTab & dr.Item("pricingunit") & vbTab & dr.Item("uom") & vbTab & vbTab & vbTab & vbTab & vbTab & vbCrLf)

                    'assign flag
                    dr.Item("exportfileid") = seq
                    dr.Item("exportfiledate") = Today.Date
                Next
                Dim fullpathname As String
                fullpathname = DS.Tables(1).Rows(0).Item("cvalue") & "\" & "ZFA2C039_GSF_MAJ_FIA" & "_" & Today.Year & Format(Today.Month, "00") & Format(seq, "000") & ".TXT"
                Using mystream As New StreamWriter(fullpathname, False, System.Text.Encoding.ASCII)
                    mystream.WriteLine(sb.ToString)
                End Using

                myret = True

                'Flag Send to SAP
                'Dim q = From n In HDBS.List
                '        Group n By key = n.item("pricechangehdid") Into Group
                '        Select key, Data = Group

                'For Each n In q
                '    CType(n.Data, DataRow).Item("exportfileid") = seq
                '    CType(n.Data, DataRow).Item("exportfiledate") = Today.Date
                'Next

                Dim ds2 = DS.GetChanges
                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.PriceChangeExportFile(Me, mye) Then
                        Logger.log(mye.message)
                    Else
                        DS.AcceptChanges()
                    End If
                End If
            End If
            

        Else
            myret = False
        End If
        Return myret
    End Function
    Public Function ExportFile002() As Boolean
        Dim myret As Boolean = True
        Dim sb As New StringBuilder

        If getdataset() Then
            If DS.Tables(0).Rows.Count > 0 Then
                sb.Append("[Variant ID]" & vbTab & "[Variant Text]" & vbTab & "LIFNR" & vbTab & "MATNR" & vbTab & "PURC_ORG" & vbTab & "WERKS" & vbTab & "DATE" & vbTab & "PRICE" & vbTab & "UNIT" & vbCrLf)
                sb.Append("-->" & vbTab & "Parameter texts" & vbTab & "Vendor" & vbTab & "Material" & vbTab & "Purch. Organization" & vbTab & "Plant" & vbTab & "Valid on" & vbTab & "Rate" & vbTab & "Pricing unit" & vbCrLf)
                sb.Append("-->" & vbTab & "Default Values" & vbTab & "#99003700" & vbTab & "#1500631606" & vbTab & "#6101" & vbTab & "#6110" & vbTab & "02.01.2006" & vbTab & "#           10000" & vbTab & "# 1000" & vbCrLf)
                sb.Append("*** Changes to the default values displayed above not effective" & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbCrLf)
                For Each dr As DataRow In DS.Tables(0).Rows
                    sb.Append(vbTab & vbTab & dr.Item("vendorcode") & vbTab & dr.Item("cmmf") & vbTab & dr.Item("purchorg") & vbTab &
                              dr.Item("plant") & vbTab & sapdate(dr.Item("validon")) & vbTab & dr.Item("price") & vbTab & dr.Item("pricingunit") & vbCrLf)

                    'assign flag
                    dr.Item("exportfileid") = seq
                    dr.Item("exportfiledate") = Today.Date
                Next
                Dim fullpathname As String
                fullpathname = DS.Tables(1).Rows(0).Item("cvalue") & "\" & "ZFA2C039_GSF_MAJ_FIA" & "_" & Today.Year & Format(Today.Month, "00") & Format(seq, "000") & ".TXT"
                Using mystream As New StreamWriter(fullpathname, False, System.Text.Encoding.ASCII)
                    mystream.WriteLine(sb.ToString)
                End Using

                myret = True

                'Flag Send to SAP
                'Dim q = From n In HDBS.List
                '        Group n By key = n.item("pricechangehdid") Into Group
                '        Select key, Data = Group

                'For Each n In q
                '    CType(n.Data, DataRow).Item("exportfileid") = seq
                '    CType(n.Data, DataRow).Item("exportfiledate") = Today.Date
                'Next

                Dim ds2 = DS.GetChanges
                If Not IsNothing(ds2) Then
                    Dim mymessage As String = String.Empty
                    Dim ra As Integer
                    Dim mye As New ContentBaseEventArgs(ds2, True, mymessage, ra, True)
                    If Not DbAdapter1.PriceChangeExportFile(Me, mye) Then
                        Logger.log(mye.message)
                    Else
                        DS.AcceptChanges()
                    End If
                End If
            End If


        Else
            myret = False
        End If
        Return myret
    End Function
    Private Function sapdate(ByVal myDate As Date) As Object
        Return (Format(myDate.Day, "00") & "." & Format(myDate.Month, "00") & "." & Format(myDate.Year, "0000"))
    End Function


End Class
