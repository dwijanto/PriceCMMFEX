Imports System.Threading
Imports System.ComponentModel
Imports PriceCMMFExt.PublicClass
Imports System.IO
Imports System.Text
Imports Microsoft.Office.Interop
Imports PriceCMMFExt.SharedClass
'Imports System.Xml
'Imports System.Xml.Xsl

Public Class ReportAVPI
    Dim myCount As Integer = 0
    Dim listcount As Integer = 0

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim QueryDelegate As New ThreadStart(AddressOf DoQuery)

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Delegate Sub ProcessReport(ByVal osheet As Excel.Worksheet)

    Dim myThread As New System.Threading.Thread(myThreadDelegate)
    Dim myQueryThread As New System.Threading.Thread(QueryDelegate)

    'Dim exclude As Boolean = True
    'Dim excludeComments As Boolean = False
    Dim SelectedPath As String = String.Empty
    Dim FullPath As String = String.Empty
    'Dim Status As Boolean = False
    'Dim ReadFileStatus As Boolean = False
    'Dim Dataset1 As DataSet
    'Dim sb As StringBuilder
    'Dim aprocesses() As Process = Nothing '= Process.GetProcesses
    'Dim aprocess As Process = Nothing
    'Dim Source As String
    'Dim excelcaption As String
    Dim hwnd As System.IntPtr
    'Dim Marketbs As New BindingSource
    'Dim VendorBs As New BindingSource
    'Dim SAOBs As New BindingSource
    'Dim comboid As String = String.Empty
    'Dim selectedCheckedListbox As CheckedListBox
    'Dim mycriteria As String = String.Empty
    'Dim myexception As String = String.Empty
    Dim startdate As Date
    Dim enddate As Date
    Dim VendorList As String
    'Dim commentstartdate As Date
    'Dim commentenddate As Date

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If myQueryThread.IsAlive Then
            ProgressReport(5, "Please wait...")
            Exit Sub
        End If

        If Not myThread.IsAlive Then
            'get Criteria
            ToolStripStatusLabel1.Text = ""
            ToolStripStatusLabel2.Text = ""
            ToolStripStatusLabel3.Text = ""

            startdate = DateTimePicker1.Value.Date
            enddate = DateTimePicker2.Value.Date


            ProgressReport(5, "")
            Dim DirectoryBrowser As FolderBrowserDialog = New FolderBrowserDialog
            DirectoryBrowser.Description = "Which directory do you want to use?"

            If (DirectoryBrowser.ShowDialog() = Windows.Forms.DialogResult.OK) Then
                SelectedPath = DirectoryBrowser.SelectedPath

                Try
                    myThread = New System.Threading.Thread(myThreadDelegate)
                    myThread.SetApartmentState(ApartmentState.MTA)
                    myThread.Start()
                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try
            End If

        Else
            MsgBox("Please wait until the current process is finished")
        End If
    End Sub

    Sub DoQuery()
        Dim ds As New DataSet
        Dim sqlstr = "select miropostingdate from miro order by miropostingdate desc limit 1;" &
                     "select vendorname::character varying from orderlinemembers ol" &
                     " left join vendor v on v.vendorcode = ol.customercode" &
                     "  where ol.orderlineid = 15 order by vendorname"
        'Dim myresult As Date
        Dim mymessage As String = String.Empty
        'If DbAdapter1.ExecuteScalar(sqlstr, myresult, mymessage) Then
        '    ProgressReport(6, String.Format("{0:dd-MMM-yyyy}", myresult.Date))
        'End If
        If DbAdapter1.TbgetDataSet(sqlstr, ds, mymessage) Then
            ProgressReport(6, String.Format("{0:dd-MMM-yyyy}", ds.Tables(0).Rows(0).Item(0)))
            VendorList = ""
            For i = 0 To ds.Tables(1).Rows.Count - 1
                VendorList = VendorList + IIf(VendorList = "", "", ",") + ds.Tables(1).Rows(i).Item(0).ToString
            Next
        Else
            ProgressReport(5, mymessage)

        End If

    End Sub

    Sub DoWork()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(5, "Export To Excel..")
        Dim status As Boolean = False
        Dim sr As New ReportAVPIExt
        sr.startdate = startdate
        sr.enddate = enddate

        status = GenerateReport(sr)
        If status Then
            sw.Stop()
            ProgressReport(5, "")
            ProgressReport(5, String.Format("Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            'ProgressReport(2, TextBox2.Text & "Done.")

            If MsgBox("File name: " & FullPath & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                Process.Start(FullPath)
            End If

        Else
            errSB.Append(sr.errormsg & vbCrLf)
            ProgressReport(5, errSB.ToString)
        End If
        sw.Stop()


    End Sub


    Private Function GenerateReport(ByRef sr As ReportAVPIExt) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty

        Try
            'Create Object Excel 
            ProgressReport(5, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.ScreenUpdating = False
            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(5, "Opening Template...")
            ProgressReport(5, "Generating records..")

            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\AVPTemplate.xltx")

            sr.oxl = oXl
            sr.owb = oWb
            sr.isheet = 2

            'Dim events As New List(Of ManualResetEvent)()
            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")

            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter
            Dim mydate1 = DateFormatyyyyMMdd(sr.startdate)
            Dim mydate2 = DateFormatyyyyMMdd(sr.enddate)



            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(validsbu(c.comfam,vpi.sbuname),pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(sbu1.sbuname,pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
            '            " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(cvp.lastprice::numeric - cvp.agv2::numeric) as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '            " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (cvp.lastprice::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
            '            " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
            '            " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '            " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
            '            " LEFT JOIN activity ac ON ac.activitycode = c.rir" &
            '            " LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
            '            " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
            '            " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '            " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '            " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '            " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '            " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '            " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '            " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '            " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '            " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '            " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
            '            " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
            '            " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(validsbu(c.comfam,vpi.sbuname),pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(sbu1.sbuname,pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
            '            " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric)  as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '            " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (cvp.lastprice::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
            '            " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
            '            " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '            " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
            '            " LEFT JOIN activity ac ON ac.activitycode = c.rir" &
            '            " LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
            '            " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
            '            " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '            " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '            " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '            " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '            " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '            " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '            " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '            " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '            " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '            " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
            '            " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
            '            " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(validsbu(c.comfam,vpi.sbuname),pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(sbu1.sbuname,pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
            '            " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric)  as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '            " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
            '            " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
            '            " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '            " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
            '            " LEFT JOIN activity ac ON ac.activitycode = c.rir" &
            '            " LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
            '            " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
            '            " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '            " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '            " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '            " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '            " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '            " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '            " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '            " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '            " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '            " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
            '            " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
            '            " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(validsbu(c.comfam,vpi.sbuname),pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(validsbu(c.comfam,sbu1.sbuname),pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
            '            " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric)  as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '            " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
            '            " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
            '            " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '            " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
            '            " LEFT JOIN activity ac ON ac.activitycode = c.rir" &
            '            " LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
            '            " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
            '            " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '            " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '            " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '            " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '            " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '            " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '            " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '            " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '            " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '            " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
            '            " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
            '            " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, mm.rri,mm.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, mm.familylv1 as comfam, f.familyname, validvpi(validsbu(mm.familylv1,vpi.sbuname),pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(validsbu(mm.familylv1,sbu1.sbuname),pg.purchasinggroup,pmo.officersebname) as sbu, mm.plant, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(mm.cmmf,m.miropostingdate)) AS stdprice," & _
            '            " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(mm.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(mm.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric)  as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '            " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(mm.cmmf,m.miropostingdate))) as towstd, " & _
            '            " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode,s.sbuname as sbusap  FROM pomiro pm" & _
            '            " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '            " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf LEFT JOIN materialmaster mm ON mm.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = mm.familylv1 left join range r on r.range = mm.range" & _
            '            " LEFT JOIN activity ac ON ac.activitycode = mm.rri" &
            '            " LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid Left join sbusap s on s.sbuid = mm.sbu" & _
            '            " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
            '            " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '            " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '            " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '            " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '            " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '            " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '            " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '            " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '            " union all (select * from  getporeversedcurrsbu(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '            " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying,sbuname2 character varying ))" & _
            '            " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null,null" & _
            '            " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, mm.rri,mm.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, mm.familylv1 as comfam, f.familyname, validvpi(validsbu(mm.familylv1,s.pi_vpi),pg.purchasinggroup) AS vpi, validvpi(s.pi_sbuname,pg.purchasinggroup) as sbuname,validvpi(validsbu(mm.familylv1,s.pi_sbu),pg.purchasinggroup) as sbu, mm.plant, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(mm.cmmf,m.miropostingdate)) AS stdprice," & _
                        " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(mm.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(mm.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric)  as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
                        " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (getlkpamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp.lastprice)::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(mm.cmmf,m.miropostingdate))) as towstd, " & _
                        " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode,s.sbuname as sbusap  FROM pomiro pm" & _
                        " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
                        " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
                        " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf LEFT JOIN materialmaster mm ON mm.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = mm.familylv1 left join range r on r.range = mm.range" & _
                        " LEFT JOIN activity ac ON ac.activitycode = mm.rri" &
                        " LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid Left join sbusap s on s.sbuid = mm.sbu" & _
                        " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
                        " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
                        " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
                        " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
                        " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
                        " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
                        " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
                        " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
                        " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
                        " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
                        " union all (select * from  getporeversedcurrsbu(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
                        " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying,sbuname2 character varying ))" & _
                        " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null,null" & _
                        " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
                        " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            obj.Name = "DATA"
            obj.osheet = oWb.Worksheets("DATA")
            If CreateWorksheet(obj) > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable1(sr)
                'ProgressReport(5, "Creating Charts..")
                'CreateChart1(oWb, 1, sr)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            FullPath = ValidateFileName(SelectedPath & "\API-" & String.Format("{0:yyyyMMdd}", sr.startdate) & "-" & String.Format("{0:yyyyMMdd}", sr.enddate) & ".xlsx")
            ProgressReport(5, "Done ")
            ProgressReport(2, "Saving File ...")
            oWb.SaveAs(FullPath)
            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
            oXl.ScreenUpdating = True

        Catch ex As Exception
            sr.errormsg = ex.Message
        Finally
            'ProgressReport(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
        Return result
    End Function
    Private Function GenerateReport1(ByRef sr As ReportAVPIExt) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty

        Try
            'Create Object Excel 
            ProgressReport(5, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            oXl.ScreenUpdating = False
            oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(5, "Opening Template...")
            ProgressReport(5, "Generating records..")

            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\AVPTemplate.xltx")

            sr.oxl = oXl
            sr.owb = oWb
            sr.isheet = 2

            'Dim events As New List(Of ManualResetEvent)()
            Dim counter As Integer = 0
            ProgressReport(5, "Creating Worksheet...")

            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter
            Dim mydate1 = DateFormatyyyyMMdd(sr.startdate)
            Dim mydate2 = DateFormatyyyyMMdd(sr.enddate)

            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,ph.payt,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(validsbu(c.comfam,vpi.sbuname),pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(sbu1.sbuname,pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
            '             " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(cvp.lastprice::numeric - cvp.agv2::numeric) as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '             " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (cvp.lastprice::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
            '             " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
            '             " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '             " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '             " LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
            '             " LEFT JOIN activity ac ON ac.activitycode = c.rir LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
            '             " LEFT JOIN paymentterm pt on pt.payt = ph.payt" & _
            '             " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '             " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '             " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '             " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '             " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '             " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '             " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '             " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '             " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '             " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '             " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
            '             " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
            '             " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '             " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            'obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(vpi.sbuname,pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(sbu1.sbuname,pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
            '            " sdhd.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(cvp.lastprice::numeric - cvp.agv2::numeric) as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
            '            " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (cvp.lastprice::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
            '            " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
            '            " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
            '            " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
            '            " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
            '            " LEFT JOIN activity ac ON ac.activitycode = c.rir LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
            '            " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
            '            " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
            '            " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
            '            " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
            '            " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
            '            " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
            '            " LEFT JOIN aasdpo sdpo ON sdpo.pohd = ph.pohd AND sdpo.poitem = pd.polineno" & _
            '            " LEFT JOIN aasdhd sdhd ON sdhd.salesdoc = sdpo.salesdoc LEFT JOIN customer cust ON cust.customercode = sdhd.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
            '            " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
            '            " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
            '            " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
            '            " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
            '            " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
            '            " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            obj.strsql = "( SELECT ph.pohd, pd.polineno, ph.pono, pd.cmmf, c.rir,c.materialdesc, ph.purchasinggroup, m.vendorcode, v.vendorname,v.shortname,e.termsofpayment,  m.supplierinvoicenum, m.mironumber, m.miropostingdate, pm.crcy as originalcurrency, getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) as amount,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty) as newamount ,  validstdprice(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountfp,validmould(pg.purchasinggroup,getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) -( validnum(agv.value) * pm.qty)) as newamountwomoulds, pm.qty, pd.oun, pm.pomiroid, c.comfam, f.familyname, validvpi(vpi.sbuname,pg.purchasinggroup,pmo.officersebname) AS vpi, validvpi(sbu.sbuname,pg.purchasinggroup,pmo.officersebname) as sbuname,validvpi(sbu1.sbuname,pg.purchasinggroup,pmo.officersebname) as sbu, c.plnt, sct.category, c.cmmftype, (getvalidpricesap(pd.cmmf,m.vendorcode,m.miropostingdate) / getexrate(ph.pohd,pd.polineno))::numeric(18,4) AS validpricesap, c.eol, validstdprice(pg.purchasinggroup,getstdcost(c.cmmf,m.miropostingdate)) AS stdprice," & _
                        " sdpo.shiptoparty, cust.customername AS shiptopartyname,validofficername(of.officername) as spm,pmo.officersebname as pm, validnum(agv.value) as amort, ((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty) - validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1 as variance,(((getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate) / pm.qty )- validnum(agv.value) - validstdprice(pg.purchasinggroup, getstdcost(c.cmmf,m.miropostingdate))) * -1) * pm.qty as ""tovariance"" , cvp.averprice::numeric as ""averpricey-1"",(cvp.lastprice::numeric - cvp.agv2::numeric) as ""lastpricey-1"",getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice) - cvp1.agv1::numeric as ""initialprice""," & _
                        " case when cvp.averprice is null then (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else cvp.averprice::numeric * qty::numeric End as ""towavpy-1"", case when cvp.lastprice is null then  (getinitialpriceamount(pd.cmmf,m.vendorcode,date_part('year',m.miropostingdate)::integer,cvp1.initialprice)::numeric - cvp1.agv1::numeric) * qty::numeric Else (cvp.lastprice::numeric - cvp.agv2::numeric) * qty::numeric End as ""towlkpy-1"", qty::numeric * validstdprice(pg.purchasinggroup,(getstdcost(c.cmmf,m.miropostingdate))) as towstd, " & _
                        " case when odm.customercode is null then validgroupact(gs.groupsbuname,pmo.officersebname) Else case when  ph.purchasinggroup = 'FOD' or ph.purchasinggroup = 'FOF' or ph.purchasinggroup = 'FOG' then  null else 'SUPOR'::text end end as groupsbu,validgroupact(gs1.groupsbuname,pmo.officersebname) as groupact,pt.days as avrpayt,pt.days::numeric * (getpocurramount(ph.pohd,pd.polineno,pm.crcy,pm.amount,pm.qty,m.miropostingdate)::numeric -( validnum(agv.value)::numeric * pm.qty::numeric)) as amtwpayt,r.range,r.rangedesc,c.modelcode FROM pomiro pm" & _
                        " LEFT JOIN miro m ON m.miroid = pm.miroid" & _
                        " LEFT JOIN podtl pd ON pd.podtlid = pm.podtlid" & _
                        " Left join ekko e on e.po = pd.pohd LEFT JOIN pohd ph ON ph.pohd = pd.pohd LEFT JOIN cmmf c ON c.cmmf = pd.cmmf  LEFT JOIN family f ON f.familyid = c.comfam left join range r on r.rangeid = c.rangeid" & _
                        " LEFT JOIN activity ac ON ac.activitycode = c.rir LEFT JOIN sbu vpi ON vpi.sbuid = ac.sbuidvpi  LEFT JOIN sbu ON sbu.sbuid = ac.sbuidlg left join sbu sbu1 on sbu1.sbuid = ac.sbuid" & _
                        " LEFT JOIN paymentterm pt on pt.payt = e.termsofpayment" & _
                        " LEFT JOIN vendor v ON v.vendorcode = m.vendorcode" & _
                        " left join supplierspanel spl on spl.vendorcode = v.vendorcode" & _
                        " Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
                        " Left JOIN officer of on of.officerid = v.officerid left join officerseb pmo on pmo.ofsebid = v.pmid  " & _
                        " left join purchasinggroup pg on pg.purchasinggroup = ph.purchasinggroup left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = m.vendorcode left join vendor vs on vs.vendorcode = odm.customercode" & _
                        " left join groupsbu gs1 on gs1.groupsbuid = pg.groupact" & _
                        " LEFT JOIN cxsebpodtl sdpo ON sdpo.sebasiapono = ph.pohd AND sdpo.polineno = pd.polineno" & _
                        " LEFT JOIN customer cust ON cust.customercode = sdpo.shiptoparty  left join agreementtx agtx on agtx.material = pd.cmmf and agtx.postingdate = m.miropostingdate and agtx.status left join agvalue agv on agv.agreement = agtx.agreement  " & _
                        " left join cmmfvendorprice cvp on cvp.cmmf = pd.cmmf and cvp.vendorcode = m.vendorcode and cvp.myyear = " & Year(sr.startdate) - 1 & " left join cmmfvendorprice cvp1 on cvp1.cmmf = pd.cmmf and cvp1.vendorcode = m.vendorcode and cvp1.myyear = " & Year(sr.startdate) & " where ph.purchasinggroup <> 'FOE' and m.miropostingdate >= " & mydate1 & " and m.miropostingdate <= " & mydate2 & ")" & _
                        " union all (select * from  getporeversedcurr(" & Year(sr.startdate) & "," & DateFormatyyyyMMdd(sr.startdate) & "," & DateFormatyyyyMMdd(sr.enddate) & ") as pr(pohd bigint , polineno integer,pono character varying,cmmf bigint,rir character varying,materialdesc character,purchasinggroup character varying,vendorcode bigint,vendorname character,shortname2 text,payt character varying,supplierinvoicenum character varying,mironumber bigint,miropostingdate date," & _
                        " crcy character varying, amount numeric,newamount numeric,newamountfp numeric,newamountwomoulds numeric, qty numeric,oun character varying,reversedby bigint,comfam integer,familyname character,vpi text,  sbuname text,sbu text,plnt integer,category character,cmmftype character varying,validpricesap numeric,eol boolean,stdprice numeric,shiptoparty bigint,shiptopartyname character,spm text,pm character,amort numeric, variance numeric,""tovariance"" numeric,""averpricey-1"" numeric,""lastprice-y"" numeric,""initialprice"" numeric,""towavpy-1"" numeric, ""towlkpy-1"" numeric, towstd numeric,groupsbu text,groupact text, avrpayt integer, amtwpayt numeric,range character varying,rangedesc character varying,modelcode character varying ))" & _
                        " union all (select null,null,null,null,null,ma.description,'FO9', ma.vendorcode,v.vendorname,v.shortname2 as shortname,null,null,null,ma.period as miropostingdate,null,ma.amount,ma.amount as newamount,validstdprice(ma.pg,ma.amount) as newamountfp,validmould(ma.pg,ma.amount) as newamountwomoulds,null,null, null,ma.familyid, f.familyname,vpi.sbuname as vpiname,null,sbu.sbuname as sbu,null,sct.category,'A',null,null,null,null,null,of.officername,pm.officersebname,null,null,ma.amount * - 1 ,null,null ,null ,null,null,null,case when odm.customercode is null then  gs.groupsbuname Else vs.ShortName2 end as groupsbu,gs1.groupsbuname as groupact,null,null::numeric,null,null,null" & _
                        " from manualadjustment ma left join vendor v on v.vendorcode = ma.vendorcode left join family f on f.familyid = ma.familyid left join officer of on of.officerid = ma.ssm::text left join officerseb pm on pm.ofsebid = v.pmid left join groupingcodesbu gcs on gcs.groupingcode = ma.grouping left join sbu on sbu.sbuid = gcs.sbuid left join sbu vpi on vpi.sbuid = gcs.vpiid left join orderlinemembers odm on odm.orderlineid = 15 and odm.customercode = ma.vendorcode left join vendor vs on vs.vendorcode = odm.customercode left join purchasinggroup pg on pg.purchasinggroup = ma.pg left join groupsbu gs on gs.groupsbuid = pg.groupsbuid left join groupsbu gs1 on gs1.groupsbuid = pg.groupact left join supplierspanel spl on spl.vendorcode = ma.vendorcode  Left join supplierscategory sct on sct.supplierscategoryid = spl.supplierscategoryid" & _
                        " where period >= " & mydate1 & " and period <=  " & mydate2 & ")"

            obj.Name = "DATA"
            obj.osheet = oWb.Worksheets("DATA")
            If CreateWorksheet(obj) > 1 Then
                ProgressReport(5, "Generating Pivot Tables..")
                CreatePivotTable1(sr)
                'ProgressReport(5, "Creating Charts..")
                'CreateChart1(oWb, 1, sr)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            FullPath = ValidateFileName(SelectedPath & "\API-" & String.Format("{0:yyyyMMdd}", sr.startdate) & "-" & String.Format("{0:yyyyMMdd}", sr.enddate) & ".xlsx")
            ProgressReport(5, "Done ")
            ProgressReport(2, "Saving File ...")
            oWb.SaveAs(FullPath)
            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
            oXl.ScreenUpdating = True

        Catch ex As Exception
            sr.errormsg = ex.Message
        Finally
            'ProgressReport(3, "Releasing Memory...")
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
        Return result
    End Function

    'Private Shared Function WaitForAll(ByVal events As ManualResetEvent()) As Boolean
    '    Dim result As Boolean = False
    '    Try
    '        If events IsNot Nothing Then
    '            For i As Integer = 0 To events.Length - 1
    '                events(i).WaitOne()
    '            Next
    '            result = True
    '        End If
    '    Catch
    '        result = False
    '    End Try
    '    Return result
    'End Function
    Private Sub ProgressReport(ByVal id As Integer, ByRef message As String)
        If Me.Label4.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 2
                    'TextBox2.Text = message
                    Me.ToolStripStatusLabel1.Text = message
                Case 3
                    'TextBox3.Text = message
                    Me.ToolStripStatusLabel2.Text = message
                Case 4
                    'TextBox1.Text = message
                    Me.ToolStripStatusLabel3.Text = message
                Case 5
                    'ToolStripStatusLabel1.Text = message
                    'ComboBox1.DataSource = bs
                    'ComboBox1.DisplayMember = "typeofitem"
                    'ComboBox1.ValueMember = "typeofitemid"
                    'Me.Label4.Text = message
                    Me.ToolStripStatusLabel1.Text = message
                Case 6
                    Label4.Text = message

                Case 7

            End Select

        End If

    End Sub

    Private Sub ReportAVPI_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Application.DoEvents()
        'Load the query in background
        myQueryThread.Start()
    End Sub

    'Sub DoQuery()
    '    Dataset1 = New DataSet
    '    Dim sqlstr As String = "select 'All' as customername,0 as shiptoparty union all select 'All Market',1 as shiptoparty union all (select distinct c.customername,s.shiptoparty from cxsebpodtl s" &
    '                           " left join customer c on c.customercode = s.shiptoparty order by c.customername);" &
    '                           " select 'All' as vendorname,0 as vendorcode union all select 'All Supplier',1  union all(select v.vendorname,p.vendorcode from (select distinct vendorcode from povendor ) as p" &
    '                           " left join vendor v on v.vendorcode = p.vendorcode order by v.vendorname);" &
    '                           " select 'All' as officersebname,0 as ofsebid union all select 'All SAO',1  union all(select distinct o.officersebname ,c.ofsebid from (select distinct s.shiptoparty from cxsebpodtl s ) as s" &
    '                           " left join customer c on c.customercode = s.shiptoparty " &
    '                           " left join officerseb o on o.ofsebid = c.ofsebid where not o.ofsebid isnull order by officersebname );" &
    '                           " select customername from orderlinemembers om" &
    '                           " left join orderline o on o.orderlineid = om.orderlineid " &
    '                           " where o.orderlinename = 'ExcludeComments'"

    '    If DbAdapter1.TbgetDataSet(sqlstr, Dataset1) Then
    '        Dataset1.Tables(0).TableName = "Market"
    '        Marketbs.DataSource = Dataset1.Tables(0)
    '        VendorBs.DataSource = Dataset1.Tables(1)
    '        SAOBs.DataSource = Dataset1.Tables(2)
    '        Dataset1.Tables(3).TableName = "ExcludeComments"
    '        ProgressReport(6, "")

    '    Else
    '        ProgressReport(5, "Error while loading Dataset.")
    '    End If

    'End Sub

    Private Function CreateWorksheet(ByVal obj As Object) As Long
        Dim osheet = DirectCast(obj, ThreadPoolObj).osheet
        osheet.Name = DirectCast(obj, ThreadPoolObj).Name
        ProgressReport(5, "Waiting for the query to be executed..." & DirectCast(obj, ThreadPoolObj).osheet.Name)
        Dim sqlstr = DirectCast(obj, ThreadPoolObj).strsql
        FillWorksheet(osheet, sqlstr, DbAdapter1)
        Dim lastrow = osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row
        Return lastrow

    End Function

    Private Sub CreatePivotTable1(ByRef sr As ReportAVPIExt)
        ProgressReport(5, "Pivot Table 1 of 10...")


        Dim osheet As Excel.Worksheet
        Dim owb = sr.owb
        Dim isheet = sr.isheet
        Dim oxl = sr.oxl
        Dim PivotSource As Integer
        owb.Worksheets(isheet).select()
        osheet = owb.Worksheets(isheet)
        owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        PivotSource = isheet
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With


        'osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.Range("A9").Group(True, True, Periods:={False, False, False, False, True, False, True})
        'osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        'For Each item As Object In osheet.PivotTables("PivotTable1").pivotfields("Years").pivotitems
        '    Dim obj = DirectCast(item, Excel.PivotItem)
        '    If obj.Value.ToString <> startdate.Year.ToString Then
        '        obj.Visible = False
        '    End If
        'Next
        'osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        'osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        ''osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").numberformat = "MMM-yy"
        'osheet.Columns("A:A").numberformat = "mmm-yy"

        'Calculated Fields
        osheet.PivotTables("PivotTable1").calculatedfields.add("AmountK", "= newamount/1000", True)
        osheet.PivotTables("PivotTable1").CalculatedFields.Add("QTYK", "= qty/1000", True)
        osheet.PivotTables("PivotTable1").CalculatedFields.Add("TOVarianceK", "= tovariance/1000", True)
        osheet.PivotTables("PivotTable1").CalculatedFields.Add("YTDvsstdY", "= if(iserror(((newamountfp -towstd )/towstd +1)*100),0,((newamountfp -towstd )/towstd +1)*100)", True)
        osheet.PivotTables("PivotTable1").CalculatedFields.Add("YTDvsY-1", "= if(iserror(((newamountwomoulds-'towavpy-1')/'towavpy-1'+1)*100),0,((newamountwomoulds-'towavpy-1')/'towavpy-1'+1)*100)", True)
        osheet.PivotTables("PivotTable1").CalculatedFields.Add("YTDvsLKP-1", "= if(iserror(((newamountwomoulds-'towlkpy-1' )/'towlkpy-1' +1)*100),0,((newamountwomoulds-'towlkpy-1' )/'towlkpy-1' +1)*100)", True)


        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("groupsbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("groupsbu").Caption = "Group"
        osheet.PivotTables("PivotTable1").PivotFields("sbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("sbu").Caption = " SBU"
        'oSheet.PivotTables("PivotTable1").PivotFields("Group").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"
        osheet.Range("C10").Group(True, True, Periods:={False, False, False, False, True, False, True})
        'osheet.Range("D9").Group(Start:=True, End:=True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable1").PivotFields(" Month").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        Call HideList("(blank)", osheet, "PivotTable1", " SBU")

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "TO Variance Y", Excel.XlConsolidationFunction.xlSum)


        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"
        osheet.Cells.EntireColumn.AutoFit()

        Call Beauty(oxl, osheet, "10:11", 3)


        'Worksheet by supplier
        ProgressReport(5, "Pivot Table 2 of 10...")
        owb.Worksheets("By Supplier").Select()
        osheet = owb.Worksheets("By Supplier")
        osheet.Name = "By_Supplier"
        'owb.Worksheets("Summary").PivotTables("PivotTable1").PivotCache.CreatePivotTable("PivotTables!R7C10", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        owb.Worksheets("Summary").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C2", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "By Supplier"



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = " Supplier Name"
        osheet.PivotTables("PivotTable1").PivotFields(" Supplier Name").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'oSheet.PivotTables("PivotTable1").PivotFields("purchasinggroup").Orientation =Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").PivotFields("purchasinggroup").Caption = " Purchasing Group"
        osheet.PivotTables("PivotTable1").PivotFields("category").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("category").Caption = " Panel"

        osheet.PivotTables("PivotTable1").PivotFields(" Panel").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").PivotFields(" Purchasing Group").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"

        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("variance"), " Variance Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "To Variance Y", Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"

        osheet.Cells.EntireColumn.AutoFit()



        Call HideList(getVendorList, osheet, "PivotTable1", " Supplier Name")
        osheet.PivotTables("PivotTable1").PivotFields(" Supplier Name").AutoSort(Excel.XlSortOrder.xlDescending, "Pur.Amount Y")
        osheet.Range("A12:A13").Select()
        osheet.Range("A12:A13").AutoFill(Destination:=osheet.Range("A12:A" & getRow(osheet) - 1))
        'copycells
        '    oSheet.Cells(12, 1) = "=IF(B12<>"""",IF(A11<>"""",A11+1, IF(A10<>"""", A10+1,IF(A9<>"""", A9+1, IF(A8<>"""",A8+1,IF(A7<>"""",A7+1, FALSE))))),"""")"
        '    oSheet.Range("A12").Copy
        '    oSheet.Range("A13:A" & getRow(oSheet) - 1).PasteSpecial (xlPasteAll)

        Call Beauty(oxl, osheet, "10:11", 4)

        ProgressReport(5, "Pivot Table 3 of 10...")
        'mysheet = mysheet + 1
        owb.Worksheets("By SPM").Select()
        osheet = owb.Worksheets("By SPM")
        osheet.Name = "By_SPM"
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "By SPM"



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("spm").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("spm").Caption = " SPM"

        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = " FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField


        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"

        'osheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = "Supplier Name"
        'osheet.PivotTables("PivotTable1").PivotFields("comfam").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("comfam").Caption = "Family Code"
        'osheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("familyname").Caption = "Family"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("variance"), " Variance Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "To Variance Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'osheet.PivotTables("PivotTable1").PivotFields("Family Code").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"

        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = "Supplier Name"
        osheet.PivotTables("PivotTable1").PivotFields("comfam").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("comfam").Caption = "Family Code"
        osheet.PivotTables("PivotTable1").PivotFields("Family Code").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("familyname").Caption = "Family"
        osheet.Cells.EntireColumn.AutoFit()



        Call HideList(getVendorList, osheet, "PivotTable1", "Supplier Name")
        Call Beauty(oxl, osheet, "10:11", 5)
        'Pivot 4
        ProgressReport(5, "Pivot Table 4 of 10...")
        'mysheet = mysheet + 1
        owb.Worksheets("By VPI").Select()
        osheet = owb.Worksheets("By VPI")
        osheet.Name = "By_VPI"
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "By VPI"



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("vpi").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("vpi").Caption = " VPI"
        osheet.PivotTables("PivotTable1").PivotFields("comfam").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("comfam").Caption = "Family Code"
        osheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("familyname").Caption = "Family"


        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = "Supplier"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        Call HideList(getVendorList, osheet, "PivotTable1", "Supplier")
        Call HideList("(blank)", osheet, "PivotTable1", " VPI")

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("variance"), " Variance Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "TO Variance Y", Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"
        osheet.PivotTables("PivotTable1").PivotFields("Family Code").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.Cells.EntireColumn.AutoFit()


        Call Beauty(oxl, osheet, "10:11", 4)


        'Pivot 5
        ProgressReport(5, "Pivot Table 5 of 10...")
        'mysheet = mysheet + 1
        owb.Worksheets("By SBU").Select()
        osheet = owb.Worksheets("By SBU")
        osheet.Name = "By_SBU"
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "By SBU"



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("sbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("sbu").Caption = " SBU"
        osheet.PivotTables("PivotTable1").PivotFields("comfam").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("comfam").Caption = "Family Code"
        osheet.PivotTables("PivotTable1").PivotFields("familyname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("familyname").Caption = " Family"
        osheet.PivotTables("PivotTable1").PivotFields("Family Code").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = "Supplier"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("variance"), " Variance Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "TO Variance Y", Excel.XlConsolidationFunction.xlSum)
        Call HideList("(blank)", osheet, "PivotTable1", " SBU")
        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"
        osheet.Cells.EntireColumn.AutoFit()



        Call HideList(getVendorList, osheet, "PivotTable1", "Supplier")
        Call Beauty(oxl, osheet, "10:11", 4)
        'Pivot 6
        ProgressReport(5, "Pivot Table 6 of 10...")
        'mysheet = mysheet + 1
        owb.Worksheets("By Site").Select()
        osheet = owb.Worksheets("By Site")
        osheet.Name = "By_Site"
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "By Site"


        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("shiptopartyname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("shiptopartyname").Caption = " Ship To Party Name"


        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = "Supplier"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField


        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("variance"), " Variance Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "TO Variance Y", Excel.XlConsolidationFunction.xlSum)


        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"
        Call HideList(getVendorList, osheet, "PivotTable1", "Supplier")
        osheet.Cells.EntireColumn.AutoFit()


        Call HideList(getVendorList, osheet, "PivotTable1", "Supplier")
        Call Beauty(oxl, osheet, "10:11", 2)

        'mysheet = mysheet + 1
        ProgressReport(5, "Pivot Table 7 of 10...")
        owb.Worksheets("TO By Supplier").Select()
        osheet = owb.Worksheets("TO By Supplier")
        osheet.Name = "TOBySupplier"
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "TO By Supplier"




        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "TURNOVER BY PANEL CATEGORY", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("category").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("category").Caption = "Panel Category"

        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"
        osheet.PivotTables("PivotTable1").PivotFields("groupsbu").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupsbu").Caption = "Criteria"




        osheet.PivotTables("PivotTable1").DisplayErrorString = True

        osheet.PivotTables("PivotTable1").PivotFields("TURNOVER BY PANEL CATEGORY").NumberFormat = "#,#0"
        osheet.Cells.EntireColumn.AutoFit()
        Call HideList("SSEAC,SUPOR,MOULD,SP", osheet, "PivotTable1", "Criteria")



        osheet.Cells(1, 1) = "TURNOVER BY SUPPLIER and PANEL CATEGORY"
        osheet.Cells(2, 1) = "Purchasing Amount in K USD"



        osheet.Rows("1:8").Font.Bold = True
        osheet.Name = "TOBySupplier"
        Dim myIrow As Integer
        myIrow = getRow(osheet)
        ProgressReport(5, "Pivot Table 8 of 10...")
        osheet.Rows(myIrow + 1 & ":" & myIrow + 5).Font.Bold = True
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R" & myIrow + 5 & "C1", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        'oSheet.Name = "TO By Supplier"
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With

        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "TURNOVER BY PANEL CATEGORY", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable2").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").PivotFields("groupact").Caption = "FP/CP"
        osheet.PivotTables("PivotTable2").PivotFields("category").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").PivotFields("category").Caption = "Panel Category"
        '    oSheet.PivotTables("PivotTable2").PivotFields("shortname2").Orientation =Excel.XlPivotFieldOrientation.xlRowField
        '    oSheet.PivotTables("PivotTable2").PivotFields("shortname2").Caption = "Short Name"
        osheet.PivotTables("PivotTable2").PivotFields("shortname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable2").PivotFields("shortname").Caption = "Short Name"

        osheet.PivotTables("PivotTable2").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable2").PivotFields("miropostingdate").Caption = " Month"

        osheet.PivotTables("PivotTable1").TableStyle2 = "Variance Report 2"
        osheet.PivotTables("PivotTable2").TableStyle2 = "Variance Report 2"
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.Name = "TO By Supplier"
        osheet.Cells.EntireColumn.AutoFit()



        owb.Worksheets("Summary FP").Select()
        osheet = owb.Worksheets("Summary FP")
        osheet.Name = "SummaryFP"
        ProgressReport(5, "Pivot Table 9 of 10...")
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "Summary FP"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("groupsbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("groupsbu").Caption = "Group"
        osheet.PivotTables("PivotTable1").PivotFields("sbu").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("sbu").Caption = " SBU"
        'oSheet.PivotTables("PivotTable1").PivotFields("Group").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"
        'oSheet.Range("D9").Group Start:=True, End:=True, Periods:=Array(False, False, False, False, True, False, True)
        osheet.PivotTables("PivotTable1").PivotFields(" Month").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("purchasinggroup").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("purchasinggroup").Caption = " Purchasing Group"
        osheet.PivotTables("PivotTable1").PivotFields(" Purchasing Group").CurrentPage = "FO9"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("cmmftype").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("cmmftype").Caption = "CMMF TYPE"
        osheet.PivotTables("PivotTable1").PivotFields("CMMF TYPE").CurrentPage = "A"

        Call HideList("(blank)", osheet, "PivotTable1", " SBU")

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "TO Variance Y", Excel.XlConsolidationFunction.xlSum)


        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"
        osheet.Cells.EntireColumn.AutoFit()



        Call Beauty(oxl, osheet, "10:11", 3)



        ProgressReport(5, "Pivot Table 10 of 10...")

        owb.Worksheets("By CMMF").Select()
        osheet = owb.Worksheets("By CMMF")
        osheet.Name = "By_CMMF"
        owb.Worksheets(PivotSource).PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R9C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With
        osheet.Name = "By CMMF"



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("QTYK"), "QTY Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("AmountK"), "Pur.Amount Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields("spm").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("spm").Caption = " SPM"

        osheet.PivotTables("PivotTable1").PivotFields("groupact").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("groupact").Caption = " FP/CP"
        osheet.PivotTables("PivotTable1").PivotFields("Years").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("category").Orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").PivotFields("category").Caption = "Panel"
        osheet.PivotTables("PivotTable1").PivotFields(" FP/CP").CurrentPage = "FP"

        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("miropostingdate").Caption = " Month"

        osheet.PivotTables("PivotTable1").PivotFields("pm").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").PivotFields("vendorname").Caption = "Supplier Name"
        'osheet.PivotTables("PivotTable1").PivotFields("cmmf").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        ''oSheet.PivotTables("PivotTable1").PivotFields("comfam").Caption = "Family Code"
        'osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Caption = "Description"
        ''oSheet.PivotTables("PivotTable1").PivotFields("shortname2").Orientation =Excel.XlPivotFieldOrientation.xlRowField
        ''oSheet.PivotTables("PivotTable1").PivotFields("shortname2").Caption = "Supplier Short Name"
        'osheet.PivotTables("PivotTable1").PivotFields("shortname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("shortname").Caption = "Supplier Short Name"
        'osheet.PivotTables("PivotTable1").PivotFields("cmmf").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsY-1"), "Actual vs Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsLKP-1"), "Actual vs LKP Y-1", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("YTDvsstdY"), "Actual vs STD Y", Excel.XlConsolidationFunction.xlSum)
        'oSheet.PivotTables("PivotTable1").AddDataField oSheet.PivotTables("PivotTable1").PivotFields("variance"), " Variance Y",Excel.XlConsolidationFunction.xlSum
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tovarianceK"), "To Variance Y", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        'osheet.PivotTables("PivotTable1").PivotFields("Description").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").DisplayErrorString = True
        osheet.PivotTables("PivotTable1").PivotFields("QTY Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Pur.Amount Y").NumberFormat = "#,#0"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs LKP Y-1").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("Actual vs STD Y").NumberFormat = "#,#0.00"
        osheet.PivotTables("PivotTable1").PivotFields("TO Variance Y").NumberFormat = "#,#0.0"
        osheet.PivotTables("PivotTable1").PivotFields("cmmf").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("cmmf").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").PivotFields("comfam").Caption = "Family Code"
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("materialdesc").Caption = "Description"
        osheet.PivotTables("PivotTable1").PivotFields("Description").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'oSheet.PivotTables("PivotTable1").PivotFields("shortname2").Orientation =Excel.XlPivotFieldOrientation.xlRowField
        'oSheet.PivotTables("PivotTable1").PivotFields("shortname2").Caption = "Supplier Short Name"
        osheet.PivotTables("PivotTable1").PivotFields("shortname").Orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("shortname").Caption = "Supplier Short Name"


        osheet.Cells.EntireColumn.AutoFit()



        Call HideList("", osheet, "PivotTable1", " SPM")
        owb.ShowPivotTableFieldList = False

        Call Beauty(oxl, osheet, "9:11", 6)

        owb.Worksheets(1).Select()

    End Sub

    Private Sub HideList(ByRef myList As String, ByVal oSheet As Excel.Worksheet, ByVal myPivot As String, ByVal myField As String)
        Dim i As Integer
        Try
            For i = 1 To oSheet.PivotTables(myPivot).PivotFields(myField).PivotItems.Count
                If InStr(1, myList, oSheet.PivotTables(myPivot).PivotFields(myField).PivotItems(i).value) > 0 Or Len(oSheet.PivotTables(myPivot).PivotFields(myField).PivotItems(i).value) = 0 Then '
                    oSheet.PivotTables(myPivot).PivotFields(myField).PivotItems(i).Visible = False
                End If
            Next i
        Catch ex As Exception
            Debug.Print(ex.Message)
        End Try
        
    End Sub

    Private Sub Beauty(ByRef oXL As Excel.Application, ByRef oSheet As Excel.Worksheet, ByVal myRange As String, ByVal FirstCol As Integer)
        Dim ytdg As Integer
        Dim lastg As Integer
        Dim maxcol As Integer

        maxcol = getColumn(oSheet)

        oSheet.Cells(1, 1) = oSheet.Name
        oSheet.Cells(2, 1) = "Purchasing Amount in K USD"
        oSheet.Cells(3, 1) = "Quantity in K PCS"
        oSheet.Cells(4, 1) = "TO Variance in K USD: (-) Overcost/(+) Saving vs STD"

        oSheet.Rows("1:11").Font.Bold = True
        oSheet.PivotTables("PivotTable1").TableStyle2 = "Variance Report 2"
        oSheet.Range(oSheet.Cells(1, FirstCol), oSheet.Cells(1, maxcol)).ColumnWidth = 8.5

        oSheet.Columns(FirstCol).Select()
        oXL.ActiveWindow.FreezePanes = True

        With oSheet.Range(myRange)
            .WrapText = True
            .HorizontalAlignment = Excel.Constants.xlCenter
            .VerticalAlignment = Excel.Constants.xlTop
            .EntireRow.AutoFit()
        End With


        ytdg = maxcol - 6
        lastg = ytdg - 6
        If lastg > FirstCol Then
            oSheet.Range(oSheet.Cells(1, FirstCol), oSheet.Cells(1, ytdg)).Columns.Group()
            oSheet.Range(oSheet.Cells(1, FirstCol), oSheet.Cells(1, lastg)).Columns.Group()
            oSheet.Outline.ShowLevels(RowLevels:=0, ColumnLevels:=2)
        End If


    End Sub
    Public Function getColumn(ByVal mysheet As Excel.Worksheet) As Long
        Dim mydata As String
        Dim myarr As Object

        mydata = mysheet.UsedRange.Address
        myarr = Split(mydata, "$")
        getColumn = mysheet.Range(myarr(3) & "1").Column

    End Function
    Public Function getRow(ByVal mysheet As Excel.Worksheet) As Long
        Dim mydata As String
        mydata = mysheet.UsedRange.Address
        getRow = CLng(Mid(mydata, InStrRev(mydata, "$")))
    End Function
    Private Function getVendorList() As String
        Return vendorlist
        'Dim myStr() As String
        'Dim rsTmp As ADODB.Recordset
        'Dim i As Integer
        'Sqlstr = "select vendorname from orderlinemembers ol" & _
        '         " left join vendor v on v.vendorcode = ol.customercode" & _
        '         "  where ol.orderlineid = 15"
        'rsTmp = oCon.Execute(Sqlstr)
        'getVendorList = ""
        'i = 0
        'While Not rsTmp.EOF
        'ReDim Preserve myStr(i) As String

        '    myStr(i) = Trim(rsTmp!vendorname)
        '    i = i + 1
        '    rsTmp.MoveNext()
        'End While
        'getVendorList = Join(myStr, ",")
        'Erase myStr
        'rsTmp = Nothing
    End Function

    Private Sub CreatePivotTable2(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "IPLT!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C27", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        osheet.PivotTables("PivotTable2").calculatedfields.add("targetiplt", "=0.95", True)
        With osheet.PivotTables("PivotTable2")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With


        osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("AA9").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable2").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable2").Pivotfields("postingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Columns("AA:AA").numberformat = "mmm-yy"

        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("postingdate - shipdate"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("iplt<=5"), " iplt<=5", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable2").PivotFields(" Nbr of days").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable2").PivotFields(" targetiplt").numberformat = "0.0%"
        osheet.PivotTables("PivotTable2").PivotFields(" iplt<=5").numberformat = "0%"

        oWb.Names.Add("myrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C27,0,0,COUNTA('PivotTables'!C28)-4,1)")
        oWb.Names.Add("iplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C28,0,0,COUNTA('PivotTables'!C28)-4,1)")
        oWb.Names.Add("targetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C29,0,0,COUNTA('PivotTables'!C28)-4,1)")
        oWb.Names.Add("ipltle5", RefersToR1C1:="=OFFSET('PivotTables'!R8C30,0,0,COUNTA('PivotTables'!C28)-4,1)")

        osheet.Cells.EntireColumn.AutoFit()
        osheet.Name = "PivotTables"

        osheet.Cells.EntireColumn.AutoFit()
    End Sub
    Private Sub CreatePivotTable3(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "IPLT Miro!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C35", "PivotTable3", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable3")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
        End With

        osheet.PivotTables("PivotTable3").calculatedfields.add("targetiplt", "=0.95", True)

        osheet.PivotTables("PivotTable3").Pivotfields("invoicepostingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("AI9").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable3").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

        For Each item As Object In osheet.PivotTables("PivotTable3").pivotfields("Years").pivotitems
            Dim obj = DirectCast(item, Excel.PivotItem)
            If obj.Value.ToString <> startdate.Year.ToString Then
                obj.Visible = False
            End If
        Next

        osheet.PivotTables("PivotTable3").Pivotfields("invoicepostingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable3").Pivotfields("miropostingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Columns("AI:AI").numberformat = "mmm-yy"

        'osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("miropostingdate - shipdate"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("LeadTime"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable3").AddDataField(osheet.PivotTables("PivotTable3").PivotFields("miro iplt<=5"), " miro iplt<=5", Excel.XlConsolidationFunction.xlAverage)
        osheet.PivotTables("PivotTable3").PivotFields(" Nbr of days").NumberFormat = "0.00"
        osheet.PivotTables("PivotTable3").PivotFields(" targetiplt").numberformat = "0.0%"
        osheet.PivotTables("PivotTable3").PivotFields(" miro iplt<=5").numberformat = "0%"

        oWb.Names.Add("miromyrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C35,0,0,COUNTA('PivotTables'!C36)-4,1)")
        oWb.Names.Add("miroiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C36,0,0,COUNTA('PivotTables'!C36)-4,1)")
        oWb.Names.Add("mirotargetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C37,0,0,COUNTA('PivotTables'!C36)-4,1)")
        oWb.Names.Add("miroipltle5", RefersToR1C1:="=OFFSET('PivotTables'!R8C38,0,0,COUNTA('PivotTables'!C36)-4,1)")

        osheet.Cells.EntireColumn.AutoFit()
    End Sub
    'Private Sub CreatePivotTable(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal startdate As Date)
    '    Dim osheet As Excel.Worksheet

    '    oWb.Worksheets(isheet).select()
    '    osheet = oWb.Worksheets(isheet)
    '    oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "WOR!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)

    '    With osheet.PivotTables("PivotTable1")
    '        .ingriddropzones = True
    '        .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
    '    End With


    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.Range("A9").Group(True, True, Periods:={False, False, False, False, True, False, True})
    '    osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

    '    For Each item As Object In osheet.PivotTables("PivotTable1").pivotfields("Years").pivotitems
    '        Dim obj = DirectCast(item, Excel.PivotItem)
    '        If obj.Value.ToString <> startdate.Year.ToString Then
    '            obj.Visible = False
    '        End If
    '    Next
    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.PivotTables("PivotTable1").Pivotfields("shipdaterange").numberformat = "MMM-yy"

    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetsasl", "=0.85", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetfcr", "=0.85", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetcltslt", "=28", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetfsl", "=0.90", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetoplt", "=0.90", True)
    '    osheet.PivotTables("PivotTable1").calculatedfields.add("targetotd", "=0.90", True)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("ordertype"), " count", Excel.XlConsolidationFunction.xlCount)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl"), " sasl", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sasl<=7"), " sasl<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetsasl"), " targetsasl", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fcr"), " fcr", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fcr<=7"), " fcr<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetfcr"), " targetfcr", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("clt"), " clt", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("slt"), " slt", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetcltslt"), " targetcltslt", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fsl"), " fsl", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("fsl<=7"), " fsl<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetfsl"), " targetfsl", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("pireply"), " pireply", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetoplt"), " targetoplt", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("0-5days"), " 0-5days", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("not conf"), " not conf", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("OTD<=7"), " OTD<=7", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable1").addDatafield(osheet.PivotTables("PivotTable1").PivotFields("targetotd"), " targetotd", Excel.XlConsolidationFunction.xlAverage)

    '    osheet.PivotTables("PivotTable1").PivotFields(" sasl").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" sasl<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetsasl").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" fcr").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" fcr<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetfcr").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" clt").NumberFormat = "0.0"
    '    osheet.PivotTables("PivotTable1").PivotFields(" slt").numberformat = "0.0"
    '    osheet.PivotTables("PivotTable1").PivotFields("targetcltslt").NumberFormat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" fsl").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" fsl<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetfsl").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" pireply").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable1").PivotFields(" 0-5days").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetoplt").numberformat = "0%"
    '    osheet.PivotTables("PivotTable1").PivotFields("not conf").numberformat = "0%"

    '    osheet.PivotTables("PivotTable1").PivotFields(" OTD<=7").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable1").PivotFields(" targetotd").numberformat = "0%"

    '    oWb.Names.Add("MonthRange", RefersToR1C1:="=OFFSET('PivotTables'!R8C1,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("sasl", RefersToR1C1:="=OFFSET('PivotTables'!R8C3,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("saslle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C4,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetsasl", RefersToR1C1:="=OFFSET('PivotTables'!R8C5,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fcr", RefersToR1C1:="=OFFSET('PivotTables'!R8C6,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fcrle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C7,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetfcr", RefersToR1C1:="=OFFSET('PivotTables'!R8C8,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("clt", RefersToR1C1:="=OFFSET('PivotTables'!R8C9,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("slt", RefersToR1C1:="=OFFSET('PivotTables'!R8C10,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetcltslt", RefersToR1C1:="=OFFSET('PivotTables'!R8C11,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fsl", RefersToR1C1:="=OFFSET('PivotTables'!R8C12,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("fslle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C13,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetfsl", RefersToR1C1:="=OFFSET('PivotTables'!R8C14,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("pireply", RefersToR1C1:="=OFFSET('PivotTables'!R8C15,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetoplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C16,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("lefive", RefersToR1C1:="=OFFSET('PivotTables'!R8C17,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("notconf", RefersToR1C1:="=OFFSET('PivotTables'!R8C18,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("otdle7", RefersToR1C1:="=OFFSET('PivotTables'!R8C19,0,0,COUNTA('PivotTables'!C1)-3,1)")
    '    oWb.Names.Add("targetotd", RefersToR1C1:="=OFFSET('PivotTables'!R8C20,0,0,COUNTA('PivotTables'!C1)-3,1)")

    '    oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "IPLT!ExternalData_1").CreatePivotTable(osheet.Name & "!R6C27", "PivotTable2", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
    '    osheet.PivotTables("PivotTable2").calculatedfields.add("targetiplt", "=0.95", True)
    '    With osheet.PivotTables("PivotTable2")
    '        .ingriddropzones = True
    '        .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
    '    End With


    '    osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.Range("AA9").Group(True, True, Periods:={False, False, False, False, True, False, True})
    '    osheet.PivotTables("PivotTable2").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlPageField

    '    For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
    '        Dim obj = DirectCast(item, Excel.PivotItem)
    '        If obj.Value.ToString <> startdate.Year.ToString Then
    '            obj.Visible = False
    '        End If
    '    Next

    '    osheet.PivotTables("PivotTable2").Pivotfields("postingdate").orientation = Excel.XlPivotFieldOrientation.xlHidden
    '    osheet.PivotTables("PivotTable2").Pivotfields("postingdaterange").orientation = Excel.XlPivotFieldOrientation.xlRowField
    '    osheet.Columns("AA:AA").numberformat = "mmm-yy"

    '    osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("postingdate - shipdate"), " Nbr of days", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("targetiplt"), " targetiplt", Excel.XlConsolidationFunction.xlSum)
    '    osheet.PivotTables("PivotTable2").AddDataField(osheet.PivotTables("PivotTable2").PivotFields("iplt<=5"), " iplt<=5", Excel.XlConsolidationFunction.xlAverage)
    '    osheet.PivotTables("PivotTable2").PivotFields(" Nbr of days").NumberFormat = "0.00"
    '    osheet.PivotTables("PivotTable2").PivotFields(" targetiplt").numberformat = "0.0%"
    '    osheet.PivotTables("PivotTable2").PivotFields(" iplt<=5").numberformat = "0%"

    '    oWb.Names.Add("myrange", RefersToR1C1:="=OFFSET('PivotTables'!R8C27,0,0,COUNTA('PivotTables'!C28)-4,1)")
    '    oWb.Names.Add("iplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C28,0,0,COUNTA('PivotTables'!C28)-4,1)")
    '    oWb.Names.Add("targetiplt", RefersToR1C1:="=OFFSET('PivotTables'!R8C29,0,0,COUNTA('PivotTables'!C28)-4,1)")
    '    oWb.Names.Add("ipltle5", RefersToR1C1:="=OFFSET('PivotTables'!R8C30,0,0,COUNTA('PivotTables'!C28)-4,1)")

    '    osheet.Cells.EntireColumn.AutoFit()
    '    osheet.Name = "PivotTables"


    'End Sub



    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.

    End Sub



    'Private Sub RadioButton1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged, RadioButton3.CheckedChanged
    '    Dim obj = DirectCast(sender, RadioButton)
    '    Select Case obj.Text
    '        Case "Market"
    '            CheckedListBox1.Enabled = True
    '            CheckedListBox2.Enabled = False
    '            CheckedListBox3.Enabled = False
    '        Case "Supplier"
    '            CheckedListBox1.Enabled = False
    '            CheckedListBox2.Enabled = True
    '            CheckedListBox3.Enabled = False
    '        Case "SAO"
    '            CheckedListBox1.Enabled = False
    '            CheckedListBox2.Enabled = False
    '            CheckedListBox3.Enabled = True
    '    End Select
    'End Sub


    'Private Sub CheckedListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckedListBox1.SelectedIndexChanged, CheckedListBox2.SelectedIndexChanged, CheckedListBox3.SelectedIndexChanged
    '    CheckedListBox_SelectedIndexChanged(sender, e)
    'End Sub


    'Private Function getselected(ByVal sender As Object) As String
    '    Dim myobj = DirectCast(sender, CheckedListBox)


    '    Dim sb As New StringBuilder
    '    Return sb.ToString
    'End Function

    Private Sub CreateChart1(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal sr As ReportAVPIExt)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

        osheet.Cells(11, 1) = "Logistic Indicators for Components " & sr.startdate.Year
        'osheet.Cells(12, 1) = sr.dr.Item(0)

        Dim ochart As New Excel.Chart

        ochart = osheet.ChartObjects("SASL").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!sasl"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetsasl"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!saslle7"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("FCR").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!clt"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetfcr"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!fcrle7"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("CLT").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!clt"
        ochart.SeriesCollection(2).Values = "='PivotTables'!slt"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!targetcltslt"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("FSL").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!fsl"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetfsl"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!fslle7"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("OPLT").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!pireply"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetoplt"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!lefive"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(4).Values = "='PivotTables'!notconf"
        ochart.SeriesCollection(4).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("OTD").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!otdle7"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetotd"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"

        ochart = osheet.ChartObjects("SSL").Chart
        ochart.SeriesCollection(1).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(1).Values = "='PivotTables'!ordercount"
        ochart.SeriesCollection(2).Values = "='PivotTables'!targetssl"
        ochart.SeriesCollection(2).XValues = "='PivotTables'!MonthRange"
        ochart.SeriesCollection(3).Values = "='PivotTables'!ssl"
        ochart.SeriesCollection(3).XValues = "='PivotTables'!MonthRange"

    End Sub
    'Private Sub CreateChart2(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal sr As Components.ScoreboardReport)
    '    Dim osheet As Excel.Worksheet

    '    oWb.Worksheets(isheet).select()
    '    osheet = oWb.Worksheets(isheet)

    '    osheet.Cells(11, 1) = "Logistic Indicators for Components " & sr.startdate.Year
    '    osheet.Cells(12, 1) = sr.dr.Item(0)

    '    Dim ochart As New Excel.Chart

    '    ochart = osheet.ChartObjects("IPLT").Chart
    '    ochart.SeriesCollection(1).XValues = "='PivotTables'!MyRange"
    '    ochart.SeriesCollection(1).Values = "='PivotTables'!iplt"
    '    ochart.SeriesCollection(2).Values = "='PivotTables'!targetiplt"
    '    ochart.SeriesCollection(2).XValues = "='PivotTables'!myRange"
    '    ochart.SeriesCollection(3).Values = "='PivotTables'!ipltle5"
    '    ochart.SeriesCollection(3).XValues = "='PivotTables'!myRange"

    'End Sub
    'Private Sub CreateChart3(ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal sr As Components.ScoreboardReport)
    '    Dim osheet As Excel.Worksheet

    '    oWb.Worksheets(isheet).select()
    '    osheet = oWb.Worksheets(isheet)

    '    osheet.Cells(11, 1) = "Logistic Indicators for Components " & sr.startdate.Year
    '    osheet.Cells(12, 1) = sr.dr.Item(0)

    '    Dim ochart As New Excel.Chart

    '    ochart = osheet.ChartObjects("IPLT Miro").Chart
    '    ochart.SeriesCollection(1).XValues = "='PivotTables'!miroMyRange"
    '    ochart.SeriesCollection(1).Values = "='PivotTables'!miroiplt"
    '    ochart.SeriesCollection(2).Values = "='PivotTables'!mirotargetiplt"
    '    ochart.SeriesCollection(2).XValues = "='PivotTables'!miromyRange"
    '    ochart.SeriesCollection(3).Values = "='PivotTables'!miroipltle5"
    '    ochart.SeriesCollection(3).XValues = "='PivotTables'!miromyRange"

    'End Sub

    'Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged, CheckBox2.CheckedChanged
    '    exclude = CheckBox1.Checked
    '    excludeComments = CheckBox2.Checked
    'End Sub

    'Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
    '    Me.Close()
    'End Sub



    'Private Sub CheckBox2_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox2.CheckedChanged
    '    Dim obj = CType(sender, CheckBox)
    '    DateTimePicker3.Enabled = obj.Checked
    '    DateTimePicker4.Enabled = obj.Checked
    'End Sub


End Class

Public Class ReportAVPIExt
    Public Property filename As String
    Public Property errormsg As String
    Public Property startdate As Date
    Public Property enddate As Date
    Public Property oxl As Excel.Application
    Public Property owb As Excel.Workbook
    Public Property osheet As Excel.Worksheet
    Public Property isheet As Integer
End Class