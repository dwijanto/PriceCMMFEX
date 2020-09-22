Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass

Public Class FormImportAverPrice

    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByRef message As String)
    Private Property FileName As String

    Dim myThreadDelegate As New ThreadStart(AddressOf DoWork)
    Dim myThread As New System.Threading.Thread(myThreadDelegate)

    Private DS As DataSet
    Dim AddPriceSB As StringBuilder
    Dim UpdPriceSB As StringBuilder

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
        Dim mylist As New List(Of CMMFVendorPriceModel)
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
                sb.Append("select myyear,cmmf,vendorcode,qtyshipped,purchasevalue,averprice,averpricefixcurr from cmmfvendorprice;")
                mymessage = String.Empty
                If Not DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
                    ProgressReport(2, mymessage)
                    Exit Sub
                End If

                DS.Tables(0).TableName = "CMMFVendorPrice"
                Dim idx0(2) As DataColumn
                idx0(0) = DS.Tables(0).Columns(0)
                idx0(1) = DS.Tables(0).Columns(1)
                idx0(2) = DS.Tables(0).Columns(2)
                DS.Tables(0).PrimaryKey = idx0

                ProgressReport(2, "Read Text File...")
                Try
                    Do Until .EndOfData
                        myrecord = .ReadFields                        
                        If count >= 1 Then
                            Dim mydata As New CMMFVendorPriceModel With {.myyear = myrecord(0),
                                                                     .cmmf = myrecord(1),
                                                                     .vendorcode = myrecord(2),
                                                                     .qtyshipped = myrecord(3),
                                                                     .purchasevalue = myrecord(4),
                                                                     .averprice = myrecord(5),
                                                                     .averpricefixcurr = myrecord(6)}
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

                AddPriceSB = New StringBuilder
                UpdPriceSB = New StringBuilder
                
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
                        Dim pkey1(2) As Object
                        pkey1(0) = CInt(mydata.myyear)
                        pkey1(1) = CLng(mydata.cmmf)
                        pkey1(2) = CLng(mydata.vendorcode)

                        result = DS.Tables(0).Rows.Find(pkey1)
                        If IsNothing(result) Then
                            'Only Update Data no create
                            Debug.Print("Add")
                        Else
                            'myyear	cmmf	vendorcode	qtyshipped	purchasevalue	averprice	averpricefixcurr
                            If Not result.Item("qtyshipped").ToString = mydata.qtyshipped Then
                                updateFlag = True
                            ElseIf Not result.Item("purchasevalue").ToString = mydata.purchasevalue Then
                                updateFlag = True
                            ElseIf Not result.Item("averprice").ToString = mydata.averprice Then
                                updateFlag = True
                            ElseIf Not result.Item("averpricefixcurr").ToString = mydata.averpricefixcurr Then
                                updateFlag = True
                            End If
                            If updateFlag Then
                                If UpdPriceSB.Length > 0 Then
                                    UpdPriceSB.Append(",")
                                End If
                                UpdPriceSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying,'{2}'::character varying,'{3}'::character varying,'{4}'::character varying,'{5}'::character varying,'{6}'::character varying]", mydata.myyear, mydata.cmmf, mydata.vendorcode, validint(mydata.qtyshipped), validreal(mydata.purchasevalue), validreal(mydata.averprice), validreal(mydata.averpricefixcurr)))
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

            If UpdPriceSB.Length > 0 Then
                ProgressReport(1, "Update Data...")
                'myyear	cmmf	vendorcode	qtyshipped	purchasevalue	averprice	averpricefixcurr
                sqlstr = "update cmmfvendorprice set qtyshipped= foo.qtyshipped1::integer,purchasevalue = foo.purchasevalue1::numeric,averprice=foo.averprice1::numeric,averpricefixcurr = foo.averpricefixcurr1::numeric from (select * from array_to_set7(Array[" & UpdPriceSB.ToString &
                         "]) as tb (myyear1 character varying,cmmf1 character varying,vendorcode1 character varying,qtyshipped1	character varying,purchasevalue1 character varying,averprice1 character varying,averpricefixcurr1 character varying))foo where myyear = foo.myyear1::integer and cmmf = foo.cmmf1::bigint and vendorcode = foo.vendorcode1::bigint;"
                Dim ra As Long
                If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, errmsg) Then
                    ProgressReport(1, errmsg)
                    myret = False
                    Err.Raise(513, Description:=errmsg & " ::Update Data")
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

    Private Sub StatusStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub OpenFileDialog1_FileOk(ByVal sender As System.Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub
End Class

Public Class CMMFVendorPriceModel
    Public Property myyear As String
    Public Property cmmf As String
    Public Property vendorcode As String
    Public Property qtyshipped As String
    Public Property purchasevalue As String
    Public Property averprice As String
    Public Property averpricefixcurr As String
End Class