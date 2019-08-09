Imports System.Threading
Imports System.Text
Imports PriceCMMFExt.PublicClass
Imports PriceCMMFExt.SharedClass
Public Class FormDocumentHeader
    Dim limit As String = " limit 1"
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim bsheader As New BindingSource
    Dim bshistory As New BindingSource
    Dim DS As DataSet
    Dim sb As New StringBuilder
    Dim creator As String
    Dim myuser As String = String.Empty
    Dim headerid As Long = 0
    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        If Not myThread.IsAlive Then
            'If creator = "" Then
            '    MessageBox.Show("You are not allowed to create new Task.")
            '    Exit Sub
            'End If

            'Dim myrow As DataRowView = bsheader.AddNew
            'myrow.Row.Item("creatorname") = creator.Trim
            'myrow.Row.Item("creator") = myuser
            'myrow.Row.Item("submitdate") = Today.Date
            ''myrow.Row.Item("negotiateddate") = Today.Date
            'myrow.Row.Item("pricetype") = "FOB"
            'myrow.Row.Item("status") = 2
            'DS.Tables(0).Rows.Add(myrow.Row)

            Dim myform = New FormDocumentVendor()
            'myform.ComboBox1.SelectedIndex = 0
            If myform.ShowDialog = DialogResult.OK Then
                'MessageBox.Show("Add New One")
                'bsheader.CancelEdit()
                'Dim dr As DataRow = myform.bsheader.Current
                'DS.Tables(0).Rows.Add(dr)
                'DataGridView1.Invalidate()
                DS.Merge(myform.DS.Tables(0))
                'DataGridView1.Invalidate()
                'loaddata()
                'bsheader.RemoveCurrent()
            Else
                'bsheader.EndEdit()

            End If

        Else
            MessageBox.Show("Still loading... Please wait.")
        End If


    End Sub

  

    Private Sub loaddata()
        myuser = HelperClass1.UserId.ToLower
        'MessageBox.Show(myuser)
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Sub DoWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")
        'ProgressReport(4, "InitData")
        '2 Dataset 
        '1 contains All tx except Completed
        'the other only contains Completed

        DS = New DataSet
        Dim mymessage As String = String.Empty
        sb.Clear()
        'Admin checking first
        'sb.Append("select * from pricechangehd ph where (ph.creator = '" & HelperClass1.UserId & "')")

        myuser = HelperClass1.UserId.ToLower
        'myuser = "as\dlie"
        'myuser = "as\elai"
        'myuser = "as\rlo"
        'myuser = "AS\afok".ToLower
        'myuser = "as\weho"
        'myuser = "AS\shxu".ToLower
        'myuser = "as\jdai"
        'myuser = "AS\SCHAN".ToLower
        'sb.Append("select h.* ,o.officersebname::text as username,o2.officersebname as validatorname,o3.officersebname as cc1name ,o4.officersebname as cc2name,o5.officersebname as cc3name,o6.officersebname as cc4name" &
        '          " from doc.header h" &
        '          " left join officerseb o on o.userid = h.userid" &
        '          " left join officerseb o2 on o2.userid = h.validator" &
        '          " left join officerseb o3 on o3.userid = h.cc1" &
        '          " left join officerseb o4 on o4.userid = h.cc2" &
        '          " left join officerseb o5 on o5.userid = h.cc3" &
        '          " left join officerseb o6 on o6.userid = h.cc4")
        sb.Append("select distinct h.*,doc.sp_getsuppliers(h.id) as suppliername,doc.sp_getdoctypename(h.id) as doctypename ,o.officersebname::text as username,o2.officersebname::text as validatorname,o3.officersebname::text as cc1name ,o4.officersebname::text as cc2name,o5.officersebname::text as cc3name,o6.officersebname::text as cc4name from doc.header h " &
                  " left join officerseb o on o.userid = h.userid" &
                  " left join officerseb o2 on o2.userid = h.validator" &
                  " left join officerseb o3 on o3.userid = h.cc1" &
                  " left join officerseb o4 on o4.userid = h.cc2" &
                  " left join officerseb o5 on o5.userid = h.cc3" &
                  " left join officerseb o6 on o6.userid = h.cc4" &
                  " where h.userid = '" & myuser & "' order by h.id desc;")
        sb.Append("select * from officerseb o  where lower(o.userid) = '" & myuser & "' limit 1;")

        If DbAdapter1.TbgetDataSet(sb.ToString, DS, mymessage) Then
            Try

                DS.Tables(0).TableName = "Header"


            Catch ex As Exception
                ProgressReport(1, "Loading Data. Error::" & ex.Message)
                ProgressReport(5, "Continuous")
                Exit Sub
            End Try
            ProgressReport(4, "InitData")
        Else
            ProgressReport(1, "Loading Data. Error::" & mymessage)
            ProgressReport(5, "Continuous")
            Exit Sub
        End If
        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Try


                Select Case id
                    Case 1
                        ToolStripStatusLabel1.Text = message
                    Case 2
                        ToolStripStatusLabel1.Text = message
                    Case 4
                        Try
                            bsheader = New BindingSource

                            Dim pk(0) As DataColumn
                            pk(0) = DS.Tables(0).Columns("id")
                            DS.Tables(0).PrimaryKey = pk
                            DS.Tables(0).Columns(0).AutoIncrement = True
                            DS.Tables(0).Columns(0).AutoIncrementSeed = 0
                            DS.Tables(0).Columns(0).AutoIncrementStep = -1
                            DS.Tables(0).TableName = "Header"

                            

                            bsheader.DataSource = DS.Tables(0)

                            DataGridView1.AutoGenerateColumns = False
                            DataGridView1.DataSource = bsheader
                        Catch ex As Exception
                            message = ex.Message
                        End Try

                    Case 5
                        ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                    Case 6
                        ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                End Select
            Catch ex As Exception

            End Try
        End If

    End Sub

    Private Sub FormMyTask_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        loaddata()
        'ComboBox1.SelectedIndex = 1
    End Sub

    Private Sub DataGridView1_CellBeginEdit(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellCancelEventArgs) Handles DataGridView1.CellBeginEdit
        'MessageBox.Show(DataGridView1.Columns(e.ColumnIndex).HeaderText)
    End Sub

    Private Sub DataGridView1_EditingControlShowing(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewEditingControlShowingEventArgs) Handles DataGridView1.EditingControlShowing
        Dim tb As DataGridViewTextBoxEditingControl = DirectCast(e.Control, DataGridViewTextBoxEditingControl)
        RemoveHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
        AddHandler (tb.KeyDown), AddressOf datagridviewTextBox_Keypdown
    End Sub

    Private Sub datagridviewTextBox_Keypdown(ByVal sender As Object, ByVal e As KeyEventArgs)
        If e.KeyValue = 112 Then 'F1 
            MessageBox.Show("Help")
        End If
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.Text = Me.Text & "-" & HelperClass1.UserId
    End Sub

    'Private Sub DataGridView1_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellDoubleClick
    '    If Not myThread.IsAlive Then
    '        Dim myrow As DataRowView = bsheader.Current
    '        Dim myform = New FormPriceChange(bsheader, DS, False)
    '        Select Case myrow.Row.Item("status")
    '            Case 2, 4
    '                myform.ToolStripButton2.Visible = False
    '                myform.ToolStripButton7.Visible = False
    '            Case 3
    '                myform.ToolStripButton4.Visible = False
    '                myform.ToolStripButton5.Visible = False
    '            Case 5
    '                myform.ToolStripButton2.Visible = False
    '                myform.ToolStripButton4.Visible = False
    '                myform.ToolStripButton5.Visible = False
    '        End Select

    '        If Not myform.ShowDialog = DialogResult.OK Then

    '            bsheader.CancelEdit()
    '        Else
    '            bsheader.EndEdit()
    '        End If
    '        loaddata()
    '    End If

    'End Sub

   

    Private Sub DataGridView1_DataError(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        If Not myThread.IsAlive Then
            loaddata()
        End If
    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        If Not IsNothing(bsheader.Current) Then
            If Not myThread.IsAlive Then
                Dim drv As DataRowView = bsheader.Current
                Dim myform = New FormDocumentVendor(drv.Row.Item("id"))
                'myform.ComboBox1.SelectedIndex = 0
                If myform.ShowDialog = DialogResult.OK Then
                    'MessageBox.Show("Add New One")
                    'bsheader.CancelEdit()
                    'Dim dr As DataRow = myform.bsheader.Current
                    'DS.Tables(0).Rows.Add(dr)
                    'DataGridView1.Invalidate()
                    'DS.Merge(myform.DS.Tables(0))
                    'DataGridView1.Invalidate()
                    loaddata()
                    'bsheader.RemoveCurrent()
                Else
                    'bsheader.EndEdit()

                End If

            Else
                MessageBox.Show("Still loading... Please wait.")
            End If
        End If
    End Sub
End Class