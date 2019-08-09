Imports System.Text
Imports PriceCMMFExt.PublicClass

Public Class UpdatePriceChangeReason
    Private Filename As String
    Public ErrorMessage As String
    Dim ReasonDict As New Dictionary(Of String, Integer)
    Dim Parent As Object

    Dim DataReaderCallback As FormatReportDelegate = AddressOf DataReader

    Public Sub New(ByVal filename As String, ByRef Parent As Object)
        Me.Filename = filename
        PopulateDict()
        Me.Parent = Parent
    End Sub

    Public Function Run() As Boolean
        Dim sbError As New StringBuilder
        Dim myret As Boolean = False
        Dim UpdateReasonSB As New StringBuilder
        Dim ReasonErrorDict As New Dictionary(Of String, String)

        Dim sw As New Stopwatch
        Dim DS As New DataSet
        Dim mystr As New StringBuilder
        Dim myInsert As New System.Text.StringBuilder
        Dim myrecord() As String
        Dim mylist As New List(Of String())
        ProgressReport(6, "Marque")
        ProgressReport(2, "Build records..")
        sw.Start()
        ' Try
        Using objTFParser = New FileIO.TextFieldParser(Filename)
            With objTFParser
                .TextFieldType = FileIO.FieldType.Delimited
                .SetDelimiters(Chr(9))
                .HasFieldsEnclosedInQuotes = True
                Dim count As Long = 0

                Do Until .EndOfData
                    myrecord = .ReadFields
                    If count > 0 Then
                        mylist.Add(myrecord)
                    End If
                    count += 1
                Loop

                For i = 0 To mylist.Count - 1                    
                    'ProgressReport(2, i + 1 & "/" & mylist.Count)
                    myrecord = mylist(i)
                    If ReasonDict.ContainsKey(myrecord(1)) Then
                        If UpdateReasonSB.Length > 0 Then
                            UpdateReasonSB.Append(",")
                        End If
                        UpdateReasonSB.Append(String.Format("['{0}'::character varying,'{1}'::character varying]", myrecord(0), ReasonDict(myrecord(1))))
                    Else
                        If Not ReasonErrorDict.ContainsKey(myrecord(1)) Then
                            sbError.Append(String.Format("Reason ""{0}"" is not registered.{1}", myrecord(1), vbCrLf))
                            ReasonErrorDict.Add(myrecord(1), myrecord(1))
                        End If
                    End If
                Next


            End With
        End Using
        'update record

        If UpdateReasonSB.Length > 0 Then
            ProgressReport(2, "Update Reason")
            'cmmf,rir,itemid,materialdesc,vendorcode,comfam,loadingcode,brandid
            Dim sqlstr = "update pricechangehd set reasonid= foo.reasonid::integer from (select * from array_to_set2(Array[" & UpdateReasonSB.ToString &
                     "]) as tb (id character varying,reasonid character varying))foo where pricechangehdid = foo.id::bigint;"
            Dim ra As Long
            If Not DbAdapter1.ExecuteNonQuery(sqlstr, ra, ErrorMessage) Then
                ProgressReport(2, ErrorMessage)
                myret = False
                Err.Raise(513, Description:=ErrorMessage & " ::Update Reason Price")
            End If
        End If
        
        sw.Stop()
        ProgressReport(2, String.Format("Done. Elapsed Time: {0}:{1}.{2}", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))

        If sbError.Length = 0 Then
            myret = True
        Else
            ErrorMessage = sbError.ToString
        End If
        ProgressReport(5, "Continuous")
        Return myret
    End Function


    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Parent.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try

        Else
            Select Case id
                Case 2
                    Parent.ToolStripStatusLabel1.Text = message
                Case 3
                    Parent.ToolStripStatusLabel2.Text = Trim(message)
                Case 4
                    Parent.close()
                Case 5
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select

        End If

    End Sub

    Sub DataReader(ByRef sender As Object, ByRef e As EventArgs)
        ReasonDict.Clear()
        Dim dr = DirectCast(sender, Npgsql.NpgsqlDataReader)
        While dr.Read
            ReasonDict.Add(dr.Item(1), dr.Item(0))
        End While
    End Sub

    Private Sub PopulateDict()
        Dim sqlstr = "select id,reasonname from pricechangereason order by id"
        DbAdapter1.DataReaderCallback = Me.DataReaderCallback
        If Not DbAdapter1.ExecuteReader(sqlstr, ErrorMessage) Then
            MessageBox.Show(ErrorMessage)
        End If
    End Sub

End Class
