Imports System.Threading

Public Class PriceCMMFExAuto
    Dim myThread As New System.Threading.Thread(AddressOf doWork)
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)

    Private Sub PriceCMMFExAuto_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.WindowState = FormWindowState.Minimized
        LoadMe()
    End Sub


    Private Sub PriceCMMFExAuto_Resize(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Resize
        If Me.WindowState = FormWindowState.Minimized Then
            Me.ShowInTaskbar = False
            Me.Hide()
            NotifyIcon1.Visible = True
        End If
    End Sub


    Private Sub NotifyIcon1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick
        Me.Show()
        Me.WindowState = FormWindowState.Normal
        Me.ShowInTaskbar = True
        NotifyIcon1.Visible = False
    End Sub

    Private Sub LoadMe()

        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(AddressOf doWork)
                myThread.TrySetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Sub doWork()
        Logger.log("--------Start----------")

        'Email For Validator
        Logger.log("Send email to Validator")
        Dim myValidator = New RoleTasks(RoleTasks.Role.Validator)
        If Not myValidator.Execute Then
            Logger.log(myValidator.errormessage)
        End If

        'Create SAP File
        Logger.log("Create New SAP File")
        Dim Export As New ExportSAPClass
        If Not Export.ExportFile() Then
            Logger.log(Export.mymessage)
        End If

        'Validate Detail with SAP ZFA037
        Logger.log("CrossCheck with PriceList - Assign Status Completed")
        Dim mycheck = New ValidateSAPPrice
        If Not mycheck.Validate Then
            Logger.log(mycheck.mymessage)
        End If

        

        Logger.log("Send Email to Creator for Status Rejected and (Validated and PriceType = STD) and Completed Also Validated")
        'Email For Creator Status Rejected and (Validated and PriceType = STD) and Completed Also Validated
        Dim mycreator = New RoleTasks(RoleTasks.Role.Creator)
        If Not mycreator.Execute Then
            Logger.log(mycreator.errormessage)
        End If

        Logger.log("Send Email to CC For status Completed")
        'Email For CC Status Completed
        Dim myCC = New RoleTasks(RoleTasks.Role.CC)
        If Not myCC.Execute Then
            Logger.log(myCC.errormessage)
        End If


        Logger.log("Send Email to WMF For status Completed")
        'Email For WMF Status Completed
        Dim myWMF = New RoleTasks(RoleTasks.Role.WMF)
        If Not myWMF.Execute Then
            Logger.log(myWMF.errormessage)
        End If

        Logger.log("--------End------------")
        ProgressReport(1, "Close Apps")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.Close()
                Case 2

                Case 3

                Case 4

                Case 5

                Case 6

                Case 7

            End Select
        End If
    End Sub

End Class