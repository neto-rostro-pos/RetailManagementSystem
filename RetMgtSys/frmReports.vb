Imports ggcAppDriver
Imports ggcRMSReports

Public Class frmReports
    Private pnActiveRow As Integer = -1
    Private p_oAppDriver As GRider
    Private ReportsArrayList As List(Of Reports) = New List(Of Reports)()

    Public WriteOnly Property Grider() As GRider
        Set(ByVal value As GRider)
            p_oAppDriver = value
        End Set
    End Property

    Private Sub initGrid()
        With DataGridView1
            .RowCount = 0

            'Set No of Columns
            .ColumnCount = 2

            'Set Column Headers
            .Columns(0).HeaderText = "No"
            .Columns(1).HeaderText = "Report Name"


            'Set Column Sizes
            .Columns(0).Width = 50
            .Columns(1).Width = 280

            .Columns(0).SortMode = DataGridViewColumnSortMode.NotSortable
            .Columns(1).SortMode = DataGridViewColumnSortMode.NotSortable

            .Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
            .Columns(1).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft

            .Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
            .Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

            'Set No of Rows
            .RowCount = 1

            pnActiveRow = 0
        End With
    End Sub

    Private Sub frmReports_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        Call initGrid()
        Call AddInfo()
        Call loadInfo()
    End Sub


    Private Sub AddInfo()
        With ReportsArrayList
            .Add(New Reports With {.Reports = "Accumulated Grand Total Report", .Number = "1"})
            .Add(New Reports With {.Reports = "Cancelled Invoice Report", .Number = "2"})
            .Add(New Reports With {.Reports = "Discounted Sales Report", .Number = "3"})
            .Add(New Reports With {.Reports = "E-Journal", .Number = "4"})
            .Add(New Reports With {.Reports = "Events Log report", .Number = "5"})
            .Add(New Reports With {.Reports = "Sales Report", .Number = "6"})
            .Add(New Reports With {.Reports = "Void Transaction Report", .Number = "7"})
            .Add(New Reports With {.Reports = "Complementary Report", .Number = "8"})
            .Add(New Reports With {.Reports = "Item List Report", .Number = "9"})
            '.Add(New Reports With {.Reports = "Charge Report", .Number = "11"})
            '.Add(New Reports With {.Reports = "Return Report", .Number = "12"})
            .Add(New Reports With {.Reports = "Bir Sales Summary Report", .Number = "10"})
            '.Add(New Reports With {.Reports = "Sales Summary Report", .Number = "8"})
            .Add(New Reports With {.Reports = "Daily Sales Report", .Number = "11"})
            .Add(New Reports With {.Reports = "Sales Ranking Report", .Number = "12"})
            '.Add(New Reports With {.Reports = "Statement of Account Detailed Report", .Number = "13"})
            '.Add(New Reports With {.Reports = "Statement of Account Summarized Report", .Number = "14"})
            .Add(New Reports With {.Reports = "Z-Reading Summary", .Number = "13"})
            .Add(New Reports With {.Reports = "Charge Invoice Report", .Number = "14"})

        End With
    End Sub

    Private Sub loadInfo()
        With DataGridView1

            .RowCount = ReportsArrayList.Count

            For lnCtr = 0 To .RowCount - 1
                .Item(0, lnCtr).Value = ReportsArrayList.Item(lnCtr).Number
                .Item(1, lnCtr).Value = ReportsArrayList.Item(lnCtr).Reports
            Next

        End With
    End Sub

    Class Reports
        Private Property psNumber As Integer
        Private Property psReportName As String

        Public Property Reports() As String
            Get
                Return psReportName
            End Get
            Set(ByVal value As String)
                psReportName = value
            End Set
        End Property

        Public Property Number() As Integer
            Get
                Return psNumber
            End Get
            Set(ByVal value As Integer)
                psNumber = value
            End Set
        End Property

    End Class

    Private Sub cmdButton02_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdButton02.Click
        Me.Close()
    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        pnActiveRow = DataGridView1.CurrentCell.RowIndex + 1
    End Sub

    Private Sub cmdButton00_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdButton00.Click
        Select Case pnActiveRow
            Case 1 'grand total  ok
                Dim loRpt As ggcRMSReports.clsGranTotalReport
                loRpt = New ggcRMSReports.clsGranTotalReport(p_oAppDriver)

                Call loRpt.ReportTrans()
           
            Case 2 'Cancelled Receipt ok
                Dim loRpt As ggcRMSReports.clsCancelledReport
                loRpt = New ggcRMSReports.clsCancelledReport(p_oAppDriver, p_sPOSNo)

                Call loRpt.ReportTrans()
            Case 3 ' Discounted Sales report
                Dim loRpt As ggcRMSReports.clsDiscountedSales
                loRpt = New ggcRMSReports.clsDiscountedSales(p_oAppDriver, p_sPOSNo)

                Call loRpt.ReportTrans()
            Case 4 ' EJournal ok
                Dim loForm As New frmTerminal
                loForm.GRider = p_oAppDriver
                loForm.Show()
            Case 5 ' Event Logs
                Dim loRpt As ggcRMSReports.clsEventLogs
                loRpt = New ggcRMSReports.clsEventLogs(p_oAppDriver)

                Call loRpt.ReportTrans()
            Case 6 ' Sales Report ok'
                Dim loRpt As ggcRMSReports.clsSalesReport
                loRpt = New ggcRMSReports.clsSalesReport(p_oAppDriver, p_sPOSNo)

                Call loRpt.ReportTrans()
            Case 7 'Void Items ok
                Dim loRpt As ggcRMSReports.clsVoidItems
                loRpt = New ggcRMSReports.clsVoidItems(p_oAppDriver, p_sPOSNo)

                Call loRpt.ReportTrans()
            Case 8 ' Complementary Report ok
                Dim loRpt As ggcRMSReports.clsComplementary
                loRpt = New ggcRMSReports.clsComplementary(p_oAppDriver, p_sPOSNo)

                Call loRpt.ReportTrans()
                'Case 11 ' Charge Report ok
                '    'Dim loRpt As ggcRMSReports.clsReturns
                '    'loRpt = New ggcRMSReports.clsReturns(p_oAppDriver, p_sPOSNo)

                '    'Call loRpt.ReportTrans()
            Case 9 ' Product List ok
                Dim loRpt As ggcRMSReports.clsProductList
                loRpt = New ggcRMSReports.clsProductList(p_oAppDriver)

                Call loRpt.ReportTrans()
            'Case 12 'return cancel/replacement report
            '    Dim loRpt As ggcRMSReports.clsReturns
            '    loRpt = New ggcRMSReports.clsReturns(p_oAppDriver, p_sPOSNo)

            '    Call loRpt.ReportTrans()
            Case 10 'BIR Sales Summary ok
                Dim loRpt As ggcRMSReports.clsBirSummary
                loRpt = New ggcRMSReports.clsBirSummary(p_oAppDriver, p_sPOSNo, p_sSerial, p_sPermit, p_sVATReg)

                Call loRpt.ReportTrans()
                'Case 8 'Sales Summary Report
                '    Dim loRpt As ggcRMSReports.clsSalesSummary
                '    loRpt = New ggcRMSReports.clsSalesSummary(p_oAppDriver, p_sPOSNo, p_sSerial, p_sPermit, p_sVATReg)

                '    Call loRpt.ReportTrans()
            Case 11 ' Daily Sales Report ok'
                Dim loRpt As ggcRMSReports.clsDailySalesReport
                loRpt = New ggcRMSReports.clsDailySalesReport(p_oAppDriver)

                Call loRpt.ReportTrans()
            Case 12 ' Daily Sales Report ok'
                Dim loRpt As ggcRMSReports.clsRankingReport
                loRpt = New ggcRMSReports.clsRankingReport(p_oAppDriver)

                Call loRpt.ReportTrans()
            'Case 13 ' SOA Detailed ok'
            '    Dim loRpt As ggcRMSReports.clsSOADetailed
            '    loRpt = New ggcRMSReports.clsSOADetailed(p_oAppDriver)

            '    Call loRpt.ReportTrans()

            'Case 14 ' SOA Summarized ok'
            '    Dim loRpt As ggcRMSReports.clsSOASummarized
            '    loRpt = New ggcRMSReports.clsSOASummarized(p_oAppDriver)

            '    Call loRpt.ReportTrans()

            Case 13 ' Z Reading Summarized ok'
                Dim loRpt As ggcRMSReports.clsZReadingSummary
                loRpt = New ggcRMSReports.clsZReadingSummary(p_oAppDriver)

                Call loRpt.ReportTrans()
        End Select
        Me.Close()
    End Sub

End Class