Imports ggcAppDriver
Imports System.Globalization
Imports System.ComponentModel
Imports System.Text
Imports MySql.Data.MySqlClient
Imports ggcRMSReports

Imports System.IO
Imports ggcReceipt

Public Class mdiMain
    Private p_oDTMaster As DataTable

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub mdiMain_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F12 Then
            If p_oAppDriver.UserLevel = xeUserRights.ENGINEER Then
                If MsgBox("You are about to reset the Retail Management Database." & vbCrLf &
                            "Do you want to continue?", MsgBoxStyle.Question + vbYesNo, "Confirm") = vbYes Then
                    If ResetDatabase() Then
                        MsgBox("Retail Management Database has been reset.", MsgBoxStyle.Information, "Success")
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub mdiMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        p_oAppDriver.MDI = Me
        tslUser.Text = Decrypt(p_oAppDriver.UserName, p_oAppDriver.Signature)
        tslDate.Text = Format(p_oAppDriver.SysDate, "MMMM dd, yyyy")
        RestoreDatabseToolStripMenuItem.Visible = False
    End Sub

    Public Function isFormOpen(ByVal lsForm As String) As Boolean
        For Each frm As Form In Me.MdiChildren
            If frm.Name = lsForm Then
                Return True
            End If
        Next

        Return False
    End Function

    Private Sub TableToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs)
        With frmInventory
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub DiscountCardsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DiscountCardsToolStripMenuItem.Click
        With frmDiscountCard
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub AffiliatesToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AffiliatesToolStripMenuItem.Click
        With frmAffiliates
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub MachineToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MachineToolStripMenuItem.Click
        With frmMachine
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub BanksToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BanksToolStripMenuItem.Click
        With frmBanks
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub BinToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles BinToolStripMenuItem.Click
        With frmBin
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub InventoryTypeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles InventoryTypeToolStripMenuItem.Click
        With frmInventory_Type
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub MeasureToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MeasureToolStripMenuItem.Click
        With frmMeasure
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub SectionToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SectionToolStripMenuItem.Click
        With frmSection
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub SizeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SizeToolStripMenuItem.Click
        With frmSize
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub ProductCategoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ProductCategoryToolStripMenuItem.Click
        With frmProduct_Category
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub InventoryToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles InventoryToolStripMenuItem.Click
        With frmInventory1
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub TermToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TermToolStripMenuItem.Click
        With frmTerm
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub SalesPromoToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SalesPromoToolStripMenuItem.Click
        With frmPromoDiscount
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub UserManagerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UserManagerToolStripMenuItem.Click
        With p_oAppDriver
            If .UserLevel <> xeUserRights.SYSADMIN And .UserLevel <> xeUserRights.ENGINEER Then
                MsgBox("Insuficient User Rights Detected.", MsgBoxStyle.Critical, "Warning")
                Exit Sub
            End If
        End With

        With frmSysUser
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub PromoAddOnsToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles PromoAddOnsToolStripMenuItem.Click
        With frmPromoAddOn
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub TerminalZReadingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TerminalZReadingToolStripMenuItem.Click
        With frmZReading
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    'Please create a environment variables for this...
    'Source of Logic: https://www.codeproject.com/Tips/726407/MYSQL-Database-Backup-Restore
    Private Sub BackupDatabase()
        Dim file1 As String
        Dim path1 As String
        sfd.Filter = "SQL Dump File (*.sql)|*.sql|All files (*.*)|*.*"
        file1 = p_oAppDriver.Database & "-" & DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") + ".sql"
        sfd.FileName = file1
        If sfd.ShowDialog = DialogResult.OK Then
            path1 = Replace(sfd.FileName, file1, "")

            Dim myProcess As New Process()
            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.WorkingDirectory = path1

            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.Start()
            Dim myStreamWriter As StreamWriter = myProcess.StandardInput
            Dim mystreamreader As StreamReader = myProcess.StandardOutput

            myStreamWriter.WriteLine("mysqldump -u" & p_oAppDriver.ServerUser & " --password=" & p_oAppDriver.ServerPassword & " -h" & p_oAppDriver.ServerName & " -B --single-transaction --max_allowed_packet=10M " & p_oAppDriver.Database & " > """ & p_oAppDriver.Database & DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss") & ".sql" & """")
            myStreamWriter.Close()
            myProcess.WaitForExit()
            myProcess.Close()
            MsgBox("Backup Created Successfully!", MsgBoxStyle.Information, "Backup")
        End If
    End Sub

    'Source of Logic: https://www.codeproject.com/Tips/726407/MYSQL-Database-Backup-Restore
    Private Sub RestoreDatabase()
        Dim file As String

        ofd.Filter = "SQL Dump File (*.sql)|*.sql|All files (*.*)|*.*"
        If ofd.ShowDialog = DialogResult.OK Then
            file = ofd.FileName

            Dim myProcess As New Process()
            myProcess.StartInfo.FileName = "cmd.exe"
            myProcess.StartInfo.UseShellExecute = False
            myProcess.StartInfo.WorkingDirectory = ""
            myProcess.StartInfo.RedirectStandardInput = True
            myProcess.StartInfo.RedirectStandardOutput = True
            myProcess.Start()
            Dim myStreamWriter As StreamWriter = myProcess.StandardInput
            Dim mystreamreader As StreamReader = myProcess.StandardOutput
            myStreamWriter.WriteLine("mysql -u" & p_oAppDriver.ServerUser & " --password=" & p_oAppDriver.ServerPassword & " -h" & p_oAppDriver.ServerName & " < " & """" & file & """")
            myStreamWriter.Close()
            myProcess.WaitForExit()
            myProcess.Close()
            MsgBox("Database Restoration Successfully!", MsgBoxStyle.Information, "Restore")
        End If
    End Sub

    Private Function ResetDatabase()
        Dim lsSQL(1, 25) As String

        'On Error GoTo errProc

        lsSQL(0, 0) = "DELETE FROM Cash_Pull_Out;"
        lsSQL(0, 1) = "DELETE FROM Charge_Invoice;"
        lsSQL(0, 2) = "DELETE FROM Check_Payment_Trans;"
        lsSQL(0, 3) = "DELETE FROM Clients;"
        lsSQL(0, 4) = "DELETE FROM Complementary;"
        lsSQL(0, 5) = "DELETE FROM Complementary_Trans;"
        lsSQL(0, 6) = "DELETE FROM Credit_Card_Trans;"
        lsSQL(0, 7) = "DELETE FROM Daily_Summary;"
        lsSQL(0, 8) = "DELETE FROM Discount;"
        lsSQL(0, 9) = "DELETE FROM Gift_Certificate_Trans;"
        lsSQL(0, 10) = "DELETE FROM Order_Split;"
        lsSQL(0, 11) = "DELETE FROM Order_Split_Detail;"
        lsSQL(0, 12) = "DELETE FROM Payment;"
        lsSQL(0, 13) = "DELETE FROM Receipt_Master;"
        lsSQL(0, 14) = "DELETE FROM Return_Detail;"
        lsSQL(0, 15) = "DELETE FROM Return_Master;"
        lsSQL(0, 16) = "DELETE FROM SO_Detail;"
        lsSQL(0, 17) = "DELETE FROM SO_Master;"
        lsSQL(0, 18) = "UPDATE Table_Master SET cStatusxx = '0', dReserved = NULL, nCapacity = 0, nOccupnts = 0;"
        lsSQL(0, 19) = "UPDATE Cash_Reg_Machine SET nSalesTot = 0, nVATSales = 0, nVATAmtxx = 0, nNonVATxx = 0, nZeroRatd = 0, nZReadCtr = 0, sORNoxxxx = '', sTransNox = '', sMasterNo='', sOrderNox='', sBillNmbr= '', nSChrgAmt = 0.00, nEODCtrxx = 0;"
        lsSQL(0, 20) = "DELETE FROM Daily_Summary;"
        lsSQL(0, 21) = "DELETE FROM Event_Master;"
        lsSQL(0, 22) = "DELETE FROM Charge_Invoice_Collection_Master;"
        lsSQL(0, 23) = "DELETE FROM Charge_Invoice_Collection_Detail;"
        lsSQL(0, 24) = "DELETE FROM Discount_Detail;"
        lsSQL(0, 25) = "DELETE FROM Charge_Invoice;"

        lsSQL(1, 0) = "Cash_Pull_Out"
        lsSQL(1, 1) = "Charge_Invoice"
        lsSQL(1, 2) = "Check_Payment_Trans"
        lsSQL(1, 3) = "Clients"
        lsSQL(1, 4) = "Complementary"
        lsSQL(1, 5) = "Complementary_Trans"
        lsSQL(1, 6) = "Credit_Card_Trans"
        lsSQL(1, 7) = "Daily_Summary"
        lsSQL(1, 8) = "Discount"
        lsSQL(1, 9) = "Gift_Certificate_Trans"
        lsSQL(1, 10) = "Order_Split"
        lsSQL(1, 11) = "Order_Split_Detail"
        lsSQL(1, 12) = "Payment"
        lsSQL(1, 13) = "Receipt_Master"
        lsSQL(1, 14) = "Return_Detail"
        lsSQL(1, 15) = "Return_Master"
        lsSQL(1, 16) = "SO_Detail"
        lsSQL(1, 17) = "SO_Master"
        lsSQL(1, 18) = "Table_Master"
        lsSQL(1, 19) = "Cash_Reg_Machine"
        lsSQL(1, 20) = "Daily_Summary"
        lsSQL(1, 21) = "Event_Master"
        lsSQL(1, 22) = "Charge_Invoice_Collection_Master"
        lsSQL(1, 23) = "Charge_Invoice_Collection_Detail"
        lsSQL(1, 24) = "Discount_Detail"
        lsSQL(1, 25) = "Charge_Invoice"

        Dim lnCtr As Integer

        p_oAppDriver.BeginTransaction()
        For lnCtr = 0 To 24
            p_oAppDriver.Execute(lsSQL(0, lnCtr), lsSQL(1, lnCtr))
        Next
        p_oAppDriver.CommitTransaction()

        Return True
errProc:
        p_oAppDriver.RollBackTransaction()
        MsgBox("Unable to Reset POS. Please Inform MIS Immediately.", MsgBoxStyle.Critical, "Warning")
        Return False
    End Function

    Private Sub BackupDatabaseToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BackupDatabaseToolStripMenuItem1.Click
        BackupDatabase()
    End Sub

    Private Sub RestoreDatabseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RestoreDatabseToolStripMenuItem.Click
        RestoreDatabase()
    End Sub

    Private Sub ItemListToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemListToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsProductList
        loRpt = New ggcRMSReports.clsProductList(p_oAppDriver)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub SalesReportToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SalesReportToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsSalesReport
        loRpt = New ggcRMSReports.clsSalesReport(p_oAppDriver, p_sPOSNo)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub VoidItemsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VoidItemsToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsVoidItems
        loRpt = New ggcRMSReports.clsVoidItems(p_oAppDriver, p_sPOSNo)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub CancelledReceiptToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CancelledReceiptToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsCancelledReport
        loRpt = New ggcRMSReports.clsCancelledReport(p_oAppDriver, p_sPOSNo)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub BIRSalesSummaryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BIRSalesSummaryToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsSalesSummary
        loRpt = New ggcRMSReports.clsSalesSummary(p_oAppDriver, p_sPOSNo, p_sSerial, p_sPermit, p_sVATReg)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub EventLogsToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EventLogsToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsEventLogs
        loRpt = New ggcRMSReports.clsEventLogs(p_oAppDriver)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub AccumulatedGrandTotalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AccumulatedGrandTotalToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsGranTotalReport
        loRpt = New ggcRMSReports.clsGranTotalReport(p_oAppDriver)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub ComplementaryToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComplementaryToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsComplementary
        loRpt = New ggcRMSReports.clsComplementary(p_oAppDriver, p_sPOSNo)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub ChargeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChargeToolStripMenuItem.Click
        Dim loRpt As ggcRMSReports.clsCharge
        loRpt = New ggcRMSReports.clsCharge(p_oAppDriver, p_sPOSNo)

        Call loRpt.ReportTrans()
    End Sub

    Private Sub SpecialDiscountToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SpecialDiscountToolStripMenuItem.Click
        With frmDiscount
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub EJournalToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EJournalToolStripMenuItem.Click
        With frmTerminal
            If isFormOpen(.Name) Then .Show()
            .MdiParent = Me
            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub YReadingToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles YReadingToolStripMenuItem.Click
        With frmReports
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub StandardToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles StandardToolStripMenuItem.Click
        With frmReports
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Grider = p_oAppDriver
            .Refresh()
        End With
    End Sub

    Private Sub BillingOfStatementToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles BillingOfStatementToolStripMenuItem.Click
        With frmSOAEntry
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub DeliveryServiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeliveryServiceToolStripMenuItem.Click
        With frmDeliveryService
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub StateOfAccountToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StateOfAccountToolStripMenuItem.Click
        With frmSOATagging
            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub ProductUploadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProductUploadToolStripMenuItem.Click
        With frmUtilProductUpload

            If isFormOpen(.Name) Then Exit Sub

            .MdiParent = Me

            .Show()
            .Refresh()
        End With
    End Sub

    Private Sub PrintToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PrintToolStripMenuItem.Click
        Dim loZReading As PRN_TZ_Reading

        loZReading = New PRN_TZ_Reading(p_oAppDriver)
        loZReading.isBackend = False
        loZReading.doPrintUtilityZReading()
    End Sub
End Class