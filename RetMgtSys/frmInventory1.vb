Imports MySql.Data.MySqlClient
Imports ggcAppDriver
Imports ggcRetailParams
Imports System.IO

Public Class frmInventory1
    Private Const pxeImagePath As String = "C:\GGC_Systems\Images\vb.net\Pedritos\Inventory\"
    Private pnLoadx As Integer
    Private poControl As Control

    Private WithEvents p_oRecord As clsInventory
    Private p_nEditMode As Integer
    Private pnSeek As Integer
    Private pnIndx As Integer

    Private Sub frmInventory_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If pnLoadx = 1 Then
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmInventory_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        Select Case e.KeyCode
            Case Keys.Return, Keys.Up, Keys.Down
                Select Case e.KeyCode
                    Case Keys.Return, Keys.Down
                        SetNextFocus()
                    Case Keys.Up
                        SetPreviousFocus()
                End Select
        End Select
    End Sub

    Private Sub frmInventory_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If pnLoadx = 0 Then
            p_oRecord = New clsInventory(p_oAppDriver)

            'Set event Handler for txtField
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "GotFocus", AddressOf txtField_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtField", "LostFocus", AddressOf txtField_LostFocus)
            Call grpCancelHandler(Me, GetType(TextBox), "txtField", "Validating", AddressOf txtField_Validating)
            Call grpKeyHandler(Me, GetType(TextBox), "txtField", "KeyDown", AddressOf txtField_KeyDown)

            Call grpEventHandler(Me, GetType(TextBox), "txtSeeks", "GotFocus", AddressOf txtSeeks_GotFocus)
            Call grpEventHandler(Me, GetType(TextBox), "txtSeeks", "LostFocus", AddressOf txtSeeks_LostFocus)
            Call grpKeyHandler(Me, GetType(TextBox), "txtSeeks", "KeyDown", AddressOf txtSeeks_KeyDown)

            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)

            clearFields()
            initButton()

            pnLoadx = 1
        End If
    End Sub

    Private Sub txtField_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 1, 2, 3, 17, 33, 8 To 14, 80 To 87
                    p_oRecord.Master(loIndex) = loTxt.Text
            End Select
        End If
    End Sub

    Private Function isEntryOk()
        If txtField85.Text = "" Then
            MsgBox("Inventory Type cannot be empty! ." & vbCrLf & _
                   "Please check entry and try again!", MsgBoxStyle.Information, "Error")
            txtField85.Focus()
            Return False
            Exit Function
        End If
        Return True
    End Function

    Private Sub txtField_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 17
                    loTxt.Text = Format(p_oRecord.Master(loIndex), "yyyy/MM/dd")
                Case 33
                    Debug.Print(p_oRecord.Master(loIndex))
                    loTxt.Text = Format(p_oRecord.Master(loIndex), "yyyy/MM/dd")
            End Select
        End If

        pnIndx = loIndex

        poControl = loTxt

        loTxt.BackColor = Color.Azure
        loTxt.SelectAll()
    End Sub

    Private Sub txtField_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtField" Then
                Select Case loIndex
                    Case 80 To 85
                        p_oRecord.SearchMaster(loIndex, loTxt.Text)
                End Select
            End If
        End If
    End Sub

    Private Sub txtField_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtField" Then
            Select Case loIndex
                Case 17
                    loTxt.Text = Format(p_oRecord.Master(loIndex), "MMM dd, yyyy")
                Case 33
                    loTxt.Text = Format(p_oRecord.Master(loIndex), "MMM dd, yyyy")
            End Select
        End If

        loTxt.BackColor = SystemColors.Window
        poControl = Nothing
    End Sub

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        loChk = CType(sender, System.Windows.Forms.Button)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        With p_oRecord
            Select Case lnIndex
                Case 1 'new
                    If p_oRecord.NewRecord Then
                        clearFields()
                        loadMaster()
                    End If
                Case 2 'save

                    .Master(86) = txtField86.Text
                    .Master(87) = txtField87.Text

                    If Not isEntryOk() Then Exit Sub
                    If .SaveRecord Then
                        MsgBox("Record Saved Successfuly.", MsgBoxStyle.Information, "Success")

                        clearFields()
                        If .OpenRecord(.Master("sStockIDx")) Then loadMaster()
                    End If
                Case 3 'search
                    Dim loTxt As TextBox

                    loTxt = CType(FindTextBox(Me, "txtField" & Format(pnIndx, "00")), TextBox)

                    Select Case pnIndx
                        Case 80 To 85
                            p_oRecord.SearchMaster(pnIndx, "%")

                            loTxt.Focus()
                    End Select
                    GoTo endProc
                Case 4 'browse
                    If pnSeek = 1 Then
                        If p_oRecord.SearchRecord(txtSeeks01.Text, True) Then loadMaster()
                    Else
                        If p_oRecord.SearchRecord(txtSeeks02.Text, False) Then loadMaster()
                    End If
                Case 5 'cancel
                    If p_oRecord.CancelUpdate() Then clearFields()
                Case 6 'update
                    If Trim(txtField01.Text) <> "" Then

                        p_oRecord.UpdateRecord()
                    End If
                Case 7 'delete
                    If MsgBox("Do you want to disable this item?", MsgBoxStyle.Question & MsgBoxStyle.YesNo, "Confirm") = MsgBoxResult.Yes Then
                        If p_oRecord.CancelRecord Then
                            MsgBox("Record has disabled successfuly.", MsgBoxStyle.Information, "Notice")
                            clearFields()
                        End If
                    End If
                Case 8 'close
                    Me.Close()
                    GoTo endProc
                Case 9 'show combo
                    p_oRecord.showComboMeals()
                    GoTo endProc
            End Select
        End With

        initButton()
endProc:
        Exit Sub
    End Sub

    Private Sub initButton()
        Dim lbShow As Integer
        Dim lnEditMode As xeEditMode = p_oRecord.EditMode

        lbShow = (lnEditMode = 1 Or lnEditMode = 2)

        cmdButton02.Visible = lbShow
        cmdButton03.Visible = lbShow
        cmdButton05.Visible = lbShow
        Panel1.Enabled = lbShow
        Panel2.Enabled = lbShow
        Panel4.Enabled = lbShow
        Panel6.Enabled = lbShow
        Panel7.Enabled = lbShow

        cmdButton01.Visible = Not lbShow
        cmdButton06.Visible = Not lbShow
        cmdButton07.Visible = Not lbShow
        cmdButton08.Visible = Not lbShow
        Panel3.Enabled = Not lbShow
        Panel2.Enabled = p_oRecord.EditMode = xeEditMode.MODE_ADDNEW
        CheckBox3.Enabled = p_oRecord.EditMode = xeEditMode.MODE_ADDNEW
        If lbShow Then
            txtField02.Focus()
        Else
            txtSeeks01.Focus()
        End If
    End Sub

    Private Sub loadMaster()
        With p_oRecord
            txtSeeks01.Text = ""
            txtSeeks02.Text = ""

            txtField01.Text = .Master(1)
            txtField02.Text = .Master(2)
            txtField03.Text = .Master(3)
            txtField80.Text = .Master(80)
            txtField81.Text = .Master(81)
            txtField82.Text = .Master(82)
            txtField08.Text = Format(.Master(8), "#,##0.00")
            txtField09.Text = Format(.Master(9), "#,##0.00")
            txtField10.Text = Format(IFNull(.Master(10), 0), "#,##0.00")
            txtField11.Text = Format(IFNull(.Master(11), 0), "#,##0.00")
            txtField12.Text = Format(IFNull(.Master(12), 0), "#,##0.00")
            txtField13.Text = Format(IFNull(.Master(13), 0), "#,##0.00")
            txtField14.Text = IFNull(.Master(14), "")
            txtField83.Text = .Master(83)
            txtField84.Text = .Master(84)
            txtField85.Text = .Master(85)
            txtField18.Text = IFNull(.Master(18))

            txtField86.Text = Format(.Master(8), "#,##0.00")
            txtField87.Text = Format(.Master(9), "#,##0.00")

            txtField17.Text = Format(IFNull(.Master(17), p_oAppDriver.getSysDate), "MMM dd, yyyy")
            txtField19.Text = IFNull(.Master(19))
            txtField20.Text = IFNull(.Master(20))
            txtField21.Text = IFNull(.Master(21))
            txtField22.Text = IFNull(.Master(22))
            txtField23.Text = IFNull(.Master(23))
            txtField24.Text = IFNull(.Master(24))

            txtField86.Text = Format(.Master(8), "#,##0.00")
            txtField87.Text = Format(.Master(9), "#,##0.00")

            CheckBox2.Checked = IIf(.Master("cWthPromo") = "1", True, False)
            CheckBox1.Checked = IIf(.Master("cComboMlx") = "1", True, False)
            CheckBox3.Checked = IIf(.Master("cRecdStat") = "1", True, False)
            CheckBox4.Checked = IIf(.Master(88) = "1", True, False)


            txtField33.Text = Format(IFNull(.Master(33), p_oAppDriver.getSysDate), "MMM dd, yyyy")

            If IFNull(.Master("sImgePath"), "") <> "" Then
                For nCtr As Integer = 0 To fileList.Items.Count - 1
                    fileList.SelectedIndex = nCtr
                    If pxeImagePath & fileList.SelectedItem = Replace(.Master("sImgePath"), "/", "\") Then
                        filePic.BackgroundImage = Image.FromFile(pxeImagePath & fileList.SelectedItem)
                        Exit For
                    End If
                Next nCtr
            End If
        End With
    End Sub

    Private Sub clearFields()
        txtSeeks01.Text = ""
        txtSeeks02.Text = ""

        txtField01.Text = ""
        txtField02.Text = ""
        txtField03.Text = ""
        txtField80.Text = ""
        txtField81.Text = ""
        txtField82.Text = ""
        txtField08.Text = Format(0.0#, "#,##0.00")
        txtField09.Text = Format(0.0#, "#,##0.00")
        txtField86.Text = Format(0.0#, "#,##0.00")
        txtField87.Text = Format(0.0#, "#,##0.00")
        txtField10.Text = Format(0.0#, "#,##0.00")
        txtField11.Text = Format(0.0#, "#,##0.00")
        txtField12.Text = Format(0.0#, "#,##0.00")
        txtField13.Text = Format(0.0#, "#,##0.00")
        txtField14.Text = ""
        txtField83.Text = ""
        txtField84.Text = ""
        txtField85.Text = ""
        txtField18.Text = "0"
        txtField17.Text = Format(p_oAppDriver.SysDate, "MMM dd, yyyy")
        txtField33.Text = Format(p_oAppDriver.SysDate, "MMM dd, yyyy")
        txtField19.Text = "0"
        txtField20.Text = "0"
        txtField21.Text = "0"
        txtField22.Text = "0"
        txtField23.Text = "0"
        txtField24.Text = "0"

        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        CheckBox4.Checked = True

        Try
            Dim ScanFolder() As String
            Dim lsFileName() As String
            ScanFolder = Directory.GetFiles(pxeImagePath, "*.jpg")

            fileList.Items.Clear()
            For Each dFile As String In ScanFolder
                lsFileName = dFile.Split("\")
                fileList.Items.Add(lsFileName(UBound(lsFileName)))
            Next

            If fileList.Items.Count > 0 Then
                fileList.SelectedIndex = 0
                filePic.BackgroundImage = Image.FromFile(pxeImagePath & "(none).jpg")
            Else
                filePic.BackgroundImage = Nothing
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private Sub p_oRecord_MasterRetrieved(ByVal Index As Integer, ByVal Value As Object) Handles p_oRecord.MasterRetrieved
        Dim loTxt As TextBox

        loTxt = CType(FindTextBox(Me, "txtField" & Format(Index, "00")), TextBox)

        Select Case Index
            Case 18, 33
                Debug.Print(Value)
                loTxt.Text = Format(Value, "MMM dd, yyyy")
            Case 8 To 13, 86, 87
                loTxt.Text = Format(Value, "#,##0.00")
            Case 80 To 85
                loTxt.Text = Value
            Case 88

                CheckBox4.Checked = IIf(Value = "1", True, False)
            Case Else
                loTxt.Text = Value
        End Select
    End Sub

    Private Sub txtSeeks_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtSeeks" Then
            Select Case loIndex
                Case 1, 2
                    pnSeek = loIndex
            End Select
        End If

        poControl = loTxt

        loTxt.BackColor = Color.Azure
        loTxt.SelectAll()
    End Sub

    Private Sub txtSeeks_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim loTxt As TextBox
        loTxt = CType(sender, System.Windows.Forms.TextBox)

        Dim loIndex As Integer
        loIndex = Val(Mid(loTxt.Name, 9))

        If Mid(loTxt.Name, 1, 8) = "txtSeeks" Then
            Select Case loIndex
            End Select
        End If

        loTxt.BackColor = SystemColors.Window
        poControl = Nothing
    End Sub

    Private Sub txtSeeks_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs)
        If e.KeyCode = Keys.F3 Or e.KeyCode = Keys.Enter Then
            Dim loTxt As TextBox
            loTxt = CType(sender, System.Windows.Forms.TextBox)
            Dim loIndex As Integer
            loIndex = Val(Mid(loTxt.Name, 9))

            If Mid(loTxt.Name, 1, 8) = "txtSeeks" Then
                Select Case loIndex
                    Case 1
                        If p_oRecord.SearchRecord(txtSeeks01.Text, True) Then
                            loadMaster()
                        Else
                            SetNextFocus()
                        End If
                    Case 2
                        If p_oRecord.SearchRecord(txtSeeks02.Text, False) Then
                            loadMaster()
                        Else
                            SetNextFocus()
                        End If
                End Select
            End If
        End If
    End Sub
    Private Sub CheckBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox4.Click
        p_oRecord.Master("cRecdStat") = IIf(CheckBox4.Checked, "1", "0")
    End Sub
    Private Sub CheckBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.Click
        p_oRecord.Master("cRecdStat") = IIf(CheckBox3.Checked, "1", "0")
    End Sub

    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click
        p_oRecord.Master("cWthPromo") = IIf(CheckBox2.Checked, "1", "0")
    End Sub

    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        p_oRecord.Master("cComboMlx") = IIf(CheckBox1.Checked, "1", "0")
    End Sub

    Private Sub fileList_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles fileList.Click
        Try
            filePic.BackgroundImage = Image.FromFile(pxeImagePath & fileList.SelectedItem)
            p_oRecord.Master("sImgePath") = pxeImagePath & fileList.SelectedItem
        Catch ex As Exception
            ex.ToString()
        End Try
    End Sub

End Class