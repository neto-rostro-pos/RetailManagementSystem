Imports ggcAppDriver
Imports System
Imports System.IO
Imports System.Text
Imports PdfSharp
Imports PdfSharp.Pdf
Imports PdfSharp.Drawing

Public Class frmEjournal
    Private pnLoadx As Integer
    Private p_bCancelled As Boolean
    Private p_oIDNumbr As String
    Private tempDirectory As String = "D:\tempEjournal"
    Private tempFile As String = "D:\tempEjournal\temp.txt"

    Public Property IDNumber As String
        Get
            Return p_oIDNumbr
        End Get
        Set(ByVal value As String)
            p_oIDNumbr = value
        End Set
    End Property

    Private Sub frmEjournal_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Activated
        If pnLoadx = 1 Then
            pnLoadx = 2
        End If
    End Sub

    Private Sub frmEjournal_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
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

    Private Sub cmdButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim loChk As Button
        Dim fileDirectory As String = "D:\ejournal\"
        Dim fileName As New List(Of String)
        Dim fileToPdf As String = ""
        Dim files() As System.IO.FileInfo
        Dim xReading As String = "X-READING"
        Dim zReading As String = "Z-READING"
        Dim dDateFrom As String = Replace(txtField00.Text, "-", "")
        Dim dDateThru As String = Replace(txtField02.Text, "-", "")
        Dim dirInfo As New System.IO.DirectoryInfo(fileDirectory)
        files = dirInfo.GetFiles("*.txt", System.IO.SearchOption.AllDirectories)

        loChk = CType(sender, System.Windows.Forms.Button)
        Dim lnIndex As Integer
        lnIndex = Val(Mid(loChk.Name, 10))

        Select Case lnIndex
            Case 0 'ok
                Select Case cmblist.SelectedIndex
                    Case 0
                        For Each file In files
                            If file.Name.Contains(xReading) Then
                                If file.Name.Substring(28, 8) >= dDateFrom And file.Name.Substring(28, 8) <= dDateThru Then
                                    fileName.Add(file.ToString)
                                End If
                            End If
                        Next
                    Case 1
                        For Each file In files
                            If file.Name.Contains(zReading) Then
                                If file.Name.Substring(28, 8) >= dDateFrom And file.Name.Substring(28, 8) <= dDateThru Then
                                    fileName.Add(file.ToString)
                                End If
                            End If
                        Next
                    Case 2
                        For Each file In files
                            If file.Name.Length = 21 Then GoTo nextEntry
                            If Not file.Name.Contains(xReading) And Not file.Name.Contains(zReading) Then
                                If file.Name.Substring(18, 8) >= dDateFrom And file.Name.Substring(18, 8) <= dDateThru Then
                                    fileName.Add(file.ToString)
                                End If
                            End If
nextEntry:
                        Next
                End Select
                p_bCancelled = False
                txtField01.Text = ""

                If fileName.Count = 0 Then
                    MessageBox.Show("Invalid search critera" & vbCrLf &
                                     "Please try again..")
                    txtField01.Text = ""
                    Exit Sub
                End If

                For lnctr = 0 To fileName.Count - 1
                    Try
                        For Each fsReadlines As String In System.IO.File.ReadAllLines(fileDirectory & fileName(lnctr))
                            txtField01.AppendText(fsReadlines + vbNewLine)
                        Next
                    Catch ex As Exception
                        MessageBox.Show("Invalid search critera" & vbCrLf &
                                        "Please try again..")
                        txtField01.Text = ""
                    End Try
                Next

                If Not File.Exists(tempFile) Then
                    Dim tmpFile As FileStream
                    tmpFile = File.Create(tempFile)
                    tmpFile.Close()
                End If
                System.IO.File.WriteAllText(tempFile, "")
                File.AppendAllText(tempFile, txtField01.Text)

            Case 2
                If txtField01.Text = "" Then Exit Sub
                Dim readFile As System.IO.TextReader
                Dim sfd As New SaveFileDialog() ' this creates an instance of the SaveFileDialog called "sfd"
                sfd.Filter = "PDF files (*.pdf)|*.pdf|All files (*.*)|*.*"
                sfd.FilterIndex = 1
                sfd.RestoreDirectory = True

                If sfd.ShowDialog() = DialogResult.OK Then
                    fileToPdf = sfd.FileName ' retrieve the full path to the file selected by the user
                    Try
                        Dim line As String

                        readFile = New StreamReader(tempFile)
                        Dim yPoint As Integer = 0
                        Dim pdf As PdfDocument = New PdfDocument
                        pdf.Info.Title = "Text File to PDF"
                        Dim pdfPage As PdfPage
                        Dim graph As XGraphics
                        Dim font As XFont = New XFont("courier new", 9, Regular)

                        yPoint = 0
nextline:
                        '842'
                        If yPoint >= 820 Then
                            yPoint = 0
                        End If

                        line = readFile.ReadLine()
                        If line Is Nothing Then GoTo AllDone

                        If yPoint = 0 Then
                            pdfPage = pdf.AddPage()
                            graph = XGraphics.FromPdfPage(pdfPage)
                            pdfPage.Size = PageSize.A4
                        End If

Page1:
                        graph.DrawString(line, font, XBrushes.Black,
                        New XRect(10, yPoint, pdfPage.Width.Point, pdfPage.Height.Point), XStringFormats.TopLeft)
                        yPoint = yPoint + 9
                        GoTo nextline

AllDone:
                        Dim pdfFilename As String = fileToPdf
                        pdf.Save(pdfFilename)
                        readFile.Close()
                        readFile = Nothing
                        Process.Start(pdfFilename)
                    Catch ex As Exception
                        MsgBox(ex.ToString)
                    End Try
                End If
            Case 3

                p_bCancelled = True
                Me.Close()
        End Select

endProc:
                Exit Sub
    End Sub

    Private Sub frmEjournal_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If pnLoadx = 0 Then
            clearFields()
            Call grpEventHandler(Me, GetType(Button), "cmdButton", "Click", AddressOf cmdButton_Click)
            Call grpCancelHandler(Me, GetType(TextBox), "txtField", "Validating", AddressOf txtField_Validating)

            If (Not System.IO.Directory.Exists(tempDirectory)) Then
                System.IO.Directory.CreateDirectory(tempDirectory)
            End If

            pnLoadx = 1
        End If
    End Sub

    Private Sub clearFields()
        txtField00.Text = Format(Now().AddDays(-1), "yyyy-MM-dd")
        txtField02.Text = Format(Now(), "yyyy-MM-dd")
        txtField01.Text = ""
        cmblist.SelectedIndex = 0
    End Sub

    Private Sub txtField_Validating(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs)
        Dim lotxt As TextBox
        lotxt = CType(sender, System.Windows.Forms.TextBox)

        Dim lnIndex As Integer
        lnIndex = Val(Mid(lotxt.Name, 9))

        If Mid(lotxt.Name, 1, 8) = "txtField" Then
            Select Case lnIndex
                Case 0
                    If Not IsDate(lotxt.Text) Then
                        lotxt.Text = Now().AddDays(-1)
                    End If
                    lotxt.Text = Format(CDate(lotxt.Text), "yyyy-MM-dd")
                Case 2
                    If Not IsDate(lotxt.Text) Then
                        lotxt.Text = Now()
                    End If
                    lotxt.Text = Format(CDate(lotxt.Text), "yyyy-MM-dd")
            End Select
        End If
    End Sub
End Class