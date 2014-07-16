Imports Seagull.BarTender.Print

Public Class WebForm3
    Inherits System.Web.UI.Page

    Private btPrintersList As Printers

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.LoadComplete
        'load dropdownlist of printers only when page is loaded for the first time
        If Not IsPostBack Then
            loadPrinterList()
            loadLabelList()
        End If
    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to set up list of printers in drop down list
    '-----------------------------------------------------------------------------------------
    Private Sub loadPrinterList()
        ' create list of default Windows printers to ignore
        Dim ignoreList As New List(Of String)
        ignoreList.Add("Fax")
        ignoreList.Add("Microsoft XPS Document Writer")
        ignoreList.Add("Send To OneNote 2010")

        ' Initialize a new BarTender print engine.
        Using btEngine As New Engine()
            ' Start the BarTender print engine.
            btEngine.Start()

            ' Get list of printers.
            btPrintersList = New Printers()

            ' datasource list for DropDownList
            Dim strPrinterList As New List(Of String)

            ' add default printer first so its on the top of the list
            strPrinterList.Add(btPrintersList.Default.PrinterName)
            ' loop through the list of printers
            For Each p In btPrintersList
                ' dont add default printer or printers in ignore list 
                If Not p.IsDefault And Not ignoreList.Contains(p.PrinterName) Then
                    strPrinterList.Add(p.PrinterName)
                End If
            Next

            ' bind data to DropDownList
            dropdownlist.DataSource = strPrinterList
            dropdownlist.DataBind()

            ' Stop the BarTender print engine.
            btEngine.Stop()
        End Using
    End Sub

    Private Sub loadLabelList()
        Dim strLabelList As New List(Of String)
        strLabelList.Add("Large Label")
        strLabelList.Add("Small Label")

        dropdownlistlabel.DataSource = strLabelList
        dropdownlistlabel.DataBind()
    End Sub

    Protected Sub printbutton_Click(sender As Object, e As EventArgs) Handles printbutton.Click
        If CopiesNumericCheck() AndAlso dropdownlistlabel.Text = "Small Label" Then
            Dim lineList() As String = ParseLines(textarea.InnerText)
            PrintCustomSmallLabel(lineList)
        ElseIf CopiesNumericCheck() AndAlso dropdownlistlabel.Text = "Large Label" Then
            Dim lineList() As String = ParseLines(textarea.InnerText)
            PrintCustomLargeLabel(lineList)
        End If
    End Sub

    '-------------------------------------------------------------------------------------------
    ' Function to check the values of "Copies" fields to ensure they are all numerical values
    '-------------------------------------------------------------------------------------------
    Private Function CopiesNumericCheck() As Boolean
        'make sure number copies entered by user is numeric
        If Not IsNumeric(copies.Text) Then
            tbWarning.Text = "Your ""Copies"" field contains a non-numeric value"
            Return False
        End If
        tbWarning.Text = ""
        Return True
    End Function

    '-------------------------------------------------------------------------------------------
    ' Function to parse lines seperated by "\n" for label printing
    '-------------------------------------------------------------------------------------------
    Private Function ParseLines(ByVal text As String) As String()
        Dim lineList() As String = text.Split(vbLf)
        Return lineList
    End Function

    Private Sub PrintCustomSmallLabel(ByVal lineList As String())
        'make sure its not 0 or >4 lines
        If lineList.Length = 0 Or lineList.Length > 2 Then
            tbWarning.Text = "You can only print a small label with 1-2 lines."
        End If

        'convert no. of copies to int and print labels
        Dim smallCopies As Integer = Convert.ToInt16(copies.Text)
        If smallCopies > 0 Then

            'print info
            Dim btEngine As New Engine(True)
            Dim labelFormat As LabelFormatDocument

            'fill in designated fields on label depending on lines entered
            Select Case lineList.Length
                Case 1
                    tbWarning.Text = ""
                    labelFormat = btEngine.Documents.Open(OneLineCustomSmallLabel)
                    labelFormat.SubStrings.SetSubString("custom1", lineList(0))
                Case 2
                    tbWarning.Text = ""
                    labelFormat = btEngine.Documents.Open(TwoLineCustomSmallLabel)
                    labelFormat.SubStrings.SetSubString("custom1", lineList(0))
                    labelFormat.SubStrings.SetSubString("custom2", lineList(1))
                Case Else
                    Exit Sub
            End Select

            'print however many copies desired
            labelFormat.PrintSetup.IdenticalCopiesOfLabel = smallCopies
            labelFormat.PrintSetup.PrinterName = dropdownlist.SelectedItem.ToString
            labelFormat.Print()

            'close print engine
            btEngine.Stop()

            'clear textarea
            clearTextArea()
        End If
        copies.Text = "1"
    End Sub

    Private Sub PrintCustomLargeLabel(ByVal lineList As String())
        'make sure its not 0 or >4 lines
        If lineList.Length = 0 Or lineList.Length > 4 Then
            tbWarning.Text = "You can only print a large label with 1-4 lines."
        End If

        'convert no. of copies to int and print labels
        Dim largeCopies As Integer = Convert.ToInt16(copies.Text)
        If largeCopies > 0 Then

            'print info
            Dim btEngine As New Engine(True)
            Dim labelFormat As LabelFormatDocument

            'fill in designated fields on label depending on lines entered
            Select Case lineList.Length
                Case 1
                    tbWarning.Text = ""
                    labelFormat = btEngine.Documents.Open(OneLineCustomLargeLabel)
                    labelFormat.SubStrings.SetSubString("custom1", lineList(0))
                Case 2
                    tbWarning.Text = ""
                    labelFormat = btEngine.Documents.Open(TwoLineCustomLargeLabel)
                    labelFormat.SubStrings.SetSubString("custom1", lineList(0))
                    labelFormat.SubStrings.SetSubString("custom2", lineList(1))
                Case 3
                    tbWarning.Text = ""
                    labelFormat = btEngine.Documents.Open(ThreeLineCustomLargeLabel)
                    labelFormat.SubStrings.SetSubString("custom1", lineList(0))
                    labelFormat.SubStrings.SetSubString("custom2", lineList(1))
                    labelFormat.SubStrings.SetSubString("custom3", lineList(2))
                Case 4
                    tbWarning.Text = ""
                    labelFormat = btEngine.Documents.Open(FourLineCustomLargeLabel)
                    labelFormat.SubStrings.SetSubString("custom1", lineList(0))
                    labelFormat.SubStrings.SetSubString("custom2", lineList(1))
                    labelFormat.SubStrings.SetSubString("custom3", lineList(2))
                    labelFormat.SubStrings.SetSubString("custom4", lineList(3))
                Case Else
                    Exit Sub
            End Select

            'print however many copies desired
            labelFormat.PrintSetup.IdenticalCopiesOfLabel = largeCopies
            labelFormat.PrintSetup.PrinterName = dropdownlist.SelectedItem.ToString
            labelFormat.Print()

            'close print engine
            btEngine.Stop()

            'clear textarea
            clearTextArea()
        End If
        copies.Text = "1"
    End Sub

    Private Sub clearTextArea()
        textarea.InnerText = ""
    End Sub
End Class