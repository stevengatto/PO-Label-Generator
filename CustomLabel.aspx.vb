'+--------------------------------------------------+
'| Author: Steven Gatto - IT Intern                 |
'| Last Updated: 7/25/14                            |
'|                                                  |
'| This program is aimed to help the receiving      |
'| dept automate label printing and increase        |
'| productivity.                                    |
'+--------------------------------------------------+

Imports Seagull.BarTender.Print

Public Class CustomLabel
    Inherits System.Web.UI.Page

    Private btPrintersList As Printers
    Dim smallLabelName As String = "Small Label (2 line max)"
    Dim largeLabelName As String = "Large Label (4 line max)"


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
        ignoreList.Add("\\IPG410-1128-04\ZDesigner TLP 2844")

        ' Initialize a new BarTender print engine.
        Using btEngine As New Engine()
            ' Start the BarTender print engine.
            btEngine.Start()

            ' Get list of printers.
            btPrintersList = New Printers()

            ' datasource list for DropDownList
            Dim strPrinterList As New List(Of String)

            ' loop through the list of printers
            For Each p In btPrintersList
                ' dont add printers in ignore list 
                If Not ignoreList.Contains(p.PrinterName) Then
                    ' change printer names to their display names
                    If p.PrinterName.Equals(AllowedPrinters.DanPrinterNetworkPath) Then
                        strPrinterList.Add(AllowedPrinters.DanPrinterDisplayName)
                    ElseIf p.PrinterName.Equals(AllowedPrinters.RonPrinterNetworkPath) Then
                        strPrinterList.Add(AllowedPrinters.RonPrinterDisplayName)
                    ElseIf p.PrinterName.Equals(AllowedPrinters.ErikPrinterNetworkPath) Then
                        strPrinterList.Add(AllowedPrinters.ErikPrinterDisplayName)
                    ElseIf p.PrinterName.Equals(AllowedPrinters.PrestonPrinterNetworkPath) Then
                        strPrinterList.Add(AllowedPrinters.PrestonPrinterDisplayName)
                    Else
                        strPrinterList.Add(p.PrinterName)
                    End If
                End If
            Next

            ' bind data to DropDownList
            dropdownlist.DataSource = strPrinterList

            'dropdownlist.DataSource = strPrinterList
            dropdownlist.DataBind()

            ' Stop the BarTender print engine.
            btEngine.Stop()
        End Using
    End Sub


    '-----------------------------------------------------------------------------------------
    ' Sub to set up list of label types in drop down list
    '-----------------------------------------------------------------------------------------
    Private Sub loadLabelList()
        Dim strLabelList As New List(Of String)
        strLabelList.Add(largeLabelName)
        strLabelList.Add(smallLabelName)

        dropdownlistlabel.DataSource = strLabelList
        dropdownlistlabel.DataBind()
    End Sub


    '-----------------------------------------------------------------------------------------
    ' Sub to handle click event of Print button. Calls helper functions and print subroutine
    '-----------------------------------------------------------------------------------------
    Protected Sub printbutton_Click(sender As Object, e As EventArgs) Handles printbutton.Click
        If CopiesNumericCheck() AndAlso dropdownlistlabel.Text = smallLabelName Then
            Dim lineList() As String = ParseLines(textarea.InnerText)
            PrintCustomSmallLabel(lineList)
        ElseIf CopiesNumericCheck() AndAlso dropdownlistlabel.Text = largeLabelName Then
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


    '-------------------------------------------------------------------------------------------
    ' Sub to print custom small labels
    '-------------------------------------------------------------------------------------------
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
            labelFormat = btEngine.Documents.Open(CustomSmallLabel)

            'fill in designated fields on label depending on lines entered
            '
            ' Note:
            '   The small custom label template has two layers of templates.
            '       - Layer one - One line that takes up the entire label
            '       - Layer two - Two lines each taking up half the label
            Select Case lineList.Length
                Case 1
                    tbWarning.Text = ""
                    labelFormat.SubStrings.SetSubString("OneLineLayer_1", lineList(0))
                Case 2
                    tbWarning.Text = ""
                    labelFormat.SubStrings.SetSubString("TwoLineLayer_1", lineList(0))
                    labelFormat.SubStrings.SetSubString("TwoLineLayer_2", lineList(1))
                Case Else
                    Exit Sub
            End Select

            'print however many copies desired
            labelFormat.PrintSetup.IdenticalCopiesOfLabel = smallCopies

            'decide which printer to print to
            Dim printerSelected As String = dropdownlist.SelectedItem.ToString
            If printerSelected = AllowedPrinters.DanPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.DanPrinterNetworkPath
            ElseIf printerSelected = AllowedPrinters.RonPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.RonPrinterNetworkPath
            ElseIf printerSelected = AllowedPrinters.PrestonPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.PrestonPrinterNetworkPath
            ElseIf printerSelected = AllowedPrinters.ErikPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.ErikPrinterNetworkPath
            Else
                labelFormat.PrintSetup.PrinterName = printerSelected
            End If

            labelFormat.Print()

            'close print engine
            btEngine.Stop()

            'clear textarea
            clearTextArea()
        End If
        copies.Text = "1"
    End Sub

        '-------------------------------------------------------------------------------------------
        ' Sub to print custom large labels
        '-------------------------------------------------------------------------------------------
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
            labelFormat = btEngine.Documents.Open(CustomLargeLabel)

            'fill in designated fields on label depending on lines entered
            '
            ' Note:
            '   The large custom label template has four layers of templates.
            '       - Layer one - One line that takes up the entire label
            '       - Layer two - Two lines each taking up half the label
            '       - Layer three - Three lines each taking up 1/3 the label
            '       - Layer four - Four lines each taking up 1/4 the label
            Select Case lineList.Length
                Case 1
                    tbWarning.Text = ""
                    labelFormat.SubStrings.SetSubString("OneLineLayer_1", lineList(0))
                Case 2
                    tbWarning.Text = ""
                    labelFormat.SubStrings.SetSubString("TwoLineLayer_1", lineList(0))
                    labelFormat.SubStrings.SetSubString("TwoLineLayer_2", lineList(1))
                Case 3
                    tbWarning.Text = ""
                    labelFormat.SubStrings.SetSubString("ThreeLineLayer_1", lineList(0))
                    labelFormat.SubStrings.SetSubString("ThreeLineLayer_2", lineList(1))
                    labelFormat.SubStrings.SetSubString("ThreeLineLayer_3", lineList(2))
                Case 4
                    tbWarning.Text = ""
                    labelFormat.SubStrings.SetSubString("FourLineLayer_1", lineList(0))
                    labelFormat.SubStrings.SetSubString("FourLineLayer_2", lineList(1))
                    labelFormat.SubStrings.SetSubString("FourLineLayer_3", lineList(2))
                    labelFormat.SubStrings.SetSubString("FourLineLayer_4", lineList(3))
                Case Else
                    Exit Sub
            End Select

            'print however many copies desired
            labelFormat.PrintSetup.IdenticalCopiesOfLabel = largeCopies

            'decide which printer to print to
            Dim printerSelected As String = dropdownlist.SelectedItem.ToString
            If printerSelected = AllowedPrinters.DanPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.DanPrinterNetworkPath
            ElseIf printerSelected = AllowedPrinters.RonPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.RonPrinterNetworkPath
            ElseIf printerSelected = AllowedPrinters.PrestonPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.PrestonPrinterNetworkPath
            ElseIf printerSelected = AllowedPrinters.ErikPrinterDisplayName Then
                labelFormat.PrintSetup.PrinterName = AllowedPrinters.ErikPrinterNetworkPath
            Else
                labelFormat.PrintSetup.PrinterName = printerSelected
            End If

            labelFormat.Print()

            'close print engine
            btEngine.Stop()

            'clear textarea
            clearTextArea()
        End If
        copies.Text = "1"
    End Sub


    '-------------------------------------------------------------------------------------------
    ' Sub to clear the contents of the textarea
    '-------------------------------------------------------------------------------------------
    Private Sub clearTextArea()
        textarea.InnerText = ""
    End Sub


End Class