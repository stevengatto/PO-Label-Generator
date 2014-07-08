'+--------------------------------------------------+
'| Author: Steven Gatto - IT Intern                 |
'| Last Updated: 7/8/14                             |
'|                                                  |
'| This program is aimed to help the receiving      |
'| dept automate label printing and increase        |
'| productivity. Enjoy.                             |
'+--------------------------------------------------+

Imports System.Data.SqlClient
Imports Seagull.BarTender.Print
Imports System.Runtime.InteropServices

Public Class WebForm1
    Inherits System.Web.UI.Page

    Dim mainReader, quantityReader, purchaserReader, vendorReader, unitsReader As SqlDataReader
    Dim conn As New SqlConnection(ConfigurationManager.ConnectionStrings("NAV-Dev-ConnectionString").ConnectionString)
    Dim btPrintersList As Printers

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.LoadComplete
        'load dropdownlist of printers only when page is loaded for the first time
        If Not IsPostBack Then loadPrinterList()
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

    '-----------------------------------------------------------------------------------------
    ' Sub to handle "PO Search" button click. Calls helper subs/functions
    '-----------------------------------------------------------------------------------------
    Protected Sub PObutton_Click(sender As Object, e As EventArgs) Handles PObutton.Click
        'set the majority of the content in the GridView
        setMainContent()

        'update Vendor, Quantity, Purchaser, and Units if data was bound
        If GridView1.Rows.Count > 0 Then
            setVendor()
            setQuantity()
            setPurchaserCode()
            setUnits()
        End If

        'show the print button and ddl as long as everything went well with pulling the data
        If GridView1.Rows.Count > 0 Then
            printbutton.Visible = True
            dropdownlist.Visible = True
        Else
            printbutton.Visible = False
            dropdownlist.Visible = False
        End If

    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to set main content of the GridView
    '-----------------------------------------------------------------------------------------
    Private Sub setMainContent()
        'sql query string for all columns except "quantity"
        Dim mainSqlString As String = SQLStrings.getMainSqlString

        'setup command with sql server and parameterize PONumber input
        Dim mainSqlCmd As New SqlCommand(mainSqlString, conn)
        Dim param As New SqlParameter("@PONumber", SqlDbType.NVarChar, 20)
        param.Value = POsearch.Text.ToUpper
        mainSqlCmd.Parameters.Add(param)

        'try to start the sql server connection and set up the gridview datasource
        Try
            conn.Open()
            mainReader = mainSqlCmd.ExecuteReader()
            GridView1.DataSource = mainReader
        Catch ex As Exception
            'tell user something went wrong if sql server connection fails
            GridView1.EmptyDataText = "There has been an error. No data retrieved."
        End Try

        'bind data to gridview no matter what the outcome (to display Data or EmptyDataText)
        GridView1.DataBind()
    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to set Vendor table values
    '-----------------------------------------------------------------------------------------
    Private Sub setVendor()
        'sql query for vendor name and address
        Dim vendorSqlString As String = SQLStrings.getVendorSqlString

        'setup command with sql server and parameterize PONumber input
        Dim vendorSqlCmd As New SqlCommand(vendorSqlString, conn)
        Dim param As New SqlParameter("@PONumber", SqlDbType.NVarChar, 20)
        param.Value = POsearch.Text.ToUpper
        vendorSqlCmd.Parameters.Add(param)

        'close the main reader if it exists and is open
        If Not IsNothing(mainReader) Then
            If Not mainReader.IsClosed() Then
                mainReader.Close()
            End If
        End If

        'get the quantities of each row in seperate sql connection (so they can be editted via textbox)
        Try
            'open connection if it needs to be and execute query
            If conn.State = ConnectionState.Closed Then conn.Open()
            vendorReader = vendorSqlCmd.ExecuteReader()
            'read all parts of vendor address
            vendorReader.Read()
            Dim vendorInfo As IDataRecord = CType(vendorReader, IDataRecord)
            'set up table to contain vendor address
            tableCol1.InnerText = vendorInfo(0).ToString
            tableCol2.InnerText = vendorInfo(1).ToString & ", " &
            vendorInfo(2).ToString & " " & vendorInfo(3).ToString

        Catch ex As Exception
            tbWarning.Text = "There has been an error. Vendor was not retrieved."
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to set "Quantity" values
    '-----------------------------------------------------------------------------------------
    Private Sub setQuantity()
        'sql query for "quantity" column
        Dim quantitySqlString As String = SQLStrings.getQuantitySqlString

        'setup command with sql server and parameterize PONumber input
        Dim quantitySqlCmd As New SqlCommand(quantitySqlString, conn)
        Dim param As New SqlParameter("@PONumber", SqlDbType.NVarChar, 20)
        param.Value = POsearch.Text.ToUpper
        quantitySqlCmd.Parameters.Add(param)

        'close the vendor reader if it exists and is open
        If Not IsNothing(vendorReader) Then
            If Not vendorReader.IsClosed() Then
                vendorReader.Close()
            End If
        End If

        'get the quantities of each row in seperate sql connection (so they can be editted via textbox)
        Try
            'open connection if it needs to be and execute query
            If conn.State = ConnectionState.Closed Then conn.Open()
            quantityReader = quantitySqlCmd.ExecuteReader()
            'loop through each row filling in quantity
            If GridView1.Rows.Count > 0 Then
                For row As Integer = 0 To GridView1.Rows.Count - 1
                    'find the textbox element of the current row
                    Dim txtQuantity As TextBox = GridView1.Rows(row).Cells(2).FindControl("quantityTextBox")
                    'move reader to next entry, record next entry, convert to string, format, and set to textbox
                    quantityReader.Read()
                    Dim record As IDataRecord = CType(quantityReader, IDataRecord)
                    Dim strQuantity = record(0).ToString
                    txtQuantity.Text = strQuantity.Substring(0, strQuantity.IndexOf("."))
                Next
            End If
        Catch ex As Exception
            tbWarning.Text = "There has been an error. Quantites were not retrieved."
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to set "Purchaser" values
    '-----------------------------------------------------------------------------------------
    Private Sub setPurchaserCode()
        'sql query for "quantity" column
        Dim purchaserSqlString As String = SQLStrings.getPurchaserSqlString

        'setup command with sql server and parameterize PONumber input
        Dim purchaserSqlCmd As New SqlCommand(purchaserSqlString, conn)
        Dim param As New SqlParameter("@PONumber", SqlDbType.NVarChar, 20)
        param.Value = POsearch.Text.ToUpper
        purchaserSqlCmd.Parameters.Add(param)

        'close the main reader if it exists and is open
        If Not IsNothing(quantityReader) Then
            If Not quantityReader.IsClosed() Then
                quantityReader.Close()
            End If
        End If

        'get the quantities of each row in seperate sql connection (so they can be editted via textbox)
        Try
            'open connection if it needs to be and execute query
            If conn.State = ConnectionState.Closed Then conn.Open()
            purchaserReader = purchaserSqlCmd.ExecuteReader()
            'loop through each row filling in quantity
            If GridView1.Rows.Count > 0 Then
                For row As Integer = 0 To GridView1.Rows.Count - 1
                    'find the textbox element of the current row
                    Dim txtPurchaser As TableCell = GridView1.Rows(row).Cells(5)
                    'move reader to next entry, record next entry, convert to string, format, and set to textbox
                    purchaserReader.Read()
                    Dim record As IDataRecord = CType(purchaserReader, IDataRecord)
                    txtPurchaser.Text = record(0).ToString
                Next
            End If
        Catch ex As Exception
            tbWarning.Text = "There has been an error. Purchasers were not retrieved."
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to set "Units" values
    '-----------------------------------------------------------------------------------------
    Private Sub setUnits()
        'sql query for "quantity" column
        Dim unitsSqlString As String = SQLStrings.getUnitsSqlString

        'setup command with sql server and parameterize PONumber input
        Dim unitsSqlCmd As New SqlCommand(unitsSqlString, conn)
        Dim param As New SqlParameter("@PONumber", SqlDbType.NVarChar, 20)
        param.Value = POsearch.Text.ToUpper
        unitsSqlCmd.Parameters.Add(param)

        'close the main reader if it exists and is open
        If Not IsNothing(purchaserReader) Then
            If Not purchaserReader.IsClosed() Then
                purchaserReader.Close()
            End If
        End If

        'get the quantities of each row in seperate sql connection (so they can be editted via textbox)
        Try
            'open connection if it needs to be and execute query
            If conn.State = ConnectionState.Closed Then conn.Open()
            unitsReader = unitsSqlCmd.ExecuteReader()
            'loop through each row filling in quantity
            If GridView1.Rows.Count > 0 Then
                For row As Integer = 0 To GridView1.Rows.Count - 1
                    'find the textbox element of the current row
                    Dim txtUnits As TextBox = GridView1.Rows(row).Cells(2).FindControl("unitsTextBox")
                    'move reader to next entry, record next entry, convert to string, format, and set to textbox
                    unitsReader.Read()
                    Dim record As IDataRecord = CType(unitsReader, IDataRecord)
                    txtUnits.Text = record(0).ToString
                Next
            End If
        Catch ex As Exception
            tbWarning.Text = "There has been an error. Units were not retrieved."
        End Try
    End Sub

    '-----------------------------------------------------------------------------------------
    ' Sub to handle click event of Print button. Calls helper functions and print subroutine
    '-----------------------------------------------------------------------------------------
    Protected Sub printbutton_Click(sender As Object, e As EventArgs) Handles printbutton.Click
        tbWarning.Text = ""
        'check that all "Copies" fields are numeric and non-zero and "Quantity" fields are numeric
        '   use AndAlso rather than And to short circuit 
        '   (first false from left to right and the rest will not be evaluated)
        If CopiesNumericCheck() AndAlso CopiesZeroCheck() AndAlso QuantityNumericCheck() Then
            'print out labels
            PrintLabels()
        End If
    End Sub

    '-------------------------------------------------------------------------------------------
    ' Function to check the values of "Copies" fields to ensure they are not all zeros
    '-------------------------------------------------------------------------------------------
    Private Function CopiesZeroCheck() As Boolean
        'find the number of rows in the GridView for looping
        Dim numOfRows As Integer = GridView1.Rows.Count

        'check to make sure a "Quantity" exists
        For rowIndex As Integer = 0 To numOfRows - 1
            'find the number of copies entered by the user
            Dim rowCells As TableCellCollection = GridView1.Rows(rowIndex).Cells()
            Dim smallTextBox As TextBox = rowCells.Item(6).FindControl("smallLabelsTextBox")
            Dim largeTextBox As TextBox = rowCells.Item(7).FindControl("largeLabelsTextBox")
            Dim smallStrCopies As String = smallTextBox.Text
            Dim largeStrCopies As String = largeTextBox.Text

            'make sure at least one text box is a non-zero
            If Not Convert.ToInt16(smallStrCopies) = 0 Then
                Return True
            End If

            'make sure at least one text box is a non-zero
            If Not Convert.ToInt16(largeStrCopies) = 0 Then
                Return True
            End If
        Next
        'if one text box is not non-zero, warn user
        tbWarning.Text = "You must enter at least one copy to print"
        Return False
    End Function

    '-------------------------------------------------------------------------------------------
    ' Function to check the values of "Copies" fields to ensure they are all numerical values
    '-------------------------------------------------------------------------------------------
    Private Function CopiesNumericCheck() As Boolean
        'find the number of rows in the GridView for looping
        Dim numOfRows As Integer = GridView1.Rows.Count

        'check to make sure all field have numeric "Label Copies" amounts
        For rowIndex As Integer = 0 To numOfRows - 1
            'find the number of copies entered by the user
            Dim rowCells As TableCellCollection = GridView1.Rows(rowIndex).Cells()
            Dim smallTextBox As TextBox = rowCells.Item(6).FindControl("smallLabelsTextBox")
            Dim largeTextBox As TextBox = rowCells.Item(7).FindControl("largeLabelsTextBox")
            Dim smallStrCopies As String = smallTextBox.Text
            Dim largeStrCopies As String = largeTextBox.Text

            'make sure number of small copies entered by user is numeric
            If Not IsNumeric(smallStrCopies) Then
                tbWarning.Text = "One of your ""Small Labels"" fields contains a non-numeric value"
                Return False
            End If

            'make sure number of large copies entered by user is numeric
            If Not IsNumeric(largeStrCopies) Then
                tbWarning.Text = "One of your ""Large Labels"" fields contains a non-numeric value"
                Return False
            End If
        Next
        Return True
    End Function

    '-------------------------------------------------------------------------------------------
    ' Function to check the values of "Quantity" fields to ensure they are all numerical values
    '-------------------------------------------------------------------------------------------
    Private Function QuantityNumericCheck() As Boolean
        'find the number of rows in the GridView for looping
        Dim numOfRows As Integer = GridView1.Rows.Count

        'check to make sure all field have numeric "Quantity" amounts
        For rowIndex As Integer = 0 To numOfRows - 1
            'find the quantity entered by the user
            Dim rowCells As TableCellCollection = GridView1.Rows(rowIndex).Cells()
            Dim textbox As TextBox = rowCells.Item(2).FindControl("quantityTextBox")
            Dim strQuantity As String = textbox.Text

            'make sure quantity entered by user is numeric
            If Not IsNumeric(strQuantity) Then
                tbWarning.Text = "One of your ""Quantity"" fields contains a non-numeric value"
                Return False
            End If
        Next
        Return True
    End Function

    '-------------------------------------------------------------------------------------------
    ' Sub to call small/large prints subroutines and zero out quantity fields
    '-------------------------------------------------------------------------------------------
    Private Sub PrintLabels()
        Dim numOfRows As Integer = GridView1.Rows.Count
        PrintSmallLabels(numOfRows)
        PrintLargeLabels(numOfRows)
        ZeroOutQuantityFields(numOfRows)
    End Sub


    '-------------------------------------------------------------------------------------------
    ' Sub to loop through small label quantities and print them
    '-------------------------------------------------------------------------------------------
    Private Sub PrintSmallLabels(ByVal numOfRows As Integer)
        'loop through rows and print labels
        For rowIndex As Integer = 0 To numOfRows - 1
            'find the number of copies entered by the user
            Dim rowCells As TableCellCollection = GridView1.Rows(rowIndex).Cells()
            Dim smallTextBox As TextBox = rowCells.Item(6).FindControl("smallLabelsTextBox")

            'convert no. of small copies to int and print labels
            Dim smallCopies As Integer = Convert.ToInt16(smallTextBox.Text)
            If smallCopies > 0 Then

                'grab data from selected row
                Dim partNo As String = rowCells.Item(0).Text
                Dim quantityTxtBox As TextBox = rowCells.Item(2).FindControl("quantityTextBox")
                Dim quantity As String = quantityTxtBox.Text
                Dim unitsTxtBox As TextBox = rowCells.Item(3).FindControl("unitsTextBox")
                Dim units As String = unitsTxtBox.Text

                'clean up strings if necessary ("&nbsp;" will exist if cell is empty)
                If partNo = "&nbsp;" Then partNo = ""

                'print info
                Dim btEngine As New Engine(True)
                Dim labelFormat As LabelFormatDocument = btEngine.Documents.Open("C:\Users\sgatto\Documents\BarTender\BarTender Documents\1x2.btw")

                'fill in designated fields on label
                labelFormat.SubStrings.SetSubString("part_no", partNo)
                labelFormat.SubStrings.SetSubString("quantity", quantity & " " & units)

                'print however many copies desired
                labelFormat.PrintSetup.IdenticalCopiesOfLabel = smallCopies
                labelFormat.PrintSetup.PrinterName = dropdownlist.SelectedItem.ToString
                labelFormat.Print()

                'close print engine
                btEngine.Stop()
            End If
        Next
    End Sub

    '-------------------------------------------------------------------------------------------
    ' Sub to loop through large label quantities and print them
    '-------------------------------------------------------------------------------------------
    Private Sub PrintLargeLabels(ByVal numOfRows As Integer)
        'loop through rows and print labels
        For rowIndex As Integer = 0 To numOfRows - 1
            'find the number of copies entered by the user
            Dim rowCells As TableCellCollection = GridView1.Rows(rowIndex).Cells()
            Dim largeTextBox As TextBox = rowCells.Item(7).FindControl("largeLabelsTextBox")
            'convert no. of large copies to int and print labels
            Dim largeCopies As Integer = Convert.ToInt16(largeTextBox.Text)
            If largeCopies > 0 Then

                'grab data from selected row
                Dim partNo As String = rowCells.Item(0).Text
                Dim desc As String = rowCells.Item(1).Text
                Dim quantityTxtBox As TextBox = rowCells.Item(2).FindControl("quantityTextBox")
                Dim quantity As String = quantityTxtBox.Text
                Dim unitsTxtBox As TextBox = rowCells.Item(3).FindControl("unitsTextBox")
                Dim units As String = unitsTxtBox.Text
                Dim location As String = rowCells.Item(4).Text
                Dim purchaser As String = rowCells.Item(5).Text

                'clean up strings if necessary ("&nbsp;" will exist if cell is empty)
                If partNo = "&nbsp;" Then partNo = ""
                If desc = "&nbsp;" Then desc = ""
                If location = "&nbsp;" Then location = ""
                If purchaser = "&nbsp;" Then purchaser = ""

                'print info
                Dim btEngine As New Engine(True)
                Dim labelFormat As LabelFormatDocument = btEngine.Documents.Open("C:\Users\sgatto\Documents\BarTender\BarTender Documents\4x2.btw")

                'fill in designated fields on label
                labelFormat.SubStrings.SetSubString("part_no", partNo)
                labelFormat.SubStrings.SetSubString("quantity", quantity & " " & units)
                labelFormat.SubStrings.SetSubString("description", desc)
                labelFormat.SubStrings.SetSubString("location", location)
                labelFormat.SubStrings.SetSubString("purchaser", purchaser)

                'print however many copies desired
                labelFormat.PrintSetup.IdenticalCopiesOfLabel = largeCopies
                labelFormat.PrintSetup.PrinterName = dropdownlist.SelectedItem.ToString
                labelFormat.Print()

                'close print engine
                btEngine.Stop()
            End If
        Next
    End Sub

    '-------------------------------------------------------------------------------------------
    ' Sub to zero out small and large label quantity fields
    '-------------------------------------------------------------------------------------------
    Private Sub ZeroOutQuantityFields(ByVal numOfRows As Integer)
        'zero out all quantities
        For rowIndex As Integer = 0 To numOfRows - 1
            'find the number of copies entered by the user
            Dim rowCells As TableCellCollection = GridView1.Rows(rowIndex).Cells()
            Dim smallTextBox As TextBox = rowCells.Item(6).FindControl("smallLabelsTextBox")
            Dim largeTextBox As TextBox = rowCells.Item(6).FindControl("largeLabelsTextBox")
            smallTextBox.Text = "0"
            largeTextBox.Text = "0"
        Next
    End Sub
End Class