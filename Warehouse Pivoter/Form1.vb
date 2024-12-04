Imports OfficeOpenXml
Imports System.IO
Imports System.Text

Public Class Form1
    Private originalListViewItems As New List(Of ListViewItem) ' Store original ListView items for filtering
    Private locationColumnIndex As Integer = -1 ' Index for the "Location" column
    Private skuColumnIndex As Integer = -1 ' Index for the "SKU" column
    Private movableUnitLabelColumnIndex As Integer = -1 ' Index for the "Movable Unit Label" column

    Private Sub AppendDebugMessage(message As String, messageType As String)
        If rtbDebuggerOutput.InvokeRequired Then
            ' Use Invoke to update the control from the background thread
            rtbDebuggerOutput.Invoke(New Action(Of String, String)(AddressOf AppendDebugMessage), message, messageType)
        Else
            ' Update the control directly if on the UI thread
            Select Case messageType.ToLower()
                Case "info"
                    rtbDebuggerOutput.SelectionColor = Color.Blue
                Case "warning"
                    rtbDebuggerOutput.SelectionColor = Color.Orange
                Case "error"
                    rtbDebuggerOutput.SelectionColor = Color.Red
                Case Else
                    rtbDebuggerOutput.SelectionColor = Color.Black ' Default color
            End Select

            ' Append the message to the RichTextBox
            rtbDebuggerOutput.AppendText($"{DateTime.Now:HH:mm:ss} - {message}{Environment.NewLine}")
            rtbDebuggerOutput.SelectionColor = Color.Black ' Reset color for next message
        End If
    End Sub


    Private Sub OpenxlsxToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenxlsxToolStripMenuItem.Click
        ' Open File Dialog to select an Excel file
        Dim openFileDialog As New OpenFileDialog With {
        .Filter = "Excel Files|*.xlsx",
        .Title = "Select an Excel File"
    }

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            Dim filePath As String = openFileDialog.FileName

            ' Log the selected file path
            AppendDebugMessage($"Selected file: {filePath}", "info")

            ' Clear ListView1 and all combo boxes before loading new data
            ListView1.Clear()
            cbLocation.Items.Clear()
            cbSKU.Items.Clear()
            cbMovableUnitLabel.Items.Clear()
            originalListViewItems.Clear()
            locationColumnIndex = -1
            skuColumnIndex = -1
            movableUnitLabelColumnIndex = -1 ' Reset column indexes

            ' Load Excel file using EPPlus
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial ' Required for EPPlus

            ' Initialize ProgressBar1
            ProgressBar1.Value = 0
            ProgressBar1.Visible = True
            ProgressBar1.Style = ProgressBarStyle.Blocks

            Try
                Using package As New ExcelPackage(New FileInfo(filePath))
                    Dim worksheet = package.Workbook.Worksheets(0) ' Load the first sheet
                    AppendDebugMessage($"Loaded worksheet: {worksheet.Name}", "info")

                    ' Load headers (first row)
                    Dim headerRow As Integer = 1
                    For col = 1 To worksheet.Dimension.Columns
                        Dim headerName As String = worksheet.Cells(headerRow, col).Text
                        ListView1.Columns.Add(headerName)

                        ' Identify important columns
                        Select Case headerName.ToLower()
                            Case "location"
                                locationColumnIndex = col - 1 ' Save index (0-based for ListView)
                            Case "sku"
                                skuColumnIndex = col - 1
                            Case "movable unit label"
                                movableUnitLabelColumnIndex = col - 1
                        End Select
                    Next

                    ' Ensure all required columns are found
                    If locationColumnIndex = -1 Or skuColumnIndex = -1 Or movableUnitLabelColumnIndex = -1 Then
                        AppendDebugMessage("Missing required columns (Location, SKU, Movable Unit Label).", "error")
                        ProgressBar1.Visible = False
                        Return
                    End If

                    AppendDebugMessage("Successfully loaded headers.", "info")

                    ' Adjust ListView columns to fit content
                    ListView1.View = View.Details
                    ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)

                    ' Prepare to load rows
                    Dim totalRows As Integer = worksheet.Dimension.Rows - 1 ' Exclude the header row
                    ProgressBar1.Maximum = totalRows

                    Dim locations As New HashSet(Of String)
                    Dim skus As New HashSet(Of String)
                    Dim movableUnitLabels As New HashSet(Of String)

                    ' Load rows (starting from the second row)
                    For row = 2 To worksheet.Dimension.Rows
                        Dim listViewItem As New ListViewItem()
                        For col = 1 To worksheet.Dimension.Columns
                            Dim cellValue As String = worksheet.Cells(row, col).Text
                            If col = 1 Then
                                listViewItem.Text = cellValue ' First column as ListView item text
                            Else
                                listViewItem.SubItems.Add(cellValue)
                            End If

                            ' Add unique values to the corresponding sets
                            If col - 1 = locationColumnIndex AndAlso Not String.IsNullOrWhiteSpace(cellValue) Then
                                locations.Add(cellValue)
                            End If
                            If col - 1 = skuColumnIndex AndAlso Not String.IsNullOrWhiteSpace(cellValue) Then
                                skus.Add(cellValue)
                            End If
                            If col - 1 = movableUnitLabelColumnIndex AndAlso Not String.IsNullOrWhiteSpace(cellValue) Then
                                movableUnitLabels.Add(cellValue)
                            End If
                        Next
                        ListView1.Items.Add(listViewItem)
                        originalListViewItems.Add(listViewItem)

                        ' Update ProgressBar
                        ProgressBar1.Value = row - 1
                    Next

                    ' Populate combo boxes with unique values
                    cbLocation.Items.AddRange(locations.ToArray())
                    cbSKU.Items.AddRange(skus.ToArray())
                    cbMovableUnitLabel.Items.AddRange(movableUnitLabels.ToArray())

                    ' Enable auto-complete for combo boxes
                    cbLocation.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    cbLocation.AutoCompleteSource = AutoCompleteSource.ListItems
                    cbSKU.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    cbSKU.AutoCompleteSource = AutoCompleteSource.ListItems
                    cbMovableUnitLabel.AutoCompleteMode = AutoCompleteMode.SuggestAppend
                    cbMovableUnitLabel.AutoCompleteSource = AutoCompleteSource.ListItems

                    AppendDebugMessage($"Successfully loaded {originalListViewItems.Count} rows.", "info")
                End Using
            Catch ex As Exception
                AppendDebugMessage($"Error loading Excel file: {ex.Message}", "error")
            Finally
                ' Hide ProgressBar after completion
                ProgressBar1.Visible = False
            End Try
        End If
    End Sub







    Private Sub FilterListView()
        Dim locationFilter As String = cbLocation.Text.ToLower()
        Dim skuFilter As String = cbSKU.Text.ToLower()
        Dim movableUnitLabelFilter As String = cbMovableUnitLabel.Text.ToLower()

        AppendDebugMessage($"Filtering with Location: {locationFilter}, SKU: {skuFilter}, Movable Unit Label: {movableUnitLabelFilter}", "info")

        ' Clear ListView1 and re-add filtered items
        ListView1.Items.Clear()

        If String.IsNullOrWhiteSpace(locationFilter) AndAlso String.IsNullOrWhiteSpace(skuFilter) AndAlso String.IsNullOrWhiteSpace(movableUnitLabelFilter) Then
            ' If no filter, restore original items
            ListView1.Items.AddRange(originalListViewItems.ToArray())
            AppendDebugMessage("Filter cleared. Showing all items.", "info")
        Else
            ' Filter items based on conditions
            Dim filteredCount As Integer = 0
            For Each item As ListViewItem In originalListViewItems
                Dim matchesLocation = String.IsNullOrWhiteSpace(locationFilter) OrElse item.SubItems(locationColumnIndex).Text.ToLower().Contains(locationFilter)
                Dim matchesSKU = String.IsNullOrWhiteSpace(skuFilter) OrElse item.SubItems(skuColumnIndex).Text.ToLower().Contains(skuFilter)
                Dim matchesMovableUnitLabel = String.IsNullOrWhiteSpace(movableUnitLabelFilter) OrElse item.SubItems(movableUnitLabelColumnIndex).Text.ToLower().Contains(movableUnitLabelFilter)

                If matchesLocation AndAlso matchesSKU AndAlso matchesMovableUnitLabel Then
                    ListView1.Items.Add(item)
                    filteredCount += 1
                End If
            Next
            AppendDebugMessage($"Filtered {filteredCount} items based on filters.", "info")
        End If

        ' Auto-resize columns
        ListView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)
    End Sub





    '''''''''''''''''''''''''''



    Private Sub btnExportLocation_Click(sender As Object, e As EventArgs) Handles btnExportLocation.Click
        ' Get the prefix from txtLocationFilter
        Dim prefix As String = txtLocationFilter.Text.Trim().ToUpper()

        ' Validate the input
        If String.IsNullOrWhiteSpace(prefix) Then
            MessageBox.Show("Please enter a valid location prefix to filter (e.g., A1, B2).", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' Ask user to choose a style: Basic or Modern
        Dim styleChoice As DialogResult = MessageBox.Show("Choose the export style: Yes for Basic, No for Modern.", "Select Export Style", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

        ' Handle user's choice
        Dim useModernStyle As Boolean
        If styleChoice = DialogResult.Yes Then
            useModernStyle = False ' Basic style
        ElseIf styleChoice = DialogResult.No Then
            useModernStyle = True ' Modern style
        Else
            AppendDebugMessage("Export canceled by user.", "info")
            Return
        End If

        AppendDebugMessage($"Filtering locations starting with '{prefix}' and exporting...", "info")

        ' Clear ListView2 before adding new results
        ListView2.Clear()

        ' Add columns to ListView2
        ListView2.Columns.Add("Location")
        ListView2.Columns.Add("SKU")
        ListView2.Columns.Add("Description")
        ListView2.Columns.Add("Total Pallets")
        ListView2.Columns.Add("Action") ' New column for manual input
        ListView2.View = View.Details

        ' Dictionary to store aggregated data for each location
        Dim locationData As New SortedDictionary(Of String, (TotalPallets As Integer, SKU As String, Description As String))

        ' Process rows in ListView1: Aggregate Total Pallets by location
        Dim dataAdded As Boolean = False
        For Each item As ListViewItem In originalListViewItems
            Dim location As String = item.SubItems(locationColumnIndex).Text.Trim().ToUpper() ' Force uppercase
            Dim normalizedLocation As String = location.Replace("-", "") ' Remove hyphens for normalization

            ' Check if the normalized location starts with the user-provided prefix
            If normalizedLocation.StartsWith(prefix) Then
                Dim sku As String = item.SubItems(skuColumnIndex).Text
                Dim descriptionColumnIndex = GetColumnIndex("Description") ' Dynamically detect the "Description" column
                Dim description As String = item.SubItems(descriptionColumnIndex).Text.Trim()

                AppendDebugMessage($"Processing Location: {location}, SKU: {sku}", "info")

                ' Aggregate Total Pallets for the location
                If locationData.ContainsKey(location) Then
                    Dim updatedTotalPallets = locationData(location).TotalPallets + 1
                    locationData(location) = (updatedTotalPallets, sku, description)

                    AppendDebugMessage($"Updated Location: {location}, Total Pallets: {updatedTotalPallets}", "info")
                Else
                    locationData(location) = (1, sku, description) ' Initialize counts for a new location
                    AppendDebugMessage($"New Location: {location}, Total Pallets: 1", "info")
                End If
                dataAdded = True
            End If
        Next

        ' Add aggregated and sorted data to ListView2
        For Each kvp In locationData
            Dim location = kvp.Key
            Dim totalPallets = kvp.Value.TotalPallets
            Dim sku = kvp.Value.SKU
            Dim description = kvp.Value.Description

            ' Add row to ListView2
            Dim listViewItem As New ListViewItem(location)
            listViewItem.SubItems.Add(sku)
            listViewItem.SubItems.Add(description)
            listViewItem.SubItems.Add(totalPallets.ToString()) ' Aggregated Total Pallets
            listViewItem.SubItems.Add("") ' Empty Action column
            ListView2.Items.Add(listViewItem)
        Next

        If dataAdded Then
            AppendDebugMessage($"Data successfully filtered for '{prefix}' and displayed in ListView2.", "info")
            ListView2.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent)

            ' Trigger HTML export
            ExportToHTML(prefix, useModernStyle)
        Else
            AppendDebugMessage($"No data matched the filter criteria for locations starting with '{prefix}'.", "warning")
            MessageBox.Show($"No data found for locations starting with '{prefix}'.", "No Results", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub


    Private Function GetColumnIndex(columnName As String) As Integer
        ' Normalize the column name and trim spaces
        columnName = columnName.Trim().ToLower()
        For i As Integer = 0 To ListView1.Columns.Count - 1
            If ListView1.Columns(i).Text.Trim().ToLower() = columnName Then
                Return i
            End If
        Next
        Throw New Exception($"Column '{columnName}' not found.")
    End Function





    Private Sub ExportToHTML(prefix As String, useModernStyle As Boolean)
        Dim desktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
        Dim filePath As String = $"{desktopPath}\{prefix}_BulletproofLogisticsCycleCounting.html"

        Dim htmlContent As New StringBuilder()

        ' Add HTML structure
        htmlContent.AppendLine("<!DOCTYPE html>")
        htmlContent.AppendLine("<html lang='en'>")
        htmlContent.AppendLine("<head>")
        htmlContent.AppendLine("<meta charset='UTF-8'>")
        htmlContent.AppendLine("<meta name='viewport' content='width=device-width, initial-scale=1.0'>")
        htmlContent.AppendLine($"<title>Bulletproof Logistics Cycle Counting - {prefix}</title>")
        htmlContent.AppendLine("<style>")

        If useModernStyle Then
            ' Modern style
            htmlContent.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }")
            htmlContent.AppendLine("h1 { text-align: center; margin-bottom: 20px; }")
            htmlContent.AppendLine("table { width: 100%; border-collapse: collapse; margin: 20px 0; font-size: 12px; }")
            htmlContent.AppendLine("table, th, td { border: 1px solid #ddd; }")
            htmlContent.AppendLine("th { background-color: #4CAF50; color: white; padding: 8px; }")
            htmlContent.AppendLine("td { padding: 8px; text-align: left; }")
            htmlContent.AppendLine("@media print { table { page-break-inside: auto; } }")
        Else
            ' Basic style
            htmlContent.AppendLine("body { font-family: Arial, sans-serif; margin: 20px; }")
            htmlContent.AppendLine("h1 { text-align: center; }")
            htmlContent.AppendLine("table { width: 100%; border-collapse: collapse; margin: 20px 0; }")
            htmlContent.AppendLine("table, th, td { border: 1px solid black; }")
            htmlContent.AppendLine("th, td { padding: 10px; text-align: left; }")
            htmlContent.AppendLine("th { background-color: #f2f2f2; }")
        End If

        htmlContent.AppendLine("</style>")
        htmlContent.AppendLine("</head>")
        htmlContent.AppendLine("<body>")

        ' Add title and date/time
        htmlContent.AppendLine($"<h1>Bulletproof Logistics Cycle Counting - {prefix}</h1>")
        htmlContent.AppendLine($"<p>Date and Time: {DateTime.Now}</p>")

        ' Add table
        htmlContent.AppendLine("<table>")
        htmlContent.AppendLine("<tr><th>Location</th><th>SKU</th><th>Description</th><th>Total Pallets</th><th>Action</th></tr>")

        ' Populate table rows
        For Each item As ListViewItem In ListView2.Items
            htmlContent.AppendLine("<tr>")
            For Each subItem As ListViewItem.ListViewSubItem In item.SubItems
                htmlContent.AppendLine($"<td>{subItem.Text}</td>")
            Next
            htmlContent.AppendLine("</tr>")
        Next

        htmlContent.AppendLine("</table>")
        htmlContent.AppendLine("</body>")
        htmlContent.AppendLine("</html>")

        ' Write HTML file
        File.WriteAllText(filePath, htmlContent.ToString())
        AppendDebugMessage($"HTML exported to {filePath}", "info")

        ' Open HTML file in default browser
        Process.Start(New ProcessStartInfo(filePath) With {.UseShellExecute = True})
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Initialize Timer
        Timer1.Interval = 2000 ' Set delay to 2 seconds
        Timer1.Enabled = False
    End Sub

    Private Sub cbLocation_TextChanged(sender As Object, e As EventArgs) Handles cbLocation.TextChanged
        Timer1.Stop() ' Stop the timer to reset the delay
        Timer1.Start() ' Restart the timer
    End Sub

    Private Sub cbSKU_TextChanged(sender As Object, e As EventArgs) Handles cbSKU.TextChanged
        Timer1.Stop() ' Stop the timer to reset the delay
        Timer1.Start() ' Restart the timer
    End Sub

    Private Sub cbMovableUnitLabel_TextChanged(sender As Object, e As EventArgs) Handles cbMovableUnitLabel.TextChanged
        Timer1.Stop() ' Stop the timer to reset the delay
        Timer1.Start() ' Restart the timer
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Timer1.Stop() ' Stop the timer to prevent repeated execution
        FilterListView() ' Perform the search logic
    End Sub



End Class
