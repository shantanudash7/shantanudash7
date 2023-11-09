Enabling Developer Mode and adding a VBA module in Excel is a simple process. Here's a step-by-step guide for a new user:

**Step 1: Open Excel**
- Open the given excel file.

**Step 2: Enable Developer Tab**
1. Click on the "File" tab in the upper left corner to open the File menu.

2. Select "Options" at the bottom of the menu. This will open the Excel Options dialog box.

3. In the Excel Options dialog, select "Customize Ribbon" on the left sidebar.

4. On the right side, you'll see a list of Main Tabs. Check the "Developer" option to enable the Developer tab.

5. Click "OK" to save your changes.

**Step 3: Access the Developer Tab**
- The Developer tab should now be visible in the Excel ribbon, usually located at the top of the Excel window. Click on the Developer tab to access its features.

**Step 4: Add a VBA Module**
1. Within the Developer tab, you'll find a "Visual Basic" button in the "Code" group. Click on it to open the Visual Basic for Applications (VBA) editor.

2. In the VBA editor, you'll see a Project Explorer pane on the left and a Code window in the center.

3. To add a new module, right-click on your workbook in the Project Explorer and select "Insert" -> "Module."

4. A new module will appear in the Project Explorer with a default name like "Module1." You can double-click on it to open the Code window for that module.

**Step 5: Write VBA Code**
- Now you can write or paste your VBA code into the Code window of the module.

```

Sub ScrapeAttend()
    Dim IE As Object
    Dim html As Object
    Dim i As Integer
    Dim ws As Worksheet
    Dim dataSheet As Worksheet
    Dim lastRow As Long
    Dim username As String
    Dim password As String
    Dim coid As String
    Dim url As String

    ' Create a new Internet Explorer instance
    Set IE = CreateObject("InternetExplorer.Application")
    IE.Visible = True

    ' Set a reference to the "Credentials" sheet
    Set ws = ThisWorkbook.Sheets("Credentials")

    ' Set a reference to the "Data" sheet
    Set dataSheet = ThisWorkbook.Sheets("Data")

    ' Get the last row in column C (where the coid values are stored)
    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).row

    ' Read the username and password from the "Credentials" sheet
    username = ws.Range("B1").Value
    password = ws.Range("B2").Value

    ' Navigate to the login page
    IE.navigate "https://acad.xlri.ac.in/"

    ' Wait for the login page to load
    Do While IE.Busy Or IE.readyState <> 4
        Application.Wait DateAdd("s", 1, Now)
    Loop

    ' Fill in the username and password and click the login button
    IE.document.getElementById("uid").Value = username
    IE.document.getElementById("pwd").Value = password
    IE.document.forms(0).submit

    ' Wait for the next page to load (you may need to customize this based on the website's behavior)
    Do While IE.Busy Or IE.readyState <> 4
        Application.Wait DateAdd("s", 1, Now)
    Loop

    ' Loop through the list of coid values
    For i = 2 To lastRow ' Assuming row 1 contains headers
        coid = ws.Cells(i, 3).Value ' Assuming coid values are in column C

        ' Construct the URL for the attendance page with the current coid
        url = "https://acad.xlri.ac.in/ais/attendance/AttendStu4Cou.php?coid=" & coid & "&&mcoid=" & coid

        ' Navigate to the attendance page with the current coid
        IE.navigate url

        ' Wait for the attendance page to load
        Do While IE.Busy Or IE.readyState <> 4
            Application.Wait DateAdd("s", 1, Now)
        Loop

        ' Get the HTML document
        Set html = IE.document

        ' Find and copy the table data
        Dim table As Object
        Set table = html.getElementById("datatables")

        ' Copy the table data to the "Data" sheet, adding to a new row for each coid
        Dim row As Object
        Dim col As Object
        Dim rowIndex As Long
        rowIndex = dataSheet.Cells(dataSheet.Rows.Count, "A").End(xlUp).row + 1

        ' Write the coid value in the first column of the row
        dataSheet.Cells(rowIndex + 1, 1).Value = coid

        ' Start writing data in the second column
        Dim dataColumn As Integer
        dataColumn = 2

        For Each row In table.getElementsByTagName("tr")
            Dim columnIndex As Integer
            columnIndex = 1

            For Each col In row.getElementsByTagName("td")
                dataSheet.Cells(rowIndex, dataColumn).Value = col.innerText
                dataColumn = dataColumn + 1
                columnIndex = columnIndex + 1
            Next col

            rowIndex = rowIndex + 1
            dataColumn = 2 ' Reset the data column for the next row
        Next row
    Next i

    ' Close IE
    IE.Quit
    Set IE = Nothing
End Sub



```

**Step 6: Save Your Workbook**
- Now save as your workbook as a macro-enabled file (usually with the ".xlsm" extension) if you want to retain the VBA code.

That's it! You've enabled Developer Mode and added a VBA module in Excel. You can start writing and running your VBA macros from the Developer tab.
