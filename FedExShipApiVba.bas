Attribute VB_Name = "Module1"
Dim apiKey As String
Dim apiPassword As String
Dim accountNumber As String
Dim accessToken As String
Dim nextRow As Long
Dim isStreetLine2 As Boolean
Dim isPR As Boolean
Dim isSD As Boolean
    Dim wsMacros As Worksheet
    Dim folderPath As String
        Dim wbCSV As Workbook
        Dim poNumber As String
        Dim trackingNumber As String
        Dim deliveryDate As String

Sub Main()

    ' Set reference to the Macros sheet
    Set wsMacros = ThisWorkbook.Sheets("BASE BEFORE")
    Call initialize
    Call CheckDateAndSelectFolder
    Call ProcessCSVFiles
    

'ask user for folder to save pdfs at start. then put a check if the global file path variable is empty, if it is ask, otherwise just run macro



End Sub


Sub initialize()
    apiKey = "your api key"
    
'    Dim apiPassword As String
    apiPassword = "your api password"
    
'    Dim accountNumber As String
    accountNumber = "your account number"
    
    ' Access token
'    Dim accessToken As String
    accessToken = GetAccessToken(apiKey, apiPassword)
    
    ' Check if access token is retrieved successfully
    If accessToken = "" Then
        MsgBox "Failed to retrieve access token."
        Exit Sub
    End If



End Sub

Sub CreateShipment()



'nextRow = 4709
    ' Set reference to the Macros sheet
'    Set wsMacros = ThisWorkbook.Sheets("BASE BEFORE")



Dim newValue As String
    'temp value to store user inputs.
    
    ' API credentials
'    Dim apiKey As String
'    apiKey = "l766ccd1dbf3db468e92772ab6961b961c"
    
'    Dim apiPassword As String
'    apiPassword = "16b30d3b46644747a3c7e1264a56aec4"
    
'    Dim accountNumber As String
'    accountNumber = "740561073"
    
    ' Access token
'    Dim accessToken As String
'    accessToken = GetAccessToken(apiKey, apiPassword)
    
    ' Check if access token is retrieved successfully
'    If accessToken = "" Then
'        MsgBox "Failed to retrieve access token."
'        Exit Sub
'    End If
    
    ' Sender information
    Dim senderName As String
    senderName = "DONGYING BAOFENG AUTO FITTING"
    
    Dim senderPhoneNumber As String
    senderPhoneNumber = "3147335490"
    
    Dim senderCompanyName As String
    senderCompanyName = "REACTION AUTO PARTS, INC"
    
    Dim senderStreetLine1 As String
    senderStreetLine1 = "7031 PREMIER PKWY"
    
    Dim senderCity As String
    senderCity = "ST PETERS"
    
    Dim senderStateCode As String
    senderStateCode = "MO"
    
    Dim senderPostalCode As String
    senderPostalCode = "63376"
    
    Dim senderCountryCode As String
    senderCountryCode = "US"
    
    ' Recipient information
    Dim recipientName As String
    recipientName = wsMacros.Range("N" & nextRow).value
    
    Dim recipientPhoneNumber As String
    recipientPhoneNumber = wsMacros.Range("P" & nextRow).value
    
    Dim invoiceNo As String
    invoiceNo = wsMacros.Range("R" & nextRow).value
    
    poNumber = wsMacros.Range("C" & nextRow).value
    
    ' Loop until the user enters a value
    Do While recipientPhoneNumber = "0"
        ' Prompt the user for a new value
        recipientPhoneNumber = InputBox("Enter phone number for " & wsMacros.Range("D" & nextRow).value, "Enter")
        
        ' Check if user entered a value
        If recipientPhoneNumber <> "" Then
        Else
            ' Notify the user that a value is required
            MsgBox "Please enter phone number.", vbExclamation
        End If
    Loop
    
'    Dim recipientCompanyName As String
'    recipientCompanyName = ThisWorkbook.Sheets("Sheet1").Range("B4").value
    
    Dim recipientStreetLine1 As String
    recipientStreetLine1 = wsMacros.Range("H" & nextRow).value
    
    Dim recipientStreetLine2 As String
    
    
    If wsMacros.Range("I" & nextRow).value = "" Then
        isStreetLine2 = False
    Else
        isStreetLine2 = True
        recipientStreetLine2 = wsMacros.Range("I" & nextRow).value
    End If
    
    Dim recipientCity As String
    recipientCity = wsMacros.Range("J" & nextRow).value
    
    Dim recipientStateCode As String
    recipientStateCode = wsMacros.Range("K" & nextRow).value
    If recipientStateCode = "PR" Then isPR = True
    
    
    Dim recipientPostalCode As String
    recipientPostalCode = wsMacros.Range("L" & nextRow).value
    
    Dim recipientCountryCode As String
    recipientCountryCode = "US"
    
    ' Package information
    Dim weightValue As Double
    weightValue = wsMacros.Range("V" & nextRow).value
    
    ' Loop until the user enters a value
    Do While weightValue = 0
        ' Prompt the user for a new value
        newValue = InputBox("Enter weight for " & wsMacros.Range("D" & nextRow).value, "Enter")
        
        ' Check if user entered a value
        If newValue <> "" Then
            ' Replace A1 value with the user's input
            weightValue = newValue * wsMacros.Range("F" & nextRow).value
        Else
            ' Notify the user that a value is required
            MsgBox "Please enter weight.", vbExclamation
        End If
    Loop
    
    If wsMacros.Range("Z" & nextRow).value = "SD" Then isSD
    
    
    
    
    ' Construct the JSON payload
    Dim jsonPayload As String

   
    jsonPayload = "{"
    jsonPayload = jsonPayload & """labelResponseOptions"": ""URL_ONLY"","
    jsonPayload = jsonPayload & """requestedShipment"": {"
    jsonPayload = jsonPayload & """shipper"": {"
    jsonPayload = jsonPayload & """contact"": {"
    jsonPayload = jsonPayload & """personName"": """ & senderName & ""","
    'jsonPayload = jsonPayload & """phoneNumber"": """ & senderPhoneNumber & ""","
    
    jsonPayload = jsonPayload & """phoneNumber"": " & senderPhoneNumber & ","
    
    jsonPayload = jsonPayload & """companyName"": """ & senderCompanyName & """"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """address"": {"
    jsonPayload = jsonPayload & """streetLines"": [""" & senderStreetLine1 & """],"
    jsonPayload = jsonPayload & """city"": """ & senderCity & ""","
    jsonPayload = jsonPayload & """stateOrProvinceCode"": """ & senderStateCode & ""","
   ' jsonPayload = jsonPayload & """postalCode"": """ & senderPostalCode & ""","
   
   jsonPayload = jsonPayload & """postalCode"": " & senderPostalCode & ","
   
   
    jsonPayload = jsonPayload & """countryCode"": """ & senderCountryCode & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """recipients"": ["
    jsonPayload = jsonPayload & "{"
    jsonPayload = jsonPayload & """contact"": {"
    jsonPayload = jsonPayload & """personName"": """ & recipientName & ""","
    'jsonPayload = jsonPayload & """phoneNumber"": """ & recipientPhoneNumber & ""","
    
    jsonPayload = jsonPayload & """phoneNumber"": " & recipientPhoneNumber & ","
    
    jsonPayload = jsonPayload & """companyName"": """ & recipientCompanyName & """"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """address"": {"
'    jsonPayload = jsonPayload & """streetLines"": [""" & recipientStreetLine1 & """,""" & recipientStreetLine2 & """],"
    
    If isStreetLine2 Then
        jsonPayload = jsonPayload & """streetLines"": [""" & recipientStreetLine1 & """,""" & recipientStreetLine2 & """],"
    Else
        jsonPayload = jsonPayload & """streetLines"": [""" & recipientStreetLine1 & """],"
    End If
    
    jsonPayload = jsonPayload & """city"": """ & recipientCity & ""","
    jsonPayload = jsonPayload & """stateOrProvinceCode"": """ & recipientStateCode & ""","
'    jsonPayload = jsonPayload & """postalCode"": """ & recipientPostalCode & ""","
    
    jsonPayload = jsonPayload & """postalCode"": " & recipientPostalCode & ","
    
    jsonPayload = jsonPayload & """countryCode"": """ & recipientCountryCode & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "],"
    jsonPayload = jsonPayload & """shipDatestamp"": """ & Format(Date, "yyyy-mm-dd") & ""","
    jsonPayload = jsonPayload & """serviceType"": ""STANDARD_OVERNIGHT"","
    jsonPayload = jsonPayload & """packagingType"": ""FEDEX_BOX"","
    jsonPayload = jsonPayload & """pickupType"": ""USE_SCHEDULED_PICKUP"","
    jsonPayload = jsonPayload & """blockInsightVisibility"": false,"
    jsonPayload = jsonPayload & """shippingChargesPayment"": {"
    jsonPayload = jsonPayload & """paymentType"": ""SENDER"""
    jsonPayload = jsonPayload & "},"
    
    
    
    
'    jsonPayload = jsonPayload & """customsClearanceDetail"": {"
'
'    jsonPayload = jsonPayload & """commercialInvoice"": {"
'    jsonPayload = jsonPayload & """customerReferences"": ["
'    jsonPayload = jsonPayload & "{"
'    jsonPayload = jsonPayload & """customerReferenceType"": ""INVOICE_NUMBER"","
'    jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
'    jsonPayload = jsonPayload & "}"
'    jsonPayload = jsonPayload & "],"
'    jsonPayload = jsonPayload & "},"
'    jsonPayload = jsonPayload & "},"
    
    
    jsonPayload = jsonPayload & """labelSpecification"": {"
    jsonPayload = jsonPayload & """imageType"": ""PDF"","
    jsonPayload = jsonPayload & """labelStockType"": ""PAPER_85X11_TOP_HALF_LABEL"""
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """requestedPackageLineItems"": ["
    jsonPayload = jsonPayload & "{"
    
    
    jsonPayload = jsonPayload & """customerReferences"": ["
    jsonPayload = jsonPayload & "{"
'    jsonPayload = jsonPayload & """customerReferenceType"": ""INVOICE_NUMBER"","
    jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
    jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & "{"
    jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
    jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("M" & nextRow).value & """"
    
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "],"
    
    
    
    
  '  jsonPayload = jsonPayload & """phoneNumber"": " & recipientPhoneNumber & ","
    
    
    
    
    
    jsonPayload = jsonPayload & """weight"": {"
    jsonPayload = jsonPayload & """value"": " & weightValue & ","
    jsonPayload = jsonPayload & """units"": ""LB"""
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
'    jsonPayload = jsonPayload & "],"
    
    jsonPayload = jsonPayload & "]"
    jsonPayload = jsonPayload & "},"
    
    jsonPayload = jsonPayload & """accountNumber"": {"
    jsonPayload = jsonPayload & """value"": """ & accountNumber & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
    
    
    'jsonPayload = jsonPayload & "}"
    

   
   
   
   
   
   

    ' Make the API request
    Dim url As String
    url = "https://apis-sandbox.fedex.com/ship/v1/shipments"

    Dim http As Object
'    Set http = CreateObject("WinHttp.WinHttpRequest")
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "POST", url, False
    http.SetRequestHeader "Authorization", "Bearer " & accessToken
    http.SetRequestHeader "Content-Type", "application/json"
 '   http.SetRequestHeader "Accept", "application/json"

    http.Send jsonPayload

    ' Check if the request was successful
    If http.Status = 200 Then
        Dim responseJson As Object
        Set responseJson = JsonConverter.ParseJson(http.responseText)
        ThisWorkbook.Sheets("Sheet1").Range("A15").value = http.responseText
        'Dim trackingNumber As String

        trackingNumber = responseJson("output")("transactionShipments")(1)("masterTrackingNumber")
 '               ThisWorkbook.Sheets("Sheet1").Range("F8").value = trackingNumber
                
                
                
                
                deliveryDate = responseJson("output")("transactionShipments")(1)("pieceResponses")(1)("deliveryDatestamp")
                Range("I13") = deliveryDate
                
                
                
                
                
                
                

        ' Retrieve the label URL from the response
        Dim labelUrl As String
'        labelUrl = responseJson("labelUrl")
       ' labelUrl = responseJson("output")("transactionShipments")(1)("serviceType")
       ' labelUrl = responseJson("output")("transactionShipments")(1)("pieceResponses")(1)("masterTrackingNumber")
        labelUrl = responseJson("output")("transactionShipments")(1)("pieceResponses")(1)("packageDocuments")(1)("url")
        
        deliveryDate = Left(deliveryDate, 4) & Mid(deliveryDate, 6, 2) & Right(deliveryDate, 2)
        
        
        ' Download and save the label as PDF
        Dim labelFilePath As String
        labelFilePath = folderPath & poNumber & ".pdf"
  '      labelFilePath = "C:\Users\ntayl\Desktop\nick vdp macro test\05.03.24\" & wsMacros.Range("C" & nextRow).value & ".pdf"

        Dim labelHttp As Object
            
            Dim labelFile As Object
'    Set labelHttp = CreateObject("Microsoft.XMLHTTP")
'    labelHttp.Open "GET", labelUrl, False
'    labelHttp.Send
      
        
               
        Dim xmlHTTP As Object
    
'    Dim pdfFilePath As String
    
   
    ' Path to save the PDF file
 '   pdfFilePath = "C:\Users\ntayl\Desktop\nick vdp macro test\05.03.24\YourFile.pdf"
    
    ' Create a new XMLHTTP object
    Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
    
    ' Open the URL
    xmlHTTP.Open "GET", labelUrl, False
    
    ' Send the request
    xmlHTTP.Send
    
    ' Check if the request was successful
    If xmlHTTP.Status = 200 Then
        ' Create a new FileStream object to write the PDF content
        Dim fso As Object
        Set fso = CreateObject("Scripting.FileSystemObject")
        Dim stream As Object
        Set stream = fso.CreateTextFile(labelFilePath, True)
        
        ' Write the response content (webpage) to the FileStream
        stream.Write xmlHTTP.responseText
        
        ' Close the FileStream
        stream.Close
        
                    
            
            wsMacros.Range("A" & nextRow) = "good"
            MsgBox "Shipment created successfully. Label saved as PDF."
        Else
            MsgBox "Failed to retrieve label. Error: " & labelHttp.Status & " - " & labelHttp.StatusText
        End If
    Else
        MsgBox "Failed to create shipment. Error: " & http.Status & " - " & http.StatusText
    End If
    
        Set xmlHTTP = Nothing
End Sub

Private Function GetAccessToken(apiKey As String, apiPassword As String) As String
    Dim url As String
    'url = "https://apis-sandbox.fedex.com/auth/oauth/v2/token"
    url = "https://apis-sandbox.fedex.com/oauth/token"

    Dim http As Object
 '   Set http = CreateObject("WinHttp.WinHttpRequest")
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    
    
    
    http.Open "POST", url, False
    http.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"

    Dim requestBody As String
    requestBody = "grant_type=client_credentials&client_id=" & apiKey & "&client_secret=" & apiPassword

    http.Send requestBody

    If http.Status = 200 Then
        Dim responseJson As Object
        Set responseJson = JsonConverter.ParseJson(http.responseText)

        GetAccessToken = responseJson("access_token")
  '      Range("A1") = GetAccessToken
    Else
        GetAccessToken = ""
    End If
End Function
Sub AddressValidation()






End Sub
Sub ProcessCSVFiles()
 '   Dim folderPath As String
    Dim csvFile As String
 '   Dim wbCSV As Workbook
  '  Dim wsMacros As Worksheet
 '   Dim nextRow As Long
    
    ' Ask user to select a folder
'    With Application.FileDialog(msoFileDialogFolderPicker)
'        .Title = "Select a folder"
'        .Show
'        If .SelectedItems.Count = 0 Then Exit Sub
'        folderPath = .SelectedItems(1) & "\"
'    End With
    

    
    ' Find the next available row in column C
    nextRow = wsMacros.Cells(wsMacros.Rows.Count, "C").End(xlUp).Row + 1
    
    ' Loop through each CSV file in the selected folder
    csvFile = Dir(folderPath & "*.csv")
    Do While csvFile <> ""
        ' Open the CSV file
        Set wbCSV = Workbooks.Open(folderPath & csvFile)
        
        ' Check if D4 has a value
        If wbCSV.Sheets(1).Range("D4").value <> "" Then
            ' Close the file without saving changes
            wbCSV.Close False
        Else
            ' Copy B2 value to next available spot in row C of Macros sheet
            wsMacros.Cells(nextRow, "C").value = wbCSV.Sheets(1).Range("B2").value
            
            ' Refresh the Macros sheet
            ThisWorkbook.RefreshAll
            
            
            
            
            
            'put shipping macro in here
            
            Call CreateShipment
            
            
            
            
            
            
            ' Close the CSV file
            wbCSV.Close False
            
            ' Move to the next row in Macros sheet
            nextRow = nextRow + 1
        End If
        
        ' Move to the next CSV file
        csvFile = Dir
    Loop
End Sub

Sub CheckDateAndSelectFolder()
    Dim ws As Worksheet
    Dim selectedFolder As Variant
    Dim objShell As Object
    
    ' Set a reference to Sheet1
    Set ws = ThisWorkbook.Sheets("Sheet1")
    
    ' Check if cell A2 contains today's date
    If ws.Range("A2").value = Date Then
    folderPath = ws.Range("A1").value
 '       MsgBox "Today's date has already been recorded. Macro will stop."
        Exit Sub
    Else
        ' Create a Shell object to browse for folder
        Set objShell = CreateObject("Shell.Application")
        ' Prompt user to select a folder
        On Error Resume Next ' In case the user cancels the folder selection
        Set selectedFolder = objShell.BrowseForFolder(0, "Select a folder", 0)
        On Error GoTo 0 ' Reset error handling
        
        ' Check if user has canceled folder selection
        If selectedFolder Is Nothing Then
            MsgBox "Folder selection canceled. Macro will stop."
            End
        End If
        
        ' Save the selected folder path in cell A1 of Sheet1
        ws.Range("A1").value = selectedFolder.Items.Item.Path & "\"
  '      MsgBox "Folder path saved successfully."
        ws.Range("A2").value = Date
        folderPath = selectedFolder.Items.Item.Path & "\"
    End If
    
    
End Sub


Sub VDP_FORMAT()
'
' VDP_FORMAT Macro
'
' Keyboard Shortcut: Ctrl+Shift+A
'
' DIR VAR = DOUBLE
' DO THIS FOR DATE ETC
' TAKE PRICE FROM ORDER

'    Dim CurrentDate As String
'    Set CurrentDate = Format(Now(), "yyyymmdd")
    Dim SelectedRow As Range
    Dim OrderSheetName As String
    Dim VdpSheetName As String
    Dim TodaysDate As String
    Dim PoNum As String
    
    
'    wsMacros
    
    wbCSV.Sheets(1).Activate
    
    
    TodaysDate = Format(Date, "yyyymmdd")
    
   ' OrderSheetName = ActiveSheet.Name
   ' VdpSheetName = "VDP PO " + TodaysDate + ".xlsm"
    
    'PoNum = Range("B2").value
    
        Application.Calculation = xlCalculationManual
        
   ' Range("U2").Select
    Range("U2:V2").NumberFormat = "@"
    Range("U2:V2") = trackingNumber
   ' ActiveSheet.PasteSpecial Format:="HTML", Link:=False, DisplayAsIcon:= _
   '     False, NoHTMLFormatting:=True
   ' Range("V2").value = Range("U2").value
    Range("D4").value = Range("C4").value
    Range("P2").value = TodaysDate
    
    'Range("Q2") = "=IF(LEFTB(I2,8)=""AUTOZONE"",TEXT(TODAY()+1,""yyyymmdd""),TEXT(WORKDAY(TODAY(),1),""yyyymmdd""))"
    
    
'    If (Range("AA2") = "SD") Then
    
'    Range("Q2") = "=TEXT(TODAY()+1,""yyyymmdd"")"
    
'    Else

'    Range("Q2") = "=TEXT(WORKDAY(TODAY(),1),""yyyymmdd"")"
    
'    End If
    
    Range("Q2") = deliveryDate
    
    
   ' Range("Q2").value = Range("Q2").value
    
    
    Range("S2").value = "LT"
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    
    
    
    
  '  Set PoLocation = Workbooks(VdpSheetName).Worksheets("BASE BEFORE").Range("C:C").Find(PoNum, LookIn:=xlValues, searchdirection:=xlPrevious)
    Range("R2") = wsMacros.Range("V" & nextRow).value
    Range("W2").value = wsMacros.Range("S" & nextRow).value
   
'    Range("R2").value = PoLocation.Offset(0, 15)
    'Range("W2").value = PoLocation.Offset(0, 12)
    Range("Y2").Formula = Range("D4").value * Range("F4").value
    Columns("A:AC").NumberFormat = "@"
    
    Do While (Len(Range("O2")) < 5)
        Range("O2") = "0" & Range("O2").value
        
    Loop
    
        Application.Calculation = xlCalculationAutomatic
    
    
    'TESTS
    
    If "20" + Mid(Range("W2"), 2, 6) <> TodaysDate Then MsgBox "Invoice Mismatch!"
    If Range("R2") = 0 Then MsgBox "No Weight!"
  '  If Range("T2") <> PoLocation.Offset(0, 16).Value Then MsgBox "Tracking Number Mismatch!"

    
    
    
'    Workbooks(VdpSheetName).Activate
    
    
    
    
'    MsgBox Mid(PoLocation.Address, 4)
    
' 18 + 19+ 20
    If wsMacros.Range("AA" & nextRow) And wsMacros.Range("AB" & nextRow) And wsMacros.Range("AC" & nextRow) Then
'    If PoLocation.Offset(0, 18) And PoLocation.Offset(0, 19) And PoLocation.Offset(0, 20) Then
'   frank added a column for saturday delivery that messed this up (221027
    wsMacros.Range("B" & nextRow).EntireRow.Select
    wsMacros.Range("B" & nextRow).EntireRow.value = wsMacros.Range("B" & nextRow).EntireRow.value
    Else
    MsgBox "Test Failed!"

    End If
End Sub

Sub test()
    Set wsMacros = ThisWorkbook.Sheets("BASE BEFORE")
    nextRow = 4709
    wsMacros.Range("B" & nextRow).EntireRow.Select
    


End Sub
