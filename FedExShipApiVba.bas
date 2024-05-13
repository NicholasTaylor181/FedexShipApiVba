Dim apiKey As String
Dim apiPassword As String
Dim accountNumber As String
Dim thirdPartyAccountNumber As String
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
Dim deliveryMethod As String
Dim isOverWeight As Boolean
Dim boxes As Integer
Dim fedexBox As String
Dim brakeSize As String
Dim brakeQuantity As Integer
Dim weight As Double
Dim oddBoxWeight As Integer
Dim isOddBox As Boolean
Dim weightPerBox As Double
Dim shipQuantity As Integer
Dim oddQuantity As Integer
Dim attempts As Integer
Dim isShipped As Boolean
Dim isFriday As Boolean

Sub Main()

    ' Set reference to the Macros sheet
    Set wsMacros = ThisWorkbook.Sheets("BASE BEFORE")
    Call initialize
    Call CheckDateAndSelectFolder
    Call ProcessCSVFiles
    '''''''''''''''''''''''''add logic for sizes other than the typical L M S
    'need to add pull sheet creator, and auto pdf print
    'add address validation
End Sub


Sub initialize()
    apiKey = "your api key"
    apiPassword = "your api password"
    accountNumber = "your account number"
    thirdPartyAccountNumber = "your third party account number"
    
    ' Access token
    accessToken = GetAccessToken(apiKey, apiPassword)
    
    ' Check if access token is retrieved successfully
    If accessToken = "" Then
        MsgBox "Failed to retrieve access token."
        Exit Sub
    End If
    
    'checks if today is friday
    If Weekday(Date, vbMonday) = 5 Then
        isFriday = True
    Else
        isFriday = False
    End If
End Sub

Sub CreateShipment()
    'temp value to store user inputs.
    Dim newValue As String

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
    If recipientStateCode = "PR" Then isPR = True Else isPR = False
    
    
    Dim recipientPostalCode As String
    recipientPostalCode = CStr(wsMacros.Range("L" & nextRow).value)
    
    Do While Len(recipientPostalCode) < 5
        recipientPostalCode = "0" & recipientPostalCode
    Loop
    
    Dim recipientCountryCode As String
    recipientCountryCode = "US"
    
    ' Package information
    weight = wsMacros.Range("W" & nextRow).value
    
    ' Loop until the user enters a value
    Do While weight = 0
        ' Prompt the user for a new value
        newValue = InputBox("Enter weight for " & wsMacros.Range("D" & nextRow).value, "Enter")
        
        ' Check if user entered a value
        If newValue <> "" Then
            ' Replace A1 value with the user's input
            weight = newValue * wsMacros.Range("F" & nextRow).value
        Else
            ' Notify the user that a value is required
            MsgBox "Please enter weight.", vbExclamation
        End If
    Loop
    
    If wsMacros.Range("AA" & nextRow).value = "SD" Then isSD = True Else isSD = False
    
    deliveryMethod = wsMacros.Range("Z" & nextRow).value
    brakeQuantity = wsMacros.Range("F" & nextRow).value
    brakeSize = wsMacros.Range("V" & nextRow).value
    
    Do While brakeSize = "0"
        ' Prompt the user for a new value
        newValue = InputBox("Enter size for " & wsMacros.Range("D" & nextRow).value, "Enter")
        
        ' Check if user entered a value
        If newValue <> "" Then
            ' Replace A1 value with the user's input
            brakeSize = newValue
        Else
            ' Notify the user that a value is required
            MsgBox "Please enter size.", vbExclamation
        End If
    Loop
    
    Call assignBoxSize
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''insert check for overweight, if so skip and put error in column a
    
    ' Construct the JSON payload
    Dim jsonPayload As String

   
    jsonPayload = "{"
    jsonPayload = jsonPayload & """labelResponseOptions"": ""URL_ONLY"","
    jsonPayload = jsonPayload & """requestedShipment"": {"
    jsonPayload = jsonPayload & """shipper"": {"
    jsonPayload = jsonPayload & """contact"": {"
    jsonPayload = jsonPayload & """personName"": """ & senderName & ""","
    jsonPayload = jsonPayload & """phoneNumber"": " & senderPhoneNumber & ","
    jsonPayload = jsonPayload & """companyName"": """ & senderCompanyName & """"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """address"": {"
    jsonPayload = jsonPayload & """streetLines"": [""" & senderStreetLine1 & """],"
    jsonPayload = jsonPayload & """city"": """ & senderCity & ""","
    jsonPayload = jsonPayload & """stateOrProvinceCode"": """ & senderStateCode & ""","
    jsonPayload = jsonPayload & """postalCode"": " & senderPostalCode & ","
    jsonPayload = jsonPayload & """countryCode"": """ & senderCountryCode & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """recipients"": ["
    jsonPayload = jsonPayload & "{"
    jsonPayload = jsonPayload & """contact"": {"
    jsonPayload = jsonPayload & """personName"": """ & recipientName & ""","
    jsonPayload = jsonPayload & """phoneNumber"": " & recipientPhoneNumber & ","
    jsonPayload = jsonPayload & """companyName"": """ & recipientCompanyName & """"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """address"": {"
    
    If isStreetLine2 Then
        jsonPayload = jsonPayload & """streetLines"": [""" & recipientStreetLine1 & """,""" & recipientStreetLine2 & """],"
    Else
        jsonPayload = jsonPayload & """streetLines"": [""" & recipientStreetLine1 & """],"
    End If
    
    jsonPayload = jsonPayload & """city"": """ & recipientCity & ""","
    jsonPayload = jsonPayload & """stateOrProvinceCode"": """ & recipientStateCode & ""","
    jsonPayload = jsonPayload & """postalCode"": """ & recipientPostalCode & ""","
    jsonPayload = jsonPayload & """countryCode"": """ & recipientCountryCode & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "],"
    jsonPayload = jsonPayload & """shipDatestamp"": """ & Format(Date, "yyyy-mm-dd") & ""","
    If isPR Then
        jsonPayload = jsonPayload & """serviceType"": ""INTERNATIONAL_ECONOMY"","
        jsonPayload = jsonPayload & """packagingType"": ""FEDEX_BOX"","
    Else
        If deliveryMethod = "GROUND" Then
            jsonPayload = jsonPayload & """serviceType"": ""FEDEX_GROUND"","
            jsonPayload = jsonPayload & """packagingType"": ""YOUR_PACKAGING"","
        
        ElseIf deliveryMethod = "STANDARD" Then
            jsonPayload = jsonPayload & """serviceType"": ""STANDARD_OVERNIGHT"","
            jsonPayload = jsonPayload & """packagingType"": ""FEDEX_BOX"","
        
        ElseIf deliveryMethod = "PRIORITY" Then
            jsonPayload = jsonPayload & """serviceType"": ""PRIORITY_OVERNIGHT"","
            jsonPayload = jsonPayload & """packagingType"": ""FEDEX_BOX"","
            If isSD And isFriday Then
                jsonPayload = jsonPayload & """shipmentSpecialServices"": {"
                jsonPayload = jsonPayload & """specialServiceTypes"": ["
                jsonPayload = jsonPayload & """SATURDAY_DELIVERY"""
                jsonPayload = jsonPayload & "]"
                jsonPayload = jsonPayload & "},"
            End If
        End If
    End If
    jsonPayload = jsonPayload & """pickupType"": ""USE_SCHEDULED_PICKUP"","
    jsonPayload = jsonPayload & """blockInsightVisibility"": false,"
    jsonPayload = jsonPayload & """shippingChargesPayment"": {"
    
    jsonPayload = jsonPayload & """paymentType"": ""THIRD_PARTY"","
    jsonPayload = jsonPayload & """payor"": {"
    jsonPayload = jsonPayload & """responsibleParty"": {"
    jsonPayload = jsonPayload & """accountNumber"": {"
    jsonPayload = jsonPayload & """value"": """ & thirdPartyAccountNumber & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "},"
    
    jsonPayload = jsonPayload & """labelSpecification"": {"
    jsonPayload = jsonPayload & """imageType"": ""PDF"","
    jsonPayload = jsonPayload & """labelStockType"": ""PAPER_85X11_TOP_HALF_LABEL"""
    jsonPayload = jsonPayload & "},"
    
    
    If isPR Then
    
        jsonPayload = jsonPayload & """customsClearanceDetail"": {"
        jsonPayload = jsonPayload & """dutiesPayment"": {"
        jsonPayload = jsonPayload & """paymentType"": ""THIRD_PARTY"","
        jsonPayload = jsonPayload & """payor"": {"
        jsonPayload = jsonPayload & """responsibleParty"": {"
        jsonPayload = jsonPayload & """accountNumber"": {"
        jsonPayload = jsonPayload & """value"": """ & thirdPartyAccountNumber & """"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "},"
        jsonPayload = jsonPayload & """isDocumentOnly"": false,"
        jsonPayload = jsonPayload & """commodities"": ["
        jsonPayload = jsonPayload & "{"
        jsonPayload = jsonPayload & """description"": ""Commercial -  - - BRAKE LININGS AND PADS FRICTION MATERIAL AND ARTICLES THEREOF (FOR EXAMPLE, SHEETS, ROLLS, STRIPS, SEGMENTS, DISCS, WASHERS, PADS), NOT MOUNTED,FOR BRAKES, FOR CLUTCHES OR THE LIKE, WITH A BASIS OF ASBESTOS, OF OTHER MINERAL SUBSTANCES OR OF CELLULOSE, WHETHER OR NOT COMBINED WITH TEXTILE OR OTHER MATERIALS: - NOT CONTAINING ASBESTOS:- - BRAKE LININGS AND PADS"","
        jsonPayload = jsonPayload & """countryOfManufacture"": ""US"","
        jsonPayload = jsonPayload & """quantity"": " & brakeQuantity & ","
        jsonPayload = jsonPayload & """quantityUnits"": ""PCS"","
        jsonPayload = jsonPayload & """unitPrice"": {"
        jsonPayload = jsonPayload & """amount"": " & wsMacros.Range("E" & nextRow).value & ","
        jsonPayload = jsonPayload & """currency"": ""USD"""
        jsonPayload = jsonPayload & "},"
        jsonPayload = jsonPayload & """customsValue"": {"
        jsonPayload = jsonPayload & """amount"": " & wsMacros.Range("E" & nextRow).value * brakeQuantity & ","
        jsonPayload = jsonPayload & """currency"": ""USD"""
        jsonPayload = jsonPayload & "},"
        jsonPayload = jsonPayload & """weight"": {"
        jsonPayload = jsonPayload & """units"": ""LB"","
        jsonPayload = jsonPayload & """value"": """ & weight / brakeQuantity & """"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "]"
            
        jsonPayload = jsonPayload & "},"
        jsonPayload = jsonPayload & """shippingDocumentSpecification"": {"
        jsonPayload = jsonPayload & """shippingDocumentTypes"": ["
        jsonPayload = jsonPayload & """COMMERCIAL_INVOICE"""
        jsonPayload = jsonPayload & "],"
        jsonPayload = jsonPayload & """commercialInvoiceDetail"": {"
        jsonPayload = jsonPayload & """documentFormat"": {"
        jsonPayload = jsonPayload & """stockType"": ""PAPER_LETTER"","
        jsonPayload = jsonPayload & """docType"": ""PDF"""
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "},"
    
    
    
    End If
    
    
    
    
    
    
    
    
    
    jsonPayload = jsonPayload & """requestedPackageLineItems"": ["
    jsonPayload = jsonPayload & "{"
    
    If deliveryMethod = "GROUND" And boxes = 1 And Not isOddBox Then
        jsonPayload = jsonPayload & """Dimensions"": {"
        jsonPayload = jsonPayload & """length"": " & 12 & ","
        jsonPayload = jsonPayload & """width"": " & 9 & ","
        jsonPayload = jsonPayload & """height"": " & 5 & ","
        jsonPayload = jsonPayload & """units"": ""IN"""
        jsonPayload = jsonPayload & "},"
    End If

    If boxes > 1 Or isOddBox Then
            jsonPayload = jsonPayload & """groupPackageCount"": " & boxes & ","
            jsonPayload = jsonPayload & """weight"": {"
            jsonPayload = jsonPayload & """value"": " & weightPerBox & ","
            jsonPayload = jsonPayload & """units"": ""LB"""
            jsonPayload = jsonPayload & "},"
            
            If deliveryMethod = "GROUND" Then
                jsonPayload = jsonPayload & """Dimensions"": {"
                jsonPayload = jsonPayload & """length"": " & 12 & ","
                jsonPayload = jsonPayload & """width"": " & 9 & ","
                jsonPayload = jsonPayload & """height"": " & 5 & ","
                jsonPayload = jsonPayload & """units"": ""IN"""
                jsonPayload = jsonPayload & "},"
            End If
            jsonPayload = jsonPayload & """customerReferences"": ["
            jsonPayload = jsonPayload & "{"
            jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
            jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
            jsonPayload = jsonPayload & "},"
            jsonPayload = jsonPayload & "{"
            jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
            jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("D" & nextRow).value & "-" & shipQuantity & """"
            jsonPayload = jsonPayload & "}"
            jsonPayload = jsonPayload & "]"
            If isOddBox Then
                jsonPayload = jsonPayload & "},"
                jsonPayload = jsonPayload & "{"
                jsonPayload = jsonPayload & """groupPackageCount"": " & 1 & ","
                jsonPayload = jsonPayload & """weight"": {"
                jsonPayload = jsonPayload & """value"": " & oddBoxWeight & ","
                jsonPayload = jsonPayload & """units"": ""LB"""
                jsonPayload = jsonPayload & "},"
            
                If deliveryMethod = "GROUND" Then
                    jsonPayload = jsonPayload & """Dimensions"": {"
                    jsonPayload = jsonPayload & """length"": " & 12 & ","
                    jsonPayload = jsonPayload & """width"": " & 9 & ","
                    jsonPayload = jsonPayload & """height"": " & 5 & ","
                    jsonPayload = jsonPayload & """units"": ""IN"""
                    jsonPayload = jsonPayload & "},"
                End If

                jsonPayload = jsonPayload & """customerReferences"": ["
                jsonPayload = jsonPayload & "{"
                jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
                jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
                jsonPayload = jsonPayload & "},"
                jsonPayload = jsonPayload & "{"
                jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
                jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("D" & nextRow).value & "-" & oddQuantity & """"
                jsonPayload = jsonPayload & "}"
                jsonPayload = jsonPayload & "]"
                jsonPayload = jsonPayload & "}"
            
            Else
                jsonPayload = jsonPayload & "}"
        
            End If
    
    
    
    
    

    Else
        jsonPayload = jsonPayload & """customerReferences"": ["
        jsonPayload = jsonPayload & "{"
        jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
        jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
        jsonPayload = jsonPayload & "},"
        jsonPayload = jsonPayload & "{"
        jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
        jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("M" & nextRow).value & """"
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "],"
        jsonPayload = jsonPayload & """weight"": {"
        jsonPayload = jsonPayload & """value"": " & weight & ","
        jsonPayload = jsonPayload & """units"": ""LB"""
        jsonPayload = jsonPayload & "}"
        jsonPayload = jsonPayload & "}"
    End If
    jsonPayload = jsonPayload & "]"
    jsonPayload = jsonPayload & "},"
    jsonPayload = jsonPayload & """accountNumber"": {"
    jsonPayload = jsonPayload & """value"": """ & accountNumber & """"
    jsonPayload = jsonPayload & "}"
    jsonPayload = jsonPayload & "}"
   
   wsMacros.Range("AP" & nextRow) = jsonPayload
   isShipped = False
   attempts = 0
   
   Do While attempts < 3
        wsMacros.Range("AQ" & nextRow) = attempts
       
        ' Make the API request
        Dim url As String
        url = "https://apis-sandbox.fedex.com/ship/v1/shipments"
    
        Dim http As Object
        Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
        http.Open "POST", url, False
        http.SetRequestHeader "Authorization", "Bearer " & accessToken
        http.SetRequestHeader "Content-Type", "application/json"
        http.Send jsonPayload
    
        ' Check if the request was successful
            wsMacros.Range("A" & nextRow) = http.Status
        If http.Status = 200 Then
            Dim responseJson As Object
            Set responseJson = JsonConverter.ParseJson(http.responseText)
            trackingNumber = responseJson("output")("transactionShipments")(1)("masterTrackingNumber")
            deliveryDate = responseJson("output")("transactionShipments")(1)("pieceResponses")(1)("deliveryDatestamp")
            
            ' Retrieve the label URL from the response
            Dim labelUrl As String
            If isPR Then
                labelUrl = responseJson("output")("transactionShipments")(1)("shipmentDocuments")(2)("url")
            Else
                If isOddBox Or boxes > 1 Then
                    labelUrl = responseJson("output")("transactionShipments")(1)("shipmentDocuments")(1)("url")
                Else
                    labelUrl = responseJson("output")("transactionShipments")(1)("pieceResponses")(1)("packageDocuments")(1)("url")
                End If
            End If
            
            ThisWorkbook.Sheets("Sheet1").Range("M2") = labelUrl
            deliveryDate = Left(deliveryDate, 4) & Mid(deliveryDate, 6, 2) & Right(deliveryDate, 2)
            
            ' Download and save the label as PDF
            Dim labelFilePath As String
            labelFilePath = folderPath & poNumber & ".pdf"
    
            Dim labelHttp As Object
            Dim labelFile As Object
            Dim xmlHTTP As Object
    
            ' Create a new XMLHTTP object
            Set xmlHTTP = CreateObject("MSXML2.XMLHTTP")
        
            ' Open the URL
            xmlHTTP.Open "GET", labelUrl, False
        
            ' Send the request
            xmlHTTP.Send
            Dim stream As Object
        
            ' Check if the request was successful
    
            If xmlHTTP.Status = 200 Then
                ' Create a new FileStream object to write the PDF content
            
            
                ' Create a new Stream object to write the PDF content
                Set stream = CreateObject("ADODB.Stream")
        
                ' Set stream properties
                stream.Type = 1 ' adTypeBinary
                stream.Open
        
                ' Write the response content (PDF) to the Stream
                stream.Write xmlHTTP.ResponseBody
        
                ' Save the Stream to a file
                stream.SaveToFile labelFilePath, 2 ' adSaveCreateOverWrite
        
                ' Close the Stream
                stream.Close
                
                Call VDP_FORMAT
                attempts = 20
                     
            Else
                attempts = attempts + 1
            End If
        Else
            attempts = attempts + 1
        End If

    Loop
    
        Set xmlHTTP = Nothing
End Sub

Private Function GetAccessToken(apiKey As String, apiPassword As String) As String
    Dim url As String
    url = "https://apis-sandbox.fedex.com/oauth/token"

    Dim http As Object
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
    Else
        GetAccessToken = ""
    End If
End Function
Sub AddressValidation()






End Sub
Sub ProcessCSVFiles()
    Dim csvFile As String

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
            
            Call CreateShipment
            
            ' Close the CSV file
            wbCSV.Close True
            
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
        ws.Range("A2").value = Date
        folderPath = selectedFolder.Items.Item.Path & "\"
    End If
End Sub

Sub VDP_FORMAT()
    Dim SelectedRow As Range
    Dim OrderSheetName As String
    Dim VdpSheetName As String
    Dim TodaysDate As String
    Dim PoNum As String
    
    wbCSV.Sheets(1).Activate
    TodaysDate = Format(Date, "yyyymmdd")
    
    Application.Calculation = xlCalculationManual
    Range("U2:V2").NumberFormat = "@"
    Range("U2:V2") = trackingNumber
    Range("D4").value = Range("C4").value
    Range("P2").value = TodaysDate
    Range("Q2") = deliveryDate
    Range("S2").value = "LT"
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    Range("R2") = wsMacros.Range("W" & nextRow).value
    Range("W2").value = wsMacros.Range("S" & nextRow).value
    Range("Y2").Formula = Range("D4").value * Range("F4").value
    Columns("A:AC").NumberFormat = "@"
    
    Do While (Len(Range("O2")) < 5)
        Range("O2") = "0" & Range("O2").value
    Loop

        Application.Calculation = xlCalculationAutomatic
    
    'TESTS
    
    If "20" + Mid(Range("W2"), 2, 6) <> TodaysDate Then MsgBox "Invoice Mismatch!"
    If Range("R2") = 0 Then MsgBox "No Weight!"
    
    wsMacros.Activate
    
    If wsMacros.Range("AB" & nextRow) And wsMacros.Range("AC" & nextRow) And wsMacros.Range("AD" & nextRow) Then
        wsMacros.Rows(nextRow).EntireRow.Select
        wsMacros.Range("B" & nextRow).EntireRow.value = wsMacros.Range("B" & nextRow).EntireRow.value
    Else
        MsgBox "Test Failed!"
    End If
End Sub

Sub assignBoxSize()
    isOddBox = False
        
    If fedexBox = "GROUND" Then
        If brakeSize = "L" Then
            If brakeQuantity = 1 Then
                boxes = 1
                weightPerBox = weight
            Else
                boxes = brakeQuantity
                weightPerBox = weight / brakeQuantity
                shipQuantity = 1
            End If
        ElseIf brakeSize = "M" Then
            boxLoop (3)
        ElseIf brakeSize = "S" Then
          boxLoop (5)
        End If
    Else
        If brakeSize = "L" Then
            boxLoop (2)
        ElseIf brakeSize = "M" Then
            boxLoop (4)
        ElseIf brakeSize = "S" Then
            boxLoop (8)
        End If
    End If
End Sub


Sub modWeight(modQuantity As Integer)
    boxes = (brakeQuantity - brakeQuantity Mod modQuantity) / modQuantity
    weightPerBox = (weight / brakeQuantity) * modQuantity
    isOddBox = False
    If brakeQuantity Mod modQuantity <> 0 Then
        isOddBox = True
        oddBoxWeight = (weight / brakeQuantity) * (brakeQuantity Mod modQuantity)
    End If
End Sub
Sub boxLoop(startMod As Integer)
    isOddBox = False
    weightPerBox = (weight / brakeQuantity) * startMod
    If startMod + 1 > brakeQuantity Then
        If weight < 20 Then
            boxes = 1
            weightPerBox = weight
        Else
            modWeight (startMod)
            Do While weightPerBox > 20
                startMod = startMod - 1
                modWeight (startMod)
            Loop
        End If
    Else
        modWeight (startMod)
        Do While weightPerBox > 20
            startMod = startMod - 1
            modWeight (startMod)
        Loop
    End If
    
    If boxes > 1 Or isOddBox Then
        shipQuantity = startMod
        If isOddBox Then
            oddQuantity = brakeQuantity Mod startMod
        End If
    End If
End Sub
