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
Dim isXL As Boolean
Dim isNoShip As Boolean

Sub Main()

    ' Set reference to the Macros sheet
    Set wsMacros = ThisWorkbook.Sheets("BASE BEFORE")
    Call initialize
    Call CheckDateAndSelectFolder
    Call MakePullSheet
    Call ProcessCSVFiles
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
        End
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
    recipientName = wsMacros.Range("N" & nextRow).Value
    
    Dim recipientPhoneNumber As String
    recipientPhoneNumber = wsMacros.Range("P" & nextRow).Value
    
    Dim invoiceNo As String
    invoiceNo = wsMacros.Range("R" & nextRow).Value
    
    poNumber = wsMacros.Range("C" & nextRow).Value
    
    fedexBox = wsMacros.Range("Z" & nextRow).Value
    
    ' Loop until the user enters a value
    Do While recipientPhoneNumber = "0"
        ' Prompt the user for a new value
        recipientPhoneNumber = InputBox("Enter phone number for " & wsMacros.Range("D" & nextRow).Value, "Enter")
        
        ' Check if user entered a value
        If recipientPhoneNumber <> "" Then
        Else
            ' Notify the user that a value is required
            MsgBox "Please enter phone number.", vbExclamation
        End If
    Loop
    
    Dim recipientStreetLine1 As String
    recipientStreetLine1 = wsMacros.Range("H" & nextRow).Value
    
    Dim recipientStreetLine2 As String
    
    
    If wsMacros.Range("I" & nextRow).Value = "" Then
        isStreetLine2 = False
    Else
        isStreetLine2 = True
        recipientStreetLine2 = wsMacros.Range("I" & nextRow).Value
    End If
    
    Dim recipientCity As String
    recipientCity = wsMacros.Range("J" & nextRow).Value
    
    Dim recipientStateCode As String
    recipientStateCode = wsMacros.Range("K" & nextRow).Value
    If recipientStateCode = "PR" Then isPR = True Else isPR = False
    
    
    Dim recipientPostalCode As String
    recipientPostalCode = CStr(wsMacros.Range("L" & nextRow).Value)
    
    Do While Len(recipientPostalCode) < 5
        recipientPostalCode = "0" & recipientPostalCode
    Loop
    
    Dim recipientCountryCode As String
    recipientCountryCode = "US"
    
    ' Package information
    weight = wsMacros.Range("W" & nextRow).Value
    
    ' Loop until the user enters a value
    Do While weight = 0
        ' Prompt the user for a new value
        newValue = InputBox("Enter weight for " & wsMacros.Range("D" & nextRow).Value, "Enter")
        
        ' Check if user entered a value
        If newValue <> "" Then
            ' Replace A1 value with the user's input
            weight = newValue * wsMacros.Range("F" & nextRow).Value
        Else
            ' Notify the user that a value is required
            MsgBox "Please enter weight.", vbExclamation
        End If
    Loop
    
    If wsMacros.Range("AA" & nextRow).Value = "SD" Then isSD = True Else isSD = False
    
    deliveryMethod = wsMacros.Range("Z" & nextRow).Value
    brakeQuantity = wsMacros.Range("F" & nextRow).Value
    brakeSize = wsMacros.Range("V" & nextRow).Value
    
    Do While brakeSize = "0"
        ' Prompt the user for a new value
        newValue = InputBox("Enter size for " & wsMacros.Range("D" & nextRow).Value, "Enter")
        
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
    
    ' Construct the JSON payload
    If weight > 149 Then
        isNoShip = True
        wsMacros.Range("A" & nextRow) = "Overweight"
    End If
    
    If Not isNoShip Then
    
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
            jsonPayload = jsonPayload & """amount"": " & wsMacros.Range("E" & nextRow).Value & ","
            jsonPayload = jsonPayload & """currency"": ""USD"""
            jsonPayload = jsonPayload & "},"
            jsonPayload = jsonPayload & """customsValue"": {"
            jsonPayload = jsonPayload & """amount"": " & wsMacros.Range("E" & nextRow).Value * brakeQuantity & ","
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
        
        If deliveryMethod = "GROUND" And isXL And boxes = 1 And Not isOddBox Then
            jsonPayload = jsonPayload & """Dimensions"": {"
            jsonPayload = jsonPayload & """length"": " & 12 & ","
            jsonPayload = jsonPayload & """width"": " & 12 & ","
            jsonPayload = jsonPayload & """height"": " & 6 & ","
            jsonPayload = jsonPayload & """units"": ""IN"""
            jsonPayload = jsonPayload & "},"
        
        ElseIf deliveryMethod = "GROUND" And boxes = 1 And Not isOddBox Then
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
                
                If deliveryMethod = "GROUND" And isXL Then
                    jsonPayload = jsonPayload & """Dimensions"": {"
                    jsonPayload = jsonPayload & """length"": " & 12 & ","
                    jsonPayload = jsonPayload & """width"": " & 12 & ","
                    jsonPayload = jsonPayload & """height"": " & 6 & ","
                    jsonPayload = jsonPayload & """units"": ""IN"""
                    jsonPayload = jsonPayload & "},"
                ElseIf deliveryMethod = "GROUND" Then
                    jsonPayload = jsonPayload & """Dimensions"": {"
                    jsonPayload = jsonPayload & """length"": " & 12 & ","
                    jsonPayload = jsonPayload & """width"": " & 9 & ","
                    jsonPayload = jsonPayload & """height"": " & 5 & ","
                    jsonPayload = jsonPayload & """units"": ""IN"""
                    jsonPayload = jsonPayload & "},"
                End If
                jsonPayload = jsonPayload & """customerReferences"": ["
                jsonPayload = jsonPayload & "{"
'                jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
                jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
                jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
                jsonPayload = jsonPayload & "},"
                jsonPayload = jsonPayload & "{"
'                jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
                jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
                jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("D" & nextRow).Value & "-" & shipQuantity & """"
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
                
                If deliveryMethod = "GROUND" And isXL Then
                    jsonPayload = jsonPayload & """Dimensions"": {"
                    jsonPayload = jsonPayload & """length"": " & 12 & ","
                    jsonPayload = jsonPayload & """width"": " & 12 & ","
                    jsonPayload = jsonPayload & """height"": " & 6 & ","
                    jsonPayload = jsonPayload & """units"": ""IN"""
                    jsonPayload = jsonPayload & "},"
                ElseIf deliveryMethod = "GROUND" Then
                    jsonPayload = jsonPayload & """Dimensions"": {"
                    jsonPayload = jsonPayload & """length"": " & 12 & ","
                    jsonPayload = jsonPayload & """width"": " & 9 & ","
                    jsonPayload = jsonPayload & """height"": " & 5 & ","
                    jsonPayload = jsonPayload & """units"": ""IN"""
                    jsonPayload = jsonPayload & "},"
                End If
    
                    jsonPayload = jsonPayload & """customerReferences"": ["
                    jsonPayload = jsonPayload & "{"
                    'jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
                    jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
                    jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
                    jsonPayload = jsonPayload & "},"
                    jsonPayload = jsonPayload & "{"
                    'jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
                    jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
                    jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("D" & nextRow).Value & "-" & oddQuantity & """"
                    jsonPayload = jsonPayload & "}"
                    jsonPayload = jsonPayload & "]"
                    jsonPayload = jsonPayload & "}"
                Else
                    jsonPayload = jsonPayload & "}"
            
                End If
        Else
            jsonPayload = jsonPayload & """customerReferences"": ["
            jsonPayload = jsonPayload & "{"
            'jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
            jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
            jsonPayload = jsonPayload & """value"": """ & invoiceNo & """"
            jsonPayload = jsonPayload & "},"
            jsonPayload = jsonPayload & "{"
            'jsonPayload = jsonPayload & """customerReferenceType"": ""CUSTOMER_REFERENCE"","
            jsonPayload = jsonPayload & """customerReferenceType"": ""P_O_NUMBER"","
            jsonPayload = jsonPayload & """value"": """ & wsMacros.Range("M" & nextRow).Value & """"
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
  '          url = "https://apis-sandbox.fedex.com/ship/v1/shipments"
            url = "https://apis.fedex.com/ship/v1/shipments"
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
                    
                    Dim acrobatPath As String
                    
                    acrobatPath = "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe"
                    
            '        Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe /p /h /s " & Chr(34) & labelFilePath & Chr(34), vbHide
     '       Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe /t " & Chr(34) & labelFilePath & Chr(34) & " " & Chr(34) & "Microsoft Print to PDF" & Chr(34), vbHide
    '                Shell "C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe /t " & Chr(34) & labelFilePath & Chr(34), vbHide
                    Shell """" & acrobatPath & """ /N /T """ & labelFilePath & """", vbHide
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
    Else
    wsMacros.Range("A" & nextRow) = "Invalid size"
    End If
End Sub

Private Function GetAccessToken(apiKey As String, apiPassword As String) As String
    Dim url As String
'    url = "https://apis-sandbox.fedex.com/oauth/token"
    url = "https://apis.fedex.com/oauth/token"
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
        If wbCSV.Sheets(1).Range("D4").Value <> "" Then
            ' Close the file without saving changes
            wbCSV.Close False
        Else
            ' Copy B2 value to next available spot in row C of Macros sheet
            wsMacros.Cells(nextRow, "C").Value = wbCSV.Sheets(1).Range("B2").Value
            
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
    If ws.Range("A2").Value = Date Then
        folderPath = ws.Range("A1").Value
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
        ws.Range("A1").Value = selectedFolder.Items.Item.Path & "\"
        ws.Range("A2").Value = Date
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
    Range("D4").Value = Range("C4").Value
    Range("P2").Value = TodaysDate
    Range("Q2") = deliveryDate
    Range("S2").Value = "LT"
    Columns("U:U").EntireColumn.AutoFit
    Columns("V:V").EntireColumn.AutoFit
    Range("R2") = wsMacros.Range("W" & nextRow).Value
    Range("W2").Value = wsMacros.Range("S" & nextRow).Value
    Range("Y2").Formula = Range("D4").Value * Range("F4").Value
    Columns("A:AC").NumberFormat = "@"
    
    Do While (Len(Range("O2")) < 5)
        Range("O2") = "0" & Range("O2").Value
    Loop

        Application.Calculation = xlCalculationAutomatic
    
    'TESTS
    
    If "20" + Mid(Range("W2"), 2, 6) <> TodaysDate Then MsgBox "Invoice Mismatch!"
    If Range("R2") = 0 Then MsgBox "No Weight!"
    
    wsMacros.Activate
    
    If wsMacros.Range("AB" & nextRow) And wsMacros.Range("AC" & nextRow) And wsMacros.Range("AD" & nextRow) Then
        wsMacros.Rows(nextRow).EntireRow.Select
        wsMacros.Range("B" & nextRow).EntireRow.Value = wsMacros.Range("B" & nextRow).EntireRow.Value
    Else
        MsgBox "Test Failed!"
    End If
End Sub

Sub assignBoxSize()
    isOddBox = False
    isXL = False
    isNoShip = False
    boxes = 0
    If fedexBox = "GROUND" Then
        If brakeSize = "XL" Then
            isXL = True
            boxLoop (2)
        ElseIf brakeSize = "L" Or brakeSize = "GL" Or brakeSize = "S1" Or brakeSize = "M2" Then
            If brakeQuantity = 1 Then
                boxes = 1
                weightPerBox = weight
            Else
                boxes = brakeQuantity
                weightPerBox = weight / brakeQuantity
                shipQuantity = 1
            End If
        ElseIf brakeSize = "M" Or brakeSize = "GS" Then
            boxLoop (3)
        ElseIf brakeSize = "S" Then
          boxLoop (5)
        ElseIf brakeSize = "GM" Then
            boxLoop (2)
        Else
            isNoShip = True
        End If

    Else
        If brakeSize = "L" Or brakeSize = "XL" Or brakeSize = "GL" Then
            boxLoop (2)
        ElseIf brakeSize = "M" Then
            boxLoop (4)
        ElseIf brakeSize = "S" Or brakeSize = "GS" Then
            boxLoop (8)
        ElseIf brakeSize = "GM" Then
            boxLoop (2)
        ElseIf brakeSize = "S1" Or brakeSize = "M2" Then
            If brakeQuantity = 1 Then
                boxes = 1
                weightPerBox = weight
                shipQuantity = 1
            Else
                boxes = brakeQuantity
                weightPerBox = weight / brakeQuantity
                shipQuantity = 1
            End If
            If boxes = 0 Then
                isNoShip = True
            End If
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

Sub MakePullSheet()

    Sheets("PULL").Activate

    Dim fso As Object
    Dim strFolder As String, strFile As String, strPath As String
    Dim objStream As Object
    Dim strData As String
    Dim arrData() As String
    Dim i As Long
    Dim a As Integer
    Dim loopNum As Integer
    
    loopNum = 1
    Sheets("PULL").Range("B2:N300").ClearContents
    a = 2
    strFolder = ThisWorkbook.Sheets("Sheet1").Range("A1").Value
    
    ' Set up the file system object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Loop through all files in the folder
    strFile = Dir(strFolder & "\*.csv")
    Do While Len(strFile) > 0
        strPath = strFolder & "\" & strFile
        
        ' Read the file data using a stream
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Type = 2 ' Text stream
        objStream.Charset = "UTF-8"
        objStream.Open
        objStream.LoadFromFile strPath
        strData = objStream.ReadText(-1)
        objStream.Close
        Application.Calculation = xlCalculationManual
        ' Split the data into rows and extract the values we want
        arrData = Split(strData, vbCrLf)
        For i = 0 To UBound(arrData)
            If InStr(1, arrData(i), "ST_NAME", vbTextCompare) > 0 Then
                If InStr(1, arrData(i), "ST_ADD2", vbTextCompare) > 0 Then
                
                    If Len(Split(arrData(3), ",")(3)) = 0 Then
                        Range("A" & a) = loopNum
                        Range("B" & a) = Split(arrData(1), ",")(1)
                        Range("C" & a) = Split(arrData(3), ",")(6)
                        Range("E" & a) = Split(arrData(3), ",")(2)
                        Range("G" & a) = Split(arrData(1), ",")(12)
                        Range("I" & a) = Split(arrData(1), ",")(8)
                        a = a + 1
                        loopNum = loopNum + 1
                    End If
  
                End If
            End If
        Next i
        
        ' Get the next file in the folder
        strFile = Dir
    Loop
    Application.Calculation = xlCalculationAutomatic
    ActiveWorkbook.Worksheets("PULL").AutoFilter.Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("PULL").AutoFilter.Sort.SortFields.Add2 Key _
        :=Range("Q1:Q200"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("PULL").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    Range("A2") = 1
    Range("A3") = 2
    Range("A2:A3").AutoFill Destination:=Range("A2:A301"), Type:=xlFillDefault
    
    If loopNum > 1 Then
        Range("A1:R" & loopNum).PrintOut
    End If
    ' Clean up the objects
    Set objStream = Nothing
    Set fso = Nothing
    
    ThisWorkbook.Sheets("BASE BEFORE").Activate
End Sub




