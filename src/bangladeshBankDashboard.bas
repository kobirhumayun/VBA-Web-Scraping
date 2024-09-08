Attribute VB_Name = "bangladeshBankDashboard"
Option Explicit

Private chromeBrowser As Selenium.ChromeDriver


Sub lcDashboard()
        
    Dim allLcInfo As Variant
    allLcInfo = ActiveCell.CurrentRegion.Value
    
    Dim allLcSheet As Worksheet
    Dim allLcWorkingRange As Range
    
    Set allLcSheet = Worksheets(1)
    Set allLcWorkingRange = allLcSheet.Range("A3:" & "AE" & allLcSheet.Range("B2").End(xlDown).Row)
    
    Dim upIssuingStatus As Variant
    upIssuingStatus = allLcWorkingRange.Value
    
    Dim lcCount As Integer
    lcCount = UBound(allLcInfo) - 2
    
    
    Set chromeBrowser = New Selenium.ChromeDriver
    
    
    'pdf printer settings
    Dim settings As String
    chromeBrowser.AddArgument "--kiosk-printing"
    
    settings = "{'appState': '{""recentDestinations"": [{""id"": ""Save as PDF"", ""origin"": ""local"", ""account"": """"}], ""selectedDestination"": ""Save as PDF"", ""version"": 2}'}"
    chromeBrowser.SetPreference "printing.print_preview_sticky_settings", settings
        
    
    chromeBrowser.Start baseUrl:="https://exp.bb.org.bd"
    
    chromeBrowser.Get "/ords/f?p=116:75:9445003520227:::::"
    chromeBrowser.Wait 1000
    
'    login start
    chromeBrowser.FindElementById("P101_USERNAME").SendKeys "CUSTOMSBOND-SYL" 'user name
    chromeBrowser.FindElementById("P101_PASSWORD").SendKeys "Sylhetbond#*2025" 'password
    chromeBrowser.FindElementByClass("t12Button").Click
    chromeBrowser.Wait 1000
'    login end

    
    ' new tab open as per LC count & rename tab as LC no. start
    Dim i As Long
    
    For i = 1 To lcCount
        ' Execute JavaScript to open a new tab and navigate to the same URL
        chromeBrowser.ExecuteScript "document.title = '" & allLcInfo(i + 2, 4) & "';"
        
        If i < lcCount Then
            chromeBrowser.Wait 500
            chromeBrowser.ExecuteScript "window.open(document.URL, '_blank');"
            chromeBrowser.SwitchToNextWindow
            
        End If
        
    Next i
    ' new tab open as per LC count & rename tab as LC no. end
    
        
    ' Switch to tab, search LC & save as pdf start
    For i = 1 To lcCount
    
    ' Switch to the new tab
    If lcCount > 1 Then
        chromeBrowser.SwitchToWindowByTitle allLcInfo(i + 2, 4)
    End If

    
    'LC search start
    Dim lcNo As Variant
    
    If allLcInfo(i + 2, 33) = "" Then
        lcNo = allLcInfo(i + 2, 4)
    Else
        lcNo = allLcInfo(i + 2, 33)
    End If
    
    chromeBrowser.FindElementById("P75_SEARCH_LC").SendKeys lcNo 'LC No.
    
    chromeBrowser.FindElementByLinkText("Search").Click
    chromeBrowser.Wait 1000
    'LC search end
    
    If chromeBrowser.FindElementById("P75_LC_VALUE").Value = "" Then ' add "0" if no data received first time
    
        lcNo = Application.Run("utility_function.InsertStringAtPosition", CStr(lcNo), "0", 5)
        
        chromeBrowser.FindElementById("P75_SEARCH_LC").Clear
        chromeBrowser.FindElementById("P75_SEARCH_LC").SendKeys lcNo 'LC No.
        chromeBrowser.FindElementByLinkText("Search").Click
        chromeBrowser.Wait 1000
    End If
    
    
    If chromeBrowser.FindElementById("P75_LC_VALUE").Value = "" Then ' if LC not uploaded then just save the page
        
        chromeBrowser.ExecuteScript "document.title = '" & allLcInfo(i + 2, 4) & "';"
        GoTo LcNotUploaded
        
    End If
    
    
    If allLcInfo(i + 2, 33) <> "" Then ' for LC input box put LC  & Bangladesh Bank Ref. both
      
        chromeBrowser.FindElementById("P75_SEARCH_LC").Clear
        chromeBrowser.FindElementById("P75_SEARCH_LC").SendKeys allLcInfo(i + 2, 4) & " " & lcNo
        
    End If
    
    
    lcNo = allLcInfo(i + 2, 4) ' for title change
    
    ' Execute JavaScript to change the page title
    chromeBrowser.ExecuteScript "document.title = '" & lcNo & "';" ' change title for pdf file name
    
    Dim currentLcProperties As Object
    
    Set currentLcProperties = Application.Run("dictionary_utility_function.dicVarifyDashboard", upIssuingStatus, lcNo) 'current LC's all properties pick from UP issuing status
    
    
    Dim currentLcResultProperties As Object 'result properties
    Set currentLcResultProperties = CreateObject("Scripting.Dictionary")
    
    
    currentLcResultProperties("is_all_properties_ok") = True ' initialize
    
    
    If chromeBrowser.FindElementById("P75_BENEFICIARY_NAME").Value = "PIONEER DENIM LIMITED" Then ' check beneficiary
    
        currentLcResultProperties("beneficiary_name") = "OK"
    
    Else
    
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("beneficiary_name") = "Beneficiary name mismatch"
    
    End If
    
    
    If IsDate(chromeBrowser.FindElementById("P75_LC_DATE|input").Value) Then
    
        If DateValue(chromeBrowser.FindElementById("P75_LC_DATE|input").Value) = DateValue(currentLcProperties("lcDate")) Then
    
            currentLcResultProperties("lc_date") = "OK"
    
        Else
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("lc_date") = "LC date mismatch"
    
        End If
    
    Else
    
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("lc_date") = "LC date not found"
    
    End If
    
    
    If IsDate(chromeBrowser.FindElementById("P75_LAST_SHIP_DATE|input").Value) Then
    
        If DateValue(chromeBrowser.FindElementById("P75_LAST_SHIP_DATE|input").Value) = DateValue(currentLcProperties("shipmentDate")) Then
    
            currentLcResultProperties("shipment_date") = "OK"
    
        ElseIf DateValue(chromeBrowser.FindElementById("P75_LAST_SHIP_DATE|input").Value) > DateValue(currentLcProperties("shipmentDate")) Then
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("shipment_date") = "Shipment date greater in dashboard may be have more LC amnd"
    
        Else
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("shipment_date") = "Shipment date mismatch"
    
        End If
    
    Else
    
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("shipment_date") = "Shipment date not found"
    
    End If
    
    
    If IsDate(chromeBrowser.FindElementById("P75_LC_EXPIRY_DATE|input").Value) Then
    
        If DateValue(chromeBrowser.FindElementById("P75_LC_EXPIRY_DATE|input").Value) = DateValue(currentLcProperties("expiryDate")) Then
    
            currentLcResultProperties("expiry_date") = "OK"
    
        ElseIf DateValue(chromeBrowser.FindElementById("P75_LC_EXPIRY_DATE|input").Value) > DateValue(currentLcProperties("expiryDate")) Then
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("expiry_date") = "Expiry date greater in dashboard may be have more LC amnd"
    
        Else
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("expiry_date") = "Expiry date mismatch"
    
        End If
    
    Else
    
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("expiry_date") = "Expiry date not found"
    
    End If
    
    
    If Application.Run("utility_function.buyerNameCompare", currentLcProperties("buyerName"), chromeBrowser.FindElementById("P75_IMPORTER").Value) Then ' check buyer name in IRC field
        currentLcResultProperties("buyerNameIrc") = "OK"
    Else
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("buyerNameIrc") = "Buyer name in IRC field mismatch"
    End If
    
    
    If Application.Run("utility_function.buyerNameCompare", currentLcProperties("buyerName"), chromeBrowser.FindElementById("P75_EXPORTER").Value) Then ' check buyer name in ERC field
        currentLcResultProperties("buyerNameErc") = "OK"
    Else
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("buyerNameErc") = "Buyer name in ERC field mismatch"
    End If
    
    
    If IsNumeric(chromeBrowser.FindElementById("P75_LC_VALUE").Value) Then
    
        If CDbl(chromeBrowser.FindElementById("P75_LC_VALUE").Value) = CDbl(currentLcProperties("value")) Then
    
            currentLcResultProperties("value") = "OK"
    
        ElseIf CDbl(chromeBrowser.FindElementById("P75_LC_VALUE").Value) > CDbl(currentLcProperties("value")) Then
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("value") = "Value greater in dashboard may be have more LC amnd"
    
        Else
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("value") = "Value mismatch = " & Round(CDbl(chromeBrowser.FindElementById("P75_LC_VALUE").Value) - CDbl(currentLcProperties("value")), 2)
    
        End If
    
    Else
    
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("value") = "Value not found"
    
    End If
    
    
    ' Find elements with the specified class name
    Dim elements As Selenium.WebElements
    Set elements = chromeBrowser.FindElementsByClass("t12data")
    
    
    Dim qtyFromDashboard As Variant
    qtyFromDashboard = elements.Item(5).Text
    
    If IsNumeric(qtyFromDashboard) Then
    
        If CDbl(qtyFromDashboard) = CDbl(currentLcProperties("qty")) Then
    
            currentLcResultProperties("qty") = "OK"
    
        ElseIf CDbl(qtyFromDashboard) > CDbl(currentLcProperties("qty")) Then
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("qty") = "Qty. greater in dashboard may be have more LC amnd"
    
        Else
    
            currentLcResultProperties("is_all_properties_ok") = False
            currentLcResultProperties("qty") = "Qty. mismatch = " & Round(CDbl(qtyFromDashboard) - CDbl(currentLcProperties("qty")), 2)
    
        End If
    
    Else
    
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("qty") = "Qty. not found"
    
    End If
    
    
    Dim mLCFromDashboard As Variant
    mLCFromDashboard = elements.Item(8).Text
    
    If Application.Run("utility_function.mLcCompare", currentLcProperties("mLC"), mLCFromDashboard) Then ' check M.LC
        currentLcResultProperties("mLc") = "OK"
    Else
        currentLcResultProperties("is_all_properties_ok") = False
        currentLcResultProperties("mLc") = "M.LC mismatch"
    End If
    
    
    Dim resultStr As Variant
    resultStr = ""
    
    If currentLcResultProperties("is_all_properties_ok") Then
        resultStr = "All Field is OK"
    Else
    
        Dim key As Variant
        For Each key In currentLcResultProperties.Keys
            ' Print key and value
            If key <> "is_all_properties_ok" Then  'exclude this properties
                If currentLcResultProperties(key) <> "OK" Then 'exclude this properties
    '                    Debug.Print key, currentLcResultProperties(key)
                    resultStr = resultStr & currentLcResultProperties(key) & ", "
                End If
            End If
        Next
        resultStr = Trim$(resultStr) 'remove leading space
        resultStr = Left$(resultStr, Len(resultStr) - 1) 'remove last comma
    
        'resize the text area
    '        chromeBrowser.ExecuteScript "document.getElementById('" & "P75_CANCEL_CAUSE" & "').style.fontSize = '12px';"
        chromeBrowser.ExecuteScript "document.getElementById('" & "P75_CANCEL_CAUSE" & "').style.width = '600px';"
        chromeBrowser.ExecuteScript "document.getElementById('" & "P75_CANCEL_CAUSE" & "').rows = '3';"
    End If
   
        
    chromeBrowser.FindElementById("P75_CANCEL_CAUSE").Clear
    chromeBrowser.FindElementById("P75_CANCEL_CAUSE").SendKeys resultStr ' send result string to remark field
   
    
LcNotUploaded: ' if LC not uploaded then direct come on this line

    chromeBrowser.Wait 1000
    
    'sends pdf file to Downloads folder
    chromeBrowser.ExecuteScript ("window.print();")
    
    chromeBrowser.Wait 1000
    
    Next i
    
    ' Switch to tab, search LC & save as pdf end
    
        
    chromeBrowser.Close
    chromeBrowser.Quit


    
End Sub


