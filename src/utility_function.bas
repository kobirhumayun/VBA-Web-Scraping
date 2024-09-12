Attribute VB_Name = "utility_function"
Option Explicit

Private Function InsertStringAtPosition(originalString As String, insertString As String, position As Integer) As Variant
    Dim length As Integer
    length = Len(originalString)
    
    If length >= position Then
        InsertStringAtPosition = Left(originalString, length - (position - 1)) & insertString & Right(originalString, (position - 1))
    Else
        ' Handle the case where the original string is shorter than 5 characters
        InsertStringAtPosition = Null
    End If
End Function


Private Function RemoveInvalidChars(ByVal inputString As String) As String
    Dim invalidChars As String
    invalidChars = " ~`!@#$%^&*()-+=[]\{}|;':"",./<>?"
    
    Dim resultString As String
    Dim i As Long
    
    For i = 1 To Len(inputString)
        Dim currentChar As String
        currentChar = Mid(inputString, i, 1)
        
        If InStr(invalidChars, currentChar) = 0 Then
            resultString = resultString & currentChar
        End If
    Next i
    
    RemoveInvalidChars = resultString
End Function


Private Function mLcCompare(mLcFromUpIssuingStatus As Variant, mLCFromDashboard As Variant) As Boolean

    On Error GoTo ErrorHandler

    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.MultiLine = True

    Dim regExReturnedMLcObject As Variant

    regEx.Pattern = ".+"
    Set regExReturnedMLcObject = regEx.Execute(mLcFromUpIssuingStatus)

    mLCFromDashboard = Replace(mLCFromDashboard, "&", "AND") ' replace & with "AND" for matching
    mLCFromDashboard = Application.Run("utility_function.RemoveInvalidChars", mLCFromDashboard) 'remove all invalid characters for matching

    Dim tempMLc As Variant
    Dim bool As Boolean

    Dim iterator As Integer

    For iterator = 0 To regExReturnedMLcObject.Count - 1

        tempMLc = Replace(regExReturnedMLcObject.Item(iterator), "&", "AND") ' replace & with "AND" for matching
        tempMLc = Application.Run("utility_function.RemoveInvalidChars", tempMLc) 'remove all invalid characters for matching

        regEx.Pattern = tempMLc
        bool = regEx.test(mLCFromDashboard)

        If bool Then
            mLcCompare = bool
            Exit Function
        End If

    Next iterator

    mLcCompare = False


    Exit Function

ErrorHandler:
    mLcCompare = False 'when an error occurs return false

End Function


Private Function buyerNameCompare(buyerFromUpIssuingStatus As Variant, buyerFromDashboard As Variant) As Boolean
    
    On Error GoTo ErrorHandler
    
    Dim regEx As Object
    Set regEx = CreateObject("VBScript.RegExp")
    regEx.Global = True
    regEx.MultiLine = True
    regEx.IgnoreCase = True
    
    buyerFromUpIssuingStatus = LCase$(buyerFromUpIssuingStatus) ' convert lower case
    buyerFromDashboard = LCase$(buyerFromDashboard) ' convert lower case
    
    buyerFromUpIssuingStatus = Replace(buyerFromUpIssuingStatus, "limited", "ltd") ' replace limited to ltd for matching
    buyerFromDashboard = Replace(buyerFromDashboard, "limited", "ltd") ' replace limited to ltd for matching

    Dim regExReturnedBuyerFromUpIssuingStatusObject, regExReturnedBuyerFromDashboardObject As Variant

    regEx.Pattern = ".+((ltd)|(limited))"

    Set regExReturnedBuyerFromUpIssuingStatusObject = regEx.Execute(buyerFromUpIssuingStatus)
    Set regExReturnedBuyerFromDashboardObject = regEx.Execute(buyerFromDashboard)

    buyerFromUpIssuingStatus = regExReturnedBuyerFromUpIssuingStatusObject.Item(0) 'take buyer name
    buyerFromDashboard = regExReturnedBuyerFromDashboardObject.Item(0) 'take buyer name

    buyerFromUpIssuingStatus = Application.Run("utility_function.RemoveInvalidChars", buyerFromUpIssuingStatus) 'remove all invalid characters for matching
    buyerFromDashboard = Application.Run("utility_function.RemoveInvalidChars", buyerFromDashboard) 'remove all invalid characters for matching

    Dim bool As Boolean
    bool = LCase$(buyerFromUpIssuingStatus) = LCase$(buyerFromDashboard)


    buyerNameCompare = bool
    
    Exit Function
    
ErrorHandler:
    buyerNameCompare = False 'when an error occurs return false

End Function

Private Function lcValueQtyAsDicFromErpReport(piReportFilePath As String) As Object

    Application.ScreenUpdating = False
        
    Dim piReportWb As Workbook
    Dim piReportWs As Worksheet
    Set piReportWb = Workbooks.Open(piReportFilePath)
    Set piReportWs = piReportWb.Worksheets(1)

    piReportWs.AutoFilterMode = False
        
    Dim temp As Variant
    temp = piReportWs.Range("A4").CurrentRegion.value

    piReportWb.Close SaveChanges:=False
        
    Dim i As Long

    Dim lcValueQtyAsDic As Object
    Set lcValueQtyAsDic = CreateObject("Scripting.Dictionary")

    For i = 2 To UBound(temp)

        If Not lcValueQtyAsDic.Exists(temp(i, 4)) Then

                'take unique LC as dictionary
            lcValueQtyAsDic.Add temp(i, 4), CreateObject("Scripting.Dictionary")

            lcValueQtyAsDic(temp(i, 4)).Add "piList", CreateObject("Scripting.Dictionary")

        End If

        If Not lcValueQtyAsDic(temp(i, 4))("piList").Exists(temp(i, 6)) Then

                'take unique PI
            lcValueQtyAsDic(temp(i, 4))("piList").Add temp(i, 6), temp(i, 6)

                'take only one time, cause total Qty duplicate
            lcValueQtyAsDic(temp(i, 4))("NetWeight") = lcValueQtyAsDic(temp(i, 4))("NetWeight") + CDec(temp(i, 23))
            lcValueQtyAsDic(temp(i, 4))("GrossWeight") = lcValueQtyAsDic(temp(i, 4))("GrossWeight") + CDec(temp(i, 25))

        End If

        lcValueQtyAsDic(temp(i, 4))("PIQty") = lcValueQtyAsDic(temp(i, 4))("PIQty") + CDec(temp(i, 18))
        lcValueQtyAsDic(temp(i, 4))("Unit") = temp(i, 19)
        lcValueQtyAsDic(temp(i, 4))("PIAmount") = lcValueQtyAsDic(temp(i, 4))("PIAmount") + CDec(temp(i, 20))
        lcValueQtyAsDic(temp(i, 4))("Currency") = temp(i, 21)
        lcValueQtyAsDic(temp(i, 4))("Buyer") = temp(i, 26)
        lcValueQtyAsDic(temp(i, 4))("BuyerAddress") = temp(i, 27)

    Next i

    Set lcValueQtyAsDicFromErpReport = lcValueQtyAsDic

End Function


