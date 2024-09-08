Attribute VB_Name = "dictionary_utility_function"
Option Explicit

Private Function dicVarifyDashboard(upIssuingStatus As Variant, lc As Variant) As Variant

    Dim currentLcProperties As Object
    Set currentLcProperties = CreateObject("Scripting.Dictionary")
    
    Dim i As Long
    
    For i = 1 To UBound(upIssuingStatus)
    
        If upIssuingStatus(i, 4) = lc Then
            
            currentLcProperties("buyerName") = upIssuingStatus(i, 2)
            currentLcProperties("buyerBank") = upIssuingStatus(i, 3)
            currentLcProperties("lcDate") = upIssuingStatus(i, 5)
            currentLcProperties("value") = currentLcProperties("value") + upIssuingStatus(i, 6)
            currentLcProperties("shipmentDate") = upIssuingStatus(i, 7)
            currentLcProperties("expiryDate") = upIssuingStatus(i, 8)
            currentLcProperties("qty") = currentLcProperties("qty") + upIssuingStatus(i, 9)
            currentLcProperties("mLC") = upIssuingStatus(i, 14)

        End If
        
    Next i
    
    Set dicVarifyDashboard = currentLcProperties
    
End Function
