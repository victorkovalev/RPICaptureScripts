''' <summary>
'''     Function will take in the a string and do a fuzzy match to a list in the field. 
'''         > fnGetValueInFieldList("Class1", "Country", "Peru1", 0.2) -> "Peru"
'''     Spaces are removed for comparison.
'''         > fnGetValueInFieldList("Class1", "Country", "P e r u", 0.2) -> "Peru"
'''     Spaces are removed for comparison.
'''         > fnGetValueInFieldList("Class1", "Country", "Peruza", 0.2) -> ""
''' </summary>
''' <remarks>Requires Logging Functions, Cedar String Compare library</remarks>
''' <param name="strClassName">Class name where the validation list is defined</param>
''' <param name="strFieldName">Field name where the validation list is present</param>
''' <param name="strValue">String value to match</param>
''' <param name="sglDistance">Distance allowed between the text that we enter and an item found in the list</param>
''' <returns>"" if no match is found with tolerance provided, Otherwise item that matches the list</returns>

Function fnGetValueInFieldList(strClassName As String, strFieldName As String, strValue As String, sglDistance As Single) As String
    On Error GoTo lbl_error

    fnLog(CDRTypeInfo, "fnGetValueInFieldList Start strClassName [" & strClassName & "] strFieldName [" & strFieldName & "] strValue [" & strValue & "]")
    fnLogIndentIncrease()

    fnGetValueInFieldList = ""

    ' Get list settings
    Dim oIVRSettings As SCBCdrListSettings
    Dim oObject As Object
    Dim oIVRFieldDef As SCBCdrFieldDef

    Set oIVRFieldDef = Project.AllClasses(strClassName).Fields.ItemByName(strFieldName)
    Set oObject = oIVRFieldDef.ValidationSettings("German")
    Set oIVRSettings = oObject

    ' Initialize compare settings
    Dim sglDist As Single
    Dim oStrComp As SCBCdrStringComp
    Set oStrComp = New SCBCdrStringComp

    oStrComp.CompType = CdrTypeLevenShtein
    oStrComp.CaseSensitive = False
    oStrComp.SearchExpression = Replace(strValue, " ", "") 'Compare without spaces

    ' Variables to store best candidates
    Dim lngItem As Long
    Dim lngBestItem As Long
    Dim sglBestItemDistance As Single

    lngBestItem = 0
    sglBestItemDistance = 1

    ' Iterate through the list, and compare each item
    For lngItem = 1 To oIVRFieldDef.ListItemCount
        'Compare without spaces
        oStrComp.Distance(Replace(oIVRSettings.ListValues.Item(lngItem), " ", ""), sglDist)
        fnLog(CDRTypeInfo, "fnGetValueInFieldList lngItem [" & CStr(lngItem) & "] itemText [" & oIVRSettings.ListValues.Item(lngItem) & "] sglDist [" & CStr(sglDist) & "]")

        'If match is better, store it
        If (sglDist < sglBestItemDistance) Then
            lngBestItem = lngItem
            sglBestItemDistance = sglDist
        End If

    Next lngItem

    ' Log the best candidate
    fnLog(CDRTypeInfo, "fnGetValueInFieldList lngBestItem [" & CStr(lngBestItem) & "] strBestItemText [" & oIVRSettings.ListValues.Item(lngBestItem) & "] sglBestItemDistance [" & CStr(sglBestItemDistance) & "]")

    ' Return text of best candidate if item in the list is within tolerance
    If (sglBestItemDistance <= sglDistance) Then fnGetValueInFieldList = oIVRSettings.ListValues.Item(lngBestItem)

lbl_end:
    fnLogIndentDecrease()
    fnLog(CDRTypeInfo, "fnGetValueInFieldList End Returning [" & CStr(fnGetValueInFieldList) & "]")
    Exit Function
lbl_error:
    fnLogIndentDecrease()
    fnLog(CDRTypeError, "fnGetValueInFieldList Error [" & Err.Description & "]")
End Function
