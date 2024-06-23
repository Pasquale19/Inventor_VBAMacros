Attribute VB_Name = "DrwBomMacro"
Public Sub StyleBalloons()
    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    
    'Set a reference to the active sheet.
    Dim oSheet As sheet
    Set oSheet = oDrawDoc.ActiveSheet
    
    'Get Layer
    
    Dim oStyleManager As StylesManager
    'Set oStyleManager = oDrawDoc.StylesManager
    Dim oLayer As Layer
    Set oLayer = oDrawDoc.StylesManager.Layers.item("BOM")
    
    Dim oBal As Balloon
    
    For Each oBal In oSheet.Balloons
        oBal.Layer = oLayer
        oBal.Leader.ArrowheadType = kFilledArrowheadType
    
    Next
    
End Sub


