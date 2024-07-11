Attribute VB_Name = "SketchElementsToLayer"
Public Sub setSketchToLayer(LayerName As String)
'set each element of the sketch to the specified Layer

Dim oSketch As Sketch
Dim txt As String
Dim t As String
t = TypeName(ThisApplication.ActiveEditObject)
If TypeOf ThisApplication.ActiveEditObject Is Sketch Then
    MsgBox ThisApplication.ActiveEditObject.name
    Set oSketch = ThisApplication.ActiveEditObject
    txt = "oSketch.Name" & " is the active Sketch"
    'MsgBox txt, vbOKOnly, , "Hier steht der Titel"
Else
    txt = "Active Edit Object is not a Sketch"
    MsgBox txt, vbOKOnly, , "Hier steht der Titel"
    
End If

    Dim oDrawDoc As DrawingDocument
    Set oDrawDoc = ThisApplication.ActiveDocument
    

    'Get Layer
    Dim oLayer As Layer
    Set oLayer = oDrawDoc.StylesManager.Layers.Item(LayerName)
    
    Dim oEnt As SketchEntity
    For Each oEnt In oSketch.SketchEntities
        oEnt.Layer = oLayer
    Next
    
    Dim oTxt As TextBox
    
    For Each oTxt In oSketch.TextBoxes
        oTxt.Layer = oLayer
    Next

End Sub

Public Sub setLayer()

    Dim name As String
    name = InputBox("enter Layername", "Layername", "Schaltung1")
    Call setSketchToLayer(name)
    
End Sub
