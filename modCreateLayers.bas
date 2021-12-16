Attribute VB_Name = "modCreateLayers"
Option Explicit

Public Sub ColorToLayer()
     
     ActiveDocument.BeginCommandGroup "ColorToLayer"
     
     Dim i As Integer
     Dim s As Shape
     Dim sr As ShapeRange
     Dim p As page
     Dim l As Layer
     Dim aux As String
     
     Dim cfg As New clsSettings
     
     Dim exists As Boolean
     
     cfg.TurnOff
     
     Set p = ActiveDocument.ActivePage
     
     i = 0
     
     p.Layers(2).Activate
     
     For Each s In p.Shapes
          
          i = i + 1
          
          s.CreateSelection
          
          Set sr = ActiveSelectionRange
          
          On Error Resume Next

          If s.Fill.Type <> cdrNoFill Then
               
               aux = s.Fill.UniformColor.RGBValue
               
               For Each l In p.Layers
                    If l.Name = aux Then
                         exists = True
                         Set l = l
                         Exit For
                    End If
               Next l
               
               If exists Then
                    sr.Item(1).MoveToLayer l
               Else
                    Set l = p.CreateLayer(aux)
                    sr.Item(1).MoveToLayer l
               End If
               
          End If
          
          exists = False
          
     Next s
     
     Call DeleteLayers(p)
     
     cfg.TurnOn
     
     ActiveDocument.EndCommandGroup
     
     MsgBox "Finished", vbInformation, "ColorToLayer"
     
     Set s = Nothing
     Set sr = Nothing
     Set l = Nothing
     Set p = Nothing
     
End Sub

Public Sub DeleteLayers(ByRef page As page)
     
     ActiveDocument.BeginCommandGroup "DeleteLayers"
     
     Dim l As Layer
     
     For Each l In ActivePage.Layers
          If l.Shapes.Count = 0 And l.Index > 1 Then l.Delete
     Next l
     
     Set l = Nothing
     
     ActiveDocument.EndCommandGroup
     
End Sub

Private Sub ActivatePage()

     ActiveDocument.BeginCommandGroup "ActivatePage"

     Dim p As page
     Dim cfg As New clsSettings
     
     cfg.TurnOff
     
     For Each p In ActiveDocument.Pages
          Call DeleteLayers(p)
     Next p
     
     cfg.TurnOn
     
     Set p = Nothing
     
     ActiveDocument.EndCommandGroup

End Sub

Public Sub ActivateSettings()
     Dim cfg As New clsSettings
     cfg.TurnOff
     cfg.TurnOn
End Sub
