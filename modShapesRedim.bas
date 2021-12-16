Attribute VB_Name = "modShapesRedim"
Option Explicit

Public Sub ShapesRedim()
     
     ActiveDocument.BeginCommandGroup ("ShapesRedim")
     
     On Error GoTo ErrHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff
     
     Dim p As page
     Dim l As Layer
     Dim sr As ShapeRange
     Dim s As Shape
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     Set p = ActiveDocument.ActivePage
     
     For Each l In p.Layers
     
          Set sr = l.Shapes.All
          sr.ConvertToCurves
          
          For Each s In sr
               s.ConvertToCurves
               s.Name = "Shape_" & Format(s.ZOrder, "00")
          Next s
          
     Next l
     
     MsgBox "Job done", vbInformation, "ShapesRedim"
     
FinishSub:
     On Error Resume Next
     cfg.TurnOn
     Set p = Nothing
     Set l = Nothing
     Set sr = Nothing
     Set s = Nothing
     Set cfg = Nothing
     ActiveDocument.EndCommandGroup
     Exit Sub
     
ErrHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "ShapesRedim"
     Debug.Print "ShapesRedim > (" & Err.Number & ") " & Err.Description & vbCr & "Source: " & Err.Source
     GoTo FinishSub
     
End Sub

Public Sub ShapesCenter()
     
     ActiveDocument.BeginCommandGroup ("ShapesCenter")
     
     On Error GoTo ErrHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff
     
     Dim p As page
     Dim l As Layer
     Dim sr As ShapeRange
     Dim s As Shape
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     Set p = ActiveDocument.ActivePage
     
     For Each l In p.Layers
          For Each s In l.Shapes
               s.Name = Format(s.ZOrder, "000000")
               s.CenterX = Round(s.CenterX, 1)
               s.CenterY = Round(s.CenterY, 1)
          Next s
     Next l
     
     MsgBox "Job done", vbInformation, "ShapesRedim"
     
FinishSub:
     On Error Resume Next
     cfg.TurnOn
     Set p = Nothing
     Set l = Nothing
     Set sr = Nothing
     Set s = Nothing
     Set cfg = Nothing
     ActiveDocument.EndCommandGroup
     Exit Sub
     
ErrHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "ShapesCenter"
     Debug.Print "ShapesCenter > (" & Err.Number & ") " & Err.Description & vbCr & "Source: " & Err.Source
     GoTo FinishSub
     
End Sub

