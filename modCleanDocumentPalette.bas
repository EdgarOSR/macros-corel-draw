Attribute VB_Name = "modCleanDocumentPalette"
Option Explicit

Public Sub RemoveColorsFromMainPalette()
     
     ActiveDocument.BeginCommandGroup "RemoveColors"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff
     
     Dim i As Integer
     
     i = ActiveDocument.Palette.ColorCount
     
     Do While i <> 0
          ActiveDocument.Palette.RemoveColor (i)
          i = i - 1
     Loop
     
     MsgBox "Job finished", vbInformation, "RemoveColors"
     
FinishSub:
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "RemoveColors"
     Debug.Print "RemoveColors", Err.Number, Err.Description, Err.Source
     GoTo FinishSub
     
End Sub

Public Sub ConvertToCMYK()

     ActiveDocument.BeginCommandGroup "ConvertToCMYK"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff

     Dim p As page
     Dim s As Shape

     For Each p In ActiveDocument.Pages
          For Each s In p.Shapes
               If s.Fill.Type = cdrUniformFill Then
                    s.Fill.UniformColor.ConvertTo (cdrColorCMYK)
               End If
          Next s
     Next p

     For Each p In ActiveDocument.Pages
          For Each s In p.Shapes
               If s.Fill.Type = cdrUniformFill Then
                    With s.Fill.UniformColor
                         If .RGBValue = 0 Or .RGBValue = 3486775 Or .RGBValue = 2697256 Then
                              .CMYKAssign 40, 0, 0, 100
                         End If
                    End With
               End If
          Next s
     Next p
     
     MsgBox "Job finished", vbInformation, "ConvertToCMYK"

FinishSub:
     Set p = Nothing
     Set s = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "ConvertToCMYK"
     Debug.Print "ConvertToCMYK", Err.Number, Err.Description, Err.Source
     GoTo FinishSub

End Sub

Public Sub ConvertLineColor()

     ActiveDocument.BeginCommandGroup "ConvertLineColor"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff

     Dim p As page
     Dim s As Shape

     For Each p In ActiveDocument.Pages
          For Each s In p.Shapes
               If s.Outline.Type = cdrOutline Then
                    With s.Outline
                         .Color.CMYKAssign 40, 0, 0, 100
                         .Width = ConvertUnits(1, cdrMillimeter, cdrInch)
                         .LineCaps = cdrOutlineRoundLineCaps
                         .LineJoin = cdrOutlineRoundLineJoin
                         .Justification = cdrOutlineJustificationMiddle
                         .MiterLimit = 45
                         .NibAngle = 0
                         .NibStretch = 100
                         .ScaleWithShape = False
                    End With
               End If
          Next s
     Next p

     MsgBox "Job finished", vbInformation, "ConvertLineColor"
     
FinishSub:
     Set p = Nothing
     Set s = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "ConvertLineColor"
     Debug.Print "ConvertLineColor", Err.Number, Err.Description, Err.Source
     GoTo FinishSub

End Sub

Public Sub ConvertLineWidth()

     ActiveDocument.BeginCommandGroup "ConvertLineWidth"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff

     Dim s As Shape
     Dim p As page
     Dim d As String, t As String

     t = InputBox("0 - Hairline" & vbCr & "1 - Milimeter")
     
     If t = "" Then GoTo FinishSub
          
     Do While CInt(t) < 0 Or CInt(t) > 1
          t = InputBox("0 - Hairline" & vbCr & "1 - Milimeter")
     Loop
     
     If CInt(t) = 0 Then
          d = ActiveDocument.ToUnits(CDbl(0.003), cdrMillimeter)
     Else
          d = InputBox("Line Width")
          If d = "" Then GoTo FinishSub
          d = ActiveDocument.ToUnits(CDbl(d), cdrMillimeter)
     End If
     
     For Each p In ActiveDocument.Pages
          For Each s In p.Shapes
               s.CreateSelection
               If s.Outline.Type = cdrOutline Then
                    With s.Outline
                         .Color.ConvertTo (cdrColorCMYK)
                         .Width = d
                         .LineCaps = cdrOutlineRoundLineCaps
                         .LineJoin = cdrOutlineRoundLineJoin
                         .Justification = cdrOutlineJustificationMiddle
                         .MiterLimit = 45
                         .NibAngle = 0
                         .NibStretch = 100
                         .ScaleWithShape = False
                    End With
               End If
          Next s
     Next p

     MsgBox "Job finished", vbInformation, "ConvertLineWidth"

FinishSub:
     Set p = Nothing
     Set s = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "ConvertLineWidth"
     Debug.Print "ConvertLineWidth", Err.Number, Err.Description, Err.Source
     GoTo FinishSub

End Sub

Private Sub RemovePowerClipX()

     ActiveDocument.BeginCommandGroup "RemovePowerClipX"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff

     Dim p As page
     Dim s As Shape
     Dim l As Layer
     Dim c As PowerClip

     For Each p In ActiveDocument.Pages
          For Each s In p.Shapes
               Set c = Nothing
               Set c = s.PowerClip
               If Not c Is Nothing Then
                    If c.Shapes.Count = 0 Then
                         c.Parent.CreateSelection
                         Application.FrameWork.Automation.InvokeItem ("7b022531-3cd7-487f-a797-9d80179dc821")
                         DoEvents
                    End If
               End If
          Next s
     Next p
     
FinishSub:
     Set p = Nothing
     Set s = Nothing
     Set l = Nothing
     Set c = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "RemovePowerClipX"
     Debug.Print "RemovePowerClipX", Err.Number, Err.Description, Err.Source
     GoTo FinishSub
     
End Sub

Public Sub RemovePowerClip()

     ActiveDocument.BeginCommandGroup "RemovePowerClip"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff

     Dim p As page
     Dim s As Shape
     Dim l As Layer
     Dim c As PowerClip

     For Each p In ActiveDocument.Pages
          
          p.Name = StrConv(p.Name, vbUpperCase)
     
          For Each l In p.Layers
               l.Name = StrConv(l.Name, vbUpperCase)
               l.Editable = True
          Next l
          
          For Each s In p.Shapes
               s.ConvertToCurves
               s.Name = StrConv(s.Name, vbLowerCase)
               
               Set c = Nothing

               If Not s.PowerClip Is Nothing Then
                    Set c = s.PowerClip
                    If c.Shapes.Count <> 0 Then
                         c.ExtractShapes
                    End If
               End If
          Next s
          
     Next p
     
     For Each p In ActiveDocument.Pages
     
          For Each s In p.Shapes
               s.ConvertToCurves
               s.Name = StrConv(s.Name, vbLowerCase)
               
               Set c = Nothing

               If Not s.PowerClip Is Nothing Then
                    Set c = s.PowerClip
                    If c.Shapes.Count <> 0 Then
                         c.ExtractShapes
                    End If
               End If
          Next s
     
     Next p
     
     Call RemovePowerClipX
     
     MsgBox "Job finished", vbInformation, "RemovePowerClip"
     
FinishSub:
     Set p = Nothing
     Set s = Nothing
     Set l = Nothing
     Set c = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "RemovePowerClip"
     Debug.Print "RemovePowerClip", Err.Number, Err.Description, Err.Source
     GoTo FinishSub

End Sub

Public Sub RemovePages()

     ActiveDocument.BeginCommandGroup "RemovePages"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff

     Dim p As page
     Dim s As Shape
     Dim l As Layer

     For Each p In ActiveDocument.Pages
          If p.Index = 1 Then
               For Each l In p.Layers
                    If l.Index > 2 Then p.Layers(l.Index).Delete
               Next l
          Else
               p.Shapes.All.Delete
               p.Delete
          End If
     Next p

     MsgBox "Job finished", vbInformation, "RemovePages"

FinishSub:
     Set p = Nothing
     Set s = Nothing
     Set l = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "RemovePages"
     Debug.Print "RemovePages", Err.Number, Err.Description, Err.Source
     GoTo FinishSub

End Sub


Public Sub CreateGuidelines()

     ActiveDocument.BeginCommandGroup "CreateGuidelines"
     
     If ActiveDocument Is Nothing Then GoTo FinishSub
     
     On Error GoTo ErrorHandler
     
     Dim cfg As New clsSettings
     cfg.TurnOff
     
     Dim h As String, d As String
     h = InputBox("Height")
     d = InputBox("Distance")
     
     If h = "" Or d = "" Then Exit Sub
     
     Dim x As Long, y As Long
     Dim s As Shape
     
     y = CDbl(h) / CDbl(d)
     
     Dim z As Double
     z = 0
     
     Dim p As page
     Set p = ActivePage
     
     d = ConvertUnits(d, cdrMillimeter, cdrInch)
     
     Do While (y + 1) > x
         Set s = ActiveDocument.MasterPage.GuidesLayer.CreateGuideAngle(p.CenterX, z, 0)
         z = z + d
         x = x + 1
     Loop
     
     MsgBox "Job finished", vbInformation, "CreateGuidelines"
  
FinishSub:
     Set p = Nothing
     Set s = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub

ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "CreateGuidelines"
     Debug.Print "CreateGuidelines", Err.Number, Err.Description, Err.Source
     GoTo FinishSub

End Sub
