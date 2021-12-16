Attribute VB_Name = "modSvgDeal"
Option Explicit

Public Sub UnlockSvg()

     ActiveDocument.BeginCommandGroup ("UnlockSvg")
     
     Dim r As ShapeRange
     Dim r2 As ShapeRange
     Dim s As Shape
     Dim l As Layer
     Dim p As page
     Dim cfg As New clsSettings
     
     On Error GoTo ErrorHandler
     
     cfg.TurnOff
     
     Set p = ActiveDocument.ActivePage
     
     
     For Each l In p.Layers
          Set r = l.Shapes.All.BreakApartEx
          
          r.Unlock
          r.UngroupAllEx
          
          For Each s In r
               
               Set r2 = s.BreakApartEx
               r2.Unlock
               r2.UngroupAllEx
               
               Debug.Print vbNewLine
               Debug.Print "--------------------"
               Debug.Print s.Properties.Count
               Debug.Print "TopY: " & Round(s.TopY, 2)
               Debug.Print "RightX: " & Round(s.RightX, 2)
               Debug.Print "BottomY: " & Round(s.BottomY, 2)
               Debug.Print "LeftX: " & Round(s.LeftX, 2)
               Debug.Print "CenterX: " & Round(s.CenterX, 2)
               Debug.Print "CenterY: " & Round(s.CenterY, 2)
               Debug.Print "--------------------"
          Next s
     Next l

FinishSub:
     Set r = Nothing
     Set s = Nothing
     Set l = Nothing
     Set p = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub
     
ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "UnlockSvg"
     Debug.Print "UnlockSvg", Err.Number, Err.Description, Err.Source
     GoTo FinishSub
     
End Sub

Public Sub SavePDF()

     If ActiveDocument Is Nothing Then
          MsgBox "There is no file opened", vbExclamation, "ShapeProperties"
          Exit Sub
     End If
     
     ActiveDocument.BeginCommandGroup ("ShapeProperties")

     Dim path As String
     Dim exp As StructSaveAsOptions
     Dim d As Document
     Dim cfg As New clsSettings
     
     On Error GoTo ErrorHandler
     
     cfg.TurnOff
     
     Set d = ActiveDocument
     
     path = "C:\Users\" & Environ$("USERNAME") & "\Desktop\"
     
     path = Replace(path & d.FileName, ".cdr", ".pdf")
          
     Set exp = CreateStructSaveAsOptions
     exp.EmbedVBAProject = False
     exp.Filter = cdrPDF
     exp.IncludeCMXData = False
     exp.Range = cdrAllPages
     exp.EmbedICCProfile = False
     exp.Version = cdrVersion15
     
     With d.PDFSettings
          .ColorMode = pdfCMYK
          .JPEGQualityFactor = 25
          .BitmapCompression = pdfJPEG
          .TextExportMode = pdfTextAsUnicode
          .TextAsCurves = True
          .PageRange = cdrAllPages
          .PrintPermissions = pdfPrintPermissionHighResolution
          .pdfVersion = pdfVersionPDFX1a
          .EmbedFonts = True
          .EmbedBaseFonts = True
     End With
     
     
     d.SaveAs path, exp
     
     MsgBox "Job done" & vbCr & path, vbInformation, "ShapeProperties"
     
FinishSub:
     Set exp = Nothing
     Set d = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub
     
ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "ShapeProperties"
     Debug.Print "ShapeProperties", Err.Number, Err.Description, Err.Source
     GoTo FinishSub
     
End Sub

Private Sub TextCurves()
     
     If ActiveDocument Is Nothing Then
          MsgBox "There is no file opened", vbExclamation, "ShapeProperties"
          Exit Sub
     End If
     
     ActiveDocument.BeginCommandGroup ("TextCurves")

     Dim path As String
     Dim exp As StructSaveAsOptions
     Dim d As Document
     Dim cfg As New clsSettings
     Dim s As Shape
     Dim sr As ShapeRange
     Dim p As page
     
     On Error GoTo ErrorHandler
     
     cfg.TurnOff
     
     Set d = ActiveDocument
     
     For Each p In d.Pages
         
         Set sr = p.Shapes.FindShapes(Type:=cdrTextShape)
         
         For Each s In sr
             s.ConvertToCurves
         Next s
     Next p
     
     MsgBox "All text converted to curves", vbInformation, "TextCurves"
     
FinishSub:
     Set exp = Nothing
     Set d = Nothing
     cfg.TurnOn
     ActiveDocument.EndCommandGroup
     Exit Sub
     
ErrorHandler:
     MsgBox Err.Number & vbCr & Err.Description, vbCritical, "TextCurves"
     Debug.Print "TextCurves", Err.Number, Err.Description, Err.Source
     GoTo FinishSub
     
End Sub


