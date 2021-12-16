VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub TurnOff()
     ActiveDocument.BeginCommandGroup "TurnOff"
     ActiveDocument.Application.StatusBar.Visible = False
     ActiveDocument.Application.Unit = cdrCentimeter
     ActiveDocument.Application.Optimization = True
     ActiveDocument.EndCommandGroup
End Sub

Public Sub TurnOn()
     ActiveDocument.BeginCommandGroup "TurnOn"
     ActiveDocument.Application.StatusBar.Visible = True
     ActiveDocument.Application.Unit = cdrCentimeter
     ActiveDocument.Application.Optimization = False
     ActiveDocument.Application.Refresh
     ActiveDocument.Application.Windows.Refresh
     ActiveDocument.ActiveWindow.ActiveView.ToFitAllObjects
     ActiveDocument.EndCommandGroup
     ActiveDocument.EditAcrossLayers = True
     ActiveDocument.Pages(1).Activate
End Sub
