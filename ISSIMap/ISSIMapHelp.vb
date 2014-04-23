Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DisplayUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.Controls
Imports System.Windows.Forms
Imports System.Windows.Forms.Cursors

Imports System

Imports System.Diagnostics
Imports Microsoft.Win32
Imports stdole

Public Class ISSIMapHelp
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button

    Public Sub New()
        'gTemplateDirPath & "\ISSIMapV1p0.pdf"
        SetESRIAddInPath()
        SaveEmbeddedResourcesFile("ISSIMapV1p0.pdf", gTemplateDirPath & "\ISSIMapV1p0.pdf")
    End Sub

    Protected Overrides Sub OnClick()
        ArcMap.Application.CurrentTool = Nothing

        'open Help File PDF
        Dim pHelpFilePath As String = ""
        pHelpFilePath = gTemplateDirPath & "\ISSIMapV1p0.pdf"
        If Microsoft.VisualBasic.FileIO.FileSystem.FileExists(pHelpFilePath) Then
            'MessageBox.Show(pHelpFilePath, "HelpFile")

            System.Diagnostics.Process.Start(pHelpFilePath)
        End If

    End Sub

    Protected Overrides Sub OnUpdate()

    End Sub
   

End Class
