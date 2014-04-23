Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DisplayUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.Controls
Imports System.Windows.Forms
Imports System.Windows.Forms.Cursors

Imports System
Imports Microsoft.Win32
Imports stdole
Public Class ISSI_Map
    Inherits ESRI.ArcGIS.Desktop.AddIns.Button
    

    Private m_hookhelper As IHookHelper

    Public Sub New()
        OnUpdate()
        CopyTemplatesLocal()
    End Sub
    
    Protected Overrides Sub OnClick()

        ArcMap.Application.CurrentTool = Nothing
        Dim pMxDoc As IMxDocument
        pMxDoc = ArcMap.Document
        Dim pActiveView As IActiveView
        pActiveView = ArcMap.Document.ActiveView

        If TypeOf pActiveView Is PageLayout Then

            ArcMap.Document.ActiveView = pMxDoc.FocusMap
            ArcMap.Document.ActiveView.Refresh()
            ArcMap.Application.RefreshWindow()

        End If

        Try
            If Not m_ISSIMapForm Is Nothing Then
                If m_ISSIMapForm.IsHandleCreated = False Then
                    m_ISSIMapForm = New FrmISSIMap

                    m_ISSIMapForm.SetApplication = Hook
                    ISSIMapFunctions.SetApplication = Hook

                    m_ISSIMapForm.Show()

                Else
                    If m_ISSIMapForm.WindowState = Windows.Forms.FormWindowState.Minimized Then
                        m_ISSIMapForm.WindowState = Windows.Forms.FormWindowState.Normal
                    Else
                        m_ISSIMapForm.Activate()
                    End If
                End If
            Else
                m_ISSIMapForm = New FrmISSIMap

                m_ISSIMapForm.SetApplication = Hook
                ISSIMapFunctions.SetApplication = Hook
                m_ISSIMapForm.ShowDialog()

            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "ISSIMap_Click", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try

    End Sub

    Protected Overrides Sub OnUpdate()
        Enabled = Not TypeOf ArcMap.Document.ActiveView Is PageLayout
    End Sub

  
End Class
