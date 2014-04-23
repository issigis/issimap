Imports System
Imports System.IO
Imports System.Windows.Forms

Imports Microsoft.VisualBasic
Imports Microsoft.Win32

Imports ESRI.ArcGIS.SystemUI
Imports ESRI.ArcGIS.OutputUI
Imports ESRI.ArcGIS.OutputExtensions
Imports ESRI.ArcGIS.Output
Imports ESRI.ArcGIS.Geometry
Imports ESRI.ArcGIS.Framework
Imports ESRI.ArcGIS.esriSystem
Imports ESRI.ArcGIS.DisplayUI
Imports ESRI.ArcGIS.Display
Imports ESRI.ArcGIS.Desktop.AddIns
Imports ESRI.ArcGIS.DataSourcesRasterUI
Imports ESRI.ArcGIS.DataSourcesRaster
Imports ESRI.ArcGIS.DataSourcesOleDB
Imports ESRI.ArcGIS.DataSourcesNetCDF
Imports ESRI.ArcGIS.DataSourcesGDB
Imports ESRI.ArcGIS.DataSourcesFile
Imports ESRI.ArcGIS.Controls
Imports ESRI.ArcGIS.CatalogUI
Imports ESRI.ArcGIS.Catalog
Imports ESRI.ArcGIS.CartoX
Imports ESRI.ArcGIS.CartoUI
Imports ESRI.ArcGIS.Carto
Imports ESRI.ArcGIS.ArcMapUI
Imports ESRI.ArcGIS.ArcMap
Imports ESRI.ArcGIS.ArcCatalogUI



Module ISSIMapFunctions

    Public gTemplateDirPath As String
    Public esriAddinPath As String

    Public m_ISSIMapForm As FrmISSIMap
    Public MapTemplateMxd As String
    Public gTemplateDir As String

    Public gISSIApp As IApplication             'Public gISSIApp As IApplication ' set as global
    Public gISSIMxApp As IMxApplication            'Public gmxApp As IMxApplication
    Public gISSIMxDoc As IMxDocument               'Public gISSIMxDoc As IMxDocument
    Public m_pPageLayout As IPageLayout3

    Dim pMxDoc As IMxDocument
    Public pMxApp As IMxApplication

    Public pSelectionTracker As ISelectionTracker

    Public gMainMapName As String
    Public gIndexMapName As String
    Public boolLogoSelected As Boolean
    Public gLogoFullPath As String

    Dim pUnusedDataFrameName As String

    Public strPageSizeLetter As String
    Public strISSILogo As String
    Public strISSIFullName As String
    Public gLogoImgDir As String

    Public WriteOnly Property SetApplication() As IApplication
        Set(value As IApplication)
            gISSIApp = value
            gISSIMxApp = value
            gISSIMxDoc = CType(gISSIApp.Document, IMxDocument)


        End Set
    End Property



    Public Sub ISSIMapAction(Template, pTemplateDirectory, PageSize, Title, Author, Project, Client, pDateString, Notes, strLogoSelected)
       
        Dim pActiveView As IActiveView
        Dim pPageLayout As IPageLayout
        Dim pImageFilePath As String    'filestring to image to load
        Dim pImgBoxname As String


        ''SET APPLICATION VARIABLES             
        pMxApp = gISSIApp   ''As IMxApplication
        pMxDoc = gISSIApp.Document ''as IMxDocument

        pActiveView = pMxDoc.ActiveView

        pPageLayout = pMxDoc.PageLayout

        strPageSizeLetter = Left(PageSize, 1)
        strISSILogo = "ISSIMap_" & strPageSizeLetter & ".png"
        strISSIFullName = gTemplateDir & "\" & strISSILogo
      
        SetMainAsFocus()

        pActiveView = gISSIMxDoc.PageLayout

        gISSIMxDoc.UpdateContents()
        gISSIMxDoc.ActiveView.Refresh()

        Dim pTextSize As Double = 8.0

        If pMxDoc.Maps.Count > 1 And gIndexMapName = "None" Then '' need to hide the extra dataframe
            SetUnusedDataFrame()
        End If


        ''LOAD USER LOGO IF ONE WAS SELECTED
        If boolLogoSelected Then

            pImageFilePath = gLogoFullPath
            pImgBoxname = "LogoDisclaimerBox"


            If Len(pImageFilePath) <> 0 Then

                InsertLogo(pImgBoxname, pImageFilePath)
              
            End If
        End If


        '  'SET TEXT ELEMENTS IN LAYOUT
        SetUserText(Title, Author, Project, Client, pDateString, Notes)


        '  'SCALE PAGE SIZE BASED ON FORM
        'ScalePageSize(PageSize)


        If pMxDoc.FocusMap.LayerCount > 0 Then
            AddProjectionInfo()
        End If



        'UNSELECT ALL ELEMENTS AND REFRESH VIEW
        Dim mActiveview As IActiveView
        mActiveview = pMxDoc.PageLayout

        pMxDoc.ActiveView = mActiveview
        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphicSelection, Nothing, Nothing)
        Dim mGraphicsContainer As IGraphicsContainer
        mGraphicsContainer = pMxDoc.PageLayout
        Dim mGCSelect As IGraphicsContainerSelect
        mGCSelect = mGraphicsContainer
        mGCSelect.UnselectAllElements()

        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()


        '   ' FIT  SCALE BAR 
        FitScaleBar()


        ScaleCenterLegend()
        Dim pScale As String = Left(PageSize, 1)
        FormatLegendByItem(pScale)

        ForceFitLegend()

        ''UNSELECT ALL ELEMENTS AND REFRESH VIEW

        mActiveview = pMxDoc.PageLayout
        pMxDoc.ActiveView = mActiveview
        pActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphicSelection, Nothing, Nothing)


        mGraphicsContainer = pPageLayout
        mGCSelect = mGraphicsContainer
        mGCSelect.SelectAllElements()
        mActiveview.Refresh()
        mGCSelect.UnselectAllElements()
        mActiveview.Refresh()

        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()

        pPageLayout.ZoomToWhole()
    End Sub

    Public Sub FitScaleBar()

        pMxDoc = gISSIApp.Document
        Dim pPageLayout As IPageLayout
        Dim pPageLayoutGraphicsContainer As IGraphicsContainer
        Dim pPagelayoutDisplay As IDisplay
        Dim pElement As IElement
        Dim pElementProperties As IElementProperties
        Dim pScaleBoxEnvelope As IEnvelope2 = New Envelope
        Dim pScaleBarEnvelope As IEnvelope2 = New Envelope

        pPageLayout = pMxDoc.PageLayout
        pPageLayoutGraphicsContainer = pPageLayout
        pPagelayoutDisplay = pMxDoc.ActiveView.ScreenDisplay
        Try


            'Check for each element in the layout that needs to be manipulated

            pPageLayoutGraphicsContainer.Reset()
            pElement = pPageLayoutGraphicsContainer.Next

            Do While Not pElement Is Nothing
                pElementProperties = pElement

                If pElementProperties.Name = "ScaleBarBox" Then

                    pElement.QueryBounds(pPagelayoutDisplay, pScaleBoxEnvelope)
                    pScaleBoxEnvelope.Expand(0.8, 0.8, True)


                End If

                If TypeOf pElementProperties Is IMapSurroundFrame Then

                    If (pElementProperties.Name Like "*Scale*Bar") Then

                        Dim pScaleBarSurround As IMapSurroundFrame
                        pScaleBarSurround = pElementProperties
                        Dim pScaleBarObject As IScaleBar
                        pScaleBarObject = pScaleBarSurround.Object
                        If pScaleBarObject.LabelSymbol.Size < 6 Then
                            pScaleBarObject.LabelSymbol.Size = 6.0
                        End If
                       

                    Else


                        pElement.Geometry = pScaleBoxEnvelope ' this was making scale bar to FAT


                    End If ' if element is scalebar

                End If ' if is IMapsurroundframe


                    pElement = pPageLayoutGraphicsContainer.Next
            Loop

        Catch ex As Exception

            'MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "FitScaleBar", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try
    End Sub

    Public Sub ForceFitLegend()


        Dim pPageLayoutDisplay As IDisplay

        Dim pPLayout As IPageLayout3
        pPLayout = pMxDoc.PageLayout

        pPageLayoutDisplay = pMxDoc.ActiveView.ScreenDisplay

        Dim pLegendBoundBox As ienvelope2 = New Envelope
        pLegendBoundBox = GetEnvelopeByName("LegendBox")

        ''GET Upperleft Anchor point
        Dim pBoundBoxWidth As Double = pLegendBoundBox.Width
        Dim pBoundBoxHeight As Double = pLegendBoundBox.Height

        Dim pLegEnv As IEnvelope2 = New Envelope
        Dim pLegNewEnv As IEnvelope2 = New Envelope

        Dim pLegendMapFrame As IMapSurroundFrame
        pLegendMapFrame = GetLegendObject()
        pLegendMapFrame.MapSurround.QueryBounds(pPageLayoutDisplay, pLegEnv, pLegEnv)
      
        Dim pNewPageSizeIndex As Integer = 0
        Dim pNewPageSizeLetter As String = "A"

        If pLegEnv.Height > pLegendBoundBox.Height Then

            If pLegEnv.Height / pLegendBoundBox.Height > 1.3 And m_ISSIMapForm.PageSizeListBox.SelectedIndex > 1 Then
                pNewPageSizeIndex = m_ISSIMapForm.PageSizeListBox.SelectedIndex - 2
                pNewPageSizeLetter = Left(m_ISSIMapForm.PageSizeListBox.Items.Item(pNewPageSizeIndex), 1)
            ElseIf pLegEnv.Height / pLegendBoundBox.Height > 1 And m_ISSIMapForm.PageSizeListBox.SelectedIndex > 0 Then
                pNewPageSizeIndex = m_ISSIMapForm.PageSizeListBox.SelectedIndex - 1
                pNewPageSizeLetter = Left(m_ISSIMapForm.PageSizeListBox.Items.Item(pNewPageSizeIndex), 1)
            Else
                pNewPageSizeIndex = 0
                pNewPageSizeLetter = "A"
            End If
            MessageBox.Show("About to resize legend for pagesize " & pNewPageSizeLetter, "ForceFitLegend")
            FormatLegendByItem(pNewPageSizeLetter)
        End If


       

    End Sub


    Public Function CreateNewEnvelope(pWidth As Double, pHeight As Double, pEnvCenter As IPoint) As IEnvelope2

     
        Dim pNewImageEnvelope As ienvelope2 = New Envelope
        Dim pEnvWidth As Double
        Dim pEnvHeight As Double

        Dim pLowerLeftX As Double
        Dim pLowerLeftY As Double
        Dim pUpperRightX As Double
        Dim pUpperRightY As Double

        Try
            pEnvWidth = pWidth
            pEnvHeight = pHeight

            pLowerLeftX = pEnvCenter.X - (pEnvWidth / 2)
            pLowerLeftY = pEnvCenter.Y - (pEnvHeight / 2)
            pUpperRightX = pEnvCenter.X + (pEnvWidth / 2)
            pUpperRightY = pEnvCenter.Y + (pEnvHeight / 2)


            pNewImageEnvelope.PutCoords(pLowerLeftX, pLowerLeftY, pUpperRightX, pUpperRightY)

        Catch ex As Exception
            pNewImageEnvelope = Nothing
        End Try


        Return pNewImageEnvelope

    End Function



    Public Sub ScaleCenterLegend()

        ''THIS MAKES IT SMALLER NOT LARGER
        Dim m_pMxDoc As IMxDocument
        Dim pActiveView As IActiveView
        Dim pPageLayout As IPageLayout
        Dim pPageLayoutGraphicsContainer As IGraphicsContainer
        Dim pDisplay As IScreenDisplay

        Dim pLElement As IElement
        Dim pLElementProperties As IElementProperties3

        Dim pAdjustedElement As IElement 'for applying the transform scale and move
        Dim pLegendBoxEnv As IEnvelope = New Envelope
        Dim pLegendEnvelope As IEnvelope = New Envelope
        Dim pMapSurroundFrame As IMapSurroundFrame
        Dim pLegendSurroundFrame As IMapSurroundFrame = Nothing
        Dim pLegendObject As ILegend3
        Dim pTransform2D As ITransform2D
        Dim pLegUpperLeft As IPoint = New Point
        m_pMxDoc = gISSIMxDoc
        'need to loop thru items on layout to find thelegend frame - mapsurround
        pActiveView = m_pMxDoc.ActiveView

        pPageLayout = m_pMxDoc.PageLayout
        pPageLayoutGraphicsContainer = pPageLayout
        pDisplay = m_pMxDoc.ActiveView.ScreenDisplay

        pPageLayoutGraphicsContainer.Reset()
        pLElement = pPageLayoutGraphicsContainer.Next

        'NEED TO FIND LEGEND BOX TO GET RESIZE ENVELOPE
        Do While Not pLElement Is Nothing
            pLElementProperties = pLElement
            If pLElementProperties.Name = "LegendBox" Then

                pLElement.QueryBounds(pDisplay, pLegendBoxEnv)

                pLegendBoxEnv.Expand(0.95, 0.95, True)
                pLegUpperLeft = pLegendBoxEnv.UpperLeft
            End If

            pLElement = pPageLayoutGraphicsContainer.Next
        Loop

        m_pMxDoc.ActiveView.Refresh()

        ''NOW SHOULD HAVE LEGEND ENVELOPE FOR POSITION AND SIZE MAX
        'Need Center of envelope
        Dim pLegendCenter As IPoint = New Point

        Dim pLegEnvCenterX As Double
        Dim pLegEnvCenterY As Double
        pLegEnvCenterX = pLegendBoxEnv.XMin + ((pLegendBoxEnv.XMax - pLegendBoxEnv.XMin) / 3)
        pLegEnvCenterY = pLegendBoxEnv.YMin + ((pLegendBoxEnv.YMax - pLegendBoxEnv.YMin) / 2)



        pLegendCenter.PutCoords(pLegEnvCenterX, pLegEnvCenterY)


        ''FIND THE LEGEND AND SET ITS ENVELOPE WITHIN THE LEGENDBOX

        pPageLayoutGraphicsContainer.Reset()
        pLElement = pPageLayoutGraphicsContainer.Next

        Do While Not pLElement Is Nothing
            pLElementProperties = pLElement
            
            If TypeOf pLElementProperties Is IMapSurroundFrame Then
                pMapSurroundFrame = pLElement
                If TypeOf pMapSurroundFrame.MapSurround Is ILegend Then
                    
                    pLegendEnvelope = New Envelope
                    pLElement.QueryBounds(pDisplay, pLegendEnvelope)

                    pLegendSurroundFrame = pLElement
                    pLegendObject = pLegendSurroundFrame.MapSurround

                    Dim pEnvHR As Double
                    pEnvHR = (pLegendBoxEnv.Height) / (pLegendEnvelope.Height)

                   
                    pLegendEnvelope.Expand(pEnvHR, pEnvHR, True)


                    Dim pEnvWR As Double
                    pEnvWR = (pLegendBoxEnv.Width) / (pLegendEnvelope.Width)
                    pLegendEnvelope.Expand(pEnvWR, pEnvWR, True)

                    Dim pNewEnv As IEnvelope = New Envelope
                    pNewEnv.UpperLeft = pLegUpperLeft

                    Dim pNewLowerRight As IPoint = New Point
                    pNewLowerRight.PutCoords(pLegUpperLeft.X + pLegendEnvelope.Width, pLegUpperLeft.Y - pLegendEnvelope.Height)
                    pNewEnv.LowerRight = pNewLowerRight


                    pLElement.Geometry = pNewEnv
                    pLElementProperties.AnchorPoint = esriAnchorPointEnum.esriTopLeftCorner


                    pPageLayoutGraphicsContainer.UpdateElement(pLElement)

                    pAdjustedElement = pLegendSurroundFrame
                    pTransform2D = pLElement

                    With pTransform2D
                        '.Scale(pLegendCenter, 0.8, 0.8)
                        .Move(0.01, 0.01)
                        .Move(-0.01, -0.01)

                    End With

                    pLElement = pTransform2D
                    pPageLayoutGraphicsContainer.UpdateElement(pLElement)

                    ''this section was commented out
                    Dim pLayoutTrackcancel As ITrackCancel
                    ''Have to create "canceltracker" to use "Draw" Command
                    pLayoutTrackcancel = New CancelTracker
                    ''**
                    pLElement.Geometry.Envelope.CenterAt(pLegendCenter)
                    pPageLayoutGraphicsContainer.UpdateElement(pLElement)
                    ''**
                    ''Draws the ELEMENT into the given display object
                    pLElement.Draw(pDisplay, pLayoutTrackcancel)
                    ''Draws the MAPSURROUND into the specified display bounds
                    pLegendObject.Draw(pDisplay, pLayoutTrackcancel, pNewEnv)

                    pLegendObject.FitToBounds(pDisplay, pNewEnv, True)

                    pLegendObject.Refresh()

                End If ' if is a legend
            End If ' if is a mapsurround

            pLElement = pPageLayoutGraphicsContainer.Next
        Loop 'while pLElement is not nothing - setting legend geometry

        If Not pLegendSurroundFrame Is Nothing Then

            pAdjustedElement = pLegendSurroundFrame
            pTransform2D = pAdjustedElement

            With pTransform2D
                .Move(0.1, 0.15)
                .Move(-0.1, -0.15)
            End With
        End If

        pPageLayoutGraphicsContainer.UpdateElement(pLElement)
        m_pMxDoc.UpdateContents()
        m_pMxDoc.ActiveView.Refresh()

    End Sub

   

    Private Sub RefreshLegend(ByRef pElement As IElement, pAV As IActiveView, bRedraw As Boolean)
        Dim pMSF As IMapSurroundFrame
        Dim pMS As IMapSurround
        Dim pLegend As ILegend3
        Dim pNewEnv As ienvelope2 = New Envelope


        pMSF = pElement
        pLegend = pMSF.MapSurround
        pNewEnv = New Envelope
        pElement = pMSF
        pMS = pMSF.MapSurround
        pMS.Refresh()
        pLegend.QueryBounds(pAV.ScreenDisplay, pElement.Geometry.Envelope, pNewEnv)
        'Assign the boundary to the map surround frame
        If bRedraw Then
            pAV.PartialRefresh(esriViewDrawPhase.esriViewGraphics, pElement, Nothing)
            pElement.Draw(pAV.ScreenDisplay, Nothing)
        End If
        pElement.Geometry = pNewEnv

        pMS.FitToBounds(pAV.ScreenDisplay, pElement.Geometry.Envelope, True)
        pMS.Refresh()
        If bRedraw Then
            pAV.PartialRefresh(esriViewDrawPhase.esriViewGraphics, pElement, Nothing)
        End If
    End Sub

    Public Function GetEnvelopeByName(ByVal pFindGraphicName) As IEnvelope2

        pMxDoc.ActiveView = pMxDoc.PageLayout
        Dim pGEEnv As ienvelope2 = New Envelope

        Dim pPageLayout As IPageLayout3
        Dim pPLGC As IGraphicsContainer
        Dim pPageLayoutDisplay As IDisplay = pMxDoc.ActiveView.ScreenDisplay
        pPageLayoutDisplay = pMxDoc.ActiveView.ScreenDisplay
        Dim pGraphicElement As IElement
        Dim pGraphicElementProps As IElementProperties3

        Dim boolLegendBox As Boolean = False

        Dim pGraphicElementName As String = ""
        pPageLayout = pMxDoc.PageLayout

        pPLGC = pMxDoc.PageLayout

        pPLGC.Reset()
        pGraphicElement = pPLGC.Next
        Do While Not pGraphicElement Is Nothing
            pGraphicElementProps = pGraphicElement
            If pGraphicElementProps.Name <> "" Then
                pGraphicElementName = pGraphicElementProps.Name
                If pGraphicElementName = pFindGraphicName Then
                    pGraphicElement.QueryBounds(pPageLayoutDisplay, pGEEnv)
                    Exit Do
                End If

            End If
            pGraphicElement = pPLGC.Next
        Loop

        Return pGEEnv

    End Function

    Public Function GetLegendObject() As IMapSurroundFrame
        Dim pGelement As IElement = Nothing
        Dim pElementProperties As IElementProperties
        Dim pMSFrame As IMapSurroundFrame = New MapSurroundFrame

        Dim pLegend As ILegend

        Dim pActiveView As IActiveView
        Dim pPageLayout As IPageLayout
        Dim pGC As IGraphicsContainer
        Dim pScreenDisplay As IDisplay
        Dim boolfoundlegend As Boolean = False
        Try

            'need to loop thru items on layout to find thelegend frame - mapsurround
            pActiveView = pMxDoc.ActiveView
            pPageLayout = pMxDoc.PageLayout
            pGC = pPageLayout
            pScreenDisplay = pMxDoc.ActiveView.ScreenDisplay

            pGC.Reset()
            pGelement = pGC.Next

            Do While Not pGelement Is Nothing

                pElementProperties = pGelement


                If TypeOf pElementProperties Is IMapSurroundFrame Then
                    pMSFrame = pGelement
                    If TypeOf pMSFrame.MapSurround Is ILegend Then
                       
                        pMSFrame = pGelement

                        pLegend = pMSFrame.MapSurround

                        boolfoundlegend = True

                        
                        Exit Try

                    End If
                End If

                pGelement = pGC.Next
            Loop 'while pelement is not nothing - setting legend geometry

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "GetLegendObject", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try



        If pMSFrame.Object Is Nothing Then
            Return Nothing
        Else

            Return pMSFrame
        End If

    End Function ' GetLegend
    Public Sub FormatLegendByItem(aLegendScale) 'should be the legendsurround
       

        pMxDoc = gISSIApp.Document
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()

        Dim pPageLayout As IPageLayout
        Dim pPagelayoutDisp As IDisplay
        Dim pPageLayoutGC2 As IGraphicsContainer
        Dim pPLGCSelect As IGraphicsContainerSelect

        Dim pGraphicElement As IElement
        Dim pGElementProp As IElementProperties3
        Dim pTransform2d As ITransform2D

        Dim pMapSurroundFrame As IMapSurroundFrame = Nothing
        Dim pMapSurround As IMapSurround
        Dim pLegend As ILegend3 = Nothing

        Dim pLegEnv As IEnvelope2 = New Envelope


        Dim pLegendBoxEnvelope As IEnvelope2 = New Envelope
        pLegendBoxEnvelope = GetEnvelopeByName("LegendBox")
        Dim pLegendBoxUpperLeft As IPoint = New Point

        pLegendBoxUpperLeft.PutCoords(pLegendBoxEnvelope.UpperLeft.X + (pLegendBoxEnvelope.Width * 0.05), pLegendBoxEnvelope.UpperLeft.Y - (pLegendBoxEnvelope.Height * 0.05))
        
        pPageLayout = pMxDoc.PageLayout
        pPageLayoutGC2 = pMxDoc.PageLayout
        pPagelayoutDisp = pMxDoc.ActiveView.ScreenDisplay

        pPLGCSelect = pPageLayoutGC2
        pPLGCSelect.SelectAllElements()
        pPLGCSelect.UnselectAllElements()

        pMxDoc.ActiveView.Refresh()

        pPageLayoutGC2.Reset()
        pGraphicElement = pPageLayoutGC2.Next

        Do Until pGraphicElement Is Nothing

            pGElementProp = pGraphicElement


            If TypeOf pGraphicElement Is IMapSurroundFrame Then
                pMapSurroundFrame = pGraphicElement
                pMapSurround = pMapSurroundFrame.MapSurround

                If TypeOf pMapSurroundFrame.MapSurround Is ILegend Then
                    pTransform2d = pGraphicElement

                    pGElementProp = pGraphicElement
                    pGElementProp.AnchorPoint = esriAnchorPointEnum.esriTopLeftCorner

                    pGraphicElement.Geometry.Envelope.UpperLeft = pLegendBoxUpperLeft

                    pLegend = pMapSurroundFrame.MapSurround
                    pLegend.QueryBounds(pPagelayoutDisp, pLegEnv, pLegEnv)

                    pLegEnv.UpperLeft.PutCoords(pLegendBoxEnvelope.UpperLeft.X + 0.15, pLegendBoxEnvelope.UpperLeft.Y - 0.15)
                    pGraphicElement.Geometry = pLegEnv

                    pMxDoc.UpdateContents()
                    pMxDoc.ActiveView.Refresh()



                    '' LOOP THRU THE FEATURELAYERS TO GET THECLASSIFICATION FOR SETTING THE LEGENDFORMAT

                    Dim pLegendItem As ILegendItem
                    Dim pLayer As ILayer             ''the layer in the  legend to compare to the legenditem
                    Dim pLayerName As String
                    Dim pLabelSymbol As ITextSymbol

                    Dim fMapLayers As IEnumLayer 'for looping thru focusmap layers
                    Dim lngCount As Integer = 0

                    For lngCount = 0 To pLegend.ItemCount - 1 ' lngcount -the # of layers

                        ''LOOP THRU TE LEGEND ITEMS AND SET TEXT AND PATCH PROPERTIES
                        pLegendItem = pLegend.Item(lngCount)

                        pLegendItem.ShowHeading = False
                        pLegendItem.ShowDescriptions = False
                
                        pLabelSymbol = pLegendItem.LegendClassFormat.LabelSymbol
                        pLabelSymbol.Font.Bold = False

                        Select Case aLegendScale
                            Case "A"
                                pLabelSymbol.Size = 7.0#
                                pLegendItem.LegendClassFormat.PatchHeight = 7.0
                                pLegendItem.LegendClassFormat.PatchWidth = 12.0
                                pLegend.Format.LayerNameGap = 4.0
                                pLegend.Format.VerticalItemGap = 4.0
                                pLegend.Format.VerticalPatchGap = 3.0

                                pLegend.Format.HeadingGap = 4.0
                                pLegend.Format.TextGap = 8.0
                                pLegend.Format.HorizontalItemGap = 8.0
                                pLegend.Format.HorizontalPatchGap = 8.0

                            Case "B"
                                pLabelSymbol.Size = 10.0#
                                pLegendItem.LegendClassFormat.PatchHeight = 9.0
                                pLegendItem.LegendClassFormat.PatchWidth = 14.0
                                pLegend.Format.LayerNameGap = 6.0
                                pLegend.Format.VerticalItemGap = 5.0
                                pLegend.Format.VerticalPatchGap = 4.0

                                pLegend.Format.HeadingGap = 6.0
                                pLegend.Format.TextGap = 10.0
                                pLegend.Format.HorizontalItemGap = 10.0
                                pLegend.Format.HorizontalPatchGap = 10.0

                            Case "C"
                                pLabelSymbol.Size = 12.0#
                                pLegendItem.LegendClassFormat.PatchHeight = 10.0
                                pLegendItem.LegendClassFormat.PatchWidth = 16.0
                                pLegend.Format.LayerNameGap = 8.0
                                pLegend.Format.VerticalItemGap = 7.0
                                pLegend.Format.VerticalPatchGap = 6.0

                                pLegend.Format.HeadingGap = 6.0
                                pLegend.Format.TextGap = 10.0
                                pLegend.Format.HorizontalItemGap = 10.0
                                pLegend.Format.HorizontalPatchGap = 10.0
                            Case "D"
                                pLabelSymbol.Size = 14.0#
                                pLegendItem.LegendClassFormat.PatchHeight = 12.0
                                pLegendItem.LegendClassFormat.PatchWidth = 18.0
                                pLegend.Format.LayerNameGap = 9.0
                                pLegend.Format.VerticalItemGap = 8.0
                                pLegend.Format.VerticalPatchGap = 7.0

                                pLegend.Format.HeadingGap = 8.0
                                pLegend.Format.TextGap = 12.0
                                pLegend.Format.HorizontalItemGap = 12.0
                                pLegend.Format.HorizontalPatchGap = 12.0
                            Case "E"
                                pLabelSymbol.Size = 18.0#
                                pLegendItem.LegendClassFormat.PatchHeight = 16.0
                                pLegendItem.LegendClassFormat.PatchWidth = 24.0
                                pLegend.Format.LayerNameGap = 14.0
                                pLegend.Format.VerticalItemGap = 12.0
                                pLegend.Format.VerticalPatchGap = 10.0

                                pLegend.Format.HeadingGap = 10.0
                                pLegend.Format.TextGap = 14.0
                                pLegend.Format.HorizontalItemGap = 14.0
                                pLegend.Format.HorizontalPatchGap = 14.0

                        End Select


                        pLegendItem.LayerNameSymbol = pLabelSymbol
                        pLegendItem.LegendClassFormat.LabelSymbol = pLabelSymbol

                        ''FIND THE LAYER FOR THE LEGEND ITEM  TO DETERMINE CLASSIFICATION
                        pLayerName = pLegendItem.Layer.Name

                        fMapLayers = pMxDoc.FocusMap.Layers
                        fMapLayers.Reset()
                        pLayer = fMapLayers.Next
                        Do While Not pLayer Is Nothing
                            If TypeOf pLayer Is IFeatureLayer Then
                                If pLayer.Name = pLayerName Then
                                    
                                    Dim pTheLayer As IFeatureLayer   ''the layer in the FocusMap
                                    pTheLayer = pLayer
                                    Dim pGeoFeatureLayer As IGeoFeatureLayer
                                    pGeoFeatureLayer = pTheLayer
                                    Dim pRenderer As IFeatureRenderer
                                    pRenderer = pGeoFeatureLayer.Renderer


                                    If TypeOf pRenderer Is SimpleRenderer Then
                                        pLegendItem.ShowLayerName = False
                                        pLegendItem.LegendClassFormat.LabelSymbol = pLabelSymbol

                                    ElseIf TypeOf pRenderer Is UniqueValueRenderer Then

                                        pLegendItem.ShowLayerName = True

                                        pLegendItem.LegendClassFormat.LabelSymbol = pLabelSymbol
                                        pLegendItem.LayerNameSymbol = pLabelSymbol


                                    End If  'IF IS CLASSIFIED
                                End If  'if current legend item name matches current layer in data frame
                            Else
                                'IF NOT FEATURELAYER THEN DO NOT DISPLAY IN LEGEND

                            End If 'if is featurelayer

                            pLayer = fMapLayers.Next
                        Loop 'DO While player is not nothing
                    Next lngCount

                    pLegend.Format.TitleGap = 1
                    pLegend.Format.ShowTitle = False

                    pLegend.Format.GroupGap = 4.0
                   
                    ''NEED TO REFRESH LEGEND GRAPHIC must get/refresh MapSurround AND MapSurroundFrame

                    pLegend.Refresh()
                    pLegend.QueryBounds(pPagelayoutDisp, pLegEnv, pLegEnv)


                    pLegEnv.UpperLeft = pLegendBoxUpperLeft
                   

                    pMapSurround.FitToBounds(pPagelayoutDisp, pLegEnv, True)
                    pMapSurround.Refresh()

                    pMapSurroundFrame.MapSurround = pMapSurround
                    pMapSurroundFrame.MapSurround.FitToBounds(pPagelayoutDisp, pLegEnv, True)
                    pMxDoc.UpdateContents()

                    pGraphicElement.Geometry = pLegEnv

                    Dim pMSEnv As IEnvelope2 = New Envelope

                    pLegend.QueryBounds(pPagelayoutDisp, pLegEnv, pLegEnv)


                    pMapSurroundFrame.MapSurround.QueryBounds(pPagelayoutDisp, pMSEnv, pMSEnv)
                   

                    pMxDoc.UpdateContents()
                    pMxDoc.ActiveView.Refresh()

                   
                 
                End If  'If TypeOf pMapSurroundFrame.MapSurround Is ILegend Then

            End If      ' If TypeOf pGraphicElement Is IMapSurroundFrame Then
            pGraphicElement = pPageLayoutGC2.Next
        Loop

        pPLGCSelect.SelectAllElements()
        pPLGCSelect.UnselectAllElements()

        pMxDoc.ActiveView.Refresh()

    End Sub

    'Public Sub FormatLegend()
    '    ''THIS is NOT USED
    '    Dim pLegend As ILegend
    '    Dim pPageLayout As IPageLayout3
    '    Dim pActiveView As IActiveView
    '    Dim plegendEnvelope As ienvelope2 = New Envelope

    '    Dim pLegSurroundFrame As IMapSurroundFrame

    '    Dim pEnumElement As IElement
    '    Dim pEnumElemProps As IElementProperties3

    '    Dim p_LegendSurround As IMapSurround
    '    Dim pPageLayoutGC As IGraphicsContainer
    '    Dim pPageLayoutScreenDisplay As IScreenDisplay
    '    'need to loop thru items on layout to find thelegend frame - mapsurround

    '    pActiveView = pMxDoc.ActiveView
    '    pPageLayout = pMxDoc.PageLayout
    '    pPageLayoutGC = pPageLayout
    '    pPageLayoutScreenDisplay = pMxDoc.ActiveView.ScreenDisplay

    '    pPageLayoutGC.Reset()
    '    pEnumElement = pPageLayoutGC.Next
    '    Do While Not pEnumElement Is Nothing
    '        pEnumElemProps = pEnumElement
    '        If pEnumElemProps.Name = "Legend" And TypeOf pEnumElemProps Is IMapSurroundFrame Then
    '            ''MessageBox.Show("Found the Legend", "FormatLegend")

    '            pLegSurroundFrame = pEnumElement
    '            p_LegendSurround = pLegSurroundFrame.MapSurround

    '            Dim m_LegSurroundEnv As ienvelope2 = New Envelope
    '            pEnumElement.QueryBounds(pPageLayoutScreenDisplay, m_LegSurroundEnv)

    '            pLegend = pLegSurroundFrame.Object

    '            If TypeOf pLegend Is Legend Then

    '                pLegend.Format.ShowTitle = False
    '                pLegend.Format.TitleGap = 2.0#
    '                'MessageBox.Show("About to set Patch Height 12 and Width 16", "FormatLegend")
    '                If m_ISSIMapForm.PageSizeListBox.SelectedIndex < 2 Then
    '                    pLegend.Format.DefaultPatchHeight = 8.0#
    '                    pLegend.Format.DefaultPatchWidth = 14.0#
    '                Else
    '                    pLegend.Format.DefaultPatchHeight = 12.0#
    '                    pLegend.Format.DefaultPatchWidth = 18.0#
    '                End If

    '                pLegend.Format.LayerNameGap = 4.0#
    '                pLegend.Format.GroupGap = 4.0#
    '                pLegend.Format.HeadingGap = 5.0#

    '                pLegend.Format.HorizontalItemGap = 6.0# 'between columns?
    '                pLegend.Format.HorizontalPatchGap = 8.0# 'between patch and desciption?


    '                pLegend.Format.TextGap = 10.0#
    '                pLegend.Format.VerticalItemGap = 4.0#
    '                pLegend.Format.VerticalPatchGap = 3.0#

    '                pLegend.AutoReorder = True
    '                pLegend.AutoAdd = True
    '                pLegend.AutoVisibility = True

    '                pLegend.Title = "Legend"
    '                pLegend.Refresh()

    '                pLegend.QueryBounds(pPageLayoutScreenDisplay, m_LegSurroundEnv, plegendEnvelope)

    '                p_LegendSurround.FitToBounds(pPageLayoutScreenDisplay, plegendEnvelope, True)
    '                p_LegendSurround.Refresh()
    '                pEnumElement.Geometry = plegendEnvelope

    '                pMxDoc.UpdateContents()
    '                pMxDoc.ActiveView.Refresh()
    '                pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
    '            Else
    '                'MessageBox.Show("m_legend is not a Legend", "Format Legend")

    '            End If 'if m_legend is a Legend

    '        End If
    '        pEnumElement = pPageLayoutGC.Next
    '    Loop

    'End Sub ''FormatLegend

    Public Sub SetUnusedDataFrame()
        
        Dim pPageLayout As IPageLayout
        Dim pPLGC As IGraphicsContainer
        Dim pPLDisplay As IDisplay

        Dim p2ElementProps As IElementProperties
        Dim p2Element As IElement

        Dim pElementtoHide As ienvelope2 = New Envelope


        ''pElement properties so can manipulate location
        Dim pElementEnv As ienvelope2
        pElementEnv = New Envelope

        pPageLayout = pMxDoc.PageLayout

        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()


        Dim pUpdateElement As Boolean

        pPLGC = pMxDoc.PageLayout
        pPLDisplay = pMxDoc.ActiveView.ScreenDisplay
        pPLGC.Reset()
        p2Element = pPLGC.Next

        Do While Not p2Element Is Nothing
            p2ElementProps = p2Element

            If TypeOf p2Element Is IMapFrame Then

                pUpdateElement = False

                If p2ElementProps.Name = pUnusedDataFrameName Then  'if its name is not the main frame - need to hide it

                    p2Element.QueryBounds(pPLDisplay, pElementtoHide)

                    pElementtoHide.Expand(0.01, 0.01, True)

                    Dim pHiddenPoint As IPoint = New Point

                    pHiddenPoint.PutCoords(-0.3, 0.1)


                    pElementtoHide.CenterAt(pHiddenPoint)
                    p2Element.Geometry = pElementtoHide
                    pUpdateElement = True
                End If


            End If 'if is a mapframe
           
            p2Element = pPLGC.Next
        Loop

        pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
        pMxDoc.UpdateContents()

    End Sub

    Public Function SetMapTemplate(ByVal TemplateToUse As String) As Boolean
        Dim pPageLayout As IPageLayout
        Dim pPageLayout2 As IPageLayout2

        Dim pGXPageLayout As IGxMapPageLayout
        Dim pGxFile As IGxFile

        Dim TemplateDirectory As String         'gTemplateDir is a global set in CopyTemplatesLocal
        TemplateDirectory = gTemplateDir

        Dim strTemplateFullName As String
        strTemplateFullName = TemplateDirectory & "\" & TemplateToUse


        If Not File.Exists(strTemplateFullName) Then
            MessageBox.Show("Map Template does not exist." & vbCrLf & strTemplateFullName, "Missing File", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Return False
            Exit Function
        End If

        Try

            pGxFile = New GxMap
            pGxFile.Path = strTemplateFullName

            pGXPageLayout = pGxFile

            pPageLayout = pGXPageLayout.PageLayout
            pPageLayout2 = pPageLayout


            pMxDoc = gISSIApp.Document
            pPageLayout2.ReplaceMaps(pMxDoc.Maps)

            pMxDoc.PageLayout = pPageLayout
            pMxDoc.ActiveView = pPageLayout
            pMxDoc.UpdateContents()
            pMxDoc.ActiveView.Refresh()

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "SetMapTemplate", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)
            Return False
        End Try

        Return True

    End Function


   


    Public Sub InsertLogo(ByVal pImgBoundBoxName, ByVal pImageFilePath)

        Dim pPicElemType As IPictureElement4 = New ESRI.ArcGIS.Carto.PictureElement
        Dim pImgTypeExt As String = ""

        'Try
        If pImageFilePath = "" Then
            Exit Sub
        Else                             ''create string of file extension to determine type of image to load
            pImgTypeExt = Microsoft.VisualBasic.Right(pImageFilePath, 3)      ''MsgBox pImgTypeExt, vbOKOnly, "THE IMaGE TYPE"
            Select Case pImgTypeExt
                Case "bmp"
                    pPicElemType = New ESRI.ArcGIS.Carto.BmpPictureElement
                Case "gif"
                    pPicElemType = New ESRI.ArcGIS.Carto.GifPictureElement
                Case "jpg"
                    pPicElemType = New ESRI.ArcGIS.Carto.JpgPictureElement
                Case "tif"
                    pPicElemType = New ESRI.ArcGIS.Carto.TifPictureElement
                Case "png"
                    pPicElemType = New ESRI.ArcGIS.Carto.PngPictureElement
                Case Else
                    MessageBox.Show("Cannot load image", "Incorrect FileType")

                    pImgTypeExt = "notfound"
            End Select
        End If

        If pImgTypeExt <> "notfound" Then           ''found imagetype to load - CONTINUE TO LOAD IMAGE              

            Dim pPicElement As IElement = New PictureElement
            ''Load the LogoImage
            Dim pMxApp As IMxApplication
            Dim pDisplay As IDisplay
            Dim pScreenDisplay As IScreenDisplay
            Dim pPageLayout As IPageLayout3
            Dim pPageLayoutDisplayTransformation As IDisplayTransformation
            Dim pGraphicsContainer As IGraphicsContainer
            Dim pEnumElement As IElement
            Dim pEnumElemProperties As IElementProperties3
            Dim pBoxEnv As IEnvelope2 = New Envelope
            Dim pLowerLeftPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim pBoxCenterPoint As IPoint = New ESRI.ArcGIS.Geometry.Point
            Dim pBoxX As Double
            Dim pBoxY As Double
            Dim pBoxRatio As Double

            pMxApp = gISSIApp
            gISSIMxDoc = pMxApp.Document
            pPageLayout = gISSIMxDoc.PageLayout
            gISSIMxDoc.ActiveView = pPageLayout

            pScreenDisplay = gISSIMxDoc.ActiveView.ScreenDisplay
            pDisplay = pMxApp.Display
            pPageLayoutDisplayTransformation = gISSIMxDoc.ActiveView.ScreenDisplay.DisplayTransformation
            pGraphicsContainer = pPageLayout


            pGraphicsContainer.Reset()
            pEnumElement = pGraphicsContainer.Next
            Do While Not pEnumElement Is Nothing

                pEnumElemProperties = pEnumElement
                Dim strpElementName As String
                strpElementName = pEnumElemProperties.Name
                If strpElementName = pImgBoundBoxName Then

                    Dim pElem As IElement
                    pElem = pEnumElement                    ''set the ielement to the current enumelement
                    pEnumElement.QueryBounds(pScreenDisplay, pBoxEnv)            ''get its envelope and scale it down 10%
                    pEnumElemProperties.AnchorPoint = esriAnchorPointEnum.esriCenterPoint

                    pBoxEnv.Expand(0.9, 0.9, True)

                    pLowerLeftPoint = pBoxEnv.LowerLeft
                    pBoxX = pLowerLeftPoint.X + (pBoxEnv.Width / 2)
                    pBoxY = pLowerLeftPoint.Y + (pBoxEnv.Height / 2)
                    pBoxCenterPoint.PutCoords(pBoxX, pBoxY)

                    pBoxRatio = pBoxEnv.Height / pBoxEnv.Width

                    Exit Do
                End If ' if name matches
                pEnumElement = pGraphicsContainer.Next
            Loop

            If pBoxCenterPoint.IsEmpty Then
                MessageBox.Show("Unable to locate Bounding Box for Logo Placement." & vbCrLf & "To add an image manually go to Menu: Insert -> Picture", "Exiting", MessageBoxButtons.OK, MessageBoxIcon.Hand)
                Exit Sub

            Else
                'set the picelem2 to the imagetypeTOBEloaded

                pPicElemType.Path = pImageFilePath
                pPicElemType.MaintainAspectRatio = True

                ' ''GET THE LOADED IMAGE SIZE
                Dim pPicPointwidth As Double
                Dim pPicPointHeight As Double
                pPicElemType.QueryIntrinsicSize(pPicPointwidth, pPicPointHeight)
                Dim dblPicElementRatio As Double
                dblPicElementRatio = pPicPointHeight / pPicPointwidth
                ''
                Dim pNewPicEnv As IEnvelope2 = New Envelope
                pNewPicEnv.LowerLeft = pBoxEnv.LowerLeft
                pNewPicEnv.UpperRight = pBoxEnv.UpperRight
                pNewPicEnv.CenterAt(pBoxCenterPoint)
                
                If pBoxRatio < 1.0 Then ' box is wide
                    If dblPicElementRatio < 1.0 Then ' it is wider than tall nee to scale to height
                        If dblPicElementRatio > pBoxRatio Then
                            pNewPicEnv.Height = pBoxEnv.Height
                            pNewPicEnv.Width = pBoxEnv.Height / dblPicElementRatio
                        Else    'If dblPicElementRatio < pBoxRatio Then
                            pNewPicEnv.Width = pBoxEnv.Width
                            pNewPicEnv.Height = pBoxEnv.Width * dblPicElementRatio
                        End If
                    Else ' h/w > 1 it is taller than wide

                        pNewPicEnv.Height = pBoxEnv.Height
                        pNewPicEnv.Width = pBoxEnv.Height / dblPicElementRatio

                    End If  'dblPicElementRatio < 1.0
                Else  ' box is tall
                    If dblPicElementRatio < 1.0 Then ' it is wider than tall nee to scale to height

                        pNewPicEnv.Height = pBoxEnv.Height
                        pNewPicEnv.Width = pBoxEnv.Height / dblPicElementRatio

                    Else ' h/w > 1 it is taller than wide

                        pNewPicEnv.Width = pBoxEnv.Width
                        pNewPicEnv.Height = pBoxEnv.Width * dblPicElementRatio

                    End If

                End If  ' If pBoxRatio < 1.0

                pNewPicEnv.CenterAt(pBoxCenterPoint)

                pNewPicEnv.Expand(0.9, 0.9, True)

                pPicElemType.SavePictureInDocument = True

                pPicElement = pPicElemType

                pPicElement.Geometry = pNewPicEnv

                Dim pPicElementProperties As IElementProperties3
                pPicElementProperties = pPicElement
                pPicElementProperties.AnchorPoint = esriAnchorPointEnum.esriCenterPoint

                'ADD NEW LOGO ELEMENTS TO ELEMENT COLLECTION TO ADD To THE GRAPHIC CONTAINER

                pGraphicsContainer.AddElement(pPicElement, 0)



                ''
                If pNewPicEnv.Width > pBoxEnv.Width Or pNewPicEnv.Height > pBoxEnv.Height Then
                    Dim newPicEnv As IEnvelope2 = New Envelope
                    pPicElement.QueryBounds(pDisplay, newPicEnv)
                    newPicEnv.LowerLeft = pLowerLeftPoint
                    newPicEnv.CenterAt(pBoxCenterPoint)
                    newPicEnv.Expand(0.95, 0.95, True)

                    pPicElement.Geometry = newPicEnv
                    gISSIMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphicSelection, pPicElement, newPicEnv)
                Else
                    gISSIMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphicSelection, pPicElement, pNewPicEnv)
                End If

                gISSIMxDoc.UpdateContents()

                gISSIMxDoc.ActiveView = gISSIMxDoc.PageLayout
                gISSIMxDoc.ActiveView.Refresh()



            End If ' if found centerpoint of bounding box
        End If       ' the image type was found - bmp tif jpg

        'Catch ex As Exception
        '    MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "InsertLogo", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        'End Try


    End Sub '       load logo image
   

   


    'Public Sub SyncLegendToLayout()
    '    Dim pMxDoc As IMxDocument
    '    pMxDoc = gISSIApp.Document
    '    SynchAllLegends(pMxDoc.PageLayout)
    '    pMxDoc.ActiveView.PartialRefresh(esriViewDrawPhase.esriViewGraphics, Nothing, Nothing)
    'End Sub

    'Public Sub SynchAllLegends(pUnk) ' as iunknown
    '    ' recurse to handle groupelements
    '    If TypeOf pUnk Is IGraphicsContainer Then
    '        Dim pGC As IGraphicsContainer
    '        pGC = pUnk

    '        pGC.Reset()
    '        Dim pElement As IElement
    '        pElement = pGC.Next
    '        Do Until pElement Is Nothing
    '            SynchAllLegends(pElement)
    '            pElement = pGC.Next
    '        Loop
    '    ElseIf TypeOf pUnk Is IMapSurroundFrame Then
    '        Dim pMSF As IMapSurroundFrame
    '        pMSF = pUnk
    '        If TypeOf pMSF.MapSurround Is ILegend Then
    '            SynchLegend(pMSF.MapSurround)
    '        End If
    '    ElseIf TypeOf pUnk Is IGroupElement Then
    '        Dim pGElement As IGroupElement
    '        pGElement = pUnk
    '        Dim l As Long
    '        For l = 0 To pGElement.ElementCount - 1
    '            SynchAllLegends(pGElement.Element(l))
    '        Next l
    '    End If
    'End Sub

    'Sub SynchLegend(pLegend As ILegend)
    '    '
    '    ' synch the legend item order with
    '    ' the legend's map layer order
    '    '
    '    Dim pColl As New Collection
    '    Dim l As Long
    '    For l = 0 To pLegend.ItemCount - 1
    '        pColl.Add(pLegend.Item(l))
    '    Next l
    '    pLegend.ClearItems()

    '    Dim k As Long
    '    For k = 0 To pLegend.Map.LayerCount - 1
    '        For l = 1 To pColl.Count
    '            Dim pLI As ILegendItem
    '            pLI = pColl.Item(l)
    '            If pLI.Layer Is pLegend.Map.Layer(k) Then
    '                pLegend.AddItem(pLI)
    '            End If
    '        Next l
    '    Next k
    'End Sub
    Public Sub SetUserText(Title, Author, Project, Client, pDateString, Notes)

        Dim pActiveview As IActiveView
        pActiveview = pMxDoc.ActiveView
        If Not TypeOf pActiveview Is PageLayout Then

            ArcMap.Document.ActiveView = pMxDoc.PageLayout
            ArcMap.Document.ActiveView.Refresh()
            ArcMap.Application.RefreshWindow()

        End If

        Dim formattedNotes As String = Notes
        Dim pPageLayout As IPageLayout
        Dim pPagelayoutDisplay As IDisplay
        Dim pPageLayoutGraphicsContainer As IGraphicsContainer

        Dim pElement As IElement
        Dim pElementProperties As IElementProperties3

        ''ITextElement Variables
        Dim strpElementName As String ' Name of object on layout to find
        Dim pElementText As ITextElement
        Dim pUserTextSymbol As ITextSymbol       'text symbol to apply to text element
        Dim pTextSize As Double
        Dim pUserTextSymbolSize As Double

        Dim pEnvCenter As IPoint = New Point
        Dim pCurrElemEnv As IEnvelope2 = New Envelope 'element envelope after manipulation
        Dim pElemCenter As IPoint = New Point 'element center after text edit and centering

        Dim pTitleSymbolSize As Double = 12.0
        Dim pNotesSymbolSize As Double = 7.0
        Select Case strPageSizeLetter
            Case "A"
                pTitleSymbolSize = 10.0
                pNotesSymbolSize = 7.0
            Case "B"
                pNotesSymbolSize = 8.0
                If MapTemplateMxd.Contains("Land") Then
                    pTitleSymbolSize = 12.0
                Else
                    pTitleSymbolSize = 14.0
                End If

            Case "C"
                pTitleSymbolSize = 18.0
                pNotesSymbolSize = 12.0
            Case "D"
                pTitleSymbolSize = 24.0
                pNotesSymbolSize = 12.0
            Case "E"
                pTitleSymbolSize = 36.0
                pNotesSymbolSize = 14.0
            Case Else
                pTitleSymbolSize = 7.0
                pNotesSymbolSize = 7.0
        End Select

        pMxDoc = gISSIApp.Document
        pPageLayout = pMxDoc.PageLayout
        pPagelayoutDisplay = pMxDoc.ActiveView.ScreenDisplay
        pPageLayoutGraphicsContainer = pPageLayout

        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''GET THE TITLE BOUNDING BOX PROPERTIES
        Dim pTitleBBox As IElement
        pTitleBBox = New RectangleElement
        pTitleBBox = FindBoundBox("TitleBoundingBox")
        Dim pTitleBBoxCenter As IPoint = New Point
        Dim pTitleBBoxENV As IEnvelope2 = New Envelope
        If Not pTitleBBox Is Nothing Then

            ''ENVELOPE
            pTitleBBox.QueryBounds(pPagelayoutDisplay, pTitleBBoxENV)
            pTitleBBoxENV.Expand(0.95, 0.95, True)

            ''CENTERPOINT   
            Dim pTitleCenterX As Double = 0.0
            Dim pTitleCenterY As Double = 0.0


            pTitleCenterX = pTitleBBoxENV.LowerLeft.X + (pTitleBBoxENV.Width / 2)
            pTitleCenterY = pTitleBBoxENV.LowerLeft.Y + (pTitleBBoxENV.Height / 2)
            pTitleBBoxCenter.PutCoords(pTitleCenterX, pTitleCenterY)

        End If ' if found titlebox

        Dim pMapDocoBox As IElement
        pMapDocoBox = New RectangleElement
        pMapDocoBox = FindBoundBox("MapDocoBox")

        Dim pMaxMapDocoWIdth As Double = 0.0
        If Not pMapDocoBox Is Nothing Then
            pMaxMapDocoWIdth = pMapDocoBox.Geometry.Envelope.Width
        End If

        ''PROJECTNOTES BOUNDINGBOX PROPERTIES
        Dim pProjNotesBBox As IElement = New RectangleElement
        pProjNotesBBox = FindBoundBox("ProjNotesBox")

        Dim pProjNoBoxCenter As IPoint = New Point
        Dim pProjNotesBBoxENV As IEnvelope2 = New Envelope


        Dim pProjNotesUpperLeft As IPoint = New Point
        If Not pProjNotesBBox Is Nothing Then
            Dim pProjNotesBBoxProperties As IElementProperties3
            pProjNotesBBoxProperties = pProjNotesBBox
            pProjNotesBBoxProperties.AnchorPoint = esriAnchorPointEnum.esriTopLeftCorner

            pProjNotesBBox.QueryBounds(pPagelayoutDisplay, pProjNotesBBoxENV)
            pProjNotesUpperLeft = pProjNotesBBoxENV.UpperLeft

            pProjNoBoxCenter.PutCoords(pProjNotesBBoxENV.LowerLeft.X + (pProjNotesBBoxENV.Width / 2), pProjNotesBBoxENV.LowerLeft.Y + (pProjNotesBBoxENV.Height / 2))

        End If ' if found project notes box


        ''''''''''''''''''''''''''''''''''''''''''''''''''''''''

        pPageLayoutGraphicsContainer.Reset()
        pElement = pPageLayoutGraphicsContainer.Next
        ''********Loop through each element in the layout that set user text
        Do While Not pElement Is Nothing
            If TypeOf pElement Is ITextElement Then
                pElementProperties = pElement

                If pElementProperties.Name <> "" Then
                    strpElementName = pElementProperties.Name  'the name of the element

                    ''CREATE A NEW TEXTELEMENT USING THE CURRENT ELEMENT SYMBOL PROPERTIES
                    pElementText = New TextElement
                    pElementText = pElement

                    ''GET THE TEXT ELEMENT SYMBOL
                    pUserTextSymbol = pElementText.Symbol
                    pTextSize = pUserTextSymbol.Size
                    
                    Select Case strpElementName

                        Case "theTitle"     'Set pElementText = pElement

                            If Len(Title) <> 0 Then
                                pUserTextSymbol.HorizontalAlignment = 1           ' esriTHALeft
                                pUserTextSymbol.Size = pTitleSymbolSize
                                pElementText.Symbol = pUserTextSymbol

                                pElementText.Text = Title

                                ''set the current pelement to the new textelement and center it in the bounding box
                                pElement = pElementText
                                pElement.Geometry = pTitleBBoxCenter

                                pPageLayoutGraphicsContainer.UpdateElement(pElement)
                                Dim pTitleTextEnv As IEnvelope2 = New Envelope
                                ''GET TEXTELEMENT ENVELOPE AFTER SETTING TEXT
                                pTitleTextEnv = New Envelope
                                pElement.QueryBounds(pPagelayoutDisplay, pTitleTextEnv)
                                pTitleBBoxCenter.PutCoords(pTitleBBoxCenter.X, pTitleBBoxCenter.Y - (pTitleTextEnv.Height / 2))
                                pTitleTextEnv.CenterAt(pTitleBBoxCenter)


                                pElement.Geometry = pTitleBBoxCenter
                                pPageLayoutGraphicsContainer.UpdateElement(pElement)


                            End If 'TITLE LENGTH NOT 0

                        Case "theClient"

                            pUserTextSymbol.HorizontalAlignment = 0           ' esriTHALeft
                            pElementText.Text = "Issued For: " & Client
                            ''CHECK width of element
                            pElement.QueryBounds(pPagelayoutDisplay, pCurrElemEnv)

                            If pMaxMapDocoWIdth <> 0.0 Then
                                If pCurrElemEnv.Width > pMaxMapDocoWIdth Then
                                    pTextSize = pUserTextSymbol.Size - 1
                                    pUserTextSymbol.Size = pTextSize
                                End If
                            End If

                            pElementText.Symbol = pUserTextSymbol

                        Case "theMapAuthor"
                            pUserTextSymbol.HorizontalAlignment = 0      ' esriTHALeft
                            pElementText.Text = "Prepared By: " & Author

                            pElement.QueryBounds(pPagelayoutDisplay, pCurrElemEnv)
                            If pMaxMapDocoWIdth <> 0.0 Then
                                If pCurrElemEnv.Width > pMaxMapDocoWIdth Then
                                    pTextSize = pUserTextSymbol.Size - 1
                                    pUserTextSymbol.Size = pTextSize
                                End If
                            End If

                            pElementText.Symbol = pUserTextSymbol

                        Case "theProjNo"
                            pUserTextSymbol.HorizontalAlignment = 0
                            pElementText.Text = "Project: " & Project

                            pElement.QueryBounds(pPagelayoutDisplay, pCurrElemEnv)
                            If pMaxMapDocoWIdth <> 0.0 Then
                                If pCurrElemEnv.Width > pMaxMapDocoWIdth Then

                                    pTextSize = pUserTextSymbol.Size - 1
                                    pUserTextSymbol.Size = pTextSize
                                End If
                            End If

                            pElementText.Symbol = pUserTextSymbol


                        Case "theDate"
                            pUserTextSymbol.HorizontalAlignment = 0
                            pElementText.Text = "Date: " & pDateString
                            pElementText.Symbol = pUserTextSymbol

                        Case "ProjectNotes"

                            pElement = pElementText

                            pElementProperties.AnchorPoint = esriAnchorPointEnum.esriTopLeftCorner
                            pElement.Geometry = pProjNotesBBoxENV.UpperLeft
                            pPageLayoutGraphicsContainer.UpdateElement(pElement)

                            If Len(Notes) <> 0 Then
                                pUserTextSymbol.HorizontalAlignment = 0

                                Select Case strPageSizeLetter
                                    Case "A"
                                        pUserTextSymbol.Size = 7.0
                                        If MapTemplateMxd.Contains("Land_") Then
                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 49) ' font size * 7
                                            Else
                                                formattedNotes = FormatByLength(Notes, 56) ' font size * 7
                                            End If
                                        Else ' Portrait
                                            
                                            If MapTemplateMxd.Contains("Vert") Then
                                                If MapTemplateMxd.Contains("noIndex") Then
                                                    formattedNotes = FormatByLength(Notes, 45) ' font size * 7
                                                Else
                                                    formattedNotes = FormatByLength(Notes, 49) ' font size * 7
                                                End If

                                            Else
                                                formattedNotes = FormatByLength(Notes, 47)
                                            End If
                                           
                                        End If

                                    Case "B"
                                        pUserTextSymbol.Size = 8.0

                                        If MapTemplateMxd.Contains("Land_") Then
                                            If MapTemplateMxd.Contains("noIndex") Then
                                                If MapTemplateMxd.Contains("Vert_") Then
                                                    formattedNotes = FormatByLength(Notes, 52) ' font size * 7
                                                Else
                                                    formattedNotes = FormatByLength(Notes, 49) ' font size * 7
                                                End If

                                            Else
                                                formattedNotes = FormatByLength(Notes, 56) ' font size * 7
                                            End If

                                        Else ' Portrait

                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 42) ' font size * 7
                                            Else
                                                formattedNotes = FormatByLength(Notes, 49) ' font size * 7
                                            End If

                                            End If

                                    Case "C"
                                        pUserTextSymbol.Size = 14.0
                                        If MapTemplateMxd.Contains("Land_Vert") Then
                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 45)
                                            Else
                                                formattedNotes = FormatByLength(Notes, 64)
                                            End If

                                        ElseIf MapTemplateMxd.Contains("Port_Vert_") Then
                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 42)
                                            Else
                                                formattedNotes = FormatByLength(Notes, 50)
                                            End If

                                        Else
                                            'Portrait
                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 50)  'font size * 5
                                            Else
                                                formattedNotes = FormatByLength(Notes, 64)  'font size * 6 ' landscape horizontal(2)
                                            End If

                                        End If

                                    Case "D"
                                        pUserTextSymbol.Size = 14.0

                                        If MapTemplateMxd.Contains("Port_Horiz_") Then
                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 50)
                                            Else
                                                formattedNotes = FormatByLength(Notes, 72)
                                            End If
                                        Else
                                            formattedNotes = FormatByLength(Notes, 72)
                                        End If


                                    Case "E"
                                            pUserTextSymbol.Size = 14.0

                                            If MapTemplateMxd.Contains("noIndex") Then
                                                formattedNotes = FormatByLength(Notes, 84)  'font size * 6
                                            Else
                                                If MapTemplateMxd.Contains("Port_Vert") Then
                                                    formattedNotes = FormatByLength(Notes, 112)  ' font size * 8
                                                Else
                                                    formattedNotes = FormatByLength(Notes, 140) ' font size * 10
                                                End If

                                            End If



                                    Case Else
                                            pUserTextSymbol.Size = 7.0
                                            formattedNotes = FormatByLength(Notes, 42)
                                End Select
                                pElementText.Text = formattedNotes

                                pElementText.Symbol = pUserTextSymbol

                                pElement = pElementText


                                pElement.Geometry = pProjNotesBBox.Geometry.Envelope.LowerLeft

                                pElementProperties.AnchorPoint = esriAnchorPointEnum.esriBottomLeftCorner


                                pPageLayoutGraphicsContainer.UpdateElement(pElement)

                                Dim pShiftedLowerLeft As IPoint = New Point
                                ' ''GET TEXTELEMENT ENVELOPE AFTER SETTING TEXT and center at ProjectNotesBox Center Point
                                Dim pNotesTextEnv As IEnvelope2 = New Envelope
                                pElement.QueryBounds(pPagelayoutDisplay, pNotesTextEnv)
                               
                                pShiftedLowerLeft.PutCoords(pProjNotesBBox.Geometry.Envelope.LowerLeft.X, pProjNotesBBox.Geometry.Envelope.LowerLeft.Y + (pProjNotesBBoxENV.Height - pNotesTextEnv.Height))

                                pNotesTextEnv.CenterAt(pShiftedLowerLeft)
                                pElement.Geometry = pShiftedLowerLeft


                                pPageLayoutGraphicsContainer.UpdateElement(pElement)


                                If pNotesTextEnv.Width > pProjNotesBBoxENV.Width Or pNotesTextEnv.Height > pProjNotesBBoxENV.Height Then
                                    pUserTextSymbolSize = pUserTextSymbol.Size - 1
                                    pUserTextSymbol.Size = pUserTextSymbolSize
                                    pElementText.Symbol = pUserTextSymbol
                                Else

                                End If
                            End If 'Notes text length is not 0


                        Case Else
                            ' NAME DOES NOT MATCH ANYTHING SO NOTHING SHOULD BE DONE


                    End Select


                End If 'if graphic has a name
            End If 'if element is a text element

            pPageLayoutGraphicsContainer.UpdateElement(pElement)
            pMxDoc.UpdateContents()
            pElement = pPageLayoutGraphicsContainer.Next
        Loop  ' do while pelement is not nothing

    

        Dim pGraphicContainer As IGraphicsContainerSelect
        pGraphicContainer = pPageLayoutGraphicsContainer
        pGraphicContainer.UnselectAllElements()

        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()


    End Sub




    Public Function FindBoundBox(strpBoxName) As IElement ' the name of the text element being edited to find its frame
        Dim pPageLayout As IPageLayout
        Dim pPagelayoutDisplay As IDisplay

        Dim pFBBelement As IElement
        Dim pBBElementProperties As IElementProperties3

        Dim pBoundBoxElement As IElement = Nothing
        Dim pBoundBoxRect As IRectangleElement
        Dim pBoundBoxProperties As IElementProperties3
      
        pPageLayout = pMxDoc.PageLayout
        pPagelayoutDisplay = pMxDoc.ActiveView.ScreenDisplay

        Dim playoutGC As IGraphicsContainer
        playoutGC = pPageLayout
        playoutGC.Reset()
        pFBBelement = playoutGC.Next
        ''********Loop through each element in the layout
        Do While Not pFBBelement Is Nothing
            pBBElementProperties = pFBBelement
            If pBBElementProperties.Name <> "" Then
                If TypeOf pFBBelement Is RectangleElement Then
                    ''set variablename for current rectangleelement with a name
                    'strpElementBoxName = pBBElementProperties.Name
                    If (strpBoxName = pBBElementProperties.Name) Then
                        pBoundBoxRect = New RectangleElement
                        pBoundBoxRect = pFBBelement
                        pBoundBoxElement = pBoundBoxRect
                        pBoundBoxProperties = pFBBelement
                        pBoundBoxProperties.Name = strpBoxName
                    End If
                End If 'if pelement has a name
            End If
            pFBBelement = playoutGC.Next
        Loop
       
        Return pBoundBoxElement

    End Function




    Public Sub ReorderMaps()
        ''Based on the MapInput Form set the Main data frame as the focusmap      
        Dim pmaps1 As IMaps
        Dim pMaps As IMaps2
        Dim pFmap As IMap = Nothing

        pMxDoc = gISSIApp.Document

        Try
            pFmap = gISSIMxDoc.FocusMap
            pUnusedDataFrameName = ""
            pmaps1 = pMxDoc.Maps
            pMaps = pmaps1

            Dim i As Integer
            For i = 0 To pmaps1.Count - 1
                If pmaps1.Item(i).Name = gMainMapName Then

                    pFmap = pmaps1.Item(i)
                Else
                    If gIndexMapName = "None" Then
                        pUnusedDataFrameName = pmaps1.Item(i).Name

                    End If
                End If
            Next i

            If pMaps.Count > 1 Then

                If gIndexMapName <> "None" Then 'they want an index map and it needs to be the top dataframe
                    If pMaps.Item(0).Name = pFmap.Name Then     ''if the top dataframe is the main dataframe
                        pMaps.MoveItem(pMaps.Item(1), 0)   ''move the 2nd dataframe to the top
                    End If
                Else    ''there is more than one data frame but they dont want an index
                    If pMaps.Item(0).Name <> gMainMapName Then
                        pMaps.MoveItem(pFmap, 0)
                    End If
                End If
            Else
                'MessageBox.Show("the top dataframe is " & pMaps.Item(0).Name, "3")
            End If

            If pMaps.Count > 1 And pUnusedDataFrameName <> "" Then 'move the gMainMapNAme to the bottom
                If pFmap.Name = pUnusedDataFrameName Then
                    pMaps.MoveItem(pFmap, pMaps.Count - 1)
                End If
            End If

            pMxDoc.UpdateContents()

            pMxDoc.ActiveView.Refresh()

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "ReorderMaps", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try



    End Sub
    Public Sub SetMainAsFocus()

        Dim pMaps As IMaps2
        Dim pfmap As IMap

        Try
            pMaps = gISSIMxDoc.Maps

            Dim i As Integer
            For i = 0 To pMaps.Count - 1

                If pMaps.Item(i).Name = gMainMapName Then
                    gISSIMxDoc.ActiveView = pMaps.Item(i)
                    pfmap = pMaps.Item(i)
                End If
            Next i


        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "SetMainasFocus", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try

    End Sub

    Public Sub AddProjectionInfo()
        Dim pMap As IMap
        Dim pActiveView As IActiveView
        Dim pMxApp As IMxApplication
        Dim pDisplay As IDisplay

        Dim pPageLayout As IPageLayout
        Dim pGraphContainer As IGraphicsContainer

        pMxApp = gISSIApp
        pMxDoc = gISSIApp.Document
        pMap = pMxDoc.FocusMap

        pActiveView = pMap
        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()

        ''Get the Projection Info for the Main Map
        ''''''
        ''THIS WILL NOT WORK IF THERE IS NO DATA IN THE DATA FRAME
        ''NEED TO CHECK FOR DATA AND BAIL IF NONE


        Dim theProjectionText As String
        Dim theSpatialReferenceName As String
        Dim theSpatialReferenceText As String
        theSpatialReferenceName = pMap.SpatialReference.Name

        theSpatialReferenceText = Replace(theSpatialReferenceName, "_", " ", 1, -1, vbTextCompare)
        theProjectionText = "Projection: " & theSpatialReferenceText

        ''Get the Lower corner coordnates for the Main Map in the layout
        pDisplay = pMxApp.Display
        pPageLayout = pMxDoc.PageLayout
        pGraphContainer = pPageLayout

        ''Get Envelope of  gMainMainName Map on Layout
        Dim pMapFrame As IMapFrame     '' the main map frame on the layout
        Dim pMapFrameEnv As ienvelope2 = New Envelope

        Dim pGElement As IElement
        pGraphContainer.Reset()
        pGElement = pGraphContainer.Next
        While Not pGElement Is Nothing
            If TypeOf pGElement Is IMapFrame Then
                pMapFrame = pGElement
                If pMapFrame.Map.Name = gMainMapName Then
                    pMapFrameEnv = pGElement.Geometry.Envelope
                End If
            End If
            pGElement = pGraphContainer.Next
        End While

        ''Need to get the lower coordinate and inset the text
        ''Set coordinates just inside map extent
        Dim pShrinkEnv As ienvelope2
        pShrinkEnv = New Envelope
        pShrinkEnv.Union(pMapFrameEnv)
        pShrinkEnv.Expand(0.96, 0.96, True)

        ''ADD TEXT Label For the LAT Long Coordinate
        Dim pTextColor As IRgbColor
        pTextColor = New RgbColor
        pTextColor.Blue = 0
        pTextColor.Green = 0
        pTextColor.Red = 0
        Dim pSimpleTextSymbol As ISimpleTextSymbol
        pSimpleTextSymbol = New TextSymbol
        With pSimpleTextSymbol
            .Size = 7
            .Color = pTextColor
            .XOffset = 0
            .YOffset = 0
            .HorizontalAlignment = esriTextHorizontalAlignment.esriTHALeft

        End With

        'Create a graphic label element
        Dim pTextElement As ITextElement
        pTextElement = New TextElement



        pTextElement.Text = theProjectionText       'pTextElement.Text = "This is some example text"
        pTextElement.Symbol = pSimpleTextSymbol     'pTextElement.Symbol = pTextSymbol

        ''Create MaskHalo for ProjectionText
        Dim pRGBClr As IRgbColor
        pRGBClr = New RgbColor
        ''set color to white
        pRGBClr.Red = 255
        pRGBClr.Blue = 255
        pRGBClr.Green = 255

        'Create a Fill Symbol for the Mask
        Dim pSmpFill As ISimpleFillSymbol
        pSmpFill = New SimpleFillSymbol
        pSmpFill.Color = pRGBClr
        pSmpFill.Style = esriSimpleFillStyle.esriSFSSolid

        ''Create Mask For Text Symbol
        Dim pMrkMask As IMask
        pMrkMask = pTextElement.Symbol

        pMrkMask.MaskSymbol = pSmpFill
        pMrkMask.MaskStyle = esriMaskStyle.esriMSHalo
        pTextElement.Symbol = pMrkMask

        Dim pPoint As IPoint
        pPoint = New Point
        pPoint.X = pShrinkEnv.LowerLeft.X
        pPoint.Y = pShrinkEnv.LowerLeft.Y

        Dim pTElement As IElement
        pTElement = pTextElement

        pTElement.Geometry = pPoint

        '** add the element to the graphics layer
        pActiveView = pMxDoc.ActiveView
        pGraphContainer = pActiveView.GraphicsContainer
        pGraphContainer.AddElement(pTElement, 0)
        pMxDoc.ActiveView.Refresh()

    End Sub

    Public Sub SetPrintPageLayout(PageSize)
        Dim pPageLayout As IPageLayout
        Dim pPage As IPage

        pMxDoc = gISSIApp.Document
        pPageLayout = pMxDoc.PageLayout
        pPage = pPageLayout.Page

        Dim MapSizeLetter As String
        MapSizeLetter = Left(PageSize, 1)
        Select Case MapSizeLetter
            Case "A"
                pPage.FormID = esriPageFormID.esriPageFormLetter
            Case "B"
                pPage.FormID = esriPageFormID.esriPageFormTabloid
            Case "C"
                pPage.FormID = esriPageFormID.esriPageFormC
            Case "D"
                pPage.FormID = esriPageFormID.esriPageFormD
            Case "E"
                pPage.FormID = esriPageFormID.esriPageFormE
        End Select

        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
        pPageLayout.ZoomToWhole()

    End Sub
    Public Sub ScalePageSize(PageSize)
        Dim pPageLayout As IPageLayout
        Dim pPage As IPage

        pMxDoc = gISSIApp.Document
        pPageLayout = pMxDoc.PageLayout
        pPage = pPageLayout.Page

        Dim MapSizeLetter As String
        MapSizeLetter = Left(PageSize, 1)
        Select Case MapSizeLetter
            Case "A"
                pPage.FormID = esriPageFormID.esriPageFormLetter
            Case "B"
                pPage.FormID = esriPageFormID.esriPageFormTabloid
            Case "C"
                pPage.FormID = esriPageFormID.esriPageFormC
            Case "D"
                pPage.FormID = esriPageFormID.esriPageFormD
            Case "E"
                pPage.FormID = esriPageFormID.esriPageFormE
        End Select

        pMxDoc.UpdateContents()
        pMxDoc.ActiveView.Refresh()
        pPageLayout.ZoomToWhole()

    End Sub
    Public Function FormatByLength(Expression As String, length As Long) As String
      
        Dim BufferSpace() As String
        Dim Buffer As String = ""

        Dim j As Long
        Dim count As Long


        BufferSpace = Split(Expression, " ")
        For j = 0 To UBound(BufferSpace)
            count = count + Len(BufferSpace(j)) + 1
            If (count <= length) Then
                Buffer = Buffer & BufferSpace(j) & " "
            Else
                count = 0
                Buffer = Buffer & vbCrLf & BufferSpace(j) & " "
                count = Len(BufferSpace(j)) + 1
            End If
        Next j
     
        ''NEED TO REMOVE TRAILING VBCRLFS
        If Right(Buffer, 1) = vbCrLf Or Right(Buffer, 1) = vbCr Then
            Buffer = Left(Buffer, Len(Buffer) - 1)
        End If
       
        Return Buffer


    End Function

   

    Public Function txtParserToLines(Expression As String, Lines As Long) As String
        ''this is run AFTER FormatByLength
        Dim strLines As String = ""
        If Len(Expression) <> 0 Then

            Dim LineItemCrLf() As String     'string array defined by carriage returns
            Dim LinestoKeep As String = ""        'the outstring
            Dim SplitbyLength As Boolean

            LineItemCrLf = Split(Expression, vbCrLf)

            If UBound(LineItemCrLf) > Lines - 1 Then

                If Len(LineItemCrLf(0)) >= 30 Then
                    SplitbyLength = True
                End If
                If Trim(LineItemCrLf(0)) = "" Then
                    LinestoKeep = LineItemCrLf(1) & vbCrLf & LineItemCrLf(2)
                Else
                    LinestoKeep = LineItemCrLf(0) & vbCrLf & LineItemCrLf(1) & vbCrLf & LineItemCrLf(2)
                End If
            Else
                SplitbyLength = True

            End If

            If SplitbyLength = True Then
                If Trim(Left(Expression, 25)) = "" Then
                    Expression = Right(Expression, (Len(Expression) - 25))
                End If

                LinestoKeep = Left(Expression, 25) & vbCrLf & Mid(Expression, 26, 25) & vbCrLf & Mid(Expression, 51, 25) & vbCrLf & Mid(Expression, 76, 25)

            End If


            strLines = LinestoKeep
        Else
            strLines = Expression

        End If 'if length is not ZERO

        Return strLines


    End Function


    Public Function FindFileDir(invalidFilePath) As String 'return the Dir to Browse For File
        Dim pPathtoFind As String
        pPathtoFind = invalidFilePath

        Dim pDirectory As String
        Dim pBrowseDir As String

        Dim pFNames As IFileNames
        pFNames = New FileNames
        pFNames.Add(pPathtoFind)
        pFNames.Reset()
        pFNames.Next()
        pFNames.Reset()
        pDirectory = pFNames.Next

        If Dir(pDirectory) = "" Then
            ''BackUp a Directory
            Dim pDirSlashIndex As Integer

            pDirSlashIndex = InStrRev(pDirectory, "\", Len(pDirectory) - 3, vbBinaryCompare)
            pBrowseDir = Left(pDirectory, Len(pDirectory) - (Len(pDirectory) - pDirSlashIndex))

        Else    ' the  directory is valid
            pBrowseDir = pDirectory

        End If
        FindFileDir = pBrowseDir

    End Function
    Public Function UserSelectDirectory(pBrowseDir As String, pDirNeeded As String) As String
        Dim pFileDirName As String = ""
        Dim pInputFileDir As IFileName = Nothing
        ''Set up FILEDIALOG WITH LIST OF DIRECTORIES
        Dim pFileDirListbox As IListDialog
        pFileDirListbox = New ListDialog

        Dim pDirtoSearch As String = ""
        Dim pFileName As String = ""
        Dim pSelectedDir As String = ""

        pDirtoSearch = pBrowseDir

        ''Create List of directories for selection
        pFileName = Dir(pDirtoSearch, vbDirectory)               ' Retrieve the first entry.
        Do While pFileName <> ""   ' Start the loop.
            If pFileName <> "." And pFileName <> ".." Then      ' Use bitwise comparison to make sure is a directory.
                If (GetAttr(pDirtoSearch & pFileName) And vbDirectory) = vbDirectory Then

                    pFileDirListbox.AddString(pFileName)

                End If   ' it represents a directory.
            End If
            pFileName = Dir()   ' Get next entry.
        Loop

        ''PROMPT USER TO SELECT A DIRECTORY
        Dim pFileDirCount As Integer
        pFileDirCount = 0

        If pFileDirListbox.DoModal("Select " & pDirNeeded, 0, gISSIApp.hWnd) <> False Then 'theyselected a directory
            Dim pSelectedFNIndex As Integer
            pSelectedFNIndex = pFileDirListbox.Choice     'the indexnumber of string selected

            pFileName = Dir(pDirtoSearch, vbDirectory)   ' Retrieve the first entry that is a directory
            Do While pFileName <> ""   ' Start the loop.
                If pFileName <> "." And pFileName <> ".." Then
                    If (GetAttr(pDirtoSearch & pFileName) And vbDirectory) = vbDirectory Then
                        If pFileDirCount = pSelectedFNIndex Then 'if the count = pSelectedFNIndex
                            'then get the assoc filename
                            pFileDirName = pFileName
                        End If
                        pFileDirCount = pFileDirCount + 1
                    End If      ' it is a directory.
                End If 'if not = ..
                pFileName = Dir()   ' Get next entry.
            Loop

            If (GetAttr(pDirtoSearch & pFileDirName) And vbDirectory) = vbDirectory Then

                pInputFileDir = New FileName
                pInputFileDir.Path = pDirtoSearch & pFileDirName
            End If
        End If 'they reset the template directory
        If Not pInputFileDir Is Nothing Then
            pSelectedDir = pInputFileDir.Path
        End If
        ''Set DIRECTORY TO USER SELECTED

        UserSelectDirectory = pSelectedDir
    End Function

    Public Sub SetESRIAddInPath() 'As String
        Dim codeBase As String = System.Reflection.Assembly.GetExecutingAssembly.CodeBase
        Dim uriBuilder As UriBuilder = New UriBuilder(codeBase)
        Dim path As String = Uri.UnescapeDataString(uriBuilder.Path)

        esriAddinPath = System.IO.Path.GetDirectoryName(path)
        gTemplateDirPath = esriAddinPath + "\ISSITemplates"
        'Return gTemplateDirPath
    End Sub
    Public Sub CopyTemplatesLocal()
        Dim pMouseCursor As IMouseCursor = New MouseCursor
        pMouseCursor.SetCursor(2)

        ''
        'Dim codeBase As String = System.Reflection.Assembly.GetExecutingAssembly.CodeBase
        'Dim uriBuilder As UriBuilder = New UriBuilder(codeBase)
        'Dim path As String = Uri.UnescapeDataString(uriBuilder.Path)
        'esriAddinPath = System.IO.Path.GetDirectoryName(path)
        ' ''MessageBox.Show("esriAddinPath = " + esriAddinPath)
        'gTemplateDirPath = esriAddinPath + "\ISSITemplates"
        ''

        SetESRIAddInPath()


        If Not File.Exists(gTemplateDirPath) Then
            'Create template folder
            Directory.CreateDirectory(gTemplateDirPath)

            Try
                ''ADd CHeck if already exists so do not overwrite everytime??
                Dim arrPageLetters() As String
                arrPageLetters = {"A", "B", "C", "D", "E"}
                Dim PL As String
                For intPL = 0 To arrPageLetters.Length - 1
                    PL = arrPageLetters(intPL)

                    SaveEmbeddedResourcesFile("Land_Vert_" & PL & ".mxd", gTemplateDirPath & "\Land_Vert_" & PL & ".mxd")
                    SaveEmbeddedResourcesFile("Land_Vert_" & PL & "_noIndex.mxd", gTemplateDirPath & "\Land_Vert_" & PL & "_noIndex.mxd")

                    SaveEmbeddedResourcesFile("Land_Horiz_" & PL & ".mxd", gTemplateDirPath & "\Land_Horiz_" & PL & ".mxd")
                    SaveEmbeddedResourcesFile("Land_Horiz_" & PL & "_noIndex.mxd", gTemplateDirPath & "\Land_Horiz_" & PL & "_noIndex.mxd")

                    SaveEmbeddedResourcesFile("Port_Horiz_" & PL & ".mxd", gTemplateDirPath & "\Port_Horiz_" & PL & ".mxd")
                    SaveEmbeddedResourcesFile("Port_Horiz_" & PL & "_noIndex.mxd", gTemplateDirPath & "\Port_Horiz_" & PL & "_noIndex.mxd")

                    SaveEmbeddedResourcesFile("Port_Vert_" & PL & ".mxd", gTemplateDirPath & "\Port_Vert_" & PL & ".mxd")
                    SaveEmbeddedResourcesFile("Port_Vert_" & PL & "_noIndex.mxd", gTemplateDirPath & "\Port_Vert_" & PL & "_noIndex.mxd")

                    SaveEmbeddedResourcesFile("ISSIMap_" & PL & ".png", gTemplateDirPath & "\ISSIMap_" & PL & ".png")
                Next intPL

                'Copy ISSI Logo Local for use later also
                SaveEmbeddedResourcesFile("ISSI_Logo.png", gTemplateDirPath & "\ISSI_logo.png")

                '' Default ISSI Logo for Map Templates
                SaveEmbeddedResourcesFile("ISSIMap.png", gTemplateDirPath & "\ISSIMap.png")
                gTemplateDir = gTemplateDirPath
                strISSILogo = "ISSIMap.png"
                strISSIFullName = gTemplateDir & "\" & strISSILogo

                ' Copy Help File Local Memory Location

                SaveEmbeddedResourcesFile("ISSIMapV1p0.pdf", gTemplateDirPath & "\ISSIMapV1p0.pdf")

            Catch ex As Exception
                'gTemplateDir = ""
                MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "CopyTemplatesLocal", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

            End Try
        Else
            gTemplateDir = gTemplateDirPath

        End If ' if template directory does not exist

        pMouseCursor.SetCursor(1)
    End Sub


    Public Sub SaveEmbeddedResourcesFile(ByVal fileToGetFromAssembly As String, ByVal LocationTosavefileTo As String)
        Dim resFilestream As Stream
        Dim Exeassembly As System.Reflection.Assembly = Reflection.Assembly.GetExecutingAssembly
        Dim Mynamespace As String = Exeassembly.GetName().Name.ToString()

        Try
            'If Not File.Exists(LocationTosavefileTo) Then
            Dim xCount As Integer
            resFilestream = Exeassembly.GetManifestResourceStream(Mynamespace + "." + fileToGetFromAssembly)
            Dim writeResources As New System.IO.FileStream(LocationTosavefileTo, FileMode.OpenOrCreate)
            For xCount = 1 To resFilestream.Length
                writeResources.WriteByte(resFilestream.ReadByte)
            Next
            writeResources.Close()
            resFilestream.Close()
            'Else
            '    MessageBox.Show("File already exists" & LocationTosavefileTo, "SaveEmbeddedResourcesFile", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)
            ' NEED to BE able to tell if file is different otherwise updated files will not replace old files because the file is already there
            'End If
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "SaveEmbeddedResourcesFile", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try
    End Sub

    Public Function GetTemplateandLogo() As System.Array


        Dim llElement As IElement
        Dim llElementProperties As IElementProperties3
        Dim strtElementName As String = ""

        Dim pPageLayoutGC As IGraphicsContainer
        Dim tPageLayout As IPageLayout3
        tPageLayout = gISSIMxDoc.PageLayout
        pPageLayoutGC = tPageLayout

        Dim pMXApp As IMxApplication = gISSIApp
        
        ''Get Page Size and orientations
        Dim dblCurrentWidth As Double
        Dim dblCurrentHeight As Double
        tPageLayout.Page.QuerySize(dblCurrentWidth, dblCurrentHeight)

        Dim strPageSize As String = dblCurrentWidth.ToString & "," & dblCurrentHeight.ToString

        Dim dblPageRatio As Double = dblCurrentHeight / dblCurrentWidth


        Dim pPageOrientation As String = "Land"
        If dblPageRatio < 1.0 Then
            pPageOrientation = "Land"
        Else
            pPageOrientation = "Port"
        End If
        Dim pLegendOrientation As String = "Vert" ' vertical(side legend) or horizontal(bottom legend)
        Dim pMapLegendOrientation As String = pPageOrientation & pLegendOrientation     '"rbLandVert" ' default

        Dim pMapLogoPath As String = ""
        Dim p_LogoName As String = ""

        Try
            'Find Text Elements
            Dim pPageLayoutGCSelect As IGraphicsContainerSelect
            pPageLayoutGCSelect = pPageLayoutGC

            pPageLayoutGC.Reset()
            llElement = pPageLayoutGC.Next

            Do While Not llElement Is Nothing
                ''PUT SELECT CASE OF TEXT ELEMENTS TO SET TEXT IN HERE
                llElementProperties = llElement

                If llElementProperties.Name <> "" Then
                    strtElementName = llElementProperties.Name
                End If
                If llElementProperties.Type = "Data Frame" Then
                    
                    If llElementProperties.Name = "Layers" Then
                        Dim pMapFrame As IMapFrame = Nothing
                        pMapFrame = llElementProperties
                        Dim pMapFrameEnv As ienvelope2 = New Envelope
                        pMapFrameEnv = llElement.Geometry.Envelope

                        Dim dblMapRatio As Double = (pMapFrameEnv.Height / pMapFrameEnv.Width)
                        
                        If pPageOrientation = "Land" Then
                            If dblMapRatio >= 0.6 Then
                                'vertical
                                pLegendOrientation = "Vert"
                            Else
                                'horizontal
                                pLegendOrientation = "Horz"
                            End If
                        Else
                            If dblMapRatio >= 1.3 Then
                                'vertical
                                pLegendOrientation = "Vert"
                            Else
                                'horizontal
                                pLegendOrientation = "Horz"
                            End If

                        End If

                    End If ' if data frame is layers


                ElseIf TypeOf llElementProperties Is IPictureElement Or llElementProperties.Type = "Picture" Then

                    Dim pmapLogo As IPictureElement5
                    pmapLogo = llElement

                    If pmapLogo.Path <> "" And pmapLogo.Path <> strISSIFullName Then
                        pMapLogoPath = pmapLogo.Path
                      
                        If Not pMapLogoPath.Contains("ISSIMap_") Then
                            p_LogoName = pMapLogoPath
                        End If ' if it is not the ISSI logo
                    End If ' ifmaplogopath is blank

                ElseIf llElementProperties.Name = "LegendBox" Then

                    Dim pLegendEnv As ienvelope2 = New Envelope
                    pLegendEnv = llElement.Geometry.Envelope

                    Dim dblLegendRatio As Double = (pLegendEnv.Height / pLegendEnv.Width)
                    
                    If dblLegendRatio >= 0.9 Then
                        'vertical
                        pLegendOrientation = "Vert"
                    Else
                        'horizontal
                        pLegendOrientation = "Horz"
                    End If


                End If ' if dataframe or textelement or pictureelement


               
                llElement = pPageLayoutGC.Next
            Loop


            gISSIMxDoc.ActiveView.Refresh()

        
            pMapLegendOrientation = pPageOrientation & pLegendOrientation
          
        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "GetTemplateandLogo", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try
       

        Dim arrPageInfo(2) As String
        arrPageInfo(0) = strPageSize
        arrPageInfo(1) = pMapLegendOrientation
        arrPageInfo(2) = p_LogoName
       
        Return arrPageInfo
    End Function

End Module
