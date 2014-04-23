Imports System
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
Imports System.Windows.Forms

Public Class FrmISSIMap
    Inherits System.Windows.Forms.Form
    Dim gMxApp As IMxApplication
    Dim pApp As IApplication
    Dim pMxApp As IMxApplication
    Public pMxDoc As IMxDocument

    ''Declare Form Variables

    Dim pTemplateDirectory As String     'filedirectory path to logos
    Dim PageSize As String
    Dim boolUseATemplate As Boolean     ' = True use "A" templateshape(8.5X11,17X22,34x44) -- boolUseATemplate = False  use "B" Template shape(11x17,22X34)

    Dim pMainMapName As String
    Dim pIndexMapName As String
    Dim IndexMapExists As Boolean
    Dim pNoIndex As String = "None"

    Dim Title As String
    Dim Notes As String
    Dim Author As String
    Dim Project As String
    Dim Client As String
    Dim pDate As String

    Public selectedPageSizeLetter As String
    Public selectedPageSizeIndex As Integer

    Public WriteOnly Property SetApplication() As IApplication
        Set(value As IApplication)
            gISSIApp = value
            gmxApp = value
            gISSIMxDoc = CType(gISSIApp.Document, IMxDocument)

        End Set
    End Property

    Private Sub BtnClose_Click(sender As System.Object, e As System.EventArgs) Handles BtnClose.Click
        Me.Close()
    End Sub

    Public Sub GetLayoutText()
        gISSIMxDoc = CType(gISSIApp.Document, IMxDocument)

        Dim tElement As IElement
        Dim tElementProperties As IElementProperties3
        Dim strtElementName As String
        Dim tTextElement As ITextElement

        Dim TitleItemVbcrlf() ' as string

        Dim pPagelayoutDisplay As IDisplay
        Dim pPageLayoutGC As IGraphicsContainer
        Dim tPageLayout As IPageLayout3
        tPageLayout = gISSIMxDoc.PageLayout
        pPageLayoutGC = tPageLayout
        pPagelayoutDisplay = pMxDoc.ActiveView.ScreenDisplay

        Dim strItemsCollected As String = ""
        Dim pPageLayoutGCSelect As IGraphicsContainerSelect
        pPageLayoutGCSelect = pPageLayoutGC
        pPageLayoutGCSelect.SelectAllElements()

        pPageLayoutGC.Reset()
        tElement = pPageLayoutGC.Next

        Try
            'Find Text Elements
            Do Until tElement Is Nothing

                ''PUT SELECT CASE OF TEXT ELEMENTS TO SET TEXT IN HERE
                tElementProperties = tElement
                strtElementName = tElementProperties.Name

                If tElementProperties.Type.ToString = "Text" Then

                    Select Case strtElementName
                        Case "theTitle"
                            tTextElement = tElement  
                            Dim pCurrentTitle As String = tTextElement.Text

                            If Microsoft.VisualBasic.Trim(pCurrentTitle) <> "" Then
                                TitleItemVbcrlf = Microsoft.VisualBasic.Split(pCurrentTitle, vbCrLf)

                                If TitleItemVbcrlf(0).trim <> "" Then
                                    txtTitle1.Text = TitleItemVbcrlf(0)
                                End If
                                If TitleItemVbcrlf(1).trim <> "" Then
                                    txtTitle2.Text = TitleItemVbcrlf(1)
                                End If
                                If TitleItemVbcrlf(2).trim <> "" Then
                                    txtTitle3.Text = TitleItemVbcrlf(2)
                                End If

                                If strItemsCollected = "" Then
                                    strItemsCollected = pCurrentTitle
                                Else
                                    strItemsCollected = strItemsCollected & "," & pCurrentTitle
                                End If
                            End If

                        Case "theClient"

                                tTextElement = tElement
                                Dim pClientText As String

                                pClientText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Issued For:", "")
                                ''SET THE ELEMENT WITH THE NEW TEXT
                                If Microsoft.VisualBasic.Trim(pClientText) <> "" Then
                                    ClientTextBox.Text = Microsoft.VisualBasic.Trim(pClientText)
                                    If strItemsCollected = "" Then
                                        strItemsCollected = pClientText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pClientText
                                    End If
                                End If

                        Case "theDate"

                                tTextElement = tElement
                                Dim pDateText As String
                                pDateText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Date:", "")
                                ''SET THE ELEMENT WITH THE NEW TEXT
                                If Microsoft.VisualBasic.Trim(pDateText) <> "" Then
                                    txtDate.Text = Microsoft.VisualBasic.Trim(pDateText)
                                    If strItemsCollected = "" Then
                                        strItemsCollected = pDateText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pDateText
                                    End If
                                End If

                        Case "theMapAuthor"

                                tTextElement = tElement
                                Dim pAuthorText As String
                                pAuthorText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Prepared By:", "")
                                If Microsoft.VisualBasic.Trim(pAuthorText) <> "" Then
                                    AuthorTextBox.Text = Microsoft.VisualBasic.Trim(pAuthorText)

                                    If strItemsCollected = "" Then
                                        strItemsCollected = pAuthorText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pAuthorText
                                    End If
                                End If

                        Case "theProjNo"

                                tTextElement = tElement
                                Dim pProjNumText As String
                                pProjNumText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Project:", "")
                                If Microsoft.VisualBasic.Trim(pProjNumText) <> "" Then
                                    ''SET THE ELEMENT WITH THE NEW TEXT
                                    ProjectTextBox.Text = pProjNumText.Trim

                                    If strItemsCollected = "" Then
                                        strItemsCollected = pProjNumText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pProjNumText
                                    End If
                                End If

                        Case "ProjectNotes"

                            tTextElement = tElement
                            Dim pProjNotesText As String
                            pProjNotesText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Notes:", "")
                            pProjNotesText = Microsoft.VisualBasic.Replace(pProjNotesText, vbCrLf, " ")
                            If Microsoft.VisualBasic.Trim(pProjNotesText) <> "" Then
                                ''SET THE ELEMENT WITH THE NEW TEXT
                                tboNotes.Text = pProjNotesText.Trim

                                If strItemsCollected = "" Then
                                    strItemsCollected = pProjNotesText
                                Else
                                    strItemsCollected = strItemsCollected & "," & pProjNotesText
                                End If
                            End If

                        Case Else


                    End Select
                Else
                    'MessageBox.Show(tElementProperties.Name & " is a " & tElement.GetType.ToString, "it is not a Textelement")

                End If ' if  textelement 

                tElement = pPageLayoutGC.Next
            Loop

            pPageLayoutGCSelect.UnselectAllElements()
            gISSIMxDoc.ActiveView.Refresh()


        Catch ex As Exception
            'MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "getLayoutText", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)      
        End Try

        Dim intTextItemCount As Integer = 0
        If strItemsCollected.Trim <> "" Then
            Dim arrTExtItems() As String = Microsoft.VisualBasic.Split(strItemsCollected, ",")
            intTextItemCount = arrTExtItems.Length

        End If

        If strItemsCollected = "" Or intTextItemCount < 2 Then
            ''Try again, missed some elements
            Try
                'Find Text Elements
                pPageLayoutGC = tPageLayout
              
                pPageLayoutGC.Reset()
                tElement = pPageLayoutGC.Next

                Do While Not tElement Is Nothing
                    ''PUT SELECT CASE OF TEXT ELEMENTS TO SET TEXT IN HERE
                    tElementProperties = tElement
                    strtElementName = tElementProperties.Name
                    If tElementProperties.Type.ToString = "Text" Then

                        Select Case strtElementName
                            Case "theTitle"
                                tTextElement = tElement

                                Dim pCurrentTitle As String = tTextElement.Text

                                If Microsoft.VisualBasic.Trim(pCurrentTitle) <> "" Then
                                    TitleItemVbcrlf = Microsoft.VisualBasic.Split(pCurrentTitle, vbCrLf)
                              
                                    txtTitle1.Text = TitleItemVbcrlf(0)
                                    txtTitle2.Text = TitleItemVbcrlf(1)
                                    txtTitle3.Text = TitleItemVbcrlf(2)
                                    If strItemsCollected = "" Then
                                        strItemsCollected = pCurrentTitle
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pCurrentTitle
                                    End If
                                End If

                            Case "theClient"

                                tTextElement = tElement
                                Dim pClientText As String

                                pClientText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Issued For:", "")
                                ''SET THE ELEMENT WITH THE NEW TEXT
                                If Microsoft.VisualBasic.Trim(pClientText) <> "" Then
                                    ClientTextBox.Text = Microsoft.VisualBasic.Trim(pClientText)

                                    If strItemsCollected = "" Then
                                        strItemsCollected = pClientText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pClientText
                                    End If
                                End If

                            Case "theDate"

                                tTextElement = tElement
                                Dim pDateText As String
                                pDateText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Date:", "")
                                ''SET THE ELEMENT WITH THE NEW TEXT
                                If Microsoft.VisualBasic.Trim(pDateText) <> "" Then
                                    txtDate.Text = Microsoft.VisualBasic.Trim(pDateText)

                                    If strItemsCollected = "" Then
                                        strItemsCollected = pDateText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pDateText
                                    End If
                                End If

                            Case "theMapAuthor"

                                tTextElement = tElement
                                Dim pAuthorText As String
                                pAuthorText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Prepared By:", "")
                                If Microsoft.VisualBasic.Trim(pAuthorText) <> "" Then
                                    AuthorTextBox.Text = Microsoft.VisualBasic.Trim(pAuthorText)

                                    If strItemsCollected = "" Then
                                        strItemsCollected = pAuthorText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pAuthorText
                                    End If
                                End If


                            Case "theProjNo"

                                tTextElement = tElement
                                Dim pProjNumText As String
                                pProjNumText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Project:", "")
                                If Microsoft.VisualBasic.Trim(pProjNumText) <> "" Then
                                    ''SET THE ELEMENT WITH THE NEW TEXT
                                    ProjectTextBox.Text = pProjNumText.Trim


                                    If strItemsCollected = "" Then
                                        strItemsCollected = pProjNumText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pProjNumText
                                    End If
                                End If

                            Case "ProjectNotes"

                                tTextElement = tElement
                                Dim pProjNotesText As String
                                pProjNotesText = Microsoft.VisualBasic.Replace(tTextElement.Text, "Notes:", "")
                                If Microsoft.VisualBasic.Trim(pProjNotesText) <> "" Then
                                    ''SET THE ELEMENT WITH THE NEW TEXT
                                    tboNotes.Text = pProjNotesText.Trim

                                    If strItemsCollected = "" Then
                                        strItemsCollected = pProjNotesText
                                    Else
                                        strItemsCollected = strItemsCollected & "," & pProjNotesText
                                    End If
                                End If

                            Case Else


                        End Select
                    Else
                        'MessageBox.Show(tElementProperties.Name & " is a " & tElement.GetType.ToString, "it is not a Textelement")

                    End If ' if  textelement 

                    tElement = pPageLayoutGC.Next
                Loop


                gISSIMxDoc.ActiveView.Refresh()





            Catch ex As Exception
                MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "getLayoutText-2loop", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

            End Try

        End If
        
    End Sub
    Public Sub CheckIMapNames()
        'Determine if Index Map Box should be available
        Dim pMaps As IMaps
        Dim pMainMap As IMap
        Dim pIndexMap As IMap
        Dim pMainMapName As String
        Dim pIndexMapName As String
        Dim IndexMapExists As Boolean
        Dim MainMapExists As Boolean

        pMaps = pMxDoc.Maps
        pMainMapName = "Layers"
        pIndexMapName = "Index Map"
        pMainMapName = cboMainMap.Text
        pIndexMapName = cboIndexMap.Text
        Try


            ''Set the property of the Index Map Check box based on the number of data frames in the map

            Dim pMapNameArray() As String 'to add data frame names to

            If pMaps.Count > 1 Then     'there is more than data frame - enable the index check box
                Dim i As Integer
                For i = 0 To pMaps.Count - 1                    'Loop thru Maps Looking For "Index Map"
                    If pMaps.Item(i).Name = pIndexMapName Then
                        pIndexMap = pMaps.Item(i)              'Set IMap to Approp Data frame by NAME
                        IndexMapExists = True
                    ElseIf pMaps.Item(i).Name = pMainMapName Then
                        pMainMap = pMaps.Item(i)
                        MainMapExists = True
                    Else
                        ReDim Preserve pMapNameArray(UBound(pMapNameArray) + 1)
                        pMapNameArray(i) = pMaps.Item(i).Name
                    End If
                Next i
            Else
                pMainMap = pMaps.Item(0)

            End If

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "checkimapnames", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try

    End Sub

    Public Function SetISSIMapDirs(gISSIMapDir) As Boolean
        'THIS was from ISSIMap Form initialize and form load
        'NOT SURE need this still
        Dim pBrowseDir As String            'FilepathString of Directory to look in for a directory
        Dim pDirNeeded As String
        Dim atemplate As String

        'IF ISSI DIR DOESNT EXIST - HAVE USER BROWSE TO ONE With ProperFiles
        If (Dir(gISSIMapDir, vbDirectory) = "") Then 'if directory was not there
            pDirNeeded = "ISSI Map Tool Directory"
            pBrowseDir = FindFileDir(gISSIMapDir)
            gISSIMapDir = UserSelectDirectory(pBrowseDir, pDirNeeded) 'returns the selected dir
        End If

        ''Set Template Directory Using MapToolDir
        gTemplateDir = gISSIMapDir & "Template"

        ''Check is valid directory
        If (Dir(gTemplateDir, vbDirectory) = "") Then   'there is no directory Template
            pDirNeeded = "ISSI Templates"
            pBrowseDir = FindFileDir(gTemplateDir)
            gTemplateDir = UserSelectDirectory(pBrowseDir, pDirNeeded)
        End If

        ''CHECK FOR MAP TEMPLATES

        If (Dir(gTemplateDir & "\*.mxd", vbNormal) = "") Then 'no mxds in the directory

            MessageBox.Show("ISSI Tools Directory does not contain Map Templates", "Exiting", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)

            atemplate = ""

       
        Else 'the dirfunction returned an mxt file

            atemplate = (Dir(gTemplateDir & "\*.mxd", vbNormal))

        End If 'if templatedir contains templates


        ''SET LOGO IMAGE DIRECTORY
        gLogoImgDir = gISSIMapDir & "Logo"
        If (Dir(gLogoImgDir, vbDirectory) = "") Then
            gLogoImgDir = gISSIMapDir
        End If

        If gTemplateDir = "" Or atemplate = "" Then
            MsgBox("Necessary ISSI Tool directories " & vbCrLf & "DO NOT EXIST or CONTAIN NO Map Templates", vbExclamation, "EXITING")
            SetISSIMapDirs = False
            'Exit Function

        Else
            ''BOTH MapTools AND TEMPLATE DIRECTORY WERE FOUND
            SetISSIMapDirs = True
        End If


    End Function

 

    Private Sub FrmISSIMap_FormClosed(sender As Object, e As System.Windows.Forms.FormClosedEventArgs) Handles Me.FormClosed
    
        m_ISSIMapForm = Nothing
   


        gLogoImgDir = ""       'Default Directory for logo/image to be loaded from

        ''Set Logo Filename and Path
        boolLogoSelected = False
      
        gMainMapName = ""
        gIndexMapName = ""


    End Sub

    Private Sub FrmISSIMap_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

    
        pApp = gISSIApp     ''As IApplication   'pApp - procedural ''gapp - Global

        pMxDoc = gISSIApp.Document 'as IMxDocument
        ''''''''''''''''''''''''''''''''''''''''''
        ''SET ALL INDEPENDENT VALUES
        pNoIndex = "None"      
        cboIndexMap.Items.Insert(0, pNoIndex)
        cboIndexMap.SelectedIndex = 0
       

        ''Set the Date to Today's Date

        txtDate.Text = MonthName(Now.Month, False) + " " + Now.Day.ToString + ", " + Now.Year.ToString
       
        Dim arrPageSizeString() As String
        arrPageSizeString = GetTemplateandLogo()
        Dim arrPageSizes() As String = Microsoft.VisualBasic.Split(arrPageSizeString(0), ",")

        Dim CurrentItem As String
        Dim x As Integer ' index of lbopagesizes
        For x = 0 To PageSizeListBox.Items.Count - 1
            CurrentItem = PageSizeListBox.Items.Item(x)

            If CurrentItem.ToString.Contains(arrPageSizes(0)) And CurrentItem.ToString.Contains(arrPageSizes(1)) Then
                PageSizeListBox.SelectedIndex = x
                Exit For
            Else
                PageSizeListBox.SelectedIndex = 0
            End If
        Next x
        
        ''''''''''''''''''''''''''''''''''''''''''''''''''''
        ''SET DEPENDENT VARIABLES
        Dim pMaps As IMaps
        pMaps = gISSIMxDoc.Maps


        Dim strMainMapName As String = ""
        Dim strIndexMapName As String = ""
        Dim intMainMap As Integer = 0
        Dim intIndexMap As Integer = 0
        Try

            Dim a As Integer
            For a = 0 To pMaps.Count - 1

                If pMaps.Item(a).Name.Contains("Index") Or pMaps.Item(a).Name = strIndexMapName Then
                    intIndexMap = a
                End If
                If pMaps.Item(a).Name.Contains("Layers") Or pMaps.Item(a).Name = strMainMapName Then
                    intMainMap = a
                End If

                cboIndexMap.Items.Add(pMaps.Item(a).Name)
                cboMainMap.Items.Add(pMaps.Item(a).Name)
            Next a

            If pMaps.Count = 1 Then

                cboIndexMap.SelectedItem = "None"
                cboIndexMap.Enabled = False
                cboMainMap.SelectedIndex = intMainMap
            Else
                cboMainMap.SelectedIndex = intMainMap
                If pMaps.Count = 2 Then
                    If intIndexMap + 1 < cboIndexMap.Items.Count Then
                        cboIndexMap.SelectedIndex = intIndexMap + 1 ' have to add 1 to get past "None" option         
                    Else
                        cboIndexMap.SelectedItem = "None"
                    End If
                Else
                    If intIndexMap + 1 < cboIndexMap.Items.Count Then
                        cboIndexMap.SelectedIndex = intIndexMap + 1 ' have to add 1 to get past "None" option 
                        If cboIndexMap.SelectedItem = cboMainMap.SelectedItem Then
                            cboIndexMap.SelectedIndex = intIndexMap + 2
                        End If
                    Else
                        cboIndexMap.SelectedItem = "None"
                    End If
                End If ' if pmaps count > 2
               

            End If
                ''

                Dim pMapTemplate As String
                pMapTemplate = arrPageSizeString(1)

                If pMapTemplate = "LandVert" Then
                    Option1_LandVert.Checked = True

                ElseIf pMapTemplate = "LandHorz" Then
                    Option2_LandHorz.Checked = True

                ElseIf pMapTemplate = "PortVert" Then
                    Option3_PortVert.Checked = True

                ElseIf pMapTemplate = "PortHorz" Then
                    Option4_PortHorz.Checked = True

                End If

                txtLogoPath.Text = arrPageSizeString(2)
               

                Me.TopMost = True

        Catch ex As Exception
            MessageBox.Show(ex.Message.ToString + " " + ex.Source.ToString, "Form Load", MessageBoxButtons.OK, MessageBoxIcon.Error, Nothing, Nothing, True)

        End Try

        Try
            GetLayoutText()
        Catch ex As Exception

        End Try

    End Sub
   

    
   

    Private Sub BtnCreateMap_Click(sender As System.Object, e As System.EventArgs) Handles BtnCreateMap.Click
        ''NEED TO MAKE SURE ARCMAP is NOT MINIMIZED
        Dim arcMapwindowPosition As IWindowPosition = gISSIApp
        If arcMapwindowPosition.State = esriWindowState.esriWSMinimize Then
            arcMapwindowPosition.State = esriWindowState.esriWSNormal

            pMxDoc.ActiveView.ScreenDisplay.UpdateWindow()
            pMxDoc.ActiveView.Refresh()
            gISSIApp.RefreshWindow()
        End If
       
        'Check for Template - CANNOT DO ANYthing without it   
        Dim MapTemplateFullPath As String = gTemplateDir & "\" & MapTemplateMxd

        If Not Microsoft.VisualBasic.FileIO.FileSystem.FileExists(MapTemplateFullPath) Then
            MessageBox.Show(MapTemplateMxd & " Template not found." & vbCrLf & "Please select another Template.", "Template Not Found.", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Exit Sub
        End If

        'Collect all input text
        Dim fNotes As String = ""
        Dim fTitle As String = ""

        Dim Title1 As String = ""
        Dim Title2 As String = ""
        Dim Title3 As String = ""

        'GET ALL TEXT INPUTS AND BOX SELECTIONS
        ''CONCATENATE THE THREE TITLE LINES FOR THE GRAPHIC TEXT ELEMENT
        Title1 = txtTitle1.Text
        Title2 = txtTitle2.Text
        Title3 = txtTitle3.Text

        ''if the title 1 is not blank append the second title to it
        If Len(Title1) <> 0 Then
            fTitle = Title1 & vbCrLf & Title2
        Else
            fTitle = Title2  'if title1 is blank set the title to the title2
        End If

        ''CHECK THE LENGTH OF THE TITLE STRING
        If Len(Trim(fTitle)) = 0 Then
            fTitle = ""             ''NEED TO SET TO EMPTY STRING TO ELIMINATE CARRAIGE RETURNS OF BLANK LINES
        End If

        If Len(Title3) <> 0 Then
            fTitle = fTitle & vbCrLf & Title3
        End If

        ''GET PROJECT NOTES TEXT AND FORMAT THE LINE LENGTH
        Dim arrNotesVbcrlf() As String
        Dim strInputNotes As String
        Dim strNotes As String = ""
        strInputNotes = tboNotes.Text
        ''NEED To replace this section with a loop for each type of "carraige return"
        strInputNotes = Microsoft.VisualBasic.Replace(strInputNotes, vbCrLf, "  ") 'remove carraige returns 
        strInputNotes = Microsoft.VisualBasic.Replace(strInputNotes, vbCr, "  ")
        strInputNotes = Microsoft.VisualBasic.Trim(strInputNotes)



        If Len(strInputNotes) <> 0 Then
            '    ''CREATE AN ARRAY FROM THE TEXT SEPARATED BY CARRIAGE RETURNS (there should ONLY BE ONE AT THE END)
            arrNotesVbcrlf = Microsoft.VisualBasic.Split(strInputNotes, vbCrLf)
            If arrNotesVbcrlf.Length > 5 Then
                strNotes = Microsoft.VisualBasic.Replace(strInputNotes, vbCrLf, " ") 'remove carraige returns
            Else
                strNotes = strInputNotes
            End If


        End If

        fNotes = Trim(strNotes)
 
        Author = AuthorTextBox.Text
        Project = ProjectTextBox.Text
        Client = ClientTextBox.Text

        Title = fTitle
        Notes = fNotes

        pDate = txtDate.Text
        Dim pDateString As String
        pDateString = txtDate.Text

        ''Set UP LOGO PATHS
        If Trim(txtLogoPath.Text) = "" Then
            boolLogoSelected = False
        Else
            boolLogoSelected = True
            gLogoFullPath = Trim(txtLogoPath.Text)

        End If

        pTemplateDirectory = gTemplateDir

        gMainMapName = cboMainMap.Text
        gIndexMapName = cboIndexMap.Text

        ReorderMaps()

        SetPrintPageLayout(PageSize)

        If SetMapTemplate(MapTemplateMxd) Then
            ''pass image infomration to ISSIMAPAction
            ISSIMapAction(MapTemplateMxd, pTemplateDirectory, PageSize, Title, Author, Project, Client, pDateString, Notes, txtLogoPath)

            Me.Close()
        End If

    End Sub

    Private Sub Option1_LandVert_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles Option1_LandVert.CheckedChanged
        If Option1_LandVert.Checked = True Then
            Option1_LandVert.PerformClick()
        End If ' if it is checked
    End Sub

    Private Sub Option4_PortHorz_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles Option4_PortHorz.CheckedChanged
        If Option4_PortHorz.Checked = True Then
            Option4_PortHorz.PerformClick()
        End If ' if is selected
    End Sub

    Private Sub Option2_LandHorz_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles Option2_LandHorz.CheckedChanged
        If Option2_LandHorz.Checked = True Then
            Option2_LandHorz.PerformClick()
        End If ' if it is checked

    End Sub

    Private Sub Option3_PortVert_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles Option3_PortVert.CheckedChanged
        If Option3_PortVert.Checked = True Then
            Option3_PortVert.PerformClick()
        End If
    End Sub

    Private Sub PageSizeListBox_Click(sender As Object, e As System.EventArgs) Handles PageSizeListBox.Click


    End Sub ' pagesize_click




    Private Sub PageSizeListBox_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles PageSizeListBox.SelectedIndexChanged

        'Get PageSize selected to determine A/C/E or B/D Template
        Dim i As Integer
        For i = 0 To 4
            If PageSizeListBox.SelectedIndex = i Then
                PageSize = PageSizeListBox.Text

                If PageSizeListBox.SelectedIndex = 0 Or PageSizeListBox.SelectedIndex = 2 Or PageSizeListBox.SelectedIndex = 4 Then
                    boolUseATemplate = True ' need this to determine which image to show on form a/c/e or b/d
                Else
                    boolUseATemplate = False


                End If

                ''This is to change the form picture based on the pagesize and layout currently selected
                If Option1_LandVert.Checked = True Then
                    Option1_LandVert.PerformClick()

                ElseIf Option2_LandHorz.Checked = True Then
                    Option2_LandHorz.PerformClick()

                ElseIf Option3_PortVert.Checked = True Then
                    Option3_PortVert.PerformClick()

                ElseIf Option4_PortHorz.Checked = True Then
                    Option4_PortHorz.PerformClick()
                Else
                    Option1_LandVert.PerformClick()

                End If
                Exit For 'once selected one is true exit loop
            End If ' if current pagesize is selected exit once get it

        Next i ' next page size in the list
    End Sub


    Private Sub cboIndexMap_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboIndexMap.SelectedIndexChanged
        If Option1_LandVert.Checked = True Then
            Option1_LandVert.PerformClick()

        ElseIf Option2_LandHorz.Checked = True Then
            Option2_LandHorz.PerformClick()

        ElseIf Option3_PortVert.Checked = True Then
            Option3_PortVert.PerformClick()

        ElseIf Option4_PortHorz.Checked = True Then
            Option4_PortHorz.PerformClick()
        Else
            Option1_LandVert.PerformClick()

        End If

    End Sub

    Private Sub cboMainMap_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboMainMap.SelectedIndexChanged

        If Option1_LandVert.Checked = True Then
            Option1_LandVert.PerformClick()

        ElseIf Option2_LandHorz.Checked = True Then
            Option2_LandHorz.PerformClick()

        ElseIf Option3_PortVert.Checked = True Then
            Option3_PortVert.PerformClick()

        ElseIf Option4_PortHorz.Checked = True Then
            Option4_PortHorz.PerformClick()
        Else
            Option1_LandVert.PerformClick()

        End If
    End Sub

    Private Sub Option1_LandVert_Click(sender As Object, e As System.EventArgs) Handles Option1_LandVert.Click
        
        selectedPageSizeIndex = PageSizeListBox.SelectedIndex
        Select Case selectedPageSizeIndex
            Case 0
                selectedPageSizeLetter = "A"
            Case 1
                selectedPageSizeLetter = "B"
            Case 2
                selectedPageSizeLetter = "C"
            Case 3
                selectedPageSizeLetter = "D"
            Case 4
                selectedPageSizeLetter = "E"

            Case Else
                selectedPageSizeLetter = "A"
        End Select
        If PageSizeListBox.SelectedIndex = 1 Or PageSizeListBox.SelectedIndex = 3 Then
            'template  is B or D use the B template
            boolUseATemplate = False
        Else
            boolUseATemplate = True
        End If
        
        If (cboIndexMap.Text <> "None") And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(4)
            MapTemplateMxd = "Land_Vert_" & selectedPageSizeLetter & ".mxd"
        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(5)
            MapTemplateMxd = "Land_Vert_" & selectedPageSizeLetter & "_noIndex.mxd"

        ElseIf (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(6)
            MapTemplateMxd = "Land_Vert_" & selectedPageSizeLetter & ".mxd"

        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(7)
            MapTemplateMxd = "Land_Vert_" & selectedPageSizeLetter & "_noIndex.mxd"

        End If
    End Sub

    Private Sub Option2_LandHorz_Click(sender As Object, e As System.EventArgs) Handles Option2_LandHorz.Click

        selectedPageSizeIndex = PageSizeListBox.SelectedIndex
        Select Case selectedPageSizeIndex
            Case 0
                selectedPageSizeLetter = "A"
            Case 1
                selectedPageSizeLetter = "B"
            Case 2
                selectedPageSizeLetter = "C"
            Case 3
                selectedPageSizeLetter = "D"
            Case 4
                selectedPageSizeLetter = "E"

            Case Else
                selectedPageSizeLetter = "A"
        End Select

        If PageSizeListBox.SelectedIndex = 1 Or PageSizeListBox.SelectedIndex = 3 Then
            'template  is B or D use the B template
            boolUseATemplate = False
        Else
            boolUseATemplate = True
        End If
        'MessageBox.Show("pNoIndex " & pNoIndex)
        If (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(0)
            MapTemplateMxd = "Land_Horiz_" & selectedPageSizeLetter & ".mxd"
        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(1)
            MapTemplateMxd = "Land_Horiz_" & selectedPageSizeLetter & "_noIndex.mxd"

        ElseIf (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(2)
            MapTemplateMxd = "Land_Horiz_" & selectedPageSizeLetter & ".mxd"
        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(3)
            MapTemplateMxd = "Land_Horiz_" & selectedPageSizeLetter & "_noIndex.mxd"

        End If
    End Sub

    Private Sub Option4_PortHorz_Click(sender As Object, e As System.EventArgs) Handles Option4_PortHorz.Click

        selectedPageSizeIndex = PageSizeListBox.SelectedIndex
        Select Case selectedPageSizeIndex
            Case 0
                selectedPageSizeLetter = "A"
            Case 1
                selectedPageSizeLetter = "B"
            Case 2
                selectedPageSizeLetter = "C"
            Case 3
                selectedPageSizeLetter = "D"
            Case 4
                selectedPageSizeLetter = "E"

            Case Else
                selectedPageSizeLetter = "A"
        End Select

        If PageSizeListBox.SelectedIndex = 1 Or PageSizeListBox.SelectedIndex = 3 Then
            'template  is B or D use the B template
            boolUseATemplate = False
        Else
            boolUseATemplate = True
        End If

        If (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(8)
            MapTemplateMxd = "Port_Horiz_" & selectedPageSizeLetter & ".mxd"
        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(9)
            MapTemplateMxd = "Port_Horiz_" & selectedPageSizeLetter & "_noIndex.mxd"
        ElseIf (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(10)
            MapTemplateMxd = "Port_Horiz_" & selectedPageSizeLetter & ".mxd"
        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(11)
            MapTemplateMxd = "Port_Horiz_" & selectedPageSizeLetter & "_noIndex.mxd"

        End If
    End Sub

    Private Sub Option3_PortVert_Click(sender As Object, e As System.EventArgs) Handles Option3_PortVert.Click

        selectedPageSizeIndex = PageSizeListBox.SelectedIndex
        Select Case selectedPageSizeIndex
            Case 0
                selectedPageSizeLetter = "A"
            Case 1
                selectedPageSizeLetter = "B"
            Case 2
                selectedPageSizeLetter = "C"
            Case 3
                selectedPageSizeLetter = "D"
            Case 4
                selectedPageSizeLetter = "E"

            Case Else
                selectedPageSizeLetter = "A"
        End Select

        If PageSizeListBox.SelectedIndex = 1 Or PageSizeListBox.SelectedIndex = 3 Then
            'template  is B or D use the B template
            boolUseATemplate = False
        Else
            boolUseATemplate = True
        End If

        If (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(12)
            MapTemplateMxd = "Port_Vert_" & selectedPageSizeLetter & ".mxd"

        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = True Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(13)
            MapTemplateMxd = "Port_Vert_" & selectedPageSizeLetter & "_noIndex.mxd"
        ElseIf (cboIndexMap.Text <> pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(14)
            MapTemplateMxd = "Port_Vert_" & selectedPageSizeLetter & ".mxd"
        ElseIf (cboIndexMap.Text = pNoIndex) And boolUseATemplate = False Then
            pboTemplateImage.Image = TemplateImageList.Images.Item(15)
            MapTemplateMxd = "Port_Vert_" & selectedPageSizeLetter & "_noIndex.mxd"

        End If
    End Sub

    Private Sub BtnSelectLogo_Click(sender As System.Object, e As System.EventArgs) Handles BtnSelectLogo.Click
        Dim pSelectedImage As String = ""

        SelectImageDialog.ShowDialog()
        pSelectedImage = SelectImageDialog.FileName

        If pSelectedImage <> "" Then
            txtLogoPath.Text = pSelectedImage
        End If

    End Sub

    Private Sub txtTitle1_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtTitle1.KeyUp
        If e.KeyCode = Keys.Tab Then
            txtTitle2.Focus()

        End If
    End Sub
    Private Sub AuthorTextBox_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles AuthorTextBox.KeyUp
        If e.KeyCode = Keys.Tab Then
            m_ISSIMapForm.ClientTextBox.Focus()
        End If

    End Sub

    Private Sub txtTitle1_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtTitle1.TextChanged

    End Sub

    Private Sub txtTitle2_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtTitle2.KeyUp
        If e.KeyCode = Keys.Tab Then
            m_ISSIMapForm.txtTitle3.Focus()
        End If
    End Sub

    Private Sub txtTitle3_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtTitle3.KeyUp
        If e.KeyCode = Keys.Tab Then
            m_ISSIMapForm.AuthorTextBox.Focus()
        End If
    End Sub

    Private Sub AuthorTextBox_TextChanged(sender As System.Object, e As System.EventArgs) Handles AuthorTextBox.TextChanged

    End Sub

    Private Sub ClientTextBox_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ClientTextBox.KeyUp
        If e.KeyCode = Keys.Tab Then
            m_ISSIMapForm.ProjectTextBox.Focus()
        End If
    End Sub


    Private Sub ProjectTextBox_KeyUp(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles ProjectTextBox.KeyUp
        If e.KeyCode = Keys.Tab Then
            m_ISSIMapForm.tboNotes.Focus()
        End If
    End Sub

    Private Sub txtDate_DoubleClick(sender As Object, e As System.EventArgs) Handles txtDate.DoubleClick
        txtDate.Text = MonthName(Now.Month, False) + " " + Now.Day.ToString + ", " + Now.Year.ToString
    End Sub
  
    Private Sub tboNotes_GotFocus(sender As Object, e As System.EventArgs) Handles tboNotes.GotFocus
        Select strPageSizeLetter
            Case "A"
                tboNotes.MaxLength = "200"
            Case "B"
                tboNotes.MaxLength = "200"
            Case "C"
                tboNotes.MaxLength = "350"
            Case "D"
                tboNotes.MaxLength = "350"
            Case "E"
                tboNotes.MaxLength = "400"

        End Select
    End Sub
 
End Class