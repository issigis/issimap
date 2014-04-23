<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmISSIMap
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmISSIMap))
        Me.BtnClose = New System.Windows.Forms.Button()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.tboNotes = New System.Windows.Forms.RichTextBox()
        Me.LblNotes = New System.Windows.Forms.Label()
        Me.LblProject = New System.Windows.Forms.Label()
        Me.LblDate = New System.Windows.Forms.Label()
        Me.LblClient = New System.Windows.Forms.Label()
        Me.LblAuthor = New System.Windows.Forms.Label()
        Me.LblTitle = New System.Windows.Forms.Label()
        Me.ProjectTextBox = New System.Windows.Forms.TextBox()
        Me.ClientTextBox = New System.Windows.Forms.TextBox()
        Me.txtDate = New System.Windows.Forms.TextBox()
        Me.AuthorTextBox = New System.Windows.Forms.TextBox()
        Me.txtTitle3 = New System.Windows.Forms.TextBox()
        Me.txtTitle2 = New System.Windows.Forms.TextBox()
        Me.txtTitle1 = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.txtLogoPath = New System.Windows.Forms.TextBox()
        Me.Panel_TemplateList = New System.Windows.Forms.Panel()
        Me.Option4_PortHorz = New System.Windows.Forms.RadioButton()
        Me.Option3_PortVert = New System.Windows.Forms.RadioButton()
        Me.Option2_LandHorz = New System.Windows.Forms.RadioButton()
        Me.TemplateImageList = New System.Windows.Forms.ImageList(Me.components)
        Me.Option1_LandVert = New System.Windows.Forms.RadioButton()
        Me.lblIndex = New System.Windows.Forms.Label()
        Me.LblLayers = New System.Windows.Forms.Label()
        Me.BtnSelectLogo = New System.Windows.Forms.Button()
        Me.cboIndexMap = New System.Windows.Forms.ComboBox()
        Me.cboMainMap = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.LblPageSize = New System.Windows.Forms.Label()
        Me.PageSizeListBox = New System.Windows.Forms.ListBox()
        Me.pboTemplateImage = New System.Windows.Forms.PictureBox()
        Me.BtnCreateMap = New System.Windows.Forms.Button()
        Me.SelectImageDialog = New System.Windows.Forms.OpenFileDialog()
        Me.PicBox_ISSIHelp = New System.Windows.Forms.PictureBox()
        Me.PicBoxISSILetterhead = New System.Windows.Forms.PictureBox()
        Me.tooltipInput = New System.Windows.Forms.ToolTip(Me.components)
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.Panel_TemplateList.SuspendLayout()
        CType(Me.pboTemplateImage, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBox_ISSIHelp, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.PicBoxISSILetterhead, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'BtnClose
        '
        Me.BtnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.BtnClose.Location = New System.Drawing.Point(541, 438)
        Me.BtnClose.Name = "BtnClose"
        Me.BtnClose.Size = New System.Drawing.Size(120, 40)
        Me.BtnClose.TabIndex = 0
        Me.BtnClose.Text = "Close Form"
        Me.BtnClose.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.tboNotes)
        Me.GroupBox1.Controls.Add(Me.LblNotes)
        Me.GroupBox1.Controls.Add(Me.LblProject)
        Me.GroupBox1.Controls.Add(Me.LblDate)
        Me.GroupBox1.Controls.Add(Me.LblClient)
        Me.GroupBox1.Controls.Add(Me.LblAuthor)
        Me.GroupBox1.Controls.Add(Me.LblTitle)
        Me.GroupBox1.Controls.Add(Me.ProjectTextBox)
        Me.GroupBox1.Controls.Add(Me.ClientTextBox)
        Me.GroupBox1.Controls.Add(Me.txtDate)
        Me.GroupBox1.Controls.Add(Me.AuthorTextBox)
        Me.GroupBox1.Controls.Add(Me.txtTitle3)
        Me.GroupBox1.Controls.Add(Me.txtTitle2)
        Me.GroupBox1.Controls.Add(Me.txtTitle1)
        Me.GroupBox1.Location = New System.Drawing.Point(8, 79)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(214, 396)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Map Information"
        '
        'tboNotes
        '
        Me.tboNotes.AcceptsTab = True
        Me.tboNotes.Location = New System.Drawing.Point(6, 321)
        Me.tboNotes.MaxLength = 200
        Me.tboNotes.Name = "tboNotes"
        Me.tboNotes.Size = New System.Drawing.Size(200, 68)
        Me.tboNotes.TabIndex = 8
        Me.tboNotes.Text = ""
        Me.tooltipInput.SetToolTip(Me.tboNotes, "Notes" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Enter Notes about the Project or Map.")
        '
        'LblNotes
        '
        Me.LblNotes.AutoSize = True
        Me.LblNotes.Location = New System.Drawing.Point(10, 301)
        Me.LblNotes.Name = "LblNotes"
        Me.LblNotes.Size = New System.Drawing.Size(38, 13)
        Me.LblNotes.TabIndex = 13
        Me.LblNotes.Text = "Notes:"
        '
        'LblProject
        '
        Me.LblProject.AutoSize = True
        Me.LblProject.Location = New System.Drawing.Point(10, 256)
        Me.LblProject.Name = "LblProject"
        Me.LblProject.Size = New System.Drawing.Size(46, 13)
        Me.LblProject.TabIndex = 12
        Me.LblProject.Text = "Project: "
        '
        'LblDate
        '
        Me.LblDate.AutoSize = True
        Me.LblDate.Location = New System.Drawing.Point(10, 211)
        Me.LblDate.Name = "LblDate"
        Me.LblDate.Size = New System.Drawing.Size(33, 13)
        Me.LblDate.TabIndex = 11
        Me.LblDate.Text = "Date:"
        '
        'LblClient
        '
        Me.LblClient.AutoSize = True
        Me.LblClient.Location = New System.Drawing.Point(10, 165)
        Me.LblClient.Name = "LblClient"
        Me.LblClient.Size = New System.Drawing.Size(59, 13)
        Me.LblClient.TabIndex = 10
        Me.LblClient.Text = "Issued For:"
        '
        'LblAuthor
        '
        Me.LblAuthor.AutoSize = True
        Me.LblAuthor.Location = New System.Drawing.Point(10, 119)
        Me.LblAuthor.Name = "LblAuthor"
        Me.LblAuthor.Size = New System.Drawing.Size(68, 13)
        Me.LblAuthor.TabIndex = 9
        Me.LblAuthor.Text = "Prepared By:"
        '
        'LblTitle
        '
        Me.LblTitle.AutoSize = True
        Me.LblTitle.Location = New System.Drawing.Point(10, 18)
        Me.LblTitle.Name = "LblTitle"
        Me.LblTitle.Size = New System.Drawing.Size(30, 13)
        Me.LblTitle.TabIndex = 8
        Me.LblTitle.Text = "Title:"
        '
        'ProjectTextBox
        '
        Me.ProjectTextBox.Location = New System.Drawing.Point(6, 275)
        Me.ProjectTextBox.MaxLength = 50
        Me.ProjectTextBox.Name = "ProjectTextBox"
        Me.ProjectTextBox.Size = New System.Drawing.Size(200, 20)
        Me.ProjectTextBox.TabIndex = 6
        '
        'ClientTextBox
        '
        Me.ClientTextBox.Location = New System.Drawing.Point(6, 184)
        Me.ClientTextBox.MaxLength = 40
        Me.ClientTextBox.Name = "ClientTextBox"
        Me.ClientTextBox.Size = New System.Drawing.Size(200, 20)
        Me.ClientTextBox.TabIndex = 4
        '
        'txtDate
        '
        Me.txtDate.Location = New System.Drawing.Point(6, 230)
        Me.txtDate.MaxLength = 30
        Me.txtDate.Name = "txtDate"
        Me.txtDate.Size = New System.Drawing.Size(200, 20)
        Me.txtDate.TabIndex = 5
        Me.tooltipInput.SetToolTip(Me.txtDate, "Date" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Double Click to Insert Today's Date")
        Me.txtDate.WordWrap = False
        '
        'AuthorTextBox
        '
        Me.AuthorTextBox.Location = New System.Drawing.Point(6, 138)
        Me.AuthorTextBox.MaxLength = 40
        Me.AuthorTextBox.Name = "AuthorTextBox"
        Me.AuthorTextBox.Size = New System.Drawing.Size(200, 20)
        Me.AuthorTextBox.TabIndex = 3
        '
        'txtTitle3
        '
        Me.txtTitle3.Location = New System.Drawing.Point(6, 90)
        Me.txtTitle3.MaxLength = 28
        Me.txtTitle3.Name = "txtTitle3"
        Me.txtTitle3.Size = New System.Drawing.Size(200, 20)
        Me.txtTitle3.TabIndex = 2
        Me.txtTitle3.WordWrap = False
        '
        'txtTitle2
        '
        Me.txtTitle2.Location = New System.Drawing.Point(6, 64)
        Me.txtTitle2.MaxLength = 28
        Me.txtTitle2.Name = "txtTitle2"
        Me.txtTitle2.Size = New System.Drawing.Size(200, 20)
        Me.txtTitle2.TabIndex = 1
        Me.txtTitle2.WordWrap = False
        '
        'txtTitle1
        '
        Me.txtTitle1.Location = New System.Drawing.Point(6, 38)
        Me.txtTitle1.MaxLength = 28
        Me.txtTitle1.Name = "txtTitle1"
        Me.txtTitle1.Size = New System.Drawing.Size(200, 20)
        Me.txtTitle1.TabIndex = 0
        Me.txtTitle1.WordWrap = False
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.txtLogoPath)
        Me.GroupBox2.Controls.Add(Me.Panel_TemplateList)
        Me.GroupBox2.Controls.Add(Me.lblIndex)
        Me.GroupBox2.Controls.Add(Me.LblLayers)
        Me.GroupBox2.Controls.Add(Me.BtnSelectLogo)
        Me.GroupBox2.Controls.Add(Me.cboIndexMap)
        Me.GroupBox2.Controls.Add(Me.cboMainMap)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Controls.Add(Me.LblPageSize)
        Me.GroupBox2.Controls.Add(Me.PageSizeListBox)
        Me.GroupBox2.Controls.Add(Me.pboTemplateImage)
        Me.GroupBox2.Location = New System.Drawing.Point(230, 79)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(434, 346)
        Me.GroupBox2.TabIndex = 2
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Layouts"
        '
        'txtLogoPath
        '
        Me.txtLogoPath.Location = New System.Drawing.Point(12, 311)
        Me.txtLogoPath.Name = "txtLogoPath"
        Me.txtLogoPath.Size = New System.Drawing.Size(185, 20)
        Me.txtLogoPath.TabIndex = 11
        Me.txtLogoPath.TabStop = False
        '
        'Panel_TemplateList
        '
        Me.Panel_TemplateList.Controls.Add(Me.Option4_PortHorz)
        Me.Panel_TemplateList.Controls.Add(Me.Option3_PortVert)
        Me.Panel_TemplateList.Controls.Add(Me.Option2_LandHorz)
        Me.Panel_TemplateList.Controls.Add(Me.Option1_LandVert)
        Me.Panel_TemplateList.Location = New System.Drawing.Point(212, 242)
        Me.Panel_TemplateList.Name = "Panel_TemplateList"
        Me.Panel_TemplateList.Size = New System.Drawing.Size(205, 89)
        Me.Panel_TemplateList.TabIndex = 10
        '
        'Option4_PortHorz
        '
        Me.Option4_PortHorz.Appearance = System.Windows.Forms.Appearance.Button
        Me.Option4_PortHorz.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option4_PortHorz.Location = New System.Drawing.Point(102, 44)
        Me.Option4_PortHorz.Name = "Option4_PortHorz"
        Me.Option4_PortHorz.Size = New System.Drawing.Size(100, 40)
        Me.Option4_PortHorz.TabIndex = 3
        Me.Option4_PortHorz.Text = "Portrait 2"
        Me.Option4_PortHorz.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Option4_PortHorz.UseVisualStyleBackColor = True
        '
        'Option3_PortVert
        '
        Me.Option3_PortVert.Appearance = System.Windows.Forms.Appearance.Button
        Me.Option3_PortVert.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option3_PortVert.Location = New System.Drawing.Point(103, 4)
        Me.Option3_PortVert.Name = "Option3_PortVert"
        Me.Option3_PortVert.Size = New System.Drawing.Size(100, 40)
        Me.Option3_PortVert.TabIndex = 2
        Me.Option3_PortVert.Text = "Portrait 1"
        Me.Option3_PortVert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Option3_PortVert.UseVisualStyleBackColor = True
        '
        'Option2_LandHorz
        '
        Me.Option2_LandHorz.Appearance = System.Windows.Forms.Appearance.Button
        Me.Option2_LandHorz.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option2_LandHorz.ImageList = Me.TemplateImageList
        Me.Option2_LandHorz.Location = New System.Drawing.Point(3, 44)
        Me.Option2_LandHorz.Name = "Option2_LandHorz"
        Me.Option2_LandHorz.Size = New System.Drawing.Size(100, 40)
        Me.Option2_LandHorz.TabIndex = 1
        Me.Option2_LandHorz.Text = "Landscape 2"
        Me.Option2_LandHorz.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Option2_LandHorz.UseVisualStyleBackColor = True
        '
        'TemplateImageList
        '
        Me.TemplateImageList.ImageStream = CType(resources.GetObject("TemplateImageList.ImageStream"), System.Windows.Forms.ImageListStreamer)
        Me.TemplateImageList.TransparentColor = System.Drawing.Color.Transparent
        Me.TemplateImageList.Images.SetKeyName(0, "Land_Horiz_A.png")
        Me.TemplateImageList.Images.SetKeyName(1, "Land_Horiz_A_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(2, "Land_Horiz_B.png")
        Me.TemplateImageList.Images.SetKeyName(3, "Land_Horiz_B_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(4, "Land_Vert_A.png")
        Me.TemplateImageList.Images.SetKeyName(5, "Land_Vert_A_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(6, "Land_Vert_B.png")
        Me.TemplateImageList.Images.SetKeyName(7, "Land_Vert_B_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(8, "Port_Horiz_A.png")
        Me.TemplateImageList.Images.SetKeyName(9, "Port_Horiz_A_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(10, "Port_Horiz_B.png")
        Me.TemplateImageList.Images.SetKeyName(11, "Port_Horiz_B_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(12, "Port_Vert_A.png")
        Me.TemplateImageList.Images.SetKeyName(13, "Port_Vert_A_noIndex.png")
        Me.TemplateImageList.Images.SetKeyName(14, "Port_Vert_B.png")
        Me.TemplateImageList.Images.SetKeyName(15, "Port_Vert_B_noIndex.png")
        '
        'Option1_LandVert
        '
        Me.Option1_LandVert.Appearance = System.Windows.Forms.Appearance.Button
        Me.Option1_LandVert.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Option1_LandVert.ImageList = Me.TemplateImageList
        Me.Option1_LandVert.Location = New System.Drawing.Point(3, 4)
        Me.Option1_LandVert.Name = "Option1_LandVert"
        Me.Option1_LandVert.Size = New System.Drawing.Size(100, 40)
        Me.Option1_LandVert.TabIndex = 0
        Me.Option1_LandVert.Text = "Landscape 1"
        Me.Option1_LandVert.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Option1_LandVert.UseVisualStyleBackColor = True
        '
        'lblIndex
        '
        Me.lblIndex.AutoSize = True
        Me.lblIndex.Location = New System.Drawing.Point(26, 188)
        Me.lblIndex.Name = "lblIndex"
        Me.lblIndex.Size = New System.Drawing.Size(60, 13)
        Me.lblIndex.TabIndex = 9
        Me.lblIndex.Text = "Index Map:"
        '
        'LblLayers
        '
        Me.LblLayers.AutoSize = True
        Me.LblLayers.Location = New System.Drawing.Point(26, 139)
        Me.LblLayers.Name = "LblLayers"
        Me.LblLayers.Size = New System.Drawing.Size(57, 13)
        Me.LblLayers.TabIndex = 8
        Me.LblLayers.Text = "Main Map:"
        '
        'BtnSelectLogo
        '
        Me.BtnSelectLogo.Location = New System.Drawing.Point(12, 271)
        Me.BtnSelectLogo.Name = "BtnSelectLogo"
        Me.BtnSelectLogo.Size = New System.Drawing.Size(100, 34)
        Me.BtnSelectLogo.TabIndex = 4
        Me.BtnSelectLogo.Text = "Select Image ..."
        Me.BtnSelectLogo.UseVisualStyleBackColor = True
        '
        'cboIndexMap
        '
        Me.cboIndexMap.FormattingEnabled = True
        Me.cboIndexMap.Location = New System.Drawing.Point(25, 211)
        Me.cboIndexMap.Name = "cboIndexMap"
        Me.cboIndexMap.Size = New System.Drawing.Size(140, 21)
        Me.cboIndexMap.TabIndex = 9
        '
        'cboMainMap
        '
        Me.cboMainMap.FormattingEnabled = True
        Me.cboMainMap.Location = New System.Drawing.Point(25, 160)
        Me.cboMainMap.Name = "cboMainMap"
        Me.cboMainMap.Size = New System.Drawing.Size(140, 21)
        Me.cboMainMap.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 121)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(70, 13)
        Me.Label1.TabIndex = 3
        Me.Label1.Text = "Data Frames:"
        '
        'LblPageSize
        '
        Me.LblPageSize.AutoSize = True
        Me.LblPageSize.Location = New System.Drawing.Point(7, 17)
        Me.LblPageSize.Name = "LblPageSize"
        Me.LblPageSize.Size = New System.Drawing.Size(58, 13)
        Me.LblPageSize.TabIndex = 2
        Me.LblPageSize.Text = "Page Size:"
        '
        'PageSizeListBox
        '
        Me.PageSizeListBox.FormattingEnabled = True
        Me.PageSizeListBox.Items.AddRange(New Object() {"A Size - 8.5 x 11", "B Size - 11 x 17", "C Size - 17 x 22", "D Size - 22 x 34", "E Size - 34 x 44"})
        Me.PageSizeListBox.Location = New System.Drawing.Point(25, 37)
        Me.PageSizeListBox.Name = "PageSizeListBox"
        Me.PageSizeListBox.Size = New System.Drawing.Size(100, 69)
        Me.PageSizeListBox.TabIndex = 1
        '
        'pboTemplateImage
        '
        Me.pboTemplateImage.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.pboTemplateImage.Image = CType(resources.GetObject("pboTemplateImage.Image"), System.Drawing.Image)
        Me.pboTemplateImage.InitialImage = CType(resources.GetObject("pboTemplateImage.InitialImage"), System.Drawing.Image)
        Me.pboTemplateImage.Location = New System.Drawing.Point(204, 13)
        Me.pboTemplateImage.Name = "pboTemplateImage"
        Me.pboTemplateImage.Size = New System.Drawing.Size(220, 220)
        Me.pboTemplateImage.TabIndex = 0
        Me.pboTemplateImage.TabStop = False
        Me.pboTemplateImage.WaitOnLoad = True
        '
        'BtnCreateMap
        '
        Me.BtnCreateMap.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.BtnCreateMap.Location = New System.Drawing.Point(405, 438)
        Me.BtnCreateMap.Name = "BtnCreateMap"
        Me.BtnCreateMap.Size = New System.Drawing.Size(120, 40)
        Me.BtnCreateMap.TabIndex = 3
        Me.BtnCreateMap.Text = "Create Map"
        Me.BtnCreateMap.UseVisualStyleBackColor = True
        '
        'SelectImageDialog
        '
        Me.SelectImageDialog.Filter = "Image Files (*.bmp, *.jpg,*.png, *.tif, *.emf)|*.bmp;*.jpg;*.png;*.tif;*.emf)"
        Me.SelectImageDialog.Title = "Select an Image"
        '
        'PicBox_ISSIHelp
        '
        Me.PicBox_ISSIHelp.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.PicBox_ISSIHelp.Image = CType(resources.GetObject("PicBox_ISSIHelp.Image"), System.Drawing.Image)
        Me.PicBox_ISSIHelp.Location = New System.Drawing.Point(238, 433)
        Me.PicBox_ISSIHelp.Name = "PicBox_ISSIHelp"
        Me.PicBox_ISSIHelp.Size = New System.Drawing.Size(56, 45)
        Me.PicBox_ISSIHelp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage
        Me.PicBox_ISSIHelp.TabIndex = 5
        Me.PicBox_ISSIHelp.TabStop = False
        '
        'PicBoxISSILetterhead
        '
        Me.PicBoxISSILetterhead.BackColor = System.Drawing.Color.White
        Me.PicBoxISSILetterhead.Dock = System.Windows.Forms.DockStyle.Top
        Me.PicBoxISSILetterhead.Image = CType(resources.GetObject("PicBoxISSILetterhead.Image"), System.Drawing.Image)
        Me.PicBoxISSILetterhead.Location = New System.Drawing.Point(0, 0)
        Me.PicBoxISSILetterhead.Name = "PicBoxISSILetterhead"
        Me.PicBoxISSILetterhead.Padding = New System.Windows.Forms.Padding(6, 6, 6, 8)
        Me.PicBoxISSILetterhead.Size = New System.Drawing.Size(674, 65)
        Me.PicBoxISSILetterhead.TabIndex = 7
        Me.PicBoxISSILetterhead.TabStop = False
        '
        'FrmISSIMap
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.BtnClose
        Me.ClientSize = New System.Drawing.Size(674, 492)
        Me.Controls.Add(Me.PicBoxISSILetterhead)
        Me.Controls.Add(Me.PicBox_ISSIHelp)
        Me.Controls.Add(Me.BtnCreateMap)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.BtnClose)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "FrmISSIMap"
        Me.ShowIcon = False
        Me.Text = "ISSI Map"
        Me.TopMost = True
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.Panel_TemplateList.ResumeLayout(False)
        CType(Me.pboTemplateImage, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBox_ISSIHelp, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.PicBoxISSILetterhead, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents BtnClose As System.Windows.Forms.Button
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents BtnCreateMap As System.Windows.Forms.Button
    Friend WithEvents pboTemplateImage As System.Windows.Forms.PictureBox
    Friend WithEvents TemplateImageList As System.Windows.Forms.ImageList
    Friend WithEvents SelectImageDialog As System.Windows.Forms.OpenFileDialog
    Friend WithEvents BtnSelectLogo As System.Windows.Forms.Button


    Friend WithEvents LblNotes As System.Windows.Forms.Label
    Friend WithEvents LblProject As System.Windows.Forms.Label
    Friend WithEvents LblDate As System.Windows.Forms.Label
    Friend WithEvents LblClient As System.Windows.Forms.Label
    Friend WithEvents LblAuthor As System.Windows.Forms.Label
    Friend WithEvents LblTitle As System.Windows.Forms.Label
    Friend WithEvents ProjectTextBox As System.Windows.Forms.TextBox
    Friend WithEvents ClientTextBox As System.Windows.Forms.TextBox
    Friend WithEvents txtDate As System.Windows.Forms.TextBox
    Friend WithEvents AuthorTextBox As System.Windows.Forms.TextBox
    Friend WithEvents txtTitle2 As System.Windows.Forms.TextBox
    Friend WithEvents txtTitle1 As System.Windows.Forms.TextBox
    Friend WithEvents PicBox_ISSIHelp As System.Windows.Forms.PictureBox
    Friend WithEvents LblPageSize As System.Windows.Forms.Label
    Friend WithEvents PageSizeListBox As System.Windows.Forms.ListBox
    Friend WithEvents Panel_TemplateList As System.Windows.Forms.Panel
    Friend WithEvents Option4_PortHorz As System.Windows.Forms.RadioButton
    Friend WithEvents Option3_PortVert As System.Windows.Forms.RadioButton
    Friend WithEvents Option2_LandHorz As System.Windows.Forms.RadioButton
    Friend WithEvents Option1_LandVert As System.Windows.Forms.RadioButton
    Friend WithEvents lblIndex As System.Windows.Forms.Label
    Friend WithEvents LblLayers As System.Windows.Forms.Label
    Friend WithEvents cboIndexMap As System.Windows.Forms.ComboBox
    Friend WithEvents cboMainMap As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents txtLogoPath As System.Windows.Forms.TextBox
    Friend WithEvents PicBoxISSILetterhead As System.Windows.Forms.PictureBox
    Friend WithEvents tooltipInput As System.Windows.Forms.ToolTip
    Friend WithEvents txtTitle3 As System.Windows.Forms.TextBox
    Friend WithEvents tboNotes As System.Windows.Forms.RichTextBox
End Class
