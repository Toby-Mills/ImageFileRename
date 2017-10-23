Imports System.IO
Imports ImageFileRename
Imports System.Linq
Imports System.Threading.Tasks

Public Class frmImageFileRename
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents cmdBrowse As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents lblDirectory As System.Windows.Forms.Label
    Friend WithEvents lblFileLabel As System.Windows.Forms.Label
    Friend WithEvents lblFile As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents cmdNext As System.Windows.Forms.Button
    Friend WithEvents cmdUpdate As System.Windows.Forms.Button
    Friend WithEvents PicCurrent As System.Windows.Forms.PictureBox
    Friend WithEvents txtFileDate As System.Windows.Forms.TextBox
    Friend WithEvents txtFileNumber As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtExtension As System.Windows.Forms.TextBox
    Friend WithEvents cmdSaveAndNext As System.Windows.Forms.Button
    Friend WithEvents cmdDelete As System.Windows.Forms.Button
    Friend WithEvents cmdBack As System.Windows.Forms.Button
    Friend WithEvents chkAutoDate As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutoNumber As System.Windows.Forms.CheckBox
    Friend WithEvents chkAutoExtension As System.Windows.Forms.CheckBox
    Friend WithEvents txtDescription2 As System.Windows.Forms.TextBox
    Friend WithEvents OpenFileDialogBrowse As System.Windows.Forms.OpenFileDialog
    Friend WithEvents cblPerson As System.Windows.Forms.CheckedListBox
    Friend WithEvents txtPerson As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Private WithEvents ShapeContainer1 As PowerPacks.ShapeContainer
    Private WithEvents RectangleShape1 As PowerPacks.RectangleShape
    Private WithEvents RectangleShape2 As PowerPacks.RectangleShape
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents lnkAll As System.Windows.Forms.LinkLabel
    Friend WithEvents lnkNone As System.Windows.Forms.LinkLabel
    Friend WithEvents lblNewFileName As System.Windows.Forms.Label
    Friend WithEvents VidCurrent As AxWMPLib.AxWindowsMediaPlayer
    Friend WithEvents lnkDateFromFile As LinkLabel
    Friend WithEvents ContextMenuDateFromFile As ContextMenuStrip
    Friend WithEvents CreatedDate As ToolStripMenuItem
    Friend WithEvents ModifiedDate As ToolStripMenuItem
    Friend WithEvents cmbDescription As System.Windows.Forms.ComboBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImageFileRename))
        Me.cmdBrowse = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.lblDirectory = New System.Windows.Forms.Label()
        Me.lblFileLabel = New System.Windows.Forms.Label()
        Me.lblFile = New System.Windows.Forms.Label()
        Me.PicCurrent = New System.Windows.Forms.PictureBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.cmdUpdate = New System.Windows.Forms.Button()
        Me.OpenFileDialogBrowse = New System.Windows.Forms.OpenFileDialog()
        Me.txtFileDate = New System.Windows.Forms.TextBox()
        Me.txtFileNumber = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtExtension = New System.Windows.Forms.TextBox()
        Me.cmdSaveAndNext = New System.Windows.Forms.Button()
        Me.cmdDelete = New System.Windows.Forms.Button()
        Me.cmdBack = New System.Windows.Forms.Button()
        Me.chkAutoDate = New System.Windows.Forms.CheckBox()
        Me.chkAutoNumber = New System.Windows.Forms.CheckBox()
        Me.chkAutoExtension = New System.Windows.Forms.CheckBox()
        Me.txtDescription2 = New System.Windows.Forms.TextBox()
        Me.cmbDescription = New System.Windows.Forms.ComboBox()
        Me.cblPerson = New System.Windows.Forms.CheckedListBox()
        Me.txtPerson = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.ShapeContainer1 = New Microsoft.VisualBasic.PowerPacks.ShapeContainer()
        Me.RectangleShape2 = New Microsoft.VisualBasic.PowerPacks.RectangleShape()
        Me.RectangleShape1 = New Microsoft.VisualBasic.PowerPacks.RectangleShape()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.lnkAll = New System.Windows.Forms.LinkLabel()
        Me.lnkNone = New System.Windows.Forms.LinkLabel()
        Me.lblNewFileName = New System.Windows.Forms.Label()
        Me.VidCurrent = New AxWMPLib.AxWindowsMediaPlayer()
        Me.lnkDateFromFile = New System.Windows.Forms.LinkLabel()
        Me.ContextMenuDateFromFile = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CreatedDate = New System.Windows.Forms.ToolStripMenuItem()
        Me.ModifiedDate = New System.Windows.Forms.ToolStripMenuItem()
        CType(Me.PicCurrent, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VidCurrent, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuDateFromFile.SuspendLayout()
        Me.SuspendLayout()
        '
        'cmdBrowse
        '
        Me.cmdBrowse.Location = New System.Drawing.Point(19, 9)
        Me.cmdBrowse.Name = "cmdBrowse"
        Me.cmdBrowse.Size = New System.Drawing.Size(72, 24)
        Me.cmdBrowse.TabIndex = 2
        Me.cmdBrowse.Text = "Open File"
        '
        'Label2
        '
        Me.Label2.BackColor = System.Drawing.Color.Wheat
        Me.Label2.Location = New System.Drawing.Point(109, 9)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(56, 16)
        Me.Label2.TabIndex = 3
        Me.Label2.Text = "Directory:"
        '
        'lblDirectory
        '
        Me.lblDirectory.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblDirectory.BackColor = System.Drawing.Color.Wheat
        Me.lblDirectory.Location = New System.Drawing.Point(173, 9)
        Me.lblDirectory.Name = "lblDirectory"
        Me.lblDirectory.Size = New System.Drawing.Size(760, 16)
        Me.lblDirectory.TabIndex = 4
        '
        'lblFileLabel
        '
        Me.lblFileLabel.BackColor = System.Drawing.Color.Wheat
        Me.lblFileLabel.Location = New System.Drawing.Point(109, 33)
        Me.lblFileLabel.Name = "lblFileLabel"
        Me.lblFileLabel.Size = New System.Drawing.Size(56, 16)
        Me.lblFileLabel.TabIndex = 5
        Me.lblFileLabel.Text = "File:"
        '
        'lblFile
        '
        Me.lblFile.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblFile.BackColor = System.Drawing.Color.Wheat
        Me.lblFile.Location = New System.Drawing.Point(173, 33)
        Me.lblFile.Name = "lblFile"
        Me.lblFile.Size = New System.Drawing.Size(760, 16)
        Me.lblFile.TabIndex = 6
        '
        'PicCurrent
        '
        Me.PicCurrent.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PicCurrent.BackColor = System.Drawing.Color.White
        Me.PicCurrent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.PicCurrent.Location = New System.Drawing.Point(16, 69)
        Me.PicCurrent.Name = "PicCurrent"
        Me.PicCurrent.Size = New System.Drawing.Size(745, 493)
        Me.PicCurrent.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.PicCurrent.TabIndex = 7
        Me.PicCurrent.TabStop = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.Location = New System.Drawing.Point(16, 575)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 16)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Description:"
        '
        'Label3
        '
        Me.Label3.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label3.BackColor = System.Drawing.Color.Wheat
        Me.Label3.Location = New System.Drawing.Point(8, 638)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(96, 16)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "New File Name:"
        '
        'cmdNext
        '
        Me.cmdNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdNext.Location = New System.Drawing.Point(840, 662)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(80, 24)
        Me.cmdNext.TabIndex = 12
        Me.cmdNext.Text = "Next >"
        '
        'cmdUpdate
        '
        Me.cmdUpdate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdUpdate.Location = New System.Drawing.Point(752, 662)
        Me.cmdUpdate.Name = "cmdUpdate"
        Me.cmdUpdate.Size = New System.Drawing.Size(80, 24)
        Me.cmdUpdate.TabIndex = 13
        Me.cmdUpdate.Text = "Save"
        '
        'txtFileDate
        '
        Me.txtFileDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileDate.Location = New System.Drawing.Point(788, 558)
        Me.txtFileDate.Name = "txtFileDate"
        Me.txtFileDate.Size = New System.Drawing.Size(80, 20)
        Me.txtFileDate.TabIndex = 14
        '
        'txtFileNumber
        '
        Me.txtFileNumber.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFileNumber.Location = New System.Drawing.Point(788, 510)
        Me.txtFileNumber.Name = "txtFileNumber"
        Me.txtFileNumber.Size = New System.Drawing.Size(80, 20)
        Me.txtFileNumber.TabIndex = 15
        '
        'Label4
        '
        Me.Label4.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label4.Location = New System.Drawing.Point(788, 542)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(64, 16)
        Me.Label4.TabIndex = 16
        Me.Label4.Text = "Date:"
        '
        'Label5
        '
        Me.Label5.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label5.Location = New System.Drawing.Point(788, 494)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(64, 16)
        Me.Label5.TabIndex = 17
        Me.Label5.Text = "Number:"
        '
        'Label6
        '
        Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label6.Location = New System.Drawing.Point(788, 582)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(80, 16)
        Me.Label6.TabIndex = 18
        Me.Label6.Text = "Extension:"
        '
        'txtExtension
        '
        Me.txtExtension.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtExtension.Location = New System.Drawing.Point(788, 598)
        Me.txtExtension.Name = "txtExtension"
        Me.txtExtension.Size = New System.Drawing.Size(48, 20)
        Me.txtExtension.TabIndex = 19
        '
        'cmdSaveAndNext
        '
        Me.cmdSaveAndNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmdSaveAndNext.Location = New System.Drawing.Point(608, 662)
        Me.cmdSaveAndNext.Name = "cmdSaveAndNext"
        Me.cmdSaveAndNext.Size = New System.Drawing.Size(128, 24)
        Me.cmdSaveAndNext.TabIndex = 20
        Me.cmdSaveAndNext.Text = "Save + Next >"
        '
        'cmdDelete
        '
        Me.cmdDelete.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdDelete.Location = New System.Drawing.Point(8, 662)
        Me.cmdDelete.Name = "cmdDelete"
        Me.cmdDelete.Size = New System.Drawing.Size(96, 24)
        Me.cmdDelete.TabIndex = 21
        Me.cmdDelete.Text = "Delete"
        '
        'cmdBack
        '
        Me.cmdBack.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.cmdBack.Location = New System.Drawing.Point(112, 662)
        Me.cmdBack.Name = "cmdBack"
        Me.cmdBack.Size = New System.Drawing.Size(88, 24)
        Me.cmdBack.TabIndex = 22
        Me.cmdBack.Text = "< Back"
        '
        'chkAutoDate
        '
        Me.chkAutoDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAutoDate.Checked = True
        Me.chkAutoDate.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAutoDate.Location = New System.Drawing.Point(772, 558)
        Me.chkAutoDate.Name = "chkAutoDate"
        Me.chkAutoDate.Size = New System.Drawing.Size(16, 16)
        Me.chkAutoDate.TabIndex = 23
        '
        'chkAutoNumber
        '
        Me.chkAutoNumber.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAutoNumber.Checked = True
        Me.chkAutoNumber.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAutoNumber.Location = New System.Drawing.Point(772, 502)
        Me.chkAutoNumber.Name = "chkAutoNumber"
        Me.chkAutoNumber.Size = New System.Drawing.Size(88, 37)
        Me.chkAutoNumber.TabIndex = 24
        '
        'chkAutoExtension
        '
        Me.chkAutoExtension.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkAutoExtension.Checked = True
        Me.chkAutoExtension.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAutoExtension.Location = New System.Drawing.Point(772, 598)
        Me.chkAutoExtension.Name = "chkAutoExtension"
        Me.chkAutoExtension.Size = New System.Drawing.Size(16, 16)
        Me.chkAutoExtension.TabIndex = 25
        '
        'txtDescription2
        '
        Me.txtDescription2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtDescription2.Location = New System.Drawing.Point(83, 599)
        Me.txtDescription2.Name = "txtDescription2"
        Me.txtDescription2.Size = New System.Drawing.Size(678, 20)
        Me.txtDescription2.TabIndex = 26
        '
        'cmbDescription
        '
        Me.cmbDescription.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cmbDescription.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.Suggest
        Me.cmbDescription.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.cmbDescription.Location = New System.Drawing.Point(83, 572)
        Me.cmbDescription.Name = "cmbDescription"
        Me.cmbDescription.Size = New System.Drawing.Size(678, 21)
        Me.cmbDescription.TabIndex = 28
        '
        'cblPerson
        '
        Me.cblPerson.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.cblPerson.CheckOnClick = True
        Me.cblPerson.FormattingEnabled = True
        Me.cblPerson.Location = New System.Drawing.Point(788, 141)
        Me.cblPerson.Name = "cblPerson"
        Me.cblPerson.Size = New System.Drawing.Size(132, 334)
        Me.cblPerson.TabIndex = 29
        '
        'txtPerson
        '
        Me.txtPerson.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtPerson.Location = New System.Drawing.Point(814, 85)
        Me.txtPerson.Name = "txtPerson"
        Me.txtPerson.Size = New System.Drawing.Size(106, 20)
        Me.txtPerson.TabIndex = 30
        '
        'Label7
        '
        Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(788, 69)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(43, 13)
        Me.Label7.TabIndex = 31
        Me.Label7.Text = "People:"
        '
        'ShapeContainer1
        '
        Me.ShapeContainer1.Location = New System.Drawing.Point(0, 0)
        Me.ShapeContainer1.Margin = New System.Windows.Forms.Padding(0)
        Me.ShapeContainer1.Name = "ShapeContainer1"
        Me.ShapeContainer1.Shapes.AddRange(New Microsoft.VisualBasic.PowerPacks.Shape() {Me.RectangleShape2, Me.RectangleShape1})
        Me.ShapeContainer1.Size = New System.Drawing.Size(936, 692)
        Me.ShapeContainer1.TabIndex = 32
        Me.ShapeContainer1.TabStop = False
        '
        'RectangleShape2
        '
        Me.RectangleShape2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RectangleShape2.BackColor = System.Drawing.Color.Wheat
        Me.RectangleShape2.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Opaque
        Me.RectangleShape2.BorderColor = System.Drawing.Color.Goldenrod
        Me.RectangleShape2.Location = New System.Drawing.Point(-13, 627)
        Me.RectangleShape2.Name = "RectangleShape2"
        Me.RectangleShape2.Size = New System.Drawing.Size(961, 78)
        '
        'RectangleShape1
        '
        Me.RectangleShape1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.RectangleShape1.BackColor = System.Drawing.Color.Wheat
        Me.RectangleShape1.BackStyle = Microsoft.VisualBasic.PowerPacks.BackStyle.Opaque
        Me.RectangleShape1.BorderColor = System.Drawing.Color.Goldenrod
        Me.RectangleShape1.Location = New System.Drawing.Point(-2, -3)
        Me.RectangleShape1.Name = "RectangleShape1"
        Me.RectangleShape1.Size = New System.Drawing.Size(951, 55)
        '
        'Label8
        '
        Me.Label8.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(16, 602)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(51, 13)
        Me.Label8.TabIndex = 33
        Me.Label8.Text = "Location:"
        '
        'Label9
        '
        Me.Label9.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(788, 91)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(29, 13)
        Me.Label9.TabIndex = 34
        Me.Label9.Text = "Add:"
        '
        'lnkAll
        '
        Me.lnkAll.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lnkAll.AutoSize = True
        Me.lnkAll.Location = New System.Drawing.Point(788, 122)
        Me.lnkAll.Name = "lnkAll"
        Me.lnkAll.Size = New System.Drawing.Size(51, 13)
        Me.lnkAll.TabIndex = 35
        Me.lnkAll.TabStop = True
        Me.lnkAll.Text = "Select All"
        '
        'lnkNone
        '
        Me.lnkNone.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lnkNone.AutoSize = True
        Me.lnkNone.Location = New System.Drawing.Point(860, 122)
        Me.lnkNone.Name = "lnkNone"
        Me.lnkNone.Size = New System.Drawing.Size(66, 13)
        Me.lnkNone.TabIndex = 36
        Me.lnkNone.TabStop = True
        Me.lnkNone.Text = "Select None"
        '
        'lblNewFileName
        '
        Me.lblNewFileName.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.lblNewFileName.AutoSize = True
        Me.lblNewFileName.BackColor = System.Drawing.Color.Wheat
        Me.lblNewFileName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNewFileName.Location = New System.Drawing.Point(92, 633)
        Me.lblNewFileName.Name = "lblNewFileName"
        Me.lblNewFileName.Size = New System.Drawing.Size(0, 20)
        Me.lblNewFileName.TabIndex = 37
        '
        'VidCurrent
        '
        Me.VidCurrent.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.VidCurrent.Enabled = True
        Me.VidCurrent.Location = New System.Drawing.Point(16, 69)
        Me.VidCurrent.Name = "VidCurrent"
        Me.VidCurrent.OcxState = CType(resources.GetObject("VidCurrent.OcxState"), System.Windows.Forms.AxHost.State)
        Me.VidCurrent.Size = New System.Drawing.Size(745, 493)
        Me.VidCurrent.TabIndex = 38
        Me.VidCurrent.Visible = False
        '
        'lnkDateFromFile
        '
        Me.lnkDateFromFile.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lnkDateFromFile.AutoSize = True
        Me.lnkDateFromFile.ContextMenuStrip = Me.ContextMenuDateFromFile
        Me.lnkDateFromFile.Location = New System.Drawing.Point(875, 560)
        Me.lnkDateFromFile.Name = "lnkDateFromFile"
        Me.lnkDateFromFile.Size = New System.Drawing.Size(43, 13)
        Me.lnkDateFromFile.TabIndex = 39
        Me.lnkDateFromFile.TabStop = True
        Me.lnkDateFromFile.Text = "from file"
        '
        'ContextMenuDateFromFile
        '
        Me.ContextMenuDateFromFile.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CreatedDate, Me.ModifiedDate})
        Me.ContextMenuDateFromFile.Name = "ContextMenuDateFromFile"
        Me.ContextMenuDateFromFile.Size = New System.Drawing.Size(150, 48)
        Me.ContextMenuDateFromFile.Text = "Created Date"
        '
        'CreatedDate
        '
        Me.CreatedDate.Name = "CreatedDate"
        Me.CreatedDate.Size = New System.Drawing.Size(149, 22)
        Me.CreatedDate.Text = "Created Date"
        '
        'ModifiedDate
        '
        Me.ModifiedDate.Name = "ModifiedDate"
        Me.ModifiedDate.Size = New System.Drawing.Size(149, 22)
        Me.ModifiedDate.Text = "Modified Date"
        '
        'frmImageFileRename
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(936, 692)
        Me.Controls.Add(Me.lnkDateFromFile)
        Me.Controls.Add(Me.VidCurrent)
        Me.Controls.Add(Me.lblNewFileName)
        Me.Controls.Add(Me.lnkNone)
        Me.Controls.Add(Me.lnkAll)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.txtPerson)
        Me.Controls.Add(Me.cblPerson)
        Me.Controls.Add(Me.cmbDescription)
        Me.Controls.Add(Me.txtDescription2)
        Me.Controls.Add(Me.chkAutoExtension)
        Me.Controls.Add(Me.chkAutoDate)
        Me.Controls.Add(Me.cmdBack)
        Me.Controls.Add(Me.cmdDelete)
        Me.Controls.Add(Me.cmdSaveAndNext)
        Me.Controls.Add(Me.txtExtension)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.txtFileNumber)
        Me.Controls.Add(Me.txtFileDate)
        Me.Controls.Add(Me.cmdUpdate)
        Me.Controls.Add(Me.cmdNext)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.PicCurrent)
        Me.Controls.Add(Me.lblFile)
        Me.Controls.Add(Me.lblFileLabel)
        Me.Controls.Add(Me.lblDirectory)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.cmdBrowse)
        Me.Controls.Add(Me.chkAutoNumber)
        Me.Controls.Add(Me.ShapeContainer1)
        Me.KeyPreview = True
        Me.Name = "frmImageFileRename"
        Me.Text = "Image File Rename 2.0"
        CType(Me.PicCurrent, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VidCurrent, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuDateFromFile.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    'Private c_strCurrent As String
    'Private c_strNext As String
    'Private c_strPrevious As String
    Private c_blnSaved As Boolean
    Private c_PhotoProperties As JSG.PhotoPropertiesLibrary.PhotoProperties
    Private c_strDescription As ArrayList
    Private c_strPerson As ArrayList
    Private c_blnLoading As Boolean
    Private c_strCurrentDescription As String
    Private c_objFileHistory As FileHistory

    Private Sub cmdBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBrowse.Click
        ShowOpenFileDialog()
    End Sub

    Private Sub ShowOpenFileDialog()
        Me.Cursor = Cursors.WaitCursor
        If Me.OpenFileDialogBrowse.ShowDialog() = DialogResult.OK Then
            c_objFileHistory.AddFile(Me.OpenFileDialogBrowse.FileName)
            LoadFile(c_objFileHistory.CurrentFile)
        End If
        Me.Cursor = Cursors.Default
    End Sub
    Private Function GetFileNameStyle(ByVal strFileName As String) As ImageFileRename.FileNameStyle

        Select Case True
            Case ImageFileRename.Filename_Canon.FileNameConforms(strFileName)
                Return New ImageFileRename.Filename_Canon
            Case ImageFileRename.Filename_Canon_Video.FileNameConforms(strFileName)
                Return New ImageFileRename.Filename_Canon_Video
            Case ImageFileRename.FileName_iPhone.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_iPhone
            Case ImageFileRename.FileName_Fuji.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Fuji
            Case ImageFileRename.FileName_Sony.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Sony
            Case ImageFileRename.FileName_Minolta.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Minolta
            Case ImageFileRename.FileName_Konika.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Konika
            Case ImageFileRename.Filename_Olympus.FileNameConforms(strFileName)
                Return New ImageFileRename.Filename_Olympus
            Case ImageFileRename.FileName_Date_Time_RandomNumber.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Date_Time_RandomNumber
            Case ImageFileRename.FileName_Film_Frame.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Film_Frame
            Case ImageFileRename.FileName_Film_FrameDescr.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Film_FrameDescr
            Case ImageFileRename.FileName_Film_Frame_Descr.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Film_Frame_Descr
            Case ImageFileRename.FileName_Date_Film_Frame_Descr.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Date_Film_Frame_Descr
            Case ImageFileRename.FileName_Date_ImageNumber_Descr.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_Date_ImageNumber_Descr
            Case ImageFileRename.Filename_Nokia.FileNameConforms(strFileName)
                Return New ImageFileRename.Filename_Nokia
            Case ImageFileRename.FileName_ImageNumber.FileNameConforms(strFileName)
                Return New ImageFileRename.FileName_ImageNumber
            Case Else
                Return New ImageFileRename.FileNameStyle
        End Select

    End Function
    Private Sub LoadFile(ByVal strFile As String)
        Dim strExtension As String
        Dim objFileNameStyle As ImageFileRename.FileNameStyle
        Dim dtePictureTaken As Date
        Dim strDescription As String

        c_blnLoading = True

        strExtension = modCommonFunctions.Extension(strFile)



        Me.lblDirectory.Text = FilePath(strFile)
        Me.lblFile.Text = FileName(strFile)


        objFileNameStyle = GetFileNameStyle(strFile)

        If Me.chkAutoNumber.Checked Then
            Me.txtFileNumber.Text = objFileNameStyle.FileNumber(strFile)
        End If

        strDescription = objFileNameStyle.FileDescription(strFile)
        If strDescription > "" Then
            If txtDescription2.Text > "" Then
                If Strings.Right(strDescription, Len(txtDescription2.Text)) = txtDescription2.Text Then
                    strDescription = Strings.Left(strDescription, Len(strDescription) - (Len(txtDescription2.Text)))
                    strDescription.TrimEnd(" ")
                    strDescription.TrimEnd(",")
                End If
            End If
            Me.cmbDescription.Text = strDescription
        End If

        If Me.chkAutoDate.Checked Then

            dtePictureTaken = objFileNameStyle.FileDate(strFile)
            If dtePictureTaken = Date.MinValue Then
                dtePictureTaken = DatePictureTaken(strFile)
            End If
            If dtePictureTaken = Date.MinValue Then
                Me.txtFileDate.Text = ""
            Else
                Me.txtFileDate.Text = dtePictureTaken.ToString("yyyyMMdd")
            End If
        End If

        If Me.chkAutoExtension.Checked Then
            Me.txtExtension.Text = objFileNameStyle.Extension(strFile)
        End If

        MetadataChanged()
        Me.cmbDescription.Focus()

        Select Case strExtension.ToLower()
            Case "jpg"
                LoadImage(strFile)
                Me.PicCurrent.Visible = True
                Me.VidCurrent.Visible = False
            Case "avi", "mov"
                LoadVideo(strFile)
                Me.VidCurrent.Visible = True
                Me.PicCurrent.Visible = False
        End Select

        c_blnSaved = False
        c_blnLoading = False

    End Sub

    Private Sub LoadVideo(ByVal strFile As String)

        Me.VidCurrent.URL = strFile

    End Sub
    Private Sub LoadImage(ByVal strFile As String)
        Dim bmpImage As Bitmap

        Me.VidCurrent.URL = ""

        Using fsImage As FileStream = New FileStream(strFile, FileMode.Open)
            bmpImage = New Bitmap(fsImage)
            fsImage.Close()
        End Using

        Me.PicCurrent.Image = bmpImage 'imgCurrent



    End Sub

    Private Enum FileDateType As Integer
        Created = 1
        Modified = 2
    End Enum

    Private Function FileDate(ByVal strString As String, enFileDateType As FileDateType) As String
        Dim filePicture As IO.FileInfo
        Dim strReturn As String
        Dim dteFileDate As Date

        filePicture = New IO.FileInfo(strString)
        Select Case enFileDateType
            Case FileDateType.Created
                dteFileDate = filePicture.CreationTime
            Case FileDateType.Modified
                dteFileDate = filePicture.LastWriteTime
        End Select

        strReturn = Format(dteFileDate, "yyyyMMdd")

        filePicture = Nothing

        Return strReturn

    End Function

    Private Function CalculateFullDescription() As String
        Dim strReturn As String = ""
        Dim strDescription As String = ""
        Dim strPeople As String = ""
        Dim strLeadCharacter As String = ""

        strPeople = BuildPeopleString()
        If strPeople > "" Then
            strReturn = strPeople
        End If

        If Me.cmbDescription.Text > "" Then
            strDescription = Me.cmbDescription.Text
            'If there are people's names, append a space
            If strReturn > "" Then
                strReturn &= " "
            Else 'otherwise capitalise the first letter of the description
                strLeadCharacter = strDescription.ToCharArray(0, 1)
                strLeadCharacter = strLeadCharacter.ToUpper
                strDescription = strDescription.Substring(1)
                strDescription = strLeadCharacter & strDescription
            End If
            strReturn &= strDescription
        End If

        If Me.txtDescription2.Text > "" Then
            If strReturn > "" Then
                strReturn &= ", "
            End If
            strReturn &= Me.txtDescription2.Text
        End If

        Return strReturn

    End Function
    Private Sub MetadataChanged()

        WarningsForSave()
        CalculateNewFileName()

    End Sub

    Private Sub CalculateNewFileName()

        Me.lblNewFileName.Text = Me.txtFileDate.Text

        If Me.txtFileNumber.Text > "" Then
            If Me.lblNewFileName.Text > "" Then
                Me.lblNewFileName.Text &= "_"
            End If
            Me.lblNewFileName.Text &= Me.txtFileNumber.Text
        End If

        If CalculateFullDescription() > "" Then
            If Me.lblNewFileName.Text > "" Then
                'dad
                'Me.txtNewFileName.Text &= " "
                Me.lblNewFileName.Text &= "_"
            End If
            Me.lblNewFileName.Text &= CalculateFullDescription()
        End If

        Me.lblNewFileName.Text &= "." & Me.txtExtension.Text

    End Sub

    Private Function BuildPeopleString() As String
        Dim strReturn As String = ""
        Dim intPerson As Integer

        For intPerson = 0 To cblPerson.CheckedItems.Count - 1
            If strReturn > "" Then
                If intPerson = cblPerson.CheckedItems.Count - 1 Then
                    strReturn &= " & "
                Else
                    strReturn &= ", "
                End If

            End If
            strReturn &= cblPerson.CheckedItems(intPerson)
        Next

        Return strReturn

    End Function

    Private Sub txtFileDate_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFileDate.TextChanged
        MetadataChanged()
    End Sub

    Private Sub txtFileNumber_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtFileNumber.TextChanged
        MetadataChanged()
    End Sub

    Private Sub cmbDescription_MouseEnter(sender As Object, e As System.EventArgs) Handles cmbDescription.MouseEnter
        cmbDescription.Focus()
    End Sub

    Private Sub cmbDescription_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbDescription.TextChanged

        MetadataChanged()

    End Sub

    Private Async Sub cmdUpdate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdUpdate.Click

        Await SaveFile()

    End Sub

    Private Async Function SaveFile() As Task(Of Boolean)
        Dim fileImage As System.IO.FileInfo
        Dim alstWarnings As ArrayList
        Dim strNewFileName As String
        Dim blnSave As Boolean = True
        Dim blnSuccess As Boolean = False

        Cursor = Cursors.WaitCursor

        alstWarnings = WarningsForSave()

        If alstWarnings.Count > 0 Then
            If MsgBox(alstWarnings.Count & " fields have no values. Are you sure you want to Save?", vbOKCancel, "Blank Fields") = MsgBoxResult.Cancel Then
                blnSave = False
            End If
        End If
        If blnSave Then
            Try
                strNewFileName = Me.lblDirectory.Text & "\" & Me.lblNewFileName.Text

                Using SourceStream As FileStream = File.Open(c_objFileHistory.CurrentFile, FileMode.Open)
                    Using DestinationStream As FileStream = File.Create(strNewFileName)
                        Await SourceStream.CopyToAsync(DestinationStream)
                    End Using

                End Using

                fileImage = New FileInfo(c_objFileHistory.CurrentFile)
                fileImage.Delete()

                c_objFileHistory.UpdateCurrentFile(strNewFileName)

                c_blnSaved = True

                If Not InDescriptionList(Me.cmbDescription.Text) Then
                    AddToDescriptionList(Me.cmbDescription.Text)
                End If

                blnSuccess = True
            Catch ex As Exception
                MsgBox(ex.Message)
                blnSuccess = False
            End Try
        End If

        Cursor = Cursors.Default

        Return blnSuccess

    End Function

    Private Function WarningsForSave() As ArrayList
        Dim alstReturn As New ArrayList

        cmbDescription.BackColor = Color.White
        txtDescription2.BackColor = Color.White
        txtFileNumber.BackColor = Color.White
        txtFileDate.BackColor = Color.White
        txtExtension.BackColor = Color.White

        If Me.cmbDescription.Text = "" Then
            cmbDescription.BackColor = Color.IndianRed
            alstReturn.Add("No Description")
        End If

        If Me.txtDescription2.Text = "" Then
            alstReturn.Add("No Location")
            txtDescription2.BackColor = Color.IndianRed
        End If

        If txtFileNumber.Text = "" Then
            alstReturn.Add("No Number")
            txtFileNumber.BackColor = Color.IndianRed
        End If

        If txtFileDate.Text = "" Then
            alstReturn.Add("No Date")
            txtFileDate.BackColor = Color.IndianRed
        End If

        If txtExtension.Text = "" Then
            alstReturn.Add("No File Extension")
            txtExtension.BackColor = Color.IndianRed
        End If

        Return alstReturn

    End Function

    Private Function SupportedFiles(strDirectory As String) As Generic.IEnumerable(Of String)
        Return Directory.EnumerateFiles(strDirectory, "*.*", SearchOption.TopDirectoryOnly).Where(SupportedFileTypesFunction)
    End Function

    Private Function SupportedFileTypesFunction() As System.Func(Of String, Boolean)
        Return Function(strFile As String) FileTypeSupported(strFile)
    End Function

    Private Function FileTypeSupported(strFileName As String) As Boolean
        Dim strExtension As String

        strExtension = modCommonFunctions.Extension(strFileName)

        FileTypeSupported = True

        Select Case strExtension.ToLower
            Case "jpg"
            Case "mov"
            Case "avi"
            Case Else
                Return False
        End Select

    End Function

    Private Function NextFileName(ByVal strCurrentFile As String) As String
        Dim files As Generic.IEnumerable(Of String)
        Dim strFile As String
        Dim intFile As Int16
        Dim strReturn As String = ""

        files = SupportedFiles(FilePath(strCurrentFile))
        For intFile = 0 To files.Count - 1
            strFile = files(intFile)
            If strFile = strCurrentFile Then
                If intFile < files.Count - 1 Then
                    strReturn = files(intFile + 1)
                Else
                    strReturn = files(0)
                End If
                Exit For
            End If
        Next

        Return strReturn

    End Function

    Private Function PreviousFileName(ByVal strCurrentFile As String) As String
        Dim files As Generic.IEnumerable(Of String)
        Dim strFile As String
        Dim intFile As Int16
        Dim strReturn As String = ""

        files = SupportedFiles(FilePath(strCurrentFile))
        For intFile = 0 To files.Count - 1
            strFile = files(intFile)
            If strFile = strCurrentFile Then
                If intFile > 0 Then
                    strReturn = files(intFile - 1)
                Else
                    strReturn = files(files.Count - 1)
                End If
                Exit For
            End If
        Next

        Return strReturn

    End Function

    Private Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click

        Cursor = Cursors.WaitCursor

        NextFile()

        Cursor = Cursors.Default

    End Sub

    Private Sub cmdSaveAndNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdSaveAndNext.Click

        SaveAndNext()

    End Sub

    Private Async Sub SaveAndNext()
        Cursor = Cursors.WaitCursor

        If Await SaveFile() Then

            NextFile()

        End If

        Cursor = Cursors.Default
    End Sub

    Private Sub NextFile()

        Dim strNextFileName As String

        If c_objFileHistory.AtLastFile Then
            strNextFileName = NextFileName(c_objFileHistory.CurrentFile)
            c_objFileHistory.AddFile(strNextFileName)
        Else
            strNextFileName = c_objFileHistory.MoveForward
        End If

        LoadFile(c_objFileHistory.CurrentFile)

    End Sub

    Private Sub cmdDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDelete.Click
        Dim intPress As Integer

        intPress = MsgBox("Are you sure?", MsgBoxStyle.OkCancel, "Delete?")
        Select Case intPress
            Case vbOK
                Cursor = Cursors.WaitCursor
                DeleteFile(c_objFileHistory.CurrentFile)
                c_objFileHistory.RemoveCurrentFile()
                NextFile()
                Cursor = Cursors.Default
        End Select
    End Sub

    Private Sub DeleteFile(ByVal strPath As String)
        Dim file As IO.FileInfo

        file = New IO.FileInfo(strPath)
        If file.Exists Then
            file.Delete()
        End If
    End Sub

    Private Sub cmdBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdBack.Click

        Me.Cursor = Cursors.WaitCursor

        If Not c_objFileHistory.AtFirstFile Then
            c_objFileHistory.MoveBack()
            LoadFile(c_objFileHistory.CurrentFile)
        End If

        Me.Cursor = Cursors.Default

    End Sub

    Private Sub txtDescription2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtDescription2.TextChanged
        MetadataChanged()
    End Sub

    Private Function DatePictureTaken(ByVal strFilePath As String) As Date
        Dim tagdatum As JSG.PhotoPropertiesLibrary.PhotoTagDatum
        Dim dteTaken As Date

        dteTaken = Date.MinValue
        Try
            c_PhotoProperties.Analyze(strFilePath)
            tagdatum = c_PhotoProperties.GetTagDatum(36867) 'Date Taken
            If Not tagdatum Is Nothing Then
                dteTaken = Date.ParseExact(tagdatum.Value, "yyyy:MM:dd HH:mm:ss", Nothing)
            End If

        Catch ex As Exception
        End Try

        Return dteTaken

    End Function

    Private Async Sub frmImageFileRename_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.Control = True Then
            Select Case e.KeyCode
                Case Keys.O
                    ShowOpenFileDialog()
                Case Keys.S
                    Await SaveFile()
                Case Keys.Enter
                    SaveAndNext()
                Case Keys.P
                    Me.cblPerson.Focus()
                Case Keys.D
                    Me.cmbDescription.Focus()
                Case Keys.L
                    Me.txtDescription2.Focus()
                Case Keys.A
                    Me.txtPerson.Focus()
            End Select
        End If

    End Sub

    Private Sub frmImageFileRename_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim objAssembly As System.Reflection.Assembly

        objAssembly = System.Reflection.Assembly.GetExecutingAssembly
        Me.Text = objAssembly.GetName.Name & " " & objAssembly.GetName.Version.Major & "." & objAssembly.GetName.Version.Minor

        c_objFileHistory = New FileHistory
        c_PhotoProperties = New JSG.PhotoPropertiesLibrary.PhotoProperties
        c_PhotoProperties.Initialize()
        c_strDescription = New ArrayList
        c_strDescription.Add("")

        c_strPerson = DefaultPersonList()
        Me.cblPerson.DataSource = c_strPerson.ToArray

        Me.cmbDescription.DataSource = c_strDescription

    End Sub

    Private Function InDescriptionList(ByVal strDescription As String) As Boolean
        Return c_strDescription.IndexOf(strDescription) > -1
    End Function

    Private Sub AddToDescriptionList(ByVal strDescription As String)
        Dim strText As String

        strText = Me.cmbDescription.Text

        c_strDescription.Insert(1, strText)
        Me.cmbDescription.DataSource = c_strDescription.ToArray

        Me.cmbDescription.Text = strText

    End Sub

    Private Sub txtPerson_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtPerson.KeyPress
        Dim strText As String

        If e.KeyChar = CChar(vbCrLf) Then
            strText = Me.txtPerson.Text

            c_strPerson.Add(strText)
            Me.cblPerson.DataSource = c_strPerson.ToArray
            Me.cblPerson.SetItemCheckState(cblPerson.Items.Count - 1, CheckState.Checked)

            txtPerson.Text = ""

            MetadataChanged()
        End If

    End Sub

    Private Sub cblPerson_ItemCheck(sender As Object, e As System.Windows.Forms.ItemCheckEventArgs) Handles cblPerson.ItemCheck

        Static blnUpdating As Boolean

        If Not blnUpdating Then
            blnUpdating = True
            cblPerson.SetItemCheckState(e.Index, e.NewValue)
            blnUpdating = False
            MetadataChanged()
        End If

    End Sub

    Private Sub lnkAll_LinkClicked(sender As System.Object, e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles lnkAll.LinkClicked
        CheckAllPeople(CheckState.Checked)
    End Sub

    Private Sub CheckAllPeople(intChecked As Windows.Forms.CheckState)
        Dim intPerson As Integer

        For intPerson = 0 To cblPerson.Items.Count - 1
            cblPerson.SetItemCheckState(intPerson, intChecked)
        Next

    End Sub

    Private Sub lnkNone_Click(sender As Object, e As System.EventArgs) Handles lnkNone.Click
        CheckAllPeople(CheckState.Unchecked)
    End Sub

    Private Function DefaultPersonList() As ArrayList
        Dim alstReturn As ArrayList

        alstReturn = New ArrayList

        With alstReturn
            .Add("Toby")
            .Add("Meris")
            .Add("Kirsten")
            .Add("Matthew")
            .Add("Sara")
            .Add("Steve")
            .Add("Frances")
            .Add("Arthur")
            .Add("Sally")
            .Add("Gary")
            .Add("Sarah")
        End With

        Return alstReturn

    End Function

    Private Sub cblPerson_MouseEnter(sender As Object, e As System.EventArgs) Handles cblPerson.MouseEnter
        cblPerson.Focus()
    End Sub

    Private Sub frmImageFileRename_PreviewKeyDown(sender As Object, e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles Me.PreviewKeyDown
    End Sub

    Private Sub lnkDateFromFile_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles lnkDateFromFile.LinkClicked

        lnkDateFromFile.ContextMenuStrip.Show(Cursor.Position)

    End Sub

    Private Sub ContextMenuDateFromFile_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles ContextMenuDateFromFile.ItemClicked
        Select Case e.ClickedItem.Name
            Case "CreatedDate"
                txtFileDate.Text = FileDate(c_objFileHistory.CurrentFile, FileDateType.Created)
            Case "ModifiedDate"
                txtFileDate.Text = FileDate(c_objFileHistory.CurrentFile, FileDateType.Modified)
        End Select
    End Sub
End Class

Public Class FileHistory
    Dim c_strFiles As ArrayList
    Dim c_intCurrentFile As Integer

    Public Sub New()
        c_strFiles = New ArrayList
        c_intCurrentFile = -1
    End Sub

    Public Sub AddFile(strPath As String)
        c_strFiles.Add(strPath)
        c_intCurrentFile = c_strFiles.Count - 1
    End Sub

    Public Function CurrentFile() As String
        Return c_strFiles(c_intCurrentFile)
    End Function

    Public Function MoveBack() As String
        c_intCurrentFile -= 1
        Return CurrentFile()
    End Function

    Public Function MoveForward() As String
        c_intCurrentFile += 1
        Return CurrentFile()
    End Function

    Public Function AtLastFile() As Boolean
        Return c_intCurrentFile = c_strFiles.Count - 1
    End Function

    Public Function AtFirstFile() As Boolean
        Return c_intCurrentFile = 0
    End Function

    Public Function FileCount() As Integer
        Return c_strFiles.Count
    End Function

    Public Sub UpdateCurrentFile(strNewPath As String)
        c_strFiles(c_intCurrentFile) = strNewPath
    End Sub

    Public Sub RemoveCurrentFile()
        c_strFiles.RemoveAt(c_intCurrentFile)
        If c_intCurrentFile > c_strFiles.Count - 1 Then
            c_intCurrentFile = c_strFiles.Count - 1
        End If
    End Sub
End Class