<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class BTA2OutlookApplication
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
		Dim ListViewGroup1 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Flights", System.Windows.Forms.HorizontalAlignment.Left)
		Dim ListViewGroup2 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Hotels", System.Windows.Forms.HorizontalAlignment.Left)
		Dim ListViewGroup3 As System.Windows.Forms.ListViewGroup = New System.Windows.Forms.ListViewGroup("Rental cars", System.Windows.Forms.HorizontalAlignment.Left)
		Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(BTA2OutlookApplication))
		Me.lst = New System.Windows.Forms.ListView()
		Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
		Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
		Me.ColumnHeader3 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
		Me.ColumnHeader4 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
		Me.ImageList1 = New System.Windows.Forms.ImageList(Me.components)
		Me.btnCreateEventsInOutlook = New System.Windows.Forms.Button()
		Me.chkAddBoardingblocker = New System.Windows.Forms.CheckBox()
		Me.lblWait = New System.Windows.Forms.Label()
		Me.Panel1 = New System.Windows.Forms.Panel()
		Me.btnSaveICSs = New System.Windows.Forms.Button()
		Me.Panel1.SuspendLayout()
		Me.SuspendLayout()
		'
		'lst
		'
		Me.lst.CheckBoxes = True
		Me.lst.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2, Me.ColumnHeader3, Me.ColumnHeader4})
		Me.lst.Dock = System.Windows.Forms.DockStyle.Fill
		Me.lst.Font = New System.Drawing.Font("Calibri", 11.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lst.FullRowSelect = True
		ListViewGroup1.Header = "Flights"
		ListViewGroup1.Name = Nothing
		ListViewGroup2.Header = "Hotels"
		ListViewGroup2.Name = Nothing
		ListViewGroup3.Header = "Rental cars"
		ListViewGroup3.Name = "cars"
		Me.lst.Groups.AddRange(New System.Windows.Forms.ListViewGroup() {ListViewGroup1, ListViewGroup2, ListViewGroup3})
		Me.lst.HideSelection = False
		Me.lst.LargeImageList = Me.ImageList1
		Me.lst.Location = New System.Drawing.Point(0, 0)
		Me.lst.Name = "lst"
		Me.lst.Size = New System.Drawing.Size(1598, 560)
		Me.lst.SmallImageList = Me.ImageList1
		Me.lst.TabIndex = 2
		Me.lst.UseCompatibleStateImageBehavior = False
		Me.lst.View = System.Windows.Forms.View.Details
		'
		'ColumnHeader1
		'
		Me.ColumnHeader1.Text = "Start"
		'
		'ColumnHeader2
		'
		Me.ColumnHeader2.Text = "End"
		'
		'ColumnHeader3
		'
		Me.ColumnHeader3.Text = ""
		'
		'ColumnHeader4
		'
		Me.ColumnHeader4.Text = "Duration"
		'
		'ImageList1
		'
		Me.ImageList1.ImageStream = CType(resources.GetObject("ImageList1.ImageStream"), System.Windows.Forms.ImageListStreamer)
		Me.ImageList1.TransparentColor = System.Drawing.Color.Transparent
		Me.ImageList1.Images.SetKeyName(0, "")
		Me.ImageList1.Images.SetKeyName(1, "")
		Me.ImageList1.Images.SetKeyName(2, "")
		'
		'btnCreateEventsInOutlook
		'
		Me.btnCreateEventsInOutlook.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
			Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
		Me.btnCreateEventsInOutlook.Image = CType(resources.GetObject("btnCreateEventsInOutlook.Image"), System.Drawing.Image)
		Me.btnCreateEventsInOutlook.ImageAlign = System.Drawing.ContentAlignment.MiddleRight
		Me.btnCreateEventsInOutlook.Location = New System.Drawing.Point(872, 8)
		Me.btnCreateEventsInOutlook.Name = "btnCreateEventsInOutlook"
		Me.btnCreateEventsInOutlook.Size = New System.Drawing.Size(716, 58)
		Me.btnCreateEventsInOutlook.TabIndex = 3
		Me.btnCreateEventsInOutlook.Text = "Create events in Outlook calendar for selected"
		Me.btnCreateEventsInOutlook.UseVisualStyleBackColor = True
		'
		'chkAddBoardingblocker
		'
		Me.chkAddBoardingblocker.AutoSize = True
		Me.chkAddBoardingblocker.Checked = True
		Me.chkAddBoardingblocker.CheckState = System.Windows.Forms.CheckState.Checked
		Me.chkAddBoardingblocker.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.chkAddBoardingblocker.Location = New System.Drawing.Point(0, 448)
		Me.chkAddBoardingblocker.Name = "chkAddBoardingblocker"
		Me.chkAddBoardingblocker.Padding = New System.Windows.Forms.Padding(15, 8, 8, 8)
		Me.chkAddBoardingblocker.Size = New System.Drawing.Size(1598, 40)
		Me.chkAddBoardingblocker.TabIndex = 4
		Me.chkAddBoardingblocker.Text = "For flights, add a one hour blocker before flight for security, boarding, etc."
		Me.chkAddBoardingblocker.UseVisualStyleBackColor = True
		'
		'lblWait
		'
		Me.lblWait.AutoSize = True
		Me.lblWait.BackColor = System.Drawing.Color.White
		Me.lblWait.Font = New System.Drawing.Font("Microsoft Sans Serif", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblWait.Location = New System.Drawing.Point(330, 157)
		Me.lblWait.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
		Me.lblWait.Name = "lblWait"
		Me.lblWait.Padding = New System.Windows.Forms.Padding(22, 23, 22, 23)
		Me.lblWait.Size = New System.Drawing.Size(683, 78)
		Me.lblWait.TabIndex = 5
		Me.lblWait.Text = "Comunicating with Outlook. Please stand by..."
		Me.lblWait.Visible = False
		'
		'Panel1
		'
		Me.Panel1.Controls.Add(Me.btnSaveICSs)
		Me.Panel1.Controls.Add(Me.btnCreateEventsInOutlook)
		Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
		Me.Panel1.Location = New System.Drawing.Point(0, 488)
		Me.Panel1.Name = "Panel1"
		Me.Panel1.Padding = New System.Windows.Forms.Padding(8)
		Me.Panel1.Size = New System.Drawing.Size(1598, 72)
		Me.Panel1.TabIndex = 6
		'
		'btnSaveICSs
		'
		Me.btnSaveICSs.Image = CType(resources.GetObject("btnSaveICSs.Image"), System.Drawing.Image)
		Me.btnSaveICSs.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
		Me.btnSaveICSs.Location = New System.Drawing.Point(10, 6)
		Me.btnSaveICSs.Name = "btnSaveICSs"
		Me.btnSaveICSs.Size = New System.Drawing.Size(840, 58)
		Me.btnSaveICSs.TabIndex = 4
		Me.btnSaveICSs.Text = "Create events and save them as files..."
		Me.btnSaveICSs.UseVisualStyleBackColor = True
		'
		'BTA2OutlookApplication
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(1598, 560)
		Me.Controls.Add(Me.lblWait)
		Me.Controls.Add(Me.chkAddBoardingblocker)
		Me.Controls.Add(Me.Panel1)
		Me.Controls.Add(Me.lst)
		Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
		Me.Name = "BTA2OutlookApplication"
		Me.Text = "BTA ITN to Outlook calendar"
		Me.Panel1.ResumeLayout(False)
		Me.ResumeLayout(False)
		Me.PerformLayout()

	End Sub
	Friend WithEvents lst As ListView
	Friend WithEvents ColumnHeader1 As ColumnHeader
	Friend WithEvents ColumnHeader2 As ColumnHeader
	Friend WithEvents ColumnHeader3 As ColumnHeader
	Friend WithEvents ImageList1 As ImageList
	Friend WithEvents ColumnHeader4 As ColumnHeader
	Friend WithEvents btnCreateEventsInOutlook As Button
	Friend WithEvents chkAddBoardingblocker As CheckBox
	Friend WithEvents lblWait As Label
	Friend WithEvents Panel1 As Panel
	Friend WithEvents btnSaveICSs As Button
End Class
