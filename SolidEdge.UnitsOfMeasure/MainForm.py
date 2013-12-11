import clr
clr.AddReference("System.Drawing")
clr.AddReference("Interop.SolidEdge")
clr.AddReference("System.Runtime.InteropServices")

import System.Windows.Forms
from System.Drawing import *
import SolidEdgeFramework as SEFramework
import System.Runtime.InteropServices as SRI
import ExternalExample as EE
import InternalExample as IE

class MainForm(System.Windows.Forms.Form):
    def __init__(self):

        listViewItem19 = System.Windows.Forms.ListViewItem(("Distance", "Meter"))
        listViewItem20 = System.Windows.Forms.ListViewItem(("Angle", "Radian"))
        listViewItem21 = System.Windows.Forms.ListViewItem(("Mass", "Kilogram"))
        listViewItem22 = System.Windows.Forms.ListViewItem(("Time", "Second"))
        listViewItem23 = System.Windows.Forms.ListViewItem(("Temperature", "Kelvin"))
        listViewItem24 = System.Windows.Forms.ListViewItem(("Charge", "Ampere"))
        listViewItem25 = System.Windows.Forms.ListViewItem(("Luminous Intensity", "Candela"))
        listViewItem26 = System.Windows.Forms.ListViewItem(("Amount of Substance", "Mole"))
        listViewItem27 = System.Windows.Forms.ListViewItem(("Solid Angle", "Steradian"))
        self.menuStrip1 = System.Windows.Forms.MenuStrip()
        self.fileToolStripMenuItem = System.Windows.Forms.ToolStripMenuItem()
        self.exitToolStripMenuItem = System.Windows.Forms.ToolStripMenuItem()
        self.externalPropertyGrid = System.Windows.Forms.PropertyGrid()
        self.internalPropertyGrid = System.Windows.Forms.PropertyGrid()
        self.splitContainer1 = System.Windows.Forms.SplitContainer()
        self.label1 = System.Windows.Forms.Label()
        self.label2 = System.Windows.Forms.Label()
        self.listView1 = System.Windows.Forms.ListView()
        self.label3 = System.Windows.Forms.Label()
        self.menuStrip1.SuspendLayout()
        self.splitContainer1.Panel1.SuspendLayout()
        self.splitContainer1.Panel2.SuspendLayout()
        self.splitContainer1.SuspendLayout()
        self.SuspendLayout()
        #
        # menuStrip1
        #
        self.menuStrip1.Items.Add(self.fileToolStripMenuItem)
        self.menuStrip1.Location = System.Drawing.Point(0, 0)
        self.menuStrip1.Name = "menuStrip1"
        self.menuStrip1.Size = System.Drawing.Size(784, 24)
        self.menuStrip1.TabIndex = 3
        self.menuStrip1.Text = "menuStrip1"
        #
        # fileToolStripMenuItem
        #
        self.fileToolStripMenuItem.DropDownItems.Add(self.exitToolStripMenuItem)
        self.fileToolStripMenuItem.Name = "fileToolStripMenuItem"
        self.fileToolStripMenuItem.Size = System.Drawing.Size(37, 20)
        self.fileToolStripMenuItem.Text = "&File"
        # 
        # exitToolStripMenuItem
        # 
        self.exitToolStripMenuItem.Name = "exitToolStripMenuItem"
        self.exitToolStripMenuItem.Size = System.Drawing.Size(92, 22)
        self.exitToolStripMenuItem.Text = "&Exit"
        self.exitToolStripMenuItem.Click += self.exitToolStripMenuItem_Click
        # 
        # externalPropertyGrid
        # 
        self.externalPropertyGrid.Dock = System.Windows.Forms.DockStyle.Fill
        self.externalPropertyGrid.Location = System.Drawing.Point(0, 23)
        self.externalPropertyGrid.Name = "externalPropertyGrid"
        self.externalPropertyGrid.PropertySort = System.Windows.Forms.PropertySort.NoSort
        self.externalPropertyGrid.Size = System.Drawing.Size(388, 305)
        self.externalPropertyGrid.TabIndex = 5
        self.externalPropertyGrid.ToolbarVisible = False
        self.externalPropertyGrid.PropertyValueChanged += self.externalPropertyGrid_PropertyValueChanged
        # 
        # internalPropertyGrid
        # 
        self.internalPropertyGrid.Dock = System.Windows.Forms.DockStyle.Fill
        self.internalPropertyGrid.Location = System.Drawing.Point(0, 23)
        self.internalPropertyGrid.Name = "internalPropertyGrid"
        self.internalPropertyGrid.PropertySort = System.Windows.Forms.PropertySort.NoSort
        self.internalPropertyGrid.Size = System.Drawing.Size(392, 305)
        self.internalPropertyGrid.TabIndex = 7
        self.internalPropertyGrid.ToolbarVisible = False
        self.internalPropertyGrid.PropertyValueChanged += self.internalPropertyGrid_PropertyValueChanged
        # 
        # splitContainer1
        # 
        self.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill
        self.splitContainer1.Location = System.Drawing.Point(0, 24)
        self.splitContainer1.Name = "splitContainer1"
        # 
        # splitContainer1.Panel1
        # 
        self.splitContainer1.Panel1.Controls.Add(self.internalPropertyGrid)
        self.splitContainer1.Panel1.Controls.Add(self.label2)
        # 
        # splitContainer1.Panel2
        # 
        self.splitContainer1.Panel2.Controls.Add(self.externalPropertyGrid)
        self.splitContainer1.Panel2.Controls.Add(self.label1)
        self.splitContainer1.Size = System.Drawing.Size(784, 328)
        self.splitContainer1.SplitterDistance = 392
        self.splitContainer1.TabIndex = 8
        # 
        # label1
        # 
        self.label1.BackColor = System.Drawing.Color.LightSteelBlue
        self.label1.Dock = System.Windows.Forms.DockStyle.Top
        self.label1.Font = System.Drawing.Font("Segoe UI", 8.25, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self.label1.Location = System.Drawing.Point(0, 0)
        self.label1.Name = "label1"
        self.label1.Padding = System.Windows.Forms.Padding(5)
        self.label1.Size = System.Drawing.Size(388, 23)
        self.label1.TabIndex = 6
        self.label1.Text = "External Units --> Internal Units"
        self.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        # 
        # label2
        # 
        self.label2.BackColor = System.Drawing.Color.LightSteelBlue
        self.label2.Dock = System.Windows.Forms.DockStyle.Top
        self.label2.Font = System.Drawing.Font("Segoe UI", 8.25, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self.label2.Location = System.Drawing.Point(0, 0)
        self.label2.Name = "label2"
        self.label2.Padding = System.Windows.Forms.Padding(5)
        self.label2.Size = System.Drawing.Size(392, 23)
        self.label2.TabIndex = 8
        self.label2.Text = "Internal Units --> External Units"
        self.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        # 
        # listView1
        #
        self.columnHeader1 = System.Windows.Forms.ColumnHeader()
        self.columnHeader2 = System.Windows.Forms.ColumnHeader()
        self.listView1.Columns.AddRange((self.columnHeader1, self.columnHeader2))
        self.listView1.Dock = System.Windows.Forms.DockStyle.Bottom
        self.listView1.FullRowSelect = True

        listitems = (listViewItem19, listViewItem20, listViewItem21, listViewItem22, listViewItem23, listViewItem24, listViewItem25, listViewItem26, listViewItem27)
        self.listView1.Items.AddRange(listitems)
        self.listView1.Location = System.Drawing.Point(0, 375)
        self.listView1.MultiSelect = False
        self.listView1.Name = "listView1"
        self.listView1.Size = System.Drawing.Size(784, 186)
        self.listView1.TabIndex = 9
        self.listView1.UseCompatibleStateImageBehavior = False
        self.listView1.View = System.Windows.Forms.View.Details
        # 
        # columnHeader1
        # 
        self.columnHeader1.Text = "Unit Type"
        self.columnHeader1.Width = 128
        # 
        # columnHeader2
        # 
        self.columnHeader2.Text = "Internal Units"
        self.columnHeader2.Width = 90
        # 
        # label3
        # 
        self.label3.BackColor = System.Drawing.Color.LightSteelBlue
        self.label3.Dock = System.Windows.Forms.DockStyle.Bottom
        self.label3.Font = System.Drawing.Font("Segoe UI", 8.25, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, 0)
        self.label3.Location = System.Drawing.Point(0, 352)
        self.label3.Name = "label3"
        self.label3.Padding = System.Windows.Forms.Padding(5)
        self.label3.Size = System.Drawing.Size(784, 23)
        self.label3.TabIndex = 10
        self.label3.Text = "Internal (Database) Unit Reference"
        self.label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        # 
        # MainForm
        # 
        self.AutoScaleDimensions = System.Drawing.SizeF(6, 13)
        self.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        self.ClientSize = System.Drawing.Size(784, 561)
        self.Controls.Add(self.splitContainer1)
        self.Controls.Add(self.label3)
        self.Controls.Add(self.listView1)
        self.Controls.Add(self.menuStrip1)
        self.Font = System.Drawing.Font("Segoe UI", 8.25, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, 0)
        self.MainMenuStrip = self.menuStrip1
        self.Name = "MainForm"
        self.Text = "Units of Measure"
        self.Load += self.MainForm_Load
        self.menuStrip1.ResumeLayout(False)
        self.menuStrip1.PerformLayout()
        self.splitContainer1.Panel1.ResumeLayout(False)
        self.splitContainer1.Panel2.ResumeLayout(False)
        self.splitContainer1.ResumeLayout(False)
        self.ResumeLayout(False)
        self.PerformLayout()

    def MainForm_Load(self, sender, event_args):
        # Connect to or start Solid Edge.
        #application = SEFramework.Contrib.ApplicationHelper.Connect(True)

        # Connect to a running instance of Solid Edge
        application = SRI.Marshal.GetActiveObject("SolidEdge.Application")

        # Ensure Solid Edge is visible.
        application.Visible = True

        # Get a reference to the Documents collection.
        documents = application.Documents

        if documents.Count == 0 :
            documents.Add("SolidEdge.PartDocument")

        self.externalPropertyGrid.SelectedObject = EE.ExternalExample()
        self.internalPropertyGrid.SelectedObject = IE.InternalExample()

    def exitToolStripMenuItem_Click(self, sender, event_args):
		self.Close()

    def externalPropertyGrid_PropertyValueChanged(self, sender, event_args):
		self.externalPropertyGrid.Refresh()

    def internalPropertyGrid_PropertyValueChanged(self, sender, event_args):
		self.internalPropertyGrid.Refresh()
