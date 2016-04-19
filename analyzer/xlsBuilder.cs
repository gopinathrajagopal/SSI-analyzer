namespace analyzer
{
    using excelTools;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Drawing;
    using System.Windows.Forms;

    public class xlsBuilder : Form
    {
        private List<report> allRep;
        private Button bCancel;
        private Button bOK;
        private IContainer components;
        private Label giCol;
        private Label giDataset;
        private GroupBox giGrp;
        private Label giRow;
        private Label goCol;
        private Label goDataset;
        private GroupBox goGrp;
        private Label goRow;
        private GroupBox groupBox2;
        private GroupBox groupBox3;
        private ListBox lstInReps;
        private ListBox lstOutReps;
        private Panel panel2;
        private Panel panel3;
        private List<sugarData> sugz;
        private ToolStripLabel toolStripLabel1;
        private ToolStripButton tsAdd;
        private ToolStripButton tsDown;
        private ToolStripButton tsRemove;
        private ToolStrip tStrip;
        private ToolStripButton tsUp;

        public xlsBuilder(List<sugarData> sugarDats)
        {
            this.sugz = sugarDats;
            this.InitializeComponent();
        }

        private void addRep()
        {
            if (this.lstOutReps.Items.Count != 0)
            {
                string item = this.lstOutReps.SelectedItem.ToString();
                this.lstInReps.Items.Add(item);
                this.lstOutReps.Items.Remove(this.lstOutReps.SelectedItem);
                this.xNameToRep(item.ToString()).xlsFlag = true;
                this.lstInReps.SelectedItem = item;
                if (this.lstOutReps.Items.Count != 0)
                {
                    this.lstOutReps.SelectedIndex = 0;
                }
            }
        }

        private void bCancel_Click(object sender, EventArgs e)
        {
            base.DialogResult = DialogResult.Cancel;
        }

        private void bOK_Click(object sender, EventArgs e)
        {
            try
            {
                object workbook = genExcelTools.newWorkbook(genExcelTools.getApp());
                object sheet = genExcelTools.getSheet(workbook, 1);
                for (int i = 0; i < this.lstInReps.Items.Count; i++)
                {
                    report report = this.xNameToRep(this.lstInReps.Items[i].ToString());
                    string sheetName = errorLog.filtIllegalXlsChars(report.xlsName);
                    if (i > 0)
                    {
                        sheet = genExcelTools.newSheetEnd(workbook, sheetName);
                    }
                    genExcelTools.CmtAzToSheet(sheet, report.rep, report.formByCol, report.formats, sheetName);
                }
                genExcelTools.showExcel(workbook, false);
            }
            catch (Exception exception)
            {
                MessageBox.Show("An error occurred while trying to create the Excel workbook." + Environment.NewLine + Environment.NewLine + exception.Message);
                return;
            }
            base.DialogResult = DialogResult.OK;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void enabler()
        {
            if (this.lstInReps.Items.Count == 0)
            {
                this.bOK.Enabled = false;
                this.tsRemove.Enabled = false;
                this.tsUp.Enabled = false;
                this.tsDown.Enabled = false;
            }
            else
            {
                this.bOK.Enabled = true;
                this.tsRemove.Enabled = true;
                if (this.lstInReps.SelectedIndex == 0)
                {
                    this.tsUp.Enabled = false;
                }
                else
                {
                    this.tsUp.Enabled = true;
                }
                if (this.lstInReps.SelectedIndex == (this.lstInReps.Items.Count - 1))
                {
                    this.tsDown.Enabled = false;
                }
                else
                {
                    this.tsDown.Enabled = true;
                }
            }
            if (this.lstOutReps.Items.Count == 0)
            {
                this.tsAdd.Enabled = false;
            }
            else
            {
                this.tsAdd.Enabled = true;
            }
        }

        private void InitializeComponent()
        {
            ComponentResourceManager manager = new ComponentResourceManager(typeof(xlsBuilder));
            this.lstOutReps = new ListBox();
            this.lstInReps = new ListBox();
            this.goGrp = new GroupBox();
            this.groupBox2 = new GroupBox();
            this.groupBox3 = new GroupBox();
            this.panel2 = new Panel();
            this.panel3 = new Panel();
            this.tStrip = new ToolStrip();
            this.goDataset = new Label();
            this.goCol = new Label();
            this.goRow = new Label();
            this.giGrp = new GroupBox();
            this.giRow = new Label();
            this.giCol = new Label();
            this.giDataset = new Label();
            this.tsAdd = new ToolStripButton();
            this.tsRemove = new ToolStripButton();
            this.tsDown = new ToolStripButton();
            this.tsUp = new ToolStripButton();
            this.toolStripLabel1 = new ToolStripLabel();
            this.bOK = new Button();
            this.bCancel = new Button();
            this.goGrp.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.tStrip.SuspendLayout();
            this.giGrp.SuspendLayout();
            base.SuspendLayout();
            this.lstOutReps.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstOutReps.Dock = DockStyle.Top;
            this.lstOutReps.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstOutReps.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstOutReps.FormattingEnabled = true;
            this.lstOutReps.Location = new Point(3, 0x10);
            this.lstOutReps.Name = "lstOutReps";
            this.lstOutReps.Size = new Size(0x8b, 0xc7);
            this.lstOutReps.TabIndex = 1;
            this.lstOutReps.DoubleClick += new EventHandler(this.lstOutReps_DoubleClick);
            this.lstOutReps.SelectedIndexChanged += new EventHandler(this.lstOutReps_SelectedIndexChanged);
            this.lstInReps.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstInReps.Dock = DockStyle.Top;
            this.lstInReps.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstInReps.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstInReps.FormattingEnabled = true;
            this.lstInReps.Location = new Point(3, 0x10);
            this.lstInReps.Name = "lstInReps";
            this.lstInReps.Size = new Size(0x8b, 0xc7);
            this.lstInReps.TabIndex = 1;
            this.lstInReps.DoubleClick += new EventHandler(this.lstInReps_DoubleClick);
            this.lstInReps.SelectedIndexChanged += new EventHandler(this.lstInReps_SelectedIndexChanged);
            this.goGrp.Controls.Add(this.goRow);
            this.goGrp.Controls.Add(this.goCol);
            this.goGrp.Controls.Add(this.goDataset);
            this.goGrp.Dock = DockStyle.Bottom;
            this.goGrp.Location = new Point(3, 0xdd);
            this.goGrp.Name = "goGrp";
            this.goGrp.Size = new Size(0x8b, 0x45);
            this.goGrp.TabIndex = 2;
            this.goGrp.TabStop = false;
            this.groupBox2.Controls.Add(this.giGrp);
            this.groupBox2.Controls.Add(this.lstInReps);
            this.groupBox2.Dock = DockStyle.Right;
            this.groupBox2.Location = new Point(0xaf, 0);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new Size(0x91, 0x125);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Included Reports";
            this.groupBox3.Controls.Add(this.goGrp);
            this.groupBox3.Controls.Add(this.lstOutReps);
            this.groupBox3.Dock = DockStyle.Left;
            this.groupBox3.Location = new Point(0, 0);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new Size(0x91, 0x125);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Excluded Reports";
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.groupBox2);
            this.panel2.Controls.Add(this.groupBox3);
            this.panel2.Dock = DockStyle.Top;
            this.panel2.Location = new Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new Size(320, 0x125);
            this.panel2.TabIndex = 7;
            this.panel3.Controls.Add(this.tStrip);
            this.panel3.Dock = DockStyle.Fill;
            this.panel3.Location = new Point(0x91, 0);
            this.panel3.Name = "panel3";
            this.panel3.Size = new Size(30, 0x125);
            this.panel3.TabIndex = 7;
            this.tStrip.Dock = DockStyle.Fill;
            this.tStrip.GripStyle = ToolStripGripStyle.Hidden;
            this.tStrip.Items.AddRange(new ToolStripItem[] { this.toolStripLabel1, this.tsAdd, this.tsRemove, this.tsUp, this.tsDown });
            this.tStrip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow;
            this.tStrip.Location = new Point(0, 0);
            this.tStrip.Name = "tStrip";
            this.tStrip.Size = new Size(30, 0x125);
            this.tStrip.TabIndex = 0;
            this.tStrip.Text = "toolStrip1";
            this.goDataset.AutoSize = true;
            this.goDataset.Location = new Point(10, 20);
            this.goDataset.Name = "goDataset";
            this.goDataset.Size = new Size(0, 13);
            this.goDataset.TabIndex = 0;
            this.goCol.AutoSize = true;
            this.goCol.Location = new Point(10, 50);
            this.goCol.Name = "goCol";
            this.goCol.Size = new Size(0, 13);
            this.goCol.TabIndex = 1;
            this.goRow.AutoSize = true;
            this.goRow.Location = new Point(0x53, 50);
            this.goRow.Name = "goRow";
            this.goRow.Size = new Size(0, 13);
            this.goRow.TabIndex = 2;
            this.giGrp.Controls.Add(this.giRow);
            this.giGrp.Controls.Add(this.giCol);
            this.giGrp.Controls.Add(this.giDataset);
            this.giGrp.Dock = DockStyle.Bottom;
            this.giGrp.Location = new Point(3, 0xdd);
            this.giGrp.Name = "giGrp";
            this.giGrp.Size = new Size(0x8b, 0x45);
            this.giGrp.TabIndex = 3;
            this.giGrp.TabStop = false;
            this.giRow.AutoSize = true;
            this.giRow.Location = new Point(0x53, 50);
            this.giRow.Name = "giRow";
            this.giRow.Size = new Size(0, 13);
            this.giRow.TabIndex = 2;
            this.giCol.AutoSize = true;
            this.giCol.Location = new Point(10, 50);
            this.giCol.Name = "giCol";
            this.giCol.Size = new Size(0, 13);
            this.giCol.TabIndex = 1;
            this.giDataset.AutoSize = true;
            this.giDataset.Location = new Point(10, 20);
            this.giDataset.Name = "giDataset";
            this.giDataset.Size = new Size(0, 13);
            this.giDataset.TabIndex = 0;
            this.tsAdd.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.tsAdd.Image = (Image) manager.GetObject("tsAdd.Image");
            this.tsAdd.ImageTransparentColor = Color.Magenta;
            this.tsAdd.Name = "tsAdd";
            this.tsAdd.Size = new Size(0x1c, 20);
            this.tsAdd.Text = "Add report";
            this.tsAdd.Click += new EventHandler(this.tsAdd_Click);
            this.tsRemove.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.tsRemove.Image = (Image) manager.GetObject("tsRemove.Image");
            this.tsRemove.ImageTransparentColor = Color.Magenta;
            this.tsRemove.Name = "tsRemove";
            this.tsRemove.Size = new Size(0x1c, 20);
            this.tsRemove.Text = "Remove report";
            this.tsRemove.Click += new EventHandler(this.tsRemove_Click);
            this.tsDown.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.tsDown.Image = (Image) manager.GetObject("tsDown.Image");
            this.tsDown.ImageTransparentColor = Color.Magenta;
            this.tsDown.Name = "tsDown";
            this.tsDown.Size = new Size(0x1c, 20);
            this.tsDown.Text = "Move report down";
            this.tsDown.Click += new EventHandler(this.tsDown_Click);
            this.tsUp.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.tsUp.Image = (Image) manager.GetObject("tsUp.Image");
            this.tsUp.ImageTransparentColor = Color.Magenta;
            this.tsUp.Name = "tsUp";
            this.tsUp.Size = new Size(0x1c, 20);
            this.tsUp.Text = "Move report up";
            this.tsUp.Click += new EventHandler(this.tsUp_Click);
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new Size(0x1c, 13);
            this.toolStripLabel1.Text = " ";
            this.bOK.Location = new Point(12, 0x12b);
            this.bOK.Name = "bOK";
            this.bOK.Size = new Size(0x4b, 0x17);
            this.bOK.TabIndex = 0;
            this.bOK.Text = "OK";
            this.bOK.UseVisualStyleBackColor = true;
            this.bOK.Click += new EventHandler(this.bOK_Click);
            this.bCancel.Location = new Point(0xe9, 0x12b);
            this.bCancel.Name = "bCancel";
            this.bCancel.Size = new Size(0x4b, 0x17);
            this.bCancel.TabIndex = 0;
            this.bCancel.Text = "Cancel";
            this.bCancel.UseVisualStyleBackColor = true;
            this.bCancel.Click += new EventHandler(this.bCancel_Click);
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            this.BackColor = Color.White;
            base.ClientSize = new Size(320, 0x148);
            base.Controls.Add(this.bCancel);
            base.Controls.Add(this.bOK);
            base.Controls.Add(this.panel2);
            base.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            base.Name = "xlsBuilder";
            this.Text = "Excel Workbook";
            base.Load += new EventHandler(this.xlsBuilder_Load);
            this.goGrp.ResumeLayout(false);
            this.goGrp.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.tStrip.ResumeLayout(false);
            this.tStrip.PerformLayout();
            this.giGrp.ResumeLayout(false);
            this.giGrp.PerformLayout();
            base.ResumeLayout(false);
        }

        private void lstInReps_DoubleClick(object sender, EventArgs e)
        {
            this.removeRep();
        }

        private void lstInReps_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.lstInReps.SelectedItems.Count == 0)
            {
                this.giGrp.Text = string.Empty;
                this.giDataset.Text = string.Empty;
                this.giCol.Text = string.Empty;
                this.giRow.Text = string.Empty;
                this.bOK.Enabled = false;
            }
            else
            {
                this.bOK.Enabled = true;
                report report = this.xNameToRep(this.lstInReps.SelectedItem.ToString());
                this.giGrp.Text = report.xlsName;
                this.giDataset.Text = "Dataset:  " + report.sugDaddy.name;
                this.giCol.Text = report.rep.Columns.Count.ToString() + " columns";
                this.giRow.Text = report.rep.Rows.Count.ToString() + " rows";
                this.enabler();
            }
        }

        private void lstOutReps_DoubleClick(object sender, EventArgs e)
        {
            this.addRep();
        }

        private void lstOutReps_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (this.lstOutReps.SelectedItems.Count == 0)
            {
                this.goGrp.Text = string.Empty;
                this.goDataset.Text = string.Empty;
                this.goCol.Text = string.Empty;
                this.goRow.Text = string.Empty;
            }
            else
            {
                report report = this.xNameToRep(this.lstOutReps.SelectedItem.ToString());
                this.goGrp.Text = report.xlsName;
                this.goDataset.Text = "Dataset:  " + report.sugDaddy.name;
                this.goCol.Text = report.rep.Columns.Count.ToString() + " columns";
                this.goRow.Text = report.rep.Rows.Count.ToString() + " rows";
                this.enabler();
            }
        }

        private void removeRep()
        {
            if (this.lstInReps.Items.Count != 0)
            {
                string item = this.lstInReps.SelectedItem.ToString();
                this.lstOutReps.Items.Add(item);
                this.lstInReps.Items.Remove(this.lstInReps.SelectedItem);
                this.xNameToRep(item.ToString()).xlsFlag = false;
                this.lstOutReps.SelectedItem = item;
                if (this.lstInReps.Items.Count != 0)
                {
                    this.lstInReps.SelectedIndex = 0;
                }
            }
        }

        private void tsAdd_Click(object sender, EventArgs e)
        {
            this.addRep();
        }

        private void tsDown_Click(object sender, EventArgs e)
        {
            string item = this.lstInReps.SelectedItem.ToString();
            int index = this.lstInReps.SelectedIndex + 1;
            this.lstInReps.Items.Remove(this.lstInReps.SelectedItem);
            this.lstInReps.Items.Insert(index, item);
            this.lstInReps.SelectedIndex = index;
        }

        private void tsRemove_Click(object sender, EventArgs e)
        {
            this.removeRep();
        }

        private void tsUp_Click(object sender, EventArgs e)
        {
            string item = this.lstInReps.SelectedItem.ToString();
            int index = this.lstInReps.SelectedIndex - 1;
            this.lstInReps.Items.Remove(this.lstInReps.SelectedItem);
            this.lstInReps.Items.Insert(index, item);
            this.lstInReps.SelectedIndex = index;
        }

        private void xlsBuilder_Load(object sender, EventArgs e)
        {
            this.allRep = new List<report>();
            foreach (sugarData data in this.sugz)
            {
                if (data.loaded)
                {
                    foreach (report report in data.repList.reps)
                    {
                        if (report.rendered)
                        {
                            report.xlsName = report.name;
                            foreach (report report2 in this.allRep)
                            {
                                if (report.xlsName == report2.xlsName)
                                {
                                    report.xlsName = report.sugDaddy.name + "." + report.xlsName;
                                }
                            }
                            this.allRep.Add(report);
                            if (report.xlsFlag)
                            {
                                this.lstInReps.Items.Add(report.xlsName);
                            }
                            else
                            {
                                this.lstOutReps.Items.Add(report.xlsName);
                            }
                        }
                    }
                }
            }
            if (this.lstInReps.Items.Count > 0)
            {
                this.lstInReps.SetSelected(0, true);
            }
            if (this.lstOutReps.Items.Count > 0)
            {
                this.lstOutReps.SetSelected(0, true);
            }
            this.enabler();
        }

        private report xNameToRep(string name)
        {
            foreach (report report in this.allRep)
            {
                if (report.xlsName == name)
                {
                    return report;
                }
            }
            return null;
        }
    }
}

