namespace analyzer
{
    using analyzer.Properties;
    using excelTools;
    using System;
    using System.Collections;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Text;
    using System.Windows.Forms;

    public class repView : Form
    {
        public bool active;
        private ToolStripMenuItem cmExHedz;
        private ContextMenuStrip cmHedz;
        private ToolStripMenuItem cmIncHedz;
        private IContainer components;
        public bool connected = true;
        private report curRep;
        public bool defLocSet;
        private DataGridView dGrid;
        private Panel panel1;
        private sessionLog SLog;
        private DataTable tbl;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripButton tsCAP;
        private ToolStripButton tsCap07;
        public ToolStripButton tsConnect;
        private ToolStripButton tsCopyAll;
        private ToolStripButton tsCopySel;
        private ToolStripButton tsDecIn;
        private ToolStripButton tsDecOut;
        public ToolStripButton tsDiscon;
        private ToolStripButton tsExcel;
        private ToolStrip tsRep;

        public repView(sessionLog sessionLogRef)
        {
            this.SLog = sessionLogRef;
            this.InitializeComponent();
        }

        private void cmExHedz_Click(object sender, EventArgs e)
        {
            try
            {
                this.cmIncHedz.Checked = false;
                this.cmExHedz.Checked = true;
                this.dGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void cmIncHedz_Click(object sender, EventArgs e)
        {
            try
            {
                this.cmIncHedz.Checked = true;
                this.cmExHedz.Checked = false;
                this.dGrid.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        public void connect()
        {
            try
            {
                this.connected = true;
                base.Icon = Resources.winConnect;
                this.tsDiscon.Visible = true;
                this.tsConnect.Visible = false;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void dGrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.Value.ToString().StartsWith("esc:"))
            {
                if (e.Value.ToString().Contains("Blue"))
                {
                    e.Value = "Blue";
                    e.CellStyle.BackColor = Color.Blue;
                    e.CellStyle.ForeColor = Color.White;
                }
                if (e.Value.ToString().Contains("Green"))
                {
                    e.Value = "Green";
                    e.CellStyle.BackColor = Color.Green;
                    e.CellStyle.ForeColor = Color.White;
                }
                if (e.Value.ToString().Contains("Yellow"))
                {
                    e.Value = "Yellow";
                    e.CellStyle.BackColor = Color.Yellow;
                    e.CellStyle.ForeColor = Color.Black;
                }
                if (e.Value.ToString().Contains("Red"))
                {
                    e.Value = "Red";
                    e.CellStyle.BackColor = Color.Red;
                    e.CellStyle.ForeColor = Color.White;
                }
                if (e.Value.ToString().Contains("Magenta"))
                {
                    e.Value = "Magenta";
                    e.CellStyle.BackColor = Color.FromArgb(0xe2, 0, 0x74);
                    e.CellStyle.ForeColor = Color.White;
                }
                if (e.Value.ToString() == "esc:noCode")
                {
                    e.Value = "N/A";
                }
            }
            else
            {
                double result = 0.0;
                if (double.TryParse(e.Value.ToString(), out result))
                {
                    string str = "0";
                    if (this.curRep.decimalVis > 0)
                    {
                        str = str + ".";
                        for (byte i = 1; i <= this.curRep.decimalVis; i = (byte) (i + 1))
                        {
                            str = str + "#";
                        }
                    }
                    string str2 = str + "%";
                    if (this.curRep.formByCol)
                    {
                        int num3 = this.dGrid.Columns.Count - this.curRep.formats.Length;
                        if ((num3 <= e.ColumnIndex) && this.curRep.formats[e.ColumnIndex - num3].EndsWith("%"))
                        {
                            e.CellStyle.Format = str2;
                            return;
                        }
                    }
                    else if (this.curRep.formats[0].EndsWith("%"))
                    {
                        e.CellStyle.Format = str2;
                        return;
                    }
                    if ((result % 1.0) == 0.0)
                    {
                        e.CellStyle.Format = "0";
                    }
                    else
                    {
                        e.CellStyle.Format = str;
                    }
                }
            }
        }

        private void dGrid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (this.tbl.Rows.Count != 0)
            {
                try
                {
                    string dir = "DESC";
                    if (this.dGrid.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection == SortOrder.Descending)
                    {
                        dir = "ASC";
                    }
                    this.sortRep(dir, e.ColumnIndex);
                    foreach (DataGridViewColumn column in this.dGrid.Columns)
                    {
                        column.HeaderCell.SortGlyphDirection = SortOrder.None;
                    }
                    if (dir == "ASC")
                    {
                        this.dGrid.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = SortOrder.Ascending;
                    }
                    else
                    {
                        this.dGrid.Columns[e.ColumnIndex].HeaderCell.SortGlyphDirection = SortOrder.Descending;
                    }
                    this.setForamt();
                }
                catch (Exception exception)
                {
                    errorLog.writeError(exception, this.SLog);
                    base.Close();
                }
            }
        }

        public void disconnect()
        {
            try
            {
                this.connected = false;
                base.Icon = Resources.winDiscon;
                this.tsDiscon.Visible = false;
                this.tsConnect.Visible = true;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && (this.components != null))
            {
                this.components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new Container();
            DataGridViewCellStyle style = new DataGridViewCellStyle();
            DataGridViewCellStyle style2 = new DataGridViewCellStyle();
            DataGridViewCellStyle style3 = new DataGridViewCellStyle();
            ComponentResourceManager manager = new ComponentResourceManager(typeof(repView));
            this.dGrid = new DataGridView();
            this.tsRep = new ToolStrip();
            this.tsDiscon = new ToolStripButton();
            this.tsConnect = new ToolStripButton();
            this.toolStripSeparator1 = new ToolStripSeparator();
            this.tsCopySel = new ToolStripButton();
            this.tsCopyAll = new ToolStripButton();
            this.toolStripSeparator2 = new ToolStripSeparator();
            this.tsExcel = new ToolStripButton();
            this.tsCAP = new ToolStripButton();
            this.tsCap07 = new ToolStripButton();
            this.toolStripSeparator3 = new ToolStripSeparator();
            this.tsDecOut = new ToolStripButton();
            this.tsDecIn = new ToolStripButton();
            this.cmHedz = new ContextMenuStrip(this.components);
            this.cmIncHedz = new ToolStripMenuItem();
            this.cmExHedz = new ToolStripMenuItem();
            this.panel1 = new Panel();
            ((ISupportInitialize) this.dGrid).BeginInit();
            this.tsRep.SuspendLayout();
            this.cmHedz.SuspendLayout();
            this.panel1.SuspendLayout();
            base.SuspendLayout();
            this.dGrid.AllowUserToAddRows = false;
            this.dGrid.AllowUserToDeleteRows = false;
            style.BackColor = Color.FromArgb(0xed, 0xed, 0xed);
            this.dGrid.AlternatingRowsDefaultCellStyle = style;
            this.dGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            this.dGrid.CellBorderStyle = DataGridViewCellBorderStyle.Sunken;
            style2.Alignment = DataGridViewContentAlignment.MiddleCenter;
            style2.BackColor = SystemColors.Control;
            style2.Font = new Font("Arial", 8.25f, FontStyle.Bold, GraphicsUnit.Point, 0);
            style2.ForeColor = SystemColors.WindowText;
            style2.SelectionBackColor = SystemColors.Highlight;
            style2.SelectionForeColor = SystemColors.HighlightText;
            style2.WrapMode = DataGridViewTriState.False;
            this.dGrid.ColumnHeadersDefaultCellStyle = style2;
            this.dGrid.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            style3.Alignment = DataGridViewContentAlignment.MiddleCenter;
            style3.BackColor = SystemColors.Window;
            style3.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            style3.ForeColor = SystemColors.ControlText;
            style3.SelectionBackColor = SystemColors.Highlight;
            style3.SelectionForeColor = SystemColors.HighlightText;
            style3.WrapMode = DataGridViewTriState.False;
            this.dGrid.DefaultCellStyle = style3;
            this.dGrid.Dock = DockStyle.Fill;
            this.dGrid.EditMode = DataGridViewEditMode.EditProgrammatically;
            this.dGrid.Location = new Point(0, 0);
            this.dGrid.Name = "dGrid";
            this.dGrid.ReadOnly = true;
            this.dGrid.RowHeadersVisible = false;
            this.dGrid.RowTemplate.Height = 0x11;
            this.dGrid.Size = new Size(0x278, 0x1a5);
            this.dGrid.TabIndex = 0;
            this.dGrid.ColumnHeaderMouseClick += new DataGridViewCellMouseEventHandler(this.dGrid_ColumnHeaderMouseClick);
            this.dGrid.CellFormatting += new DataGridViewCellFormattingEventHandler(this.dGrid_CellFormatting);
            this.tsRep.GripStyle = ToolStripGripStyle.Hidden;
            this.tsRep.Items.AddRange(new ToolStripItem[] { this.tsDiscon, this.tsConnect, this.toolStripSeparator1, this.tsCopySel, this.tsCopyAll, this.toolStripSeparator2, this.tsExcel, this.tsCAP, this.tsCap07, this.toolStripSeparator3, this.tsDecOut, this.tsDecIn });
            this.tsRep.Location = new Point(0, 0);
            this.tsRep.Name = "tsRep";
            this.tsRep.Size = new Size(0x278, 0x19);
            this.tsRep.TabIndex = 1;
            this.tsDiscon.Image = (Image) manager.GetObject("tsDiscon.Image");
            this.tsDiscon.ImageTransparentColor = Color.Magenta;
            this.tsDiscon.Name = "tsDiscon";
            this.tsDiscon.Size = new Size(0x68, 0x16);
            this.tsDiscon.Text = "Disconnect View";
            this.tsDiscon.Click += new EventHandler(this.tsDiscon_Click);
            this.tsConnect.Image = (Image) manager.GetObject("tsConnect.Image");
            this.tsConnect.ImageTransparentColor = Color.Magenta;
            this.tsConnect.Name = "tsConnect";
            this.tsConnect.Size = new Size(0x5c, 0x16);
            this.tsConnect.Text = "Connect View";
            this.tsConnect.Visible = false;
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new Size(6, 0x19);
            this.tsCopySel.Image = (Image) manager.GetObject("tsCopySel.Image");
            this.tsCopySel.ImageTransparentColor = Color.Magenta;
            this.tsCopySel.Name = "tsCopySel";
            this.tsCopySel.Size = new Size(0x62, 0x16);
            this.tsCopySel.Text = "Copy Selection";
            this.tsCopySel.MouseDown += new MouseEventHandler(this.tsCopySel_MouseDown);
            this.tsCopySel.Click += new EventHandler(this.tsCopySel_Click);
            this.tsCopyAll.Image = (Image) manager.GetObject("tsCopyAll.Image");
            this.tsCopyAll.ImageTransparentColor = Color.Magenta;
            this.tsCopyAll.Name = "tsCopyAll";
            this.tsCopyAll.Size = new Size(0x42, 0x16);
            this.tsCopyAll.Text = "Copy All";
            this.tsCopyAll.ToolTipText = "Copy whole report to clipboard";
            this.tsCopyAll.MouseDown += new MouseEventHandler(this.tsCopyAll_MouseDown);
            this.tsCopyAll.Click += new EventHandler(this.tsCopyAll_Click);
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new Size(6, 0x19);
            this.tsExcel.Image = (Image) manager.GetObject("tsExcel.Image");
            this.tsExcel.ImageTransparentColor = Color.Magenta;
            this.tsExcel.Name = "tsExcel";
            this.tsExcel.Size = new Size(100, 0x16);
            this.tsExcel.Text = "Export to Excel";
            this.tsExcel.Click += new EventHandler(this.tsExcel_Click);
            this.tsCAP.Image = (Image) manager.GetObject("tsCAP.Image");
            this.tsCAP.ImageTransparentColor = Color.Magenta;
            this.tsCAP.Name = "tsCAP";
            this.tsCAP.Size = new Size(70, 0x16);
            this.tsCAP.Text = "CAP Tool";
            this.tsCAP.Visible = false;
            this.tsCAP.Click += new EventHandler(this.tsCAP_Click);
            this.tsCap07.Image = (Image) manager.GetObject("tsCap07.Image");
            this.tsCap07.ImageTransparentColor = Color.Magenta;
            this.tsCap07.Name = "tsCap07";
            this.tsCap07.Size = new Size(0x55, 20);
            this.tsCap07.Text = "CAP Tool 07";
            this.tsCap07.Visible = false;
            this.tsCap07.Click += new EventHandler(this.tsCAP07_Click);
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new Size(6, 0x19);
            this.tsDecOut.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.tsDecOut.Image = (Image) manager.GetObject("tsDecOut.Image");
            this.tsDecOut.ImageTransparentColor = Color.Magenta;
            this.tsDecOut.Name = "tsDecOut";
            this.tsDecOut.Size = new Size(0x17, 20);
            this.tsDecOut.ToolTipText = "Move decimal places out";
            this.tsDecOut.Click += new EventHandler(this.tsDecOut_Click);
            this.tsDecIn.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.tsDecIn.Image = (Image) manager.GetObject("tsDecIn.Image");
            this.tsDecIn.ImageTransparentColor = Color.Magenta;
            this.tsDecIn.Name = "tsDecIn";
            this.tsDecIn.Size = new Size(0x17, 20);
            this.tsDecIn.ToolTipText = "Move decimal places in";
            this.tsDecIn.Click += new EventHandler(this.tsDecIn_Click);
            this.cmHedz.Items.AddRange(new ToolStripItem[] { this.cmIncHedz, this.cmExHedz });
            this.cmHedz.Name = "cmHedz";
            this.cmHedz.Size = new Size(0xa6, 0x30);
            this.cmIncHedz.Checked = true;
            this.cmIncHedz.CheckState = CheckState.Checked;
            this.cmIncHedz.Name = "cmIncHedz";
            this.cmIncHedz.Size = new Size(0xa5, 0x16);
            this.cmIncHedz.Text = "Include Headers";
            this.cmIncHedz.Click += new EventHandler(this.cmIncHedz_Click);
            this.cmExHedz.Name = "cmExHedz";
            this.cmExHedz.Size = new Size(0xa5, 0x16);
            this.cmExHedz.Text = "Exclude Headers";
            this.cmExHedz.Click += new EventHandler(this.cmExHedz_Click);
            this.panel1.Controls.Add(this.dGrid);
            this.panel1.Dock = DockStyle.Fill;
            this.panel1.Location = new Point(0, 0x19);
            this.panel1.Name = "panel1";
            this.panel1.Size = new Size(0x278, 0x1a5);
            this.panel1.TabIndex = 2;
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            base.ClientSize = new Size(0x278, 0x1be);
            base.Controls.Add(this.panel1);
            base.Controls.Add(this.tsRep);
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.Name = "repView";
            base.StartPosition = FormStartPosition.Manual;
            base.FormClosing += new FormClosingEventHandler(this.repView_FormClosing);
            ((ISupportInitialize) this.dGrid).EndInit();
            this.tsRep.ResumeLayout(false);
            this.tsRep.PerformLayout();
            this.cmHedz.ResumeLayout(false);
            this.panel1.ResumeLayout(false);
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        private void repView_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                if (this.connected)
                {
                    this.active = false;
                    e.Cancel = true;
                    base.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void setCapViz()
        {
            bool flag = true;
            if (this.curRep.formByCol)
            {
                flag = false;
            }
            if (((this.curRep.metricNms[0] != "CRT") && (this.curRep.metricNms[0] != "Calls Offered")) && (this.curRep.metricNms[0] != "Calls Handled"))
            {
                flag = false;
            }
            if ((this.curRep.grp.Count > 1) || (this.curRep.grp[0] != "Interval"))
            {
                flag = false;
            }
            this.tsCAP.Visible = flag;
            this.tsCap07.Visible = flag;
        }

        private void setForamt()
        {
            try
            {
                DataGridViewCellStyle style = new DataGridViewCellStyle(this.dGrid.DefaultCellStyle) {
                    Font = new Font(this.dGrid.Font, FontStyle.Bold),
                    Alignment = DataGridViewContentAlignment.MiddleLeft
                };
                for (int i = 0; i < this.curRep.grp.Count; i++)
                {
                    this.dGrid.Columns[i].HeaderText = this.curRep.grp[i];
                    this.dGrid.Columns[i].DefaultCellStyle = style;
                    this.dGrid.Columns[i].Frozen = true;
                }
                foreach (DataGridViewColumn column in this.dGrid.Columns)
                {
                    column.SortMode = DataGridViewColumnSortMode.Programmatic;
                    if (column.HeaderCell.Value.ToString().StartsWith("esc:"))
                    {
                        column.HeaderCell.Value = column.HeaderCell.Value.ToString().Replace("esc:", "");
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        public void setView(report rep)
        {
            try
            {
                if (rep.rendered)
                {
                    base.Visible = true;
                    this.curRep = rep;
                    this.Text = rep.name;
                    this.tbl = this.curRep.rep.Copy();
                    this.dGrid.DataSource = null;
                    this.dGrid.DataSource = this.tbl;
                    this.setCapViz();
                    this.setForamt();
                    this.active = true;
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void sortRep(string dir, int colInd)
        {
            try
            {
                this.dGrid.DataSource = null;
                if (!this.tbl.Columns[colInd].ColumnName.StartsWith("axis"))
                {
                    DataRow[] rowArray = this.tbl.Copy().Select("axis0 <> 'Total'", "[" + this.tbl.Columns[colInd].ColumnName + "] " + dir);
                    this.tbl.BeginLoadData();
                    for (int i = 0; i < rowArray.Length; i++)
                    {
                        this.tbl.Rows[i].ItemArray = rowArray[i].ItemArray;
                    }
                    this.tbl.EndLoadData();
                }
                else
                {
                    string[] destinationArray = new string[this.curRep.srts[colInd].Length];
                    Array.Copy(this.curRep.srts[colInd], destinationArray, this.curRep.srts[colInd].Length);
                    if (dir == "ASC")
                    {
                        Array.Reverse(destinationArray);
                    }
                    DataTable table = this.tbl.Copy();
                    object[] itemArray = this.tbl.Rows[this.tbl.Rows.Count - 1].ItemArray;
                    this.tbl.Clear();
                    this.tbl.BeginLoadData();
                    foreach (string str in destinationArray)
                    {
                        foreach (DataRow row in table.Select(table.Columns[colInd].ToString() + " = '" + str.Replace("'", "''") + "'"))
                        {
                            DataRow row2 = this.tbl.NewRow();
                            row2.BeginEdit();
                            row2.ItemArray = row.ItemArray;
                            row2.EndEdit();
                            this.tbl.Rows.Add(row2);
                        }
                    }
                    DataRow row3 = this.tbl.NewRow();
                    row3.BeginEdit();
                    row3.ItemArray = itemArray;
                    row3.EndEdit();
                    this.tbl.Rows.Add(row3);
                    this.tbl.EndLoadData();
                }
                this.dGrid.DataSource = this.tbl;
                if (this.connected)
                {
                    this.curRep.rep = this.tbl.Copy();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void tsCAP_Click(object sender, EventArgs e)
        {
            try
            {
                new genExcelTools();
                object app = genExcelTools.getApp();
                object sheet = genExcelTools.getSheet(genExcelTools.newWorkbook(app), 1);
                genExcelTools.CmtAzToSheet(sheet, this.tbl, this.curRep.formByCol, this.curRep.formats, errorLog.filtIllegalXlsChars(this.curRep.name));
                bool volCAP = true;
                if (this.curRep.metricNms[0] == "CRT")
                {
                    volCAP = false;
                }
                genExcelTools.addToCap(sheet, volCAP, false);
                genExcelTools.showApp(app);
            }
            catch (Exception exception)
            {
                MessageBox.Show("An error occurred while trying to create the Excel workbook." + Environment.NewLine + Environment.NewLine + exception.Message);
            }
        }

        private void tsCAP07_Click(object sender, EventArgs e)
        {
            try
            {
                new genExcelTools();
                object app = genExcelTools.getApp();
                object sheet = genExcelTools.getSheet(genExcelTools.newWorkbook(app), 1);
                genExcelTools.CmtAzToSheet(sheet, this.tbl, this.curRep.formByCol, this.curRep.formats, errorLog.filtIllegalXlsChars(this.curRep.name));
                bool volCAP = true;
                if (this.curRep.metricNms[0] == "CRT")
                {
                    volCAP = false;
                }
                genExcelTools.addToCap(sheet, volCAP, true);
                genExcelTools.showApp(app);
            }
            catch (Exception exception)
            {
                MessageBox.Show("An error occurred while trying to create the Excel workbook." + Environment.NewLine + Environment.NewLine + exception.Message);
            }
        }

        private void tsCopyAll_Click(object sender, EventArgs e)
        {
            try
            {
                StringBuilder builder = new StringBuilder();
                if (this.cmIncHedz.Checked)
                {
                    foreach (DataGridViewColumn column in this.dGrid.Columns)
                    {
                        if (column.Index > 0)
                        {
                            builder.Append("\t");
                        }
                        builder.Append(column.HeaderText);
                    }
                    builder.Append(Environment.NewLine);
                }
                foreach (DataGridViewRow row in (IEnumerable) this.dGrid.Rows)
                {
                    if (row.Index > 0)
                    {
                        builder.Append(Environment.NewLine);
                    }
                    bool flag = true;
                    foreach (DataGridViewCell cell in row.Cells)
                    {
                        if (flag)
                        {
                            flag = false;
                        }
                        else
                        {
                            builder.Append("\t");
                        }
                        builder.Append(cell.FormattedValue.ToString());
                    }
                }
                Clipboard.SetData("Text", builder.ToString());
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void tsCopyAll_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    this.cmHedz.Show(this, base.PointToClient(Cursor.Position));
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void tsCopySel_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.dGrid.SelectedCells.Count == 0)
                {
                    MessageBox.Show("You don't have anything selected.");
                }
                else
                {
                    Clipboard.SetDataObject(this.dGrid.GetClipboardContent());
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void tsCopySel_MouseDown(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    this.cmHedz.Show(this, base.PointToClient(Cursor.Position));
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void tsDecIn_Click(object sender, EventArgs e)
        {
            if (this.curRep.decimalVis != 0)
            {
                this.curRep.decimalVis = (byte) (this.curRep.decimalVis - 1);
                this.dGrid.Refresh();
            }
        }

        private void tsDecOut_Click(object sender, EventArgs e)
        {
            if (this.curRep.decimalVis <= 0x11)
            {
                this.curRep.decimalVis = (byte) (this.curRep.decimalVis + 1);
                this.dGrid.Refresh();
            }
        }

        private void tsDiscon_Click(object sender, EventArgs e)
        {
            this.disconnect();
        }

        private void tsExcel_Click(object sender, EventArgs e)
        {
            try
            {
                object workbook = genExcelTools.newWorkbook(genExcelTools.getApp());
                genExcelTools.CmtAzToSheet(genExcelTools.getSheet(workbook, 1), this.tbl, this.curRep.formByCol, this.curRep.formats, errorLog.filtIllegalXlsChars(this.curRep.name));
                genExcelTools.showExcel(workbook);
            }
            catch (Exception exception)
            {
                MessageBox.Show("An error occurred while trying to create the Excel workbook." + Environment.NewLine + Environment.NewLine + exception.Message);
            }
        }
    }
}

