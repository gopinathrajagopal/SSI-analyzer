namespace analyzer
{
    using ACD;
    using analyzer.Properties;
    using DateInterval;
    using escCoSel;
    using excelTools;
    using forecGroup;
    using Hierarchy;
    using ICM;
    using segGroup;
    using SkillGroup;
    using System;
    using System.Collections.Generic;
    using System.ComponentModel;
    using System.Data;
    using System.Drawing;
    using System.Windows.Forms;
    using VDN;

    public class main : Form
    {
        private ContextMenuStrip cmFormat;
        private ToolStripMenuItem cmFormOff;
        private ToolStripMenuItem cmFormOn;
        private ContextMenuStrip cmShowBlanksC;
        private ToolStripMenuItem cmShowBlanksCOff;
        private ToolStripMenuItem cmShowBlanksCOn;
        private ContextMenuStrip cmShowBlanksR;
        private ToolStripMenuItem cmShowBlanksROff;
        private ToolStripMenuItem cmShowBlanksROn;
        private ToolStripMenuItem cmShowZerosCOff;
        private ToolStripMenuItem cmShowZerosCOn;
        private ToolStripMenuItem cmShowZerosROff;
        private ToolStripMenuItem cmShowZerosROn;
        private ToolStripMenuItem cmTFDay;
        private ToolStripMenuItem cmTFHou;
        private ToolStripMenuItem cmTFItv;
        private ToolStripMenuItem cmTFMin;
        private ToolStripMenuItem cmTFSec;
        private ContextMenuStrip cmTimeFormz;
        private IContainer components;
        private sugarData curSug;
        private repView curView;
        private ToolStripButton dsNew;
        private ToolStripButton dsOpen;
        private ToolStripButton dsRelease;
        private ToolStripComboBox dsSelecta;
        private ACD.selector fACD;
        private DateInterval.selector fDI;
        private escCoSel.selector fEsc;
        private forecGroup.selector fForec;
        private Hierarchy.selector fHier;
        private ICM.selector fICM;
        private segGroup.selector fSeg;
        private SkillGroup.selector fSG;
        private VDN.selector fVDN;
        private ToolStripButton fwACD;
        private ToolStripButton fwDateTime;
        private ToolStripButton fwEsc;
        private ToolStripButton fwForecast;
        private ToolStripButton fwHierarchy;
        private ToolStripButton fwICM;
        private ToolStripButton fwSegGroup;
        private ToolStripButton fwSkillGroup;
        private ToolStripButton fwVDN;
        private GroupBox groupBox4;
        private GroupBox groupBox5;
        private GroupBox groupBox6;
        private GroupBox groupBox7;
        private GroupBox grpReports;
        private ListBox lstAxCol;
        private ListBox lstAxGrp;
        private ComboBox lstMetGrps;
        private ListBox lstMetrics;
        private ListBox lstReps;
        private Panel mainPan;
        private ToolStrip mainStrip;
        private bool repNmChanging;
        private ToolStrip repTools;
        private ToolStripButton rtClipboard;
        private ToolStripButton rtDelete;
        private ToolStripButton rtExcel;
        private ToolStripButton rtExcelFlag;
        private ToolStripButton rtExcelWin;
        private ToolStripButton rtRender;
        private ToolStripButton rtRenderNew;
        private bool showVis = true;
        private sessionLog SLog;
        private ToolStripStatusLabel ssMainStat;
        private ToolStripStatusLabel ssNull;
        private ToolStripDropDownButton ssVizTog;
        private StatusStrip statStrip;
        private List<sugarData> sugz = new List<sugarData>();
        private ToolStripLabel toolStripLabel1;
        private ToolStripLabel toolStripLabel2;
        private ToolStripSeparator toolStripSeparator1;
        private ToolStripSeparator toolStripSeparator2;
        private ToolStripSeparator toolStripSeparator3;
        private ToolStripSeparator toolStripSeparator4;
        private ToolStripSeparator toolStripSeparator5;
        private TextBox tRepName;
        private ToolStripMenuItem vizOff;
        private ToolStripMenuItem vizOn;

        public main(sessionLog log)
        {
            try
            {
                this.SLog = log;
                this.InitializeComponent();
                this.dsSelecta.Items.Add("Forecast");
                this.dsSelecta.Items.Add("Skills");
                this.dsSelecta.Items.Add("Agent/Skills");
                this.dsSelecta.Items.Add("Agent");
                this.dsSelecta.Items.Add("Shrinkage");
                this.dsSelecta.Items.Add("VDN");
                this.dsSelecta.Items.Add("Compliance");
                this.dsSelecta.Items.Add("ICM");
                this.dsSelecta.SelectedItem = "Skills";
                this.curView = new repView(this.SLog);
                this.ssNull.Spring = true;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void clozer(object sender, EventArgs e)
        {
            try
            {
                Form form = (Form) sender;
                bool visible = form.Visible;
                if (form.Text == "Dates")
                {
                    this.fwDateTime.Checked = visible;
                }
                if (form.Text == "Skill Groups")
                {
                    this.fwSkillGroup.Checked = visible;
                }
                if (form.Text == "ACDs")
                {
                    this.fwACD.Checked = visible;
                }
                if (form.Text == "Hierarchy")
                {
                    this.fwHierarchy.Checked = visible;
                }
                if (form.Text == "Forecast Groups")
                {
                    this.fwForecast.Checked = visible;
                }
                if (form.Text == "Shrinkage Categories")
                {
                    this.fwSegGroup.Checked = visible;
                }
                if (form.Text == "Esc. Codes")
                {
                    this.fwEsc.Checked = visible;
                }
                if (form.Text == "Call Details")
                {
                    this.fwICM.Checked = visible;
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void cmFormOff_Click(object sender, EventArgs e)
        {
            this.cmFormOn.Checked = false;
            this.cmFormOff.Checked = true;
        }

        private void cmFormOn_Click(object sender, EventArgs e)
        {
            this.cmFormOn.Checked = true;
            this.cmFormOff.Checked = false;
        }

        private void cmShowBlankCOn_Click(object sender, EventArgs e)
        {
            this.cmShowBlanksCOff.Checked = false;
            this.cmShowBlanksCOn.Checked = true;
        }

        private void cmShowBlankROn_Click(object sender, EventArgs e)
        {
            this.cmShowBlanksROff.Checked = false;
            this.cmShowBlanksROn.Checked = true;
        }

        private void cmShowBlanksCOff_Click(object sender, EventArgs e)
        {
            this.cmShowBlanksCOff.Checked = false;
            this.cmShowBlanksCOn.Checked = true;
        }

        private void cmShowBlanksROff_Click(object sender, EventArgs e)
        {
            this.cmShowBlanksROff.Checked = true;
            this.cmShowBlanksROn.Checked = false;
            this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString()).showBlankRows = false;
        }

        private void cmShowZerosCOff_Click(object sender, EventArgs e)
        {
            this.cmShowZerosCOn.Checked = false;
            this.cmShowZerosCOff.Checked = true;
        }

        private void cmShowZerosCOn_Click(object sender, EventArgs e)
        {
            this.cmShowZerosCOn.Checked = true;
            this.cmShowZerosCOff.Checked = false;
        }

        private void cmShowZerosROff_Click(object sender, EventArgs e)
        {
            this.cmShowZerosROn.Checked = false;
            this.cmShowZerosROff.Checked = true;
        }

        private void cmShowZerosROn_Click(object sender, EventArgs e)
        {
            this.cmShowZerosROn.Checked = true;
            this.cmShowZerosROff.Checked = false;
        }

        private void cmTimeFormz_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {
            string str = string.Empty;
            if (e.ClickedItem.Text == "Minutes")
            {
                str = "m";
            }
            if (e.ClickedItem.Text == "Intervals")
            {
                str = "i";
            }
            if (e.ClickedItem.Text == "Hours")
            {
                str = "h";
            }
            if (e.ClickedItem.Text == "Days")
            {
                str = "d";
            }
            string[] strArray = new string[] { this.lstMetrics.SelectedItem.ToString(), this.lstMetGrps.Text };
            DataRow row = this.curSug.metrics.Rows.Find((object[]) strArray);
            string str2 = row["format"].ToString();
            if ((str2.StartsWith("d") || str2.StartsWith("i")) || (str2.StartsWith("h") || str2.StartsWith("d")))
            {
                str2 = str2.Substring(1);
            }
            row["format"] = str + str2;
            string str3 = row["metric"].ToString();
            if (str3.Contains(" ("))
            {
                str3 = str3.Substring(0, str3.LastIndexOf(" ("));
            }
            row["metric"] = str3 + " (" + e.ClickedItem.Text + ")";
            object selectedItem = this.lstMetrics.SelectedItem;
            this.lstMetGrps_SelectedIndexChanged(this.cmTimeFormz, new EventArgs());
            this.lstMetrics.SelectedItem = selectedItem;
        }

        private void cmTimeFormz_Opening(object sender, CancelEventArgs e)
        {
            if (this.lstMetrics.SelectedItems.Count != 1)
            {
                e.Cancel = true;
            }
            else
            {
                foreach (DataRow row in this.curSug.metrics.Rows)
                {
                    if ((row["metric"].ToString() == this.lstMetrics.SelectedItem.ToString()) && (row["metGroup"].ToString() == this.lstMetGrps.Text))
                    {
                        if (!((bool) row["resTime"]))
                        {
                            break;
                        }
                        bool flag = false;
                        bool flag2 = false;
                        bool flag3 = false;
                        bool flag4 = false;
                        bool flag5 = false;
                        string str = row["format"].ToString();
                        if (str.StartsWith("m"))
                        {
                            flag2 = true;
                        }
                        if (str.StartsWith("i"))
                        {
                            flag3 = true;
                        }
                        if (str.StartsWith("h"))
                        {
                            flag4 = true;
                        }
                        if (str.StartsWith("d"))
                        {
                            flag5 = true;
                        }
                        if ((!flag2 && !flag3) && (!flag4 && !flag5))
                        {
                            flag = true;
                        }
                        this.cmTFSec.Checked = flag;
                        this.cmTFMin.Checked = flag2;
                        this.cmTFItv.Checked = flag3;
                        this.cmTFHou.Checked = flag4;
                        this.cmTFDay.Checked = flag5;
                        return;
                    }
                }
                e.Cancel = true;
            }
        }

        private string detSpclHand()
        {
            try
            {
                object[] keys = new object[] { this.lstMetrics.Items[0].ToString(), this.lstMetGrps.SelectedItem.ToString() };
                bool flag = false;
                if (this.lstAxCol.SelectedItem.ToString() == "Perspective")
                {
                    flag = true;
                }
                if (this.lstAxCol.SelectedItem.ToString() == "Forecast")
                {
                    return "forec";
                }
                foreach (object obj2 in this.lstAxGrp.SelectedItems)
                {
                    if (obj2.ToString() == "Forecast")
                    {
                        return "forec";
                    }
                    if (obj2.ToString() == "Perspective")
                    {
                        flag = true;
                    }
                }
                if ((bool) this.curSug.metrics.Rows.Find(keys)["primarySkill"])
                {
                    return "primary";
                }
                if (((this.fDI.sPersp != null) && !flag) && (this.fDI.sPersp.Select("sel = true").Length > 1))
                {
                    if ((bool) this.fDI.sPersp.Rows.Find(0)["sel"])
                    {
                        return "perspPost";
                    }
                    if ((bool) this.fDI.sPersp.Rows.Find(1)["sel"])
                    {
                        return "perspDay";
                    }
                }
                return string.Empty;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
                return null;
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

        private void doView(report rep)
        {
            try
            {
                if (this.showVis)
                {
                    if (!this.curView.connected)
                    {
                        this.curView = new repView(this.SLog);
                    }
                    if (rep.rendered)
                    {
                        if (!this.curView.defLocSet)
                        {
                            this.curView.Location = new Point(base.Location.X, base.Location.Y + base.Height);
                            this.curView.defLocSet = true;
                        }
                        this.curView.tsConnect.Click += new EventHandler(this.vwRemoteCon);
                        this.curView.setView(rep);
                        this.curView.Show();
                    }
                    else
                    {
                        this.curView.Hide();
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void dsNew_Click(object sender, EventArgs e)
        {
            try
            {
                this.Cursor = Cursors.WaitCursor;
                base.Enabled = false;
                sugarData item = new sugarData(this.dsSelecta.Text, this, this.ssMainStat, this.SLog.SID);
                this.Cursor = Cursors.Default;
                base.Enabled = true;
                if (item.loaded)
                {
                    this.dsNew.Visible = false;
                    this.dsRelease.Visible = true;
                    this.sugz.Add(item);
                    this.curSug = item;
                    string metGrp = this.curSug.metGrps[0];
                    List<string> grp = new List<string> { "Date" };
                    string[] mets = new string[] { string.Empty };
                    if (this.dsSelecta.Text == "Shrinkage")
                    {
                        this.cmShowZerosROff_Click(this, new EventArgs());
                        this.cmShowZerosCOff_Click(this, new EventArgs());
                    }
                    report rep = this.curSug.buildRep("Metric Group", grp, mets, metGrp, this.curSug.name + " default", false, false, this.cmShowZerosCOn.Checked, this.cmShowZerosROn.Checked);
                    this.curSug.repList.add(rep);
                    this.setSug();
                    this.lstReps.SetSelected(0, true);
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void dsRelease_Click(object sender, EventArgs e)
        {
            try
            {
                this.dsRelease.Visible = false;
                this.dsNew.Visible = true;
                for (int i = 0; i < this.sugz.Count; i++)
                {
                    if (this.sugz[i].name == this.curSug.name)
                    {
                        this.curSug.loaded = false;
                        this.curSug.ds.Dispose();
                        this.curSug.assocFilters.Clear();
                        this.curSug.uLog.endUsage();
                        this.sugz.RemoveAt(i);
                        this.setSug();
                        return;
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void dsSelecta_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.curSug = this.getSug(this.dsSelecta.SelectedItem.ToString());
            this.setSug();
        }

        private void enabler()
        {
            try
            {
                if (!this.curSug.loaded)
                {
                    this.grpReports.Enabled = false;
                }
                else
                {
                    this.grpReports.Enabled = true;
                    this.rtRender.Enabled = true;
                    if (this.curSug.repList.reps.Count == 1)
                    {
                        this.rtDelete.Enabled = false;
                    }
                    else
                    {
                        this.rtDelete.Enabled = true;
                    }
                    if (this.curSug.repList.reps[0].rendered)
                    {
                        this.rtRenderNew.Enabled = true;
                        this.rtClipboard.Enabled = true;
                        this.rtExcel.Enabled = true;
                        this.rtExcelFlag.Enabled = true;
                        this.rtExcelWin.Enabled = true;
                    }
                    else
                    {
                        this.rtRenderNew.Enabled = false;
                        this.rtClipboard.Enabled = false;
                        this.rtExcel.Enabled = false;
                        this.rtExcelFlag.Enabled = false;
                        this.rtExcelWin.Enabled = false;
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwACD_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.fwACD.Checked)
                {
                    this.fACD.winRefresh();
                    this.fACD.winShow();
                }
                else
                {
                    this.fACD.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwDateTime_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.fwDateTime.Checked)
                {
                    this.fDI.winRefresh();
                    this.fDI.winShow();
                }
                else
                {
                    this.fDI.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwEsc_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.fwEsc.Checked)
                {
                    this.fEsc.winRefresh();
                    this.fEsc.windShow();
                }
                else
                {
                    this.fEsc.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwForecast_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.fwForecast.Checked)
                {
                    this.fForec.winRefresh();
                    this.fForec.winShow();
                }
                else
                {
                    this.fForec.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwHierarchy_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.fwHierarchy.Checked)
                {
                    this.fHier.winRefresh();
                    this.fHier.winShow();
                }
                else
                {
                    this.fHier.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwICM_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.fwICM.Checked)
                {
                    this.fICM.winRefresh();
                    this.fICM.winShow();
                }
                else
                {
                    this.fICM.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwSegGroup_Click(object sender, EventArgs e)
        {
            try
            {
                if (this.fwSegGroup.Checked)
                {
                    this.fSeg.winRefresh();
                    this.fSeg.winShow();
                }
                else
                {
                    this.fSeg.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwSkillGroup_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.fwSkillGroup.Checked)
                {
                    this.fSG.winRefresh();
                    this.fSG.winShow();
                }
                else
                {
                    this.fSG.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void fwVDN_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                if (this.fwVDN.Checked)
                {
                    this.fVDN.winRefresh();
                    this.fVDN.winShow();
                }
                else
                {
                    this.fVDN.window.Hide();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private sugarData getSug(string sugName)
        {
            try
            {
                foreach (sugarData data in this.sugz)
                {
                    if (data.name == sugName)
                    {
                        return data;
                    }
                }
                return new sugarData(string.Empty, this, this.ssMainStat, this.SLog.SID);
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
                return null;
            }
        }

        private void InitializeComponent()
        {
            this.components = new Container();
            ComponentResourceManager manager = new ComponentResourceManager(typeof(main));
            this.grpReports = new GroupBox();
            this.repTools = new ToolStrip();
            this.rtRender = new ToolStripButton();
            this.rtRenderNew = new ToolStripButton();
            this.rtDelete = new ToolStripButton();
            this.toolStripSeparator1 = new ToolStripSeparator();
            this.rtClipboard = new ToolStripButton();
            this.rtExcel = new ToolStripButton();
            this.rtExcelFlag = new ToolStripButton();
            this.toolStripSeparator3 = new ToolStripSeparator();
            this.rtExcelWin = new ToolStripButton();
            this.groupBox7 = new GroupBox();
            this.lstMetrics = new ListBox();
            this.cmTimeFormz = new ContextMenuStrip(this.components);
            this.cmTFSec = new ToolStripMenuItem();
            this.cmTFMin = new ToolStripMenuItem();
            this.cmTFItv = new ToolStripMenuItem();
            this.cmTFHou = new ToolStripMenuItem();
            this.cmTFDay = new ToolStripMenuItem();
            this.groupBox6 = new GroupBox();
            this.lstMetGrps = new ComboBox();
            this.groupBox5 = new GroupBox();
            this.lstAxCol = new ListBox();
            this.cmShowBlanksC = new ContextMenuStrip(this.components);
            this.cmShowBlanksCOn = new ToolStripMenuItem();
            this.cmShowBlanksCOff = new ToolStripMenuItem();
            this.toolStripSeparator5 = new ToolStripSeparator();
            this.cmShowZerosCOn = new ToolStripMenuItem();
            this.cmShowZerosCOff = new ToolStripMenuItem();
            this.groupBox4 = new GroupBox();
            this.lstAxGrp = new ListBox();
            this.cmShowBlanksR = new ContextMenuStrip(this.components);
            this.cmShowBlanksROn = new ToolStripMenuItem();
            this.cmShowBlanksROff = new ToolStripMenuItem();
            this.toolStripSeparator4 = new ToolStripSeparator();
            this.cmShowZerosROn = new ToolStripMenuItem();
            this.cmShowZerosROff = new ToolStripMenuItem();
            this.tRepName = new TextBox();
            this.lstReps = new ListBox();
            this.mainStrip = new ToolStrip();
            this.toolStripLabel2 = new ToolStripLabel();
            this.dsSelecta = new ToolStripComboBox();
            this.dsNew = new ToolStripButton();
            this.dsRelease = new ToolStripButton();
            this.dsOpen = new ToolStripButton();
            this.toolStripSeparator2 = new ToolStripSeparator();
            this.toolStripLabel1 = new ToolStripLabel();
            this.fwDateTime = new ToolStripButton();
            this.fwVDN = new ToolStripButton();
            this.fwSegGroup = new ToolStripButton();
            this.fwEsc = new ToolStripButton();
            this.fwSkillGroup = new ToolStripButton();
            this.fwACD = new ToolStripButton();
            this.fwForecast = new ToolStripButton();
            this.fwICM = new ToolStripButton();
            this.fwHierarchy = new ToolStripButton();
            this.statStrip = new StatusStrip();
            this.ssMainStat = new ToolStripStatusLabel();
            this.ssNull = new ToolStripStatusLabel();
            this.ssVizTog = new ToolStripDropDownButton();
            this.vizOn = new ToolStripMenuItem();
            this.vizOff = new ToolStripMenuItem();
            this.mainPan = new Panel();
            this.cmFormat = new ContextMenuStrip(this.components);
            this.cmFormOn = new ToolStripMenuItem();
            this.cmFormOff = new ToolStripMenuItem();
            this.grpReports.SuspendLayout();
            this.repTools.SuspendLayout();
            this.groupBox7.SuspendLayout();
            this.cmTimeFormz.SuspendLayout();
            this.groupBox6.SuspendLayout();
            this.groupBox5.SuspendLayout();
            this.cmShowBlanksC.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.cmShowBlanksR.SuspendLayout();
            this.mainStrip.SuspendLayout();
            this.statStrip.SuspendLayout();
            this.mainPan.SuspendLayout();
            this.cmFormat.SuspendLayout();
            base.SuspendLayout();
            this.grpReports.Controls.Add(this.repTools);
            this.grpReports.Controls.Add(this.groupBox7);
            this.grpReports.Controls.Add(this.groupBox6);
            this.grpReports.Controls.Add(this.groupBox5);
            this.grpReports.Controls.Add(this.groupBox4);
            this.grpReports.Controls.Add(this.tRepName);
            this.grpReports.Controls.Add(this.lstReps);
            this.grpReports.Dock = DockStyle.Right;
            this.grpReports.Enabled = false;
            this.grpReports.Font = new Font("Arial", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.grpReports.ForeColor = Color.FromArgb(0xe2, 0, 0x74);
            this.grpReports.Location = new Point(0x65, 0);
            this.grpReports.Name = "grpReports";
            this.grpReports.Size = new Size(0x215, 220);
            this.grpReports.TabIndex = 2;
            this.grpReports.TabStop = false;
            this.grpReports.Text = "Reports";
            this.repTools.AutoSize = false;
            this.repTools.BackgroundImage = (Image) manager.GetObject("repTools.BackgroundImage");
            this.repTools.Dock = DockStyle.Left;
            this.repTools.GripStyle = ToolStripGripStyle.Hidden;
            this.repTools.Items.AddRange(new ToolStripItem[] { this.rtRender, this.rtRenderNew, this.rtDelete, this.toolStripSeparator1, this.rtClipboard, this.rtExcel, this.rtExcelFlag, this.toolStripSeparator3, this.rtExcelWin });
            this.repTools.Location = new Point(3, 0x11);
            this.repTools.Name = "repTools";
            this.repTools.Size = new Size(0x1a, 200);
            this.repTools.TabIndex = 4;
            this.rtRender.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtRender.Image = (Image) manager.GetObject("rtRender.Image");
            this.rtRender.ImageTransparentColor = Color.Magenta;
            this.rtRender.Name = "rtRender";
            this.rtRender.Size = new Size(0x18, 20);
            this.rtRender.ToolTipText = "Render Report";
            this.rtRender.Click += new EventHandler(this.rtRender_Click);
            this.rtRenderNew.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtRenderNew.Image = (Image) manager.GetObject("rtRenderNew.Image");
            this.rtRenderNew.ImageTransparentColor = Color.Magenta;
            this.rtRenderNew.Name = "rtRenderNew";
            this.rtRenderNew.Size = new Size(0x18, 20);
            this.rtRenderNew.ToolTipText = "Render as New Report";
            this.rtRenderNew.Click += new EventHandler(this.rtRenderNew_Click);
            this.rtDelete.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtDelete.Image = (Image) manager.GetObject("rtDelete.Image");
            this.rtDelete.ImageTransparentColor = Color.Magenta;
            this.rtDelete.Name = "rtDelete";
            this.rtDelete.Size = new Size(0x18, 20);
            this.rtDelete.Text = "toolStripButton3";
            this.rtDelete.ToolTipText = "Delete Report";
            this.rtDelete.Click += new EventHandler(this.rtDelete_Click);
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new Size(0x18, 6);
            this.rtClipboard.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtClipboard.Image = (Image) manager.GetObject("rtClipboard.Image");
            this.rtClipboard.ImageTransparentColor = Color.Magenta;
            this.rtClipboard.Name = "rtClipboard";
            this.rtClipboard.Size = new Size(0x18, 20);
            this.rtClipboard.Text = "toolStripButton4";
            this.rtClipboard.ToolTipText = "Copy report to clipboard";
            this.rtClipboard.MouseDown += new MouseEventHandler(this.showCmForm);
            this.rtClipboard.Click += new EventHandler(this.rtClipboard_Click);
            this.rtExcel.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtExcel.Image = (Image) manager.GetObject("rtExcel.Image");
            this.rtExcel.ImageTransparentColor = Color.Magenta;
            this.rtExcel.Name = "rtExcel";
            this.rtExcel.Size = new Size(0x18, 20);
            this.rtExcel.Text = "toolStripButton5";
            this.rtExcel.ToolTipText = "Export to Excel";
            this.rtExcel.Click += new EventHandler(this.rtExcel_Click);
            this.rtExcelFlag.CheckOnClick = true;
            this.rtExcelFlag.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtExcelFlag.Image = (Image) manager.GetObject("rtExcelFlag.Image");
            this.rtExcelFlag.ImageTransparentColor = Color.Magenta;
            this.rtExcelFlag.Name = "rtExcelFlag";
            this.rtExcelFlag.Size = new Size(0x18, 20);
            this.rtExcelFlag.Text = "Add to workbook list";
            this.rtExcelFlag.CheckedChanged += new EventHandler(this.rtExcelFlag_CheckedChanged);
            this.toolStripSeparator3.Name = "toolStripSeparator3";
            this.toolStripSeparator3.Size = new Size(0x18, 6);
            this.rtExcelWin.DisplayStyle = ToolStripItemDisplayStyle.Image;
            this.rtExcelWin.Image = (Image) manager.GetObject("rtExcelWin.Image");
            this.rtExcelWin.ImageTransparentColor = Color.Magenta;
            this.rtExcelWin.Name = "rtExcelWin";
            this.rtExcelWin.Size = new Size(0x18, 20);
            this.rtExcelWin.ToolTipText = "Excel Workbook Dialog";
            this.rtExcelWin.Click += new EventHandler(this.rtExcelWin_Click);
            this.groupBox7.Controls.Add(this.lstMetrics);
            this.groupBox7.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox7.ForeColor = Color.FromArgb(0, 0, 0xc0);
            this.groupBox7.Location = new Point(0x164, 0x41);
            this.groupBox7.Name = "groupBox7";
            this.groupBox7.Size = new Size(0xab, 0x97);
            this.groupBox7.TabIndex = 3;
            this.groupBox7.TabStop = false;
            this.groupBox7.Text = "Metrics";
            this.lstMetrics.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstMetrics.ContextMenuStrip = this.cmTimeFormz;
            this.lstMetrics.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstMetrics.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstMetrics.FormattingEnabled = true;
            this.lstMetrics.Location = new Point(0, 0x11);
            this.lstMetrics.Name = "lstMetrics";
            this.lstMetrics.SelectionMode = SelectionMode.MultiExtended;
            this.lstMetrics.Size = new Size(0xab, 0x86);
            this.lstMetrics.TabIndex = 0;
            this.cmTimeFormz.Items.AddRange(new ToolStripItem[] { this.cmTFSec, this.cmTFMin, this.cmTFItv, this.cmTFHou, this.cmTFDay });
            this.cmTimeFormz.Name = "cmTimeFormz";
            this.cmTimeFormz.Size = new Size(0x81, 0x72);
            this.cmTimeFormz.ItemClicked += new ToolStripItemClickedEventHandler(this.cmTimeFormz_ItemClicked);
            this.cmTimeFormz.Opening += new CancelEventHandler(this.cmTimeFormz_Opening);
            this.cmTFSec.Name = "cmTFSec";
            this.cmTFSec.Size = new Size(0x80, 0x16);
            this.cmTFSec.Text = "Seconds";
            this.cmTFMin.Name = "cmTFMin";
            this.cmTFMin.Size = new Size(0x80, 0x16);
            this.cmTFMin.Text = "Minutes";
            this.cmTFItv.Name = "cmTFItv";
            this.cmTFItv.Size = new Size(0x80, 0x16);
            this.cmTFItv.Text = "Intervals";
            this.cmTFHou.Name = "cmTFHou";
            this.cmTFHou.Size = new Size(0x80, 0x16);
            this.cmTFHou.Text = "Hours";
            this.cmTFDay.Name = "cmTFDay";
            this.cmTFDay.Size = new Size(0x80, 0x16);
            this.cmTFDay.Text = "Days";
            this.groupBox6.Controls.Add(this.lstMetGrps);
            this.groupBox6.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox6.ForeColor = Color.FromArgb(0, 0, 0xc0);
            this.groupBox6.Location = new Point(0x164, 0x11);
            this.groupBox6.Name = "groupBox6";
            this.groupBox6.Size = new Size(0xab, 0x24);
            this.groupBox6.TabIndex = 3;
            this.groupBox6.TabStop = false;
            this.groupBox6.Text = "Metric Group";
            this.lstMetGrps.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstMetGrps.DropDownStyle = ComboBoxStyle.DropDownList;
            this.lstMetGrps.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstMetGrps.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstMetGrps.FormattingEnabled = true;
            this.lstMetGrps.Location = new Point(0, 15);
            this.lstMetGrps.Name = "lstMetGrps";
            this.lstMetGrps.Size = new Size(0xab, 0x15);
            this.lstMetGrps.TabIndex = 0;
            this.lstMetGrps.SelectedIndexChanged += new EventHandler(this.lstMetGrps_SelectedIndexChanged);
            this.groupBox5.Controls.Add(this.lstAxCol);
            this.groupBox5.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox5.ForeColor = Color.FromArgb(0, 0, 0xc0);
            this.groupBox5.Location = new Point(0xb3, 40);
            this.groupBox5.Name = "groupBox5";
            this.groupBox5.Size = new Size(0x52, 0xb0);
            this.groupBox5.TabIndex = 2;
            this.groupBox5.TabStop = false;
            this.groupBox5.Text = "Columns";
            this.lstAxCol.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstAxCol.ContextMenuStrip = this.cmShowBlanksC;
            this.lstAxCol.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstAxCol.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstAxCol.FormattingEnabled = true;
            this.lstAxCol.Location = new Point(0, 0x10);
            this.lstAxCol.Name = "lstAxCol";
            this.lstAxCol.Size = new Size(0x52, 160);
            this.lstAxCol.TabIndex = 0;
            this.cmShowBlanksC.Items.AddRange(new ToolStripItem[] { this.cmShowBlanksCOn, this.cmShowBlanksCOff, this.toolStripSeparator5, this.cmShowZerosCOn, this.cmShowZerosCOff });
            this.cmShowBlanksC.Name = "cmShowBlanksC";
            this.cmShowBlanksC.Size = new Size(0xb7, 0x62);
            this.cmShowBlanksCOn.Name = "cmShowBlanksCOn";
            this.cmShowBlanksCOn.Size = new Size(0xb6, 0x16);
            this.cmShowBlanksCOn.Text = "Show Blank Columns";
            this.cmShowBlanksCOn.Click += new EventHandler(this.cmShowBlankCOn_Click);
            this.cmShowBlanksCOff.Checked = true;
            this.cmShowBlanksCOff.CheckState = CheckState.Checked;
            this.cmShowBlanksCOff.Name = "cmShowBlanksCOff";
            this.cmShowBlanksCOff.Size = new Size(0xb6, 0x16);
            this.cmShowBlanksCOff.Text = "Drop Blank Columns";
            this.cmShowBlanksCOff.Click += new EventHandler(this.cmShowBlanksCOff_Click);
            this.toolStripSeparator5.Name = "toolStripSeparator5";
            this.toolStripSeparator5.Size = new Size(0xb3, 6);
            this.cmShowZerosCOn.Checked = true;
            this.cmShowZerosCOn.CheckState = CheckState.Checked;
            this.cmShowZerosCOn.Name = "cmShowZerosCOn";
            this.cmShowZerosCOn.Size = new Size(0xb6, 0x16);
            this.cmShowZerosCOn.Text = "Show Zero Columns";
            this.cmShowZerosCOn.Click += new EventHandler(this.cmShowZerosCOn_Click);
            this.cmShowZerosCOff.Name = "cmShowZerosCOff";
            this.cmShowZerosCOff.Size = new Size(0xb6, 0x16);
            this.cmShowZerosCOff.Text = "Drop Zero Columns";
            this.cmShowZerosCOff.Click += new EventHandler(this.cmShowZerosCOff_Click);
            this.groupBox4.Controls.Add(this.lstAxGrp);
            this.groupBox4.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.groupBox4.ForeColor = Color.FromArgb(0, 0, 0xc0);
            this.groupBox4.Location = new Point(0x10c, 40);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new Size(0x52, 0xb0);
            this.groupBox4.TabIndex = 2;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "Group By";
            this.lstAxGrp.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstAxGrp.ContextMenuStrip = this.cmShowBlanksR;
            this.lstAxGrp.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstAxGrp.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstAxGrp.FormattingEnabled = true;
            this.lstAxGrp.Location = new Point(0, 0x10);
            this.lstAxGrp.Name = "lstAxGrp";
            this.lstAxGrp.SelectionMode = SelectionMode.MultiExtended;
            this.lstAxGrp.Size = new Size(0x52, 160);
            this.lstAxGrp.TabIndex = 0;
            this.cmShowBlanksR.Items.AddRange(new ToolStripItem[] { this.cmShowBlanksROn, this.cmShowBlanksROff, this.toolStripSeparator4, this.cmShowZerosROn, this.cmShowZerosROff });
            this.cmShowBlanksR.Name = "cmShowBlanks";
            this.cmShowBlanksR.Size = new Size(0xa9, 0x62);
            this.cmShowBlanksROn.Name = "cmShowBlanksROn";
            this.cmShowBlanksROn.Size = new Size(0xa8, 0x16);
            this.cmShowBlanksROn.Text = "Show Blank Rows";
            this.cmShowBlanksROn.Click += new EventHandler(this.cmShowBlankROn_Click);
            this.cmShowBlanksROff.Checked = true;
            this.cmShowBlanksROff.CheckState = CheckState.Checked;
            this.cmShowBlanksROff.Name = "cmShowBlanksROff";
            this.cmShowBlanksROff.Size = new Size(0xa8, 0x16);
            this.cmShowBlanksROff.Text = "Drop Blank Rows";
            this.cmShowBlanksROff.Click += new EventHandler(this.cmShowBlanksROff_Click);
            this.toolStripSeparator4.Name = "toolStripSeparator4";
            this.toolStripSeparator4.Size = new Size(0xa5, 6);
            this.cmShowZerosROn.Checked = true;
            this.cmShowZerosROn.CheckState = CheckState.Checked;
            this.cmShowZerosROn.Name = "cmShowZerosROn";
            this.cmShowZerosROn.Size = new Size(0xa8, 0x16);
            this.cmShowZerosROn.Text = "Show Zero Rows";
            this.cmShowZerosROn.Click += new EventHandler(this.cmShowZerosROn_Click);
            this.cmShowZerosROff.Name = "cmShowZerosROff";
            this.cmShowZerosROff.Size = new Size(0xa8, 0x16);
            this.cmShowZerosROff.Text = "Drop Zero Rows";
            this.cmShowZerosROff.Click += new EventHandler(this.cmShowZerosROff_Click);
            this.tRepName.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.tRepName.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.tRepName.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.tRepName.Location = new Point(0xb3, 0x11);
            this.tRepName.Name = "tRepName";
            this.tRepName.Size = new Size(170, 20);
            this.tRepName.TabIndex = 1;
            this.tRepName.KeyDown += new KeyEventHandler(this.tRepName_KeyDown);
            this.lstReps.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.lstReps.Font = new Font("Microsoft Sans Serif", 8.25f, FontStyle.Regular, GraphicsUnit.Point, 0);
            this.lstReps.ForeColor = Color.FromArgb(0x44, 0x44, 0x44);
            this.lstReps.FormattingEnabled = true;
            this.lstReps.Location = new Point(0x20, 0x11);
            this.lstReps.Name = "lstReps";
            this.lstReps.Size = new Size(0x8d, 0xc7);
            this.lstReps.TabIndex = 0;
            this.lstReps.SelectedIndexChanged += new EventHandler(this.lstReps_SelectedIndexChanged);
            this.mainStrip.BackColor = Color.FromArgb(0xcc, 0xcc, 0xcc);
            this.mainStrip.BackgroundImage = (Image) manager.GetObject("mainStrip.BackgroundImage");
            this.mainStrip.Dock = DockStyle.Left;
            this.mainStrip.GripStyle = ToolStripGripStyle.Hidden;
            this.mainStrip.Items.AddRange(new ToolStripItem[] { this.toolStripLabel2, this.dsSelecta, this.dsNew, this.dsRelease, this.dsOpen, this.toolStripSeparator2, this.toolStripLabel1, this.fwDateTime, this.fwVDN, this.fwSegGroup, this.fwEsc, this.fwSkillGroup, this.fwACD, this.fwForecast, this.fwICM, this.fwHierarchy });
            this.mainStrip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow;
            this.mainStrip.Location = new Point(0, 0);
            this.mainStrip.Name = "mainStrip";
            this.mainStrip.Size = new Size(0x62, 220);
            this.mainStrip.TabIndex = 0;
            this.mainStrip.Text = "toolStrip2";
            this.toolStripLabel2.Font = new Font("Arial", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.toolStripLabel2.ForeColor = Color.FromArgb(0xe2, 0, 0x74);
            this.toolStripLabel2.Name = "toolStripLabel2";
            this.toolStripLabel2.Size = new Size(0x5f, 15);
            this.toolStripLabel2.Text = "  Datasets";
            this.toolStripLabel2.TextAlign = ContentAlignment.MiddleLeft;
            this.dsSelecta.DropDownStyle = ComboBoxStyle.DropDownList;
            this.dsSelecta.FlatStyle = FlatStyle.System;
            this.dsSelecta.Name = "dsSelecta";
            this.dsSelecta.Size = new Size(0x5d, 0x15);
            this.dsSelecta.SelectedIndexChanged += new EventHandler(this.dsSelecta_SelectedIndexChanged);
            this.dsNew.ForeColor = Color.Black;
            this.dsNew.Image = (Image) manager.GetObject("dsNew.Image");
            this.dsNew.ImageAlign = ContentAlignment.MiddleLeft;
            this.dsNew.ImageTransparentColor = Color.Magenta;
            this.dsNew.Name = "dsNew";
            this.dsNew.Size = new Size(0x5f, 20);
            this.dsNew.Text = "New";
            this.dsNew.ToolTipText = "Load new data set";
            this.dsNew.Click += new EventHandler(this.dsNew_Click);
            this.dsRelease.Image = (Image) manager.GetObject("dsRelease.Image");
            this.dsRelease.ImageAlign = ContentAlignment.MiddleLeft;
            this.dsRelease.ImageTransparentColor = Color.Magenta;
            this.dsRelease.Name = "dsRelease";
            this.dsRelease.Size = new Size(0x5f, 20);
            this.dsRelease.Text = "Release";
            this.dsRelease.ToolTipText = "Release data set";
            this.dsRelease.Visible = false;
            this.dsRelease.Click += new EventHandler(this.dsRelease_Click);
            this.dsOpen.ForeColor = Color.Black;
            this.dsOpen.Image = (Image) manager.GetObject("dsOpen.Image");
            this.dsOpen.ImageAlign = ContentAlignment.MiddleLeft;
            this.dsOpen.ImageTransparentColor = Color.Magenta;
            this.dsOpen.Name = "dsOpen";
            this.dsOpen.Size = new Size(0x5f, 20);
            this.dsOpen.Text = "Open";
            this.dsOpen.Visible = false;
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new Size(0x5f, 6);
            this.toolStripLabel1.Font = new Font("Arial", 9f, FontStyle.Bold, GraphicsUnit.Point, 0);
            this.toolStripLabel1.ForeColor = Color.FromArgb(0xe2, 0, 0x74);
            this.toolStripLabel1.Name = "toolStripLabel1";
            this.toolStripLabel1.Size = new Size(0x5f, 15);
            this.toolStripLabel1.Text = "  Filter Windows";
            this.toolStripLabel1.TextAlign = ContentAlignment.MiddleLeft;
            this.fwDateTime.CheckOnClick = true;
            this.fwDateTime.ForeColor = Color.Black;
            this.fwDateTime.Image = (Image) manager.GetObject("fwDateTime.Image");
            this.fwDateTime.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwDateTime.ImageTransparentColor = Color.Magenta;
            this.fwDateTime.Name = "fwDateTime";
            this.fwDateTime.Size = new Size(0x5f, 20);
            this.fwDateTime.Text = "Date Time";
            this.fwDateTime.CheckedChanged += new EventHandler(this.fwDateTime_CheckedChanged);
            this.fwVDN.CheckOnClick = true;
            this.fwVDN.ForeColor = Color.Black;
            this.fwVDN.Image = (Image) manager.GetObject("fwVDN.Image");
            this.fwVDN.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwVDN.ImageTransparentColor = Color.Magenta;
            this.fwVDN.Name = "fwVDN";
            this.fwVDN.Size = new Size(0x5f, 20);
            this.fwVDN.Text = "VDNs";
            this.fwVDN.CheckedChanged += new EventHandler(this.fwVDN_CheckedChanged);
            this.fwSegGroup.CheckOnClick = true;
            this.fwSegGroup.ForeColor = Color.Black;
            this.fwSegGroup.Image = (Image) manager.GetObject("fwSegGroup.Image");
            this.fwSegGroup.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwSegGroup.ImageTransparentColor = Color.Magenta;
            this.fwSegGroup.Name = "fwSegGroup";
            this.fwSegGroup.Size = new Size(0x5f, 20);
            this.fwSegGroup.Text = "Segment Grps";
            this.fwSegGroup.CheckedChanged += new EventHandler(this.fwSegGroup_Click);
            this.fwEsc.CheckOnClick = true;
            this.fwEsc.ForeColor = Color.Black;
            this.fwEsc.Image = (Image) manager.GetObject("fwEsc.Image");
            this.fwEsc.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwEsc.ImageTransparentColor = Color.Magenta;
            this.fwEsc.Name = "fwEsc";
            this.fwEsc.Size = new Size(0x5f, 20);
            this.fwEsc.Text = "Esc. Codes";
            this.fwEsc.CheckedChanged += new EventHandler(this.fwEsc_Click);
            this.fwSkillGroup.CheckOnClick = true;
            this.fwSkillGroup.ForeColor = Color.Black;
            this.fwSkillGroup.Image = (Image) manager.GetObject("fwSkillGroup.Image");
            this.fwSkillGroup.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwSkillGroup.ImageTransparentColor = Color.Magenta;
            this.fwSkillGroup.Name = "fwSkillGroup";
            this.fwSkillGroup.Size = new Size(0x5f, 20);
            this.fwSkillGroup.Text = "Skill Groups";
            this.fwSkillGroup.CheckedChanged += new EventHandler(this.fwSkillGroup_CheckedChanged);
            this.fwACD.CheckOnClick = true;
            this.fwACD.ForeColor = Color.Black;
            this.fwACD.Image = (Image) manager.GetObject("fwACD.Image");
            this.fwACD.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwACD.ImageTransparentColor = Color.Magenta;
            this.fwACD.Name = "fwACD";
            this.fwACD.Size = new Size(0x35, 20);
            this.fwACD.Text = "ACDs";
            this.fwACD.CheckedChanged += new EventHandler(this.fwACD_CheckedChanged);
            this.fwForecast.CheckOnClick = true;
            this.fwForecast.ForeColor = Color.Black;
            this.fwForecast.Image = (Image) manager.GetObject("fwForecast.Image");
            this.fwForecast.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwForecast.ImageTransparentColor = Color.Magenta;
            this.fwForecast.Name = "fwForecast";
            this.fwForecast.Size = new Size(0x4a, 20);
            this.fwForecast.Text = "Forecasts";
            this.fwForecast.Click += new EventHandler(this.fwForecast_Click);
            this.fwICM.CheckOnClick = true;
            this.fwICM.ForeColor = Color.Black;
            this.fwICM.Image = (Image) manager.GetObject("fwICM.Image");
            this.fwICM.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwICM.ImageTransparentColor = Color.Magenta;
            this.fwICM.Name = "fwICM";
            this.fwICM.Size = new Size(0x51, 20);
            this.fwICM.Text = "ICM Details";
            this.fwICM.CheckedChanged += new EventHandler(this.fwICM_CheckedChanged);
            this.fwHierarchy.CheckOnClick = true;
            this.fwHierarchy.ForeColor = Color.Black;
            this.fwHierarchy.Image = (Image) manager.GetObject("fwHierarchy.Image");
            this.fwHierarchy.ImageAlign = ContentAlignment.MiddleLeft;
            this.fwHierarchy.ImageTransparentColor = Color.Magenta;
            this.fwHierarchy.Name = "fwHierarchy";
            this.fwHierarchy.Size = new Size(0x49, 20);
            this.fwHierarchy.Text = "Hierarchy";
            this.fwHierarchy.CheckedChanged += new EventHandler(this.fwHierarchy_CheckedChanged);
            this.statStrip.Items.AddRange(new ToolStripItem[] { this.ssMainStat, this.ssNull, this.ssVizTog });
            this.statStrip.Location = new Point(0, 0xde);
            this.statStrip.Name = "statStrip";
            this.statStrip.Size = new Size(0x27a, 0x16);
            this.statStrip.SizingGrip = false;
            this.statStrip.TabIndex = 3;
            this.statStrip.Text = "statusStrip1";
            this.ssMainStat.BackColor = Color.Transparent;
            this.ssMainStat.Name = "ssMainStat";
            this.ssMainStat.Size = new Size(0, 0x11);
            this.ssNull.BackColor = Color.Transparent;
            this.ssNull.Name = "ssNull";
            this.ssNull.Size = new Size(0, 0x11);
            this.ssVizTog.DropDownItems.AddRange(new ToolStripItem[] { this.vizOn, this.vizOff });
            this.ssVizTog.Image = (Image) manager.GetObject("ssVizTog.Image");
            this.ssVizTog.ImageTransparentColor = Color.Magenta;
            this.ssVizTog.Name = "ssVizTog";
            this.ssVizTog.Size = new Size(80, 20);
            this.ssVizTog.Text = "Visualizer";
            this.vizOn.Checked = true;
            this.vizOn.CheckState = CheckState.Checked;
            this.vizOn.Name = "vizOn";
            this.vizOn.Size = new Size(0x65, 0x16);
            this.vizOn.Text = "On";
            this.vizOn.Click += new EventHandler(this.vizToggle);
            this.vizOff.CheckOnClick = true;
            this.vizOff.Name = "vizOff";
            this.vizOff.Size = new Size(0x65, 0x16);
            this.vizOff.Text = "Off";
            this.vizOff.Click += new EventHandler(this.vizToggle);
            this.mainPan.Controls.Add(this.grpReports);
            this.mainPan.Controls.Add(this.mainStrip);
            this.mainPan.Dock = DockStyle.Top;
            this.mainPan.Location = new Point(0, 0);
            this.mainPan.Name = "mainPan";
            this.mainPan.Size = new Size(0x27a, 220);
            this.mainPan.TabIndex = 4;
            this.cmFormat.Items.AddRange(new ToolStripItem[] { this.cmFormOn, this.cmFormOff });
            this.cmFormat.Name = "cmFormat";
            this.cmFormat.Size = new Size(0x9a, 0x30);
            this.cmFormOn.Checked = true;
            this.cmFormOn.CheckState = CheckState.Checked;
            this.cmFormOn.Name = "cmFormOn";
            this.cmFormOn.Size = new Size(0x99, 0x16);
            this.cmFormOn.Text = "Formatting";
            this.cmFormOn.Click += new EventHandler(this.cmFormOn_Click);
            this.cmFormOff.Name = "cmFormOff";
            this.cmFormOff.Size = new Size(0x99, 0x16);
            this.cmFormOff.Text = "No Formatting";
            this.cmFormOff.Click += new EventHandler(this.cmFormOff_Click);
            base.AutoScaleDimensions = new SizeF(6f, 13f);
            base.AutoScaleMode = AutoScaleMode.Font;
            this.AutoValidate = AutoValidate.EnableAllowFocusChange;
            this.BackColor = Color.White;
            base.ClientSize = new Size(0x27a, 0xf4);
            base.Controls.Add(this.mainPan);
            base.Controls.Add(this.statStrip);
            base.FormBorderStyle = FormBorderStyle.FixedSingle;
            base.Icon = (Icon) manager.GetObject("$this.Icon");
            base.MaximizeBox = false;
            base.Name = "main";
            this.Text = "SSI Analyzer";
            base.FormClosing += new FormClosingEventHandler(this.main_FormClosing);
            this.grpReports.ResumeLayout(false);
            this.grpReports.PerformLayout();
            this.repTools.ResumeLayout(false);
            this.repTools.PerformLayout();
            this.groupBox7.ResumeLayout(false);
            this.cmTimeFormz.ResumeLayout(false);
            this.groupBox6.ResumeLayout(false);
            this.groupBox5.ResumeLayout(false);
            this.cmShowBlanksC.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.cmShowBlanksR.ResumeLayout(false);
            this.mainStrip.ResumeLayout(false);
            this.mainStrip.PerformLayout();
            this.statStrip.ResumeLayout(false);
            this.statStrip.PerformLayout();
            this.mainPan.ResumeLayout(false);
            this.mainPan.PerformLayout();
            this.cmFormat.ResumeLayout(false);
            base.ResumeLayout(false);
            base.PerformLayout();
        }

        private void lstMetGrps_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                this.lstMetrics.BeginUpdate();
                this.lstMetrics.Items.Clear();
                foreach (DataRow row in this.curSug.metrics.Select("metGroup = '" + this.lstMetGrps.Text + "'", "ordinal"))
                {
                    this.lstMetrics.Items.Add(row["metric"]);
                }
                this.lstMetrics.EndUpdate();
                this.lstMetrics.SetSelected(0, true);
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void lstReps_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (!this.repNmChanging)
                {
                    report rep = this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString());
                    if (rep != null)
                    {
                        this.tRepName.Text = rep.name;
                        for (int i = 0; i < this.lstAxGrp.Items.Count; i++)
                        {
                            if (rep.grp.Contains(this.lstAxGrp.Items[i].ToString()))
                            {
                                this.lstAxGrp.SetSelected(i, true);
                            }
                            else
                            {
                                this.lstAxGrp.SetSelected(i, false);
                            }
                        }
                        this.lstAxCol.SelectedItem = rep.col;
                        this.lstMetGrps.SelectedItem = rep.metGrpNm;
                        this.lstMetrics.BeginUpdate();
                        for (int j = 0; j < this.lstMetrics.Items.Count; j++)
                        {
                            this.lstMetrics.SetSelected(j, false);
                            for (int k = 0; k < rep.metricNms.Length; k++)
                            {
                                if (this.lstMetrics.Items[j].ToString() == rep.metricNms[k])
                                {
                                    this.lstMetrics.SetSelected(j, true);
                                }
                            }
                        }
                        this.lstMetrics.EndUpdate();
                        if (this.lstMetrics.SelectedItems.Count == 0)
                        {
                            this.lstMetrics.SetSelected(0, true);
                        }
                        if (rep.showBlankRows)
                        {
                            this.cmShowBlanksROn.Checked = true;
                            this.cmShowBlanksROff.Checked = false;
                        }
                        else
                        {
                            this.cmShowBlanksROn.Checked = false;
                            this.cmShowBlanksROff.Checked = true;
                        }
                        if (rep.showBlankCols)
                        {
                            this.cmShowBlanksCOn.Checked = true;
                            this.cmShowBlanksCOff.Checked = false;
                        }
                        else
                        {
                            this.cmShowBlanksCOn.Checked = false;
                            this.cmShowBlanksCOff.Checked = true;
                        }
                        this.curSug.swapFilters(rep.filters);
                        this.doView(rep);
                        if (this.curSug.assocFilters.Contains("Forecast"))
                        {
                            this.fForec.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("DateInterval"))
                        {
                            this.fDI.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("ACD"))
                        {
                            this.fACD.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("SkillGroup"))
                        {
                            this.fSG.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("Hierarchy"))
                        {
                            this.fHier.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("SegGroup"))
                        {
                            this.fSeg.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("EscCodes"))
                        {
                            this.fEsc.winRefresh();
                        }
                        if (this.curSug.assocFilters.Contains("ICM"))
                        {
                            this.fICM.winRefresh();
                        }
                        if (rep.xlsFlag)
                        {
                            this.rtExcelFlag.Checked = true;
                        }
                        else
                        {
                            this.rtExcelFlag.Checked = false;
                        }
                        this.curSug.repList.curSel = rep;
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void main_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                foreach (Form form in base.OwnedForms)
                {
                    form.Dispose();
                }
                foreach (sugarData data in this.sugz)
                {
                    if (data.loaded)
                    {
                        data.uLog.endUsage();
                    }
                }
                this.SLog.endSession();
                base.Dispose();
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void repChangeCommit(report rep)
        {
            try
            {
                for (int i = 0; i < this.curSug.repList.reps.Count; i++)
                {
                    if (this.curSug.repList.reps[i].staticNum == rep.staticNum)
                    {
                        this.curSug.repList.reps[i] = rep;
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void repListUpdate(string selectedRep)
        {
            try
            {
                this.lstReps.Items.Clear();
                foreach (report report in this.curSug.repList.reps)
                {
                    this.lstReps.Items.Add(report.name);
                }
                this.lstReps.SelectedItem = selectedRep;
                if (this.lstReps.Items.Count == 1)
                {
                    this.rtDelete.Enabled = false;
                }
                else
                {
                    this.rtDelete.Enabled = true;
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void rtClipboard_Click(object sender, EventArgs e)
        {
            try
            {
                this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString()).copyToClipboard(this.cmFormOn.Checked);
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void rtDelete_Click(object sender, EventArgs e)
        {
            try
            {
                string selectedRep = string.Empty;
                if (this.lstReps.SelectedIndex == (this.lstReps.Items.Count - 1))
                {
                    selectedRep = this.lstReps.Items[this.lstReps.Items.Count - 2].ToString();
                }
                else
                {
                    selectedRep = this.lstReps.Items[this.lstReps.SelectedIndex + 1].ToString();
                }
                this.curSug.repList.remove(this.lstReps.SelectedItem.ToString());
                this.repListUpdate(selectedRep);
                this.curView.setView(this.curSug.repList.findByName(selectedRep));
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void rtExcel_Click(object sender, EventArgs e)
        {
            try
            {
                object workbook = genExcelTools.newWorkbook(genExcelTools.getApp());
                object sheet = genExcelTools.getSheet(workbook, 1);
                report report = this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString());
                string sheetName = errorLog.filtIllegalXlsChars(report.name);
                genExcelTools.CmtAzToSheet(sheet, report.rep, report.formByCol, report.formats, sheetName);
                genExcelTools.showExcel(workbook);
            }
            catch (Exception exception)
            {
                MessageBox.Show("An error occurred while trying to create the Excel workbook." + Environment.NewLine + Environment.NewLine + exception.Message);
            }
        }

        private void rtExcelFlag_CheckedChanged(object sender, EventArgs e)
        {
            try
            {
                this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString()).xlsFlag = this.rtExcelFlag.Checked;
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void rtExcelWin_Click(object sender, EventArgs e)
        {
            try
            {
                new xlsBuilder(this.sugz).ShowDialog();
                if (this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString()).xlsFlag)
                {
                    this.rtExcelFlag.Checked = true;
                }
                else
                {
                    this.rtExcelFlag.Checked = false;
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void rtRender_Click(object sender, EventArgs e)
        {
            try
            {
                List<string> grp = new List<string>();
                foreach (object obj2 in this.lstAxGrp.SelectedItems)
                {
                    grp.Add(obj2.ToString());
                }
                List<string> list2 = new List<string>();
                foreach (object obj3 in this.lstMetrics.SelectedItems)
                {
                    list2.Add(obj3.ToString());
                }
                report rep = this.curSug.buildRep(this.lstAxCol.SelectedItem.ToString(), grp, list2.ToArray(), this.lstMetGrps.SelectedItem.ToString(), this.tRepName.Text, this.cmShowBlanksCOn.Checked, this.cmShowBlanksROn.Checked, this.cmShowZerosCOn.Checked, this.cmShowZerosROn.Checked);
                rep.staticNum = this.curSug.repList.nameToStatNum(this.lstReps.SelectedItem.ToString());
                rep.render(this.curSug.makeSelTbl(this.detSpclHand()));
                if (!rep.tooMuch)
                {
                    this.repChangeCommit(rep);
                    this.repListUpdate(rep.name);
                    this.doView(rep);
                    this.enabler();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void rtRenderNew_Click(object sender, EventArgs e)
        {
            try
            {
                string repName = this.curSug.repList.nameUnique(this.tRepName.Text);
                List<string> grp = new List<string>();
                foreach (object obj2 in this.lstAxGrp.SelectedItems)
                {
                    grp.Add(obj2.ToString());
                }
                List<string> list2 = new List<string>();
                foreach (object obj3 in this.lstMetrics.SelectedItems)
                {
                    list2.Add(obj3.ToString());
                }
                report rep = this.curSug.buildRep(this.lstAxCol.SelectedItem.ToString(), grp, list2.ToArray(), this.lstMetGrps.SelectedItem.ToString(), repName, this.cmShowBlanksCOn.Checked, this.cmShowBlanksROn.Checked, this.cmShowZerosCOn.Checked, this.cmShowZerosROn.Checked);
                rep.render(this.curSug.makeSelTbl(this.detSpclHand()));
                if (!rep.tooMuch)
                {
                    this.curSug.repList.add(rep);
                    this.repListUpdate(rep.name);
                    this.doView(rep);
                    this.enabler();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void setSug()
        {
            try
            {
                if (this.curSug.assocFilters.Contains("DateInterval"))
                {
                    this.fDI = this.curSug.fDI;
                    this.fDI.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwDateTime.Visible = true;
                }
                else
                {
                    if (this.fDI != null)
                    {
                        this.fDI.window.Hide();
                        this.fDI = null;
                    }
                    this.fwDateTime.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("ACD"))
                {
                    this.fACD = this.curSug.fACD;
                    this.fACD.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwACD.Visible = true;
                }
                else
                {
                    if (this.fACD != null)
                    {
                        this.fACD.window.Hide();
                        this.fACD = null;
                    }
                    this.fwACD.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("SkillGroup"))
                {
                    this.fSG = this.curSug.fSG;
                    this.fSG.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwSkillGroup.Visible = true;
                }
                else
                {
                    if (this.fSG != null)
                    {
                        this.fSG.window.Hide();
                        this.fSG = null;
                    }
                    this.fwSkillGroup.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("Hierarchy"))
                {
                    this.fHier = this.curSug.fHier;
                    this.fHier.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwHierarchy.Visible = true;
                }
                else
                {
                    if (this.fHier != null)
                    {
                        this.fHier.window.Hide();
                        this.fHier = null;
                    }
                    this.fwHierarchy.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("Forecast"))
                {
                    this.fForec = this.curSug.fForec;
                    this.fForec.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwForecast.Visible = true;
                }
                else
                {
                    if (this.fForec != null)
                    {
                        this.fForec.window.Hide();
                        this.fForec = null;
                    }
                    this.fwForecast.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("SegGroup"))
                {
                    this.fSeg = this.curSug.fSeg;
                    this.fSeg.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwSegGroup.Visible = true;
                }
                else
                {
                    if (this.fSeg != null)
                    {
                        this.fSeg.window.Hide();
                        this.fSeg = null;
                    }
                    this.fwSegGroup.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("EscCodes"))
                {
                    this.fEsc = this.curSug.fEsc;
                    this.fEsc.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwEsc.Visible = true;
                }
                else
                {
                    if (this.fEsc != null)
                    {
                        this.fEsc.window.Hide();
                        this.fEsc = null;
                    }
                    this.fwEsc.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("VDN"))
                {
                    this.fVDN = this.curSug.fVDN;
                    this.fVDN.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwVDN.Visible = true;
                }
                else
                {
                    if (this.fVDN != null)
                    {
                        this.fVDN.window.Hide();
                        this.fVDN = null;
                    }
                    this.fwVDN.Visible = false;
                }
                if (this.curSug.assocFilters.Contains("ICM"))
                {
                    this.fICM = this.curSug.fICM;
                    this.fICM.window.VisibleChanged += new EventHandler(this.clozer);
                    this.fwICM.Visible = true;
                }
                else
                {
                    if (this.fICM != null)
                    {
                        this.fICM.window.Hide();
                        this.fICM = null;
                    }
                    this.fwICM.Visible = false;
                }
                this.lstReps.BeginUpdate();
                this.lstAxCol.BeginUpdate();
                this.lstAxGrp.BeginUpdate();
                this.lstMetGrps.BeginUpdate();
                this.lstReps.Items.Clear();
                this.lstAxCol.Items.Clear();
                this.lstAxGrp.Items.Clear();
                this.lstMetGrps.Items.Clear();
                if (!this.curSug.loaded)
                {
                    this.dsNew.Visible = true;
                    this.dsRelease.Visible = false;
                    this.grpReports.Enabled = false;
                    this.lstReps.EndUpdate();
                    this.lstAxCol.EndUpdate();
                    this.lstAxGrp.EndUpdate();
                    this.lstMetrics.Items.Clear();
                    this.lstMetGrps.EndUpdate();
                    this.ssMainStat.Text = string.Empty;
                    if (((this.curView != null) && this.curView.Visible) && this.curView.connected)
                    {
                        this.curView.Hide();
                    }
                    this.enabler();
                }
                else
                {
                    this.grpReports.Enabled = true;
                    this.dsNew.Visible = false;
                    this.dsRelease.Visible = true;
                    foreach (string str in this.curSug.axes)
                    {
                        if ((str != "Metric Group") && (str != "All"))
                        {
                            this.lstAxGrp.Items.Add(str);
                        }
                        this.lstAxCol.Items.Add(str);
                    }
                    foreach (string str2 in this.curSug.metGrps)
                    {
                        this.lstMetGrps.Items.Add(str2);
                    }
                    foreach (report report in this.curSug.repList.reps)
                    {
                        this.lstReps.Items.Add(report.name);
                        if ((this.curSug.repList.curSel != null) && (report.name == this.curSug.repList.curSel.name))
                        {
                            this.lstReps.SetSelected(this.lstReps.Items.Count - 1, true);
                        }
                    }
                    this.ssMainStat.Text = this.curSug.statMsg;
                    this.lstReps.EndUpdate();
                    this.lstAxCol.EndUpdate();
                    this.lstAxGrp.EndUpdate();
                    this.lstMetrics.EndUpdate();
                    this.lstMetGrps.EndUpdate();
                    this.enabler();
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void showCmForm(object sender, MouseEventArgs e)
        {
            try
            {
                if (e.Button == MouseButtons.Right)
                {
                    this.cmFormat.Show(this, base.PointToClient(Cursor.Position));
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void tRepName_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.Enter)
                {
                    report report = this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString());
                    if (report != null)
                    {
                        this.repNmChanging = true;
                        report.name = this.tRepName.Text;
                        for (int i = 0; i < this.lstReps.Items.Count; i++)
                        {
                            if (this.lstReps.GetSelected(i))
                            {
                                this.lstReps.Items[i] = report.name;
                            }
                        }
                        this.repNmChanging = false;
                    }
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void vizToggle(object sender, EventArgs e)
        {
            try
            {
                ToolStripMenuItem item = (ToolStripMenuItem) sender;
                if (item.Text == "On")
                {
                    this.vizOn.Checked = true;
                    this.vizOff.Checked = false;
                    this.showVis = true;
                    if (this.lstReps.SelectedItem != null)
                    {
                        report rep = this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString());
                        this.doView(rep);
                    }
                    this.ssVizTog.Image = Resources.vizOn;
                }
                else
                {
                    this.vizOn.Checked = false;
                    this.vizOff.Checked = true;
                    this.showVis = false;
                    this.ssVizTog.Image = Resources.vizOff;
                }
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }

        private void vwRemoteCon(object sender, EventArgs e)
        {
            try
            {
                this.curView.disconnect();
                ToolStripButton button = (ToolStripButton) sender;
                ToolStrip owner = button.Owner;
                this.curView = (repView) owner.FindForm();
                this.curView.connect();
                report rep = this.curSug.repList.findByName(this.lstReps.SelectedItem.ToString());
                this.doView(rep);
            }
            catch (Exception exception)
            {
                errorLog.writeError(exception, this.SLog);
                base.Close();
            }
        }
    }
}

