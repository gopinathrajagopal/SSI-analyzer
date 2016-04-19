namespace analyzer
{
    using ACD;
    using DateInterval;
    using escCoSel;
    using forecGroup;
    using Hierarchy;
    using ICM;
    using segGroup;
    using SkillGroup;
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Data.SqlClient;
    using System.Drawing;
    using System.Text;
    using System.Windows.Forms;
    using VDN;

    public class sugarData
    {
        public List<string> assocFilters = new List<string>();
        public List<string> axes = new List<string>();
        private SqlConnection con = new SqlConnection(@"server=psqltdy01\tdy01;database=cmt;uid=CMT_User;pwd=Bonds756");
        public DataSet ds;
        public ACD.selector fACD;
        public DateInterval.selector fDI;
        public escCoSel.selector fEsc;
        public forecGroup.selector fForec;
        public Hierarchy.selector fHier;
        public ICM.selector fICM;
        public Form frmRef;
        public ToolStripStatusLabel frmSS;
        public segGroup.selector fSeg;
        public SkillGroup.selector fSG;
        public VDN.selector fVDN;
        public Point genStrt;
        public bool loaded;
        public string[] metGrps;
        public DataTable metrics;
        public string name;
        public reportList repList = new reportList();
        public int SID;
        public string statMsg = string.Empty;
        public usageLog uLog;

        public sugarData(string type, Form formReference, ToolStripStatusLabel ss, int sessionID)
        {
            if (!datter.connectionAvailable())
            {
                MessageBox.Show("The Analyzer is unable to establish a conntection to the CMT database at the moment.  Ensure that you are connected to the network and are on the GSM1900 domain.  If you are properly connected but continue to get this message for more than a few minutes, please reach out to the NRP Command Center to report the possible outage at 1-877-792-7286", "SSI");
            }
            else
            {
                this.frmRef = formReference;
                this.frmSS = ss;
                this.SID = sessionID;
                this.genStrt = new Point(this.frmRef.Location.X, this.frmRef.Location.Y + 0x75);
                if (type == "Forecast")
                {
                    this.newForec();
                }
                else if (type == "Skills")
                {
                    this.newSkills();
                }
                else if (type == "Agent/Skills")
                {
                    this.newAgent(true);
                }
                else if (type == "Agent")
                {
                    this.newAgent(false);
                }
                else if (type == "Shrinkage")
                {
                    this.newShrink();
                }
                else if (type == "VDN")
                {
                    this.newVDN();
                }
                else if (type == "Compliance")
                {
                    this.newCTS();
                }
                else if (type == "ICM")
                {
                    this.newICM();
                }
            }
        }

        private void alterMetricTimeName()
        {
            foreach (DataRow row in this.metrics.Rows)
            {
                if (row["format"].ToString().Substring(0, 1) == "m")
                {
                    row["metric"] = row["metric"].ToString() + " (minutes)";
                }
                if (row["format"].ToString().Substring(0, 1) == "i")
                {
                    row["metric"] = row["metric"].ToString() + " (intervals)";
                }
                if (row["format"].ToString().Substring(0, 1) == "h")
                {
                    row["metric"] = row["metric"].ToString() + " (hours)";
                }
                if (row["format"].ToString().Substring(0, 1) == "d")
                {
                    row["metric"] = row["metric"].ToString() + " (days)";
                }
            }
        }

        private void buildMets(string metGrpDBName)
        {
            this.metrics = new DataTable("metrics");
            new SqlDataAdapter("select metric, metGroup, formula, div0Check, primarySkill, format, resTime, metGroup, ordinal from vwMetGrp where " + metGrpDBName + " = 1 order by vwMetGrp.metGroup, ordinal", this.con).Fill(this.metrics);
            this.metrics.PrimaryKey = new DataColumn[] { this.metrics.Columns["metric"], this.metrics.Columns["metGroup"] };
            DataTable dataTable = new DataTable();
            new SqlDataAdapter("select Name from metGroup where " + metGrpDBName + " = 1 order by ordinal", this.con).Fill(dataTable);
            this.metGrps = new string[dataTable.Rows.Count];
            for (int i = 0; i < dataTable.Rows.Count; i++)
            {
                this.metGrps[i] = dataTable.Rows[i]["Name"].ToString();
            }
            this.alterMetricTimeName();
        }

        public report buildRep(string col, List<string> grp, string[] mets, string metGrp, string repName, bool showBlankCols, bool showBlankRows, bool showZeroCols, bool showZeroRows)
        {
            report report = new report {
                name = repName,
                sugDaddy = this,
                col = col,
                axCol = this.getAxis(col, mets[0], metGrp)
            };
            if ((mets.Length > 1) && (report.axCol.TableName != "metAx"))
            {
                report.axCol = this.compAx(report.axCol, mets, metGrp);
            }
            report.showBlankCols = showBlankCols;
            report.showBlankRows = showBlankRows;
            report.showZeroCols = showZeroCols;
            report.showZeroRows = showZeroRows;
            List<string> list = new List<string>();
            List<string> list2 = new List<string>();
            foreach (string str in grp)
            {
                switch (str)
                {
                    case "Site":
                    case "Department":
                    case "Role":
                    case "Manager":
                    case "Coach":
                    case "Rep":
                        list.Add(str);
                        list2.Add(str);
                        break;
                }
            }
            bool flag = false;
            if (list.Count > 1)
            {
                report.axGrp.Add(this.fHier.getCompoundAxis(list.ToArray()));
                flag = true;
            }
            bool flag2 = false;
            if (grp.Contains("Manufacturer") && grp.Contains("Model"))
            {
                flag2 = true;
                report.axGrp.Add(this.fICM.getMakeModelAxis());
            }
            int index = 0;
            foreach (string str2 in grp)
            {
                bool flag3 = true;
                if (((((str2 == "Site") || (str2 == "Department")) || ((str2 == "Role") || (str2 == "Manager"))) || ((str2 == "Coach") || (str2 == "Rep"))) && flag)
                {
                    flag3 = false;
                }
                if (((str2 == "Manufacturer") || (str2 == "Model")) && flag2)
                {
                    flag3 = false;
                }
                if (flag3)
                {
                    report.axGrp.Add(this.getAxis(str2, mets[0], metGrp));
                }
                if ((((str2 != "Site") && (str2 != "Department")) && ((str2 != "Role") && (str2 != "Manager"))) && ((str2 != "Coach") && (str2 != "Rep")))
                {
                    if (flag)
                    {
                        list2.Add(str2);
                    }
                    else
                    {
                        list2.Insert(index, str2);
                        index++;
                    }
                }
            }
            report.grp = list2;
            report.srts = this.makeGBSorts(grp, mets[0], metGrp);
            List<string> list3 = new List<string>();
            string[] strArray = mets;
            for (int i = 0; i < strArray.Length; i++)
            {
                list3.Add(strArray[i].ToString());
            }
            report.metricNms = list3.ToArray();
            report.metGrpNm = metGrp;
            report.metrics = this.metrics;
            report.filters = this.getFilters();
            return report;
        }

        private void buildStatMessage(int dataRowSize)
        {
            this.statMsg = this.ds.Tables["raw"].Rows.Count.ToString() + " rows. (approx ";
            long num = this.ds.Tables["raw"].Rows.Count * dataRowSize;
            double num2 = 0.0;
            if (num > 0x40000000L)
            {
                num2 = ((double) num) / 1073741824.0;
                this.statMsg = this.statMsg + num2.ToString("#0.0") + " GB)";
            }
            else if (num > 0x100000L)
            {
                num2 = ((double) num) / 1048576.0;
                this.statMsg = this.statMsg + num2.ToString("#0.0") + " MB)";
            }
            else if (num > 0x400L)
            {
                this.statMsg = this.statMsg + ((((double) num) / 1024.0)).ToString("#0.0") + " KB)";
            }
            else
            {
                this.statMsg = this.statMsg + num.ToString() + " bytes)";
            }
        }

        private DataTable compAx(DataTable oldAx, string[] selMets, string metGrp)
        {
            DataTable table = new DataTable("compAx");
            table.Columns.Add("title", System.Type.GetType("System.String"));
            table.Columns.Add("filter", System.Type.GetType("System.String"));
            table.Columns.Add("metric", System.Type.GetType("System.String"));
            foreach (DataRow row in oldAx.Rows)
            {
                foreach (string str in selMets)
                {
                    DataRow row2 = table.NewRow();
                    row2.BeginEdit();
                    row2["title"] = row["title"].ToString() + " - " + str.ToString();
                    row2["filter"] = row["filter"].ToString();
                    object[] keys = new object[] { str.ToString(), metGrp };
                    row2["metric"] = this.metrics.Rows.Find(keys)["metric"].ToString();
                    table.Rows.Add(row2);
                }
            }
            return table;
        }

        private DataTable getAxis(string axType, string metric, string metGrp)
        {
            if (axType == "Date")
            {
                return this.fDI.getDateAxis();
            }
            if (axType == "Day of Week")
            {
                return this.fDI.getDoWAxis();
            }
            if (axType == "Interval")
            {
                return this.fDI.getIntervalAxis();
            }
            if (axType == "ACD")
            {
                return this.fACD.getACDAxis();
            }
            if (axType == "Parent Site")
            {
                return this.fACD.getParentSiteAxis();
            }
            if (axType == "Forecast")
            {
                return this.fForec.getFGAxis();
            }
            if ((axType == "Skill") || (axType == "Skill Group"))
            {
                object[] keys = new object[] { metric, metGrp };
                bool flag = (bool) this.metrics.Rows.Find(keys)["primarySkill"];
                if (axType == "Skill")
                {
                    if (flag)
                    {
                        return this.fSG.getSplitPriAxis();
                    }
                    return this.fSG.getSplitAxis();
                }
                if (axType == "Skill Group")
                {
                    if (flag)
                    {
                        return this.fSG.getSkillGroupPriAxis();
                    }
                    return this.fSG.getSkillGroupAxis();
                }
            }
            if (axType == "Site")
            {
                return this.fHier.getSiteAxis();
            }
            if (axType == "Department")
            {
                return this.fHier.getDeptAxis();
            }
            if (axType == "Role")
            {
                return this.fHier.getRoleAxis();
            }
            if (axType == "Manager")
            {
                return this.fHier.getMgrAxis();
            }
            if (axType == "Coach")
            {
                return this.fHier.getCoachAxis();
            }
            if (axType == "Rep")
            {
                return this.fHier.getRepAxis();
            }
            if (axType == "Segment Group")
            {
                return this.fSeg.getSegGrpAxis();
            }
            if (axType == "Segment")
            {
                return this.fSeg.getSegCodeAxis();
            }
            if (axType == "Esc. Code")
            {
                return this.fEsc.getEscCoAxis();
            }
            if (axType == "Perspective")
            {
                return this.fDI.getSPerspAxis();
            }
            if (axType == "VDN")
            {
                return this.fVDN.getVdnAx();
            }
            if (axType == "Queue")
            {
                return this.fVDN.getQueueAx();
            }
            if (axType == "Transfer")
            {
                return this.fVDN.getXferAx();
            }
            if (axType == "Vector")
            {
                return this.fVDN.getVectorAx(this.ds.Tables["raw"]);
            }
            if (axType == "Account Type")
            {
                return this.fICM.getAcctTypeAxis();
            }
            if (axType == "Area Code")
            {
                return this.fICM.getAreaCodeAxis();
            }
            if (axType == "Split")
            {
                return this.fICM.getSplitAxis();
            }
            if (axType == "Exit State")
            {
                return this.fICM.getExitStateAxis();
            }
            if (axType == "Manufacturer")
            {
                return this.fICM.getManuAxis();
            }
            if (axType == "Model")
            {
                return this.fICM.getModelAxis();
            }
            if (axType == "Metric Group")
            {
                DataTable table = new DataTable("metAx");
                table.Columns.Add("title", System.Type.GetType("System.String"));
                table.Columns.Add("expression", System.Type.GetType("System.String"));
                foreach (DataRow row in this.metrics.Select("metGroup = '" + metGrp + "'", "ordinal"))
                {
                    DataRow row2 = table.NewRow();
                    row2.BeginEdit();
                    row2["title"] = row["metric"].ToString();
                    row2["expression"] = row["formula"].ToString();
                    row2.EndEdit();
                    table.Rows.Add(row2);
                }
                return table;
            }
            DataTable table2 = new DataTable("selAx");
            table2.Columns.Add("title", System.Type.GetType("System.String"));
            table2.Columns.Add("filter", System.Type.GetType("System.String"));
            DataRow row3 = table2.NewRow();
            row3.BeginEdit();
            row3["title"] = "All";
            row3["filter"] = "true";
            row3.EndEdit();
            table2.Rows.Add(row3);
            return table2;
        }

        private DataSet getFilters()
        {
            DataSet set = new DataSet("filters");
            foreach (DataTable table in this.ds.Tables)
            {
                if (table.TableName != "raw")
                {
                    set.Tables.Add(table.Copy());
                }
            }
            return set;
        }

        public List<string[]> makeGBSorts(List<string> grp, string metric, string metGrp)
        {
            List<string[]> list = new List<string[]>();
            foreach (string str in grp)
            {
                switch (str)
                {
                    case "Date":
                        list.Add(this.fDI.getDateSort());
                        break;

                    case "Day of Week":
                        list.Add(this.fDI.getDoWSort());
                        break;

                    case "Interval":
                        list.Add(this.fDI.getIntervalSort());
                        break;

                    case "ACD":
                        list.Add(this.fACD.getACDSort());
                        break;

                    case "Parent Site":
                        list.Add(this.fACD.getParentSiteSort());
                        break;

                    case "Forecast":
                        list.Add(this.fForec.getFGSort());
                        break;

                    case "Skill":
                    {
                        object[] keys = new object[] { metric, metGrp };
                        if ((bool) this.metrics.Rows.Find(keys)["primarySkill"])
                        {
                            list.Add(this.fSG.getSplitPriSort());
                        }
                        else
                        {
                            list.Add(this.fSG.getSplitSort());
                        }
                        break;
                    }
                }
                if (str == "Skill Group")
                {
                    object[] objArray2 = new object[] { metric, metGrp };
                    if ((bool) this.metrics.Rows.Find(objArray2)["primarySkill"])
                    {
                        list.Add(this.fSG.getSkillGroupPriSort());
                    }
                    else
                    {
                        list.Add(this.fSG.getSkillGroupSort());
                    }
                }
                if (str == "Site")
                {
                    list.Add(this.fHier.getSiteSort());
                }
                if (str == "Department")
                {
                    list.Add(this.fHier.getDeptSort());
                }
                if (str == "Role")
                {
                    list.Add(this.fHier.getRoleSort());
                }
                if (str == "Manager")
                {
                    list.Add(this.fHier.getManagerSort());
                }
                if (str == "Coach")
                {
                    list.Add(this.fHier.getCoachSort());
                }
                if (str == "Rep")
                {
                    list.Add(this.fHier.getRepSort());
                }
                if (str == "Segment Group")
                {
                    list.Add(this.fSeg.getSegGrpSort());
                }
                if (str == "Segment")
                {
                    list.Add(this.fSeg.getSegCodeSort());
                }
                if (str == "Esc. Code")
                {
                    list.Add(this.fEsc.getEscCoSort());
                }
                if (str == "Perspective")
                {
                    list.Add(this.fDI.getSPerspSort());
                }
                if (str == "VDN")
                {
                    list.Add(this.fVDN.getVdnSort());
                }
                if (str == "Transfer")
                {
                    list.Add(this.fVDN.getXferSort());
                }
                if (str == "Queue")
                {
                    list.Add(this.fVDN.getQueueSort());
                }
                if (str == "Vector")
                {
                    list.Add(this.fVDN.getVectorSort(this.ds.Tables["raw"]));
                }
                if (str == "All")
                {
                    string[] item = new string[] { "All" };
                    list.Add(item);
                }
            }
            return list;
        }

        public DataTable makeSelTbl(string specialHandling)
        {
            string filterExpression = string.Empty;
            if (this.name == "Forecast")
            {
                if (specialHandling == "forec")
                {
                    filterExpression = "parent(dates).sel = true and parent(forec).sel = true ";
                }
                else
                {
                    filterExpression = "parent(dates).sel = true and parent(forec).hiSel = true ";
                }
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true ";
                }
            }
            else if (this.name == "Skills")
            {
                filterExpression = "parent(dates).sel = true and parent(ACDs).sel = true ";
                if (specialHandling == "primary")
                {
                    filterExpression = filterExpression + "and ((parent(splitKey).sel = true and parent(splitKey).pri = true)";
                    string str2 = this.fSG.getManSel(true);
                    if (str2 == string.Empty)
                    {
                        filterExpression = filterExpression + ")";
                    }
                    else
                    {
                        filterExpression = filterExpression + " or " + str2 + ")";
                    }
                }
                else
                {
                    filterExpression = filterExpression + "and ((parent(splitKey).sel = true and parent(splitKey).manEnt = false)";
                    string str3 = this.fSG.getManSel(false);
                    if (str3 == string.Empty)
                    {
                        filterExpression = filterExpression + ")";
                    }
                    else
                    {
                        filterExpression = filterExpression + " or " + str3 + ")";
                    }
                }
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true ";
                }
            }
            else if (this.name == "Agent/Skills")
            {
                filterExpression = "parent(dates).sel = true and parent(hier).sel = true and parent(hier).rProf = true and parent(hier).dProf = true and ((parent(splitKey).sel = true and parent(splitKey).manEnt = false)";
                string str4 = this.fSG.getManSel(false);
                if (str4 == string.Empty)
                {
                    filterExpression = filterExpression + ")";
                }
                else
                {
                    filterExpression = filterExpression + " or " + str4 + ")";
                }
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true ";
                }
            }
            else if (this.name == "Agent")
            {
                filterExpression = "parent(dates).sel = true and parent(hier).sel = true and parent(hier).rProf = true and parent(hier).dProf = true ";
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true ";
                }
            }
            else if (this.name == "Shrinkage")
            {
                filterExpression = "parent(dates).sel = true and parent(sPersp).sel = true and parent(hier).sel = true and parent(hier).rProf = true and parent(hier).dProf = true and parent(segCode).sel = true and parent(escCo).sel = true ";
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + "and parent(interval).sel = true";
                }
                if (specialHandling == "perspPost")
                {
                    filterExpression = filterExpression + " and persp = 0";
                }
                if (specialHandling == "perspDay")
                {
                    filterExpression = filterExpression + " and persp = 1";
                }
            }
            else if (this.name == "VDN")
            {
                filterExpression = "parent(dates).sel = true and parent(ACDs).sel = true and parent(VDN).sel = true ";
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true ";
                }
            }
            else if (this.name == "Compliance")
            {
                filterExpression = "parent(dates).sel = true and parent(hier).sel = true and parent(hier).rProf = true and parent(hier).dProf = true and parent(escCo).sel = true ";
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true ";
                }
            }
            else if (this.name == "ICM")
            {
                filterExpression = "parent(dates).sel = true and parent(acctType).sel = true ";
                if (this.fICM.usesAreaCode)
                {
                    filterExpression = filterExpression + "and parent(areaCode).sel = true ";
                }
                if (this.fICM.usesEsS)
                {
                    filterExpression = filterExpression + "and parent(exitState).sel = true and parent(split).sel = true ";
                }
                if (this.fICM.usesMakeModel)
                {
                    filterExpression = filterExpression + "and parent(makeModel).sel = true ";
                }
                if (this.fDI.grain == 0)
                {
                    filterExpression = filterExpression + " and parent(interval).sel = true";
                }
            }
            DataTable table = this.ds.Tables["raw"].Clone();
            table.BeginLoadData();
            foreach (DataRow row in this.ds.Tables["raw"].Select(filterExpression))
            {
                DataRow row2 = table.NewRow();
                row2.BeginEdit();
                row2.ItemArray = row.ItemArray;
                row2.EndEdit();
                table.Rows.Add(row2);
            }
            return table;
        }

        private DataTable makeTable(string query)
        {
            return this.makeTable(query, this.con);
        }

        private DataTable makeTable(string query, bool view)
        {
            DataTable table = this.makeTable(query);
            if (view)
            {
                foreach (DataRow row in table.Rows)
                {
                    row.SetModified();
                }
            }
            return table;
        }

        private DataTable makeTable(string query, SqlConnection altCon)
        {
            this.frmRef.Cursor = Cursors.WaitCursor;
            this.frmSS.Text = "Querying the CMT...";
            this.frmRef.Update();
            SqlDataAdapter adapter = new SqlDataAdapter(query, altCon);
            DataTable dataTable = new DataTable();
            adapter.SelectCommand.CommandTimeout = 180;
            adapter.Fill(dataTable);
            this.frmSS.Text = "";
            this.frmRef.Cursor = Cursors.Default;
            return dataTable;
        }

        private DataTable makeTable(string query, string keyColumn)
        {
            DataTable table = this.makeTable(query);
            table.PrimaryKey = new DataColumn[] { table.Columns[keyColumn] };
            return table;
        }

        private DataTable makeTable(string query, bool view, SqlConnection altCon)
        {
            DataTable table = this.makeTable(query, altCon);
            if (view)
            {
                foreach (DataRow row in table.Rows)
                {
                    row.SetModified();
                }
            }
            return table;
        }

        private DataTable makeTable(string query, string keyColumn, bool view)
        {
            DataTable table = this.makeTable(query, view);
            table.PrimaryKey = new DataColumn[] { table.Columns[keyColumn] };
            return table;
        }

        private void newAgent(bool withSkill)
        {
            int dataRowSize = 0;
            string str = string.Empty;
            if (withSkill)
            {
                this.name = "Agent/Skills";
            }
            else
            {
                this.name = "Agent";
            }
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, true, true);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int startDate = this.fDI.dateFrom();
                int endDate = this.fDI.dateTo();
                byte xpdOption = 0;
                string vwAppend = "D";
                if (this.fDI.grain == 0)
                {
                    vwAppend = "Itv";
                }
                if (this.fDI.grain == 2)
                {
                    vwAppend = "W";
                }
                if (this.fDI.grain == 3)
                {
                    vwAppend = "M";
                }
                string str3 = "Unac";
                if (withSkill)
                {
                    str3 = "Split";
                }
                Hierarchy.initialSelect select2 = new Hierarchy.initialSelect(this.con, startDate, endDate, xpdOption, this.genStrt, "vwAg" + str3, vwAppend);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fHier = select2.derivedSel;
                    this.fHier.window.Owner = this.frmRef;
                    SkillGroup.initialSelect select3 = null;
                    if (withSkill)
                    {
                        select3 = new SkillGroup.initialSelect(this.con, startDate, endDate, true, this.genStrt);
                        if (select3.result == DialogResult.Cancel)
                        {
                            return;
                        }
                        this.fSG = select3.derivedSel;
                        this.fSG.window.Owner = this.frmRef;
                    }
                    StringBuilder builder = new StringBuilder("select * from ");
                    if (withSkill)
                    {
                        str = "vwAgSplit" + select2.vwType;
                        dataRowSize = 160;
                    }
                    else
                    {
                        str = "vwAgUnac" + select2.vwType;
                        dataRowSize = 180;
                    }
                    builder.Append(str);
                    builder.Append(" where " + select.CMT_Query);
                    builder.Append(" and " + select2.CMT_Query);
                    if (withSkill)
                    {
                        builder.Append(" and " + select3.CMT_Query);
                    }
                    DataTable table = this.makeTable(builder.ToString(), true);
                    if (table.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no data for the combination of parameters that you have selected.");
                    }
                    else
                    {
                        table.TableName = "raw";
                        this.ds = new DataSet("Agent");
                        if (withSkill)
                        {
                            this.ds.DataSetName = "Agent/Skills";
                        }
                        this.ds.Tables.Add(table);
                        this.ds.Tables.Add(this.fDI.dates);
                        this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                        if (this.fDI.grain == 0)
                        {
                            this.ds.Tables.Add(this.fDI.intervals);
                            this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                        }
                        if (this.fHier.xpdLevel < 3)
                        {
                            foreach (DataRow row in table.Select("mgr is null"))
                            {
                                row["mgr"] = "Unknown";
                            }
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            foreach (DataRow row2 in table.Select("coach is null"))
                            {
                                row2["coach"] = "Unknown";
                            }
                        }
                        this.ds.Tables.Add(this.fHier.hier);
                        List<DataColumn> list = new List<DataColumn> {
                            this.ds.Tables["hier"].Columns["deptID"],
                            this.ds.Tables["hier"].Columns["roleID"],
                            this.ds.Tables["hier"].Columns["rootID"],
                            this.ds.Tables["hier"].Columns["siteID"]
                        };
                        if (this.fHier.xpdLevel < 3)
                        {
                            list.Add(this.ds.Tables["hier"].Columns["mgrID"]);
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            list.Add(this.ds.Tables["hier"].Columns["coachID"]);
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            list.Add(this.ds.Tables["hier"].Columns["fk_emp"]);
                        }
                        List<DataColumn> list2 = new List<DataColumn> {
                            this.ds.Tables["raw"].Columns["deptID"],
                            this.ds.Tables["raw"].Columns["roleID"],
                            this.ds.Tables["raw"].Columns["rootID"],
                            this.ds.Tables["raw"].Columns["siteID"]
                        };
                        if (this.fHier.xpdLevel < 3)
                        {
                            list2.Add(this.ds.Tables["raw"].Columns["mgrID"]);
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            list2.Add(this.ds.Tables["raw"].Columns["coachID"]);
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            list2.Add(this.ds.Tables["raw"].Columns["fk_emp"]);
                        }
                        this.ds.Relations.Add("hier", list.ToArray(), list2.ToArray());
                        if (withSkill)
                        {
                            this.ds.Tables.Add(this.fSG.splits);
                            this.ds.Tables.Add(this.fSG.skillGroups);
                            this.splitKeyConsolidation();
                            this.ds.Relations.Add("splitKey", this.ds.Tables["split"].Columns["pk_split"], this.ds.Tables["raw"].Columns["fk_split"]);
                        }
                        this.axes.Add("Date");
                        if (this.fDI.grain < 2)
                        {
                            this.axes.Add("Day of Week");
                        }
                        if (this.fDI.grain == 0)
                        {
                            this.axes.Add("Interval");
                        }
                        this.axes.Add("Site");
                        this.axes.Add("Department");
                        this.axes.Add("Role");
                        if (this.fHier.xpdLevel < 3)
                        {
                            this.axes.Add("Manager");
                            dataRowSize += 20;
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            this.axes.Add("Coach");
                            dataRowSize += 20;
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            this.axes.Add("Rep");
                            dataRowSize += 20;
                        }
                        if (withSkill)
                        {
                            this.axes.Add("Skill");
                            this.axes.Add("Skill Group");
                        }
                        this.axes.Add("Metric Group");
                        this.axes.Add("All");
                        if (withSkill)
                        {
                            this.buildMets("agentSkill");
                        }
                        else
                        {
                            this.buildMets("agent");
                        }
                        this.assocFilters.Add("DateInterval");
                        this.assocFilters.Add("Hierarchy");
                        if (withSkill)
                        {
                            this.assocFilters.Add("SkillGroup");
                        }
                        this.uLog = new usageLog(this.SID, str, this.ds.Tables["raw"].Rows.Count);
                        this.buildStatMessage(dataRowSize);
                        this.frmSS.Text = this.statMsg;
                        this.loaded = true;
                    }
                }
            }
        }

        private void newCTS()
        {
            int dataRowSize = 0x6a;
            string viewName = string.Empty;
            this.name = "Compliance";
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, true, true);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int startDate = this.fDI.dateFrom();
                int endDate = this.fDI.dateTo();
                string vwAppend = "D";
                if (this.fDI.grain == 0)
                {
                    vwAppend = "Itv";
                }
                if (this.fDI.grain == 2)
                {
                    vwAppend = "W";
                }
                if (this.fDI.grain == 3)
                {
                    vwAppend = "M";
                }
                Hierarchy.initialSelect select2 = new Hierarchy.initialSelect(this.con, startDate, endDate, 0, this.genStrt, "vwCTS", vwAppend);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fHier = select2.derivedSel;
                    this.fHier.window.Owner = this.frmRef;
                    viewName = "vwCTS" + select2.vwType;
                    string query = "select * from vwCTS" + select2.vwType + " where " + select.CMT_Query + " and " + select2.CMT_Query;
                    DataTable table = this.makeTable(query, true);
                    if (table.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no data for the combination of parameters that you have selected.");
                    }
                    else
                    {
                        table.TableName = "raw";
                        this.ds = new DataSet("CTS");
                        this.ds.Tables.Add(table);
                        escCoSel.initialSelect select3 = new escCoSel.initialSelect(this.con);
                        this.fEsc = select3.derivedSel;
                        this.fEsc.window.Owner = this.frmRef;
                        this.ds.Tables.Add(this.fDI.dates);
                        this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                        if (this.fDI.grain == 0)
                        {
                            this.ds.Tables.Add(this.fDI.intervals);
                            this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                        }
                        this.ds.Tables.Add(this.fEsc.escCo);
                        this.ds.Relations.Add("escCo", this.ds.Tables["escCo"].Columns["code"], this.ds.Tables["raw"].Columns["code"]);
                        this.ds.Tables.Add(this.fHier.hier);
                        List<DataColumn> list = new List<DataColumn>();
                        List<DataColumn> list2 = new List<DataColumn> {
                            this.ds.Tables["hier"].Columns["deptid"]
                        };
                        list.Add(this.ds.Tables["raw"].Columns["deptid"]);
                        list2.Add(this.ds.Tables["hier"].Columns["roleid"]);
                        list.Add(this.ds.Tables["raw"].Columns["roleid"]);
                        list2.Add(this.ds.Tables["hier"].Columns["rootid"]);
                        list.Add(this.ds.Tables["raw"].Columns["rootid"]);
                        list2.Add(this.ds.Tables["hier"].Columns["siteid"]);
                        list.Add(this.ds.Tables["raw"].Columns["siteid"]);
                        if (this.fHier.xpdLevel < 3)
                        {
                            list2.Add(this.ds.Tables["hier"].Columns["mgrid"]);
                            list.Add(this.ds.Tables["raw"].Columns["mgrid"]);
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            list2.Add(this.ds.Tables["hier"].Columns["coachid"]);
                            list.Add(this.ds.Tables["raw"].Columns["coachid"]);
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            list2.Add(this.ds.Tables["hier"].Columns["fk_emp"]);
                            list.Add(this.ds.Tables["raw"].Columns["fk_emp"]);
                        }
                        this.ds.Relations.Add("hier", list2.ToArray(), list.ToArray());
                        this.axes.Add("Date");
                        if (this.fDI.grain < 2)
                        {
                            this.axes.Add("Day of Week");
                        }
                        if (this.fDI.grain == 0)
                        {
                            this.axes.Add("Interval");
                        }
                        this.axes.Add("Site");
                        this.axes.Add("Department");
                        this.axes.Add("Role");
                        if (this.fHier.xpdLevel < 3)
                        {
                            this.axes.Add("Manager");
                            dataRowSize += 20;
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            this.axes.Add("Coach");
                            dataRowSize += 20;
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            this.axes.Add("Rep");
                            dataRowSize += 20;
                        }
                        this.axes.Add("Esc. Code");
                        this.axes.Add("Metric Group");
                        this.axes.Add("All");
                        this.buildMets("cts");
                        this.assocFilters.Add("DateInterval");
                        this.assocFilters.Add("Hierarchy");
                        this.assocFilters.Add("EscCodes");
                        this.uLog = new usageLog(this.SID, viewName, this.ds.Tables["raw"].Rows.Count);
                        this.buildStatMessage(dataRowSize);
                        this.frmSS.Text = this.statMsg;
                        this.loaded = true;
                    }
                }
            }
        }

        private void newForec()
        {
            int dataRowSize = 0;
            this.name = "Forecast";
            string str = string.Empty;
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, false, false);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int startDate = this.fDI.dateFrom();
                int endDate = this.fDI.dateTo();
                forecGroup.initialSelect select2 = new forecGroup.initialSelect(this.con, startDate, endDate, this.genStrt);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fForec = select2.derivedSel;
                    this.fForec.window.Owner = this.frmRef;
                    StringBuilder builder = new StringBuilder("select * from ");
                    if (this.fDI.grain == 0)
                    {
                        str = "vwForGrpISplit";
                        dataRowSize = 0x85;
                    }
                    else
                    {
                        str = "vwForGrpDSplit";
                        dataRowSize = 0x84;
                    }
                    builder.Append(str);
                    builder.Append(" where " + select.CMT_Query);
                    builder.Append(" and " + select2.getCMT_Query());
                    DataTable table = this.makeTable(builder.ToString(), true);
                    if (table.Rows.Count == 0)
                    {
                        MessageBox.Show("There's no data found for the combination of forecasts and dates that you have selected.");
                    }
                    else
                    {
                        table.TableName = "raw";
                        this.ds = new DataSet("Forecast");
                        this.ds.Tables.Add(table);
                        this.ds.Tables.Add(this.fDI.dates);
                        this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                        if (this.fDI.grain == 0)
                        {
                            this.ds.Tables.Add(this.fDI.intervals);
                            this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                        }
                        this.ds.Tables.Add(this.fForec.FG);
                        this.ds.Relations.Add("forec", this.ds.Tables["forec"].Columns["pk_fg"], this.ds.Tables["raw"].Columns["fk_fg"]);
                        this.axes.Add("Date");
                        if (this.fDI.grain < 2)
                        {
                            this.axes.Add("Day of Week");
                        }
                        if (this.fDI.grain == 0)
                        {
                            this.axes.Add("Interval");
                        }
                        this.axes.Add("Forecast");
                        this.axes.Add("Metric Group");
                        this.axes.Add("All");
                        this.buildMets("forecast");
                        this.assocFilters.Add("DateInterval");
                        this.assocFilters.Add("Forecast");
                        this.uLog = new usageLog(this.SID, str, this.ds.Tables["raw"].Rows.Count);
                        this.buildStatMessage(dataRowSize);
                        this.frmSS.Text = this.statMsg;
                        this.loaded = true;
                    }
                }
            }
        }

        private void newICM()
        {
            int dataRowSize = 13;
            string viewName = string.Empty;
            this.name = "ICM";
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, true, true);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int sDate = this.fDI.dateFrom();
                int eDate = this.fDI.dateTo();
                string dateGrain = "D";
                if (this.fDI.grain == 0)
                {
                    dateGrain = "Itv";
                }
                if (this.fDI.grain == 2)
                {
                    dateGrain = "W";
                }
                if (this.fDI.grain == 3)
                {
                    dateGrain = "M";
                }
                SqlConnection cmtCon = new SqlConnection(@"server=psqltdy01\tdy01;database=careCallsLog;uid=CMT_User;pwd=Bonds756");
                ICM.initialSelect select2 = new ICM.initialSelect(cmtCon, sDate, eDate, dateGrain, this.genStrt);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fICM = select2.derivedSel;
                    string query = "select * from " + select2.vwName + " where " + select.CMT_Query + select2.whereSql;
                    DataTable table = this.makeTable(query, true, cmtCon);
                    if (table.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no data for the combination of parameters that you have selected.");
                    }
                    else
                    {
                        table.TableName = "raw";
                        this.ds = new DataSet("ICM");
                        this.ds.Tables.Add(table);
                        this.ds.Tables.Add(this.fDI.dates);
                        this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                        if (this.fDI.grain == 0)
                        {
                            this.ds.Tables.Add(this.fDI.intervals);
                            this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                        }
                        this.ds.Tables.Add(this.fICM.tblAcctType);
                        this.ds.Relations.Add("acctType", this.ds.Tables["acctType"].Columns["fk_acctType"], this.ds.Tables["raw"].Columns["fk_acctType"]);
                        if (this.fICM.usesAreaCode)
                        {
                            this.ds.Tables.Add(this.fICM.tblAreaCode);
                            this.ds.Relations.Add("areaCode", this.ds.Tables["areaCode"].Columns["areacode"], this.ds.Tables["raw"].Columns["areacode"]);
                        }
                        if (this.fICM.usesEsS)
                        {
                            this.ds.Tables.Add(this.fICM.tblExitState);
                            this.ds.Relations.Add("exitState", this.ds.Tables["exitState"].Columns["fk_exitState"], this.ds.Tables["raw"].Columns["fk_exitState"]);
                            this.ds.Tables.Add(this.fICM.tblSplit);
                            this.ds.Relations.Add("split", this.ds.Tables["split"].Columns["split"], this.ds.Tables["raw"].Columns["split"]);
                        }
                        if (this.fICM.usesMakeModel)
                        {
                            this.ds.Tables.Add(this.fICM.tblMakeModel);
                            this.ds.Relations.Add("makeModel", this.ds.Tables["makeModel"].Columns["fk_careMakeModel"], this.ds.Tables["raw"].Columns["fk_careMakeModel"]);
                        }
                        this.axes.Add("Date");
                        if (this.fDI.grain < 2)
                        {
                            this.axes.Add("Day of Week");
                        }
                        if (this.fDI.grain == 0)
                        {
                            this.axes.Add("Interval");
                        }
                        this.axes.Add("Account Type");
                        if (this.fICM.usesAreaCode)
                        {
                            this.axes.Add("Area Code");
                        }
                        if (this.fICM.usesEsS)
                        {
                            this.axes.Add("Split");
                            this.axes.Add("Exit State");
                        }
                        if (this.fICM.usesMakeModel)
                        {
                            this.axes.Add("Manufacturer");
                            this.axes.Add("Model");
                        }
                        this.axes.Add("Metric Group");
                        this.axes.Add("All");
                        if (select2.vwName.ToLower().Contains("sivr"))
                        {
                            this.buildMets("ICMSivr");
                            viewName = "vwSIVRCalls";
                        }
                        else
                        {
                            this.buildMets("ICMCare");
                            viewName = "vwCareCalls";
                        }
                        this.assocFilters.Add("DateInterval");
                        this.assocFilters.Add("ICM");
                        this.uLog = new usageLog(this.SID, viewName, this.ds.Tables["raw"].Rows.Count);
                        this.buildStatMessage(dataRowSize);
                        this.frmSS.Text = this.statMsg;
                        this.loaded = true;
                    }
                }
            }
        }

        private void newShrink()
        {
            int dataRowSize = 0x6a;
            string viewName = string.Empty;
            this.name = "Shrinkage";
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, true, true, true);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int startDate = this.fDI.dateFrom();
                int endDate = this.fDI.dateTo();
                string vwBaseName = "vwShrink";
                if (this.fDI.grain == 0)
                {
                    vwBaseName = vwBaseName + "Itv";
                }
                if (this.fDI.grain == 1)
                {
                    vwBaseName = vwBaseName + "Day";
                }
                if (this.fDI.grain == 2)
                {
                    vwBaseName = vwBaseName + "Week";
                }
                if (this.fDI.grain == 3)
                {
                    vwBaseName = vwBaseName + "Mon";
                }
                Hierarchy.initialSelect select2 = new Hierarchy.initialSelect(this.con, startDate, endDate, 0, this.genStrt, vwBaseName, (bool) this.fDI.sPersp.Rows.Find(0)["sel"], (bool) this.fDI.sPersp.Rows.Find(1)["sel"], (bool) this.fDI.sPersp.Rows.Find(2)["sel"]);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fHier = select2.derivedSel;
                    this.fHier.window.Owner = this.frmRef;
                    viewName = vwBaseName + select2.vwType + "X";
                    string str3 = string.Empty;
                    if ((bool) this.fDI.sPersp.Rows.Find(0)["sel"])
                    {
                        str3 = "select * from " + vwBaseName + select2.vwType + "P";
                    }
                    if ((bool) this.fDI.sPersp.Rows.Find(1)["sel"])
                    {
                        if (str3 != string.Empty)
                        {
                            str3 = str3 + " union ";
                        }
                        string str5 = str3;
                        str3 = str5 + "select * from " + vwBaseName + select2.vwType + "D";
                    }
                    if ((bool) this.fDI.sPersp.Rows.Find(2)["sel"])
                    {
                        if (str3 != string.Empty)
                        {
                            str3 = str3 + " union ";
                        }
                        string str6 = str3;
                        str3 = str6 + "select * from " + vwBaseName + select2.vwType + "W";
                    }
                    string query = "select * from (" + str3 + ") tbl where " + select.CMT_Query + " and " + select2.CMT_Query;
                    DataTable table = this.makeTable(query, true);
                    if (table.Rows.Count == 0)
                    {
                        MessageBox.Show("There is no data for the combination of parameters that you have selected.");
                    }
                    else
                    {
                        table.TableName = "raw";
                        List<DataColumn> list = new List<DataColumn>();
                        if (this.fDI.grain == 0)
                        {
                            list.Add(table.Columns["interval"]);
                        }
                        list.Add(table.Columns["date"]);
                        list.Add(table.Columns["fk_segCode"]);
                        list.Add(table.Columns["fk_segGroup"]);
                        list.Add(table.Columns["persp"]);
                        list.Add(table.Columns["code"]);
                        list.Add(table.Columns["siteID"]);
                        list.Add(table.Columns["siteParID"]);
                        list.Add(table.Columns["deptID"]);
                        list.Add(table.Columns["roleID"]);
                        list.Add(table.Columns["rootID"]);
                        list.Add(table.Columns["RSH"]);
                        this.ds = new DataSet("Shrink");
                        this.ds.Tables.Add(table);
                        escCoSel.initialSelect select3 = new escCoSel.initialSelect(this.con);
                        this.fEsc = select3.derivedSel;
                        this.fEsc.window.Owner = this.frmRef;
                        segGroup.initialSelect select4 = new segGroup.initialSelect(this.con, startDate, endDate, table);
                        this.fSeg = select4.derivedSel;
                        this.fSeg.window.Owner = this.frmRef;
                        this.ds.Tables.Add(this.fDI.dates);
                        this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                        if (this.fDI.grain == 0)
                        {
                            this.ds.Tables.Add(this.fDI.intervals);
                            this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                        }
                        this.ds.Tables.Add(this.fDI.sPersp);
                        this.ds.Relations.Add("sPersp", this.ds.Tables["sPersp"].Columns["persp"], this.ds.Tables["raw"].Columns["persp"]);
                        this.ds.Tables.Add(this.fEsc.escCo);
                        this.ds.Relations.Add("escCo", this.ds.Tables["escCo"].Columns["code"], this.ds.Tables["raw"].Columns["code"]);
                        this.ds.Tables.Add(this.fSeg.segCode);
                        this.ds.Relations.Add("segCode", this.ds.Tables["segCode"].Columns["pk_segCode"], this.ds.Tables["raw"].Columns["fk_segCode"]);
                        this.ds.Tables.Add(this.fSeg.segGroup);
                        this.ds.Tables.Add(this.fHier.hier);
                        List<DataColumn> list2 = new List<DataColumn>();
                        List<DataColumn> list3 = new List<DataColumn> {
                            this.ds.Tables["hier"].Columns["deptid"]
                        };
                        list2.Add(this.ds.Tables["raw"].Columns["deptid"]);
                        list3.Add(this.ds.Tables["hier"].Columns["roleid"]);
                        list2.Add(this.ds.Tables["raw"].Columns["roleid"]);
                        list3.Add(this.ds.Tables["hier"].Columns["rootid"]);
                        list2.Add(this.ds.Tables["raw"].Columns["rootid"]);
                        list3.Add(this.ds.Tables["hier"].Columns["siteid"]);
                        list2.Add(this.ds.Tables["raw"].Columns["siteid"]);
                        if (this.fHier.xpdLevel < 3)
                        {
                            list3.Add(this.ds.Tables["hier"].Columns["mgrid"]);
                            list2.Add(this.ds.Tables["raw"].Columns["mgrid"]);
                            list.Add(table.Columns["mgrID"]);
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            list3.Add(this.ds.Tables["hier"].Columns["coachid"]);
                            list2.Add(this.ds.Tables["raw"].Columns["coachid"]);
                            list.Add(table.Columns["coachID"]);
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            list3.Add(this.ds.Tables["hier"].Columns["fk_emp"]);
                            list2.Add(this.ds.Tables["raw"].Columns["fk_emp"]);
                            list.Add(table.Columns["fk_emp"]);
                        }
                        this.ds.Relations.Add("hier", list3.ToArray(), list2.ToArray());
                        this.axes.Add("Date");
                        if (this.fDI.grain < 2)
                        {
                            this.axes.Add("Day of Week");
                        }
                        if (this.fDI.grain == 0)
                        {
                            this.axes.Add("Interval");
                        }
                        this.axes.Add("Site");
                        this.axes.Add("Department");
                        this.axes.Add("Role");
                        if (this.fHier.xpdLevel < 3)
                        {
                            this.axes.Add("Manager");
                            dataRowSize += 20;
                        }
                        if (this.fHier.xpdLevel < 2)
                        {
                            this.axes.Add("Coach");
                            dataRowSize += 20;
                        }
                        if (this.fHier.xpdLevel == 0)
                        {
                            this.axes.Add("Rep");
                            dataRowSize += 20;
                        }
                        this.axes.Add("Segment Group");
                        this.axes.Add("Segment");
                        this.axes.Add("Esc. Code");
                        if (this.fDI.sPersp.Select("sel = true").Length > 1)
                        {
                            this.axes.Add("Perspective");
                        }
                        this.axes.Add("Metric Group");
                        this.axes.Add("All");
                        this.buildMets("shrink");
                        this.assocFilters.Add("DateInterval");
                        this.assocFilters.Add("Hierarchy");
                        this.assocFilters.Add("SegGroup");
                        this.assocFilters.Add("EscCodes");
                        this.uLog = new usageLog(this.SID, viewName, this.ds.Tables["raw"].Rows.Count);
                        this.buildStatMessage(dataRowSize);
                        this.frmSS.Text = this.statMsg;
                        this.loaded = true;
                    }
                }
            }
        }

        private void newSkills()
        {
            int dataRowSize = 0;
            this.name = "Skills";
            string str = string.Empty;
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, true, true);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int startDate = this.fDI.dateFrom();
                int endDate = this.fDI.dateTo();
                ACD.initialSelect select2 = new ACD.initialSelect(this.con, startDate, endDate, this.genStrt);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fACD = select2.derivedSel;
                    this.fACD.window.Owner = this.frmRef;
                    SkillGroup.initialSelect select3 = new SkillGroup.initialSelect(this.con, startDate, endDate, false, this.genStrt);
                    if (select3.result != DialogResult.Cancel)
                    {
                        this.fSG = select3.derivedSel;
                        this.fSG.window.Owner = this.frmRef;
                        StringBuilder builder = new StringBuilder("select * from ");
                        if (this.fDI.grain == 0)
                        {
                            str = "vwKeyISplit";
                            dataRowSize = 120;
                        }
                        else if (this.fDI.grain == 1)
                        {
                            str = "vwKeyDSplit ";
                            dataRowSize = 0x77;
                        }
                        else if (this.fDI.grain == 2)
                        {
                            str = "vwKeyWSplit ";
                            dataRowSize = 0x77;
                        }
                        else if (this.fDI.grain == 3)
                        {
                            str = "vwKeyMSplit ";
                            dataRowSize = 0x77;
                        }
                        builder.Append(str);
                        builder.Append(" where " + select.CMT_Query);
                        builder.Append(" and " + select2.CMT_Query);
                        builder.Append(" and " + select3.CMT_Query);
                        DataTable table = this.makeTable(builder.ToString(), true);
                        if (table.Rows.Count == 0)
                        {
                            MessageBox.Show("There is no data for the combination of dates, skills, and ACD's that you have selected.");
                        }
                        else
                        {
                            table.TableName = "raw";
                            this.ds = new DataSet("Skills");
                            this.ds.Tables.Add(table);
                            this.ds.Tables.Add(this.fDI.dates);
                            this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                            if (this.fDI.grain == 0)
                            {
                                this.ds.Tables.Add(this.fDI.intervals);
                                this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                            }
                            this.ds.Tables.Add(this.fACD.ACDs);
                            this.ds.Relations.Add("ACDs", this.ds.Tables["ACDs"].Columns["pk_acd"], this.ds.Tables["raw"].Columns["fk_acd"]);
                            this.ds.Tables.Add(this.fSG.splits);
                            this.ds.Tables.Add(this.fSG.skillGroups);
                            this.splitKeyConsolidation();
                            this.ds.Relations.Add("splitKey", this.ds.Tables["split"].Columns["pk_split"], this.ds.Tables["raw"].Columns["fk_split"]);
                            this.axes.Add("Date");
                            if (this.fDI.grain < 2)
                            {
                                this.axes.Add("Day of Week");
                            }
                            if (this.fDI.grain == 0)
                            {
                                this.axes.Add("Interval");
                            }
                            this.axes.Add("Parent Site");
                            this.axes.Add("ACD");
                            this.axes.Add("Skill");
                            this.axes.Add("Skill Group");
                            this.axes.Add("Metric Group");
                            this.axes.Add("All");
                            this.buildMets("skill");
                            this.assocFilters.Add("ACD");
                            this.assocFilters.Add("DateInterval");
                            this.assocFilters.Add("SkillGroup");
                            this.uLog = new usageLog(this.SID, str, this.ds.Tables["raw"].Rows.Count);
                            this.buildStatMessage(dataRowSize);
                            this.frmSS.Text = this.statMsg;
                            this.loaded = true;
                        }
                    }
                }
            }
        }

        private void newVDN()
        {
            int dataRowSize = 0;
            this.name = "VDN";
            string str = string.Empty;
            DateInterval.initialSelect select = new DateInterval.initialSelect(this.genStrt, true, true, true, true);
            if (select.result != DialogResult.Cancel)
            {
                this.fDI = select.derivedSel;
                this.fDI.window.Owner = this.frmRef;
                int startDate = this.fDI.dateFrom();
                int endDate = this.fDI.dateTo();
                ACD.initialSelect select2 = new ACD.initialSelect(this.con, startDate, endDate, this.genStrt, false);
                if (select2.result != DialogResult.Cancel)
                {
                    this.fACD = select2.derivedSel;
                    this.fACD.window.Owner = this.frmRef;
                    VDN.initialSelect select3 = new VDN.initialSelect(this.con, startDate, endDate, this.genStrt);
                    if (select3.result != DialogResult.Cancel)
                    {
                        this.fVDN = select3.derivedSel;
                        this.fVDN.window.Owner = this.frmRef;
                        StringBuilder builder = new StringBuilder("select * from ");
                        if (this.fDI.grain == 0)
                        {
                            str = "vwVDNi ";
                            dataRowSize = 120;
                        }
                        else if (this.fDI.grain == 1)
                        {
                            str = "vwVDNd ";
                            dataRowSize = 0x77;
                        }
                        else if (this.fDI.grain == 2)
                        {
                            str = "vwVDNw ";
                            dataRowSize = 0x77;
                        }
                        else if (this.fDI.grain == 3)
                        {
                            str = "vwVDNm ";
                            dataRowSize = 0x77;
                        }
                        builder.Append(str);
                        builder.Append(" where " + select.CMT_Query);
                        builder.Append(" and " + select2.CMT_Query);
                        builder.Append(" and " + select3.CMT_Query);
                        DataTable table = this.makeTable(builder.ToString(), true);
                        if (table.Rows.Count == 0)
                        {
                            MessageBox.Show("There is no data for the combination of dates, VDNs, and ACDs that you have selected.");
                        }
                        else
                        {
                            table.TableName = "raw";
                            this.ds = new DataSet("VDN");
                            this.ds.Tables.Add(table);
                            this.ds.Tables.Add(this.fDI.dates);
                            this.ds.Relations.Add("dates", this.ds.Tables["dates"].Columns["date"], this.ds.Tables["raw"].Columns["date"]);
                            if (this.fDI.grain == 0)
                            {
                                this.ds.Tables.Add(this.fDI.intervals);
                                this.ds.Relations.Add("interval", this.ds.Tables["intervals"].Columns["interval"], this.ds.Tables["raw"].Columns["interval"]);
                            }
                            this.ds.Tables.Add(this.fACD.ACDs);
                            this.ds.Relations.Add("ACDs", this.ds.Tables["ACDs"].Columns["pk_acd"], this.ds.Tables["raw"].Columns["fk_acd"]);
                            this.ds.Tables.Add(this.fVDN.vdns);
                            List<DataColumn> list = new List<DataColumn> {
                                this.ds.Tables["vdn"].Columns["vdn"],
                                this.ds.Tables["vdn"].Columns["queue"],
                                this.ds.Tables["vdn"].Columns["xfer"]
                            };
                            List<DataColumn> list2 = new List<DataColumn> {
                                this.ds.Tables["raw"].Columns["vdn"],
                                this.ds.Tables["raw"].Columns["skill1"],
                                this.ds.Tables["raw"].Columns["skillXfer"]
                            };
                            this.ds.Relations.Add("vdn", list.ToArray(), list2.ToArray());
                            this.axes.Add("Date");
                            if (this.fDI.grain < 2)
                            {
                                this.axes.Add("Day of Week");
                            }
                            if (this.fDI.grain == 0)
                            {
                                this.axes.Add("Interval");
                            }
                            this.axes.Add("ACD");
                            this.axes.Add("VDN");
                            this.axes.Add("Queue");
                            this.axes.Add("Transfer");
                            this.axes.Add("Vector");
                            this.axes.Add("Metric Group");
                            this.axes.Add("All");
                            this.buildMets("VDN");
                            this.assocFilters.Add("ACD");
                            this.assocFilters.Add("DateInterval");
                            this.assocFilters.Add("VDN");
                            this.uLog = new usageLog(this.SID, str, this.ds.Tables["raw"].Rows.Count);
                            this.buildStatMessage(dataRowSize);
                            this.frmSS.Text = this.statMsg;
                            this.loaded = true;
                        }
                    }
                }
            }
        }

        private void splitKeyConsolidation()
        {
            DataRow[] rowArray = this.ds.Tables["split"].Select("manEnt = true");
            if (rowArray.Length != 0)
            {
                int num = Convert.ToInt32(rowArray[0]["pk_split"]);
                StringBuilder builder = new StringBuilder();
                foreach (DataRow row in this.ds.Tables["split"].Rows)
                {
                    if (builder.Length > 1)
                    {
                        builder.Append(" and ");
                    }
                    builder.Append("fk_split <> " + row["pk_split"]);
                }
                foreach (DataRow row2 in this.ds.Tables["raw"].Select(builder.ToString()))
                {
                    row2["fk_split"] = num;
                }
            }
        }

        public void swapFilters(DataSet filters)
        {
            foreach (DataTable table in filters.Tables)
            {
                string tableName = table.TableName;
                if (this.ds.Tables.Contains(tableName))
                {
                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        this.ds.Tables[tableName].Rows[i].BeginEdit();
                        this.ds.Tables[tableName].Rows[i].ItemArray = table.Rows[i].ItemArray;
                        this.ds.Tables[tableName].Rows[i].EndEdit();
                    }
                    foreach (object obj2 in table.ExtendedProperties.Keys)
                    {
                        this.ds.Tables[tableName].ExtendedProperties[obj2] = table.ExtendedProperties[obj2];
                    }
                }
            }
        }
    }
}

