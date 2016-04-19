namespace analyzer
{
    using System;
    using System.Collections.Generic;
    using System.Data;
    using System.Text;
    using System.Windows.Forms;

    public class report
    {
        public DataTable axCol;
        public List<DataTable> axGrp = new List<DataTable>();
        public string col = string.Empty;
        public byte decimalVis = 2;
        public DataSet filters;
        public string[] formats;
        public bool formByCol;
        public List<string> grp = new List<string>();
        public string metGrpNm = string.Empty;
        public string[] metricNms;
        public DataTable metrics;
        public string name = string.Empty;
        public bool rendered;
        public DataTable rep;
        public bool showBlankCols;
        public bool showBlankRows;
        public bool showZeroCols;
        public bool showZeroRows;
        public List<string[]> srts = new List<string[]>();
        public int staticNum;
        public sugarData sugDaddy;
        public bool tooMuch;
        public bool xlsFlag;
        public string xlsName = string.Empty;

        private DataTable buildAxRow(DataTable selected)
        {
            if (this.axGrp.Count == 1)
            {
                return this.axGrp[0];
            }
            double num = 1.0;
            foreach (DataTable table in this.axGrp)
            {
                num *= table.Rows.Count;
            }
            this.tooMuch = false;
            if ((num > 100000.0) && (MessageBox.Show("The combination of 'Group By' items that you selected may take a while a will be a strain on you computer.  Do you want to continue?", "SSI Analyzer", MessageBoxButtons.OKCancel) == DialogResult.Cancel))
            {
                this.tooMuch = true;
                return null;
            }
            foreach (DataTable table2 in this.axGrp)
            {
                table2.AcceptChanges();
                foreach (DataColumn column in table2.Columns)
                {
                    if (column.ColumnName.StartsWith("title"))
                    {
                        foreach (DataRow row in table2.Rows)
                        {
                            if (row[column].ToString() == "Total")
                            {
                                row.Delete();
                            }
                        }
                        table2.AcceptChanges();
                    }
                }
            }
            DataTable baseAx = this.axGrp[0];
            for (int i = 1; i < this.axGrp.Count; i++)
            {
                baseAx = this.explodeAx(baseAx, this.axGrp[i], selected);
            }
            if (!this.grp.Contains("Forecast") && !this.grp.Contains("Perspective"))
            {
                DataRow row2 = baseAx.NewRow();
                row2.BeginEdit();
                for (int j = 1; j < (row2.ItemArray.Length - 1); j++)
                {
                    row2.ItemArray[j] = "";
                }
                row2[0] = "Total";
                row2[row2.ItemArray.Length - 1] = "true";
                baseAx.Rows.Add(row2);
            }
            return baseAx;
        }

        public void copyToClipboard(bool useFormat)
        {
            if (this.rendered)
            {
                StringBuilder builder = new StringBuilder();
                foreach (DataColumn column in this.rep.Columns)
                {
                    if (column.Ordinal > 0)
                    {
                        builder.Append("\t");
                    }
                    if (column.Caption != string.Empty)
                    {
                        builder.Append(column.Caption);
                    }
                    else
                    {
                        builder.Append(column.ColumnName);
                    }
                }
                foreach (DataRow row in this.rep.Rows)
                {
                    builder.Append(Environment.NewLine);
                    bool flag = true;
                    int num = this.rep.Columns.Count - this.formats.Length;
                    for (int i = 0; i < this.rep.Columns.Count; i++)
                    {
                        if (flag)
                        {
                            flag = false;
                        }
                        else
                        {
                            builder.Append("\t");
                        }
                        if (this.rep.Columns[i].ColumnName.StartsWith("axis") || !useFormat)
                        {
                            builder.Append(row[i].ToString());
                        }
                        else if (row[i].ToString() != string.Empty)
                        {
                            double num3 = double.Parse(row[i].ToString());
                            if (this.formByCol)
                            {
                                builder.Append(num3.ToString(this.formats[i - num]));
                            }
                            else
                            {
                                builder.Append(num3.ToString(this.formats[0]));
                            }
                        }
                    }
                }
                Clipboard.SetData("Text", builder.ToString());
            }
        }

        private void dropNullAxEnts(DataTable ax, DataTable selected)
        {
            if (ax.TableName != "metAx")
            {
                ax.AcceptChanges();
                foreach (DataRow row in ax.Rows)
                {
                    if (selected.Select(row["filter"].ToString()).Length == 0)
                    {
                        row.Delete();
                    }
                }
                ax.AcceptChanges();
            }
        }

        private void dropZeroAxEnts(DataTable ax, DataTable selected)
        {
            if (ax.TableName != "metAx")
            {
                string expression = string.Empty;
                string filter = string.Empty;
                string str3 = string.Empty;
                object[] keys = new object[2];
                keys[1] = this.metGrpNm;
                ax.AcceptChanges();
                foreach (DataRow row in ax.Rows)
                {
                    if (ax.TableName == "compAx")
                    {
                        keys[0] = row["metric"].ToString();
                    }
                    else
                    {
                        keys[0] = this.metricNms[0];
                    }
                    filter = row["filter"].ToString();
                    expression = this.metrics.Rows.Find(keys)["div0Check"].ToString();
                    str3 = this.metrics.Rows.Find(keys)["formula"].ToString();
                    bool flag = false;
                    if (expression == string.Empty)
                    {
                        expression = str3;
                        flag = true;
                    }
                    object obj2 = selected.Compute(expression, filter);
                    if ((obj2 != DBNull.Value) && (Convert.ToDouble(obj2) > 0.0))
                    {
                        if (!flag)
                        {
                            object obj3 = selected.Compute(str3, filter);
                            if ((obj3 == DBNull.Value) || (Convert.ToDouble(obj3) == 0.0))
                            {
                                row.Delete();
                            }
                        }
                    }
                    else
                    {
                        row.Delete();
                    }
                }
                ax.AcceptChanges();
            }
        }

        private DataTable explodeAx(DataTable baseAx, DataTable newAx, DataTable selected)
        {
            DataTable table = new DataTable("selAx");
            foreach (DataColumn column in baseAx.Columns)
            {
                if (column.ColumnName != "filter")
                {
                    table.Columns.Add("title" + table.Columns.Count.ToString(), System.Type.GetType("System.String"));
                }
            }
            foreach (DataColumn column2 in newAx.Columns)
            {
                if (column2.ColumnName != "filter")
                {
                    table.Columns.Add("title" + table.Columns.Count.ToString(), System.Type.GetType("System.String"));
                }
            }
            table.Columns.Add("filter", System.Type.GetType("System.String"));
            foreach (DataRow row in baseAx.Rows)
            {
                foreach (DataRow row2 in newAx.Rows)
                {
                    if (selected.Select("(" + row["filter"].ToString() + ") and (" + row2["filter"].ToString() + ")").Length > 0)
                    {
                        DataRow row3 = table.NewRow();
                        row3.BeginEdit();
                        for (int i = 0; i < (baseAx.Columns.Count - 1); i++)
                        {
                            row3[i] = row[i].ToString();
                        }
                        for (int j = 0; j < (newAx.Columns.Count - 1); j++)
                        {
                            row3[(j + baseAx.Columns.Count) - 1] = row2[j].ToString();
                        }
                        row3["filter"] = "(" + row["filter"].ToString() + ") and (" + row2["filter"].ToString() + ")";
                        row3.EndEdit();
                        table.Rows.Add(row3);
                    }
                }
            }
            return table;
        }

        private double preFormat(double v, string formz)
        {
            if (formz != string.Empty)
            {
                if (formz.Substring(0, 1) == "m")
                {
                    return (v / 60.0);
                }
                if (formz.Substring(0, 1) == "i")
                {
                    return (v / 1800.0);
                }
                if (formz.Substring(0, 1) == "h")
                {
                    return (v / 3600.0);
                }
                if (formz.Substring(0, 1) == "d")
                {
                    return (v / 86400.0);
                }
            }
            return v;
        }

        public void render(DataTable selected)
        {
            if (((this.axCol != null) && (this.axGrp.Count != 0)) && (selected != null))
            {
                DataTable ax = this.buildAxRow(selected);
                if (!this.tooMuch)
                {
                    if (!this.showBlankRows)
                    {
                        this.dropNullAxEnts(ax, selected);
                    }
                    if (!this.showBlankCols)
                    {
                        this.dropNullAxEnts(this.axCol, selected);
                    }
                    if (!this.showZeroRows)
                    {
                        this.dropZeroAxEnts(ax, selected);
                    }
                    if (!this.showZeroCols)
                    {
                        this.dropZeroAxEnts(this.axCol, selected);
                    }
                    if (this.axCol.Rows.Count >= 0x100)
                    {
                        MessageBox.Show("The report that you're trying to load contains too many columns." + Environment.NewLine + "Try using the Group-by categories to show details rather than the columns.");
                        this.tooMuch = true;
                    }
                    else
                    {
                        this.rep = new DataTable("report");
                        this.rep.BeginLoadData();
                        for (int i = 0; i < this.grp.Count; i++)
                        {
                            DataColumn column = new DataColumn("axis" + i.ToString(), System.Type.GetType("System.String")) {
                                Caption = this.grp[i]
                            };
                            this.rep.Columns.Add(column);
                        }
                        foreach (DataRow row in this.axCol.Rows)
                        {
                            this.rep.Columns.Add(row["title"].ToString(), System.Type.GetType("System.Double"));
                        }
                        string expression = string.Empty;
                        string str2 = string.Empty;
                        string filter = string.Empty;
                        string formz = string.Empty;
                        object[] keys = new object[2];
                        keys[1] = this.metGrpNm;
                        DataView view = new DataView(selected);
                        foreach (DataRow row2 in ax.Rows)
                        {
                            DataRow row3 = this.rep.NewRow();
                            row3.BeginEdit();
                            for (int k = 0; k < (ax.Columns.Count - 1); k++)
                            {
                                row3[k] = row2[k].ToString();
                            }
                            view.RowFilter = row2["filter"].ToString();
                            DataTable table2 = view.ToTable();
                            foreach (DataRow row4 in this.axCol.Rows)
                            {
                                if (row4.Table.TableName == "metAx")
                                {
                                    keys[0] = row4["title"].ToString();
                                    expression = this.metrics.Rows.Find(keys)["formula"].ToString();
                                    str2 = this.metrics.Rows.Find(keys)["div0Check"].ToString();
                                    filter = "true";
                                    formz = this.metrics.Rows.Find(keys)["format"].ToString();
                                }
                                else if (row2.Table.TableName == "metAx")
                                {
                                    keys[0] = row2["title"].ToString();
                                    expression = this.metrics.Rows.Find(keys)["formula"].ToString();
                                    str2 = this.metrics.Rows.Find(keys)["div0Check"].ToString();
                                    filter = row4["filter"].ToString();
                                    formz = this.metrics.Rows.Find(keys)["format"].ToString();
                                }
                                else if (row4.Table.TableName == "compAx")
                                {
                                    keys[0] = row4["metric"].ToString();
                                    expression = this.metrics.Rows.Find(keys)["formula"].ToString();
                                    str2 = this.metrics.Rows.Find(keys)["div0Check"].ToString();
                                    filter = "(" + row4["filter"] + ")";
                                    formz = this.metrics.Rows.Find(keys)["format"].ToString();
                                }
                                else
                                {
                                    keys[0] = this.metricNms[0];
                                    expression = this.metrics.Rows.Find(keys)["formula"].ToString();
                                    str2 = this.metrics.Rows.Find(keys)["div0Check"].ToString();
                                    filter = "(" + row4["filter"] + ")";
                                    formz = this.metrics.Rows.Find(keys)["format"].ToString();
                                }
                                bool flag = false;
                                if (str2 == string.Empty)
                                {
                                    str2 = expression;
                                    flag = true;
                                }
                                object obj2 = table2.Compute(str2, filter);
                                if (obj2 != DBNull.Value)
                                {
                                    double v = Convert.ToDouble(obj2);
                                    if (flag)
                                    {
                                        row3[row4["title"].ToString()] = this.preFormat(v, formz);
                                    }
                                    else if (v != 0.0)
                                    {
                                        object obj3 = table2.Compute(expression, filter);
                                        if (obj3 != DBNull.Value)
                                        {
                                            row3[row4["title"].ToString()] = this.preFormat(Convert.ToDouble(obj3), formz);
                                        }
                                        else
                                        {
                                            row3[row4["title"].ToString()] = 0;
                                        }
                                    }
                                }
                            }
                            row3.EndEdit();
                            this.rep.Rows.Add(row3);
                        }
                        this.rep.EndLoadData();
                        if (this.axCol.TableName == "metAx")
                        {
                            this.formByCol = true;
                            this.formats = new string[this.axCol.Rows.Count];
                            for (int m = 0; m < this.axCol.Rows.Count; m++)
                            {
                                keys[0] = this.axCol.Rows[m]["title"].ToString();
                                this.formats[m] = this.metrics.Rows.Find(keys)["format"].ToString();
                            }
                        }
                        else if (this.axCol.TableName == "compAx")
                        {
                            this.formByCol = true;
                            this.formats = new string[this.axCol.Rows.Count];
                            for (int n = 0; n < this.axCol.Rows.Count; n++)
                            {
                                keys[0] = this.axCol.Rows[n]["metric"].ToString();
                                this.formats[n] = this.metrics.Rows.Find(keys)["format"].ToString();
                            }
                        }
                        else
                        {
                            this.formByCol = false;
                            this.formats = new string[1];
                            keys[0] = this.metricNms[0];
                            this.formats[0] = this.metrics.Rows.Find(keys)["format"].ToString();
                        }
                        for (int j = 0; j < this.formats.Length; j++)
                        {
                            this.formats[j] = this.formats[j].Replace("m", string.Empty);
                            this.formats[j] = this.formats[j].Replace("i", string.Empty);
                            this.formats[j] = this.formats[j].Replace("h", string.Empty);
                            this.formats[j] = this.formats[j].Replace("d", string.Empty);
                        }
                        this.sugDaddy.uLog.renderCnt++;
                        this.rendered = true;
                    }
                }
            }
        }
    }
}

