namespace analyzer
{
    using System;
    using System.Collections.Generic;
    using System.Data;

    public class reportList
    {
        public report curSel;
        public List<report> reps = new List<report>();

        public void add(report rep)
        {
            int num = 0;
            bool flag = true;
            while (flag)
            {
                flag = false;
                foreach (report report in this.reps)
                {
                    if (report.staticNum == num)
                    {
                        flag = true;
                        break;
                    }
                }
                if (flag)
                {
                    num++;
                }
            }
            rep.staticNum = num;
            this.reps.Add(rep);
        }

        private static int compareByName(report x, report y)
        {
            return string.Compare(x.name, y.name);
        }

        public void copy(report oldRep, string newName)
        {
            report rep = new report {
                name = newName,
                filters = oldRep.filters.Copy(),
                staticNum = this.reps.Count,
                col = oldRep.col,
                grp = oldRep.grp,
                metricNms = oldRep.metricNms,
                metGrpNm = oldRep.metGrpNm,
                metrics = oldRep.metrics.Copy(),
                axCol = oldRep.axCol.Copy()
            };
            foreach (DataTable table in oldRep.axGrp)
            {
                rep.axGrp.Add(table.Copy());
            }
            rep.rendered = oldRep.rendered;
            if (rep.rendered)
            {
                rep.rep = oldRep.rep.Copy();
            }
            this.add(rep);
            this.sortByName();
        }

        public report findByName(string name)
        {
            foreach (report report in this.reps)
            {
                if (report.name == name)
                {
                    return report;
                }
            }
            return null;
        }

        public report findByStatNum(int num)
        {
            foreach (report report in this.reps)
            {
                if (report.staticNum == num)
                {
                    return report;
                }
            }
            return null;
        }

        public int nameToStatNum(string name)
        {
            foreach (report report in this.reps)
            {
                if (report.name == name)
                {
                    return report.staticNum;
                }
            }
            return this.reps.Count;
        }

        public string nameUnique(string name)
        {
            bool flag = true;
            foreach (report report in this.reps)
            {
                if (report.name == name)
                {
                    flag = false;
                    break;
                }
            }
            if (flag)
            {
                return name;
            }
            if (name.Contains(" (") && name.EndsWith(")"))
            {
                int result = 0;
                int startIndex = name.LastIndexOf(" (") + 2;
                if (int.TryParse(name.Substring(startIndex, (name.Length - startIndex) - 1), out result))
                {
                    name = name.Substring(0, name.LastIndexOf(" ("));
                }
            }
            int num3 = 1;
            string str = string.Empty;
            while (!flag)
            {
                str = name + " (" + num3.ToString() + ")";
                flag = true;
                foreach (report report2 in this.reps)
                {
                    if (report2.name == str)
                    {
                        flag = false;
                        break;
                    }
                }
                num3++;
            }
            return str;
        }

        public void remove(report rep)
        {
            this.reps.Remove(this.findByStatNum(rep.staticNum));
        }

        public void remove(int statNum)
        {
            this.reps.Remove(this.findByStatNum(statNum));
        }

        public void remove(string repName)
        {
            this.reps.Remove(this.findByName(repName));
        }

        public void sortByName()
        {
            this.reps.Sort(new Comparison<report>(reportList.compareByName));
        }
    }
}

