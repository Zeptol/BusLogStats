using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace BusDataStats
{
    public partial class Index : Form
    {
        private int _minDiff;

        public Index()
        {
            InitializeComponent();
        }

        private void Index_Load(object sender, EventArgs e)
        {
            var list =
            (from f in FileToList()
                orderby f.LineCode, f.Direction, f.StationNum, f.StationCode, f.VehCode, f.EstimatedTime
                select f).ToList();
            var srcList = new List<BusDataStats>();
            var temp = new List<BusData>();
            while (list.Any())
            {
                foreach (var data in list)
                {
                    var last = temp.LastOrDefault();
                    if (last == null)
                    {
                        temp.Add(data);
                        continue;
                    }
                    var diff = (data.ProcessTime - last.ProcessTime).TotalMinutes;
                    diff = diff < 0 ? diff + 1440 : diff;
                    if (diff > 1)
                        break;
                    temp.Add(data);
                }
                if (temp.Count > 1)
                {
                    var first = temp.First();
                    srcList.Add(new BusDataStats
                    {
                        LineCode = first.LineCode,
                        Direction = first.Direction,
                        StationNum = first.StationNum,
                        StationCode = first.StationCode,
                        VehCode = first.VehCode,
                        FirstEstimatedTime = first.EstimatedTime,
                        FirstProcessTime = first.ProcessTime,
                        LastEstimatedTime = temp.Last().EstimatedTime,
                        LastProcessTime = temp.Last().ProcessTime
                    });
                }
                list = list.Skip(temp.Count).ToList();
                temp.Clear();
            }
            dataGridViewStats.DataSource = srcList.OrderBy(t => t.TimeDiffMins).ToList();
            dataGridViewTotal.DataSource = (from t in srcList
                group t by t.TimeDiffMins
                into g
                orderby g.Key
                select new {TimeDiffMins = g.Key, Count = g.Count()}).ToList();

            dataGridViewMatrix.RowHeadersWidth = 60;
            var minDiff = (int) Math.Floor(srcList.Min(t => t.TimeDiffMins));
            _minDiff = minDiff = minDiff % 2 == 0 ? minDiff : minDiff - 1;
            var maxDiff = (int) Math.Ceiling(srcList.Max(t => t.TimeDiffMins));
            maxDiff = maxDiff % 2 == 0 ? maxDiff + 2 : maxDiff + 1;
            var maxTotal = (int) Math.Ceiling(srcList.Max(t => t.ActualTotalMins));
            maxTotal = maxTotal % 2 == 0 ? maxTotal + 2 : maxTotal + 1;
            var colCount = maxTotal / 2;
            var rowCount = (maxDiff - minDiff) / 2;
            var dt = new DataTable();
            for (var i = 0; i < colCount; i++)
            {
                dt.Columns.Add(string.Format("{0} ~ {1}", 2 * i, 2 * (i + 1)));
            }
            for (var j = 0; j < rowCount; j++)
            {
                var dr = dt.NewRow();
                for (var i = 0; i < colCount; i++)
                {
                    dr[i] =
                        srcList.Count(
                            t =>
                                t.ActualTotalMins >= 2 * i && t.ActualTotalMins < 2 * (i + 1) &&
                                t.TimeDiffMins >= minDiff + 2 * j && t.TimeDiffMins < minDiff + 2 * (j + 1));
                }
                dt.Rows.Add(dr);
            }
            dataGridViewMatrix.DataSource = dt;
            for (var i = 0; i < dataGridViewMatrix.Columns.Count; i++)
            {
                dataGridViewMatrix.Columns[i].Width = 25;
                dataGridViewMatrix.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void dataGridViewMatrix_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            var rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                dataGridViewTotal.RowHeadersWidth + 20, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics,
                string.Format("{0} ~ {1}", e.RowIndex * 2 + _minDiff, (e.RowIndex + 1) * 2 + _minDiff),
                new Font("tahoma", 8, FontStyle.Regular), rectangle,
                dataGridViewTotal.RowHeadersDefaultCellStyle.ForeColor,
                TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        }

        private static List<BusData> FileToList()
        {
            var res = new List<BusData>();
            var sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "data.txt", Encoding.UTF8);
            string line;
            while ((line = sr.ReadLine()) != null)
            {
                var arr = line.Split('|');
                if (!arr.Any() || arr.Length < 8 || string.IsNullOrWhiteSpace(arr[6]) ||
                    string.IsNullOrWhiteSpace(arr[7])) continue;
                var data = new BusData
                {
                    LineCode = arr[1],
                    Direction = (Direction) Enum.Parse(typeof(Direction), arr[2]),
                    StationNum = int.Parse(arr[3]),
                    StationCode = arr[4],
                    VehCode = arr[5],
                    EstimatedTime = DateTime.Parse(arr[6]),
                    ProcessTime = DateTime.Parse(arr[7])
                };
                data.ProcessTime = data.ProcessTime.AddSeconds(-data.ProcessTime.Second);
                res.Add(data);
            }
            sr.Dispose();
            return res;
        }

    }
}
