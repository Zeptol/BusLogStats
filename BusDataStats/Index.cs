using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace BusDataStats
{
    public partial class Index : Form
    {
        private DataTable _dtMatrix;
        private readonly ExcelData _data = new ExcelData();

        public Index()
        {
            InitializeComponent();
        }

        //private void dataGridViewMatrix_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        //{
        //    //var rectangle = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
        //        //dataGridViewTotal.RowHeadersWidth + 20, e.RowBounds.Height);
        //    //TextRenderer.DrawText(e.Graphics,
        //    //    string.Format("{0} ~ {1}", e.RowIndex * 2 + _minDiff, (e.RowIndex + 1) * 2 + _minDiff),
        //    //    new Font("tahoma", 8, FontStyle.Regular), rectangle,
        //    //    dataGridViewTotal.RowHeadersDefaultCellStyle.ForeColor,
        //    //    TextFormatFlags.VerticalCenter | TextFormatFlags.Right);
        //}

        private static List<BusData> FileToList(Stream stream)
        {
            var res = new List<BusData>();
            //var sr = new StreamReader(AppDomain.CurrentDomain.BaseDirectory + "data.txt", Encoding.UTF8);
            var sr = new StreamReader(stream);
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

        private void btnChoose_Click(object sender, EventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                AddExtension = true,
                CheckFileExists = true,
                CheckPathExists = true,
                Filter = @"文本文件（*.txt）|*.txt",
                RestoreDirectory = true
            };
            if (dialog.ShowDialog() != DialogResult.OK) return;
            _data.FileName = dialog.FileName;
            var list =
            (from f in FileToList(dialog.OpenFile())
                orderby f.LineCode, f.Direction, f.StationNum, f.StationCode, f.VehCode, f.ProcessTime
                select f).ToList();
            var srcList = new List<BusDataStats>();
            var temp = new List<BusData>();
            var timer = new Stopwatch();
            timer.Start();
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
                    if (diff > 30)
                        break;
                    temp.Add(data);
                }
                if (temp.Count > 1)
                {
                    var first = temp.First();
                    var stats = new BusDataStats
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
                    };
                    if (!(stats.TimeDiffMins < -30))
                        srcList.Add(stats);
                }
                list = list.Skip(temp.Count).ToList();
                temp.Clear();
            }
            timer.Stop();
            MessageBox.Show(this, $"统计耗时：{(float) timer.ElapsedMilliseconds / 1000:F3}秒", @"提示",
                MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            dataGridViewStats.DataSource = srcList.OrderBy(t => t.TimeDiffMins).ToList();
            var statsList = (from t in srcList
                group t by t.TimeDiffMins
                into g
                orderby g.Key
                select new {TimeDiffMins = g.Key, Count = g.Count()}).ToList();
            dataGridViewTotal.DataSource = statsList;
            _data.Total = statsList.Sum(t => t.Count);
            _data.Ratio = (float) statsList.Where(t => t.TimeDiffMins >= -1 && t.TimeDiffMins <= 1).Sum(t => t.Count) /
                          _data.Total;

            dataGridViewMatrix.RowHeadersWidth = 60;
            var minDiff = (int) Math.Floor(srcList.Min(t => t.TimeDiffMins));
            minDiff = minDiff % 2 == 0 ? minDiff : minDiff - 1;
            var maxDiff = (int) Math.Ceiling(srcList.Max(t => t.TimeDiffMins));
            maxDiff = maxDiff % 2 == 0 ? maxDiff + 2 : maxDiff + 1;
            var maxTotal = (int) Math.Ceiling(srcList.Max(t => t.ActualTotalMins));
            maxTotal = maxTotal % 2 == 0 ? maxTotal + 2 : maxTotal + 1;
            var colCount = maxTotal / 2;
            var rowCount = (maxDiff - minDiff) / 2;
            var dt = new DataTable();
            dt.Columns.Add("次数");
            for (var i = 0; i < colCount; i++)
            {
                dt.Columns.Add($"{2 * i} ~ {2 * (i + 1)}");
            }
            for (var j = 0; j < rowCount; j++)
            {
                var dr = dt.NewRow();
                dr[0] = $"{j * 2 + minDiff} ~ {(j + 1) * 2 + minDiff}";
                for (var i = 0; i < colCount; i++)
                {
                    dr[i + 1] =
                        srcList.Count(
                            t =>
                                t.ActualTotalMins >= 2 * i && t.ActualTotalMins < 2 * (i + 1) &&
                                t.TimeDiffMins >= minDiff + 2 * j && t.TimeDiffMins < minDiff + 2 * (j + 1));
                }
                dt.Rows.Add(dr);
            }
            dataGridViewMatrix.DataSource = _dtMatrix = dt;
            for (var i = 0; i < dataGridViewMatrix.Columns.Count; i++)
            {
                dataGridViewMatrix.Columns[i].Width = 25;
                dataGridViewMatrix.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (_dtMatrix != null)
            {
                try
                {
                    ExcelHelper.Export(_dtMatrix, _data);
                    MessageBox.Show(this, @"导出成功！", @"提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(this, ex.Message, @"错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else MessageBox.Show(this, @"请选择文件！", @"提示", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

    }
}
