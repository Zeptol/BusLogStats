using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace BusDataStats
{
    public class ExcelHelper
    {
        private static readonly IWorkbook Wb = new HSSFWorkbook();

        private static readonly ICellStyle CellStyle = GetCellStyle(Wb, false, HorizontalAlignment.Center,
            FillPattern.NoFill);

        private static readonly ICellStyle BoldStyle = GetCellStyle(Wb, true, HorizontalAlignment.Center,
            FillPattern.NoFill);

        private static readonly ICellStyle FilledStyle = GetCellStyle(Wb, false, HorizontalAlignment.Center,
            FillPattern.SolidForeground);

        private static readonly ICellStyle LeftAlignStyle = GetCellStyle(Wb, false, HorizontalAlignment.Left,
            FillPattern.NoFill);

        public static void Export(DataTable dt, ExcelData data)
        {
            // create a new sheet
            ISheet sheet = Wb.CreateSheet();
            // declare a row object reference
            // declare a cell object reference
            ICell c;
            int lastCol = dt.Columns.Count;

            InitHeader(sheet, lastCol, data);

            // Titles
            var r = sheet.CreateRow(2);
            for (int i = 0; i < lastCol; i++)
            {
                c = r.CreateCell(i + 1);
                if (i > 0) c.SetCellValue(dt.Columns[i].ToString().Replace(" ", ""));
                c.CellStyle = BoldStyle;
            }

            for (int rowNum = 0; rowNum < dt.Rows.Count; rowNum++)
            {
                r = sheet.CreateRow(rowNum + 3);
                var row = dt.Rows[rowNum];
                c = r.CreateCell(0);
                if (rowNum == 0)
                    c.SetCellValue("预计行程时间-实际行程时间（分钟）");
                c.CellStyle = CellStyle;
                for (int cellNum = 0; cellNum < lastCol; cellNum++)
                {
                    c = r.CreateCell(cellNum + 1);
                    var val = row[cellNum];
                    int intVal;
                    if (int.TryParse(val.ToString(), out intVal))
                    {
                        c.SetCellValue(intVal);
                        c.CellStyle = intVal > 0 ? FilledStyle : CellStyle;
                    }
                    else
                    {
                        c.SetCellValue(val.ToString());
                        c.CellStyle = BoldStyle;
                    }
                }
            }
            //合并单元格
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 0, 1));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 2, 7));
            sheet.AddMergedRegion(new CellRangeAddress(0, 0, 8, lastCol));
            sheet.AddMergedRegion(new CellRangeAddress(1, 2, 0, 1));
            sheet.AddMergedRegion(new CellRangeAddress(1, 1, 2, lastCol));
            sheet.AddMergedRegion(new CellRangeAddress(3, dt.Rows.Count + 2, 0, 0));
            for (int i = 0; i < lastCol; i++)
            {
                //设置自适应宽度
                sheet.AutoSizeColumn(i);
            }
            // Save
            var fs = new FileStream(data.FileName.Replace(".txt", "") + ".xls", FileMode.OpenOrCreate);
            Wb.Write(fs);
            fs.Close();

        }

        private static void InitHeader(ISheet sheet, int lastCol, ExcelData data)
        {
            IRow r = sheet.CreateRow(0);
            ICell c;
            for (int i = 0; i <= lastCol; i++)
            {
                c = r.CreateCell(i);
                if (i == 0)
                {
                    c.SetCellValue("线路：");
                    c.CellStyle = CellStyle;
                }
                else if (i == 2)
                {
                    c.SetCellValue("数据日期：");
                    c.CellStyle = LeftAlignStyle;
                }
                else if (i == 8)
                {
                    c.SetCellValue(string.Format("数据量：{0}  百分比：{1:P}", data.Total, data.Ratio));
                    c.CellStyle = LeftAlignStyle;
                }
                else
                {
                    c.CellStyle = CellStyle;
                }
            }

            r = sheet.CreateRow(1);
            for (int i = 0; i <= lastCol; i++)
            {
                c = r.CreateCell(i);
                if (i == 0)
                {
                    c.SetCellValue("次数");
                    c.CellStyle = BoldStyle;
                }
                else if (i == 2)
                {
                    c.SetCellValue("实际行程时间（分钟）");
                    c.CellStyle = BoldStyle;
                }
                else c.CellStyle = CellStyle;
            }
        }

        private static ICellStyle GetCellStyle(IWorkbook wb, bool isBold, HorizontalAlignment hAlign,
            FillPattern fillType)
        {
            ICellStyle cellStyle = wb.CreateCellStyle();
            cellStyle.Alignment = hAlign;
            cellStyle.VerticalAlignment = VerticalAlignment.Center;
            cellStyle.BorderTop = BorderStyle.Thin;
            cellStyle.BorderLeft = BorderStyle.Thin;
            cellStyle.BorderRight = BorderStyle.Thin;
            cellStyle.BorderBottom = BorderStyle.Thin;
            cellStyle.FillForegroundColor = IndexedColors.Grey25Percent.Index;
            cellStyle.FillPattern = fillType;
            cellStyle.WrapText = true;
            IFont f = wb.CreateFont();
            f.FontName = "宋体";
            f.FontHeightInPoints = 12;
            f.IsBold = isBold;
            cellStyle.SetFont(f);
            return cellStyle;
        }


    }
}
