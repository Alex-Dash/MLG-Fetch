using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace MLG_Fetch
{
    class Excel
    {
        Object missingObj = System.Reflection.Missing.Value;
        Object trueObj = true;
        Object falseObj = false;
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        Worksheet ws;

        public Excel(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        public void Visible()
        {
            excel.Visible = true;
        }

        public void AddSheet() {
            wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);

        }

        public void SortShit(int Sheet, int end_col, string byColLetter) {
            ws = wb.Sheets[Sheet];

            _Excel.Range rng = ws.Range[ws.Cells[2, 1], ws.Cells[2 + Globals.REGIONS_COUNT, end_col]];

            ws.Sort.SortFields.Clear();
            ws.Sort.SortFields.Add(rng.Columns[byColLetter], _Excel.XlSortOn.xlSortOnValues, _Excel.XlSortOrder.xlDescending, System.Type.Missing, _Excel.XlSortDataOption.xlSortNormal);
            var sort = ws.Sort;
            sort.SetRange(rng.Rows);
            sort.Header = _Excel.XlYesNoGuess.xlYes;
            sort.MatchCase = false;
            sort.Orientation = _Excel.XlSortOrientation.xlSortColumns;
            sort.SortMethod = _Excel.XlSortMethod.xlPinYin;
            sort.Apply();

        }

        public _Excel.Range GetRng(int Start_Row, int Start_Col, int End_Row, int End_Col, int Sheet){
            ws = wb.Sheets[Sheet];
            _Excel.Range rng = ws.Range[ws.Cells[Start_Row, Start_Col], ws.Cells[End_Row, End_Col]];
            return rng;
        }

        public void ApplyCondForm(int Sheet, int s_row, int s_col, int e_row, int e_col,_Excel.ColorScale[] arr, int scaler) {
            ws = wb.Sheets[Sheet];
            _Excel.Range rng = ws.Range[ws.Cells[s_row, s_col], ws.Cells[e_row, e_col]];

            arr[scaler] = (_Excel.ColorScale)(rng.FormatConditions.AddColorScale(3));
            arr[scaler].ColorScaleCriteria[1].Type = _Excel.XlConditionValueTypes.xlConditionValueLowestValue;
            arr[scaler].ColorScaleCriteria[1].FormatColor.Color = System.Drawing.Color.DodgerBlue; 

            arr[scaler].ColorScaleCriteria[2].Type = _Excel.XlConditionValueTypes.xlConditionValuePercentile;
            arr[scaler].ColorScaleCriteria[2].Value = 50;
            arr[scaler].ColorScaleCriteria[2].FormatColor.Color = System.Drawing.Color.White;  

            arr[scaler].ColorScaleCriteria[3].Type = _Excel.XlConditionValueTypes.xlConditionValueHighestValue;
            arr[scaler].ColorScaleCriteria[3].FormatColor.Color = System.Drawing.Color.IndianRed;  
     
        }
    
        public void PaintSheet(_Excel.Range rng, System.Drawing.Color Color)
        {
            rng.Interior.Color = Color;
        }

        public _Excel.Range GetUsedRangeOf(int Sheet)
        {
            ws = wb.Sheets[Sheet];
            return ws.UsedRange;
        }

        public void FixTheShit(int Sheet)
        {
            ws = wb.Sheets[Sheet];
            _Excel.Range src;
            _Excel.Range dest;
            int lastCol = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            int lastRow = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            
            switch (Sheet)
            {
                case 10:
                    src = GetRng(2,1,4,3, Sheet);
                    dest = GetRng(6, 1, 8, 3, Sheet);
                    src.Copy(dest);
                    src = GetRng(5, 1, 8, 3, Sheet);
                    dest = GetRng(2, 1, 5, 3, Sheet);
                    src.Copy(dest);
                    for(int ro = 6; ro < 10; ro++)
                    {
                        ws.Rows[ro].Clear();
                    }
                    
                    break;
                case 11:
                    src = GetRng(2, 1, 4, 3, Sheet);
                    dest = GetRng(6, 1, 8, 3, Sheet);
                    src.Copy(dest);
                    src = GetRng(5, 1, 8, 3, Sheet);
                    dest = GetRng(2, 1, 5, 3, Sheet);
                    src.Copy(dest);
                    for (int ro = 6; ro < 10; ro++)
                    {
                        ws.Rows[ro].Clear();
                    }
                    break;

                case 12:
                    ws.Rows[5].Copy(ws.Rows[7]);
                    ws.Rows[2].Copy(ws.Rows[8]);
                    ws.Rows[3].Copy(ws.Rows[9]);
                    ws.Rows[4].Copy(ws.Rows[10]);

                    ws.Rows[7].Copy(ws.Rows[2]);
                    ws.Rows[8].Copy(ws.Rows[3]);
                    ws.Rows[9].Copy(ws.Rows[4]);
                    ws.Rows[10].Copy(ws.Rows[5]);

                    //fix labels
                    for(int co = 2; co<= lastCol; co++)
                    {
                        string or = ws.Cells[1, co].Value2;
                        string[] arr = or.Split(" - ".ToCharArray());
                        ws.Cells[1, co].Value2 = arr[0].Substring(0, 5) + " - " + arr[3].Substring(0, 5);
                    }

                    for (int ro = 7; ro < 11; ro++)
                    {
                        ws.Rows[ro].Clear();
                    }
                    break;
                case 13:
                    //fix labels
                    for (int co = 2; co <= lastCol; co++)
                    {
                        string or = ws.Cells[1, co].Value2;
                        string[] arr = or.Split(" - ".ToCharArray());
                        ws.Cells[1, co].Value2 = arr[0].Substring(0, 5) + " - " + arr[3].Substring(0, 5);
                    }

                    break;
                case 4:
                    //fix labels
                    string p1 = ws.Cells[1, 2].Value2;
                    string p2 = ws.Cells[1, 5].Value2;
                    string newp1 =" "+ p1.Split(" - ".ToCharArray())[0].Substring(0,5) + " - " + p1.Split(" - ".ToCharArray())[3].Substring(0, 5);
                    string newp2 =" "+ p2.Split(" - ".ToCharArray())[0].Substring(0, 5) + " - " + p2.Split(" - ".ToCharArray())[3].Substring(0, 5);
                    ws.Cells[2, 2].Value2 += newp1;
                    ws.Cells[2, 3].Value2 += newp1;
                    ws.Cells[2, 4].Value2 += newp1;

                    ws.Cells[2, 5].Value2 = "% Сообщений" + newp2;
                    ws.Cells[2, 6].Value2 = "Прирост сообщ., %";
                    ws.Cells[2, 7].Value2 = "Медиа-Индекс" + newp2;

                    //change order of cols

                    ws.Columns[5].Copy(ws.Columns[9]);
                    ws.Columns[6].Copy(ws.Columns[10]);
                    ws.Columns[4].Copy(ws.Columns[11]);
                    ws.Columns[7].Copy(ws.Columns[12]);
                    ws.Columns[8].Copy(ws.Columns[13]);

                    for (int co = 9; co<14; co++)
                    {
                        ws.Columns[co].Copy(ws.Columns[co - 5]);
                        ws.Columns[co].Clear();
                    }
                    ws.Rows[1].Clear();
                    
                    for(int ro = 2; ro<8; ro++)
                    {
                        ws.Rows[ro].Copy(ws.Rows[ro - 1]);
                    }
                    ws.Columns[9].Clear();

                    ws.Rows[5].Copy(ws.Rows[8]);
                    ws.Rows[2].Copy(ws.Rows[9]);
                    ws.Rows[3].Copy(ws.Rows[10]);
                    ws.Rows[4].Copy(ws.Rows[11]);

                    for(int ro = 8; ro<12; ro++)
                    {
                        ws.Rows[ro].Copy(ws.Rows[ro-6]);
                        ws.Rows[ro].Clear();
                    }
                    ws.Rows[7].Clear();

                    break;
                case 7:
                    //fix labels
                    p1 = ws.Cells[1, 2].Value2;
                    p2 = ws.Cells[1, 5].Value2;
                    newp1 = " " + p1.Split(" - ".ToCharArray())[0].Substring(0, 5) + " - " + p1.Split(" - ".ToCharArray())[3].Substring(0, 5);
                    newp2 = " " + p2.Split(" - ".ToCharArray())[0].Substring(0, 5) + " - " + p2.Split(" - ".ToCharArray())[3].Substring(0, 5);
                    ws.Cells[2, 2].Value2 += newp1;
                    ws.Cells[2, 3].Value2 += newp1;
                    ws.Cells[2, 4].Value2 += newp1;

                    ws.Cells[2, 5].Value2 = "% Сообщений" + newp2;
                    ws.Cells[2, 6].Value2 = "Прирост сообщ., %";
                    ws.Cells[2, 7].Value2 = "Медиа-Индекс" + newp2;

                    //change order of cols

                    ws.Columns[5].Copy(ws.Columns[9]);
                    ws.Columns[6].Copy(ws.Columns[10]);
                    ws.Columns[4].Copy(ws.Columns[11]);
                    ws.Columns[7].Copy(ws.Columns[12]);
                    ws.Columns[8].Copy(ws.Columns[13]);

                    for (int co = 9; co < 14; co++)
                    {
                        ws.Columns[co].Copy(ws.Columns[co - 5]);
                        ws.Columns[co].Clear();
                    }
                    ws.Rows[1].Clear();

                    for (int ro = 2; ro < 8; ro++)
                    {
                        ws.Rows[ro].Copy(ws.Rows[ro - 1]);
                    }
                    ws.Columns[9].Clear();

                    ws.Rows[5].Copy(ws.Rows[8]);
                    ws.Rows[2].Copy(ws.Rows[9]);
                    ws.Rows[3].Copy(ws.Rows[10]);
                    ws.Rows[4].Copy(ws.Rows[11]);

                    for (int ro = 8; ro < 12; ro++)
                    {
                        ws.Rows[ro].Copy(ws.Rows[ro - 6]);
                        ws.Rows[ro].Clear();
                    }
                    ws.Rows[7].Clear();

                    break;

                case 14:
                    ws.Rows[5].Copy(ws.Rows[7]);
                    ws.Rows[2].Copy(ws.Rows[8]);
                    ws.Rows[3].Copy(ws.Rows[9]);
                    ws.Rows[4].Copy(ws.Rows[10]);

                    ws.Rows[7].Copy(ws.Rows[2]);
                    ws.Rows[8].Copy(ws.Rows[3]);
                    ws.Rows[9].Copy(ws.Rows[4]);
                    ws.Rows[10].Copy(ws.Rows[5]);

                    //fix labels
                    for (int co = 2; co <= lastCol; co++)
                    {
                        string or = ws.Cells[1, co].Value2;
                        string[] arr = or.Split(" - ".ToCharArray());
                        ws.Cells[1, co].Value2 = arr[0].Substring(0, 5) + " - " + arr[3].Substring(0, 5);
                    }
                    for (int ro = 7; ro < 11; ro++)
                    {
                        ws.Rows[ro].Clear();
                    }
                    break;

                case 8:
                    src = GetRng(1,6,lastRow, lastCol,Sheet);
                    dest = GetRng(1,10,lastRow, lastCol+4,Sheet);
                    src.Copy(dest);
                    src = GetRng(1, lastCol+1, lastRow, lastCol + 4, Sheet);
                    dest = GetRng(1, 6, lastRow, 10, Sheet);
                    src.Copy(dest);
                    ws.Columns[lastCol + 1].clear();
                    ws.Columns[lastCol + 2].clear();
                    ws.Columns[lastCol + 3].clear();
                    ws.Columns[lastCol + 4].clear();
                    break;

                case 9:
                    src = GetRng(1, 5, lastRow, lastCol, Sheet);
                    dest = GetRng(1, 8, lastRow, lastCol + 3, Sheet);
                    src.Copy(dest);
                    src = GetRng(1, lastCol + 1, lastRow, lastCol + 3, Sheet);
                    dest = GetRng(1, 5, lastRow, 8, Sheet);
                    src.Copy(dest);
                    ws.Columns[lastCol + 1].clear();
                    ws.Columns[lastCol + 2].clear();
                    ws.Columns[lastCol + 3].clear();
                    break;


                default:
                    break;
            }
        }

        public void AddTotalsFix(double[] MessagesTotals, double[] IndexesTotal, double[] LastMessagesTotals, double[] LastIndexesTotal, int RegionCount)
        {
            double LastMessagesSum = LastMessagesTotals[0]+ LastMessagesTotals[1] + LastMessagesTotals[2] + LastMessagesTotals[3];
            double MessagesSum = MessagesTotals[0] + MessagesTotals[1] + MessagesTotals[2] + MessagesTotals[3];
            

            //messages
            ws = wb.Sheets[8];
            for(int partId = 0; partId<4; partId++)
            {
                string MT = FindPercent(MessagesTotals[partId], MessagesSum);
                string LMT = FindPercent(LastMessagesTotals[partId], LastMessagesSum);
                ws.Cells[RegionCount + 3, 4 * partId + 3].Value2 = MT;
                ws.Cells[RegionCount + 3, 4 * partId + 4].Value2 = LMT;

                string MTclear = MT.Remove(MT.Length - 1, 1);
                string LMTclear = LMT.Remove(LMT.Length - 1, 1);

                ws.Cells[RegionCount + 3, 4 * partId + 5].Value2 = (Convert.ToDouble(MTclear) - Convert.ToDouble(LMTclear)).ToString() + "%";
            }

            //indexes
            ws = wb.Sheets[9];
            for (int partId = 0; partId < 4; partId++)
            {
                ws.Cells[RegionCount + 3, 3 * partId + 4].Value2 = (IndexesTotal[partId] - LastIndexesTotal[partId]).ToString();
            }

        }


        public void PlotChart(int Sheet, string rng_start, string rng_end, string filename) {

            
            ws = wb.Sheets[Sheet];

            _Excel.Range xlRange = ws.UsedRange;
            _Excel.Range chartRange;
            _Excel.ChartObjects xlCharts = (_Excel.ChartObjects)
            ws.ChartObjects(Type.Missing);
            _Excel.ChartObject myChart = (_Excel.ChartObject) xlCharts.Add(10, 80, 500, 250);
            

            _Excel.Chart chartPage = myChart.Chart;
            chartRange = ws.get_Range(rng_start, rng_end);
            chartPage.SetSourceData(chartRange, missingObj);
            chartPage.ChartType = _Excel.XlChartType.xl3DColumnClustered;
            chartPage.Perspective = 10;
            chartPage.Rotation = 10;

            chartPage.ApplyDataLabels();

            // Export chart as picture file
            string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath); 
            chartPage.Export(base_path+"\\"+filename,"BMP", missingObj);
            

        }
        public void PlotCSumChart(int Sheet, string filename)
        {
            ws = wb.Sheets[Sheet];
            var charts = ws.ChartObjects() as _Excel.ChartObjects;
            var chartObject = charts.Add(60, 10, 780, 400) as _Excel.ChartObject;
            var chart = chartObject.Chart;

            // Set chart range.
            //var range = worksheet.get_Range(topLeft, bottomRight);
            var range = ws.UsedRange;
            chart.SetSourceData(range);
            // Set chart properties.
            chart.ChartType = _Excel.XlChartType.xlLine;
            chart.ChartWizard(Source: range,
                Title: "Кумулятивная сумма по датам",
                CategoryTitle: "Даты",
                ValueTitle: "Количество");
            chart.PlotBy = XlRowCol.xlRows;
            chart.ApplyDataLabels(_Excel.XlDataLabelsType.xlDataLabelsShowLabel, true, true, false, false, false, true, true, false, false);
            string base_path = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            
            chart.Export(base_path + "\\" + filename, "BMP", missingObj);
        }
        public void AllignRange(string Row, int Sheet)
        {
            ws = wb.Sheets[Sheet];
            ws.Rows[Row].Cells.Orientation = _Excel.XlOrientation.xlUpward;
            ws.Rows[Row].Cells.VerticalAlignment = _Excel.XlVAlign.xlVAlignCenter;
            ws.Rows[Row].Cells.HorizontalAlignment = _Excel.XlHAlign.xlHAlignCenter;
            ws.Rows[Row].Cells.RowHeight = 90;
            ws.Rows[Row].Cells.WrapText = true;
        }

        public void WrapDetailed()
        {
            ws = wb.Sheets[8];
            ws.Columns[1].ColumnWidth = 22;
            ws.Columns[1].Cells.WrapText = true;
            ws = wb.Sheets[9];
            ws.Columns[1].ColumnWidth = 22;
            ws.Columns[1].Cells.WrapText = true;

        }

        public int ListCount()
        {
            return wb.Worksheets.Count;
        }

        public string ReadCell(int i, int j, int Sheet)
        {
            ws = wb.Worksheets[Sheet];

            if (ws.Cells[i, j].Value2 != null)
            {
                return ws.Cells[i, j].Value2.ToString();
            }
            else {
                return "";
            }


        }
        public void WriteToCell(int i, int j, string content, int Sheet) {
            ws = wb.Worksheets[Sheet];
            ws.Cells[i, j].Value2 = content;
        }

        public double RangeSum(int row1, int col1, int row2, int col2, int Sheet)
        {
            double sum = 0;
            ws = wb.Worksheets[Sheet];
            for (int col = col1; col < col2+1; col++)
            {
                for (int row = row1; row < row2; row++)
                {
                    double cell = Convert.ToDouble(ReadCell(row, col, Sheet));

                    sum += cell;
                    
                }
            }
            return sum;

        }

        public string FindPercent(double what, double where)
        {
            
            return(Math.Round((what / where),3) * 100).ToString() + "%";

        }

        public Tuple<int, int> FindCell(string content, int Sheet)
        {
            
            int row = 0;
            int col = 0;
            ws = wb.Worksheets[Sheet];
            try
            {
                row = ws.Cells.Find(content).Row;
                col = ws.Cells.Find(content).Column;
            }
            catch {

                int lastCol = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;

                         int i = 1;
                     for (int j = 1; j < lastCol; j++) {
                     if ((ReadCell(i,j,Sheet) != "") & (ReadCell(i, j, Sheet) == content)) {
                             row = i;
                             col = j;
                             MessageBox.Show(row.ToString() + "  " + col.ToString(), "Found");
                             break;
                         }
                     }
                 

            }
            
            return Tuple.Create(row, col);

        }

        public void Save()
        {
            wb.Save();
            MessageBox.Show("Файл сохранен как " + Globals.FILE_NAME, "Автоматическое Сохнанение");
        }

        public string SaveAs( string fileName = "") {
            if (fileName == "")
            {
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.InitialDirectory = "c:\\";
                saveFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
                saveFileDialog1.FilterIndex = 0;
                saveFileDialog1.RestoreDirectory = true;

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        wb.SaveAs(saveFileDialog1.FileName);
                        MessageBox.Show("Файл сохранен как " + saveFileDialog1.FileName, "Ручное сохранение");
                        wb.Close();
                        return saveFileDialog1.FileName;
                    }
                    catch {
                        MessageBox.Show("Невозможно сохранить файл. Выберите другое имя или путь." + saveFileDialog1.FileName, "Ошибка сохранения");

                        SaveAs();
                    }
                    
                }
            } else
            {
                wb.SaveAs(fileName);
                MessageBox.Show("Файл сохранен как " + fileName, "Автоматическое Сохнанение");
                wb.Close();
                return fileName;
            }
            return "none";
        }

        public void DelCellContents(int row, int col, int Sheet) {
            ws = wb.Worksheets[Sheet];
            ws.Cells[row, col].ClearContents();
            ws.Cells[row, col].ClearFormats();

        }

        public void ClearListContents(int Sheet) {
            ws = wb.Worksheets[Sheet];
            ws.Cells.ClearContents();

        }
        
        public void Close() {
            try { wb.Close();
                excel.Quit();
            }
            catch {
            }
               
        }


    }
}
