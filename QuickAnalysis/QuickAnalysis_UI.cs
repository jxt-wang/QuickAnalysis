using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;

using System.Configuration;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;

namespace QuickAnalysis
{
    [ComVisible(true)]
    public class QuickAnalysis_UI : Office.IRibbonExtensibility
    {

        #region Private Variables

        private Office.IRibbonUI ribbon;
        
        #endregion

        public QuickAnalysis_UI()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("QuickAnalysis.QuickAnalysis_UI.xml");
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Button Callbacks

        public void click_exportChart(Office.IRibbonControl control)
        {

            // Get Active Elements
            Excel.Chart chart_active = GetActiveChart();
            Excel.Workbook workbook_active = GetActiveWorkbook();

            if (chart_active != null && workbook_active != null)
            {

                // Save File Dialog
                SaveFileDialog dialog_saveFile = new SaveFileDialog();
                dialog_saveFile.Filter = "JPEG Image|*.jpg|Bitmap Image|*.bmp|GIF Image|*.gif|PNG Image|*.png|JPE Image|*.jpe";
                dialog_saveFile.Title = "Export Chart As...";

                // Set Default File Name
                dialog_saveFile.FileName = chart_active.ChartTitle.Text;

                // Set Default Directory
                SetDefaultDirectory(dialog_saveFile, workbook_active);

                // Show Save File Dialog
                dialog_saveFile.ShowDialog();

                // Export Chart
                if (!String.IsNullOrEmpty(dialog_saveFile.FileName))
                {
                    chart_active.Export(dialog_saveFile.FileName);
                }

            }

        }

        public void click_importData(Office.IRibbonControl control)
        {

            // Initialize Open File Dialog
            OpenFileDialog dialog_openFile = new OpenFileDialog();
            dialog_openFile.Filter = "Excel Files (*.xls;*.xlsx; *.xlsm; *.csv)| *.xls; *.csv; *.xlsx; *.xlsm|Unicode Text|*.txt";
            dialog_openFile.Title = "Select Files To Import...";
            dialog_openFile.Multiselect = true;

            // Set Active Elements
            Excel.Workbook workbook_active = GetActiveWorkbook();

            // Set Default Directory
            SetDefaultDirectory(dialog_openFile, workbook_active);

            // Show Open File Dialog
            DialogResult result_openFile = dialog_openFile.ShowDialog();

            if(result_openFile == System.Windows.Forms.DialogResult.OK)
            {

                DataSet dataset_import = new DataSet();

                foreach (String fileName in dialog_openFile.FileNames)
                {
                    try
                    {

                        DataTable datatable_temp = new DataTable(Path.GetFileNameWithoutExtension(fileName));
                        string ext = Path.GetExtension(fileName);
                        
                        using (TextFieldParser parser = new TextFieldParser(@fileName))
                        {

                            parser.TextFieldType = FieldType.Delimited;
                            
                            string delimeter = ConfigurationManager.AppSettings[ext + "_delimeter"];
                            parser.SetDelimiters(delimeter);
                            parser.HasFieldsEnclosedInQuotes = Convert.ToBoolean(ConfigurationManager.AppSettings[ext + "_quotes"]);
                            
                            if (ext == ".csv")
                            {

                                string str_checkName = ConfigurationManager.AppSettings[ext + "_name"];
                                string str_checkValue = ConfigurationManager.AppSettings[ext + "_value"];
                                string str_checkTitle = ConfigurationManager.AppSettings[ext + "_title"];

                                while (!parser.EndOfData)
                                {

                                    // TODO: DO NOT HARDCODE FIRST VAL and TYPEOF VALUES

                                    string[] fields_temp = parser.ReadFields();

                                    if (fields_temp[0] == str_checkName)
                                    {
                                        for (int i = 1; i < fields_temp.Length; i++)
                                        {
                                            datatable_temp.Columns.Add(new DataColumn(fields_temp[i], typeof(double)));
                                        }
                                    }
                                    else if (fields_temp[0] == str_checkValue)
                                    {

                                        double[] values_temp = Array.ConvertAll<string, double>(SubArray<string>(fields_temp, 1, fields_temp.Length - 1), Convert.ToDouble);

                                        if (values_temp.Length != datatable_temp.Columns.Count)
                                        {
                                            throw new Exception("Column mismatch.");
                                        }

                                        DataRow datarow_temp = datatable_temp.NewRow();

                                        for (int i = 0; i < values_temp.Length; i++)
                                        {
                                            datarow_temp[i] = values_temp[i];
                                        }

                                        datatable_temp.Rows.Add(datarow_temp);

                                    }
                                    else if (fields_temp[1] == str_checkTitle)
                                    {
                                        datatable_temp.TableName = fields_temp[2];
                                    }

                                }

                            }
                            
                            
                        } // End of 'using'

                        dataset_import.Tables.Add(datatable_temp);

                    }

                    catch (Exception IE)
                    {
                        MessageBox.Show(IE.Message + " error: could not load file");
                        continue;
                    }

                }

                Excel.Worksheet worksheet_new = CreateNewWorksheet();

                string[] arr_names = new string[dataset_import.Tables.Count + 1];
                
                DataTable voltage_table = dataset_import.Tables[0];
                arr_names[0] = "Vgs";
                double[] voltage_values = voltage_table.AsEnumerable().Select(r => r.Field<double>("Vgs")).ToArray();

                for (int v = 0; v < voltage_table.Rows.Count; v++)
                {
                    worksheet_new.Cells[v + 2, 1] = voltage_values[v];
                }
                
                for (int i = 0; i < dataset_import.Tables.Count; i++)
                {

                    DataTable datatable_temp = dataset_import.Tables[i];
                    arr_names[i+1] = "Id of " + datatable_temp.TableName;
                    double[] values_fromTable = datatable_temp.AsEnumerable().Select(r => r.Field<double>("Id")).ToArray();
                    
                    for(int j = 0; j < datatable_temp.Rows.Count; j++)
                    {
                        worksheet_new.Cells[j + 2, i + 2] = -1*values_fromTable[j];
                    }
                    
                }

                for (int h = 0; h < arr_names.Length; h++)
                {
                    worksheet_new.Cells[1, h + 1] = arr_names[h];
                }

                object misValue = System.Reflection.Missing.Value;

                Excel.Range range_chart;

                Excel.ChartObjects xlCharts = (Excel.ChartObjects)worksheet_new.ChartObjects(Type.Missing);
                Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(50, 150, 500, 250);
                Excel.Chart chartPage = myChart.Chart;

                range_chart = worksheet_new.UsedRange;
                Excel.Range range_tempA = range_chart.Columns[1].Find("-40");
                range_chart = range_chart.Resize[range_tempA.Row];

                Excel.Range range_tempB = worksheet_new.Cells[range_chart.Columns[1].Find("-20").Row, 2];
                Excel.Range c2 = worksheet_new.Cells[range_chart.Rows.Count, range_chart.Columns.Count];
                Excel.Range range_trendline = (Excel.Range)worksheet_new.get_Range(range_tempB, c2);
                
                range_chart.Cells[1, 1].Value = "";
                chartPage.SetSourceData(range_chart, misValue);
                //chartPage.SetSourceData(range_trendline, misValue);
                chartPage.ChartType = Excel.XlChartType.xlXYScatter;

                foreach (Excel.Series series in chartPage.SeriesCollection())
                {
                    Excel.Trendline trendline = series.Trendlines().Add(Excel.XlTrendlineType.xlLinear, System.Type.Missing, System.Type.Missing, 20, System.Type.Missing, System.Type.Missing, true, true, System.Type.Missing);
                    Excel.DataLabel datalabel_temp = trendline.DataLabel;
                    datalabel_temp.NumberFormat = "0.0000E+00";
                }

                Excel.Range a1 = range_chart.Columns[1].Find("-20");
                Excel.Range a2 = worksheet_new.Cells[range_chart.Rows.Count, 1];
                Excel.Range a3 = (Excel.Range)worksheet_new.get_Range(a1, a2);
                for (int i = 1; i <= range_trendline.Columns.Count; i++)
                {
                    Excel.Series series_temp = chartPage.SeriesCollection().Add(range_trendline.Columns[i]);
                    
                    series_temp.XValues = a3;
                    Excel.Trendlines trendlines_temp = series_temp.Trendlines();
                    Excel.Trendline trendline_temp = trendlines_temp.Add(Excel.XlTrendlineType.xlLinear, System.Type.Missing, System.Type.Missing, 20, System.Type.Missing, System.Type.Missing, true, true, System.Type.Missing);
                    Excel.DataLabel datalabel_temp = trendline_temp.DataLabel;
                    
                    datalabel_temp.NumberFormat = "0.0000E+00";
                  
                }
                

            }
            
        }

        public void click_analyzeChart(Office.IRibbonControl control)
        {
            Excel.Chart chart_active = GetActiveChart();

            List<double> x_intercepts = new List<double>();

            foreach (Excel.Series series in chart_active.SeriesCollection())
            {
                Excel.Trendlines trendlines_temp = series.Trendlines();
                if(trendlines_temp.Count > 0)
                {

                    Excel.DataLabel dlabel_temp = trendlines_temp.Item(1).DataLabel;
                    string label = dlabel_temp.Text;
                    int x = label.IndexOf("R²");
                    int y = label.IndexOf("y");

                    string equation = label.Substring(y, x);
                    string r_val = label.Substring(x);

                    Regex pattern = new Regex(@"[\d]*\.?[\d]+(E[-+][\d]+)?");
                    Match match = pattern.Match(r_val);
                    double r_squared = Double.Parse(match.Value, System.Globalization.NumberStyles.Float);

                    if (r_squared < 0.9)
                    {
                        trendlines_temp.Item(1).Delete();
                        if (r_squared < 0.7)
                        {
                            series.Delete();
                        }
                        continue;
                    }
                    

                    Regex pattern_x = new Regex(@"[\d]*\.?[\d]+(E[-+][\d]+)?");
                    MatchCollection matches = pattern_x.Matches(equation);


                    Stack<double> eq_vals = new Stack<double>();

                    foreach (Match match_temp in matches)
                    {
                        double val = Double.Parse(match_temp.Value, System.Globalization.NumberStyles.Float);
                        eq_vals.Push(val);
                    }

                    double y_int = eq_vals.Pop();
                    double y_slope = eq_vals.Pop();

                    if(y_slope < 1e-8)
                    {
                        trendlines_temp.Item(1).Delete();
                        series.Delete();
                        continue;
                    }

                    double x_int = (-1*y_int) / y_slope;
                    
                    x_intercepts.Add(x_int);




                    
                    
                }
            }

            if(x_intercepts.Count > 0)
            {
                MessageBox.Show(x_intercepts.Average().ToString());
            }

        }

        public void click_saveData(Office.IRibbonControl control)
        {
            MessageBox.Show("TODO: Implement Save");
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        private static void SetDefaultDirectory(FileDialog dialog, Excel.Workbook wb)
        {
            try
            {
                if (!String.IsNullOrEmpty(wb.Path))
                {
                    string paramPath = wb.Path + @"\";
                    dialog.InitialDirectory = paramPath;
                }
            }
            catch(Exception IE)
            {
                MessageBox.Show("error: could not select current directory as default.");
            }
        }

        //private static void UpdateCells<T>(Excel.Worksheet ws, T[][] data_import, T[] header_import)
        //{
        //    for (int h = 0; h < arr_names.Length; h++)
        //    {
        //        worksheet_new.Cells[1, h + 1] = arr_names[h];
        //    }
        //}

        //private static T[][] InitData<T>(DataSet dataset_import, string param_x, string param_y)
        //{
            
        //    T[][] data_extract = new T[dataset_import.Tables.Count + 1][];
        //    T[] x_values = dataset_import.Tables[0].AsEnumerable().Select(r => r.Field<T>(param_x)).ToArray();

        //    // check if all tables have the same set of x data
        //    for (int i = 0; i < dataset_import.Tables.Count; i++)
        //    {
        //        T[] x_valuesTemp = dataset_import.Tables[0].AsEnumerable().Select(r => r.Field<T>(param_x)).ToArray();
        //        if (x_valuesTemp.SequenceEqual(x_values))
        //        {
        //            return null;
        //        }
        //    }

        //    data_extract[0] = x_values;

        //    for (int i = 0; i < dataset_import.Tables.Count; i++)
        //    {

        //        DataTable datatable_temp = dataset_import.Tables[i];
        //        T[] y_values = datatable_temp.AsEnumerable().Select(r => r.Field<T>(param_y)).ToArray();

        //        data_extract[i + 1] = (y_values);

        //    }

        //    return data_extract;

        //}

        //private static string[] InitHeaders(DataSet dataset_import, string param_x, string param_y)
        //{

        //    string[] arr_names = new string[dataset_import.Tables.Count + 1];
        //    arr_names[0] = param_x;
            
        //    for (int i = 1; i < dataset_import.Tables.Count + 1; i++)
        //    {
        //        DataTable datatable_temp = dataset_import.Tables[i];

        //        string table_name = "";
                
        //        arr_names[i] = param_y + " | " + table_name;

        //        arr_names[i + 1] = "Id of " + datatable_temp.TableName;
        //        double[] values_fromTable = datatable_temp.AsEnumerable().Select(r => r.Field<double>("Id")).ToArray();

        //        for (int j = 0; j < datatable_temp.Rows.Count; j++)
        //        {
        //            worksheet_new.Cells[j + 2, i + 2] = -1 * values_fromTable[j];
        //        }

        //    }

        //    for (int h = 0; h < arr_names.Length; h++)
        //    {
        //        worksheet_new.Cells[1, h + 1] = arr_names[h];
        //    }



        //}

        private static Excel.Chart GetActiveChart()
        {
            try
            {
                return ((Excel.Chart)Globals.ThisAddIn.Application.ActiveWorkbook.ActiveChart);

            }
            catch (Exception IE)
            {
                MessageBox.Show("error: could not load active chart.");
                return null;
            }
        }

        private static Excel.Workbook GetActiveWorkbook()
        {
            try
            {
                return ((Excel.Workbook)Globals.ThisAddIn.Application.ActiveWorkbook);
            }
            catch (Exception IE)
            {
                MessageBox.Show("error: could not load active workbook.");
                return null;
            }

        }

        private static Excel.Worksheet CreateNewWorksheet()
        {
            try
            {
                return (Excel.Worksheet)Globals.ThisAddIn.Application.Worksheets.Add();
            }
            catch (Exception IE)
            {
                MessageBox.Show("error: could not create new worksheet");
                return null;
            }
        }

        private static T[] SubArray<T>(T[] data, int index, int length)
        {
            T[] result = new T[length];
            Array.Copy(data, index, result, 0, length);
            return result;
        }
        
        #endregion

        #region Images

        public Bitmap GetImage(Office.IRibbonControl control)
        {

            switch (control.Id)
            {
                case "button_importData": return new Bitmap(Properties.Resources.import);
                case "button_saveData": return new Bitmap(Properties.Resources.save_as);
                case "button_exportChart": return new Bitmap(Properties.Resources.export);
                case "button_analyzeChart": return new Bitmap(Properties.Resources.tick_marks);
            }

            return null;

        }

        #endregion
    }
}
