using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToCSV
{
    public class Program
    {
        public static void Main(string[] args)
        {
            //string filePath = "C:\\Folder1\\Prev.xlsx";

            string filePath = "D:\\DATA\\Downloads\\Programming Challenge - Copy of ProductList (1).xlsx";
            //Open the Excel file using ClosedXML.
            using (XLWorkbook workBook = new XLWorkbook(filePath))
            {
                //Read the first Sheet from Excel file.
                IXLWorksheet workSheet = workBook.Worksheet(1);

                //Create a new DataTable.
                DataTable dt = new DataTable();

                //Loop through the Worksheet rows.
                bool firstRow = true;
                foreach (IXLRow row in workSheet.Rows())
                {
                    //Use the first row to add columns to DataTable.
                    if (firstRow)
                    {
                        foreach (IXLCell cell in row.Cells())
                        {
                            if(cell.Value.ToString() == "MfrPN")
                                dt.Columns.Add("Mfr P//N");
                            else if (cell.Value.ToString() == "Cost")
                                    dt.Columns.Add("Price");
                            else if (cell.Value.ToString() == "Coo")
                                dt.Columns.Add("COO");
                            else
                            dt.Columns.Add(cell.Value.ToString());
                        }
                        firstRow = false;
                    }
                    else
                    {
                        //Add rows to DataTable.
                        dt.Rows.Add();
                        int i = 0;
                        foreach (IXLCell cell in row.Cells())
                        {
                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
                            i++;
                        }
                    }

                    //GridView1.DataSource = dt;
                    //GridView1.DataBind();
                }
                var ii = 0;
                //changing values for all rows for single column
                //dt.AsEnumerable().ToList<DataRow>().ForEach(r =>  r["COO"] = r["COO"].ToString() == "TW1" ? "sud" : r["Coo"]);
                dt.AsEnumerable().ToList<DataRow>().ForEach(r => r["COO"] = string.IsNullOrEmpty(r["COO"].ToString()) ? "TW" : r["Coo"]);
                dt.AsEnumerable().ToList<DataRow>().ForEach(r => r["UOM"] = string.IsNullOrEmpty(r["UOM"].ToString()) ? "EA" : r["UOM"]);
                dt.AsEnumerable().ToList<DataRow>().ForEach(r => r["Price"] = GetPrice(r["Price"].ToString()));
                var io = 0;


                //creating csv file
                string strFilePath = @"D:\DATA\Downloads\Programming Challenge - Copy of ProductList" +".csv";
                StreamWriter sw = new StreamWriter(strFilePath, false);
                //headers    
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    sw.Write(dt.Columns[i]);
                    if (i < dt.Columns.Count - 1)
                    {
                        sw.Write(",");
                    }
                }
                sw.Write(sw.NewLine);
                foreach (DataRow dr in dt.Rows)
                {
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        if (!Convert.IsDBNull(dr[i]))
                        {
                            string value = dr[i].ToString();
                            if (value.Contains(','))
                            {
                                value = String.Format("\"{0}\"", value);
                                sw.Write(value);
                            }
                            else
                            {
                                sw.Write(dr[i].ToString());
                            }
                        }
                        if (i < dt.Columns.Count - 1)
                        {
                            sw.Write(",");
                        }
                    }
                    sw.Write(sw.NewLine);
                }
                sw.Close();
            }
        }
        public static decimal GetPrice(string price)
        {
            int x = 0;
            Int32.TryParse(price, out x);
            return x + (x * 20/100) ;
        }
    }
}

//    using (XLWorkbook workBook = new XLWorkbook(filePath))
//            {
//                //Read the first Sheet from Excel file.
//                IXLWorksheet workSheet = workBook.Worksheet(1);

//    //Create a new DataTable.
//    DataTable dt = new DataTable();

//    //Loop through the Worksheet rows.
//    bool firstRow = true;
//    //Range for reading the cells based on the last cell used.  
//    string readRange = "1:1";
//                foreach (IXLRow row in workSheet.RowsUsed())
//                {
//                    //Use the first row to add columns to DataTable.
//                    if (firstRow)
//                    {
//                        //Checking the Last cellused for column generation in datatable  
//                        readRange = string.Format("{0}:{1}", 1, row.LastCellUsed().Address.ColumnNumber);
//                        foreach (IXLCell cell in row.Cells(readRange))
//                        {
//                            dt.Columns.Add(cell.Value.ToString());
//                        }
//firstRow = false;
//                    }
//                    else
//                    {
//                        //Add rows to DataTable.
//                        dt.Rows.Add();
//                        int i = 0;
//                        foreach (IXLCell cell in row.Cells(readRange))
//                        {
//                            dt.Rows[dt.Rows.Count - 1][i] = cell.Value.ToString();
//                            i++;
//                        }
//                    }
                   
//                    //GridView1.DataSource = dt;
//                    //GridView1.DataBind();
//                }
//                var ii = 0;
//            }
