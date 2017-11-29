using System;
using System.Data;
using xls = Microsoft.Office.Interop.Excel; //import thu vien interop.excel vao de lam viec
using FORECAST_QUANTITY_DTO;
using System.Collections.Generic;
using System.Reflection;

namespace FORECAST_QUANTITY_DAL
{
    public class DAL_Excel
    {
        private InfoCustomer Info = new InfoCustomer(); 

        private string ReadExcel(string duongdan)
        {
            DataTable oTbl = new DataTable();

            xls.Application ExcelApp = new xls.Application();
            var temp = Info.GetType().GetProperties();

            try
            {
                ExcelApp.Workbooks.Open(duongdan);
                List<string> columns = new List<string>();
                //oTbl.Columns.Add(, typeof(string));//tao mot cot ten la MA_DDO co kieu du lieu string
                //oTbl.Columns.Add("B", typeof(string));
                //oTbl.Columns.Add("C", typeof(string));
                //oTbl.Columns.Add("D", typeof(string));
                //oTbl.Columns.Add("E", typeof(string));

                foreach (var item in temp)
                {
                    oTbl.Columns.Add(item.Name,item.GetType());
                }

                //doc du lieu tung sheet cua excel
                foreach (xls.Worksheet WSheet in ExcelApp.Worksheets)
                {
                    #region content

                    //tao mot datarow de lay du lieu cho tung cell
                    DataRow dr = oTbl.NewRow();//dataRow co kieu du lieu cung voi oTbl
                                               //bien i de doc tung dong trong sheet
                    long i = 2;

                    do
                    {
                        if (WSheet.Range["A" + i].Value == null && WSheet.Range["B" + i].Value == null && WSheet.Range["C" + i].Value == null && WSheet.Range["D" + i].Value == null && WSheet.Range["E" + i].Value == null)
                        {
                            break;//neu tro den dong cuoi cung cua sheet thi dung lai
                        }
                        dr = oTbl.NewRow();//DataRow co kieu du lieu cung voi oTbl
                        dr["A"] = WSheet.Range["A" + i].Value;
                        dr["B"] = WSheet.Range["B" + i].Value;
                        dr["C"] = WSheet.Range["C" + i].Value;
                        dr["D"] = WSheet.Range["D" + i].Value;
                        dr["E"] = WSheet.Range["E" + i].Value;
                        oTbl.Rows.Add(dr);
                        i++;
                    }
                    while (true);


                    #endregion
                }


                return "";
            }
            catch (Exception ex)
            {
                return ex.ToString();
            }
        }
    }
}
