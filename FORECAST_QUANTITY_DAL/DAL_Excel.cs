using System;
using System.Data;
using FORECAST_QUANTITY_DTO;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Text;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;

namespace FORECAST_QUANTITY_DAL
{
    public class DAL_Excel
    {
        private InfoCustomer Info = new InfoCustomer();
        public DBConnect dbConnect = new DBConnect();

        private DataTable ReadExcel(string path)
        {
            DataTable oTbl = new DataTable();
            var temp = Info.GetType().GetProperties(); // get property infocustomer

            foreach (var item in temp)
                oTbl.Columns.Add(item.Name, item.GetType());

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(path)))
                {
                    if (package.Workbook.Worksheets.Count > 0)
                    {
                        var workSheet = package.Workbook.Worksheets.FirstOrDefault();

                        foreach (var firstRowCell in workSheet.Cells[1, 1, 1, workSheet.Dimension.End.Column])
                            oTbl.Columns.Add(firstRowCell.Text);

                        int rowNumber = 0;

                        for (rowNumber = 2; rowNumber <= workSheet.Dimension.End.Row; rowNumber++)
                        {
                            var row = workSheet.Cells[rowNumber, 1, rowNumber, workSheet.Dimension.End.Column];
                            var newRow = oTbl.NewRow();

                            foreach (var cell in row)
                                newRow[cell.Start.Column - 1] = cell.Text;

                            oTbl.Rows.Add(newRow);
                        }
                    }
                }

                return oTbl;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public void ExcelToDatabase(string path)
        {
            try
            {
                DataTable tb = ReadExcel(path);
                StringBuilder str = new StringBuilder();

                foreach (DataRow row in tb.Rows)
                {
                    str.AppendLine("insert into MH_CHISO_DUBAO('MA_DDO','THANG','NAM','SAN_LUONG','NGAY_TAO','NGUOI_TAO')" +
                        " values ('" + row["MA_DDO"] + "','" + row["THANG"] + "','" + row["NAM"] + "','" + row["SAN_LUONG"] + "','" + DateTime.Now.ToShortDateString() + "','DUY') ;");
                }

                //Datatable to Oracle

                OracleCommand command = new OracleCommand(str.ToString(), dbConnect.Connection);
                command.CommandType = CommandType.Text;
                dbConnect.OpenConn();
                int recordInserted = command.ExecuteNonQuery();
                dbConnect.CloseConn();

                //
            }
            catch (Exception)
            {
                dbConnect.Connection.Dispose();
            }
        }

        #region oracle


        #endregion
    }
}
