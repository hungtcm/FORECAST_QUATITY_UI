using System;
using System.Data;
using FORECAST_QUANTITY_DTO;
using OfficeOpenXml;
using System.IO;
using System.Linq;
using System.Text;
using Oracle.DataAccess;
using Oracle.DataAccess.Client;
using System.Collections.Generic;

namespace FORECAST_QUANTITY_DAL
{
    public class DAL_Excel
    {

        public DBConnect dbConnect = new DBConnect();
        private DataTable ReadExcel(string path)
        {
            DataTable oTbl = new DataTable();

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
                StringBuilder str = new StringBuilder();
                str.Append("INSERT ALL ");

                DataTable tb = ReadExcel(path); // content of excel
                List<Customer> customerList = new List<Customer>();
                List<ChiSo> chisoList = new List<ChiSo>();

                for (int i = 0; i < tb.Rows.Count; i++)
                {
                    Customer customer = new Customer();
                    ChiSo cs = new ChiSo();

                    customer.MA_DDO = (tb.Rows[i]["MA_DDO"]).ToString();
                    customer.MA_SOGCS = (tb.Rows[i]["MA_SOGCS"]).ToString();
                    customer.STT = Convert.ToInt32(tb.Rows[i]["STT"]);
                    customer.TEN_KHANG = (tb.Rows[i]["TEN_KHANG"]).ToString();
                    customer.DCHI_HDON = (tb.Rows[i]["DCHI_HDON"]).ToString();
                    customer.SO_TBI = (tb.Rows[i]["SO_TBI"]).ToString();
                    customer.MA_CAPDA = Convert.ToInt32(tb.Rows[i]["MA_CAPDA"]);
                    customer.KIMUA_CSPK = Convert.ToInt32(tb.Rows[i]["KIMUA_CSPK"]);
                    customer.SO_HO = Convert.ToInt32(tb.Rows[i]["SO_HO"]);
                    customer.MA_TRAM = (tb.Rows[i]["MA_TRAM"]).ToString();
                    customer.TEN_TRAM = (tb.Rows[i]["TEN_TRAM"]).ToString();
                    customer.LOAI_TRAM = (tb.Rows[i]["LOAI_TRAM"]).ToString();
                    customer.DINH_DANH = (tb.Rows[i]["DINH_DANH"]).ToString();
                    customer.SO_DTHOAI = (tb.Rows[i]["SO_DTHOAI"]).ToString();
                    customer.ID_HDONG = (tb.Rows[i]["ID_HDONG"]).ToString();
                    customer.TEN_DVIDCHINH = (tb.Rows[i]["TEN_DVIDCHINH"]).ToString();
                    customer.MA_KHANG = (tb.Rows[i]["MA_KHANG"]).ToString();
                    customer.NGAY_TAO = DateTime.Now;
                    customer.NGUOI_TAO = "Duy"; //dua tu login vao

                    customerList.Add(customer);

                    cs.MA_DDO = (tb.Rows[i]["MA_DDO"]).ToString();
                    cs.NAM = Convert.ToInt32(tb.Rows[i]["NAM"]);
                    cs.THANG = Convert.ToInt32(tb.Rows[i]["THANG"]);
                    cs.SAN_LUONG = Convert.ToDouble(tb.Rows[i]["SAN_LUONG"]);
                    cs.NGAY_TAO = DateTime.Now;
                    cs.NGUOI_TAO = "DUY";
                    chisoList.Add(cs);
                }

                foreach (var item in chisoList)
                    str.AppendLine("into MH_CHISO_DUBAO(MA_DDO,THANG,NAM,SAN_LUONG,NGAY_TAO,NGUOI_TAO) values ('" + item.MA_DDO + "'," + item.THANG + "," + item.NAM + "," + item.SAN_LUONG + ", to_date('" + DateTime.Now + "','MM/DD/YYYY HH:MI:SS AM'), 'DUY') ");

                str.AppendLine("SELECT 1 FROM DUAL ");

                OracleCommand command_chiso = new OracleCommand(str.ToString(), dbConnect.Connection);
                command_chiso.CommandType = CommandType.Text;
                dbConnect.OpenConn();
                int inserted = command_chiso.ExecuteNonQuery();
                dbConnect.CloseConn();
            }
            catch (Exception Ex)
            {
                return;
            }
        }

        #region oracle


        #endregion
    }
}
