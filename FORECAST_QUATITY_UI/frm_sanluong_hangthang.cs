using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using ExcelDataReader;
using xls = Microsoft.Office.Interop.Excel;//import thu vien interop.excel vao de lam viec


namespace FORECAST_QUATITY_UI
{
    public partial class frm_sanluong_hangthang : Form
    {
        //tao mot bang tam de chua du lieu lay tu file excel vao
        private DataTable oTbl;
        //mot file de chua thong tin
        private string fileName;
        DataSet result;
        //method doc du lieu tu file excel vao datatable oTbl
        private void readExcel()
        {
            fileName = txt_duongdan.Text;
            //kiem tra xem du lieu fileName da co chua
            if (fileName == null)
            {
                MessageBox.Show("Chưa chọn file excel");
            }
            else
            {
                //neu da ton tai file excel thi mo file va doc du lieu vao
                xls.Application ExcelApp = new xls.Application();//tao mot app lam viec moi
                //kiem tra xem co mo duoc du lieu tu fileName khong
                try
                {
                    ExcelApp.Workbooks.Open(fileName);

                }
                catch
                {
                    MessageBox.Show("Khong the mo file du lieu");
                }
                //tao cau truc cho Table oTbl
                oTbl = new DataTable();
                oTbl.Columns.Add("A", typeof(string));//tao mot cot ten la MA_DDO co kieu du lieu string
                oTbl.Columns.Add("B", typeof(string));
                oTbl.Columns.Add("C", typeof(string));
                oTbl.Columns.Add("D", typeof(string));
                oTbl.Columns.Add("E", typeof(string));
                
                //doc du lieu tung sheet cua excel
                foreach(xls.Worksheet WSheet in ExcelApp.Worksheets)
                {
                    //tao mot datarow de lay du lieu cho tung cell
                    DataRow dr = oTbl.NewRow();//dataRow co kieu du lieu cung voi oTbl
                    //bien i de doc tung dong trong sheet
                    long i = 2;
                    try
                    {
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

                    }
                    catch
                    {
                        MessageBox.Show("Có lỗi khi thực hiện");
                    }
                }

            }
        }
        public frm_sanluong_hangthang()
        {
            InitializeComponent();
        }
        //viet su kien khi clik nut mo_file
        private void btn_mofile_Click(object sender, EventArgs e)
        {
            //khi click vao nut mo file thi se mo ra hop thoai chon file
            OpenFileDialog fDialog = new OpenFileDialog();//tao mot hop thoai de cho file
            fDialog.Filter = "excel file |*.xls;*.xlsx";//chi lay nhung file co dinh dang xls hoac xlsx
            fDialog.FilterIndex = 1;//tro vao vi tri dau tien cua bo loc
            fDialog.RestoreDirectory = true;//nho duong dan cua lan truy cap truoc
            fDialog.Multiselect = false;//khong cho phep chon nhieu file cung luc
            fDialog.Title = "Chon file excel";//tieu de cua hop thoai
            // neu chon file thanh cong
            if(fDialog.ShowDialog() == DialogResult.OK)
            {
                //set gia tri cho textbox
                txt_duongdan.Text = fDialog.FileName;
                //thuc thi method readexcel
                readExcel();
            }

        }

        private void btn_capnhat_Click(object sender, EventArgs e)
        {
            //truyen du lieu vao datagrid
            if (oTbl != null)
            {
                dtg_excel.DataSource = oTbl;
            }
            else
            {
                MessageBox.Show("Khong co du lieu");

            }
        }

        private void btn_thoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btn_browse2_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xls;*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    FileStream fs = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                    IExcelDataReader reader = ExcelReaderFactory.CreateOpenXmlReader(fs);
                    //reader.IsFirstRowASColumnNames = true;
                    DataSet result = reader.AsDataSet();
                    comboBox1.Items.Clear();
                    foreach (DataTable dt in result.Tables)
                        comboBox1.Items.Add(dt.TableName);
                    reader.Close();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dtg_excel.DataSource = result.Tables[comboBox1.SelectedIndex];
        }
    }
}
