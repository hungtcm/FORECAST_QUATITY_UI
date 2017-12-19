using System;
using System.Data;
using System.Windows.Forms;
using FORECAST_QUANTITY_BLL;



namespace FORECAST_QUATITY_UI
{
    public partial class frm_sanluong_hangthang : Form
    {
        public Bus bus = new Bus();

        public frm_sanluong_hangthang()
        {
            InitializeComponent();
        }

        private void btn_capnhat_Click(object sender, EventArgs e)
        {
            bus.ExcelToDatabase(txt_duongdan.Text);
        }

        private void btn_thoat_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        #region helper

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
            if (fDialog.ShowDialog() == DialogResult.OK)
            {
                //set gia tri cho textbox
                txt_duongdan.Text = fDialog.FileName;
                //thuc thi method readexcel
                //BLL.ReadExcel();
            }

        }


        #endregion

    }
}
