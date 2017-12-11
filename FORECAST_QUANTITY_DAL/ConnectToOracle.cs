using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
//using System.Data.OleDb;
using Oracle.DataAccess.Client;
using Oracle.DataAccess.Types;

namespace CT_THOATHUANTHANHTOAN_PCTHUDUC.Model
{
    class ConnectToOracle
    {
        #region Availible
        //private OleDbConnection Conn;
        private OracleConnection Conn;
        //private OleDbCommand _cmd;
        private OracleCommand _cmd;

        public OracleCommand Cmd
        {
            get { return _cmd; }
            set { _cmd = value; }
        }

        /*public OleDbCommand Cmd
        {
            get { return _cmd; }
            set { _cmd = value; }
        }*/

        //public OleDbConnection Connection { get { return Conn; } }
        public OracleConnection Connection { get { return Conn; } }
        private string error;
        public string Error
        {
            get { return error; }
            set { error = value; }
        }
        string StrCon;

        #endregion

        #region Contructor
        public ConnectToOracle()
        {

            //StrCon = @"Provider=MSDAORA;DATA SOURCE=vtg-cmis-scan.hcmpc.com.vn:1521/CMIS2;USER ID=QTVUTHUY;Password=conlaumoibiet;Unicode=True";
            //StrCon = @"Provider=MSDAORA;DATA SOURCE=MINHHUNG;USER ID=QTVUTHUY;Password=conlaumoibiet;Unicode=True";
            //StrCon = @"Provider=MSDAORA;Data Source=MINHHUNG;User ID=DLTHUDUC;Password=hcm#pc357;Unicode=True";

            //StrCon = @"Data Source = (DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST = vtg-cmis-scan.hcmpc.com.vn))(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = CMISAPP)));User Id = QTVUTHUY; Password = conlaumoibiet";
            //StrCon = @"SERVER = (DESCRIPTION = (ADDRESS = (PROTOCOL = TCP)(HOST =vtg-cmis-scan.hcmpc.com.vn)(PORT = 1521))(CONNECT_DATA = (SERVICE_NAME = CMISAPP)));uid = QTVUTHUY; pwd = conlaumoibiet; Unicode=True";
            StrCon = "Data Source = CMIS2TD; User Id = QTVUTHUY; Password = conlaumoibiet";
            Conn = new OracleConnection(StrCon);
            //Conn = new OleDbConnection(StrCon);
        }

        #endregion

        #region Methods

        public bool OpenConn()
        {
            try
            {
                if (Conn.State == ConnectionState.Closed)
                    Conn.Open();
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            return true;
        }

        public bool CloseConn()
        {
            try
            {
                if (Conn.State == ConnectionState.Open)
                    Conn.Close();
            }
            catch (Exception ex)
            {
                error = ex.Message;
                return false;
            }
            return true;
        }

        #endregion 
    }
}
