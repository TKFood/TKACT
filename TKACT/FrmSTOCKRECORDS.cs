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
using System.Data.SqlClient;
using System.Configuration;
using System.Text.RegularExpressions;
using FastReport;
using FastReport.Data;
using TKITDLL;
using System.Data.OleDb;
using System.Net;

namespace TKACT
{
    public partial class FrmSTOCKRECORDS : Form
    {
        StringBuilder sbSql = new StringBuilder();
       
        SqlDataAdapter adapter = new SqlDataAdapter();
        SqlCommandBuilder sqlCmdBuilder = new SqlCommandBuilder();
        SqlTransaction tran;
        SqlCommand cmd = new SqlCommand();
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        string tablename = null;
        int rownum = 0;
        DataGridViewRow row;
        int result;
        string SortedColumn = string.Empty;
        string SortedModel = string.Empty;


        public FrmSTOCKRECORDS()
        {
            InitializeComponent();
        }

        #region FUNCTION

        public void Search(string STOCKACCOUNTNUMBER,string STOCKNAME)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();

            if(!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                sbSqlQuery1.AppendFormat(@" AND STOCKACCOUNTNUMBER LIKE '%{0}%'", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if(!string.IsNullOrEmpty(STOCKNAME))
            {
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE '%{0}%'", STOCKNAME);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }


            sbSql.AppendFormat(@"
                                SELECT 
                                [STOCKACCOUNTNUMBER] AS '戶號'
                                ,[STOCKNAME] AS '股東姓名'
                                ,[IDNUMBER] AS '身份證字號或統一編號'
                                ,[POSTALCODE] AS '通訊地郵遞區號'
                                ,[MAILINGADDRESS] AS '通訊地址'
                                ,[REGISTEREDPOSTALCODE] AS '戶籍地郵遞區號'
                                ,[REGISTEREDADDRESS] AS '戶籍地址'
                                ,[DATEOFBIRTH] AS '出生/設立日期'
                                ,[BANKNAME] AS '銀行名稱'
                                ,[BRANCHNAME] AS '分行名稱'
                                ,[BANKCODE] AS '銀行代碼'
                                ,[ACCOUNTNUMBER] AS '帳號'
                                ,[HOMEPHONENUMBER] AS '住家電話'
                                ,[MOBILEPHONENUMBER] AS '手機號碼'
                                ,[EMAIL] AS 'e-mail'
                                ,[PASSPORTNUMBER] AS '護照號碼'
                                ,[ENGLISHNAME] AS '英文名'
                                ,[FATHER] AS '父'
                                ,[MOTHER] AS '母'
                                ,[SPOUSE] AS '配偶'
                                ,CONVERT(nvarchar,[CREATEDATES],112) AS '建立時間'
                                FROM [TKACT].[dbo].[TKSTOCKS]
                                WHERE 1=1
                                {0}
                                {1}
                                ORDER BY [STOCKACCOUNTNUMBER] 
                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString());
            sbSql.AppendFormat(@"  ");

            SEARCH_MANULINE(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);
        }

        #endregion
        public void SEARCH_MANULINE(string QUERY, DataGridView DataGridViewNew, string SortedColumn, string SortedModel)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            SqlDataAdapter SqlDataAdapterNEW = new SqlDataAdapter();
            SqlCommandBuilder SqlCommandBuilderNEW = new SqlCommandBuilder();
            DataSet DataSetNEW = new DataSet();

            try
            {
                //20210902密
                Class1 TKID = new Class1();//用new 建立類別實體
                SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

                //資料庫使用者密碼解密
                sqlsb.Password = TKID.Decryption(sqlsb.Password);
                sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

                String connectionString;
                sqlConn = new SqlConnection(sqlsb.ConnectionString);

                SqlDataAdapterNEW = new SqlDataAdapter(@"" + sbSql, sqlConn);

                SqlCommandBuilderNEW = new SqlCommandBuilder(SqlDataAdapterNEW);
                sqlConn.Open();
                DataSetNEW.Clear();
                SqlDataAdapterNEW.Fill(DataSetNEW, "DataSetNEW");
                sqlConn.Close();


                DataGridViewNew.DataSource = null;

                if (DataSetNEW.Tables["DataSetNEW"].Rows.Count >= 1)
                {
                    //DataGridViewNew.Rows.Clear();
                    DataGridViewNew.DataSource = DataSetNEW.Tables["DataSetNEW"];
                    DataGridViewNew.AutoResizeColumns();
                    //DataGridViewNew.CurrentCell = dataGridView1[0, rownum];
                    //dataGridView20SORTNAME
                    //dataGridView20SORTMODE
                    if (!string.IsNullOrEmpty(SortedColumn))
                    {
                        if (SortedModel.Equals("Ascending"))
                        {
                            DataGridViewNew.Sort(DataGridViewNew.Columns["" + SortedColumn + ""], ListSortDirection.Ascending);
                        }
                        else
                        {
                            DataGridViewNew.Sort(DataGridViewNew.Columns["" + SortedColumn + ""], ListSortDirection.Descending);
                        }
                    }
                }

            }
            catch
            {

            }
            finally
            {
                sqlConn.Close();
            }
        }

        public void CHECKADD()
        {
            string MESSAGES = "";

            //戶號
            if (string.IsNullOrEmpty(textBox3.Text))
            {
                MESSAGES = MESSAGES + "戶號 不可空白";
            }
            if (!string.IsNullOrEmpty(textBox3.Text))
            {
                string input = textBox3.Text;
                int number;

                if (input.Length == 4 && int.TryParse(input, out number))
                {
                    // 輸入為 4 位數字
                    // 在這裡處理符合條件的情況
                }
                else
                {
                    MESSAGES = MESSAGES + " 戶號 要為4位數字 ";
                }
               
            }

            //股東姓名
            if (string.IsNullOrEmpty(textBox4.Text))
            {
                MESSAGES = MESSAGES + " 股東姓名 不可空白";
            }
            if (!string.IsNullOrEmpty(textBox4.Text))
            {
                string input = textBox4.Text;     

                if (input.Length >=3 )
                {
                    // 輸入為 4 位數字
                    // 在這裡處理符合條件的情況
                }
                else
                {
                    MESSAGES = MESSAGES + " 股東姓名 至少3個中文字 ";
                }

            }

            //身份證字號或統一編號
            //通訊地郵遞區號
            //通訊地址
            //戶籍地郵遞區號
            //戶籍地址
            //出生 / 設立日期
            //銀行名稱
            //分行名稱
            //銀行代碼
            //帳號
            //住家電話
            //手機號碼
            //e - mail
            //護照號碼
            //英文名
            //父
            //母
            //配偶



            //MESSAGES
            if (!string.IsNullOrEmpty(MESSAGES))
            {
                MessageBox.Show(MESSAGES);
            }
            
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
           
        }
        private void textBox3_Leave(object sender, EventArgs e)
        {
            CHECKADD();
        }
        private void textBox4_Leave(object sender, EventArgs e)
        {
            CHECKADD();
        }


        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox1.Text,textBox2.Text);
        }
        private void button2_Click(object sender, EventArgs e)
        {

        }



        #endregion

     
    }
}
