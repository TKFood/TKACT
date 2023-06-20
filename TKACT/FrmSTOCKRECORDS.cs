﻿using System;
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

            comboBox1load();
            comboBox2load();

        }

        #region FUNCTION
        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT  [ID],[KINDS],[NAMES],[KEYS] FROM [TKACT].[dbo].[TBPARAS] WHERE KINDS='異動原因' ORDER BY ID", "KEYS", "KEYS");
        }
        public void comboBox2load()
        {
            LoadComboBoxData(comboBox2, "SELECT  [ID],[KINDS],[NAMES],[KEYS] FROM [TKACT].[dbo].[TBPARAS] WHERE KINDS='異動原因' ORDER BY ID", "KEYS", "KEYS");
        }
        public void LoadComboBoxData(ComboBox comboBox, string query, string valueMember, string displayMember)
        {
            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
            {
                SqlCommand command = new SqlCommand(query, connection);
                connection.Open();

                SqlDataAdapter adapter = new SqlDataAdapter(command);
                DataTable dataTable = new DataTable();
                adapter.Fill(dataTable);

                comboBox.DataSource = dataTable;
                comboBox.ValueMember = valueMember;
                comboBox.DisplayMember = displayMember;
            }
        }
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
                                ,'民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([DATEOFBIRTH],1,4))-1911))+'年'+SUBSTRING([DATEOFBIRTH],6,2)+'月'+SUBSTRING([DATEOFBIRTH],9,2) +'日' AS '出生/設立日期'
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
                                ,[COMMENTS] AS '備註'
                                ,CONVERT(nvarchar,[CREATEDATES],112) AS '建立時間'
                                ,[DATEOFBIRTH] 
                                FROM [TKACT].[dbo].[TKSTOCKS]
                                WHERE 1=1
                                {0}
                                {1}
                                ORDER BY [STOCKACCOUNTNUMBER] 
                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString());
            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView1, SortedColumn, SortedModel);
        }

        public void Search_DG2(string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                sbSqlQuery1.AppendFormat(@" AND STOCKACCOUNTNUMBER LIKE '%{0}%'", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
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
                                ,'民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([DATEOFBIRTH],1,4))-1911))+'年'+SUBSTRING([DATEOFBIRTH],6,2)+'月'+SUBSTRING([DATEOFBIRTH],9,2) +'日' AS '出生/設立日期'
                               
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
                                ,[COMMENTS] AS '備註'
                                ,CONVERT(nvarchar,[CREATEDATES],112) AS '建立時間'
                                ,[DATEOFBIRTH] 
                                ,[ID]
                                FROM [TKACT].[dbo].[TKSTOCKS]
                                WHERE 1=1
                                {0}
                                {1}
                                ORDER BY [STOCKACCOUNTNUMBER] 
                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString());
            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView2, SortedColumn, SortedModel);
        }

        public void Search_DG3(string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();
            sbSqlQuery3.Clear();

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                sbSqlQuery1.AppendFormat(@" AND STOCKACCOUNTNUMBER LIKE '%{0}%'", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE '%{0}%'", STOCKNAME);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }
      


            sbSql.AppendFormat(@"
                                SELECT 
                                [SERNO] AS '流水號'
                                ,[STOCKACCOUNTNUMBER] AS '戶號'
                                ,[STOCKNAME] AS '股東姓名'
                                ,[IDNUMBER] AS '身份證字號或統一編號'
                                ,[POSTALCODE] AS '通訊地郵遞區號'
                                ,[MAILINGADDRESS] AS '通訊地址'
                                ,[REGISTEREDPOSTALCODE] AS '戶籍地郵遞區號'
                                ,[REGISTEREDADDRESS] AS '戶籍地址'
                                ,'民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([DATEOFBIRTH],1,4))-1911))+'年'+SUBSTRING([DATEOFBIRTH],6,2)+'月'+SUBSTRING([DATEOFBIRTH],9,2) +'日' AS '出生/設立日期'
                                
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
                                ,[COMMENTS] AS '備註'
                                ,CONVERT(nvarchar,[CREATEDATES],112) AS '建立時間'
                                ,[ISUPDATE] AS '是否更新'
                                ,[DATEOFBIRTH] 
                                FROM [TKACT].[dbo].[TKSTOCKSCHAGES]
                                WHERE 1=1
                                {0}
                                {1}
                                {2}
                                ORDER BY [STOCKACCOUNTNUMBER] 
                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString());
            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView3, SortedColumn, SortedModel);
        }

        public void Search_DG4(string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();
            sbSqlQuery3.Clear();

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                sbSqlQuery1.AppendFormat(@" AND STOCKACCOUNTNUMBER LIKE '%{0}%'", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE '%{0}%'", STOCKNAME);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }



            sbSql.AppendFormat(@"
                                
                                SELECT 
                                '民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([CAPITALINCREASERECORDDATE],1,4))-1911))+'年'+SUBSTRING([CAPITALINCREASERECORDDATE],6,2)+'月'+SUBSTRING([CAPITALINCREASERECORDDATE],9,2) +'日' AS '增資基準日'
                                ,[REASONFORCHANGE] AS '異動原因'
                                ,[STOCKACCOUNTNUMBER] AS '戶號'
                                ,[STOCKNAME] AS '股東姓名'
                                ,[INCREASEDSHARES] AS '增資股數'
                                ,[PARVALUPERSHARE] AS '每股面額'
                                ,[TRADINGPRICEPERSHARE] AS '每股成交價格'
                                ,[TOTALTRADINGAMOUNT] AS '成交總額'
                                ,[INCREASEDSHARESHUNDREDTHOUSANDS] AS '增資股票號碼(十萬股)'
                                ,[INCREASEDSHARESTENSOFTHOUSANDS] AS '增資股票號碼(萬股)'
                                ,[INCREASEDSHARESTHOUSANDS] AS '增資股票號碼(千股)'
                                ,[INCREASEDSHARESIRREGULARLOTS] AS '增資股票號碼(不定額股)'
                                ,[HOLDINGSHARES] AS '持有股數'
                                ,[SERNO]
                                ,[CAPITALINCREASERECORDDATE]

                                FROM [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                WHERE 1=1
                                {0}
                                {1}
                                ORDER BY [STOCKACCOUNTNUMBER],SERNO

                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString());

            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView4, SortedColumn, SortedModel);
        }

        public void Search_DG5(string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();
            sbSqlQuery3.Clear();

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                sbSqlQuery1.AppendFormat(@" AND STOCKACCOUNTNUMBER LIKE '%{0}%'", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE '%{0}%'", STOCKNAME);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }



            sbSql.AppendFormat(@"
                                SELECT
                                [SERNO] AS '流水號'
                                ,'民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([DATEOFCHANGE],1,4))-1911))+'年'+SUBSTRING([DATEOFCHANGE],6,2)+'月'+SUBSTRING([DATEOFCHANGE],9,2) +'日' AS '異動日期'
                                ,[REASOFORCHANGE] AS '異動原因'
                                ,[STOCKACCOUNTNUMBERFORM] AS '轉讓人戶號'
                                ,[STOCKNAMEFORM] AS '轉讓人姓名'
                                ,[STOCKACCOUNTNUMBERTO] AS '受讓人戶號'
                                ,[STOCKNAMETO] AS '受讓人姓名'
                                ,[TRANSFERREDSHARES] AS '轉讓股數'
                                ,[PARVALUEPERSHARE] AS '每股面額'
                                ,[TRADINGPRICEPERSHARE] AS '每股成交價格'
                                ,[TOTALTRADINGAMOUNT] AS '成交總額'
                                ,[SECURITIESTRANSACTIONTAXAMOUNT] AS '證券交易稅額'
                                ,[TRANSFERREDSHARESHUNDREDTHOUSANDS] AS '轉讓股票號碼(十萬股)'
                                ,[TRANSFERREDSHARESTENSOFTHOUSANDS] AS '轉讓股票號碼(萬股)'
                                ,[TRANSFERREDSHARESTHOUSANDS] AS '轉讓股票號碼(千股)'
                                ,[TRANSFERREDSHARESIRREGULARLOTS] AS '轉讓股票號碼(不定額股)'
                                ,[HOLDINGSHARES] AS '持有股數'
                                ,[IDFORM]
                                ,[IDTO]
                                ,[DATEOFCHANGE]
                                FROM [TKACT].[dbo].[TKSTOCKSTRANS]
                                WHERE 1=1
                                {0}
                                {1}
                                ORDER BY [DATEOFCHANGE], [SERNO]

                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString());

            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView5, SortedColumn, SortedModel);
        }

        public void SEARCH(string QUERY, DataGridView DataGridViewNew, string SortedColumn, string SortedModel)
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

        public void TKSTOCKS_ADD(
                                string CREATEDATES
                                , string STOCKACCOUNTNUMBER
                                , string STOCKNAME
                                , string IDNUMBER
                                , string POSTALCODE
                                , string MAILINGADDRESS
                                , string REGISTEREDPOSTALCODE
                                , string REGISTEREDADDRESS
                                , string DATEOFBIRTH
                                , string BANKNAME
                                , string BRANCHNAME
                                , string BANKCODE
                                , string ACCOUNTNUMBER
                                , string HOMEPHONENUMBER
                                , string MOBILEPHONENUMBER
                                , string EMAIL
                                , string PASSPORTNUMBER
                                , string ENGLISHNAME
                                , string FATHER
                                , string MOTHER
                                , string SPOUSE
                                , string COMMENTS
            )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKACT].[dbo].[TKSTOCKS]
                                    (
                                    [CREATEDATES]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[IDNUMBER]
                                    ,[POSTALCODE]
                                    ,[MAILINGADDRESS]
                                    ,[REGISTEREDPOSTALCODE]
                                    ,[REGISTEREDADDRESS]
                                    ,[DATEOFBIRTH]
                                    ,[BANKNAME]
                                    ,[BRANCHNAME]
                                    ,[BANKCODE]
                                    ,[ACCOUNTNUMBER]
                                    ,[HOMEPHONENUMBER]
                                    ,[MOBILEPHONENUMBER]
                                    ,[EMAIL]
                                    ,[PASSPORTNUMBER]
                                    ,[ENGLISHNAME]
                                    ,[FATHER]
                                    ,[MOTHER]
                                    ,[SPOUSE]
                                    ,[COMMENTS]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    ,'{14}'
                                    ,'{15}'
                                    ,'{16}'
                                    ,'{17}'
                                    ,'{18}'
                                    ,'{19}'
                                    ,'{20}'
                                    ,'{21}'
                                    )
                                        
                                        ", CREATEDATES
                                        , STOCKACCOUNTNUMBER
                                        , STOCKNAME
                                        , IDNUMBER
                                        , POSTALCODE
                                        , MAILINGADDRESS
                                        , REGISTEREDPOSTALCODE
                                        , REGISTEREDADDRESS
                                        , DATEOFBIRTH
                                        , BANKNAME
                                        , BRANCHNAME
                                        , BANKCODE
                                        , ACCOUNTNUMBER
                                        , HOMEPHONENUMBER
                                        , MOBILEPHONENUMBER
                                        , EMAIL
                                        , PASSPORTNUMBER
                                        , ENGLISHNAME
                                        , FATHER
                                        , MOTHER
                                        , SPOUSE
                                        , COMMENTS
                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");

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

        public void TKSTOCKSCHAGES_ADD(
                               string CREATEDATES
                               , string STOCKACCOUNTNUMBER
                               , string STOCKNAME
                               , string IDNUMBER
                               , string POSTALCODE
                               , string MAILINGADDRESS
                               , string REGISTEREDPOSTALCODE
                               , string REGISTEREDADDRESS
                               , string DATEOFBIRTH
                               , string BANKNAME
                               , string BRANCHNAME
                               , string BANKCODE
                               , string ACCOUNTNUMBER
                               , string HOMEPHONENUMBER
                               , string MOBILEPHONENUMBER
                               , string EMAIL
                               , string PASSPORTNUMBER
                               , string ENGLISHNAME
                               , string FATHER
                               , string MOTHER
                               , string SPOUSE
                               , string ISUPDATE
                               , string ID
                               , string COMMENTS
           )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@" 
                                    INSERT INTO [TKACT].[dbo].[TKSTOCKSCHAGES]
                                    (
                                    [CREATEDATES]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[IDNUMBER]
                                    ,[POSTALCODE]
                                    ,[MAILINGADDRESS]
                                    ,[REGISTEREDPOSTALCODE]
                                    ,[REGISTEREDADDRESS]
                                    ,[DATEOFBIRTH]
                                    ,[BANKNAME]
                                    ,[BRANCHNAME]
                                    ,[BANKCODE]
                                    ,[ACCOUNTNUMBER]
                                    ,[HOMEPHONENUMBER]
                                    ,[MOBILEPHONENUMBER]
                                    ,[EMAIL]
                                    ,[PASSPORTNUMBER]
                                    ,[ENGLISHNAME]
                                    ,[FATHER]
                                    ,[MOTHER]
                                    ,[SPOUSE]
                                    ,[ISUPDATE]
                                    ,[ID]
                                    ,[COMMENTS]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                    ,'{14}'
                                    ,'{15}'
                                    ,'{16}'
                                    ,'{17}'
                                    ,'{18}'
                                    ,'{19}'
                                    ,'{20}'
                                    ,'{21}'
                                    ,'{22}'
                                    ,'{23}'
                                    )
                                        
                                        ", CREATEDATES
                                        , STOCKACCOUNTNUMBER
                                        , STOCKNAME
                                        , IDNUMBER
                                        , POSTALCODE
                                        , MAILINGADDRESS
                                        , REGISTEREDPOSTALCODE
                                        , REGISTEREDADDRESS
                                        , DATEOFBIRTH
                                        , BANKNAME
                                        , BRANCHNAME
                                        , BANKCODE
                                        , ACCOUNTNUMBER
                                        , HOMEPHONENUMBER
                                        , MOBILEPHONENUMBER
                                        , EMAIL
                                        , PASSPORTNUMBER
                                        , ENGLISHNAME
                                        , FATHER
                                        , MOTHER
                                        , SPOUSE
                                        , ISUPDATE
                                        , ID
                                        , COMMENTS
                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    //MessageBox.Show("完成");

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

        public void UPDATE_TO_TKSTOCKS(
            string ID
            , string STOCKACCOUNTNUMBER
            , string STOCKNAME
            , string IDNUMBER
            , string POSTALCODE
            , string MAILINGADDRESS
            , string REGISTEREDPOSTALCODE
            , string REGISTEREDADDRESS
            , string DATEOFBIRTH
            , string BANKNAME
            , string BRANCHNAME
            , string BANKCODE
            , string ACCOUNTNUMBER
            , string HOMEPHONENUMBER
            , string MOBILEPHONENUMBER
            , string EMAIL
            , string PASSPORTNUMBER
            , string ENGLISHNAME
            , string FATHER
            , string MOTHER
            , string SPOUSE
            , string COMMENTS
            )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                    
                                    UPDATE [TKACT].[dbo].[TKSTOCKS]
                                    SET 
                                    [STOCKACCOUNTNUMBER]='{1}'
                                    ,[STOCKNAME]='{2}'
                                    ,[IDNUMBER]='{3}'
                                    ,[POSTALCODE]='{4}'
                                    ,[MAILINGADDRESS]='{5}'
                                    ,[REGISTEREDPOSTALCODE]='{6}'
                                    ,[REGISTEREDADDRESS]='{7}'
                                    ,[DATEOFBIRTH]='{8}'
                                    ,[BANKNAME]='{9}'
                                    ,[BRANCHNAME]='{10}'
                                    ,[BANKCODE]='{11}'
                                    ,[ACCOUNTNUMBER]='{12}'
                                    ,[HOMEPHONENUMBER]='{13}'
                                    ,[MOBILEPHONENUMBER]='{14}'
                                    ,[EMAIL]='{15}'
                                    ,[PASSPORTNUMBER]='{16}'
                                    ,[ENGLISHNAME]='{17}'
                                    ,[FATHER]='{18}'
                                    ,[MOTHER]='{19}'
                                    ,[SPOUSE]='{20}'
                                    ,[COMMENTS]='{21}'
                                    WHERE [ID]='{0}'
                                        
                                        ", ID
                                        , STOCKACCOUNTNUMBER
                                        , STOCKNAME
                                        , IDNUMBER
                                        , POSTALCODE
                                        , MAILINGADDRESS
                                        , REGISTEREDPOSTALCODE
                                        , REGISTEREDADDRESS
                                        , DATEOFBIRTH
                                        , BANKNAME
                                        , BRANCHNAME
                                        , BANKCODE
                                        , ACCOUNTNUMBER
                                        , HOMEPHONENUMBER
                                        , MOBILEPHONENUMBER
                                        , EMAIL
                                        , PASSPORTNUMBER
                                        , ENGLISHNAME
                                        , FATHER
                                        , MOTHER
                                        , SPOUSE
                                        , COMMENTS
                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");

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

        private void dataGridView2_SelectionChanged(object sender, EventArgs e)
        {
            string ID = "";
            textBox26.Text = "";
            textBox27.Text = "";
            textBox28.Text = "";
            textBox29.Text = "";
            textBox30.Text = "";
            textBox31.Text = "";
            textBox32.Text = "";                
            textBox25.Text = "";
            textBox33.Text = "";
            textBox34.Text = "";
            textBox35.Text = "";
            textBox24.Text = "";
            textBox36.Text = "";
            textBox37.Text = "";
            textBox38.Text = "";
            textBox39.Text = "";
            textBox40.Text = "";
            textBox41.Text = "";
            textBox42.Text = "";
            textBox43.Text = "";
            textBox71.Text = "";

            if (dataGridView2.CurrentRow != null)
            {
                int rowindex = dataGridView2.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView2.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();
                    textBox26.Text = row.Cells["戶號"].Value.ToString();
                    textBox27.Text = row.Cells["股東姓名"].Value.ToString();
                    textBox28.Text = row.Cells["身份證字號或統一編號"].Value.ToString();
                    textBox29.Text = row.Cells["通訊地郵遞區號"].Value.ToString();
                    textBox30.Text = row.Cells["通訊地址"].Value.ToString();
                    textBox31.Text = row.Cells["戶籍地郵遞區號"].Value.ToString();
                    textBox32.Text = row.Cells["戶籍地址"].Value.ToString();
                    textBox25.Text = row.Cells["銀行名稱"].Value.ToString();
                    textBox33.Text = row.Cells["分行名稱"].Value.ToString();
                    textBox34.Text = row.Cells["銀行代碼"].Value.ToString();
                    textBox35.Text = row.Cells["帳號"].Value.ToString();
                    textBox24.Text = row.Cells["住家電話"].Value.ToString();
                    textBox36.Text = row.Cells["手機號碼"].Value.ToString();
                    textBox37.Text = row.Cells["e-mail"].Value.ToString();
                    textBox38.Text = row.Cells["護照號碼"].Value.ToString();
                    textBox39.Text = row.Cells["英文名"].Value.ToString();
                    textBox40.Text = row.Cells["父"].Value.ToString();
                    textBox41.Text = row.Cells["母"].Value.ToString();
                    textBox42.Text = row.Cells["配偶"].Value.ToString();
                    textBox71.Text = row.Cells["備註"].Value.ToString();
                    textBox43.Text = row.Cells["ID"].Value.ToString();

                    dateTimePicker2.Value = Convert.ToDateTime(row.Cells["DATEOFBIRTH"].Value.ToString());

                }
                else
                {


                }
            }
        }

        private void dataGridView4_SelectionChanged(object sender, EventArgs e)
        {
            string SERNO = "";
            textBox46.Text = "";
            textBox47.Text = "";
            textBox48.Text = "";
            textBox49.Text = "";
            textBox50.Text = "";
            textBox51.Text = "";
            textBox52.Text = "";
            textBox53.Text = "";
            textBox54.Text = "";
            textBox55.Text = "";
            textBox56.Text = "";
            textBox57.Text = "";

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    SERNO = row.Cells["SERNO"].Value.ToString();
                    textBox58.Text = row.Cells["SERNO"].Value.ToString();
                    textBox59.Text = row.Cells["戶號"].Value.ToString();
                    textBox60.Text = row.Cells["股東姓名"].Value.ToString();
                    textBox61.Text = row.Cells["增資股數"].Value.ToString();
                    textBox62.Text = row.Cells["每股面額"].Value.ToString();
                    textBox63.Text = row.Cells["每股成交價格"].Value.ToString();
                    textBox64.Text = row.Cells["成交總額"].Value.ToString();
                    textBox65.Text = row.Cells["增資股票號碼(十萬股)"].Value.ToString();
                    textBox66.Text = row.Cells["增資股票號碼(萬股)"].Value.ToString();
                    textBox67.Text = row.Cells["增資股票號碼(千股)"].Value.ToString();
                    textBox68.Text = row.Cells["增資股票號碼(不定額股)"].Value.ToString();
                    textBox69.Text = row.Cells["持有股數"].Value.ToString();

                    dateTimePicker4.Value= Convert.ToDateTime(row.Cells["CAPITALINCREASERECORDDATE"].Value.ToString());
                    
                    comboBox2.SelectedValue = row.Cells["異動原因"].Value.ToString();


                }
            }
        }

        public void TKSTOCKSTRANSADD_ADD(
             string CAPITALINCREASERECORDDATE
            , string REASONFORCHANGE
            , string STOCKACCOUNTNUMBER
            , string STOCKNAME
            , string INCREASEDSHARES
            , string PARVALUPERSHARE
            , string TRADINGPRICEPERSHARE
            , string TOTALTRADINGAMOUNT
            , string INCREASEDSHARESHUNDREDTHOUSANDS
            , string INCREASEDSHARESTENSOFTHOUSANDS
            , string INCREASEDSHARESTHOUSANDS
            , string INCREASEDSHARESIRREGULARLOTS
            , string HOLDINGSHARES
            , string ID
            )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                  
                                   
                                    INSERT INTO [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                    (
                                   [CAPITALINCREASERECORDDATE]
                                    ,[REASONFORCHANGE]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[INCREASEDSHARES]
                                    ,[PARVALUPERSHARE]
                                    ,[TRADINGPRICEPERSHARE]
                                    ,[TOTALTRADINGAMOUNT]
                                    ,[INCREASEDSHARESHUNDREDTHOUSANDS]
                                    ,[INCREASEDSHARESTENSOFTHOUSANDS]
                                    ,[INCREASEDSHARESTHOUSANDS]
                                    ,[INCREASEDSHARESIRREGULARLOTS]
                                    ,[HOLDINGSHARES]
                                    ,[ID]
                                    )
                                    VALUES
                                    (
                                    '{0}'
                                    ,'{1}'
                                    ,'{2}'
                                    ,'{3}'
                                    ,'{4}'
                                    ,'{5}'
                                    ,'{6}'
                                    ,'{7}'
                                    ,'{8}'
                                    ,'{9}'
                                    ,'{10}'
                                    ,'{11}'
                                    ,'{12}'
                                    ,'{13}'
                                
                                    )
                                        
                                        ",
                                         CAPITALINCREASERECORDDATE
                                        , REASONFORCHANGE
                                        , STOCKACCOUNTNUMBER
                                        , STOCKNAME
                                        , INCREASEDSHARES
                                        , PARVALUPERSHARE
                                        , TRADINGPRICEPERSHARE
                                        , TOTALTRADINGAMOUNT
                                        , INCREASEDSHARESHUNDREDTHOUSANDS
                                        , INCREASEDSHARESTENSOFTHOUSANDS
                                        , INCREASEDSHARESTHOUSANDS
                                        , INCREASEDSHARESIRREGULARLOTS
                                        , HOLDINGSHARES
                                        , ID
                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");

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

        public void TKSTOCKSTRANSADD_UPDATE(
          string SERNO
          , string CAPITALINCREASERECORDDATE
          , string REASONFORCHANGE
          , string STOCKACCOUNTNUMBER
          , string STOCKNAME
          , string INCREASEDSHARES
          , string PARVALUPERSHARE
          , string TRADINGPRICEPERSHARE
          , string TOTALTRADINGAMOUNT
          , string INCREASEDSHARESHUNDREDTHOUSANDS
          , string INCREASEDSHARESTENSOFTHOUSANDS
          , string INCREASEDSHARESTHOUSANDS
          , string INCREASEDSHARESIRREGULARLOTS
          , string HOLDINGSHARES
         
          )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                
                                   
                                    UPDATE [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                    SET 
                                    [CAPITALINCREASERECORDDATE]='{1}'
                                    ,[REASONFORCHANGE]='{2}'
                                    ,[STOCKACCOUNTNUMBER]='{3}'
                                    ,[STOCKNAME]='{4}'
                                    ,[INCREASEDSHARES]='{5}'
                                    ,[PARVALUPERSHARE]='{6}'
                                    ,[TRADINGPRICEPERSHARE]='{7}'
                                    ,[TOTALTRADINGAMOUNT]='{8}'
                                    ,[INCREASEDSHARESHUNDREDTHOUSANDS]='{9}'
                                    ,[INCREASEDSHARESTENSOFTHOUSANDS]='{10}'
                                    ,[INCREASEDSHARESTHOUSANDS]='{11}'
                                    ,[INCREASEDSHARESIRREGULARLOTS]='{12}'
                                    ,[HOLDINGSHARES]='{13}'
                                    WHERE [SERNO]='{0}'
                                        
                                        ",SERNO
                                        , CAPITALINCREASERECORDDATE
                                        , REASONFORCHANGE
                                        , STOCKACCOUNTNUMBER
                                        , STOCKNAME
                                        , INCREASEDSHARES
                                        , PARVALUPERSHARE
                                        , TRADINGPRICEPERSHARE
                                        , TOTALTRADINGAMOUNT
                                        , INCREASEDSHARESHUNDREDTHOUSANDS
                                        , INCREASEDSHARESTENSOFTHOUSANDS
                                        , INCREASEDSHARESTHOUSANDS
                                        , INCREASEDSHARESIRREGULARLOTS
                                        , HOLDINGSHARES
                                        
                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");

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

        public void TKSTOCKSTRANSADD_DELETE(string SERNO)
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

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


                sqlConn.Close();
                sqlConn.Open();
                tran = sqlConn.BeginTransaction();

                sbSql.Clear();

                sbSql.AppendFormat(@"                                
                                   
                                    DELETE  [TKACT].[dbo].[TKSTOCKSTRANSADD]                                  
                                    WHERE [SERNO]='{0}'
                                        
                                        ", SERNO                                      

                                        );


                cmd.Connection = sqlConn;
                cmd.CommandTimeout = 60;
                cmd.CommandText = sbSql.ToString();
                cmd.Transaction = tran;
                result = cmd.ExecuteNonQuery();

                if (result == 0)
                {
                    tran.Rollback();    //交易取消
                }
                else
                {
                    tran.Commit();      //執行交易  

                    MessageBox.Show("完成");

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
        public void CHECKADD(TextBox TEXTBOXIN)
        {
            string MESSAGES = "";

            //戶號
            if(TEXTBOXIN.Name.Equals("textBox3"))
            {
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
            }

            //戶號
            if (TEXTBOXIN.Name.Equals("textBox26"))
            {
                if (string.IsNullOrEmpty(textBox26.Text))
                {
                    MESSAGES = MESSAGES + "戶號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox26.Text))
                {
                    string input = textBox26.Text;
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
            }

            //股東姓名
            if (TEXTBOXIN.Name.Equals("textBox4"))
            {
                if (string.IsNullOrEmpty(textBox4.Text))
                {
                    MESSAGES = MESSAGES + " 股東姓名 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox4.Text))
                {
                    string input = textBox4.Text;

                    if (input.Length >= 3)
                    {
                        // 輸入為 4 位數字
                        // 在這裡處理符合條件的情況
                    }
                    else
                    {
                        MESSAGES = MESSAGES + " 股東姓名 至少3個中文字 ";
                    }

                }
            }

            //股東姓名
            if (TEXTBOXIN.Name.Equals("textBox27"))
            {
                if (string.IsNullOrEmpty(textBox27.Text))
                {
                    MESSAGES = MESSAGES + " 股東姓名 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox27.Text))
                {
                    string input = textBox27.Text;

                    if (input.Length >= 3)
                    {
                        // 輸入為 4 位數字
                        // 在這裡處理符合條件的情況
                    }
                    else
                    {
                        MESSAGES = MESSAGES + " 股東姓名 至少3個中文字 ";
                    }

                }
            }


            //身份證字號或統一編號
            if (TEXTBOXIN.Name.Equals("textBox5"))
            {
                if (string.IsNullOrEmpty(textBox5.Text))
                {
                    MESSAGES = MESSAGES + " 身份證字號或統一編號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox5.Text))
                {
                    string input = textBox5.Text;

                    if (input.Length == 8 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else if (input.Length == 10 && Regex.IsMatch(input, @"^[A-Za-z]\d{9}$"))
                    {
                        // 符合條件 2：長度為 10 位，開頭為一個英文字母，其餘 9 位為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + " 法人：8位數字、自然人1英文字+9位數字 ";
                    }

                }
            }
            //身份證字號或統一編號
            if (TEXTBOXIN.Name.Equals("textBox28"))
            {
                if (string.IsNullOrEmpty(textBox28.Text))
                {
                    MESSAGES = MESSAGES + " 身份證字號或統一編號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox28.Text))
                {
                    string input = textBox28.Text;

                    if (input.Length == 8 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else if (input.Length == 10 && Regex.IsMatch(input, @"^[A-Za-z]\d{9}$"))
                    {
                        // 符合條件 2：長度為 10 位，開頭為一個英文字母，其餘 9 位為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + " 法人：8位數字、自然人1英文字+9位數字 ";
                    }

                }
            }

            //通訊地郵遞區號
            if (TEXTBOXIN.Name.Equals("textBox6"))
            {
                if (string.IsNullOrEmpty(textBox6.Text))
                {
                    MESSAGES = MESSAGES + " 通訊地郵遞區號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox6.Text))
                {
                    string input = textBox6.Text;

                    if (input.Length == 6 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "通訊地郵遞區號 6位數字 ";
                    }

                }
            }
            //通訊地郵遞區號
            if (TEXTBOXIN.Name.Equals("textBox29"))
            {
                if (string.IsNullOrEmpty(textBox29.Text))
                {
                    MESSAGES = MESSAGES + " 通訊地郵遞區號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox29.Text))
                {
                    string input = textBox29.Text;

                    if (input.Length == 6 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "通訊地郵遞區號 6位數字 ";
                    }

                }
            }

            //通訊地址
            if (TEXTBOXIN.Name.Equals("textBox7"))
            {
                if (string.IsNullOrEmpty(textBox7.Text))
                {
                    MESSAGES = MESSAGES + " 通訊地址 不可空白";
                }
            }

            //通訊地址
            if (TEXTBOXIN.Name.Equals("textBox30"))
            {
                if (string.IsNullOrEmpty(textBox30.Text))
                {
                    MESSAGES = MESSAGES + " 通訊地址 不可空白";
                }
            }

            //戶籍地郵遞區號
            if (TEXTBOXIN.Name.Equals("textBox8"))
            {
                if (string.IsNullOrEmpty(textBox8.Text))
                {
                    MESSAGES = MESSAGES + " 通訊地郵遞區號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox8.Text))
                {
                    string input = textBox8.Text;

                    if (input.Length == 6 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "通訊地郵遞區號 6位數字 ";
                    }

                }
            }

            //戶籍地郵遞區號
            if (TEXTBOXIN.Name.Equals("textBox31"))
            {
                if (string.IsNullOrEmpty(textBox31.Text))
                {
                    MESSAGES = MESSAGES + " 通訊地郵遞區號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox31.Text))
                {
                    string input = textBox31.Text;

                    if (input.Length == 6 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "通訊地郵遞區號 6位數字 ";
                    }

                }
            }

            //戶籍地址
            if (TEXTBOXIN.Name.Equals("textBox9"))
            {
                if (string.IsNullOrEmpty(textBox9.Text))
                {
                    MESSAGES = MESSAGES + " 戶籍地址 不可空白";
                }
            }
            //戶籍地址
            if (TEXTBOXIN.Name.Equals("textBox32"))
            {
                if (string.IsNullOrEmpty(textBox32.Text))
                {
                    MESSAGES = MESSAGES + " 戶籍地址 不可空白";
                }
            }

            //出生 / 設立日期

            //銀行名稱
            if (TEXTBOXIN.Name.Equals("textBox10"))
            {
                if (string.IsNullOrEmpty(textBox10.Text))
                {
                    MESSAGES = MESSAGES + " 銀行名稱 不可空白";
                }
            }

            //銀行名稱
            if (TEXTBOXIN.Name.Equals("textBox25"))
            {
                if (string.IsNullOrEmpty(textBox25.Text))
                {
                    MESSAGES = MESSAGES + " 銀行名稱 不可空白";
                }
            }
            //分行名稱
            if (TEXTBOXIN.Name.Equals("textBox11"))
            {
                if (string.IsNullOrEmpty(textBox11.Text))
                {
                    MESSAGES = MESSAGES + " 分行名稱 不可空白";
                }
            }
            //分行名稱
            if (TEXTBOXIN.Name.Equals("textBox33"))
            {
                if (string.IsNullOrEmpty(textBox33.Text))
                {
                    MESSAGES = MESSAGES + " 分行名稱 不可空白";
                }
            }

            //銀行代碼
            if (TEXTBOXIN.Name.Equals("textBox12"))
            {
                if (string.IsNullOrEmpty(textBox12.Text))
                {
                    MESSAGES = MESSAGES + " 銀行代碼 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox12.Text))
                {
                    string input = textBox12.Text;

                    if (input.Length == 7 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "銀行代碼 7位數字 ";
                    }

                }
            }
            //銀行代碼
            if (TEXTBOXIN.Name.Equals("textBox34"))
            {
                if (string.IsNullOrEmpty(textBox34.Text))
                {
                    MESSAGES = MESSAGES + " 銀行代碼 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox34.Text))
                {
                    string input = textBox34.Text;

                    if (input.Length == 7 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "銀行代碼 7位數字 ";
                    }

                }
            }

            //帳號
            if (TEXTBOXIN.Name.Equals("textBox13"))
            {
                if (string.IsNullOrEmpty(textBox13.Text))
                {
                    MESSAGES = MESSAGES + " 帳號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox13.Text))
                {
                    string input = textBox13.Text;

                    if (input.Length >= 11 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "帳號 11~14碼數字 ";
                    }

                }
            }
            //帳號
            if (TEXTBOXIN.Name.Equals("textBox35"))
            {
                if (string.IsNullOrEmpty(textBox35.Text))
                {
                    MESSAGES = MESSAGES + " 帳號 不可空白";
                }
                if (!string.IsNullOrEmpty(textBox35.Text))
                {
                    string input = textBox35.Text;

                    if (input.Length >= 11 && Regex.IsMatch(input, @"^\d+$"))
                    {
                        // 符合條件 1：長度為 8 位且全為數字                    
                    }
                    else
                    {
                        MESSAGES = MESSAGES + "帳號 11~14碼數字 ";
                    }

                }
            }

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
            CHECKADD(textBox3);
        }
        private void textBox4_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox4);
        }
        private void textBox5_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox5);
        }
        private void textBox6_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox6);
        }

        private void textBox7_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox7);
        }

        private void textBox8_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox8);
        }

        private void textBox9_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox9);
        }

        private void textBox10_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox10);
        }

        private void textBox11_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox11);
        }

        private void textBox12_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox12);
        }

        private void textBox13_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox13);
        }
        private void textBox26_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox26);
        }

        private void textBox27_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox27);
        }

        private void textBox28_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox28);
        }

        private void textBox29_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox29);
        }

        private void textBox30_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox30);
        }

        private void textBox31_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox31);
        }

        private void textBox32_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox32);
        }

        private void textBox25_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox25);
        }

        private void textBox33_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox33);
        }

        private void textBox34_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox34);
        }

        private void textBox35_Leave(object sender, EventArgs e)
        {
            CHECKADD(textBox35);
        }
        private void textBox48_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
            {
                e.Handled = true; // 阻止輸入非數字字符
            }

            // 檢查輸入長度是否超過 7
            if (textBox1.Text.Length >= 7 && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; // 阻止輸入超過指定長度的字符
            }
        }
        private void textBox50_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && e.KeyChar != '.')
            {
                e.Handled = true; // 阻止輸入非數字、控制字符或小數點
            }

            // 檢查是否已輸入小數點
            if (e.KeyChar == '.' && textBox1.Text.Contains('.'))
            {
                e.Handled = true; // 阻止輸入多個小數點
            }

            // 檢查小數點後位數
            if (textBox1.Text.Contains('.'))
            {
                string[] parts = textBox1.Text.Split('.');
                if (parts.Length > 1 && parts[1].Length >= 5 && !char.IsControl(e.KeyChar))
                {
                    e.Handled = true; // 阻止輸入超過五位小數
                }
            }
        }
        private void textBox48_TextChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX51();
        }

        private void textBox50_TextChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX51();
        }

        public void SET_TEXTBOX51()
        {
            if(!string.IsNullOrEmpty(textBox48.Text) && !string.IsNullOrEmpty(textBox50.Text))
            {
                decimal result = Convert.ToDecimal(textBox48.Text) * Convert.ToDecimal(textBox50.Text);
                int roundedResult = (int)Math.Round(result);
                string roundedResultString = roundedResult.ToString();
                textBox51.Text = roundedResultString;
            }
        }
        #endregion


        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox1.Text,textBox2.Text);
        }
        private void button2_Click(object sender, EventArgs e)
        {
          

            TKSTOCKS_ADD(DateTime.Now.ToString("yyyy/MM/dd")
                , textBox3.Text
                , textBox4.Text
                , textBox5.Text
                , textBox6.Text
                , textBox7.Text
                , textBox8.Text
                , textBox9.Text
                , dateTimePicker1.Value.ToString("yyyy/MM/dd") 
                , textBox10.Text
                , textBox11.Text
                , textBox12.Text
                , textBox13.Text
                , textBox15.Text
                , textBox14.Text
                , textBox16.Text
                , textBox17.Text
                , textBox18.Text
                , textBox19.Text
                , textBox20.Text
                , textBox21.Text
                , textBox70.Text
                );

            Search(textBox1.Text, textBox2.Text);
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Search_DG2(textBox22.Text, textBox23.Text);
            Search_DG3(textBox22.Text, textBox23.Text);
        }


        private void button4_Click(object sender, EventArgs e)
        {


            TKSTOCKSCHAGES_ADD(DateTime.Now.ToString("yyyy/MM/dd")
                , textBox26.Text
                , textBox27.Text
                , textBox28.Text
                , textBox29.Text
                , textBox30.Text
                , textBox31.Text
                , textBox32.Text
                , dateTimePicker2.Value.ToString("yyyy/MM/dd")
                , textBox25.Text
                , textBox33.Text
                , textBox34.Text
                , textBox35.Text
                , textBox24.Text
                , textBox36.Text
                , textBox37.Text
                , textBox38.Text
                , textBox39.Text
                , textBox40.Text
                , textBox41.Text
                , textBox42.Text
                , "N"
                , textBox43.Text
                , textBox71.Text
                );

            UPDATE_TO_TKSTOCKS(
                               textBox43.Text
                               , textBox26.Text
                               , textBox27.Text
                               , textBox28.Text
                               , textBox29.Text
                               , textBox30.Text
                               , textBox31.Text
                               , textBox32.Text
                               , dateTimePicker2.Value.ToString("yyyy/MM/dd")
                               , textBox25.Text
                               , textBox33.Text
                               , textBox34.Text
                               , textBox35.Text
                               , textBox24.Text
                               , textBox36.Text
                               , textBox37.Text
                               , textBox38.Text
                               , textBox39.Text
                               , textBox40.Text
                               , textBox41.Text
                               , textBox42.Text
                               , textBox71.Text
                               );

            Search_DG2(textBox22.Text, textBox23.Text);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                
            }
        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            Search_DG4(textBox44.Text, textBox45.Text);
        }

        private void button6_Click(object sender, EventArgs e)
        {
            TKSTOCKSTRANSADD_ADD(
             dateTimePicker3.Value.ToString("yyyy/MM/dd")
            , comboBox1.SelectedValue.ToString()
            , textBox46.Text
            , textBox47.Text
            , textBox48.Text
            , textBox49.Text
            , textBox50.Text
            , textBox51.Text
            , textBox52.Text
            , textBox53.Text
            , textBox54.Text
            , textBox55.Text
            , textBox56.Text
            , textBox57.Text
            );

            Search_DG4(textBox44.Text, textBox45.Text);
        }


        private void button7_Click(object sender, EventArgs e)
        {
            TKSTOCKSTRANSADD_UPDATE(
            textBox58.Text
         , dateTimePicker3.Value.ToString("yyyy/MM/dd")
         , comboBox2.SelectedValue.ToString()
         , textBox59.Text
         , textBox60.Text
         , textBox61.Text
         , textBox62.Text
         , textBox63.Text
         , textBox64.Text
         , textBox65.Text
         , textBox66.Text
         , textBox67.Text
         , textBox68.Text
         , textBox69.Text
        
         );

            Search_DG4(textBox44.Text, textBox45.Text);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                TKSTOCKSTRANSADD_DELETE(textBox58.Text);
                Search_DG4(textBox44.Text, textBox45.Text);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            Search_DG5(textBox72.Text, textBox73.Text);
        }


        #endregion


    }
}
