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
        SqlConnection sqlConn = new SqlConnection();
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

        private bool isTextBox76Changing = false;
        private bool isTextBox77Changing = false;


        public FrmSTOCKRECORDS()
        {
            InitializeComponent();

            comboBox1load();
            comboBox2load();
            comboBox3load();

            comboBox5load();

        }

        #region FUNCTION
        public void comboBox1load()
        {
            LoadComboBoxData(comboBox1, "SELECT  [ID],[KINDS],[NAMES],[KEYS] FROM [TKACT].[dbo].[TBPARAS] WHERE KINDS='異動原因' ORDER BY ID", "KEYS", "KEYS");
        }
        public void comboBox2load()
        {
            //LoadComboBoxData(comboBox2, "SELECT  [ID],[KINDS],[NAMES],[KEYS] FROM [TKACT].[dbo].[TBPARAS] WHERE KINDS='異動原因' ORDER BY ID", "KEYS", "KEYS");
        }
        public void comboBox3load()
        {
            LoadComboBoxData(comboBox3, "SELECT [ID],[KINDS],[NAMES],[KEYS] FROM [TKACT].[dbo].[TBPARAS] WHERE [KINDS]='異動原因轉讓' ORDER BY [ID]", "KEYS", "KEYS");
        }

        public void comboBox5load()
        {
            LoadComboBoxData(comboBox5, "SELECT [ID],[KINDS],[NAMES],[KEYS] FROM [TKACT].[dbo].[TBPARAS] WHERE [KINDS]='REPORTS' ORDER BY [ID]", "KEYS", "KEYS");
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
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE N'%{0}%'", STOCKNAME);
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
                                ,[REGISTEREDADDRESS] AS '戶籍/設立地址'
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
                                FROM [TKACT].[dbo].[TKSTOCKSNAMES]
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
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE N'%{0}%'", STOCKNAME);
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
                                ,[REGISTEREDADDRESS] AS '戶籍/設立地址'
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
                                FROM [TKACT].[dbo].[TKSTOCKSNAMES]
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
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE N'%{0}%'", STOCKNAME);
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
                                ,[REGISTEREDADDRESS] AS '戶籍/設立地址'
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
                                ,[STOCKSHARES] AS '增資股票(不定額股)股數'
                                ,[HOLDINGSHARES] AS '持有股數'
                                ,[SERNO]
                                ,[CAPITALINCREASERECORDDATE]
                                ,[ID]

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
                sbSqlQuery1.AppendFormat(@" AND (STOCKACCOUNTNUMBERFORM LIKE '%{0}%' OR STOCKACCOUNTNUMBERTO LIKE '%{0}%' )", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                sbSqlQuery2.AppendFormat(@" AND ([STOCKNAMEFORM] LIKE '%{0}%' OR [STOCKNAMETO] LIKE '%{0}%')", STOCKNAME);
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


        public void Search_DG6(string STOCKACCOUNTNUMBER, string STOCKNAME)
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
                sbSqlQuery2.AppendFormat(@" AND STOCKNAME LIKE N'%{0}%'", STOCKNAME);
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
                                ,'民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([EXDIVIDENDINTERESTRECORDDATE],1,4))-1911))+'年'+SUBSTRING([EXDIVIDENDINTERESTRECORDDATE],6,2)+'月'+SUBSTRING([EXDIVIDENDINTERESTRECORDDATE],9,2) +'日' AS '除權/息基準日'
                                ,'民國'+CONVERT(NVARCHAR,(CONVERT(INT,SUBSTRING([CASHDIVIDENDPAYMENTDATE],1,4))-1911))+'年'+SUBSTRING([CASHDIVIDENDPAYMENTDATE],6,2)+'月'+SUBSTRING([CASHDIVIDENDPAYMENTDATE],9,2) +'日' AS '現金股利發放日'

                                ,[CASHDIVIDENDPERSHARE] AS '每股配發現金股利'
                                ,[STOCKDIVIDEND] AS '每股配發股票股利'
                                ,[DIVAMOUNTS] AS '每股配發資本公積'
                                ,[DECLAREDCASHDIVIDEND] AS '應發股利現金股利'
                                ,[DECLAREDSTOCKDIVIDEND] AS '應發股利股票股利'
                                ,[SUPPLEMENTARYPREMIUMTOBEDEDUCTED] AS '應扣補充保費'
                                ,[ACTUALCASHDIVIDENDPAID] AS '實發現金股利'
                                ,[CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES] AS '盈餘增資配股數'
                                ,[CAPITALIZATIONOFCAPITALSURPLUSSHARES] AS '資本公積增資配股數'
                                ,[DIVYEARS] AS '分配年度'
                               
                                ,[EXDIVIDENDINTERESTRECORDDATE]
                                ,[CASHDIVIDENDPAYMENTDATE]

                                FROM [TKACT].[dbo].[TKSTOCKSDIV]                             
                                WHERE 1=1
                                {0}
                                {1}
                                ORDER BY  [SERNO]

                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString());

            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView6, SortedColumn, SortedModel);
        }

        public void Search_DG7(string STOCKID, string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();
            sbSqlQuery3.Clear();

            if (!string.IsNullOrEmpty(STOCKID))
            {
                sbSqlQuery1.AppendFormat(@" AND STOCKID LIKE '%{0}%'", STOCKID);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                sbSqlQuery2.AppendFormat(@" AND STOCKACCOUNTNUMBER LIKE '%{0}%'", STOCKACCOUNTNUMBER);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                sbSqlQuery3.AppendFormat(@" AND STOCKNAME LIKE N'%{0}%'", STOCKNAME);
            }
            else
            {
                sbSqlQuery3.AppendFormat(@" ");
            }



            sbSql.AppendFormat(@"
                                SELECT
                                [STOCKID] AS '股票號碼'
                                ,[PARVALUPER] AS '每股面額'
                                ,[STOCKSHARES] AS '股數'
                                ,[STOCKACCOUNTNUMBER] AS '戶號'
                                ,[STOCKNAME] AS '股東姓名'
                                ,[STOCKIDKEY]
                                FROM [TKACT].[dbo].[TKSTOCKSREORDS]
                                WHERE 1=1
                                {0}
                                {1}
                                {2}
                                ORDER BY [STOCKACCOUNTNUMBER],[STOCKID]

                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString(), sbSqlQuery3.ToString());

            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView7, SortedColumn, SortedModel);
        }

        private void dataGridView7_SelectionChanged(object sender, EventArgs e)
        {
            textBox94.Text = "";
            textBox95.Text = "";
            textBox96.Text = "";
            textBox97.Text = "";
            textBox98.Text = "";
            textBox99.Text = "";
            textBox102.Text = "0";
            textBox103.Text = "";
            textBox104.Text = "0";

            if (dataGridView7.CurrentRow != null)
            {
                int rowindex = dataGridView7.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView7.Rows[rowindex];
                    textBox94.Text = row.Cells["股票號碼"].Value.ToString();
                    textBox95.Text = row.Cells["戶號"].Value.ToString();
                    textBox96.Text = row.Cells["股東姓名"].Value.ToString();
                    textBox97.Text = row.Cells["每股面額"].Value.ToString();
                    textBox98.Text = row.Cells["股數"].Value.ToString();
                    textBox99.Text = row.Cells["STOCKIDKEY"].Value.ToString();
                    textBox102.Text = row.Cells["每股面額"].Value.ToString();
                    Search_DG8(textBox94.Text.Trim(),"N");
                }
                else
                {

                }

            }
        }

        public void Search_DG8(string OLDSTOCKID,string VALIDS)
        {
            StringBuilder sbSqlQuery1 = new StringBuilder();
            StringBuilder sbSqlQuery2 = new StringBuilder();
            StringBuilder sbSqlQuery3 = new StringBuilder();

            sbSql.Clear();
            sbSqlQuery1.Clear();
            sbSqlQuery2.Clear();
            sbSqlQuery3.Clear();

            if (!string.IsNullOrEmpty(OLDSTOCKID))
            {
                sbSqlQuery1.AppendFormat(@" AND OLDSTOCKID LIKE '%{0}%'", OLDSTOCKID);
            }
            else
            {
                sbSqlQuery1.AppendFormat(@" ");
            }

            if (!string.IsNullOrEmpty(VALIDS))
            {
                sbSqlQuery2.AppendFormat(@" AND VALIDS LIKE '%{0}%'", VALIDS);
            }
            else
            {
                sbSqlQuery2.AppendFormat(@" ");
            }


            sbSql.AppendFormat(@"
                                SELECT 
                                [NEWSTOCKID] AS '分割後的股票號碼'
                                ,[NEWPARVALUPER] AS '分割後的每股面額'
                                ,[NEWSTOCKSHARES] AS '分割後的股數'
                                ,[OLDSTOCKID] AS '分割前的股票號碼'
                                ,[OLDPARVALUPER] AS '分割前的每股面額'
                                ,[OLDSTOCKSHARES] AS '分割前的股數'
                                ,[STOCKACCOUNTNUMBER] AS '戶號'
                                ,[STOCKNAME] AS '股東姓名'
                                ,[VALIDS]
                                ,[STOCKIDKEY] 
                                FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                                WHERE 1=1
                                {0}
                                {1}
                               

                                  ", sbSqlQuery1.ToString(), sbSqlQuery2.ToString());

            sbSql.AppendFormat(@"  ");

            SEARCH(sbSql.ToString(), dataGridView8, SortedColumn, SortedModel);
        }

        private void dataGridView8_SelectionChanged(object sender, EventArgs e)
        {
            textBox103.Text = "";
            textBox104.Text = "0";

            if (dataGridView8.CurrentRow != null)
            {
                int rowindex = dataGridView8.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView8.Rows[rowindex];
                    textBox103.Text = row.Cells["分割後的股票號碼"].Value.ToString();

                    //計算暫時分割的股數
                    DataTable DT= FINE_TKSTOCKSREORDSDIV(textBox94.Text.Trim(),"N");

                    textBox104.Text = DT.Rows[0]["NEWSTOCKSHARES"].ToString();



                }
                else
                {
                  
                }

            }
          
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
     

        public void TKSTOCKSNAMES_ADD(
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
                                    INSERT INTO [TKACT].[dbo].[TKSTOCKSNAMES]
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
                                    ,N'{2}'
                                    ,N'{3}'
                                    ,N'{4}'
                                    ,N'{5}'
                                    ,N'{6}'
                                    ,N'{7}'
                                    ,'{8}'
                                    ,N'{9}'
                                    ,N'{10}'
                                    ,N'{11}'
                                    ,N'{12}'
                                    ,N'{13}'
                                    ,N'{14}'
                                    ,N'{15}'
                                    ,N'{16}'
                                    ,N'{17}'
                                    ,N'{18}'
                                    ,N'{19}'
                                    ,N'{20}'
                                    ,N'{21}'
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

        public void TKSTOCKSNAMES_DELETE(string ID)
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
                                   
                                    DELETE  [TKACT].[dbo].[TKSTOCKSNAMES]                                   
                                    WHERE [ID]='{0}'
                                    ", ID

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
                                    ,N'{1}'
                                    ,N'{2}'
                                    ,N'{3}'
                                    ,N'{4}'
                                    ,N'{5}'
                                    ,N'{6}'
                                    ,N'{7}'
                                    ,'{8}'
                                    ,N'{9}'
                                    ,N'{10}'
                                    ,N'{11}'
                                    ,N'{12}'
                                    ,N'{13}'
                                    ,N'{14}'
                                    ,N'{15}'
                                    ,N'{16}'
                                    ,N'{17}'
                                    ,N'{18}'
                                    ,N'{19}'
                                    ,N'{20}'
                                    ,N'{21}'
                                    ,N'{22}'
                                    ,N'{23}'
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

        public void TKSTOCKSNAMES_UPDATE(
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
                                    UPDATE [TKACT].[dbo].[TKSTOCKSNAMES]
                                    SET 
                                    [STOCKACCOUNTNUMBER]=@STOCKACCOUNTNUMBER
                                    ,[STOCKNAME]=@STOCKNAME
                                    ,[IDNUMBER]=@IDNUMBER
                                    ,[POSTALCODE]=@POSTALCODE
                                    ,[MAILINGADDRESS]=@MAILINGADDRESS
                                    ,[REGISTEREDPOSTALCODE]=@REGISTEREDPOSTALCODE
                                    ,[REGISTEREDADDRESS]=@REGISTEREDADDRESS
                                    ,[DATEOFBIRTH]=@DATEOFBIRTH
                                    ,[BANKNAME]=@BANKNAME
                                    ,[BRANCHNAME]=@BRANCHNAME
                                    ,[BANKCODE]=@BANKCODE
                                    ,[ACCOUNTNUMBER]=@ACCOUNTNUMBER
                                    ,[HOMEPHONENUMBER]=@HOMEPHONENUMBER
                                    ,[MOBILEPHONENUMBER]=@MOBILEPHONENUMBER
                                    ,[EMAIL]=@EMAIL
                                    ,[PASSPORTNUMBER]=@PASSPORTNUMBER
                                    ,[ENGLISHNAME]=@ENGLISHNAME
                                    ,[FATHER]=@FATHER
                                    ,[MOTHER]=@MOTHER
                                    ,[SPOUSE]=@SPOUSE
                                    ,[COMMENTS]=@COMMENTS
                                    WHERE [ID]=@ID
                                        
                                        ");

                using (SqlConnection connection = new SqlConnection(sqlsb.ConnectionString))
                {
                    connection.Open();

                    using (SqlCommand command = new SqlCommand(sbSql.ToString(), connection))
                    {
                        command.Parameters.AddWithValue("@ID", ID);
                        command.Parameters.AddWithValue("@STOCKACCOUNTNUMBER", STOCKACCOUNTNUMBER);
                        command.Parameters.AddWithValue("@STOCKNAME", STOCKNAME);
                        command.Parameters.AddWithValue("@IDNUMBER", IDNUMBER);
                        command.Parameters.AddWithValue("@POSTALCODE", POSTALCODE);
                        command.Parameters.AddWithValue("@MAILINGADDRESS", MAILINGADDRESS);
                        command.Parameters.AddWithValue("@REGISTEREDPOSTALCODE", REGISTEREDPOSTALCODE);
                        command.Parameters.AddWithValue("@REGISTEREDADDRESS", REGISTEREDADDRESS);
                        command.Parameters.AddWithValue("@DATEOFBIRTH", DATEOFBIRTH);
                        command.Parameters.AddWithValue("@BANKNAME", BANKNAME);
                        command.Parameters.AddWithValue("@BRANCHNAME", BRANCHNAME);
                        command.Parameters.AddWithValue("@BANKCODE", BANKCODE);
                        command.Parameters.AddWithValue("@ACCOUNTNUMBER", ACCOUNTNUMBER);
                        command.Parameters.AddWithValue("@HOMEPHONENUMBER", HOMEPHONENUMBER);
                        command.Parameters.AddWithValue("@MOBILEPHONENUMBER", MOBILEPHONENUMBER);
                        command.Parameters.AddWithValue("@EMAIL", EMAIL);
                        command.Parameters.AddWithValue("@PASSPORTNUMBER", PASSPORTNUMBER);
                        command.Parameters.AddWithValue("@ENGLISHNAME", ENGLISHNAME);
                        command.Parameters.AddWithValue("@FATHER", FATHER);
                        command.Parameters.AddWithValue("@MOTHER", MOTHER);
                        command.Parameters.AddWithValue("@SPOUSE", SPOUSE);
                        command.Parameters.AddWithValue("@COMMENTS", COMMENTS);

                        command.ExecuteNonQuery();
                        MessageBox.Show("完成");
                    }


                }

                //sbSql.AppendFormat(@"                                    
                //                    UPDATE [TKACT].[dbo].[TKSTOCKSNAMES]
                //                    SET 
                //                    [STOCKACCOUNTNUMBER]=N'{1}'
                //                    ,[STOCKNAME]=N'{2}'
                //                    ,[IDNUMBER]=N'{3}'
                //                    ,[POSTALCODE]=N'{4}'
                //                    ,[MAILINGADDRESS]=N'{5}'
                //                    ,[REGISTEREDPOSTALCODE]=N'{6}'
                //                    ,[REGISTEREDADDRESS]=N'{7}'
                //                    ,[DATEOFBIRTH]='{8}'
                //                    ,[BANKNAME]=N'{9}'
                //                    ,[BRANCHNAME]=N'{10}'
                //                    ,[BANKCODE]=N'{11}'
                //                    ,[ACCOUNTNUMBER]=N'{12}'
                //                    ,[HOMEPHONENUMBER]=N'{13}'
                //                    ,[MOBILEPHONENUMBER]=N'{14}'
                //                    ,[EMAIL]=N'{15}'
                //                    ,[PASSPORTNUMBER]=N'{16}'
                //                    ,[ENGLISHNAME]=N'{17}'
                //                    ,[FATHER]=N'{18}'
                //                    ,[MOTHER]=N'{19}'
                //                    ,[SPOUSE]=N'{20}'
                //                    ,[COMMENTS]=N'{21}'
                //                    WHERE [ID]='{0}'

                //                        ", ID
                //                        , STOCKACCOUNTNUMBER
                //                        , STOCKNAME
                //                        , IDNUMBER
                //                        , POSTALCODE
                //                        , MAILINGADDRESS
                //                        , REGISTEREDPOSTALCODE
                //                        , REGISTEREDADDRESS
                //                        , DATEOFBIRTH
                //                        , BANKNAME
                //                        , BRANCHNAME
                //                        , BANKCODE
                //                        , ACCOUNTNUMBER
                //                        , HOMEPHONENUMBER
                //                        , MOBILEPHONENUMBER
                //                        , EMAIL
                //                        , PASSPORTNUMBER
                //                        , ENGLISHNAME
                //                        , FATHER
                //                        , MOTHER
                //                        , SPOUSE
                //                        , COMMENTS
                //                        );

                //cmd.Connection = sqlConn;
                //cmd.CommandTimeout = 60;
                //cmd.CommandText = sbSql.ToString();
                //cmd.Transaction = tran;
                //result = cmd.ExecuteNonQuery();

                //if (result == 0)
                //{
                //    tran.Rollback();    //交易取消
                //}
                //else
                //{
                //    tran.Commit();      //執行交易  

                //    MessageBox.Show("完成");

                //}

            }
            catch
            {

            }

            finally
            {
                sqlConn.Close();
            }
        }
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            string ID = "";
            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
            textBox7.Text = "";
            textBox8.Text = "";
            textBox9.Text = "";
            textBox10.Text = "";
            textBox11.Text = "";
            textBox12.Text = "";
            textBox13.Text = "";
            textBox14.Text = "";
            textBox15.Text = "";
            textBox16.Text = "";
            textBox17.Text = "";
            textBox18.Text = "";
            textBox19.Text = "";
            textBox20.Text = "";
            textBox21.Text = "";
            textBox70.Text = "";
            textBox130.Text = "";

            if (dataGridView1.CurrentRow != null)
            {
                int rowindex = dataGridView1.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView1.Rows[rowindex];
                    ID = row.Cells["ID"].Value.ToString();
                    textBox3.Text = row.Cells["戶號"].Value.ToString();
                    textBox4.Text = row.Cells["股東姓名"].Value.ToString();
                    textBox5.Text = row.Cells["身份證字號或統一編號"].Value.ToString();
                    textBox6.Text = row.Cells["通訊地郵遞區號"].Value.ToString();
                    textBox7.Text = row.Cells["通訊地址"].Value.ToString();
                    textBox8.Text = row.Cells["戶籍地郵遞區號"].Value.ToString();
                    textBox9.Text = row.Cells["戶籍/設立地址"].Value.ToString();
                    textBox10.Text = row.Cells["銀行名稱"].Value.ToString();
                    textBox11.Text = row.Cells["分行名稱"].Value.ToString();
                    textBox12.Text = row.Cells["銀行代碼"].Value.ToString();
                    textBox13.Text = row.Cells["帳號"].Value.ToString();
                    textBox15.Text = row.Cells["住家電話"].Value.ToString();
                    textBox14.Text = row.Cells["手機號碼"].Value.ToString();
                    textBox16.Text = row.Cells["e-mail"].Value.ToString();
                    textBox17.Text = row.Cells["護照號碼"].Value.ToString();
                    textBox18.Text = row.Cells["英文名"].Value.ToString();
                    textBox19.Text = row.Cells["父"].Value.ToString();
                    textBox20.Text = row.Cells["母"].Value.ToString();
                    textBox21.Text = row.Cells["配偶"].Value.ToString();
                    textBox70.Text = row.Cells["備註"].Value.ToString();
                    textBox130.Text = row.Cells["ID"].Value.ToString();

                    dateTimePicker1.Value = Convert.ToDateTime(row.Cells["DATEOFBIRTH"].Value.ToString());

                }
                else
                {


                }
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
                    textBox32.Text = row.Cells["戶籍/設立地址"].Value.ToString();
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
            textBox58.Text = "";

            if (dataGridView4.CurrentRow != null)
            {
                int rowindex = dataGridView4.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView4.Rows[rowindex];
                    SERNO = row.Cells["SERNO"].Value.ToString();
                    textBox58.Text = row.Cells["SERNO"].Value.ToString();
                    textBox57.Text = row.Cells["ID"].Value.ToString();
                    textBox46.Text = row.Cells["戶號"].Value.ToString();
                    textBox47.Text = row.Cells["股東姓名"].Value.ToString();
                    textBox48.Text = row.Cells["增資股數"].Value.ToString();
                    textBox49.Text = row.Cells["每股面額"].Value.ToString();
                    textBox50.Text = row.Cells["每股成交價格"].Value.ToString();
                    textBox51.Text = row.Cells["成交總額"].Value.ToString();
                    textBox52.Text = row.Cells["增資股票號碼(十萬股)"].Value.ToString();
                    textBox53.Text = row.Cells["增資股票號碼(萬股)"].Value.ToString();
                    textBox54.Text = row.Cells["增資股票號碼(千股)"].Value.ToString();
                    textBox55.Text = row.Cells["增資股票號碼(不定額股)"].Value.ToString();
                    textBox56.Text = row.Cells["增資股票(不定額股)股數"].Value.ToString();
                    dateTimePicker3.Value= Convert.ToDateTime(row.Cells["CAPITALINCREASERECORDDATE"].Value.ToString());

                    comboBox1.SelectedValue = row.Cells["異動原因"].Value.ToString();


                }
            }
        }

        private void dataGridView5_SelectionChanged(object sender, EventArgs e)
        {
            string SERNO = "";
            textBox76.Text = "";
            textBox77.Text = "";
            textBox78.Text = "";
            textBox79.Text = "";
            textBox80.Text = "";
            textBox81.Text = "";
            textBox82.Text = "";
            textBox83.Text = "";
            textBox84.Text = "";
            textBox85.Text = "";
            textBox86.Text = "";
            textBox87.Text = "";
            textBox88.Text = "";
            
       
            textBox74.Text = "";
            textBox75.Text = "";
            textBox106.Text = "";

            if (dataGridView5.CurrentRow != null)
            {
                int rowindex = dataGridView5.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView5.Rows[rowindex];
                    SERNO = row.Cells["流水號"].Value.ToString();
                    textBox106.Text = row.Cells["流水號"].Value.ToString();
                    textBox76.Text = row.Cells["轉讓人戶號"].Value.ToString();
                    textBox77.Text = row.Cells["轉讓人姓名"].Value.ToString();
                    textBox78.Text = row.Cells["受讓人戶號"].Value.ToString();
                    textBox79.Text = row.Cells["受讓人姓名"].Value.ToString();
                    textBox80.Text = row.Cells["轉讓股數"].Value.ToString();
                    textBox81.Text = row.Cells["每股面額"].Value.ToString();
                    textBox82.Text = row.Cells["每股成交價格"].Value.ToString();
                    textBox83.Text = row.Cells["成交總額"].Value.ToString();
                    textBox84.Text = row.Cells["證券交易稅額"].Value.ToString();
                    textBox85.Text = row.Cells["轉讓股票號碼(十萬股)"].Value.ToString();
                    textBox86.Text = row.Cells["轉讓股票號碼(萬股)"].Value.ToString();
                    textBox87.Text = row.Cells["轉讓股票號碼(千股)"].Value.ToString();
                    textBox88.Text = row.Cells["轉讓股票號碼(不定額股)"].Value.ToString();
                  
                    textBox74.Text = row.Cells["IDFORM"].Value.ToString();
                    textBox75.Text = row.Cells["IDTO"].Value.ToString();

                    dateTimePicker5.Value = Convert.ToDateTime(row.Cells["DATEOFCHANGE"].Value.ToString());

                    comboBox3.SelectedValue = row.Cells["異動原因"].Value.ToString();

                }
            }
        }

        private void dataGridView6_SelectionChanged(object sender, EventArgs e)
        {
            string SERNO = "";

            textBox109.Text = "";
            textBox110.Text = "";
            textBox111.Text = "";
            textBox112.Text = "";
            textBox113.Text = "";
            textBox114.Text = "";
            textBox115.Text = "";
            textBox116.Text = "";
            textBox117.Text = "";
            textBox118.Text = "";
            textBox129.Text = "";
            textBox120.Text = "";
            textBox119.Text = "";

            if (dataGridView6.CurrentRow != null)
            {
                int rowindex = dataGridView6.CurrentRow.Index;
                if (rowindex >= 0)
                {
                    DataGridViewRow row = dataGridView6.Rows[rowindex];
                    SERNO = row.Cells["流水號"].Value.ToString();
                    textBox109.Text = row.Cells["戶號"].Value.ToString();
                    textBox110.Text = row.Cells["股東姓名"].Value.ToString();
                    textBox111.Text = row.Cells["每股配發現金股利"].Value.ToString();
                    textBox112.Text = row.Cells["每股配發股票股利"].Value.ToString();
                    textBox113.Text = row.Cells["應發股利現金股利"].Value.ToString();
                    textBox114.Text = row.Cells["應發股利股票股利"].Value.ToString();
                    textBox115.Text = row.Cells["應扣補充保費"].Value.ToString();
                    textBox116.Text = row.Cells["實發現金股利"].Value.ToString();
                    textBox117.Text = row.Cells["盈餘增資配股數"].Value.ToString();
                    textBox118.Text = row.Cells["資本公積增資配股數"].Value.ToString();
                    textBox129.Text = row.Cells["流水號"].Value.ToString();
                    textBox120.Text = row.Cells["分配年度"].Value.ToString();
                    textBox119.Text = row.Cells["每股配發資本公積"].Value.ToString();
                    dateTimePicker7.Value = Convert.ToDateTime(row.Cells["EXDIVIDENDINTERESTRECORDDATE"].Value.ToString());
                    dateTimePicker8.Value = Convert.ToDateTime(row.Cells["CASHDIVIDENDPAYMENTDATE"].Value.ToString());



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
            , string INCREASEDSHARESHUNDREDTHOUSANDS_ST
            , string INCREASEDSHARESTENSOFTHOUSANDS_ST
            , string INCREASEDSHARESTHOUSANDS_ST
            , string INCREASEDSHARESIRREGULARLOTS_ST
            , string HOLDINGSHARES
            , string INCREASEDSHARESHUNDREDTHOUSANDS_END
            , string INCREASEDSHARESTENSOFTHOUSANDS_END
            , string INCREASEDSHARESTHOUSANDS_END
            , string INCREASEDSHARESIRREGULARLOTS_END
            , string ID
            , string STOCKSHARES
            )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();

            try
            {
                sbSql.Clear();

                int INCREASEDSHARESHUNDREDTHOUSANDS_COUNT = 0;
                int INCREASEDSHARESTENSOFTHOUSANDS_COUNT = 0;
                int INCREASEDSHARESTHOUSANDS_COUNT = 0;
                int INCREASEDSHARESIRREGULARLOTS_COUNT = 0;

                //INCREASEDSHARESHUNDREDTHOUSANDS_COUNT
                if (INCREASEDSHARESHUNDREDTHOUSANDS_ST .Length>=7&& INCREASEDSHARESHUNDREDTHOUSANDS_END .Length>=7&& !string.IsNullOrEmpty(INCREASEDSHARESHUNDREDTHOUSANDS_ST)&&!string.IsNullOrEmpty(INCREASEDSHARESHUNDREDTHOUSANDS_END))
                {
                    int START = 0;
                    int END = 0;

                    START = Convert.ToInt32(INCREASEDSHARESHUNDREDTHOUSANDS_ST.ToString().Substring(INCREASEDSHARESHUNDREDTHOUSANDS_ST.Length-7,7));
                    END = Convert.ToInt32(INCREASEDSHARESHUNDREDTHOUSANDS_END.ToString().Substring(INCREASEDSHARESHUNDREDTHOUSANDS_END.Length - 7, 7));

                    INCREASEDSHARESHUNDREDTHOUSANDS_COUNT =  END- START+1;

                }
                else if (!string.IsNullOrEmpty(INCREASEDSHARESHUNDREDTHOUSANDS_ST))
                {
                    INCREASEDSHARESHUNDREDTHOUSANDS_COUNT = 1;
                }
                else
                {
                    INCREASEDSHARESHUNDREDTHOUSANDS_COUNT = 0;
                }
                //INCREASEDSHARESTENSOFTHOUSANDS_COUNT
                if (INCREASEDSHARESTENSOFTHOUSANDS_ST.Length >= 7 && INCREASEDSHARESTENSOFTHOUSANDS_END.Length >= 7 && !string.IsNullOrEmpty(INCREASEDSHARESTENSOFTHOUSANDS_ST) && !string.IsNullOrEmpty(INCREASEDSHARESTENSOFTHOUSANDS_END))
                {
                    int START = 0;
                    int END = 0;

                    START = Convert.ToInt32(INCREASEDSHARESTENSOFTHOUSANDS_ST.ToString().Substring(INCREASEDSHARESTENSOFTHOUSANDS_ST.Length - 7, 7));
                    END = Convert.ToInt32(INCREASEDSHARESTENSOFTHOUSANDS_END.ToString().Substring(INCREASEDSHARESTENSOFTHOUSANDS_END.Length - 7, 7));

                    INCREASEDSHARESTENSOFTHOUSANDS_COUNT = END - START+1;

                }
                else if (!string.IsNullOrEmpty(INCREASEDSHARESTENSOFTHOUSANDS_ST))
                {
                    INCREASEDSHARESTENSOFTHOUSANDS_COUNT = 1;
                }
                else
                {
                    INCREASEDSHARESTENSOFTHOUSANDS_COUNT = 0;
                }
                //INCREASEDSHARESTHOUSANDS_COUNT
                if (INCREASEDSHARESTHOUSANDS_ST.Length >= 7 && INCREASEDSHARESTHOUSANDS_END.Length >= 7 && !string.IsNullOrEmpty(INCREASEDSHARESTHOUSANDS_ST) && !string.IsNullOrEmpty(INCREASEDSHARESTHOUSANDS_END))
                {
                    int START = 0;
                    int END = 0;

                    START = Convert.ToInt32(INCREASEDSHARESTHOUSANDS_ST.ToString().Substring(INCREASEDSHARESTHOUSANDS_ST.Length - 7, 7));
                    END = Convert.ToInt32(INCREASEDSHARESTHOUSANDS_END.ToString().Substring(INCREASEDSHARESTHOUSANDS_END.Length - 7, 7));

                    INCREASEDSHARESTHOUSANDS_COUNT = END - START+1;

                }
                else if (!string.IsNullOrEmpty(INCREASEDSHARESTHOUSANDS_ST))
                {
                    INCREASEDSHARESTHOUSANDS_COUNT = 1;
                }
                else
                {
                    INCREASEDSHARESTHOUSANDS_COUNT = 0;
                }
                //INCREASEDSHARESIRREGULARLOTS_COUNT
                if (INCREASEDSHARESIRREGULARLOTS_ST.Length >= 7 && INCREASEDSHARESIRREGULARLOTS_END.Length >= 7 && !string.IsNullOrEmpty(INCREASEDSHARESIRREGULARLOTS_ST) && !string.IsNullOrEmpty(INCREASEDSHARESIRREGULARLOTS_END))
                {
                    int START = 0;
                    int END = 0;

                    START = Convert.ToInt32(INCREASEDSHARESIRREGULARLOTS_ST.ToString().Substring(INCREASEDSHARESIRREGULARLOTS_ST.Length - 7, 7));
                    END = Convert.ToInt32(INCREASEDSHARESIRREGULARLOTS_END.ToString().Substring(INCREASEDSHARESIRREGULARLOTS_END.Length - 7, 7));

                    INCREASEDSHARESIRREGULARLOTS_COUNT = END - START+1;

                }
                else if (!string.IsNullOrEmpty(INCREASEDSHARESIRREGULARLOTS_ST))
                {
                    INCREASEDSHARESIRREGULARLOTS_COUNT = 1;
                }
                else
                {
                    INCREASEDSHARESIRREGULARLOTS_COUNT = 0;
                }

                ///增資股票(不定額股)股數
                int number;
                if (int.TryParse(STOCKSHARES, out number))
                {
                    STOCKSHARES = STOCKSHARES;
                }
                else
                {
                    STOCKSHARES = "1";
                }

                //增資股票號碼(十萬股) 
                //INCREASEDSHARESHUNDREDTHOUSANDS_COUNT
                //SQL
                if (INCREASEDSHARESHUNDREDTHOUSANDS_COUNT>=2)
                {
                    //sbSql.Clear();

                    string INCREASEDSHARE = "";
                    string INCREASEDSHARE_PRE = INCREASEDSHARESHUNDREDTHOUSANDS_ST.Substring(0, INCREASEDSHARESHUNDREDTHOUSANDS_ST.Length-7);
                    int INCREASEDSHARE_COUT= Convert.ToInt32(INCREASEDSHARESHUNDREDTHOUSANDS_ST.Substring(INCREASEDSHARESHUNDREDTHOUSANDS_ST.Length - 7,7));
                    for (int i=1;i<= INCREASEDSHARESHUNDREDTHOUSANDS_COUNT;i++)
                    {
                        INCREASEDSHARE = INCREASEDSHARE_PRE + PadNumberWithZero7(INCREASEDSHARE_COUT);

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
                                        ,[HOLDINGSHARES]        
                                        ,[ID]            
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                                 
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
                                           , INCREASEDSHARE
                                           , HOLDINGSHARES
                                           , ID
                                           ,"10"
                                           ,"100000"
                                           );
                    

                    INCREASEDSHARE_COUT++;

                    }
                }
                else if(INCREASEDSHARESHUNDREDTHOUSANDS_COUNT ==1)
                {
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
                                        ,[HOLDINGSHARES]   
                                        ,[ID]     
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                                             
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
                                            , INCREASEDSHARESHUNDREDTHOUSANDS_ST
                                            , HOLDINGSHARES
                                            , ID
                                            , "10"
                                            , "100000"
                                            );
                }

                //增資股票號碼(萬股) 
                //INCREASEDSHARESTENSOFTHOUSANDS_COUNT
                //SQL
                if (INCREASEDSHARESTENSOFTHOUSANDS_COUNT >= 2)
                {
                    //sbSql.Clear();

                    string INCREASEDSHARE = "";
                    string INCREASEDSHARE_PRE = INCREASEDSHARESTENSOFTHOUSANDS_ST.Substring(0, INCREASEDSHARESTENSOFTHOUSANDS_ST.Length - 7);
                    int INCREASEDSHARE_COUT = Convert.ToInt32(INCREASEDSHARESTENSOFTHOUSANDS_ST.Substring(INCREASEDSHARESTENSOFTHOUSANDS_ST.Length - 7, 7));
                    for (int i = 1; i <= INCREASEDSHARESTENSOFTHOUSANDS_COUNT; i++)
                    {
                        INCREASEDSHARE = INCREASEDSHARE_PRE + PadNumberWithZero7(INCREASEDSHARE_COUT);

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
                                        ,[INCREASEDSHARESTENSOFTHOUSANDS]                                      
                                        ,[HOLDINGSHARES]        
                                        ,[ID]       
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                           
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
                                           , INCREASEDSHARE
                                           , HOLDINGSHARES
                                           , ID
                                           , "10"
                                           , "10000"
                                           );


                        INCREASEDSHARE_COUT++;

                    }
                }
                else if (INCREASEDSHARESTENSOFTHOUSANDS_COUNT == 1)
                {
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
                                        ,[INCREASEDSHARESTENSOFTHOUSANDS]                                      
                                        ,[HOLDINGSHARES]   
                                        ,[ID]      
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                                 
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
                                            , INCREASEDSHARESTENSOFTHOUSANDS_ST
                                            , HOLDINGSHARES
                                            , ID
                                            , "10"
                                           , "10000"
                                            );
                }

                //增資股票號碼(千股) 
                //INCREASEDSHARESTHOUSANDS_COUNT
                //SQL
                if (INCREASEDSHARESTHOUSANDS_COUNT >= 2)
                {
                    //sbSql.Clear();

                    string INCREASEDSHARE = "";
                    string INCREASEDSHARE_PRE = INCREASEDSHARESTHOUSANDS_ST.Substring(0, INCREASEDSHARESTHOUSANDS_ST.Length - 7);
                    int INCREASEDSHARE_COUT = Convert.ToInt32(INCREASEDSHARESTHOUSANDS_ST.Substring(INCREASEDSHARESTHOUSANDS_ST.Length - 7, 7));
                    for (int i = 1; i <= INCREASEDSHARESTHOUSANDS_COUNT; i++)
                    {
                        INCREASEDSHARE = INCREASEDSHARE_PRE + PadNumberWithZero7(INCREASEDSHARE_COUT);

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
                                        ,[INCREASEDSHARESTHOUSANDS]                                      
                                        ,[HOLDINGSHARES]        
                                        ,[ID]       
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                          
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
                                           , INCREASEDSHARE
                                           , HOLDINGSHARES
                                           , ID
                                           , "10"
                                           , "1000"
                                           );


                        INCREASEDSHARE_COUT++;

                    }
                }
                else if (INCREASEDSHARESTHOUSANDS_COUNT == 1)
                {
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
                                        ,[INCREASEDSHARESTHOUSANDS]                                      
                                        ,[HOLDINGSHARES]   
                                        ,[ID]   
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                                  
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
                                            , INCREASEDSHARESTHOUSANDS_ST
                                            , HOLDINGSHARES
                                            , ID
                                            , "10"
                                             , "1000"
                                            );
                }

                //增資股票號碼(不定額股)
                //INCREASEDSHARESIRREGULARLOTS_COUNT
                //SQL
                if (INCREASEDSHARESIRREGULARLOTS_COUNT >= 2)
                {
                    //sbSql.Clear();

                    string INCREASEDSHARE = "";
                    string INCREASEDSHARE_PRE = INCREASEDSHARESIRREGULARLOTS_ST.Substring(0, INCREASEDSHARESIRREGULARLOTS_ST.Length - 7);
                    int INCREASEDSHARE_COUT = Convert.ToInt32(INCREASEDSHARESIRREGULARLOTS_ST.Substring(INCREASEDSHARESIRREGULARLOTS_ST.Length - 7, 7));
                    for (int i = 1; i <= INCREASEDSHARESIRREGULARLOTS_COUNT; i++)
                    {
                        INCREASEDSHARE = INCREASEDSHARE_PRE + PadNumberWithZero7(INCREASEDSHARE_COUT);

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
                                        ,[INCREASEDSHARESIRREGULARLOTS]                                      
                                        ,[HOLDINGSHARES]        
                                        ,[ID]   
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                            
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
                                           , INCREASEDSHARE
                                           , HOLDINGSHARES
                                           , ID
                                           ,"10"
                                           , STOCKSHARES
                                           );


                        INCREASEDSHARE_COUT++;

                    }
                }
                else if (INCREASEDSHARESIRREGULARLOTS_COUNT == 1)
                {
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
                                        ,[INCREASEDSHARESIRREGULARLOTS]                                      
                                        ,[HOLDINGSHARES]   
                                        ,[ID]    
                                        ,[PARVALUPER]
                                        ,[STOCKSHARES]                                  
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
                                            , INCREASEDSHARESIRREGULARLOTS_ST
                                            , HOLDINGSHARES
                                            , ID
                                            , "10"
                                            , STOCKSHARES
                                            );
                }

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

                //sbSql.Clear();

                //sbSql.AppendFormat(@"  ");                                  
               


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

        public string PadNumberWithZero7(int number)
        {
            return number.ToString().PadLeft(7, '0');
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
                                    DELETE  [TKACT].[dbo].[TKSTOCKSREORDS]
                                    WHERE [STOCKID] IN (
                                    SELECT 
                                    ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'')
                                    FROM [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                    WHERE SERNO='{0}'
                                    )

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

        public void TKSTOCKSTRANS_ADD(
            string IDFORM
            , string IDTO
            , string DATEOFCHANGE
            , string REASOFORCHANGE
            , string STOCKACCOUNTNUMBERFORM
            , string STOCKNAMEFORM
            , string STOCKACCOUNTNUMBERTO
            , string STOCKNAMETO
            , string TRANSFERREDSHARES
            , string PARVALUEPERSHARE
            , string TRADINGPRICEPERSHARE
            , string TOTALTRADINGAMOUNT
            , string SECURITIESTRANSACTIONTAXAMOUNT
            , string TRANSFERREDSHARESHUNDREDTHOUSANDS_ST
            , string TRANSFERREDSHARESTENSOFTHOUSANDS_ST
            , string TRANSFERREDSHARESTHOUSANDS_ST
            , string TRANSFERREDSHARESIRREGULARLOTS_ST
            , string HOLDINGSHARES
            , string TRANSFERREDSHARESHUNDREDTHOUSANDS_END
            , string TRANSFERREDSHARESTENSOFTHOUSANDS_END
            , string TRANSFERREDSHARESTHOUSANDS_END
            , string TRANSFERREDSHARESIRREGULARLOTS_END
            )
        {
            SqlConnection sqlConn = new SqlConnection();
            SqlCommand sqlComm = new SqlCommand();
            sbSql.Clear();

            int SHARESHUNDREDTHOUSANDS_COUNT = 0;
            int SHARESTENSOFTHOUSANDS_COUNT = 0;
            int SHARESTHOUSANDS_COUNT = 0;
            int SHARESIRREGULARLOTS_COUNT = 0;

            //INCREASEDSHARESHUNDREDTHOUSANDS_COUNT
            if (TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.Length >= 7 && TRANSFERREDSHARESHUNDREDTHOUSANDS_END.Length >= 7 && !string.IsNullOrEmpty(TRANSFERREDSHARESHUNDREDTHOUSANDS_ST) && !string.IsNullOrEmpty(TRANSFERREDSHARESHUNDREDTHOUSANDS_END))
            {
                int START = 0;
                int END = 0;

                START = Convert.ToInt32(TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.ToString().Substring(TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.Length - 7, 7));
                END = Convert.ToInt32(TRANSFERREDSHARESHUNDREDTHOUSANDS_END.ToString().Substring(TRANSFERREDSHARESHUNDREDTHOUSANDS_END.Length - 7, 7));

                SHARESHUNDREDTHOUSANDS_COUNT = END - START + 1;

            }
            else if (!string.IsNullOrEmpty(TRANSFERREDSHARESHUNDREDTHOUSANDS_ST))
            {
                SHARESHUNDREDTHOUSANDS_COUNT = 1;
            }
            else
            {
                SHARESHUNDREDTHOUSANDS_COUNT = 0;
            }
            //INCREASEDSHARESTENSOFTHOUSANDS_COUNT
            if (TRANSFERREDSHARESTENSOFTHOUSANDS_ST.Length >= 7 && TRANSFERREDSHARESTENSOFTHOUSANDS_END.Length >= 7 && !string.IsNullOrEmpty(TRANSFERREDSHARESTENSOFTHOUSANDS_ST) && !string.IsNullOrEmpty(TRANSFERREDSHARESTENSOFTHOUSANDS_END))
            {
                int START = 0;
                int END = 0;

                START = Convert.ToInt32(TRANSFERREDSHARESTENSOFTHOUSANDS_ST.ToString().Substring(TRANSFERREDSHARESTENSOFTHOUSANDS_ST.Length - 7, 7));
                END = Convert.ToInt32(TRANSFERREDSHARESTENSOFTHOUSANDS_END.ToString().Substring(TRANSFERREDSHARESTENSOFTHOUSANDS_END.Length - 7, 7));

                SHARESTENSOFTHOUSANDS_COUNT = END - START + 1;

            }
            else if (!string.IsNullOrEmpty(TRANSFERREDSHARESTENSOFTHOUSANDS_ST))
            {
                SHARESTENSOFTHOUSANDS_COUNT = 1;
            }
            else
            {
                SHARESTENSOFTHOUSANDS_COUNT = 0;
            }
            //INCREASEDSHARESTHOUSANDS_COUNT
            if (TRANSFERREDSHARESTHOUSANDS_ST.Length >= 7 && TRANSFERREDSHARESTHOUSANDS_END.Length >= 7 && !string.IsNullOrEmpty(TRANSFERREDSHARESTHOUSANDS_ST) && !string.IsNullOrEmpty(TRANSFERREDSHARESTHOUSANDS_END))
            {
                int START = 0;
                int END = 0;

                START = Convert.ToInt32(TRANSFERREDSHARESTHOUSANDS_ST.ToString().Substring(TRANSFERREDSHARESTHOUSANDS_ST.Length - 7, 7));
                END = Convert.ToInt32(TRANSFERREDSHARESTHOUSANDS_END.ToString().Substring(TRANSFERREDSHARESTHOUSANDS_END.Length - 7, 7));

                SHARESTHOUSANDS_COUNT = END - START + 1;

            }
            else if (!string.IsNullOrEmpty(TRANSFERREDSHARESTHOUSANDS_ST))
            {
                SHARESTHOUSANDS_COUNT = 1;
            }
            else
            {
                SHARESTHOUSANDS_COUNT = 0;
            }
            //INCREASEDSHARESIRREGULARLOTS_COUNT
            if (TRANSFERREDSHARESIRREGULARLOTS_ST.Length >= 7 && TRANSFERREDSHARESIRREGULARLOTS_END.Length >= 7 && !string.IsNullOrEmpty(TRANSFERREDSHARESIRREGULARLOTS_ST) && !string.IsNullOrEmpty(TRANSFERREDSHARESIRREGULARLOTS_END))
            {
                int START = 0;
                int END = 0;

                START = Convert.ToInt32(TRANSFERREDSHARESIRREGULARLOTS_ST.ToString().Substring(TRANSFERREDSHARESIRREGULARLOTS_ST.Length - 7, 7));
                END = Convert.ToInt32(TRANSFERREDSHARESIRREGULARLOTS_END.ToString().Substring(TRANSFERREDSHARESIRREGULARLOTS_END.Length - 7, 7));

                SHARESIRREGULARLOTS_COUNT = END - START + 1;

            }
            else if (!string.IsNullOrEmpty(TRANSFERREDSHARESIRREGULARLOTS_ST))
            {
                SHARESIRREGULARLOTS_COUNT = 1;
            }
            else
            {
                SHARESIRREGULARLOTS_COUNT = 0;
            }

            //轉讓股票號碼(十萬股)
            //INCREASEDSHARESHUNDREDTHOUSANDS_COUNT
            //SQL
            if (SHARESHUNDREDTHOUSANDS_COUNT >= 2)
            {
                //sbSql.Clear();

                string SHARE = "";
                string SHARE_PRE = TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.Substring(0, TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.Length - 7);
                int SHARE_COUT = Convert.ToInt32(TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.Substring(TRANSFERREDSHARESHUNDREDTHOUSANDS_ST.Length - 7, 7));
                for (int i = 1; i <= SHARESHUNDREDTHOUSANDS_COUNT; i++)
                {
                    SHARE = SHARE_PRE + PadNumberWithZero7(SHARE_COUT);

                    sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESHUNDREDTHOUSANDS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                        ,IDFORM
                                        ,IDTO
                                        ,DATEOFCHANGE
                                        ,REASOFORCHANGE
                                        ,STOCKACCOUNTNUMBERFORM
                                        ,STOCKNAMEFORM
                                        ,STOCKACCOUNTNUMBERTO
                                        ,STOCKNAMETO
                                        ,TRANSFERREDSHARES
                                        ,PARVALUEPERSHARE
                                        ,TRADINGPRICEPERSHARE
                                        ,TOTALTRADINGAMOUNT
                                        ,SECURITIESTRANSACTIONTAXAMOUNT
                                        ,SHARE
                                        ,HOLDINGSHARES
                                       );


                    SHARE_COUT++;

                }
            }
            else if (SHARESHUNDREDTHOUSANDS_COUNT == 1)
            {
                sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESHUNDREDTHOUSANDS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                         , IDFORM
                                         , IDTO
                                         , DATEOFCHANGE
                                         , REASOFORCHANGE
                                         , STOCKACCOUNTNUMBERFORM
                                         , STOCKNAMEFORM
                                         , STOCKACCOUNTNUMBERTO
                                         , STOCKNAMETO
                                         , TRANSFERREDSHARES
                                         , PARVALUEPERSHARE
                                         , TRADINGPRICEPERSHARE
                                         , TOTALTRADINGAMOUNT
                                         , SECURITIESTRANSACTIONTAXAMOUNT
                                         , TRANSFERREDSHARESHUNDREDTHOUSANDS_ST
                                         , HOLDINGSHARES
                                        );
            }

            //轉讓股票號碼(萬股)
            //TRANSFERREDSHARESTENSOFTHOUSANDS
            //SQL
            if (SHARESTENSOFTHOUSANDS_COUNT >= 2)
            {
                //sbSql.Clear();

                string SHARE = "";
                string SHARE_PRE = TRANSFERREDSHARESTENSOFTHOUSANDS_ST.Substring(0, TRANSFERREDSHARESTENSOFTHOUSANDS_ST.Length - 7);
                int SHARE_COUT = Convert.ToInt32(TRANSFERREDSHARESTENSOFTHOUSANDS_ST.Substring(TRANSFERREDSHARESTENSOFTHOUSANDS_ST.Length - 7, 7));
                for (int i = 1; i <= SHARESTENSOFTHOUSANDS_COUNT; i++)
                {
                    SHARE = SHARE_PRE + PadNumberWithZero7(SHARE_COUT);

                    sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESTENSOFTHOUSANDS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                        , IDFORM
                                        , IDTO
                                        , DATEOFCHANGE
                                        , REASOFORCHANGE
                                        , STOCKACCOUNTNUMBERFORM
                                        , STOCKNAMEFORM
                                        , STOCKACCOUNTNUMBERTO
                                        , STOCKNAMETO
                                        , TRANSFERREDSHARES
                                        , PARVALUEPERSHARE
                                        , TRADINGPRICEPERSHARE
                                        , TOTALTRADINGAMOUNT
                                        , SECURITIESTRANSACTIONTAXAMOUNT
                                        , SHARE
                                        , HOLDINGSHARES
                                       );


                    SHARE_COUT++;

                }
            }
            else if (SHARESTENSOFTHOUSANDS_COUNT == 1)
            {
                sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESTENSOFTHOUSANDS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                         , IDFORM
                                         , IDTO
                                         , DATEOFCHANGE
                                         , REASOFORCHANGE
                                         , STOCKACCOUNTNUMBERFORM
                                         , STOCKNAMEFORM
                                         , STOCKACCOUNTNUMBERTO
                                         , STOCKNAMETO
                                         , TRANSFERREDSHARES
                                         , PARVALUEPERSHARE
                                         , TRADINGPRICEPERSHARE
                                         , TOTALTRADINGAMOUNT
                                         , SECURITIESTRANSACTIONTAXAMOUNT
                                         , TRANSFERREDSHARESTENSOFTHOUSANDS_ST
                                         , HOLDINGSHARES
                                        );
            }

            //轉讓股票號碼(千股)
            //SHARESTHOUSANDS_COUNT
            //SQL
            if (SHARESTHOUSANDS_COUNT >= 2)
            {
                //sbSql.Clear();

                string SHARE = "";
                string SHARE_PRE = TRANSFERREDSHARESTHOUSANDS_ST.Substring(0, TRANSFERREDSHARESTHOUSANDS_ST.Length - 7);
                int SHARE_COUT = Convert.ToInt32(TRANSFERREDSHARESTHOUSANDS_ST.Substring(TRANSFERREDSHARESTHOUSANDS_ST.Length - 7, 7));
                for (int i = 1; i <= SHARESTHOUSANDS_COUNT; i++)
                {
                    SHARE = SHARE_PRE + PadNumberWithZero7(SHARE_COUT);

                    sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESTHOUSANDS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                        , IDFORM
                                        , IDTO
                                        , DATEOFCHANGE
                                        , REASOFORCHANGE
                                        , STOCKACCOUNTNUMBERFORM
                                        , STOCKNAMEFORM
                                        , STOCKACCOUNTNUMBERTO
                                        , STOCKNAMETO
                                        , TRANSFERREDSHARES
                                        , PARVALUEPERSHARE
                                        , TRADINGPRICEPERSHARE
                                        , TOTALTRADINGAMOUNT
                                        , SECURITIESTRANSACTIONTAXAMOUNT
                                        , SHARE
                                        , HOLDINGSHARES
                                       );


                    SHARE_COUT++;

                }
            }
            else if (SHARESTHOUSANDS_COUNT == 1)
            {
                sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESTHOUSANDS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                         , IDFORM
                                         , IDTO
                                         , DATEOFCHANGE
                                         , REASOFORCHANGE
                                         , STOCKACCOUNTNUMBERFORM
                                         , STOCKNAMEFORM
                                         , STOCKACCOUNTNUMBERTO
                                         , STOCKNAMETO
                                         , TRANSFERREDSHARES
                                         , PARVALUEPERSHARE
                                         , TRADINGPRICEPERSHARE
                                         , TOTALTRADINGAMOUNT
                                         , SECURITIESTRANSACTIONTAXAMOUNT
                                         , TRANSFERREDSHARESTHOUSANDS_ST
                                         , HOLDINGSHARES
                                        );
            }

            //轉讓股票號碼(不定額股)
            //SHARESIRREGULARLOTS_COUNT
            //SQL
            if (SHARESIRREGULARLOTS_COUNT >= 2)
            {
                //sbSql.Clear();

                string SHARE = "";
                string SHARE_PRE = TRANSFERREDSHARESIRREGULARLOTS_ST.Substring(0, TRANSFERREDSHARESIRREGULARLOTS_ST.Length - 7);
                int SHARE_COUT = Convert.ToInt32(TRANSFERREDSHARESIRREGULARLOTS_ST.Substring(TRANSFERREDSHARESIRREGULARLOTS_ST.Length - 7, 7));
                for (int i = 1; i <= SHARESIRREGULARLOTS_COUNT; i++)
                {
                    SHARE = SHARE_PRE + PadNumberWithZero7(SHARE_COUT);

                    sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESIRREGULARLOTS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                        , IDFORM
                                        , IDTO
                                        , DATEOFCHANGE
                                        , REASOFORCHANGE
                                        , STOCKACCOUNTNUMBERFORM
                                        , STOCKNAMEFORM
                                        , STOCKACCOUNTNUMBERTO
                                        , STOCKNAMETO
                                        , TRANSFERREDSHARES
                                        , PARVALUEPERSHARE
                                        , TRADINGPRICEPERSHARE
                                        , TOTALTRADINGAMOUNT
                                        , SECURITIESTRANSACTIONTAXAMOUNT
                                        , SHARE
                                        , HOLDINGSHARES
                                       );


                    SHARE_COUT++;

                }
            }
            else if (SHARESIRREGULARLOTS_COUNT == 1)
            {
                sbSql.AppendFormat(@"
                                        INSERT INTO[TKACT].[dbo].[TKSTOCKSTRANS]
                                        (
                                        [IDFORM]
                                        ,[IDTO]
                                        ,[DATEOFCHANGE]
                                        ,[REASOFORCHANGE]
                                        ,[STOCKACCOUNTNUMBERFORM]
                                        ,[STOCKNAMEFORM]
                                        ,[STOCKACCOUNTNUMBERTO]
                                        ,[STOCKNAMETO]
                                        ,[TRANSFERREDSHARES]
                                        ,[PARVALUEPERSHARE]
                                        ,[TRADINGPRICEPERSHARE]
                                        ,[TOTALTRADINGAMOUNT]
                                        ,[SECURITIESTRANSACTIONTAXAMOUNT]
                                        ,[TRANSFERREDSHARESIRREGULARLOTS]
                                        ,[HOLDINGSHARES]
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
                                        )"
                                         , IDFORM
                                         , IDTO
                                         , DATEOFCHANGE
                                         , REASOFORCHANGE
                                         , STOCKACCOUNTNUMBERFORM
                                         , STOCKNAMEFORM
                                         , STOCKACCOUNTNUMBERTO
                                         , STOCKNAMETO
                                         , TRANSFERREDSHARES
                                         , PARVALUEPERSHARE
                                         , TRADINGPRICEPERSHARE
                                         , TOTALTRADINGAMOUNT
                                         , SECURITIESTRANSACTIONTAXAMOUNT
                                         , TRANSFERREDSHARESIRREGULARLOTS_ST
                                         , HOLDINGSHARES
                                        );
            }

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

                //sbSql.Clear();

                //sbSql.AppendFormat(@"                                
                                   
                //                    INSERT INTO  [TKACT].[dbo].[TKSTOCKSTRANS]
                //                    (
                //                    [IDFORM]
                //                    ,[IDTO]
                //                    ,[DATEOFCHANGE]
                //                    ,[REASOFORCHANGE]
                //                    ,[STOCKACCOUNTNUMBERFORM]
                //                    ,[STOCKNAMEFORM]
                //                    ,[STOCKACCOUNTNUMBERTO]
                //                    ,[STOCKNAMETO]
                //                    ,[TRANSFERREDSHARES]
                //                    ,[PARVALUEPERSHARE]
                //                    ,[TRADINGPRICEPERSHARE]
                //                    ,[TOTALTRADINGAMOUNT]
                //                    ,[SECURITIESTRANSACTIONTAXAMOUNT]
                //                    ,[TRANSFERREDSHARESHUNDREDTHOUSANDS]
                //                    ,[TRANSFERREDSHARESTENSOFTHOUSANDS]
                //                    ,[TRANSFERREDSHARESTHOUSANDS]
                //                    ,[TRANSFERREDSHARESIRREGULARLOTS]
                //                    ,[HOLDINGSHARES]
                //                    )
                //                    VALUES
                //                    (
                //                    '{0}'
                //                    ,'{1}'
                //                    ,'{2}'
                //                    ,'{3}'
                //                    ,'{4}'
                //                    ,'{5}'
                //                    ,'{6}'
                //                    ,'{7}'
                //                    ,'{8}'
                //                    ,'{9}'
                //                    ,'{10}'
                //                    ,'{11}'
                //                    ,'{12}'
                //                    ,'{13}'
                //                    ,'{14}'
                //                    ,'{15}'
                //                    ,'{16}'
                //                    ,'{17}'

                //                    )
                                   
                                        
                //                        ",IDFORM
                //                    ,IDTO
                //                    ,DATEOFCHANGE
                //                    ,REASOFORCHANGE
                //                    ,STOCKACCOUNTNUMBERFORM
                //                    ,STOCKNAMEFORM
                //                    ,STOCKACCOUNTNUMBERTO
                //                    ,STOCKNAMETO
                //                    ,TRANSFERREDSHARES
                //                    ,PARVALUEPERSHARE
                //                    ,TRADINGPRICEPERSHARE
                //                    ,TOTALTRADINGAMOUNT
                //                    ,SECURITIESTRANSACTIONTAXAMOUNT
                //                    ,TRANSFERREDSHARESHUNDREDTHOUSANDS
                //                    ,TRANSFERREDSHARESTENSOFTHOUSANDS
                //                    ,TRANSFERREDSHARESTHOUSANDS
                //                    ,TRANSFERREDSHARESIRREGULARLOTS
                //                    ,HOLDINGSHARES

                //                        );


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

        public void TKSTOCKSTRANS_UPDATE(
            string SERNO
            , string IDFORM
            , string IDTO
            , string DATEOFCHANGE
            , string REASOFORCHANGE
            , string STOCKACCOUNTNUMBERFORM
            , string STOCKNAMEFORM
            , string STOCKACCOUNTNUMBERTO
            , string STOCKNAMETO
            , string TRANSFERREDSHARES
            , string PARVALUEPERSHARE
            , string TRADINGPRICEPERSHARE
            , string TOTALTRADINGAMOUNT
            , string SECURITIESTRANSACTIONTAXAMOUNT
            , string TRANSFERREDSHARESHUNDREDTHOUSANDS
            , string TRANSFERREDSHARESTENSOFTHOUSANDS
            , string TRANSFERREDSHARESTHOUSANDS
            , string TRANSFERREDSHARESIRREGULARLOTS
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
                                   
                                    UPDATE  [TKACT].[dbo].[TKSTOCKSTRANS]
                                    SET
                                    [IDFORM]='{1}'
                                    ,[IDTO]='{2}'
                                    ,[DATEOFCHANGE]='{3}'
                                    ,[REASOFORCHANGE]='{4}'
                                    ,[STOCKACCOUNTNUMBERFORM]='{5}'
                                    ,[STOCKNAMEFORM]='{6}'
                                    ,[STOCKACCOUNTNUMBERTO]='{7}'
                                    ,[STOCKNAMETO]='{8}'
                                    ,[TRANSFERREDSHARES]='{9}'
                                    ,[PARVALUEPERSHARE]='{10}'
                                    ,[TRADINGPRICEPERSHARE]='{11}'
                                    ,[TOTALTRADINGAMOUNT]='{12}'
                                    ,[SECURITIESTRANSACTIONTAXAMOUNT]='{13}'
                                    ,[TRANSFERREDSHARESHUNDREDTHOUSANDS]='{14}'
                                    ,[TRANSFERREDSHARESTENSOFTHOUSANDS]='{15}'
                                    ,[TRANSFERREDSHARESTHOUSANDS]='{16}'
                                    ,[TRANSFERREDSHARESIRREGULARLOTS]='{17}'
                                    ,[HOLDINGSHARES]='{18}'
                                    WHERE [SERNO]='{0}'
                                    

                                    ", SERNO
                                    ,  IDFORM
                                    ,  IDTO
                                    ,  DATEOFCHANGE
                                    ,  REASOFORCHANGE
                                    ,  STOCKACCOUNTNUMBERFORM
                                    ,  STOCKNAMEFORM
                                    ,  STOCKACCOUNTNUMBERTO
                                    ,  STOCKNAMETO
                                    ,  TRANSFERREDSHARES
                                    ,  PARVALUEPERSHARE
                                    ,  TRADINGPRICEPERSHARE
                                    ,  TOTALTRADINGAMOUNT
                                    ,  SECURITIESTRANSACTIONTAXAMOUNT
                                    ,  TRANSFERREDSHARESHUNDREDTHOUSANDS
                                    ,  TRANSFERREDSHARESTENSOFTHOUSANDS
                                    ,  TRANSFERREDSHARESTHOUSANDS
                                    ,  TRANSFERREDSHARESIRREGULARLOTS
                                    ,  HOLDINGSHARES

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

        public void TKSTOCKSTRANS_DELETE(string SERNO)
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
                                   UPDATE [TKACT].[dbo].[TKSTOCKSREORDS]
                                    SET  [STOCKIDKEY]=TEMP.[IDFORM],[STOCKACCOUNTNUMBER]=TEMP.[STOCKACCOUNTNUMBERFORM],[STOCKNAME]=TEMP.[STOCKNAMEFORM]
                                    FROM 
                                    (
                                    SELECT ISNULL([TRANSFERREDSHARESHUNDREDTHOUSANDS],'')+ISNULL([TRANSFERREDSHARESTENSOFTHOUSANDS],'')+ISNULL([TRANSFERREDSHARESTHOUSANDS],'')+ISNULL([TRANSFERREDSHARESIRREGULARLOTS],'') AS 'STOCKID',[IDFORM],[STOCKACCOUNTNUMBERFORM],[STOCKNAMEFORM]
                                    FROM [TKACT].[dbo].[TKSTOCKSTRANS]
                                    WHERE SERNO='{0}'
                                    ) AS TEMP
                                    WHERE TEMP.STOCKID=[TKSTOCKSREORDS].STOCKID

                                    DELETE  [TKACT].[dbo].[TKSTOCKSTRANS]                                   
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

        public void TKSTOCKSDIV_ADD(
            string STOCKACCOUNTNUMBER
            , string STOCKNAME
            , string EXDIVIDENDINTERESTRECORDDATE
            , string CASHDIVIDENDPAYMENTDATE
            , string CASHDIVIDENDPERSHARE
            , string STOCKDIVIDEND
            , string DECLAREDCASHDIVIDEND
            , string DECLAREDSTOCKDIVIDEND
            , string SUPPLEMENTARYPREMIUMTOBEDEDUCTED
            , string ACTUALCASHDIVIDENDPAID
            , string CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES
            , string CAPITALIZATIONOFCAPITALSURPLUSSHARES
            , string ID
            , string DIVYEARS
            , string DIVAMOUNTS
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
                                    INSERT INTO [TKACT].[dbo].[TKSTOCKSDIV] 
                                    (
                                    [STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[EXDIVIDENDINTERESTRECORDDATE]
                                    ,[CASHDIVIDENDPAYMENTDATE]
                                    ,[CASHDIVIDENDPERSHARE]
                                    ,[STOCKDIVIDEND]
                                    ,[DECLAREDCASHDIVIDEND]
                                    ,[DECLAREDSTOCKDIVIDEND]
                                    ,[SUPPLEMENTARYPREMIUMTOBEDEDUCTED]
                                    ,[ACTUALCASHDIVIDENDPAID]
                                    ,[CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES]
                                    ,[CAPITALIZATIONOFCAPITALSURPLUSSHARES]
                                    ,[DIVYEARS]
                                    ,[DIVAMOUNTS]
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
                                    ,'{14}'
                                    )

                                        ", STOCKACCOUNTNUMBER
                                    , STOCKNAME
                                    , EXDIVIDENDINTERESTRECORDDATE
                                    , CASHDIVIDENDPAYMENTDATE
                                    , CASHDIVIDENDPERSHARE
                                    , STOCKDIVIDEND
                                    , DECLAREDCASHDIVIDEND
                                    , DECLAREDSTOCKDIVIDEND
                                    , SUPPLEMENTARYPREMIUMTOBEDEDUCTED
                                    , ACTUALCASHDIVIDENDPAID
                                    , CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES
                                    , CAPITALIZATIONOFCAPITALSURPLUSSHARES
                                    , DIVYEARS
                                    , DIVAMOUNTS
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

        public void TKSTOCKSDIV_UPDATE(
             string SERNO
            , string STOCKACCOUNTNUMBER
            , string STOCKNAME
            , string EXDIVIDENDINTERESTRECORDDATE
            , string CASHDIVIDENDPAYMENTDATE
            , string CASHDIVIDENDPERSHARE
            , string STOCKDIVIDEND
            , string DECLAREDCASHDIVIDEND
            , string DECLAREDSTOCKDIVIDEND
            , string SUPPLEMENTARYPREMIUMTOBEDEDUCTED
            , string ACTUALCASHDIVIDENDPAID
            , string CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES
            , string CAPITALIZATIONOFCAPITALSURPLUSSHARES
            , string DIVYEARS
            , string DIVAMOUNTS

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
                                   UPDATE  [TKACT].[dbo].[TKSTOCKSDIV] 
                                    SET 
                                    [STOCKACCOUNTNUMBER]='{1}'
                                    ,[STOCKNAME]='{2}'
                                    ,[EXDIVIDENDINTERESTRECORDDATE]='{3}'
                                    ,[CASHDIVIDENDPAYMENTDATE]='{4}'
                                    ,[CASHDIVIDENDPERSHARE]='{5}'
                                    ,[STOCKDIVIDEND]='{6}'
                                    ,[DECLAREDCASHDIVIDEND]='{7}'
                                    ,[DECLAREDSTOCKDIVIDEND]='{8}'
                                    ,[SUPPLEMENTARYPREMIUMTOBEDEDUCTED]='{9}'
                                    ,[ACTUALCASHDIVIDENDPAID]='{10}'
                                    ,[CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES]='{11}'
                                    ,[CAPITALIZATIONOFCAPITALSURPLUSSHARES]='{12}'
                                    ,[DIVYEARS]='{13}'
                                    ,[DIVAMOUNTS]='{14}'
                                    WHERE  [SERNO]='{0}'
                                   
                                    ", SERNO
                                    , STOCKACCOUNTNUMBER
                                    , STOCKNAME
                                    , EXDIVIDENDINTERESTRECORDDATE
                                    , CASHDIVIDENDPAYMENTDATE
                                    , CASHDIVIDENDPERSHARE
                                    , STOCKDIVIDEND
                                    , DECLAREDCASHDIVIDEND
                                    , DECLAREDSTOCKDIVIDEND
                                    , SUPPLEMENTARYPREMIUMTOBEDEDUCTED
                                    , ACTUALCASHDIVIDENDPAID
                                    , CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES
                                    , CAPITALIZATIONOFCAPITALSURPLUSSHARES
                                    , DIVYEARS
                                    , DIVAMOUNTS
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

        public void TKSTOCKSDIV_DELETE(string SERNO)
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
                                   
                                    DELETE  [TKACT].[dbo].[TKSTOCKSDIV]                                   
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

        public DataTable FINE_TKSTOCKSNAMES_STOCKACCOUNTNUMBER(string STOCKNAME)
        {
            DataTable DT = new DataTable();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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


                sbSql.Clear();
                QUERYS.Clear();
                
                
                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [ID]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    FROM [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [STOCKNAME] LIKE '%{0}%'

                                    ", STOCKNAME);




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                {
                    return ds1.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }

        public DataTable FINE_TKSTOCKSNAMES_STOCKNAME(string STOCKACCOUNTNUMBER)
        {
            DataTable DT = new DataTable();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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


                sbSql.Clear();
                QUERYS.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT 
                                    [ID]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    FROM [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [STOCKACCOUNTNUMBER] LIKE '{0}%'

                                    ", STOCKACCOUNTNUMBER);




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                {
                    return ds1.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                return null;
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

            //戶籍/設立地址
            if (TEXTBOXIN.Name.Equals("textBox9"))
            {
                if (string.IsNullOrEmpty(textBox9.Text))
                {
                    MESSAGES = MESSAGES + " 戶籍/設立地址 不可空白";
                }
            }
            //戶籍/設立地址
            if (TEXTBOXIN.Name.Equals("textBox32"))
            {
                if (string.IsNullOrEmpty(textBox32.Text))
                {
                    MESSAGES = MESSAGES + " 戶籍/設立地址 不可空白";
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
                Int64 roundedResult = (Int64)Math.Round(result);
                string roundedResultString = roundedResult.ToString();
                textBox51.Text = roundedResultString;
            }
        }
        private void textBox80_TextChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX83();
        }

        private void textBox82_TextChanged(object sender, EventArgs e)
        {
            SET_TEXTBOX83();
        }
        public void SET_TEXTBOX83()
        {
            if (!string.IsNullOrEmpty(textBox80.Text) && !string.IsNullOrEmpty(textBox82.Text))
            {
                decimal result = Convert.ToDecimal(textBox80.Text) * Convert.ToDecimal(textBox82.Text);
                Int64 roundedResult = (Int64)Math.Round(result);
                string roundedResultString = roundedResult.ToString();
                textBox83.Text = roundedResultString;
                
            }
        }



        private void textBox83_TextChanged(object sender, EventArgs e)
        {
            if(!string.IsNullOrEmpty(textBox83.Text))
            {
                if (comboBox3.SelectedValue.ToString().Equals("贈與"))
                {
                    textBox84.Text = "0";
                }
                else
                {
                    decimal result = Convert.ToDecimal(textBox83.Text) * 3 / 1000;
                    Int64 roundedResult = (Int64)Math.Round(result);
                    string roundedResultString = roundedResult.ToString();
                    textBox84.Text = roundedResultString;
                }
            }
           
        }

        private void textBox46_TextChanged(object sender, EventArgs e)
        {
            
            
        }

        private void textBox46_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox47.Text = "";
            if (!string.IsNullOrEmpty(textBox46.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox46.Text);
                if (DT != null)
                {
                    textBox47.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox57.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox47.Text = "";
                    textBox57.Text = "";
                }
            }
           
        }

        private void textBox46_Leave(object sender, EventArgs e)
        {
            textBox47.Text = "";
            if (!string.IsNullOrEmpty(textBox46.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox46.Text);
                if (DT != null)
                {
                    textBox47.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox57.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox47.Text = "";
                    textBox57.Text = "";
                }
            }
        }

        private void textBox4_KeyPress(object sender, KeyPressEventArgs e)
        {
           
               
        }

        private void textBox47_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox46.Text = "";
            if (!string.IsNullOrEmpty(textBox47.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKACCOUNTNUMBER(textBox47.Text);
                if (DT != null)
                {
                    textBox46.Text = DT.Rows[0]["STOCKACCOUNTNUMBER"].ToString();
                    textBox57.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox46.Text = "";
                    textBox57.Text = "";
                }
            }
        }

        private void textBox76_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox77.Text = "";
            if (!string.IsNullOrEmpty(textBox76.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox76.Text);
                if (DT != null)
                {
                    textBox77.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox74.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox77.Text = "";
                    textBox74.Text = "";
                }
            }
        }

        private void textBox76_Leave(object sender, EventArgs e)
        {
            textBox77.Text = "";
            if (!string.IsNullOrEmpty(textBox76.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox76.Text);
                if (DT != null)
                {
                    textBox77.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox74.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox77.Text = "";
                    textBox74.Text = "";
                }
            }
        }


        private void textBox77_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox76.Text = "";
            if (!string.IsNullOrEmpty(textBox77.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKACCOUNTNUMBER(textBox77.Text);
                if (DT != null)
                {
                    textBox76.Text = DT.Rows[0]["STOCKACCOUNTNUMBER"].ToString();
                    textBox75.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox76.Text = "";
                    textBox75.Text = "";
                }
            }
        }

        private void textBox78_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox79.Text = "";
            if (!string.IsNullOrEmpty(textBox78.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox78.Text);
                if (DT != null)
                {
                    textBox79.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox75.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox79.Text = "";
                    textBox75.Text = "";
                }
            }
        }

        private void textBox78_Leave(object sender, EventArgs e)
        {
            textBox79.Text = "";
            if (!string.IsNullOrEmpty(textBox78.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox78.Text);
                if (DT != null)
                {
                    textBox79.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox75.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox79.Text = "";
                    textBox75.Text = "";
                }
            }
        }
        private void textBox79_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox78.Text = "";
            if (!string.IsNullOrEmpty(textBox79.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKACCOUNTNUMBER(textBox79.Text);
                if (DT != null)
                {
                    textBox78.Text = DT.Rows[0]["STOCKACCOUNTNUMBER"].ToString();
                    textBox75.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox78.Text = "";
                    textBox75.Text = "";
                }
            }
        }

        private void textBox52_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox52.Text,textBox131.Text, textBox136);
        }

        private void textBox131_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox52.Text, textBox131.Text, textBox136);
        }

        private void textBox53_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox53.Text, textBox132.Text, textBox137);
        }

        private void textBox132_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox53.Text, textBox132.Text, textBox137);
        }

        private void textBox54_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox54.Text, textBox133.Text, textBox138);
        }

        private void textBox133_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox54.Text, textBox133.Text, textBox138);
        }

        private void textBox55_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox55.Text, textBox134.Text, textBox139);
        }

        private void textBox134_TextChanged(object sender, EventArgs e)
        {
            SET_NUM(textBox55.Text, textBox134.Text, textBox139);
        }

        public void SET_NUM(string START,string END,TextBox TEXTBOXNUM)
        {
            try
            {
                if (START.Length >= 7 && END.Length >= 7 && !string.IsNullOrEmpty(START) && !string.IsNullOrEmpty(END))
                {
                    int STARTNUM = 0;
                    int ENDNUM = 0;

                    STARTNUM = Convert.ToInt32(START.ToString().Substring(START.Length - 7, 7));
                    ENDNUM = Convert.ToInt32(END.ToString().Substring(END.Length - 7, 7));

                    TEXTBOXNUM.Text = (ENDNUM - STARTNUM + 1).ToString();
                }
            }
            catch
            { }
            finally
            { }
            
        }

        public void TKSTOCKSREORDS_ADD()
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
                                    SET ID=[TKSTOCKSNAMES].ID
                                    FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]=[TKSTOCKSTRANSADD].[STOCKACCOUNTNUMBER]
                                    AND [TKSTOCKSTRANSADD].ID<>[TKSTOCKSNAMES].ID


                                   INSERT INTO [TKACT].[dbo].[TKSTOCKSREORDS]
                                    (
                                    [STOCKID]
                                    ,[PARVALUPER]
                                    ,[STOCKSHARES]
                                    ,[STOCKIDKEY]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[VALID]
                                    )
                                    SELECT ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') 
                                    ,[PARVALUPER]
                                    ,[STOCKSHARES]
                                    ,[ID]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,'Y'
                                    FROM  [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                    WHERE (ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') )<>''
                                    AND (ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') ) NOT IN (SELECT [STOCKID] FROM [TKACT].[dbo].[TKSTOCKSREORDS])
                                    AND (ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') ) NOT IN (SELECT [OLDSTOCKID] FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV] WHERE [VALIDS] IN ('Y') GROUP BY [OLDSTOCKID])





                                     "

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

        public void TKSTOCKSREORDS_UPDATE()
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
                                    UPDATE [TKACT].[dbo].[TKSTOCKSTRANS]
                                    SET [IDFORM]=[TKSTOCKSNAMES].ID
                                    FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [TKSTOCKSTRANS].[STOCKACCOUNTNUMBERFORM]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]
                                    AND [TKSTOCKSTRANS].[IDFORM]<>[TKSTOCKSNAMES].ID

                                    UPDATE [TKACT].[dbo].[TKSTOCKSTRANS]
                                    SET [IDTO]=[TKSTOCKSNAMES].ID
                                    FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [TKSTOCKSTRANS].[STOCKACCOUNTNUMBERTO]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]
                                    AND [TKSTOCKSTRANS].[IDTO]<>[TKSTOCKSNAMES].ID                        
                                 
                                    UPDATE [TKACT].[dbo].[TKSTOCKSREORDS]
                                    SET [STOCKIDKEY]=TEMP2.IDTO
                                    ,[STOCKACCOUNTNUMBER]=TEMP2.[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]=TEMP2.[STOCKNAME]
                                    FROM (
	                                    SELECT *
	                                    FROM 
	                                    (
		                                    SELECT
		                                    [STOCKID]
		                                    ,[PARVALUPER]
		                                    ,[STOCKSHARES]
		                                    ,[STOCKIDKEY]
		                                    ,(SELECT TOP 1 [IDTO] FROM  [TKACT].[dbo].[TKSTOCKSTRANS] WHERE ([TRANSFERREDSHARESHUNDREDTHOUSANDS]=[STOCKID]  OR [TRANSFERREDSHARESTENSOFTHOUSANDS]=[STOCKID] OR [TRANSFERREDSHARESTHOUSANDS]=[STOCKID] OR [TRANSFERREDSHARESIRREGULARLOTS]=[STOCKID] ) ORDER BY [SERNO] DESC ) AS 'IDTO'
		                                    FROM [TKACT].[dbo].[TKSTOCKSREORDS]
	                                    ) AS TEMP
	                                    LEFT JOIN [TKACT].[dbo].[TKSTOCKSNAMES] ON [TKSTOCKSNAMES].ID=IDTO
	                                    WHERE ISNULL(TEMP.IDTO,'')<>''
                                    ) AS TEMP2
                                    WHERE TEMP2.STOCKID=[TKSTOCKSREORDS].STOCKID
                                    AND [TKSTOCKSREORDS].[STOCKIDKEY]<>TEMP2.IDTO
                                    "

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

        public void TKSTOCKSREORDS_UPDATE_BEFROEDELETE(string SERNO)
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
                                    
                                   UPDATE [TKACT].[dbo].[TKSTOCKSREORDS]
                                    SET [STOCKIDKEY]=TEMP.IDFORM
                                    ,[STOCKACCOUNTNUMBER]=TEMP.[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]=TEMP.[STOCKNAME]
                                    FROM 
                                    (
	                                    SELECT *
	                                    FROM 
	                                    (
		                                    SELECT
		                                    [STOCKID]
		                                    ,[PARVALUPER]
		                                    ,[STOCKSHARES]
		                                    ,[STOCKIDKEY]
		                                    ,(SELECT TOP 1 [IDFORM] FROM  [TKACT].[dbo].[TKSTOCKSTRANS] WHERE ([TRANSFERREDSHARESHUNDREDTHOUSANDS]=[STOCKID]  OR [TRANSFERREDSHARESTENSOFTHOUSANDS]=[STOCKID] OR [TRANSFERREDSHARESTHOUSANDS]=[STOCKID] OR [TRANSFERREDSHARESIRREGULARLOTS]=[STOCKID] ) AND SERNO='{0}' ORDER BY [SERNO] DESC ) AS 'IDFORM'
		                                    FROM [TKACT].[dbo].[TKSTOCKSREORDS]
	                                    ) AS TEMP
	                                    LEFT JOIN  [TKACT].[dbo].[TKSTOCKSNAMES] ON [TKSTOCKSNAMES].ID=IDFORM
	                                    WHERE ISNULL(TEMP.IDFORM,'')<>''
                                    ) AS TEMP
                                    WHERE TEMP.STOCKID=[TKSTOCKSREORDS].STOCKID
                                    AND [TKSTOCKSREORDS].[STOCKIDKEY]<>TEMP.IDFORM



                                     ", SERNO );


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
        private void textBox109_KeyPress(object sender, KeyPressEventArgs e)
        {
            textBox110.Text = "";
            if (!string.IsNullOrEmpty(textBox109.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox109.Text);
                if (DT != null)
                {
                    textBox110.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox63.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox110.Text = "";
                    textBox63.Text = "";
                }
            }
        }

        public void SETFASTREPORT(string REPORTS,string STOCKACCOUNTNUMBER)
        {          
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();

            if (REPORTS.Equals("股票明細清單"))
            {
                report1.Load(@"REPORT\股權明細.frx");

                SQL = SETSQL(STOCKACCOUNTNUMBER);
            } 
            else if (REPORTS.Equals("股東名冊"))
            {
                report1.Load(@"REPORT\股東名冊ALL.frx");

                SQL= SETSQL2(STOCKACCOUNTNUMBER);
            }
            else if (REPORTS.Equals("每股投資損益彙總表"))
            {
                report1.Load(@"REPORT\每股投資損益彙總表.frx");

                SQL = SETSQL3(STOCKACCOUNTNUMBER);
            }
            else
            {

            }


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();        

            report1.Preview = previewControl1;
            report1.Show();
        }

        public StringBuilder SETSQL(string STOCKACCOUNTNUMBER)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();

            if(!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                SBQUERY1.AppendFormat(@" AND ([TKSTOCKSREORDS].STOCKACCOUNTNUMBER LIKE '%{0}%'  OR [TKSTOCKSREORDS].STOCKNAME LIKE '%{0}%')", STOCKACCOUNTNUMBER);
            }
            else
            {
                SBQUERY1.AppendFormat(@" ");
            }

            SB.AppendFormat(@" 
                            
                            SELECT 
                            [TKSTOCKSREORDS].[STOCKID] AS '股票號碼'
                            ,[TKSTOCKSREORDS].[PARVALUPER] 
                            ,CONVERT(INT,[TKSTOCKSREORDS].[STOCKSHARES] )AS '股數'
                            ,[TKSTOCKSREORDS].[STOCKIDKEY] 
                            ,[TKSTOCKSREORDS].[STOCKACCOUNTNUMBER] 
                            ,[TKSTOCKSREORDS].[STOCKNAME] 
                            ,[CREATEDATES]
                            ,[TKSTOCKSNAMES].[ID]
                            ,[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] AS '戶號'
                            ,[TKSTOCKSNAMES].[STOCKNAME] AS '股東姓名'
                            ,[TKSTOCKSNAMES].[IDNUMBER] AS '身份證字號或統一編號'
                            ,[TKSTOCKSNAMES].[POSTALCODE] AS '通訊地郵遞區號'
                            ,[TKSTOCKSNAMES].[MAILINGADDRESS] AS '通訊地址'
                            ,[TKSTOCKSNAMES].[REGISTEREDPOSTALCODE] AS '戶籍地郵遞區號'
                            ,[TKSTOCKSNAMES].[REGISTEREDADDRESS] AS '戶籍/設立地址'
                            ,[TKSTOCKSNAMES].[DATEOFBIRTH] AS '出生/設立日期'
                            ,[TKSTOCKSNAMES].[BANKNAME] AS '銀行名稱'
                            ,[TKSTOCKSNAMES].[BRANCHNAME] AS '分行名稱'
                            ,[TKSTOCKSNAMES].[BANKCODE] AS '銀行代碼'
                            ,[TKSTOCKSNAMES].[ACCOUNTNUMBER] AS '帳號'
                            ,[TKSTOCKSNAMES].[HOMEPHONENUMBER] AS '住家電話'
                            ,[TKSTOCKSNAMES].[MOBILEPHONENUMBER] AS '手機號碼'
                            ,[TKSTOCKSNAMES].[EMAIL] AS 'e-mail'
                            ,[TKSTOCKSNAMES].[PASSPORTNUMBER] AS '護照號碼'
                            ,[TKSTOCKSNAMES].[ENGLISHNAME] AS '英文名'
                            ,[TKSTOCKSNAMES].[FATHER] AS '父'
                            ,[TKSTOCKSNAMES].[MOTHER] AS '母'
                            ,[TKSTOCKSNAMES].[SPOUSE] AS '配偶'
                            ,[TKSTOCKSNAMES].[COMMENTS] AS '備註'
                            FROM [TKACT].[dbo].[TKSTOCKSREORDS]
                            LEFT JOIN [TKACT].[dbo].[TKSTOCKSNAMES] ON [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]=[TKSTOCKSREORDS].[STOCKACCOUNTNUMBER]
                            WHERE 1=1
                            AND  [TKSTOCKSREORDS].[VALID]='Y'
                            {0}
                            ORDER BY [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER],CONVERT(INT,[TKSTOCKSREORDS].[STOCKSHARES] ) DESC,[TKSTOCKSREORDS].[STOCKID]
                            ", SBQUERY1.ToString());

            return SB;

        }

        public StringBuilder SETSQL2(string STOCKACCOUNTNUMBER)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                SBQUERY1.AppendFormat(@" AND ([TKSTOCKSNAMES].STOCKACCOUNTNUMBER LIKE '%{0}%'  OR [TKSTOCKSNAMES].STOCKNAME LIKE '%{0}%')", STOCKACCOUNTNUMBER);
            }
            else
            {
                SBQUERY1.AppendFormat(@"  ");
            }

            SB.AppendFormat(@" 
                            SELECT 
                            [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] AS '戶號'
                            ,[TKSTOCKSNAMES].[STOCKNAME] AS '股東姓名'
                            ,[TKSTOCKSNAMES].[IDNUMBER] AS '身份證字號或統一編號'
                            ,(SELECT ISNULL(SUM(CONVERT(INT,[STOCKSHARES])),0) FROM  [TKACT].[dbo].[TKSTOCKSREORDS] WHERE   [TKSTOCKSREORDS].[VALID]='Y' AND [TKSTOCKSREORDS].[STOCKACCOUNTNUMBER]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]) AS '股數'
                            ,'10' AS '每股面額(元)'
                            ,(SELECT ISNULL(SUM(CONVERT(INT,[STOCKSHARES])),0)*10 FROM  [TKACT].[dbo].[TKSTOCKSREORDS] WHERE [TKSTOCKSREORDS].[VALID]='Y' AND [TKSTOCKSREORDS].[STOCKACCOUNTNUMBER]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]) AS '股款'
                            FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                            WHERE 1=1
                            AND (SELECT ISNULL(SUM(CONVERT(INT,[STOCKSHARES])),0) FROM  [TKACT].[dbo].[TKSTOCKSREORDS] WHERE [TKSTOCKSREORDS].[VALID]='Y' AND [TKSTOCKSREORDS].[STOCKACCOUNTNUMBER]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER])>0
                            {0}
                            ORDER BY [STOCKACCOUNTNUMBER]
                            ", SBQUERY1.ToString());

            return SB;

        }

        public StringBuilder SETSQL3(string STOCKACCOUNTNUMBER)
        { 
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                SBQUERY1.AppendFormat(@" AND (戶號 LIKE '%{0}%'  OR 股東姓名 LIKE '%{0}%')", STOCKACCOUNTNUMBER);
            }
            else
            {
                SBQUERY1.AppendFormat(@"  ");
            }

            SB.AppendFormat(@"                      
                             SELECT *
                            ,((平均買入每股成本-結餘每股成本)/平均買入每股成本) AS '每股平均投資損益%'
                            FROM 
                            (
                            SELECT *
                            ,(增資股數 + 轉入股數) AS '買進/轉讓股數'
                            ,(增資股數金額 + 轉入股數金額) AS '買進總成本'
                            ,CONVERT(DECIMAL(16, 5), (增資股數金額 + 轉入股數金額) / (增資股數 + 轉入股數)) AS '平均買入每股成本'
                            ,轉出股數 AS '賣出/轉讓股數'
                            ,(轉出股數金額-轉出股數金額的證券交易稅額) AS '賣出淨額'
                            ,CONVERT(DECIMAL(16, 5), (((增資股數金額 + 轉入股數金額) - 歷年分配現金股利淨配 - 轉出股數金額) / 結餘股數)) AS '結餘每股成本'
                            FROM (
                            SELECT [STOCKACCOUNTNUMBER] AS '戶號'
                            ,[STOCKNAME] AS '股東姓名 '
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES])), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANSADD] WHERE  REASONFORCHANGE NOT IN ('資本公積分配','盈餘分配')  AND [TKSTOCKSTRANSADD].[STOCKACCOUNTNUMBER] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '增資股數'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES])), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANSADD] WHERE  REASONFORCHANGE IN ('資本公積分配','盈餘分配') AND [TKSTOCKSTRANSADD].[STOCKACCOUNTNUMBER] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] )AS '歷年分配股數'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES])), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANS] WHERE [STOCKACCOUNTNUMBERTO] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '轉入股數'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES])), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANS] WHERE [STOCKACCOUNTNUMBERFORM] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]) AS '轉出股數'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES])), 0) FROM [TKACT].[dbo].[TKSTOCKSREORDS] WHERE [TKSTOCKSREORDS].[VALID]='Y' AND [TKSTOCKSREORDS].[STOCKACCOUNTNUMBER] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '結餘股數'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES]) * [TRADINGPRICEPERSHARE]), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANSADD] WHERE REASONFORCHANGE NOT IN ('資本公積分配','盈餘分配') AND [TKSTOCKSTRANSADD].[STOCKACCOUNTNUMBER] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '增資股數金額'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES]) * [TRADINGPRICEPERSHARE]), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANS]  WHERE [STOCKACCOUNTNUMBERTO] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '轉入股數金額'
                            ,(SELECT ISNULL(SUM(CONVERT(INT, [STOCKSHARES]) * [TRADINGPRICEPERSHARE]), 0) FROM [TKACT].[dbo].[TKSTOCKSTRANS]  WHERE [STOCKACCOUNTNUMBERFORM] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '轉出股數金額'

                            ,(SELECT ISNULL(SUM([ACTUALCASHDIVIDENDPAID]), 0) FROM [TKACT].[dbo].[TKSTOCKSDIV] WHERE [TKSTOCKSDIV].[STOCKACCOUNTNUMBER] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '歷年分配現金股利淨配'
                            ,(SELECT CONVERT(INT, ISNULL(SUM(CONVERT(INT, [STOCKSHARES]) * [TRADINGPRICEPERSHARE]), 0) * 0.003) FROM [TKACT].[dbo].[TKSTOCKSTRANS] WHERE  [STOCKACCOUNTNUMBERFORM] = [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] ) AS '轉出股數金額的證券交易稅額'
                            FROM [TKACT].[dbo].[TKSTOCKSNAMES]
                            ) AS TEMP
                            WHERE 結餘股數 > 0
                            ) AS TEMP2
                            WHERE 1=1
                            {0}
                            ORDER BY 戶號

                            ", SBQUERY1.ToString());

            return SB;

        }

        public void SETFASTREPORT_TKSTOCKSNAMES()
        {
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();

            report1.Load(@"REPORT\股東明細表.frx");

            SQL = SETSQL_TKSTOCKSNAMES();


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl2;
            report1.Show();
        }

        public StringBuilder SETSQL_TKSTOCKSNAMES()
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();

            SB.AppendFormat(@" 
                            
                            SELECT 
                            [CREATEDATES]
                            ,[TKSTOCKSNAMES].[ID]
                            ,[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER] AS '戶號'
                            ,[TKSTOCKSNAMES].[STOCKNAME] AS '股東姓名'
                            ,[TKSTOCKSNAMES].[IDNUMBER] AS '身份證字號或統一編號'
                            ,[TKSTOCKSNAMES].[POSTALCODE] AS '通訊地郵遞區號'
                            ,[TKSTOCKSNAMES].[MAILINGADDRESS] AS '通訊地址'
                            ,[TKSTOCKSNAMES].[REGISTEREDPOSTALCODE] AS '戶籍地郵遞區號'
                            ,[TKSTOCKSNAMES].[REGISTEREDADDRESS] AS '戶籍/設立地址'
                            ,'民國 '+CONVERT(NVARCHAR,DATEPART(YEAR,(CONVERT(DATETIME,[TKSTOCKSNAMES].[DATEOFBIRTH])))-1911)+'年'+CONVERT(NVARCHAR,DATEPART(MONTH,(CONVERT(DATETIME,[TKSTOCKSNAMES].[DATEOFBIRTH]))))+'月'+CONVERT(NVARCHAR,DATEPART(DAY,(CONVERT(DATETIME,[TKSTOCKSNAMES].[DATEOFBIRTH]))))+'日' AS '出生/設立日期'
                            ,[TKSTOCKSNAMES].[BANKNAME] AS '銀行名稱'
                            ,[TKSTOCKSNAMES].[BRANCHNAME] AS '分行名稱'
                            ,[TKSTOCKSNAMES].[BANKCODE] AS '銀行代碼'
                            ,[TKSTOCKSNAMES].[ACCOUNTNUMBER] AS '帳號'
                            ,[TKSTOCKSNAMES].[HOMEPHONENUMBER] AS '住家電話'
                            ,[TKSTOCKSNAMES].[MOBILEPHONENUMBER] AS '手機號碼'
                            ,[TKSTOCKSNAMES].[EMAIL] AS 'e-mail'
                            ,[TKSTOCKSNAMES].[PASSPORTNUMBER] AS '護照號碼'
                            ,[TKSTOCKSNAMES].[ENGLISHNAME] AS '英文名'
                            ,[TKSTOCKSNAMES].[FATHER] AS '父'
                            ,[TKSTOCKSNAMES].[MOTHER] AS '母'
                            ,[TKSTOCKSNAMES].[SPOUSE] AS '配偶'
                            ,[TKSTOCKSNAMES].[COMMENTS] AS '備註'
                            FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                            ORDER BY [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]
                            
                            ");

            return SB;

        }


        public void SETFASTREPORT_TKSTOCKSTRANSADD(string SDATE, string EDATE, string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();

            report1.Load(@"REPORT\股權增資.frx");

            SQL = SETSQL_TKSTOCKSTRANSADD(SDATE, EDATE, STOCKACCOUNTNUMBER, STOCKNAME);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl3;
            report1.Show();
        }

        public StringBuilder SETSQL_TKSTOCKSTRANSADD(string SDATE,string EDATE,string STOCKACCOUNTNUMBER,string STOCKNAME)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();
            StringBuilder SBQUERY3 = new StringBuilder();

            if(!string.IsNullOrEmpty(SDATE)&& !string.IsNullOrEmpty(EDATE))
            {
                SBQUERY1.AppendFormat(@"
                                       AND [CAPITALINCREASERECORDDATE]>='{0}' AND [CAPITALINCREASERECORDDATE]<='{1}'
                                        ", SDATE, EDATE);
            }
            else
            {
                SBQUERY1.AppendFormat(@" ");
            }
            if(!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                SBQUERY2.AppendFormat(@" 
                                        AND STOCKACCOUNTNUMBER LIKE '%{0}%'
                                        ", STOCKACCOUNTNUMBER);
            }
            else
            {
                SBQUERY2.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                SBQUERY3.AppendFormat(@" 
                                        AND STOCKNAME LIKE '%{0}%'
                                        ", STOCKNAME);
            }
            else
            {
                SBQUERY3.AppendFormat(@" ");
            }


            SB.AppendFormat(@"                      
                            SELECT 
                             [SERNO]
                            ,[ID]
                            ,[CAPITALINCREASERECORDDATE] AS '增資基準日'
                            ,[REASONFORCHANGE] AS '異動原因'
                            ,[STOCKACCOUNTNUMBER] AS '戶號'
                            ,[STOCKNAME] AS '股東姓名'
                            ,CONVERT(INT,[INCREASEDSHARES]) AS '增資股數'
                            ,[PARVALUPERSHARE] AS '每股面額'
                            ,CONVERT(INT,[TRADINGPRICEPERSHARE]) AS '每股成交價格'
                            ,CONVERT(INT,[TOTALTRADINGAMOUNT]) AS '成交總額'
                            ,[INCREASEDSHARESHUNDREDTHOUSANDS] AS '增資股票號碼(十萬股)'
                            ,[INCREASEDSHARESTENSOFTHOUSANDS] AS '增資股票號碼(萬股)'
                            ,[INCREASEDSHARESTHOUSANDS] AS '增資股票號碼(千股)'
                            ,[INCREASEDSHARESIRREGULARLOTS] AS '增資股票號碼(不定額股)'
                            ,CONVERT(INT,[HOLDINGSHARES]) AS '持有股數'
                            ,[PARVALUPER] 
                            ,[STOCKSHARES]
                            FROM [TKACT].[dbo].[TKSTOCKSTRANSADD]
                            WHERE 1=1
                            {0}
                            {1}
                            {2} 
                            ORDER BY  [CAPITALINCREASERECORDDATE]
                            
                            ", SBQUERY1.ToString(), SBQUERY2.ToString(), SBQUERY3.ToString());

            return SB;

        }

        public void SETFASTREPORT_TKSTOCKSTRANS(string SDATE, string EDATE, string STOCKACCOUNTNUMBERFORM, string STOCKNAMEFORM,string STOCKACCOUNTNUMBERTO,string STOCKNAMETO)
        {
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();

            report1.Load(@"REPORT\股權轉讓.frx");

            SQL = SETSQL_TKSTOCKSTRANS(SDATE, EDATE, STOCKACCOUNTNUMBERFORM, STOCKNAMEFORM, STOCKACCOUNTNUMBERTO, STOCKNAMETO);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl4;
            report1.Show();
        }

        public StringBuilder SETSQL_TKSTOCKSTRANS(string SDATE, string EDATE, string STOCKACCOUNTNUMBERFORM, string STOCKNAMEFORM,string STOCKACCOUNTNUMBERTO, string STOCKNAMETO)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();
            StringBuilder SBQUERY2 = new StringBuilder();
            StringBuilder SBQUERY3 = new StringBuilder();
            StringBuilder SBQUERY4 = new StringBuilder();
            StringBuilder SBQUERY5 = new StringBuilder();

            if (!string.IsNullOrEmpty(SDATE) && !string.IsNullOrEmpty(EDATE))
            {
                SBQUERY1.AppendFormat(@"
                                       AND [DATEOFCHANGE]>='{0}' AND [DATEOFCHANGE]<='{1}'
                                        ", SDATE, EDATE);
            }
            else
            {
                SBQUERY1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBERFORM))
            {
                SBQUERY2.AppendFormat(@" 
                                        AND STOCKACCOUNTNUMBERFORM LIKE '%{0}%'
                                        ", STOCKACCOUNTNUMBERFORM);
            }
            else
            {
                SBQUERY2.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAMEFORM))
            {
                SBQUERY3.AppendFormat(@" 
                                        AND STOCKNAMEFORM LIKE '%{0}%'
                                        ", STOCKNAMEFORM);
            }
            else
            {
                SBQUERY3.AppendFormat(@" ");
            }

            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBERTO))
            {
                SBQUERY4.AppendFormat(@" 
                                        AND STOCKACCOUNTNUMBERTO LIKE '%{0}%'
                                        ", STOCKACCOUNTNUMBERTO);
            }
            else
            {
                SBQUERY4.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAMETO))
            {
                SBQUERY5.AppendFormat(@" 
                                        AND STOCKNAMETO LIKE '%{0}%'
                                        ", STOCKNAMETO);
            }
            else
            {
                SBQUERY5.AppendFormat(@" ");
            }

            SB.AppendFormat(@"   
                            SELECT
                            [SERNO]
                            ,[IDFORM]
                            ,[IDTO]
                            ,[DATEOFCHANGE] AS '異動日期'
                            ,[REASOFORCHANGE] AS '異動原因'
                            ,[STOCKACCOUNTNUMBERFORM] AS '轉讓人戶號'
                            ,[STOCKNAMEFORM] AS '轉讓人股東姓名'
                            ,[STOCKACCOUNTNUMBERTO] AS '受讓人戶號'
                            ,[STOCKNAMETO] AS '受讓人股東姓名'
                            ,CONVERT(INT,[TKSTOCKSTRANS].[STOCKSHARES]) AS '轉讓股數'
                            ,[PARVALUEPERSHARE] AS '每股面額'
                            ,CONVERT(DECIMAL(16,2),[TRADINGPRICEPERSHARE]) AS '每股成交價格'
                            ,(CONVERT(INT,[TKSTOCKSTRANS].[STOCKSHARES]))*CONVERT(DECIMAL(16,2),[TRADINGPRICEPERSHARE]) AS '成交總額'
                            ,(CONVERT(DECIMAL(16,2),(CONVERT(INT,[TKSTOCKSTRANS].[STOCKSHARES]))*CONVERT(DECIMAL(16,2),[TRADINGPRICEPERSHARE]))*0.003) AS '證券交易稅額'
                            ,[TRANSFERREDSHARESHUNDREDTHOUSANDS] AS '轉讓股票號碼(十萬股)'
                            ,[TRANSFERREDSHARESTENSOFTHOUSANDS] AS '轉讓股票號碼(萬股)'
                            ,[TRANSFERREDSHARESTHOUSANDS] AS '轉讓股票號碼(千股)'
                            ,[TRANSFERREDSHARESIRREGULARLOTS] AS '轉讓股票號碼(不定額股)'
                            ,[HOLDINGSHARES] AS '持有股數'
                            FROM [TKACT].[dbo].[TKSTOCKSTRANS]
                            LEFT JOIN [TKACT].[dbo].[TKSTOCKSREORDS] ON ([TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESHUNDREDTHOUSANDS] OR [TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESTENSOFTHOUSANDS] OR [TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESTHOUSANDS] OR  [TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESIRREGULARLOTS] )
                            WHERE 1=1
                            {0}
                            {1}
                            {2}
                            {3}
                            {4}
                            ORDER BY [DATEOFCHANGE] 
                            ", SBQUERY1.ToString(), SBQUERY2.ToString(), SBQUERY3.ToString(), SBQUERY4.ToString(), SBQUERY5.ToString());

            return SB;

        }

        public void SETFASTREPORT_TKSTOCKSDIV(string SDATE, string EDATE, string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();

            report1.Load(@"REPORT\除權息.frx"); 

            SQL = SETSQL_TKSTOCKSDIV(SDATE, EDATE, STOCKACCOUNTNUMBER, STOCKNAME);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl5;
            report1.Show();
        }

        public StringBuilder SETSQL_TKSTOCKSDIV(string SDATE, string EDATE, string STOCKACCOUNTNUMBER, string STOCKNAME)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder(); 
            StringBuilder SBQUERY2 = new StringBuilder();
            StringBuilder SBQUERY3 = new StringBuilder();
            StringBuilder SBQUERY4 = new StringBuilder();
            StringBuilder SBQUERY5 = new StringBuilder();

            if (!string.IsNullOrEmpty(SDATE) && !string.IsNullOrEmpty(EDATE))
            {
                SBQUERY1.AppendFormat(@" 
                                       AND [EXDIVIDENDINTERESTRECORDDATE]>='{0}' AND [EXDIVIDENDINTERESTRECORDDATE]<='{1}'
                                        ", SDATE, EDATE);
            }
            else
            {
                SBQUERY1.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKACCOUNTNUMBER))
            {
                SBQUERY2.AppendFormat(@" 
                                        AND STOCKACCOUNTNUMBER LIKE '%{0}%'
                                        ", STOCKACCOUNTNUMBER);
            }
            else
            {
                SBQUERY2.AppendFormat(@" ");
            }
            if (!string.IsNullOrEmpty(STOCKNAME))
            {
                SBQUERY3.AppendFormat(@" 
                                        AND STOCKNAME LIKE '%{0}%'
                                        ", STOCKNAME);
            }
            else
            {
                SBQUERY3.AppendFormat(@" ");
            }

            SB.AppendFormat(@"  
                            SELECT *
                            ,(增資+轉入-轉出) AS '除息基準日持有股數'
                            FROM (
                            SELECT 
                            [SERNO]
                            ,[ID]
                            ,[STOCKACCOUNTNUMBER] AS '戶號'
                            ,[STOCKNAME] AS '股東姓名'
                            ,[EXDIVIDENDINTERESTRECORDDATE] AS '除權/息基準日'
                            ,[CASHDIVIDENDPAYMENTDATE] AS '現金股利發放日'
                            ,CONVERT(DECIMAL(16,5),[DIVAMOUNTS]) AS '每股配發資本公積'
                            ,CONVERT(DECIMAL(16,5),[CASHDIVIDENDPERSHARE]) AS '每股配發現金股利'
                            ,CONVERT(DECIMAL(16,5),[STOCKDIVIDEND]) AS '每股配發股票股利'
                            ,CONVERT(DECIMAL(16,5),[DECLAREDCASHDIVIDEND]) AS '應發股利現金股利'
                            ,CONVERT(DECIMAL(16,5),[DECLAREDSTOCKDIVIDEND]) AS '應發股利股票股利'
                            ,CONVERT(INT,[SUPPLEMENTARYPREMIUMTOBEDEDUCTED]) AS '應扣補充保費'
                            ,CONVERT(INT,[ACTUALCASHDIVIDENDPAID]) AS '實發現金股利'
                            ,CONVERT(INT,[CAPITALIZATIONOFSURPLUSDISTRIBUTIONSHARES]) AS '盈餘增資配股數'
                            ,CONVERT(INT,[CAPITALIZATIONOFCAPITALSURPLUSSHARES]) AS '資本公積增資配股數'
                            ,(SELECT ISNULL(SUM(CONVERT(INT,[TKSTOCKSTRANSADD].[STOCKSHARES])),0) FROM [TKACT].[dbo].[TKSTOCKSTRANSADD] WHERE [TKSTOCKSTRANSADD].[STOCKACCOUNTNUMBER]=[TKSTOCKSDIV].[STOCKACCOUNTNUMBER] AND [TKSTOCKSTRANSADD].[CAPITALINCREASERECORDDATE]<=[TKSTOCKSDIV].[EXDIVIDENDINTERESTRECORDDATE]) AS '增資'
                            ,(SELECT ISNULL(SUM(CONVERT(INT,[TKSTOCKSTRANS].[STOCKSHARES])),0) FROM  [TKACT].[dbo].[TKSTOCKSTRANS],[TKACT].[dbo].[TKSTOCKSREORDS] WHERE ([TKSTOCKSTRANS].[TRANSFERREDSHARESHUNDREDTHOUSANDS]=[TKSTOCKSREORDS].STOCKID OR [TKSTOCKSTRANS].[TRANSFERREDSHARESTENSOFTHOUSANDS]=[TKSTOCKSREORDS].STOCKID OR [TKSTOCKSTRANS].[TRANSFERREDSHARESTHOUSANDS]=[TKSTOCKSREORDS].STOCKID OR [TKSTOCKSTRANS].[TRANSFERREDSHARESIRREGULARLOTS]=[TKSTOCKSREORDS].STOCKID ) AND [STOCKACCOUNTNUMBERTO]=[TKSTOCKSDIV].[STOCKACCOUNTNUMBER] AND [TKSTOCKSTRANS].[DATEOFCHANGE]<=[TKSTOCKSDIV].[EXDIVIDENDINTERESTRECORDDATE]) AS '轉入'
                            ,(SELECT ISNULL(SUM(CONVERT(INT,[TKSTOCKSTRANS].[STOCKSHARES])),0) FROM  [TKACT].[dbo].[TKSTOCKSTRANS],[TKACT].[dbo].[TKSTOCKSREORDS] WHERE ([TKSTOCKSTRANS].[TRANSFERREDSHARESHUNDREDTHOUSANDS]=[TKSTOCKSREORDS].STOCKID OR [TKSTOCKSTRANS].[TRANSFERREDSHARESTENSOFTHOUSANDS]=[TKSTOCKSREORDS].STOCKID OR [TKSTOCKSTRANS].[TRANSFERREDSHARESTHOUSANDS]=[TKSTOCKSREORDS].STOCKID OR [TKSTOCKSTRANS].[TRANSFERREDSHARESIRREGULARLOTS]=[TKSTOCKSREORDS].STOCKID ) AND [STOCKACCOUNTNUMBERFORM]=[TKSTOCKSDIV].[STOCKACCOUNTNUMBER] AND [TKSTOCKSTRANS].[DATEOFCHANGE]<=[TKSTOCKSDIV].[EXDIVIDENDINTERESTRECORDDATE]) AS '轉出'
                            FROM [TKACT].[dbo].[TKSTOCKSDIV]
                            WHERE 1=1  
                            {0}
                            {1}
                            {2} 
 
                            ) AS TEMP
                            ORDER BY '除權/息基準日'
                            ", SBQUERY1.ToString(), SBQUERY2.ToString(), SBQUERY3.ToString());

            return SB;

        }


        public void SETFASTREPORT_TKSTOCKSREORDSDIV(string OLDSTOCKID)
        {
            StringBuilder SQL = new StringBuilder();

            Report report1 = new Report();

            report1.Load(@"REPORT\分割股權.frx");

            SQL = SETSQL_TKSTOCKSREORDSDIV(OLDSTOCKID);


            //20210902密
            Class1 TKID = new Class1();//用new 建立類別實體
            SqlConnectionStringBuilder sqlsb = new SqlConnectionStringBuilder(ConfigurationManager.ConnectionStrings["dbconn"].ConnectionString);

            //資料庫使用者密碼解密
            sqlsb.Password = TKID.Decryption(sqlsb.Password);
            sqlsb.UserID = TKID.Decryption(sqlsb.UserID);

            String connectionString;
            sqlConn = new SqlConnection(sqlsb.ConnectionString);

            report1.Dictionary.Connections[0].ConnectionString = sqlsb.ConnectionString;
            TableDataSource table = report1.GetDataSource("Table") as TableDataSource;
            table.SelectCommand = SQL.ToString();

            report1.Preview = previewControl6;
            report1.Show();
        }

        public StringBuilder SETSQL_TKSTOCKSREORDSDIV(string OLDSTOCKID)
        {
            StringBuilder SB = new StringBuilder();
            StringBuilder SBQUERY1 = new StringBuilder();
       

            if (!string.IsNullOrEmpty(OLDSTOCKID) )
            {
                SBQUERY1.AppendFormat(@"
                                       AND OLDSTOCKID LIKE '%{0}%' 
                                        ", OLDSTOCKID);
            }
            else
            {
                SBQUERY1.AppendFormat(@" ");
            }
         
            SB.AppendFormat(@"     
                            SELECT  
                            [NEWSTOCKID] AS '分割新股票號碼'
                            ,[NEWPARVALUPER] AS '分割新每股面額'
                            ,[NEWSTOCKSHARES] AS '分割新股數'
                            ,[OLDSTOCKID] AS '待分割的股票號碼'
                            ,[OLDPARVALUPER] AS '待分割的每股面額'
                            ,[OLDSTOCKSHARES] AS '待分割的股數'
                            ,[STOCKACCOUNTNUMBER] AS '戶號'
                            ,[STOCKNAME] AS '股東姓名'
                            ,CASE WHEN [VALIDS]='N' THEN '未完分割' ELSE '已分割' END AS '狀態'
                            ,[STOCKIDKEY] 
                            ,[VALIDS]
                            FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                            WHERE 1=1
                            {0}
                            ORDER BY [OLDSTOCKID],[NEWSTOCKID]

                            ", SBQUERY1.ToString());

            return SB;

        }


        public void TKSTOCKSREORDSDIV_ADD(
                                            string NEWSTOCKID
                                            , string NEWPARVALUPER
                                            , string NEWSTOCKSHARES
                                            , string OLDSTOCKID
                                            , string OLDPARVALUPER
                                            , string OLDSTOCKSHARES
                                            , string STOCKIDKEY
                                            , string STOCKACCOUNTNUMBER
                                            , string STOCKNAME
                                            , string VALIDS
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
                                    INSERT INTO  [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                                    (
                                    [NEWSTOCKID]
                                    ,[NEWPARVALUPER]
                                    ,[NEWSTOCKSHARES]
                                    ,[OLDSTOCKID]
                                    ,[OLDPARVALUPER]
                                    ,[OLDSTOCKSHARES]
                                    ,[STOCKIDKEY]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[VALIDS])
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
                                    )


                                     ", NEWSTOCKID
                                            , NEWPARVALUPER
                                            , NEWSTOCKSHARES
                                            , OLDSTOCKID
                                            , OLDPARVALUPER
                                            , OLDSTOCKSHARES
                                            , STOCKIDKEY
                                            , STOCKACCOUNTNUMBER
                                            , STOCKNAME
                                            , VALIDS);


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

        public void TKSTOCKSREORDSDIV_DELETE(string NEWSTOCKID,string OLDSTOCKID)
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
                                    DELETE [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                                    WHERE NEWSTOCKID='{0}'
                                    AND OLDSTOCKID='{1}'
                                    


                                     ", NEWSTOCKID , OLDSTOCKID
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

        public DataTable FINE_TKSTOCKSREORDSDIV(string OLDSTOCKID,string VALIDS)
        {
            DataTable DT = new DataTable();
            SqlConnection sqlConn = new SqlConnection();
            SqlDataAdapter adapter1 = new SqlDataAdapter();
            SqlCommandBuilder sqlCmdBuilder1 = new SqlCommandBuilder();
            DataSet ds1 = new DataSet();
            StringBuilder QUERYS = new StringBuilder();

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


                sbSql.Clear();
                QUERYS.Clear();


                sbSql.AppendFormat(@"  
                                    SELECT ISNULL(SUM(NEWSTOCKSHARES),0) NEWSTOCKSHARES
                                    FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                                    WHERE [OLDSTOCKID]='{0}'
                                    AND [VALIDS]='N'

                                    ", OLDSTOCKID);




                adapter1 = new SqlDataAdapter(@"" + sbSql, sqlConn);

                sqlCmdBuilder1 = new SqlCommandBuilder(adapter1);
                sqlConn.Open();
                ds1.Clear();
                adapter1.Fill(ds1, "TEMPds1");
                sqlConn.Close();


                if (ds1.Tables["TEMPds1"].Rows.Count > 0)
                {
                    return ds1.Tables["TEMPds1"];
                }
                else
                {
                    return null;
                }


            }
            catch
            {
                return null;
            }
            finally
            {
                sqlConn.Close();
            }
        }


        public void TKSTOCKSREORDS_AFTER_DIV(string OLDSTOCKID)
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
                                    INSERT INTO  [TKACT].[dbo].[TKSTOCKSREORDS]
                                    (
                                    [STOCKID]
                                    ,[PARVALUPER]
                                    ,[STOCKSHARES]
                                    ,[STOCKIDKEY]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[VALID]
                                    )
                                    SELECT 
                                    [NEWSTOCKID] 
                                    ,[NEWPARVALUPER] 
                                    ,[NEWSTOCKSHARES]
                                    ,[STOCKIDKEY] 
                                    ,[STOCKACCOUNTNUMBER] 
                                    ,[STOCKNAME] 
                                    ,'Y'
                                    FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                                    WHERE OLDSTOCKID='{0}' AND [VALIDS]='N'

                                    UPDATE [TKACT].[dbo].[TKSTOCKSREORDS]
                                    SET [VALID]='N'
                                    WHERE STOCKID='{0}'

                                    UPDATE [TKACT].[dbo].[TKSTOCKSREORDSDIV]
                                    SET [VALIDS]='Y'
                                    WHERE OLDSTOCKID='{0}'

                                    UPDATE [TKACT].[dbo].[TKSTOCKSREORDS]
                                    SET [VALID]='N'
                                    WHERE [STOCKID] IN (SELECT [OLDSTOCKID] FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV] WHERE [VALIDS]='Y' GROUP BY [OLDSTOCKID])

                                     ", OLDSTOCKID
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

        /// <summary>
        /// 增資+異動的過帳
        /// </summary>
        public void TKSTOCKSDIV_AFTER()
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
                                    --轉讓的[STOCKSHARES]
                                    UPDATE  [TKACT].[dbo].[TKSTOCKSTRANS]
                                    SET [TKSTOCKSTRANS].[STOCKSHARES]=[TKSTOCKSREORDS].[STOCKSHARES]
                                    FROM [TKACT].[dbo].[TKSTOCKSREORDS]
                                    WHERE [TKSTOCKSTRANS].[STOCKSHARES]=0
                                    AND [TKSTOCKSTRANS].[STOCKSHARES]<>[TKSTOCKSREORDS].[STOCKSHARES]
                                    AND ([TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESHUNDREDTHOUSANDS] OR [TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESTENSOFTHOUSANDS]  OR [TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESTHOUSANDS] OR [TKSTOCKSREORDS].STOCKID=[TRANSFERREDSHARESIRREGULARLOTS] )


                                    --增資的過帳
                                    UPDATE [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                    SET ID=[TKSTOCKSNAMES].ID
                                    FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]=[TKSTOCKSTRANSADD].[STOCKACCOUNTNUMBER]
                                    AND [TKSTOCKSTRANSADD].ID<>[TKSTOCKSNAMES].ID


                                    INSERT INTO [TKACT].[dbo].[TKSTOCKSREORDS]
                                    (
                                    [STOCKID]
                                    ,[PARVALUPER]
                                    ,[STOCKSHARES]
                                    ,[STOCKIDKEY]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,[VALID]
                                    )
                                    SELECT ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') 
                                    ,[PARVALUPER]
                                    ,[STOCKSHARES]
                                    ,[ID]
                                    ,[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]
                                    ,'Y'
                                    FROM  [TKACT].[dbo].[TKSTOCKSTRANSADD]
                                    WHERE (ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') )<>''
                                    AND (ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') ) NOT IN (SELECT [STOCKID] FROM [TKACT].[dbo].[TKSTOCKSREORDS])
                                    AND (ISNULL([INCREASEDSHARESHUNDREDTHOUSANDS],'')+ISNULL([INCREASEDSHARESTENSOFTHOUSANDS],'')+ISNULL([INCREASEDSHARESTHOUSANDS],'')+ISNULL([INCREASEDSHARESIRREGULARLOTS],'') ) NOT IN (SELECT [OLDSTOCKID] FROM [TKACT].[dbo].[TKSTOCKSREORDSDIV] WHERE [VALIDS] IN ('Y') GROUP BY [OLDSTOCKID])


                                    --轉讓的過帳
                                    UPDATE [TKACT].[dbo].[TKSTOCKSTRANS]
                                    SET [IDFORM]=[TKSTOCKSNAMES].ID
                                    FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [TKSTOCKSTRANS].[STOCKACCOUNTNUMBERFORM]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]
                                    AND [TKSTOCKSTRANS].[IDFORM]<>[TKSTOCKSNAMES].ID

                                    UPDATE [TKACT].[dbo].[TKSTOCKSTRANS]
                                    SET [IDTO]=[TKSTOCKSNAMES].ID
                                    FROM  [TKACT].[dbo].[TKSTOCKSNAMES]
                                    WHERE [TKSTOCKSTRANS].[STOCKACCOUNTNUMBERTO]=[TKSTOCKSNAMES].[STOCKACCOUNTNUMBER]
                                    AND [TKSTOCKSTRANS].[IDTO]<>[TKSTOCKSNAMES].ID                        
                                 
                                    UPDATE [TKACT].[dbo].[TKSTOCKSREORDS]
                                    SET [STOCKIDKEY]=TEMP2.IDTO
                                    ,[STOCKACCOUNTNUMBER]=TEMP2.[STOCKACCOUNTNUMBER]
                                    ,[STOCKNAME]=TEMP2.[STOCKNAME]
                                    FROM (
	                                    SELECT *
	                                    FROM 
	                                    (
		                                    SELECT
		                                    [STOCKID]
		                                    ,[PARVALUPER]
		                                    ,[STOCKSHARES]
		                                    ,[STOCKIDKEY]
		                                    ,(SELECT TOP 1 [IDTO] FROM  [TKACT].[dbo].[TKSTOCKSTRANS] WHERE ([TRANSFERREDSHARESHUNDREDTHOUSANDS]=[STOCKID]  OR [TRANSFERREDSHARESTENSOFTHOUSANDS]=[STOCKID] OR [TRANSFERREDSHARESTHOUSANDS]=[STOCKID] OR [TRANSFERREDSHARESIRREGULARLOTS]=[STOCKID] ) ORDER BY [SERNO] DESC ) AS 'IDTO'
		                                    FROM [TKACT].[dbo].[TKSTOCKSREORDS]
	                                    ) AS TEMP
	                                    LEFT JOIN [TKACT].[dbo].[TKSTOCKSNAMES] ON [TKSTOCKSNAMES].ID=IDTO
	                                    WHERE ISNULL(TEMP.IDTO,'')<>''
                                    ) AS TEMP2
                                    WHERE TEMP2.STOCKID=[TKSTOCKSREORDS].STOCKID
                                    AND [TKSTOCKSREORDS].[STOCKIDKEY]<>TEMP2.IDTO

                                    "
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

        private void textBox113_TextChanged(object sender, EventArgs e)
        {           
            CHECK_textBox116(textBox113.Text, textBox115.Text);
        }

        private void textBox115_TextChanged(object sender, EventArgs e)
        {
            CHECK_textBox116(textBox113.Text, textBox115.Text);
        }

        public void CHECK_textBox116(string textBox113,string textBox115)
        {            
            decimal value1;
            decimal value2;

            if(!string.IsNullOrEmpty(textBox113)&& !string.IsNullOrEmpty(textBox115) )
            {
                if (decimal.TryParse(textBox113, out value1) && value1 >= 0 && (decimal.TryParse(textBox115, out value2) && value2 >= 0))
                {
                    CAL_textBox116(value1, value2);
                }
                else
                {
                    // 輸入無效或小於等於0
                    MessageBox.Show("請輸入有效的正小數且大於0。");
                }
            }
           
        }
        public void CAL_textBox116(decimal textBox113, decimal textBox115)
        {
            if (textBox113>0 && textBox115>0)
            {
                textBox116.Text = (textBox113 - textBox115).ToString();
            }
        }
        private void dateTimePicker8_ValueChanged(object sender, EventArgs e)
        {
            textBox120.Text = dateTimePicker8.Value.ToString("yyyy");
        }



        #endregion


        #region BUTTON

        private void button1_Click(object sender, EventArgs e)
        {
            Search(textBox1.Text,textBox2.Text); 
        }
        private void button2_Click(object sender, EventArgs e)
        {


            TKSTOCKSNAMES_ADD(DateTime.Now.ToString("yyyy /MM/dd")
                , textBox3.Text.Trim()
                , textBox4.Text.Trim()
                , textBox5.Text.Trim()
                , textBox6.Text.Trim()
                , textBox7.Text.Trim()
                , textBox8.Text.Trim()
                , textBox9.Text.Trim()
                , dateTimePicker1.Value.ToString("yyyy/MM/dd") 
                , textBox10.Text.Trim()
                , textBox11.Text.Trim()
                , textBox12.Text.Trim()
                , textBox13.Text.Trim()
                , textBox15.Text.Trim()
                , textBox14.Text.Trim()
                , textBox16.Text.Trim()
                , textBox17.Text.Trim()
                , textBox18.Text.Trim()
                , textBox19.Text.Trim()
                , textBox20.Text.Trim()
                , textBox21.Text.Trim()
                , textBox70.Text.Trim()
                );

            Search(textBox1.Text.Trim(), textBox2.Text.Trim());
        }


        private void button3_Click(object sender, EventArgs e)
        {
            Search_DG2(textBox22.Text, textBox23.Text);
            Search_DG3(textBox22.Text, textBox23.Text);
        }


        private void button4_Click(object sender, EventArgs e)
        {


            TKSTOCKSCHAGES_ADD(DateTime.Now.ToString("yyyy/MM/dd")
                , textBox26.Text.Trim()
                , textBox27.Text.Trim()
                , textBox28.Text.Trim()
                , textBox29.Text.Trim()
                , textBox30.Text.Trim()
                , textBox31.Text.Trim()
                , textBox32.Text.Trim()
                , dateTimePicker2.Value.ToString("yyyy/MM/dd")
                , textBox25.Text.Trim()
                , textBox33.Text.Trim()
                , textBox34.Text.Trim()
                , textBox35.Text.Trim()
                , textBox24.Text.Trim()
                , textBox36.Text.Trim()
                , textBox37.Text.Trim()
                , textBox38.Text.Trim()
                , textBox39.Text.Trim()
                , textBox40.Text.Trim()
                , textBox41.Text.Trim()
                , textBox42.Text.Trim()
                , "N"
                , textBox43.Text.Trim()
                , textBox71.Text.Trim()
                );

            TKSTOCKSNAMES_UPDATE(
                               textBox43.Text.Trim()
                               , textBox26.Text.Trim()
                               , textBox27.Text.Trim()
                               , textBox28.Text.Trim()
                               , textBox29.Text.Trim()
                               , textBox30.Text.Trim()
                               , textBox31.Text.Trim()
                               , textBox32.Text.Trim()
                               , dateTimePicker2.Value.ToString("yyyy/MM/dd")
                               , textBox25.Text.Trim()
                               , textBox33.Text.Trim()
                               , textBox34.Text.Trim()
                               , textBox35.Text.Trim()
                               , textBox24.Text.Trim()
                               , textBox36.Text.Trim()
                               , textBox37.Text.Trim()
                               , textBox38.Text.Trim()
                               , textBox39.Text.Trim()
                               , textBox40.Text.Trim()
                               , textBox41.Text.Trim()
                               , textBox42.Text.Trim()
                               , textBox71.Text.Trim()
                               );

            Search_DG2(textBox22.Text, textBox23.Text); 
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Search_DG4(textBox44.Text, textBox45.Text);
        }

     

        private void button6_Click(object sender, EventArgs e)
        {
            //ID
            if (!string.IsNullOrEmpty(textBox46.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox46.Text);
                if (DT != null)
                {
                    textBox47.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox57.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox47.Text = "";
                    textBox57.Text = "";
                }
            }

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
            , "0"
            , textBox131.Text
            , textBox132.Text
            , textBox133.Text
            , textBox134.Text
            , textBox57.Text
            , textBox56.Text
            );

            //TKSTOCKSREORDS_ADD();
            TKSTOCKSDIV_AFTER();

            Search_DG4(textBox44.Text, textBox45.Text);
        }


        private void button7_Click(object sender, EventArgs e)
        {
            TKSTOCKSTRANSADD_UPDATE(
            textBox58.Text
         , dateTimePicker3.Value.ToString("yyyy/MM/dd")
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
         , "0"
        
         );


            //TKSTOCKSREORDS_ADD();
            TKSTOCKSDIV_AFTER();

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

                //TKSTOCKSREORDS_ADD();
                TKSTOCKSDIV_AFTER();

                Search_DG4(textBox44.Text, textBox45.Text);
            }
        }
        private void button9_Click(object sender, EventArgs e)
        {
            Search_DG5(textBox72.Text, textBox73.Text);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            //[IDFORM]
            if (!string.IsNullOrEmpty(textBox76.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox76.Text);
                if (DT != null)
                {
                    textBox77.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox74.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox77.Text = "";
                    textBox74.Text = "";
                }
            }
            //[IDTO]
            if (!string.IsNullOrEmpty(textBox78.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox78.Text);
                if (DT != null)
                {
                    textBox79.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox75.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox79.Text = "";
                    textBox75.Text = "";
                }
            }

            TKSTOCKSTRANS_ADD(
             textBox74.Text
             , textBox75.Text
             , dateTimePicker5.Value.ToString("yyyy/MM/dd")
             , comboBox3.SelectedValue.ToString()
             , textBox76.Text
             , textBox77.Text
             , textBox78.Text
             , textBox79.Text
             , textBox80.Text
             , textBox81.Text
             , textBox82.Text
             , textBox83.Text
             , textBox84.Text
             , textBox85.Text
             , textBox86.Text
             , textBox87.Text
             , textBox88.Text
             , "0"
             , textBox59.Text
             , textBox60.Text
             , textBox61.Text
             , textBox62.Text
            
             );

            //TKSTOCKSREORDS_UPDATE();
            TKSTOCKSDIV_AFTER();

            Search_DG5(textBox72.Text, textBox73.Text);
        }


        private void button11_Click(object sender, EventArgs e)
        {
            //[IDFORM]
            if (!string.IsNullOrEmpty(textBox76.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox76.Text);
                if (DT != null)
                {
                    textBox77.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox74.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox77.Text = "";
                    textBox74.Text = "";
                }
            }
            //[IDTO]
            if (!string.IsNullOrEmpty(textBox78.Text))
            {
                DataTable DT = FINE_TKSTOCKSNAMES_STOCKNAME(textBox78.Text);
                if (DT != null)
                {
                    textBox79.Text = DT.Rows[0]["STOCKNAME"].ToString();
                    textBox75.Text = DT.Rows[0]["ID"].ToString();
                }
                else
                {
                    textBox79.Text = "";
                    textBox75.Text = "";
                }
            }

            TKSTOCKSTRANS_UPDATE
                (
                 textBox106.Text.Trim()
                , textBox74.Text.Trim()
                , textBox75.Text.Trim()
                , dateTimePicker5.Value.ToString("yyyy/MM/dd")
                , comboBox3.SelectedValue.ToString()
                , textBox76.Text.Trim()
                , textBox77.Text.Trim()
                , textBox78.Text.Trim()
                , textBox79.Text.Trim()
                , textBox80.Text.Trim()
                , textBox81.Text.Trim()
                , textBox82.Text.Trim()
                , textBox83.Text.Trim()
                , textBox84.Text.Trim()
                , textBox85.Text.Trim()
                , textBox86.Text.Trim()
                , textBox87.Text.Trim()
                , textBox88.Text.Trim()
                , "0"

                );

            //TKSTOCKSREORDS_UPDATE();
            TKSTOCKSDIV_AFTER();

            Search_DG5(textBox72.Text, textBox73.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {           
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                //TKSTOCKSREORDS_UPDATE_BEFROEDELETE(textBox106.Text.Trim());
                TKSTOCKSTRANS_DELETE(textBox106.Text.Trim());

                //TKSTOCKSREORDS_UPDATE();
                TKSTOCKSDIV_AFTER();

                Search_DG5(textBox72.Text, textBox73.Text);

            }
        }
        private void button13_Click(object sender, EventArgs e)
        {
            Search_DG6(textBox107.Text, textBox108.Text);
        }

        private void button14_Click(object sender, EventArgs e)
        {
            TKSTOCKSDIV_ADD(
            textBox109.Text
            , textBox110.Text
            , dateTimePicker7.Value.ToString("yyyy/MM/dd")
            , dateTimePicker8.Value.ToString("yyyy/MM/dd")
            , textBox111.Text
            , textBox112.Text
            , textBox113.Text
            , textBox114.Text
            , textBox115.Text
            , textBox116.Text
            , textBox117.Text
            , textBox118.Text
            , textBox63.Text
            , textBox120.Text
            , textBox119.Text
            );
            Search_DG6(textBox107.Text, textBox108.Text);
        }

        private void button15_Click(object sender, EventArgs e)
        {
            TKSTOCKSDIV_UPDATE(
           textBox129.Text
           , textBox109.Text
           , textBox110.Text
           , dateTimePicker7.Value.ToString("yyyy/MM/dd")
           , dateTimePicker8.Value.ToString("yyyy/MM/dd")
           , textBox111.Text
           , textBox112.Text
           , textBox113.Text
           , textBox114.Text
           , textBox115.Text
           , textBox116.Text
           , textBox117.Text
           , textBox118.Text
           , textBox120.Text
           , textBox119.Text

           );
            Search_DG6(textBox107.Text, textBox108.Text);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                TKSTOCKSDIV_DELETE(textBox129.Text);

                Search_DG6(textBox107.Text, textBox108.Text);

            }

        }




        private void button17_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                if(!string.IsNullOrEmpty(textBox130.Text))
                {
                    TKSTOCKSNAMES_DELETE(textBox130.Text);

                    Search(textBox1.Text, textBox2.Text);
                }
             

            }
        }

        
        private void button18_Click(object sender, EventArgs e)
        {
            SETFASTREPORT(comboBox5.Text.ToString(),textBox121.Text.Trim());           
        }


        private void button19_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_TKSTOCKSNAMES();  
        }

        private void button20_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_TKSTOCKSTRANSADD(dateTimePicker4.Value.ToString("yyyy/MM/dd"), dateTimePicker6.Value.ToString("yyyy/MM/dd"),textBox64.Text, textBox65.Text);
        }
        private void button21_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_TKSTOCKSTRANS(dateTimePicker9.Value.ToString("yyyy/MM/dd"), dateTimePicker10.Value.ToString("yyyy/MM/dd"), textBox66.Text, textBox67.Text, textBox68.Text, textBox69.Text);
        }
        private void button22_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_TKSTOCKSDIV(dateTimePicker11.Value.ToString("yyyy/MM/dd"), dateTimePicker12.Value.ToString("yyyy/MM/dd"), textBox90.Text, textBox91.Text);
        }


        private void button23_Click(object sender, EventArgs e)
        {
            Search_DG7(textBox89.Text.Trim(), textBox92.Text.Trim(), textBox93.Text.Trim());
        }



        private void button24_Click(object sender, EventArgs e)
        {
            TKSTOCKSREORDSDIV_ADD(textBox100.Text.Trim()
                                            , textBox102.Text.Trim()
                                            , textBox101.Text.Trim()
                                            , textBox94.Text.Trim()
                                            , textBox97.Text.Trim()
                                            , textBox98.Text.Trim()
                                            , textBox99.Text.Trim()
                                            , textBox95.Text.Trim()
                                            , textBox96.Text.Trim()
                                            , "N"
                                            );

            Search_DG8(textBox94.Text.Trim(), "N");
        }

        private void button26_Click(object sender, EventArgs e)
        {
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                if (textBox98.Text.Equals(textBox104.Text))
                {
                    //分割後新增、刪除股票
                    TKSTOCKSREORDS_AFTER_DIV(textBox94.Text.Trim());
                    TKSTOCKSDIV_AFTER();

                    Search_DG7(textBox89.Text.Trim(), textBox92.Text.Trim(), textBox93.Text.Trim());
                    Search_DG8(textBox94.Text.Trim(), "N");

                    MessageBox.Show("完成分割");

                }
                else
                {
                    MessageBox.Show("分割前後的總股權不相同，不能執行分割");
                }


            }

        }

        private void button25_Click(object sender, EventArgs e)
        {           
            // 顯示確認對話框
            DialogResult result = MessageBox.Show("確定要執行此操作嗎？", "確認", MessageBoxButtons.OKCancel);

            // 檢查使用者是否按下了確定按鈕
            if (result == DialogResult.OK)
            {
                // 確認後執行的動作
                // TODO: 在這裡執行您的程式碼
                // 例如：
                if (!string.IsNullOrEmpty(textBox103.Text))
                {
                    TKSTOCKSREORDSDIV_DELETE(textBox103.Text.Trim(), textBox94.Text.Trim());
                    Search_DG8(textBox94.Text.Trim(), "N");

                }


            }

           
        }
        private void button28_Click(object sender, EventArgs e)
        {
            SETFASTREPORT_TKSTOCKSREORDSDIV(textBox105.Text.Trim());
        }



        #endregion

      
    }
}
