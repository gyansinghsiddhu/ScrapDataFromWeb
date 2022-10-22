using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System.Data.OleDb;

using System.IO;
using System.Threading;
using System.Security.Cryptography.X509Certificates;
using System.Net.Security;


using Excel = Microsoft.Office.Interop.Excel;
using ADOX;



namespace ScrapDataFromWeb
{
    public partial class Form3 : Form
    {

        public IWebDriver driver;

        DataTable objDT, objDT2;
        DataTable distinctTable;

        public static string SelectedTable = string.Empty;

        DataRow objDR,objDR2;
        bool TT6 = false; bool TT7 = false; bool TT8 = false;

        string PName = "", Price = "", MRP = "", ASIS = "", AGE = "", Bullet_1 = "", Bullet_2 = "", Bullet_3 = "", Bullet_4 = "", Bullet_5 = "", Bullet_6 = "", Bullet_7 = "", Bullet_8 = "", ImgUrl1 = "", ImgUrl2 = "", ImgUrl3 = "", ImgUrl4 = "", ImgUrl5 = "", ImgUrl6 = "", PDESC = "", PDetail = "", URL = "", size = "";
        
        string StrUrl = "";

        string U_R_L = "";
       
        string T1 = "", T2 = "", T3 = "", T4 = "", T5 = "", T6 = "", T7 = "", T8 = "",T9="",T10 ="",T11="";
        int totRec=0;
       
      
       
        public Form3()
        {
            InitializeComponent();
        }

        public void Open_Browser()
        {

            var driverService = ChromeDriverService.CreateDefaultService();
            driverService.HideCommandPromptWindow = true;

            var options = new ChromeOptions();
            options.AddArgument("--headless");
            //options.AddArgument(" — incognito");

           
            driver = new ChromeDriver(driverService);


            try
            {
                driver.Navigate().GoToUrl("https://www.amazon.in/");
               // driver.Navigate().GoToUrl("https://www.amazon.in/dp/" + distinctTable.Rows[0][0].ToString());
            }
            catch
            {
                MessageBox.Show("First Attach ASIN FILE/ CHECK iNTERNET CONNECTION", "ASIN FILE NOT CORRECT/UPLOADED", MessageBoxButtons.OK);
                throw;
            }


        }
        private void btnBrowser_Click(object sender, EventArgs e)
        {
            btnBrowser.Enabled = true;
            Open_Browser();
            btnScrap1.Visible = true;
        }

        public void NewCart2()
        {
            objDT2 = new DataTable("Cart2");
            objDT2.Columns.Add("URL", typeof(string));
            AddProduct2();
        }
        public void AddProduct2()
        {

            objDR2 = objDT2.NewRow();
            objDR2["URL"] = U_R_L; // URL;
            objDT2.Rows.Add(objDR2);
            objDT2.AcceptChanges();
        }
 

        /// <summary>
        /// ASIN	NAME	DESCRIPTION	MRP	SALE PRICE	Image URL1

        /// </summary>
       

        public void NewCart()
        {
            objDT = new DataTable("Cart");

            objDT.Columns.Add("ASIN", typeof(string));
            objDT.Columns.Add("Prouduct Name", typeof(string));
            objDT.Columns.Add("ProductDesc", typeof(string));
            objDT.Columns.Add("MRP", typeof(string));
            objDT.Columns.Add("Price", typeof(string));
            objDT.Columns.Add("Age", typeof(string));

            objDT.Columns.Add("Bullet_1", typeof(string));
            objDT.Columns.Add("Bullet_2", typeof(string));
            objDT.Columns.Add("Bullet_3", typeof(string));
            objDT.Columns.Add("Bullet_4", typeof(string));
            objDT.Columns.Add("Bullet_5", typeof(string));
            objDT.Columns.Add("Bullet_6", typeof(string));
            objDT.Columns.Add("Bullet_7", typeof(string));
            

          
            objDT.Columns.Add("ImgUrl_1", typeof(string));
            objDT.Columns.Add("ImgUrl_2", typeof(string));
            objDT.Columns.Add("ImgUrl_3", typeof(string));
            objDT.Columns.Add("ImgUrl_4", typeof(string));
            objDT.Columns.Add("ImgUrl_5", typeof(string));
            objDT.Columns.Add("ImgUrl_6", typeof(string));
          
           // objDT.Columns.Add("ProductDetails", typeof(string));
            //objDT.Columns.Add("Size", typeof(string));
           // objDT.Columns.Add("Colour", typeof(string));
            AddProduct();

        }

        public void AddProduct()
        {

            objDR = objDT.NewRow();
            objDR["ASIN"] = ASIS.Trim();
            objDR["Prouduct Name"] = PName.Trim();
            objDR["ProductDesc"] = PDESC.Replace("</p>", "");
            objDR["Price"] = Price.Trim();
            objDR["MRP"] = MRP.Trim();
            objDR["Age"] = AGE.Trim();

            objDR["Bullet_1"] = Bullet_1;
            objDR["Bullet_2"] = Bullet_2;
            objDR["Bullet_3"] = Bullet_3;
            objDR["Bullet_4"] = Bullet_4;
            objDR["Bullet_5"] = Bullet_5;
            objDR["Bullet_6"] = Bullet_6;
            objDR["Bullet_7"] = Bullet_7;
           

            objDR["ImgUrl_1"] = ImgUrl1;
            objDR["ImgUrl_2"] = ImgUrl2;
            objDR["ImgUrl_3"] = ImgUrl3;
            objDR["ImgUrl_4"] = ImgUrl4;
            objDR["ImgUrl_5"] = ImgUrl5;
            objDR["ImgUrl_6"] = ImgUrl6;

          
           // objDR["ProductDetails"] = PDetail;
            //objDR["Size"] = size;
            //objDR["Colour"] = Colour;
           
            objDT.Rows.Add(objDR);
            objDT.AcceptChanges();
            

        }




        public void ReadGoogleSheet()
        {
            try
            {
                if (txtColumnAddress.Text == "" || txtSheetAddress.Text == "") { MessageBox.Show("Please Enter Google Sheet Address and Sheet Address", "Enter Google Sheet Path and Column Address"); return; }
                var SheetId = txtSheetAddress.Text.Trim(); // "1t4HqyzqM2ByGKBHhalTKiFh5AafRX6WxJgSh_c4q21M";
                String range = txtColumnAddress.Text.Trim(); // "Sheet24!B2:B";
                var service = AuthorizeGoogleAppForSheetsService();
                SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SheetId, range);
                ValueRange response = request.Execute();
                IList<IList<Object>> values = response.Values;
                if (values != null && values.Count > 0)
                {
                    foreach (var row in values)
                    {
                        try
                        {
                            U_R_L = row[0].ToString();
                            if (objDT2 == null) NewCart2();
                            else AddProduct2();
                            Application.DoEvents();
                        }
                        catch
                        {
                            U_R_L = "https://shopee.co.th/flash_sale";
                            if (objDT2 == null) NewCart2();
                            else AddProduct2();
                            Application.DoEvents();
                        }
                    }
                }
                else
                {

                }
                distinctTable = objDT2;
                //MessageBox.Show(distinctTable.Rows.Count.ToString()+ " SKU LINK IS UPLOADED", "SKU LINK UPLODED");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), "ERROR MSG");
            }
        }

        

        private void ScrapData()
        {
            //StrUrl = txtFileName.Text;
            btnScrap1.Enabled = true;

           

            //if (Convert.ToDateTime(DateTime.Now.ToString("dd-MMM-yyyy")) > Convert.ToDateTime("30-Dec-2022"))
            //{
            //    MessageBox.Show("Ohhh!!! Sorry Software Trail Version Expire. Please Contact Gyan Singh Email- gyansinghsiddhu@gmail.com or whatsapp = +91-9837808476", "Trail Version Expire", MessageBoxButtons.OK, MessageBoxIcon.Information);
            //    this.Close();

            //}





            HtmlAgilityPack.HtmlDocument htmlDocument = new HtmlAgilityPack.HtmlDocument();

            distinctTable = null;
            objDT = null;
            objDT2 = null;
            totRec = 0;
            ReadGoogleSheet();

            // for (int i = 0; i <= distinctTable.Rows.Count; i++)

            for (int i = 0; i <= distinctTable.Rows.Count; i++)
            {
                try
                {
                    //#zg-ordered-list > li:nth-child(2) > span > div > span > a
                    try
                    {
                        driver.Navigate().GoToUrl("https://www.amazon.in/dp/" + distinctTable.Rows[i][0].ToString());

                    }
                    catch
                    {
                        continue;// continue;// break;
                    }


                    try
                    {

                        //1 ASIN NO
                        ASIS = distinctTable.Rows[i][0].ToString();


                        //2 PRODUCT NAME 
                        try
                        {
                            IWebElement t1 = driver.FindElement(By.Id("productTitle")); PName = t1.Text.Replace("\r\n", "");//   .GetAttribute("innerHTML");

                        }
                        catch
                        {
                            PName = "N/A";
                        }

                        //3 PRICE
                        try
                        {

                            IWebElement t2 = driver.FindElement(By.Id("priceblock_ourprice")); Price = t2.Text;  ///.GetAttribute("innerHTML");

                        }
                        catch
                        {
                            try
                            {
                                IWebElement t2 = driver.FindElement(By.Id("priceblock_saleprice")); Price = t2.Text;   ///.GetAttribute("innerHTML");
                            }
                            catch
                            {
                                Price = "N/A";
                            }

                        }

                        //4 MRP
                        try
                        {
                            IWebElement t3 = driver.FindElement(By.XPath("//*[@id='price']/table/tbody/tr[1]/td[2]/span[1]")); MRP = t3.Text; // GetAttribute("innerHTML");
                        }
                        catch
                        {
                            MRP = "N/A";
                        }

                        // AGE GROUP 
                        try
                        {
                            //IWebElement t4 = driver.FindElement(By.XPath("//*[@id='detail_bullets_id']/table/tbody/tr"));
                            IWebElement t4 = driver.FindElement(By.Id("prodDetails"));
                            List<IWebElement> lstTrElem = new List<IWebElement>(t4.FindElements(By.TagName("tr")));
                            //Date first available at Amazon.in:
                            //Customer Reviews
                            //Amazon Bestsellers Rank
                            for (int itm = 0; itm <= lstTrElem.Count; itm++)
                            {
                                IWebElement t_5 = lstTrElem[itm];
                                string abc = t_5.GetAttribute("innerHTML");
                                if (abc.Contains(" Age"))
                                {
                                    AGE = abc.Replace("<td class=\"label\">", "").Replace("</td><td class=\"value\">", " : ").Replace("</td>", "");
                                    string[] words = AGE.Split(':');
                                    AGE = words[1];
                                    break;
                                }


                            }


                        }
                        catch
                        {


                        }





                        // Feature Bullet
                        try
                        {
                            IWebElement t4 = driver.FindElement(By.Id("feature-bullets"));
                            List<IWebElement> lstTrElem = new List<IWebElement>(t4.FindElements(By.TagName("li")));

                            if (lstTrElem.Count > 0)
                            {
                                try
                                {
                                    Bullet_1 = lstTrElem[0].Text;
                                }
                                catch
                                {
                                    Bullet_1 = "";
                                }
                                try
                                {
                                    Bullet_2 = lstTrElem[1].Text;
                                }
                                catch
                                {
                                    Bullet_2 = "";
                                }
                                try
                                {
                                    Bullet_3 = lstTrElem[2].Text;
                                }
                                catch
                                {
                                    Bullet_3 = "";
                                }
                                try
                                {
                                    Bullet_4 = lstTrElem[3].Text;
                                }
                                catch
                                {
                                    Bullet_4 = "";
                                }
                                try
                                {
                                    Bullet_5 = lstTrElem[4].Text;
                                }
                                catch
                                {
                                    Bullet_5 = "";
                                }
                                try
                                {
                                    Bullet_6 = lstTrElem[5].Text;
                                }
                                catch
                                {
                                    Bullet_6 = "";
                                }
                                try
                                {
                                    Bullet_7 = lstTrElem[6].Text;
                                }
                                catch
                                {
                                    Bullet_7 = "";
                                }

                            }




                        }
                        catch
                        {


                        }




                        // product Description 
                        try
                        {

                            IWebElement t9 = driver.FindElement(By.CssSelector("#productDescription"));
                            PDESC = t9.GetAttribute("innerHTML");
                            int firstStringPosition = PDESC.IndexOf("<p>");
                            int secondStringPosition = PDESC.IndexOf("</p>");
                            PDESC = (PDESC.Substring(firstStringPosition, secondStringPosition - firstStringPosition)).Replace("<p>", "").Replace("</p>", "").Replace("\r\n", "").Trim();



                        }
                        catch
                        {
                            try
                            {
                                IWebElement t9 = driver.FindElement(By.Id("productDescription")); string pdesc = t9.GetAttribute("innerHTML").Replace("\r\n", "");
                                PDESC = pdesc.Replace("\t", "").Replace("<p>", "").Replace("</p>", "");

                            }
                            catch
                            {
                                try
                                {
                                    IWebElement t9 = driver.FindElement(By.XPath("//*[@id='productDescription']/p/text()")); PDESC = t9.GetAttribute("innerHTML").Replace("\r\n", "");
                                }
                                catch
                                {
                                    try
                                    {
                                        //*[@id="aplus"]/div/div[2]/div/div/p/text()
                                        IWebElement t9 = driver.FindElement(By.XPath("/html/body/div[2]/div[2]/div[4]/div[12]/div/div/div[2]/p")); PDESC = t9.GetAttribute("innerHTML").Replace("\r\n", "");
                                    }
                                    catch
                                    {
                                        try
                                        {
                                            IWebElement t9 = driver.FindElement(By.ClassName("a-section launchpad-text-left-justify")); PDESC = t9.Text.Replace("\r\n", ""); //.GetAttribute("innerHTML");

                                        }
                                        catch
                                        {
                                            PDESC = "N/A";
                                        }

                                    }

                                }
                            }

                            // T9 = "N/A"; 

                        }




                        String ImgSrtUrl = "https://www.amazon.in/";

                        // Product Images 

                        try
                        {
                            IWebElement t5 = driver.FindElement(By.CssSelector("#landingImage"));
                            t5.Click();
                            System.Threading.Thread.Sleep(3000);





                            IWebElement t_5 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                            T5 = t_5.GetAttribute("src");

                            if (T5.Contains("CB485921387_.gif"))
                            {
                                System.Threading.Thread.Sleep(3000);
                                t_5 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                                T5 = t_5.GetAttribute("src");
                            }
                            else if (T5.Contains("CB485921387_.gif"))
                            {
                                T5 = "";
                            }




                            IWebElement t6 = driver.FindElement(By.XPath("//*[@id='ivImage_1']/div"));
                            t6.Click();
                            System.Threading.Thread.Sleep(2000);

                            IWebElement t_6 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                            T6 = t_6.GetAttribute("src");
                            if (T6.Contains("CB485921387_.gif"))
                            {
                                System.Threading.Thread.Sleep(3000);
                                t_6 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                                T6 = t_6.GetAttribute("src");

                            }
                            else if (T6.Contains("CB485921387_.gif"))
                            {
                                T6 = "";
                            }





                            IWebElement t7 = driver.FindElement(By.XPath("//*[@id='ivImage_2']/div"));
                            t7.Click();
                            System.Threading.Thread.Sleep(2000);
                            IWebElement t_7 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                            T7 = t_7.GetAttribute("src");
                            if (T7.Contains("CB485921387_.gif"))
                            {
                                System.Threading.Thread.Sleep(3000);
                                t_7 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                                T7 = t_7.GetAttribute("src");
                            }
                            else if (T7.Contains("CB485921387_.gif"))
                            {
                                T7 = "";
                            }



                            IWebElement t8 = driver.FindElement(By.XPath("//*[@id='ivImage_3']/div"));
                            t8.Click();
                            System.Threading.Thread.Sleep(2000);
                            IWebElement t_8 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                            T8 = t_8.GetAttribute("src");
                            if (T8.Contains("CB485921387_.gif"))
                            {
                                System.Threading.Thread.Sleep(3000);
                                t_8 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                                T8 = t_8.GetAttribute("src");
                            }
                            else if (T8.Contains("CB485921387_.gif"))
                            {
                                T8 = "";
                            }



                            IWebElement t9 = driver.FindElement(By.XPath("//*[@id='ivImage_4']/div"));
                            t9.Click();
                            System.Threading.Thread.Sleep(2000);
                            IWebElement t_9 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                            T9 = t_9.GetAttribute("src");
                            if (T9.Contains("CB485921387_.gif"))
                            {
                                System.Threading.Thread.Sleep(2000);
                                t_9 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                                T9 = t_9.GetAttribute("src");
                            }
                            else if (T9.Contains("CB485921387_.gif"))
                            {
                                T9 = "";
                            }



                            IWebElement t10 = driver.FindElement(By.XPath("//*[@id='ivImage_5']/div"));
                            t10.Click();
                            System.Threading.Thread.Sleep(2000);
                            IWebElement t_10 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                            T10 = t_10.GetAttribute("src");
                            if (T10.Contains("CB485921387_.gif"))
                            {
                                System.Threading.Thread.Sleep(3000);
                                t_10 = driver.FindElement(By.XPath("//*[@id='ivLargeImage']/img"));
                                T10 = t_10.GetAttribute("src");
                            }
                            else if (T10.Contains("CB485921387_.gif"))
                            {
                                T10 = "";
                            }



                        }
                        catch (Exception ex)
                        {

                        }












                        ImgUrl1 = T5.Trim();
                        ImgUrl2 = T6.Trim();
                        ImgUrl3 = T7.Trim();
                        ImgUrl4 = T8.Trim();
                        ImgUrl5 = T9.Trim();
                        ImgUrl6 = T10.Trim();


                        //Colour = T3.Trim();
                        // PDetail = T10.Trim().Replace("\r\n", "*");
                        //size = T11.Trim();

                        if (objDT == null) NewCart();
                        else AddProduct();

                        totRec++;
                        dgview.DataSource = objDT;
                        label2.Text = totRec.ToString();
                    }
                    catch (Exception ex) { MessageBox.Show("ERROR1: In Get Area COde 1: " + ex); }
                    dgview.Columns[0].Width = 100;// The id column 
                    dgview.Columns[1].Width = 300;
                    dgview.Columns[2].Width = 200;
                    dgview.Columns[3].Width = 70;
                    dgview.Columns[4].Width = 70;
                    dgview.Columns[5].Width = 80;
                    dgview.Columns[6].Width = 80;
                    dgview.Columns[7].Width = 80;
                    dgview.Columns[8].Width = 80;
                    dgview.Columns[9].Width = 80;
                    dgview.Columns[10].Width = 80;
                    this.SelectNextControl(dgview, true, true, false, true);
                    Application.DoEvents();    //this does magic
                    dgview.Focus();
                    //driver.Navigate().Back();

                    TT6 = false;
                    TT7 = false;
                    TT8 = false;
                    T5 = "";
                    T6 = "";
                    T7 = "";
                    T8 = "";
                    T9 = "";
                    T10 = "";

                    ASIS = "";
                    PName = "";
                    PDESC = ""; ;
                    MRP = "";
                    Price = "";
                    AGE = "";

                    Bullet_1 = "";
                    Bullet_2 = "";
                    Bullet_3 = "";
                    Bullet_4 = "";
                    Bullet_5 = "";
                    Bullet_6 = "";
                    Bullet_7 = "";


                    if (btnCtn.Text == "STOP") break;

                }

                catch (Exception)
                {
                    //break;
                }


            }

            try
            {
                driver.Navigate().GoToUrl("https://www.amazon.in/");
                exportGoogleSheet();
                dgview.DataSource = null;
            }
            catch
            {
                MessageBox.Show("First Attach ASIN FILE/ CHECK iNTERNET CONNECTION", "ASIN FILE NOT CORRECT/UPLOADED", MessageBoxButtons.OK);
                throw;
            }
        }


        private void btnScrap1_Click(object sender, EventArgs e)
        {

            ScrapData();
            timer3.Start();
                     
        
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (dgview.DataSource == null)
            {
                MessageBox.Show("Sorry nothing to export into excel sheet..", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            int rowsTotal = 0;
            int colsTotal = 0;
            int I = 0;
            int j = 0;
            int iC = 0;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();

            
            try
            {
                Excel.Workbook excelBook = xlApp.Workbooks.Add();
                System.Threading.Thread.Sleep(10000);
                Excel.Worksheet excelWorksheet = (Excel.Worksheet)excelBook.Worksheets[1];
                xlApp.Visible = true;

                rowsTotal = dgview.RowCount - 1;
                colsTotal = dgview.Columns.Count - 1;
                var _with1 = excelWorksheet;
                _with1.Cells.Select();
                _with1.Cells.Delete();
                for (iC = 0; iC <= colsTotal; iC++)
                {
                    _with1.Cells[1, iC + 1].Value = dgview.Columns[iC].HeaderText;
                }
                for (I = 0; I <= rowsTotal - 1; I++)
                {
                    for (j = 0; j <= colsTotal; j++)
                    {
                        _with1.Cells[I + 2, j + 1].value = dgview.Rows[I].Cells[j].Value;
                    }
                }
                System.Threading.Thread.Sleep(5000);
                _with1.Rows["1:1"].Font.FontStyle = "Bold";
                _with1.Rows["1:1"].Font.Size = 12;

                _with1.Cells.Columns.AutoFit();
                _with1.Cells.Select();
                _with1.Cells.EntireColumn.AutoFit();
                _with1.Cells[1, 1].Select();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                //RELEASE ALLOACTED RESOURCES
                System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
                xlApp = null;
            }
        }

        private void btnCtn_Click(object sender, EventArgs e)
        {
            btnCtn.Text = "STOP";
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            driver.Close();
            Application.Exit();
            //btnStop.Visible = false;
            //pbStatus2.Visible = false;
        }

        private void Form3_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            distinctTable = null;
            OpenFileDialog fdlg = new OpenFileDialog();
            fdlg.Title = "Select file";
            fdlg.InitialDirectory = @"e:\";
            //fdlg.InitialDirectory = @"c:";
            fdlg.FileName = txtFileName.Text;
            fdlg.Filter = "Excel Sheet(*.xls)|*.xls|All Files(*.*)|*.*";
            //fdlg.Filter = "Excel Sheet(*.xlsx)|*.xlsx|All Files(*.*)|*.*";
            fdlg.FilterIndex = 1;
            fdlg.RestoreDirectory = true;
            if (fdlg.ShowDialog() == DialogResult.OK)
            {
                txtFileName.Text = fdlg.FileName;
                Import();
                Application.DoEvents();
            }
        }



        private void Import()
        {
            if (txtFileName.Text.Trim() != string.Empty)
            {
                try
                {
                    string[] strTables = GetTableExcel(txtFileName.Text);

                    SelectTable objSelectTable = new SelectTable(strTables);
                    objSelectTable.ShowDialog(this);
                    objSelectTable.Dispose();
                    if ((SelectedTable != string.Empty) && (SelectedTable != null))
                    {
                        distinctTable = GetDataTableExcel(txtFileName.Text, SelectedTable);
                        MessageBox.Show(distinctTable.Rows.Count+" "+"ASIN NUMBER SUCCESSFULLY ADDED.NOW CLICK ON OPEN BROWSER AFTER THAT CLICK ON SCRP BUTTON", "ASIN UPLOADED", MessageBoxButtons.OK);
                       
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
        }




        public static DataTable GetDataTableExcel(string strFileName, string Table)
        {
            System.Data.OleDb.OleDbConnection conn = new System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OleDb.4.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";");
            conn.Open();
            string strQuery = "SELECT * FROM [" + Table + "]";
            System.Data.OleDb.OleDbDataAdapter adapter = new System.Data.OleDb.OleDbDataAdapter(strQuery, conn);
            System.Data.DataSet ds = new System.Data.DataSet();
            adapter.Fill(ds);
            return ds.Tables[0];
        }

        public static string[] GetTableExcel(string strFileName)
        {
            string[] strTables = new string[100];
            Catalog oCatlog = new Catalog();
            ADOX.Table oTable = new ADOX.Table();
            ADODB.Connection oConn = new ADODB.Connection();
            oConn.Open("Provider=Microsoft.Jet.OleDb.4.0; Data Source = " + strFileName + "; Extended Properties = \"Excel 8.0;HDR=Yes;IMEX=1\";", "", "", 0);
            oCatlog.ActiveConnection = oConn;
            if (oCatlog.Tables.Count > 0)
            {
                int item = 0;
                foreach (ADOX.Table tab in oCatlog.Tables)
                {
                    if (tab.Type == "TABLE")
                    {
                        strTables[item] = tab.Name;
                        item++;
                    }
                }
            }
            return strTables;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Your Hard Drive SrNo is:      " + HardwareInfo.GetHDDSerialNo().ToString() , "This Code Send To Developer For Software Missuse", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            exportGoogleSheet();
        }



        public void exportGoogleSheet()
        {
            if (dgview.DataSource == null)
            {
                MessageBox.Show("Sorry nothing to export into Google sheet..", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }


            var SheetId = txtSheetAddress.Text; // "1t4HqyzqM2ByGKBHhalTKiFh5AafRX6WxJgSh_c4q21M";
            String newRange = txtPriceAddrees.Text.Trim(); // "Sheet24!B2:B";
            var service = AuthorizeGoogleAppForSheetsService();
            SpreadsheetsResource.ValuesResource.GetRequest request = service.Spreadsheets.Values.Get(SheetId, newRange);

            IList<IList<Object>> objNeRecords2 = GenerateData2();

            UpdatGoogleSheetinBatch(objNeRecords2, SheetId, newRange, service);
            //MessageBox.Show("Data has been Successfully Exported in Google Sheet", "Export Google Sheet");
            label2.Text = label2.Text;

        }

        private static SheetsService AuthorizeGoogleAppForSheetsService()
        {
            // If modifying these scopes, delete your previously saved credentials  
            // at ~/.credentials/sheets.googleapis.com-dotnet-quickstart.json  
            string[] Scopes = { SheetsService.Scope.Spreadsheets };
            string ApplicationName = "Google Sheets API .NET Quickstart";
            UserCredential credential;
            using (var stream =
               new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);

            }

            // Create Google Sheets API service.  
            var service = new SheetsService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            return service;
        }



        protected static string GetRange(SheetsService service, string SheetId)
        {
            // Define request parameters.  
            String spreadsheetId = SheetId;
            String range = "A:A";

            SpreadsheetsResource.ValuesResource.GetRequest getRequest =
                       service.Spreadsheets.Values.Get(spreadsheetId, range);
            System.Net.ServicePointManager.ServerCertificateValidationCallback = delegate(object sender, X509Certificate certificate, X509Chain chain, SslPolicyErrors sslPolicyErrors) { return true; };
            ValueRange getResponse = getRequest.Execute();
            IList<IList<Object>> getValues = getResponse.Values;
            if (getValues == null)
            {
                // spreadsheet is empty return Row A Column A  
                return "A:A";
            }

            int currentCount = getValues.Count() + 1;
            String newRange = "A" + currentCount + ":A";
            return newRange;
        }


        public IList<IList<Object>> GenerateData()
        {
            int rowsTotal = 0;
            int colsTotal = 0;

            rowsTotal = dgview.RowCount - 1;
            colsTotal = dgview.Columns.Count - 1;


            List<IList<Object>> objNewRecords = new List<IList<Object>>();


            //IList<Object> obj2 = new List<Object>();
            //obj2.Add("Total SKU: " + label2.Text + " || Total Variation: " + countT.ToString() + " ||  Start Time: " + STRTTIME.ToString() + " ||  End Time: " + ENDTIME.ToString());
            //objNewRecords.Add(obj2);


            IList<Object> obj = new List<Object>();
            for (var x = 0; x <= dgview.ColumnCount - 1; x++)
            {
                obj.Add(dgview.Columns[x].HeaderText.ToString());

            }
            objNewRecords.Add(obj);


            return objNewRecords;
        }


        public IList<IList<Object>> GenerateData2()
        {

            DataTable objDT = (DataTable)this.dgview.DataSource;
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            try
            {
                for (var i = 0; i <= objDT.Rows.Count - 1; i++)
                {
                    IList<Object> obj = new List<Object>();
                    for (var j = 1; j <= objDT.Columns.Count - 1; j++)
                    {
                        try
                        {
                            obj.Add(objDT.Rows[i][j].ToString());
                        }
                        catch { }
                    }

                    objNewRecords.Add(obj);







                }

                return objNewRecords;

            }
            catch
            {
                return objNewRecords;
            }
        }

        private static void UpdatGoogleSheetinBatch(IList<IList<Object>> values, string spreadsheetId, string newRange, SheetsService service)
        {

            SpreadsheetsResource.ValuesResource.UpdateRequest update = service.Spreadsheets.Values.Update(new ValueRange() { Values = values }, spreadsheetId, newRange);
            update.ValueInputOption = SpreadsheetsResource.ValuesResource.UpdateRequest.ValueInputOptionEnum.RAW;
            UpdateValuesResponse result2 = update.Execute();

            try
            {
                //var response = request.Execute();
                var response = update.Execute();
            }
            catch
            { }
        }

        private void timer3_Tick(object sender, EventArgs e)
        {
            timer3.Enabled = false;
           
            int x = Convert.ToInt32(txtDly.Text);
            int y = x * 60000; // 1800000
            var delay = System.Threading.Tasks.Task.Delay(TimeSpan.FromMilliseconds(y));
            var seconds = 0;
            int mn = y / 1000;
            while (!delay.IsCompleted)
            {
                // While we're waiting, note the time ticking past
                seconds++;
                Thread.Sleep(TimeSpan.FromMilliseconds(1000));
                lblWait.Text = "Waiting For Next Scrapping :  " + seconds + " / " + mn;
                Application.DoEvents();

                if (btnCtn.Text == "STOP") break; 

            }
            lblWait.Text = " ";
            ScrapData();
            
            timer3.Enabled = true;
            btnCtn.Text = "Stop";
            Application.DoEvents();

        }



        public static void DeleteTmpFile()
        {
            try
            {
                // Delete the temp file (if it exists)
                string[] files = Directory.GetFiles(System.IO.Path.GetTempPath());
                foreach (string file in files)
                {
                    try { File.Delete(file); }
                    catch { }

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error deleteing TEMP file: " + ex.Message);
            }
        }


        public void Delay(int dly)
        {
            var delay = System.Threading.Tasks.Task.Delay(TimeSpan.FromMilliseconds(dly));
            var seconds = 0;

            while (!delay.IsCompleted)
            {
                // While we're waiting, note the time ticking past
                seconds++;
                Thread.Sleep(TimeSpan.FromMilliseconds(1000));

                Application.DoEvents();

            }

        }



    }
}
