using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Net.Mail;
using FirebirdSql.Data.FirebirdClient;




namespace Demo2m
{
    public class Methods
    {
        private readonly FirefoxDriver _firefox;
        private List<string> messageContent = new List<string>();
        private RowDataList preparationPcsWeb282 = new RowDataList();
        private RowDataList preparationPcsDbList = new RowDataList();

        public Methods(FirefoxDriver firefox)
        {
            _firefox = firefox;
        }


        public string MessageContent(List<string> list)
        {
            var sb = new StringBuilder();
            foreach (var str in list)
            {
                sb.AppendLine(str);
                sb.AppendLine("<br>");

            }
            return sb.ToString();
        }



        public void WaitForAjax()
        {
            while (true) // Handle timeout somewhere
            {
                var ajaxIsComplete = (bool)(_firefox as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete)
                {
                    Thread.Sleep(2000);
                    break;
                }
                Thread.Sleep(2000);
            }
        }

        public void TryToLoadPage(string url, string waitPresenceAllElementsByXPath)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(60));
            for (int i = 0; i < 5; i++)
            {
                try
                {
                    _firefox.Navigate().GoToUrl(url);
                    wait.Until(
                        ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(waitPresenceAllElementsByXPath)));
                    return;
                }
                catch (Exception)
                {
                    Console.WriteLine("Load page 282. Attempt №" + i);
                    i++;
                }

            }

        }

        public void TryToClickWithoutException(string locator, IWebElement element)
        {
            var maxElementRetries = 100;
            var action = new Actions(_firefox);
            var retries = 0;
            while (true)
            {
                try
                {
                    WebDriverWait wait = new WebDriverWait(new SystemClock(), _firefox, TimeSpan.FromSeconds(120),
                        TimeSpan.FromSeconds(5));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                    _firefox.FindElement(By.XPath(locator)).Click();
                    WaitForAjax();
                    return;
                }
                catch (Exception e)
                {
                    if (retries < maxElementRetries)
                    {
                        retries++;
                    }
                    else
                    {
                        throw e;
                    }
                }
            }
        }


        public void WaitForTextInTitleAttribute(string locator, string text)
        {
            const int waitRetryDelayMs = 1000; //шаг итерации (задержка)
            const int timeOut = 500; //время тайм маута 
            bool first = true;

            for (int milliSecond = 0; ; milliSecond += waitRetryDelayMs)
            {
                WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));

                try
                {
                    if (milliSecond > timeOut * 10000)
                    {
                        Debug.WriteLine("Timeout: Text " + text + " is not found ");
                        break; //если время ожидания закончилось (элемент за выделенное время не был найден)
                    }
                    _firefox.Manage().Timeouts().SetPageLoadTimeout(TimeSpan.FromSeconds(60));
                    _firefox.Manage().Timeouts().SetScriptTimeout(TimeSpan.FromSeconds(60));
                    wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(locator)));
                    Debug.WriteLine(_firefox.FindElement(By.XPath(locator)).GetAttribute("title") + "  tittle in method");
                    if (_firefox.FindElement(By.XPath(locator)).GetAttribute("title") == text)
                    {
                        if (!first) Debug.WriteLine("Text is found: " + text);
                        break; //если элемент найден
                    }

                    if (first) Debug.WriteLine("Waiting for text is present: " + text);

                    first = false;
                    Thread.Sleep(waitRetryDelayMs);
                    Debug.WriteLine(milliSecond + "миллисекунд" + " / text example - " + text + " : " +
                                    "text was found - " + _firefox.FindElement(By.XPath(locator)).GetAttribute("title"));
                }
                catch (StaleElementReferenceException a)
                {
                    if (milliSecond < timeOut * 10000)
                        continue;
                    else
                    {
                        throw a;
                    }
                }
                catch (NoSuchElementException b)
                {
                    if (milliSecond < timeOut * 10000)
                        continue;
                    else
                    {
                        throw b;
                    }
                }
            }

        }

        public void LoginDemo2mPage()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            _firefox.Navigate().GoToUrl("http://pharmxplorer.com.ua");
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("submit")));
            pageElements.LoginElement.SendKeys("full_test");
            pageElements.PasswordElement.SendKeys("aspirin222");
            TryToClickWithoutException(PageElements.LoginButtonXPath, pageElements.LoginButton);
            WaitForAjax();
            TryToLoadPage("http://pharmxplorer.com.ua/282", ".//*[@id='MainContainer']");

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchMarketButtonXPath)));

            TryToClickWithoutException(PageElements.SearchMarketButtonXPath, pageElements.SearchMarketButton);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.InputMarketFieldXPath)));
            pageElements.InputMarketField.SendKeys("Sandoz");

            WaitForTextInTitleAttribute(".//*[@class='QvFrame Document_LB1431']/div[3]/div/div[1]/div",
                "Sandoz  (Switzerland)");
            pageElements.SelectedMarketField.Click();
            WaitForAjax();
            Thread.Sleep(500);
            TryToClickWithoutException(PageElements.ContinueButtonXPath, pageElements.ContinueButton);
            WaitForAjax();
            Thread.Sleep(6000);
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='MainContainer']")));

        }

        public void SetUpFiltersSalesOut()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.MarketAnalysisTabXPath)));
            TryToClickWithoutException(PageElements.MarketAnalysisTabXPath, pageElements.MarketAnalysisTab);
            WaitForAjax();
            TryToClickWithoutException(PageElements.SelectDimensionSaleOutXPath, pageElements.SelectDimensionSaleOut);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SelectDimensionSaleOutDropDownXPath)));
            TryToClickWithoutException(PageElements.BrandXPath, pageElements.Brand);
            WaitForAjax();
            Thread.Sleep(2000);
            TryToClickWithoutException(PageElements.MeasureXPath, pageElements.Measure);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.MeasureDropDownXPath)));
            Console.WriteLine(_firefox.FindElement(By.XPath(PageElements.PcsMeasureXPath)).GetAttribute("title") +
                              " - title of pcs");
            TryToClickWithoutException(PageElements.PcsMeasureXPath, pageElements.PcsMeasure);
            WaitForAjax();
            TryToClickWithoutException(PageElements.TopAllXPath, pageElements.TopAll);
            WaitForAjax();
            Console.WriteLine("Filters have been set up.");
        }

        public void SetUpFiltersPromo()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.PromoAnalysisTabXPath)));
            TryToClickWithoutException(PageElements.PromoAnalysisTabXPath, pageElements.PromoAnalysisTab);
            WaitForAjax();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SelectDimentionPromoXPath)));
            TryToClickWithoutException(PageElements.SelectDimentionPromoXPath, pageElements.SelectDimentionPromo);
            wait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath(PageElements.DropDownMenuSelectDimentionPromoBrandXPath)));
            TryToClickWithoutException(PageElements.DropDownMenuSelectDimentionPromoBrandXPath,
                pageElements.DropDownMenuSelectDimentionBrand);
            WaitForAjax();
            TryToClickWithoutException(PageElements.ConfigTabPromoXPath, pageElements.ConfigTabPromo);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.CountryRegionPromoXPath)));
            TryToClickWithoutException(PageElements.CountryRegionPromoXPath, pageElements.CountryRegionPromo);
            WaitForAjax();
            TryToClickWithoutException(PageElements.TopAllPromoXPath, pageElements.TopAllPromo);
            WaitForAjax();

        }

        public void SetUpFiltersTV_Press_Radio(string advert)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.AdvertisingAnalysisTabXPath)));
            TryToClickWithoutException(PageElements.AdvertisingAnalysisTabXPath, pageElements.AdvertisingAnalysisTab);
            WaitForAjax();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SelectDimentionTvPressRadioXPath)));
            if (advert == "TV") TryToClickWithoutException(PageElements.TvButtonXPath, pageElements.TvButton);
            if (advert == "Press")
            {
                TryToClickWithoutException(PageElements.PressButtonXPath, pageElements.PressButton);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.MeasurePopUpCloseButtonXPath)));
                TryToClickWithoutException(PageElements.MeasurePopUpUAHXPath, pageElements.MeasurePopUpUAH);
            }
            if (advert == "Radio")
            {
                TryToClickWithoutException(PageElements.RadioButtonXPath, pageElements.RadioButton);
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.MeasurePopUpCloseButtonXPath)));
                TryToClickWithoutException(PageElements.MeasurePopUpUAHXPath, pageElements.MeasurePopUpUAH);
            }
            WaitForAjax();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.MeasurePopUpCloseButtonXPath)));
            TryToClickWithoutException(PageElements.MeasurePopUpCloseButtonXPath, pageElements.MeasurePopUpCloseButton);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.TopAllTV_Press_RadioXPath)));
            TryToClickWithoutException(PageElements.SelectDimentionTvPressRadioXPath, pageElements.SelectDimentionTvPressRadio);
            wait.Until(
                ExpectedConditions.ElementIsVisible(By.XPath(PageElements.DropDownMenuSelectDimentionPromoBrandXPath)));
            TryToClickWithoutException(PageElements.DropDownMenuSelectDimentionPromoBrandXPath,
                pageElements.DropDownMenuSelectDimentionBrand);
            WaitForAjax();

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.ConfigTabTV_Press_RadioXPath)));

            TryToClickWithoutException(PageElements.ConfigTabTV_Press_RadioXPath, pageElements.ConfigTabTV_Press_Radio);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.CountryRegionTV_Press_RadioXPath)));
            TryToClickWithoutException(PageElements.CountryRegionTV_Press_RadioXPath, pageElements.CountryRegionTV_Press_Radio);
            WaitForAjax();
            TryToClickWithoutException(PageElements.TopAllTV_Press_RadioXPath, pageElements.TopAllTV_Press_Radio);
            WaitForAjax();

        }

        public void SetUpPeriodSaleOut(string period)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            if (period.Contains('_'))
            {
                TryToClickWithoutException(PageElements.MonthPeriodXPath, pageElements.MonthPeriod);
            }
            else
            {
                if (period.Contains('Q'))
                {
                    TryToClickWithoutException(PageElements.QrtPeriodXPath, pageElements.QrtPeriod);
                }
                else
                {
                    TryToClickWithoutException(PageElements.YearPeriodXPath, pageElements.YearPeriod);
                }
            }
            WaitForAjax();
            TryToClickWithoutException(PageElements.DropDownPeriodButtonXPath, pageElements.DropDownPeriodButton);

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.DropDownMenuPeriodXPath)));
            action.ContextClick(pageElements.FirstFieldDropDownMenu).Perform();

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchOptionXPath)));
            pageElements.SearchOption.Click();
            Thread.Sleep(1000);

            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".PopupSearch>input")));
            pageElements.PeriodInputField.SendKeys(period + Keys.Enter);
            WaitForAjax();
            Thread.Sleep(2000);
        }

        public void SetUpPeriodPromo(string period)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            if (period.Contains('_'))
            {
                TryToClickWithoutException(PageElements.MonthPeriodPromoXPath, pageElements.MonthPeriodPromo);
            }
            else
            {
                if (period.Contains('Q'))
                {
                    TryToClickWithoutException(PageElements.QrtPeriodPromoXPath, pageElements.QrtPeriodPromo);
                }
                else
                {
                    TryToClickWithoutException(PageElements.YearPeriodPromoXPath, pageElements.YearPeriodPromo);
                }
            }
            WaitForAjax();
            TryToClickWithoutException(PageElements.PeriodPromoXPath, pageElements.PeriodPromo);

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.PeriodDropDownPromoXPath)));
            action.MoveToElement(pageElements.PeriodDropDownPromo, 10, 10).ContextClick().Perform();

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchOptionXPath)));
            TryToClickWithoutException(PageElements.SearchOptionXPath, pageElements.SearchOption);

            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".PopupSearch>input")));
            pageElements.PeriodInputField.SendKeys(period + Keys.Enter);
            WaitForAjax();
            Thread.Sleep(2000);
        }

        public void SetUpPeriodTvPressRadio(string period)
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            if (period.Contains('_'))
            {
                TryToClickWithoutException(PageElements.MonthPeriodTV_Press_RadioXPath, pageElements.MonthPeriodTV_Press_Radio);
            }
            else
            {
                if (period.Contains('Q'))
                {
                    TryToClickWithoutException(PageElements.QrtPeriodPeriodTV_Press_RadioXPath, pageElements.QrtPeriodPeriodTV_Press_Radio);
                }
                else
                {
                    TryToClickWithoutException(PageElements.YearPeriodPeriodTV_Press_RadioXPath, pageElements.YearPeriodPeriodTV_Press_Radio);
                }
            }
            WaitForAjax();
            TryToClickWithoutException(PageElements.PeriodTV_Press_RadioXPath, pageElements.PeriodTV_Press_Radio);

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.PeriodDropDownPromoXPath)));
            action.MoveToElement(pageElements.PeriodDropDownPromo, 10, 10).ContextClick().Perform();

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchOptionXPath)));
            TryToClickWithoutException(PageElements.SearchOptionXPath, pageElements.SearchOption);

            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".PopupSearch>input")));
            pageElements.PeriodInputField.SendKeys(period + Keys.Enter);
            WaitForAjax();
            Thread.Sleep(2000);
        }

        public void SendToExcel()
        {

            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            TryToClickWithoutException(PageElements.SendToExcelButtonXPath, pageElements.SendToExcelButtonSaleOut);

        }

        public void SendToExcelPromo()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            TryToClickWithoutException(PageElements.SendToExcelPromoXPath, pageElements.SendToExcelPromo);

        }

        public void SendToExcelTV_Press_Radio()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            TryToClickWithoutException(PageElements.SendToExcelTV_Press_RadioXPath, pageElements.SendToExcelTV_Press_Radio);

        }

        public void StoreExcelDataFromWeb()
        {
            var directory = new DirectoryInfo(@"D:\DownloadTest");
            var myFile = (from f in directory.GetFiles()
                          orderby f.LastWriteTime descending
                          select f).First();

            Console.WriteLine(myFile.Name);

            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\DownloadTest\" + myFile, "SELECT * from [Sheet1$A2:B]");
            /* WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\DownloadTest\temp.xls", "SELECT * from [Sheet1$A2:B]");*/

            // Console.WriteLine(dt.Rows.Count);
            dt.Rows[0].Delete();
            dt.Rows[dt.Rows.Count - 1].Delete();
            dt.AcceptChanges();
            //Console.WriteLine(dt.Rows.Count);

            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var brand = row["Brand"].ToString().Trim().Replace("\u00A0", " ").ToLower();
                //Console.WriteLine(brand);
                var rowData = new RowData
                {
                    Brand = Regex.Replace(brand, @"\s+", " ").Substring(6),
                    ComparedValue = Convert.ToDecimal(row["pcs"])

                };
                preparationPcsWeb282.Add(rowData);
            }
            /*Console.WriteLine("Данные дашборда");
               for (int i = 0; i < 9; i++)
               {

                   Console.WriteLine(preparationPcsWeb282[i].Brand + " / " + preparationPcsWeb282[i].Upakovki);

               }*/
        }

        public void StoreExcelDataFromWebPromo()
        {
            var directory = new DirectoryInfo(@"D:\DownloadTest");
            var myFile = (from f in directory.GetFiles()
                          orderby f.LastWriteTime descending
                          select f).First();


            Console.WriteLine(myFile.Name);

            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\DownloadTest\" + myFile, "SELECT * from [Sheet1$A1:C]");

            dt.Columns.Remove("Region");
            /* Console.WriteLine(dt.Columns[0].ColumnName + " / "+ dt.Columns[1].ColumnName);*/
            dt.Rows[0].Delete();
            dt.AcceptChanges();


            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var brand = row["Brand"].ToString().Trim().Replace("\u00A0", " ").ToLower();
                //Console.WriteLine(brand);
                var rowData = new RowData
                {
                    Brand = Regex.Replace(brand, @"\s+", " ").Substring(6),
                    ComparedValue = Convert.ToDecimal(row["Total"])

                };
                preparationPcsWeb282.Add(rowData);
            }
            /* Console.WriteLine("Данные дашборда");
                for (int i = 0; i < 9; i++)
                {

                    Console.WriteLine(preparationPcsWeb282[i].Brand + " / " + preparationPcsWeb282[i].Upakovki);

                }*/
        }

        public void StoreExcelDataFromWebTV_Press_Radio()
        {
            var directory = new DirectoryInfo(@"D:\DownloadTest");
            var myFile = (from f in directory.GetFiles()
                          orderby f.LastWriteTime descending
                          select f).First();


            Console.WriteLine(myFile.Name);

            DataTable dt = new DataTable();
            WorkWithExcelFile.ExcelFileToDataTable(out dt, @"D:\DownloadTest\" + myFile, "SELECT * from [Sheet1$B1:C]");
            dt.Columns[1].ColumnName = "Total";

            foreach (DataRow row in dt.Rows)
            {
                if (row[0] == DBNull.Value) continue;

                var brand = row["Brand"].ToString().Trim().Replace("\u00A0", " ").ToLower();
                //Console.WriteLine(brand);
                var rowData = new RowData
                {
                    Brand = Regex.Replace(brand, @"\s+", " ").Substring(6),
                    ComparedValue = Convert.ToDecimal(row["Total"])

                };
                preparationPcsWeb282.Add(rowData);
            }
           /* Console.WriteLine("Данные дашборда");
            for (int i = 0; i < 5; i++)
            {

                Console.WriteLine(preparationPcsWeb282[i].Brand + " / " + preparationPcsWeb282[i].ComparedValue);

            }*/
        }

        public void StoreExcelDataFromDB(string periodFrom, string periodTo) //period 201601006
        {
            string connectionString = "Database=morion_2009:morion_database; User=TESTER; Password=654654";
            FbConnection conn = new FbConnection(connectionString);
            conn.Open();

            FbDatabaseInfo inf = new FbDatabaseInfo(conn);
            Console.WriteLine("Info: " + inf.ServerClass + ";" + inf.ServerVersion);
            if (conn.State == ConnectionState.Closed) conn.Open();

            FbTransaction myTransaction = conn.BeginTransaction();
            FbCommand myCommand = new FbCommand("select iif(d.BRAND_ID = 0, d.NAME_ENG, b.NAME_ENG) as BRAND, " +
                                                 " sum(p.Q) as Q " +
                                                 "from QV_BALD_SO(" + periodFrom + ", " + periodTo + ", 1, '3') p " +
                                                 "left join rep_v_drugs d on d.id = p.drugs_id " +
                                                 "left join M_BRANDS b on b.id = d.BRAND_ID " +
                                                 "where d.category_id = 1 " +
                                                 "group by 1 " +
                                                 "order by 2 desc", conn);
            myCommand.Transaction = myTransaction;
            FbDataReader reader = myCommand.ExecuteReader();

            try
            {
                while (reader.Read())
                {

                    var rowData = new RowData
                    {
                        Brand = reader["BRAND"].ToString().Trim().ToLower(),
                        ComparedValue = Convert.ToDecimal(reader["Q"])
                    };

                    preparationPcsDbList.Add(rowData);
                }
            }
            finally
            {
                conn.Close();
            }
            /*Console.WriteLine("Данные БД");
            for (int i = 0; i < 20; i++)
            {
                Console.WriteLine(preparationPcsDbList[i].Brand + " / " + preparationPcsDbList[i].Upakovki);
            }*/
            myCommand.Dispose();
        }

        public void StoreExcelDataFromDB_Promo(string periodFrom, string periodTo) //period 201601006
        {
            string connectionString = "Database=morion_2009:morion_database; User=TESTER; Password=654654";
            FbConnection conn = new FbConnection(connectionString);
            conn.Open();

            FbDatabaseInfo inf = new FbDatabaseInfo(conn);
            Console.WriteLine("Info: " + inf.ServerClass + ";" + inf.ServerVersion);
            if (conn.State == ConnectionState.Closed) conn.Open();

            FbTransaction myTransaction = conn.BeginTransaction();
            FbCommand myCommand = new FbCommand("select iif(md.BRAND_ID=0, md.NAME_ENG, b.NAME_ENG) as BRAND, " +
                                                "sum(k_25) as K25 " +
                                                "From  qv_bald_promo_ar(" + periodFrom + "," + periodTo + ", 1) pr " +
                                                "left join M_drugS md on md.id = pr.drugs_id " +
                                                "left join M_BRANDS b on b.id = md.brand_id " +
                                                "left join rep_data_layers rdl on rdl.id = pr.data_layer_id " +
                                                "where pr.promo_type_id = 6  and  md.category_id = 1 " +
                                                "group by 1 " +
                                                "order by 2 desc", conn);

            myCommand.Transaction = myTransaction;
            FbDataReader reader = myCommand.ExecuteReader();

            try
            {
                while (reader.Read())
                {
                    var rowData = new RowData
                    {
                        Brand = reader["BRAND"].ToString().Trim().ToLower(),
                        ComparedValue = Convert.ToDecimal(reader["K25"])
                    };

                    preparationPcsDbList.Add(rowData);
                }
            }
            finally
            {
                conn.Close();
            }
            /*  Console.WriteLine("Данные БД");
              for (int i = 0; i < 20; i++)
              {
                  Console.WriteLine(preparationPcsDbList[i].Brand + " / " + preparationPcsDbList[i].ComparedValue);
              }*/
            myCommand.Dispose();
        }

        public void StoreExcelDataFromDB_TV_Press_Radio(string periodFrom, string periodTo, string advert) //period 201601006
        {

            string queryTV = "select iif(d.BRAND_ID = 0, d.NAME_ENG, b.NAME_ENG) as BRAND," +
                          " sum(p.FIELD_M_1067) as Total " +
                          "from QV_UD_106_m1(" + periodFrom + "," + periodTo + ", 1, 1) p " +
                          "left join rep_v_drugs d on d.id = p.drugs_id " +
                          "left join M_BRANDS b on b.id = d.BRAND_ID " +
                          "where d.category_id = 1 " +
                          "group by 1 " +
                             "order by Total desc";

            var queryPress = "select iif(d.BRAND_ID = 0, d.NAME_ENG, b.NAME_ENG) as BRAND, " +
                             "sum(p.ADV_V_UAH) as Total " +
                             "from QV_UD_179_m1(" + periodFrom + "," + periodTo + ", 1, 1) p " +
                             "left join rep_v_drugs d on d.id = p.drugs_id " +
                             "left join M_BRANDS b on b.id = d.BRAND_ID " +
                             "where d.category_id = 1 " +
                             "group by 1 " +
                             "order by Total desc";

            var queryRadio = "select iif(d.BRAND_ID = 0, d.NAME_ENG, b.NAME_ENG) as BRAND, " +
                             "sum(p.ADV_V_UAH) as Total " +
                             "from QV_UD_184_m1(" + periodFrom + "," + periodTo + ", 1, 1) p " +
                             "left join rep_v_drugs d on d.id = p.drugs_id " +
                             "left join M_BRANDS b on b.id = d.BRAND_ID " +
                             "where d.category_id = 1 " +
                             "group by 1 " +
                             "order by Total desc";

            string connectionString = "Database=morion_2009:morion_database; User=TESTER; Password=654654";
            FbConnection conn = new FbConnection(connectionString);
            conn.Open();

            FbDatabaseInfo inf = new FbDatabaseInfo(conn);
            Console.WriteLine("Info: " + inf.ServerClass + ";" + inf.ServerVersion);
            if (conn.State == ConnectionState.Closed) conn.Open();

            FbTransaction myTransaction = conn.BeginTransaction();
            FbCommand myCommand = new FbCommand();
            if (advert == "TV") myCommand.CommandText = queryTV;
            if (advert == "Press") myCommand.CommandText = queryPress;
            if (advert == "Radio") myCommand.CommandText = queryRadio;
            myCommand.Connection = conn;
            myCommand.Transaction = myTransaction;
            FbDataReader reader = myCommand.ExecuteReader();

            try
            {
                while (reader.Read())
                {
                    var rowData = new RowData
                    {
                        Brand = reader["BRAND"].ToString().Trim().ToLower(),
                        ComparedValue = Convert.ToDecimal(reader["Total"])
                    };

                    preparationPcsDbList.Add(rowData);
                }
            }
            finally
            {
                conn.Close();
            }
            /*Console.WriteLine("Данные БД");
            for (int i = 0; i < 10; i++)
            {
                Console.WriteLine(preparationPcsDbList[i].Brand + " / " + preparationPcsDbList[i].ComparedValue);
            }*/
            myCommand.Dispose();
        }

        public void Compare()
        {
            var differenceTotal = RowDataList.CompareTotal(preparationPcsWeb282, preparationPcsDbList);
            var difference = RowDataList.ComparePCS(preparationPcsWeb282, preparationPcsDbList);
            Console.WriteLine(differenceTotal);
            if (difference.Count > 0)
            {
                /* foreach (var d in difference)
                 {
                     Console.WriteLine(d);
                 }*/
                string path = @"D:\Sneghka\WriteAllLinesDifferenceDemo2m.xls";

                File.WriteAllLines(path, difference);

            }
            else
            {

                Console.WriteLine("Данные совпадают");
            }

        }

        public void email_send(string subject)
        {
            var mail = new MailMessage();
            mail.IsBodyHtml = true;
            var smtpServer = new SmtpClient("post.morion.ua");
            mail.From = new MailAddress("snizhana.nomirovska@proximaresearch.com");
            mail.To.Add("snizhana.nomirovska@proximaresearch.com");
            //mail.To.Add("nataly.tenkova@proximaresearch.com");
            mail.Subject = subject;
            mail.Body = MessageContent(messageContent);
            smtpServer.Send(mail);

        }
    }
}
