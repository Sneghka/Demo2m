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
            //Console.WriteLine("In tryToClick method + xpath: " + locator);
            var maxElementRetries = 100;
            var action = new Actions(_firefox);
            var retries = 0;
            while (true)
            {
                /*WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));*/

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
                        Debug.WriteLine(retries);
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
                    Debug.WriteLine(milliSecond + "миллисекунд" + " / text example - " + text + " : " + "text was found - " + _firefox.FindElement(By.XPath(locator)).GetAttribute("title"));
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
            _firefox.Navigate().GoToUrl("http://pharmxplorer.com.ua/282");
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.Id("submit")));
            pageElements.LoginElement.SendKeys("full_test");
            pageElements.PasswordElement.SendKeys("aspirin222");
            TryToClickWithoutException(PageElements.LoginButtonXPath, pageElements.LoginButton);
            WaitForAjax();

            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchMarketButtonXPath)));
            
            TryToClickWithoutException(PageElements.SearchMarketButtonXPath, pageElements.SearchMarketButton);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.InputMarketFieldXPath)));
            pageElements.InputMarketField.SendKeys("Sandoz");

            WaitForTextInTitleAttribute(".//*[@class='QvFrame Document_LB1431']/div[3]/div/div[1]/div", "Sandoz  (Switzerland)");
            pageElements.SelectedMarketField.Click();
            WaitForAjax();
            TryToClickWithoutException(PageElements.ContinueButtonXPath, pageElements.ContinueButton);
            WaitForAjax();
            Thread.Sleep(6000);
            wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(".//*[@id='MainContainer']")));
            TryToClickWithoutException(PageElements.MarketAnalysisTabXPath, pageElements.MarketAnalysisTab);
            WaitForAjax();
        }

        public void SetUpFiltersSalesOut()
        {
            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var action = new Actions(_firefox);
            var pageElements = new PageElements(_firefox);
            TryToClickWithoutException(PageElements.SelectDimensionXPath, pageElements.SelectDimension);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SelectDimensionDropDownXPath)));
            TryToClickWithoutException(PageElements.BrandXPath, pageElements.Brand);
            WaitForAjax();
            Thread.Sleep(2000);
            TryToClickWithoutException(PageElements.MeasureXPath, pageElements.Measure);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.MeasureDropDownXPath)));
            Debug.WriteLine(_firefox.FindElement(By.XPath(PageElements.PcsMeasureXPath)).GetAttribute("title") + " - title of pcs");
            TryToClickWithoutException(PageElements.PcsMeasureXPath, pageElements.PcsMeasure);
            WaitForAjax();
            TryToClickWithoutException(PageElements.TopAllXPath, pageElements.TopAll);
            WaitForAjax();
            Debug.WriteLine("Filters have been set up.");
        }

        public void SetUpPeriod(string period)
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
            Thread.Sleep(4000);
            pageElements.DropDownPeriod.Click();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.DropDownMenuPeriodXPath)));
            action.ContextClick(pageElements.DropDownMenuPeriod).Perform();
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.SearchOptionXPath)));
            pageElements.SearchOption.Click();
            Thread.Sleep(2000);
            Debug.WriteLine("try to send key - period");
            wait.Until(ExpectedConditions.ElementIsVisible(By.CssSelector(".PopupSearch>input")));
            Debug.WriteLine("try to send key - period");
            pageElements.PeriodInputField.SendKeys(period);
            Thread.Sleep(4000);
            pageElements.SelectedPeriod.Click();
            WaitForAjax();
            Thread.Sleep(4000);
        }


        public void SendToExcel()
        {

            WebDriverWait wait = new WebDriverWait(_firefox, TimeSpan.FromSeconds(120));
            var pageElements = new PageElements(_firefox);
            var action = new Actions(_firefox);

            TryToClickWithoutException(PageElements.SendToExcelButtonXPath, pageElements.SendToExcelButton);
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath(PageElements.OpenHereXlsLinkXPath)));
            // TryToClickWithoutException(PageElements.OpenHereXlsLinkXPath, pageElements.OpenHereXlsLink);

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
                    Upakovki = Convert.ToDecimal(row["pcs"])

                };
                preparationPcsWeb282.Add(rowData);
            }
            /*Console.WriteLine("Данные дашборда");
               for (int i = 0; i < 9; i++)
               {

                   Console.WriteLine(preparationPcsWeb282[i].Brand + " / " + preparationPcsWeb282[i].Upakovki);

               }*/
        }

        public void StoreExcelDataFromDB()
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
                                                 "from QV_BALD_SO(201601006, 201601006, 1, '3') p " +
                                                 "left join rep_v_drugs d on d.id = p.drugs_id " +
                                                 "left join M_BRANDS b on b.id = d.BRAND_ID " +
                                                 "where d.category_id = 1 " +
                                                 "group by 1 " +
                                                 "order by 2 desc", conn);
            myCommand.Transaction = myTransaction;
            FbDataReader reader = myCommand.ExecuteReader();

            List<string> idList = new List<string>();
            List<string> parentOrgList = new List<string>();
            List<string> nameList = new List<string>();
            try
            {
                while (reader.Read())
                {

                    var rowData = new RowData
                    {
                        Brand = reader["BRAND"].ToString().Trim().ToLower(),
                        Upakovki = Convert.ToDecimal(reader["Q"])
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

        public void Compare()
        {
            var differenceTotal = RowDataList.CompareTotal(preparationPcsWeb282, preparationPcsDbList);
            var difference = RowDataList.ComparePCS(preparationPcsWeb282, preparationPcsDbList);
            Console.WriteLine(differenceTotal);
            if (difference.Count > 0)
            {
                foreach (var d in difference)
                {
                    Console.WriteLine(d);
                }
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
