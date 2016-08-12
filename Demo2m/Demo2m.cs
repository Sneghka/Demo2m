using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using NUnit.Framework;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;





namespace Demo2m
{
    [TestFixture]
    public class Demo2m
    {
        [Test]
        public void Demo2m_SalesOut()
        {
            var downloadPath = @"D:\DownloadTest";
            FirefoxProfile firefoxProfile = new FirefoxProfile();
            firefoxProfile.SetPreference("browser.download.folderList", 2);
            firefoxProfile.SetPreference("browser.download.dir", downloadPath);
            firefoxProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
            var firefox = new FirefoxDriver(firefoxProfile);
            var methods = new Methods(firefox);
           methods.LoginDemo2mPage();
            methods.SetUpFiltersSalesOut();
            methods.SendToExcel();
            Thread.Sleep(10000);
            methods.StoreExcelDataFromWeb();
            methods.StoreExcelDataFromDB();
            methods.Compare();
            
            firefox.Quit();

        }


    }
}
