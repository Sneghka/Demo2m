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
        public FirefoxProfile SetUpFirefoxProfile()
        {
            var downloadPath = @"D:\DownloadTest";
            FirefoxProfile firefoxProfile = new FirefoxProfile();
            firefoxProfile.SetPreference("browser.download.folderList", 2);
            firefoxProfile.SetPreference("browser.download.dir", downloadPath);
            firefoxProfile.SetPreference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel");
            return firefoxProfile;
        }

        [Test]
        public void Demo2m_SalesOut_ByBrand()
        {

            var firefox = new FirefoxDriver(SetUpFirefoxProfile());
            var methods = new Methods(firefox);
            methods.LoginDemo2mPage();
            methods.SetUpFiltersSalesOut();
            /* methods.SetUpPeriod("2016 06_Jun");*/
            methods.SendToExcel();
            Thread.Sleep(15000);
            methods.StoreExcelDataFromWeb();
            methods.StoreExcelDataFromDB("201601008", "201601008");
            methods.Compare();
            firefox.Quit();

        }

        [Test]
        public void Demo2m_Promo30_ByBrand()
        {
            var firefox = new FirefoxDriver(SetUpFirefoxProfile());
            var methods = new Methods(firefox);
            methods.LoginDemo2mPage();
            methods.SetUpFiltersPromo();
          /*  methods.SetUpPeriodPromo("2016 07_Jul");*/
            methods.SendToExcelPromo();
            Thread.Sleep(15000);
            methods.StoreExcelDataFromWebPromo();
            methods.StoreExcelDataFromDB_Promo("201601008", "201601008");
            methods.Compare();
          
            firefox.Quit();
        }

        [Test]
        public void Demo2m_Press_TV_Radio_ByBrand()
        {
            var firefox = new FirefoxDriver(SetUpFirefoxProfile());
            var methods = new Methods(firefox);
            methods.LoginDemo2mPage();
            methods.SetUpFiltersTV_Press_Radio("Press");   // parametr TV || Press || Radio
            //methods.SetUpPeriodTvPressRadio("2016 06_Jun");
            methods.SendToExcelTV_Press_Radio();
            Thread.Sleep(15000);
            methods.StoreExcelDataFromWebTV_Press_Radio();
            methods.StoreExcelDataFromDB_TV_Press_Radio("201601008", "201601008", "Press");  // parametr TV || Press || Radio
            methods.Compare();
            firefox.Quit();

        }

    }
}
