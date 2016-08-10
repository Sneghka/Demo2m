using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Support.UI;


namespace Demo2m
{
    public class PageElements
    {
        private readonly FirefoxDriver _firefox;

        public PageElements(FirefoxDriver firefox)
        {
            _firefox = firefox;

        }

        #region PageElementXpath

        public const string LoginButtonXPath = ".//*[@id='submit']";
        public const string SearchMarketButtonXPath = ".//*[@class='QvFrame Document_LB1431']/div[2]/div[1]/div";
        public const string InputMarketFieldXPath = "html/body/div[2]/input";
        public const string SelectedMarketXPath = ".//*[@class='QvFrame Document_LB1431']/div[3]/div/div[1]/div";
        public const string ContinueButtonXPath = ".//*[@class='QvFrame Document_BU954']/div[3]/button";
        public const string MarketAnalysisTabXPath = ".//*[@rel='DocumentSH42']/a";
        public const string SelectDimensionXPath = ".//*[@class='QvFrame Document_MB289']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string SelectDimensionDropDownXPath = ".//*[@class='QvFrame DS']/div/div/div[1]";
        public const string SkuXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div[25]";
        public const string BrandXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div[9]";
        public const string MeasureXPath = ".//*[@class='QvFrame Document_BU1267']/div[3]/button";
        public const string MeasureDropDownXPath = ".//*[@class='QvFrame Document_LB1124']/div[2]/div/div[1]";
        public const string PcsMeasureXPath = ".//*[@class='QvFrame Document_LB1124']/div[2]/div/div[1]/div[4]";
        public const string TopAllXPath = ".//*[@class='QvFrame Document_LB2238']/div[3]/div/div[1]/div[13]";
        public const string QrtPeriodXPath = ".//*[@class='QvFrame Document_LB2232']/div[3]/div/div[1]/div[2]/div[1]";
        public const string MonthPeriodXPath = ".//*[@class='QvFrame Document_LB2232']/div[3]/div/div[1]/div[1]/div[1]";
        public const string YearPeriodXPath = ".//*[@class='QvFrame Document_LB2232']/div[3]/div/div[1]/div[3]/div[1]";
        public const string DropDownPeriodButtonXPath = ".//*[@class='QvFrame Document_MB535']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string DropDownMenuPeriodXPath = ".//*[@class='QvFrame DS']/div/div/div[1]";
        public const string SearchOptionXPath = "html/body/ul/li[1]/a";
        //public const string PeriodInputFieldXPath = "html/body/div[3]/input";
        public const string SelectedPeriodXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div/div[1]";
        public const string SendToExcelButtonXPath = ".//*[@class='QvFrame Document_CH658']/div[2]/div[1]/div[1]";
        public const string OpenHereXlsLinkXPath = "html/body/div[12]/div[2]/div/a";
        #endregion

        #region LoginElements

        public IWebElement LoginElement
        {
            get { return _firefox.FindElement(By.Id("username")); }
        }
        public IWebElement PasswordElement
        {
            get { return _firefox.FindElement(By.Id("password")); }
        }

        public IWebElement LoginButton
        {
            get { return _firefox.FindElement(By.Id("submit")); }
        }

        public IWebElement SearchMarketButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB1431']/div[2]/div[1]/div")); }
        }

        public IWebElement InputMarketField
        {
            get { return _firefox.FindElement(By.XPath("html/body/div[2]/input")); }
        }
        public IWebElement SelectedMarketField
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB1431']/div[3]/div/div[1]/div")); }
        }

        public IWebElement ContinueButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_BU954']/div[3]/button")); }
        }
        public IWebElement MarketAnalysisTab
        {
            get { return _firefox.FindElement(By.XPath(".//*[@rel='DocumentSH42']/a")); }
        }
        #endregion

        #region SetUpFilter

        public IWebElement SelectDimension
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB289']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }
        public IWebElement SelectDimensionDropDown
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]")); }
        }
        public IWebElement Sku
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div[25]")); }
        }

        public IWebElement Brand
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div[9]")); }
        }

        public IWebElement Measure
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_BU1267']/div[3]/button")); }
        }
        public IWebElement MeasureDropDown
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB1124']/div[2]/div/div[1]")); }
        }
        public IWebElement PcsMeasure
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB1124']/div[2]/div/div[1]/div[4]")); }
        }
        public IWebElement TopAll
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2238']/div[3]/div/div[1]/div[13]")); }
        }


        #endregion

        #region Period

        public IWebElement MonthPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2232']/div[3]/div/div[1]/div[1]/div[1]")); }
        }

        public IWebElement QrtPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2232']/div[3]/div/div[1]/div[2]/div[1]")); }
        }

        public IWebElement YearPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2232']/div[3]/div/div[1]/div[3]/div[1]")); }
        }
        public IWebElement DropDownPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB535']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }
        public IWebElement DropDownMenuPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]")); }
        }
        public IWebElement PeriodInputField
        {
            get { return _firefox.FindElement(By.CssSelector(".PopupSearch>input")); }
        }
        public IWebElement SelectedPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div/div[1]")); }
        }
        public IWebElement SearchOption
        {
            get { return _firefox.FindElement(By.XPath("html/body/ul/li[1]/a")); }
        }

        #endregion

        #region SendToExcel

        public IWebElement SendToExcelButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_CH658']/div[2]/div[1]/div[1]")); }
        }

        public IWebElement OpenHereXlsLink
        {
            get { return _firefox.FindElement(By.XPath("html/body/div[12]/div[2]/div/a")); }
        }
        #endregion
    }
}
