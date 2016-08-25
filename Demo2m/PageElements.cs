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

        #region PageElementSaleOutXpath
        public const string LoginButtonXPath = ".//*[@id='submit']";
        public const string SearchMarketButtonXPath = ".//*[@class='QvFrame Document_LB1431']/div[2]/div[1]/div";
        public const string InputMarketFieldXPath = "html/body/div[2]/input";
        public const string SelectedMarketXPath = ".//*[@class='QvFrame Document_LB1431']/div[3]/div/div[1]/div";
        public const string ContinueButtonXPath = ".//*[@class='QvFrame Document_BU954']/div[3]/button";
        public const string MarketAnalysisTabXPath = ".//*[@rel='DocumentSH42']/a";
        public const string PromoAnalysisTabXPath = ".//*[@rel='DocumentSH73']/a";
        public const string AdvertisingAnalysisTabXPath = ".//*[@rel='DocumentSH57']/a";
        public const string SelectDimensionSaleOutXPath = ".//*[@class='QvFrame Document_MB289']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string SelectDimensionSaleOutDropDownXPath = ".//*[@class='QvFrame DS']/div/div/div[1]";
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
        public const string PeriodInputFieldXPath = "html/body/div[3]/input";
        public const string FirstFieldDropDownMenuXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div[1]";
        public const string SelectedPeriodXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div/div[1]";
        public const string SendToExcelButtonXPath = ".//*[@class='QvFrame Document_CH658']/div[2]/div[1]/div[1]";
        public const string OpenHereXlsLinkXPath = "html/body/div[12]/div[2]/div/a";
        #endregion

        #region PageElementsPromoXPath

        public const string SelectDimentionPromoXPath = ".//*[@class='QvFrame Document_MB293']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string DropDownMenuSelectDimentionPromo = ".//*[@class='QvFrame DS']";
        public const string DropDownMenuSelectDimentionPromoBrandXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div[9]/div[1]";
        public const string DropDownDimentionPromoSelectedOptionXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div/div[1]";
        public const string TopAllPromoXPath = ".//*[@class='QvFrame Document_LB2259']/div[3]/div/div[1]/div[13]/div[1]";
        public const string PeriodPromoXPath = ".//*[@class='QvFrame Document_MB530']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string PeriodDropDownPromoXPath = ".//*[@class='QvFrame DS']/div/div/div[1]/div[1]/div[1]";
        public const string ConfigTabPromoXPath = ".//*[@class='QvFrame Document_TX1208']/div[2]/table/tbody/tr/td";
        public const string CountryRegionPromoXPath = ".//*[@class='QvFrame Document_LB2387']/div[3]/div/div[1]/div[1]";
        public const string SendToExcelPromoXPath = ".//*[@class='QvFrame Document_CH916']/div[2]/div[1]/div[3]";

        public const string MonthPeriodPromoXPath = ".//*[@class='QvFrame Document_LB2253']/div[3]/div/div[1]/div[1]/div[1]";
        public const string QrtPeriodPromoXPath = ".//*[@class='QvFrame Document_LB2253']/div[3]/div/div[1]/div[2]/div[1]";
        public const string YearPeriodPromoXPath = ".//*[@class='QvFrame Document_LB2253']/div[3]/div/div[1]/div[3]/div[1]";

        #endregion

        #region PageElementsTV_Press_Radio

        public const string SelectDimentionTvPressRadioXPath = ".//*[@class='QvFrame Document_MB294']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string TopAllTV_Press_RadioXPath = ".//*[@class='QvFrame Document_LB2266']/div[3]/div/div[1]/div[13]/div[1]";
        public const string PeriodTV_Press_RadioXPath = ".//*[@class='QvFrame Document_MB531']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]";
        public const string ConfigTabTV_Press_RadioXPath = ".//*[@class='QvFrame Document_TX1142']/div[2]/table/tbody/tr/td";
        public const string CountryRegionTV_Press_RadioXPath = ".//*[@class='QvFrame Document_LB2389']/div[3]/div/div[1]/div[1]";
        public const string SendToExcelTV_Press_RadioXPath = ".//*[@class='QvFrame Document_CH634']/div[2]/div[1]/div[4]";

        public const string MonthPeriodTV_Press_RadioXPath = ".//*[@class='QvFrame Document_LB2260']/div[3]/div/div[1]/div[1]/div[1]";
        public const string QrtPeriodPeriodTV_Press_RadioXPath = ".//*[@class='QvFrame Document_LB2260']/div[3]/div/div[1]/div[2]/div[1]";
        public const string YearPeriodPeriodTV_Press_RadioXPath = ".//*[@class='QvFrame Document_LB2260']/div[3]/div/div[1]/div[3]/div[1]";

        public const string TvButtonXPath = ".//*[@class='QvFrame Document_BU1311']/div[3]/button";
        public const string RadioButtonXPath = ".//*[@class='QvFrame Document_BU1313']/div[3]/button";
        public const string PressButtonXPath = ".//*[@class='QvFrame Document_BU1312']/div[3]/button";
        public const string MeasurePopUpCloseButtonXPath = ".//*[@class='QvFrame Document_TX1203']/div[2]/table/tbody/tr/td";
        public const string MeasurePopUpUAHXPath = ".//*[@class='QvFrame Document_LB1162']/div[2]/div/div[1]/div[3]/div[1]";

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

        #region SetUpFilter SaleOUT

        public IWebElement SelectDimensionSaleOut
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

        #region SetUpFilter Promo

        public IWebElement PromoAnalysisTab
        {
            get { return _firefox.FindElement(By.XPath(".//*[@rel='DocumentSH73']/a")); }
        }

        public IWebElement SelectDimentionPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB293']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }

        public IWebElement DropDownMenuSelectDimention
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']")); }
        }
        public IWebElement DropDownMenuSelectDimentionBrand
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div[9]/div[1]")); }
        }
        public IWebElement DropDownDimentionPromoSelectedOption
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div/div[1]")); }
        }
        public IWebElement TopAllPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2259']/div[3]/div/div[1]/div[13]/div[1]")); }
        }
        public IWebElement PeriodPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB530']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }
        public IWebElement PeriodDropDownPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div[1]/div[1]")); }
        }
        public IWebElement  ConfigTabPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_TX1208']/div[2]/table/tbody/tr/td")); }
        }
        public IWebElement CountryRegionPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2387']/div[3]/div/div[1]/div[1]")); }
        }
        #endregion

        #region SetUpFilters TV_Press_Radio

        public IWebElement AdvertisingAnalysisTab
        {
            get { return _firefox.FindElement(By.XPath(".//*[@rel='DocumentSH57']/a")); }
        }

        public IWebElement TopAllTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2266']/div[3]/div/div[1]/div[13]/div[1]"));}
        }

        public IWebElement PeriodTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB531']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }

        public IWebElement ConfigTabTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_TX1142']/div[2]/table/tbody/tr/td")); }
        }

        public IWebElement CountryRegionTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2389']/div[3]/div/div[1]/div[1]")); }
        }

        public IWebElement SelectDimentionTvPressRadio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB294']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }
        public IWebElement TvButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_BU1311']/div[3]/button")); }
        }
        public IWebElement RadioButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_BU1313']/div[3]/button")); }
        }
        public IWebElement PressButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_BU1312']/div[3]/button")); }
        }

        public IWebElement MeasurePopUpCloseButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_TX1203']/div[2]/table/tbody/tr/td")); }
        }

        public IWebElement MeasurePopUpUAH
        {
            get{return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB1162']/div[2]/div/div[1]/div[3]/div[1]"));}
        }

        #endregion

        #region PeriodSaleOut

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
        public IWebElement DropDownPeriodButton
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_MB535']/div[3]/div/div[1]/div[5]/div/div[1]/div[1]")); }
        }
        public IWebElement DropDownMenuPeriod
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]")); }
        }

        public IWebElement FirstFieldDropDownMenu
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame DS']/div/div/div[1]/div[1]")); }
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

        #region PeriodPromo

        public IWebElement MonthPeriodPromo
        {
            get{return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2253']/div[3]/div/div[1]/div[1]/div[1]"));}
        }

        public IWebElement QrtPeriodPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2253']/div[3]/div/div[1]/div[2]/div[1]")); }
        }
        public IWebElement YearPeriodPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2253']/div[3]/div/div[1]/div[3]/div[1]")); }
        }

        #endregion

        #region PeriodTV_Press_Radio

        public IWebElement MonthPeriodTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2260']/div[3]/div/div[1]/div[1]/div[1]")); }
        }

        public IWebElement QrtPeriodPeriodTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2260']/div[3]/div/div[1]/div[2]/div[1]")); }
        }
        public IWebElement YearPeriodPeriodTV_Press_Radio
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_LB2260']/div[3]/div/div[1]/div[3]/div[1]")); }
        }

        #endregion


        #region SendToExcel

        public IWebElement SendToExcelButtonSaleOut
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_CH658']/div[2]/div[1]/div[1]")); }
        }

        public IWebElement SendToExcelPromo
        {
            get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_CH916']/div[2]/div[1]/div[3]")); }
        }

        public IWebElement SendToExcelTV_Press_Radio
        {
          get { return _firefox.FindElement(By.XPath(".//*[@class='QvFrame Document_CH634']/div[2]/div[1]/div[4]")); }
        }


        #endregion
    }
}
