using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Threading;
using OpenQA.Selenium.Interactions;
using ExcelDataReader;

namespace Selenium.Extensions
{
    public static class WebdriverExtensions
    {
        public class ElementNotFoundException : Exception
        {
            public ElementNotFoundException(By by)
            {
                throw new Exception($"**UI Element with locator - '{by}' doesn't exist on page!!**");
            }
        }
        public class IncorrectLandingPageException : Exception
        {
            public IncorrectLandingPageException(string str)
            {
                throw new Exception($"**{str}**!!");
            }
        }
        public static void WaitForAJAXComplete(this IWebDriver driver, bool RaiseTimeoutException = true)
        {
            DateTime initial = DateTime.Now;
            while (DateTime.Now.Subtract(initial).TotalSeconds < 120)
            {
                var ajaxIsComplete = (bool)(driver as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (ajaxIsComplete)
                {
                    var timeElapsed = DateTime.Now.Subtract(initial);
                    Console.WriteLine($"Time taken to load all widgets - {Math.Round(timeElapsed.TotalSeconds, 3)} secs");
                    return;
                }
            }
            if (RaiseTimeoutException)
            {
                throw new Exception("Widgets loading for more than a minute. WebDriver timed out.!!");
            }
        }
        public static bool IsValuePresentInDropdown(this IWebDriver driver, By by, string valueToChooseFromDropdown)
        {
            if (driver.IsElementPresent(by))
            {
                var elt = driver.FindElement(by);
                var selectElement = new SelectElement(elt);
                return selectElement.GetAllValuesFromDropdown().Any(s => s == valueToChooseFromDropdown);
            }
            else
                throw new Exception($"Element {by} not present!!!");
        }
        //Captures screenshot, assign the pagename to the screenshot image and save the image file to localstorage.
        /* If the url of the page is https://testsite.com/testpath/dashboard.aspx
         * A .png screenshot is stored on the local file storage with name - dashboard.png */
        public static void CaptureScreenshot(this IWebDriver driver, string textToAppend = "")
        {
            driver.WaitforPageLoad();
            var path = @"C:\Selenium\";
            var img = (driver as ITakesScreenshot).GetScreenshot();
            var fromPoint = driver.Url.LastIndexOf("/") + 1;
            string filename = driver.Url.Substring(fromPoint).RemoveInvalidCharacters() + (!textToAppend.IsNullOrEmpty() ? $" ({textToAppend})" : "");
            if (File.Exists($@"{path}\{filename}.png"))
            {
                var filenameTemp = filename;
                int i = 2;
                if (textToAppend.IsNullOrEmpty())
                {
                    while (File.Exists($@"{path}\{filenameTemp}.png"))
                    {
                        filenameTemp = filename + i.ToString();
                        i++;
                    }
                }
                else
                {
                    var temp = filenameTemp.Substring(0, filenameTemp.IndexOf("(")).Trim();
                    while (File.Exists($@"{path}\{filenameTemp}.png"))
                    {
                        filenameTemp = temp + i.ToString() + $" ({textToAppend})";
                        i++;
                    }
                }
                filename = filenameTemp;
            }
            img.SaveAsFile($@"{path}\{filename}.png", ScreenshotImageFormat.Png);
            Console.WriteLine($"   Screenshot Created - {filename}");
        }
        //Parse an excel file and return the total number of rows
        public static int ParseExcelFileAndReturnTotalRows(this IWebDriver driver, string folderPath, string file)
        {
            int rowsCount;
            using (var stream1 = File.Open($@"{folderPath}\{file}", FileMode.Open, FileAccess.Read))
            using (var reader1 = ExcelReaderFactory.CreateCsvReader(stream1))
            {
                var resultDatatable = reader1.AsDataSet().Tables[0];
                rowsCount = resultDatatable.Rows.Count;
            }
            if (rowsCount > 1)
            {
                Console.WriteLine($"Total Records found in excel- {rowsCount - 1}");
            }
            else
                throw new Exception("Exported file is empty");
            return rowsCount;
        }
        public static IWebElement FindElementWithScroll(this IWebDriver driver, By by)
        {
            var elementExists = driver.IsElementPresent(by);
            if (elementExists)
            {
                IWebElement elt = driver.FindElement(by);
                IJavaScriptExecutor je = (IJavaScriptExecutor)driver;
                je.ExecuteScript("return (document.readyState == 'complete')");
                je.ExecuteScript("arguments[0].scrollIntoView(true);", elt);
                je.ExecuteScript("arguments[0].style.outline = '2px solid #3DEF96'", elt);
                return elt;
            }
            else
                throw new ElementNotFoundException(by);
        }
        public static IEnumerable<IWebElement> FindElementAndHighlight(this IWebDriver driver, By by)
        {
            var elementExists = driver.IsElementPresent(by);
            if (elementExists)
            {
                var elts = driver.FindElements(by);
                var jsDriver = (IJavaScriptExecutor)driver;
                foreach (var elt in elts)
                {
                    jsDriver.ExecuteScript("arguments[0].style.outline = '2px solid #3DEF96'", elt);
                }
                return elts;
            }
            else
                throw new ElementNotFoundException(by);
        }
        public static void ClickElementJS(this IWebDriver driver, By by)
        {
            Thread.Sleep(500);
            var elt = driver.FindElement(by);
            var jsDriver = (IJavaScriptExecutor)driver;
            jsDriver.ExecuteScript("arguments[0].click();", elt);
            if (!driver.IsAlertPresent())
            {
                driver.WaitforPageLoad();
            }
        }
        public static void PressTab(this IWebDriver driver)
        {
            Actions act = new Actions(driver);
            act.SendKeys(Keys.Tab).Build().Perform();
        }
        public static void ClickElement(this IWebDriver driver, By by)
        {
            driver.FindElement(by).Click();
            if (!driver.IsAlertPresent())
            {
                driver.WaitforPageLoad();
            }
        }
        public static bool IsAlertPresent(this IWebDriver driver)
        {
            try
            {
                driver.SwitchTo().Alert();
                return true;
            }
            catch (NoAlertPresentException Ex)
            {
                return false;
            }
        }
        public static void WaitforPageLoad(this IWebDriver driver, bool RaiseTimeoutException = true)
        {
            DateTime initial = DateTime.Now;
            while (DateTime.Now.Subtract(initial).TotalSeconds < 60)
            {
                var pageloadIsComplete = (driver as IJavaScriptExecutor).ExecuteScript("return document.readyState").ToString().Equals("complete");
                if (pageloadIsComplete)
                    return;
            }
            if (RaiseTimeoutException)
            {
                throw new Exception("Page or Ajax Widgets loading for more than a minute. WebDriver timed out.!!");
            }
        }

        public static void EnterDate(this IWebDriver driver, By by, string date)
        {
            var elt = driver.FindElement(by);
            var jsDriver = (IJavaScriptExecutor)driver;
            jsDriver.ExecuteScript($"arguments[0].value = '{date}'", elt);
        }
        public static void EnterDate(this IWebDriver driver, string IDOfTheDatepicker, string date)
        {
            driver.FindElementAndHighlight(By.Id(IDOfTheDatepicker));
            var js = driver as IJavaScriptExecutor;
            js.ExecuteScript($"document.getElementById('{IDOfTheDatepicker}').value = '{date}'");
        }

        public static void OpenInNewTab(this IWebDriver driver, string url)
        {
            var js = driver as IJavaScriptExecutor;
            js.ExecuteScript("window.open();");
            driver.SwitchToLatestTab();
            driver.Url = url;
        }
        public static void SwitchToLatestTab(this IWebDriver driver)
        {
            var handlesCount = driver.WindowHandles.Count;
            if (handlesCount > 0)
                driver.SwitchTo().Window(driver.WindowHandles[handlesCount - 1]);
        }
        public static void CloseLatestTabAndSwitchToParentTab(this IWebDriver driver)
        {
            driver.Close();
            driver.SwitchTo().Window(driver.WindowHandles.Last());
        }
        public static bool IsElementPresent(this IWebDriver driver, By by, int timeoutInMilliSeconds = 6000)
        {
            driver.ImplicitWait(TimeSpan.FromMilliseconds(timeoutInMilliSeconds));
            return driver.FindElements(by).Count > 0;
        }
        public static string AcceptAlert(this IWebDriver driver)
        {
            var alert = driver.SwitchTo().Alert();
            var alertText = alert.Text;
            alert.Accept();
            Thread.Sleep(500);
            return alertText;
        }
        public static void ImplicitWait(this IWebDriver driver, TimeSpan t)
        {
            driver.Manage().Timeouts().ImplicitWait = t;
        }
        public static void PressEnter(this IWebDriver driver)
        {
            Actions action = new Actions(driver);
            action.SendKeys(Keys.Enter).Build().Perform();
        }
    }
}