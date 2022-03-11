using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.Collections.Generic;
using System.Linq;

namespace Selenium.Extensions
{
    public static class WebElementExtensions
    {
        public static void SelectRandomValueFromDropdown(this IWebElement element)
        {
            int i = 0; var random = new Random();
            var elt = new SelectElement(element);
            var maxIndexOfDrodown = elt.Options.Count;
            if (maxIndexOfDrodown == 1)
                elt.SelectByIndex(0);
            else
                elt.SelectByIndex(random.Next(1, maxIndexOfDrodown));
            var selectedText = elt.SelectedOption.Text;
            while (selectedText.IsNullOrEmpty())
            {
                i++;
                elt.SelectByIndex(random.Next(1, maxIndexOfDrodown));
                selectedText = elt.SelectedOption.Text;
                if (i > 5)
                    throw new Exception($"**Nothing to select. Dropdown Empty!!!!");
            }
        }
        public static void SelectValueFromDropdown(this IWebElement element, string valueToSelect)
        {
            var selectElement = new SelectElement(element);
            if (selectElement.GetAllValuesFromDropdown().Any(s => s == valueToSelect))
                selectElement.SelectByValue(valueToSelect);
            else
                throw new Exception($"The value intended to select-'{valueToSelect}' is not contained in the dropdown");
        }
        public static void SelectFromDropdown(this IWebElement element, string optionToSelect)
        {
            var selectElement = new SelectElement(element);
            if (selectElement.GetAllValuesFromDropdown().Any(s => s.Contains(optionToSelect)))
                selectElement.SelectByValue(optionToSelect);
            else
                throw new Exception($"There's no option in the dropdown that matches-'{optionToSelect}'!!!***");
        }
        public static void SelectTextFromDropdown(this IWebElement element, string textToSelect)
        {
            var elt = new SelectElement(element);
            if (elt.GetAllTextFromDropdown().Any(s => s == textToSelect))
                elt.SelectByText(textToSelect);
            else
                throw new Exception($"The text intended to select-'{textToSelect}' is not contained in the dropdown");
        }
        public static string SelectedText(this IWebElement elt)
        {
            return new SelectElement(elt).SelectedOption.Text;
        }
        public static bool IsChecked(this IWebElement elt)
        {
            try
            {
                return elt.Selected;
            }
            catch (NoSuchElementException)
            {
                return false;
            }
        }
        public static IEnumerable<string> GetAllValuesFromDropdown(this SelectElement elt)
        {
            List<string> Values = new List<string>();
            foreach (var option in elt.Options)
            {
                Values.Add(option.GetValueForElement());
            }
            return Values;
        }
        static IEnumerable<string> GetAllTextFromDropdown(this SelectElement elt)
        {
            List<string> Texts = new List<string>();
            foreach (var option in elt.Options)
            {
                Texts.Add(option.Text);
            }
            return Texts;
        }
        public static long GetChildrenCount(this IWebElement element, IWebDriver driver)
        {
            var jsDriver = (IJavaScriptExecutor)driver;
            return (long)jsDriver.ExecuteScript("return arguments[0].children.length", element);
        }
        public static void EnterText(this IWebElement elt, string txt)
        {
            elt.Clear();
            elt.SendKeys(txt);
        }
        public static string GetValueForElement(this IWebElement elt)
        {
            return elt.GetAttribute("value");
        }
        public static string GetTextContent(this IWebElement elt)
        {
            return elt.GetAttribute("textContent");
        }
    }
}