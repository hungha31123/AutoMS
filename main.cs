using OpenQA.Selenium;
using System;
using System.Threading;
using AutoMS;

namespace AutoMS
{
    public class SearchHelper
    {
        public static void SearchPC(int time, string[] key, IWebDriver driver)
        {
            // Thực hiện tìm kiếm từng từ khoá
            foreach (var keyword in key)
            {
                string searchUrl = "https://www.bing.com/search?q=" + Uri.EscapeDataString(keyword);
                driver.Navigate().GoToUrl(searchUrl);

                // Thực hiện các hoạt động khác sau khi tìm kiếm
                Thread.Sleep(time);
            }

            driver.Navigate().GoToUrl("https://www.bing.com");
            Thread.Sleep(5000);

            try
            {
                IWebElement element = driver.FindElement(By.XPath("//span[@id='id_rc' and contains(@class, 'hp')]"));
                string value = element.Text;
                Console.WriteLine(value);
            }
            catch
            {
                return;
            }
        }


    }
}
