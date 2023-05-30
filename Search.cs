using OpenQA.Selenium;
using System;
using System.Threading;
using AutoMS;
using System.Windows.Forms;
using OpenQA.Selenium.Chrome;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Button;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using System.IO;
using Newtonsoft.Json;
using System.Collections.Generic;
using System.Net.Http.Headers;
using System.Net.Http;
using System.Threading.Tasks;
using HtmlAgilityPack;
using OpenQA.Selenium.DevTools;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using static System.Net.Mime.MediaTypeNames;
using System.Linq;

namespace AutoMS
{
    public class SearchHelper
    {
        public void autoSearch(bool mbs, int time, string[] key, IWebDriver driver, DataGridView tb, DataGridViewCellEventArgs e)
        {
            DataGridViewRow slRow = tb.Rows[e.RowIndex];
            // Thực hiện tìm kiếm từng từ khoá
            int count = 1;
            foreach (var keyword in key)
            {
                string searchUrl = "https://www.bing.com/search?q=" + Uri.EscapeDataString(keyword);
                driver.Navigate().GoToUrl(searchUrl);
                // Thực hiện các hoạt động khác sau khi tìm kiếm
                Thread.Sleep(time);
                // Tìm thẻ span có ID "id_rc" và lấy nội dung của nó
                if (mbs == false)
                {
                    IWebElement spanElement = driver.FindElement(By.Id("id_rc"));
                    string pointsText = spanElement.Text;
                    slRow.Cells[3].Value = pointsText;
                    slRow.Cells[6].Value = $"Seach PC: {count}/{key.Count()}";
                }
                else
                {
                    slRow.Cells[6].Value = $"Seach Mobile: {count}/{key.Count()}";
                }
                //slRow.Cells[6].Value = $"Seach: {count}/{key.Count()}";
                count++;
            }
            driver.Navigate().GoToUrl("https://rewards.bing.com/");
        }

        public static string GetChromeProfilePath(string email)
        {
            // Đường dẫn đến thư mục chứa các profile trình duyệt đã lưu trước đó
            string userNamePc = Environment.UserName;
            string profilesDirPath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";

            // Đường dẫn đến thư mục của profile trình duyệt
            string profileDirPath = Path.Combine(profilesDirPath, email);

            return profileDirPath;
        }

        public ChromeOptions webPro(string email, int width, int height, TextBox notificationTextBox, bool useWindowSize)
        {
            string profileDirPath = GetChromeProfilePath(email);


            // Kiểm tra xem profile trình duyệt có tồn tại hay không
            if (!Directory.Exists(profileDirPath))
            {
                MessageBox.Show("Profile not found!");
                return null;
            }
            else
            {
                //notificationTextBox.Text += "\r\n" + email + "--> Open...";
                // Tạo cấu hình trình duyệt Chrome với profile trình duyệt đã có sẵn
                ChromeOptions options = new ChromeOptions();
                if (useWindowSize)
                {
                    options.AddArgument($"--window-size={width},{height}");
                }

                options.AddArgument("user-data-dir=" + profileDirPath);
                options.AddArgument("--start-maximized");

                return options;
            }
        }


        public async Task<string> sPc(int ws, string email, int width, int height, TextBox notificationTextBox, bool useWindowSize, bool mb, DataGridView dt, DataGridViewCellEventArgs e)
        {
            ChromeOptions options = webPro(email, width, height, notificationTextBox, useWindowSize);
            // Tạo cấu hình cho ChromeDriverService để ẩn cửa sổ cmd
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            // Khởi tạo trình duyệt Chrome với cấu hình đã tạo
            IWebDriver driver = new ChromeDriver(service, options);
            List<string> wPC = await this.GetRandomWords(ws);
            this.autoSearch(false,3000, wPC.ToArray(), driver, dt, e);
            string points = this.GetPoints(driver, mb);
            
            return points;
        }


        public async Task<string> sMb(int ws, string email, int width, int height, TextBox notificationTextBox, bool useWindowSize, DataGridView dt, DataGridViewCellEventArgs e)
        {
            bool mb = true;
            ChromeOptions options = webPro(email, width, height, notificationTextBox, useWindowSize);
            options.AddArgument("--user-agent=Mozilla/5.0 (iPhone; CPU iPhone OS 14_0 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0 Mobile/15E148 Safari/604.1");
            // Tạo cấu hình cho ChromeDriverService để ẩn cửa sổ cmd
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            // Khởi tạo trình duyệt Chrome với cấu hình đã tạo
            IWebDriver driver = new ChromeDriver(service, options);
            List<string> wPC = await this.GetRandomWords(ws);
            this.autoSearch(true, 3000, wPC.ToArray(), driver, dt, e);
            string points = this.GetPoints(driver, mb);
            driver.Quit();
            return points;

            
        }

        public async void Auto(int wsPc, int wsMb, string email, int width, int height, TextBox notificationTextBox, bool useWindowSize, bool Mb, DataGridView dt, DataGridViewCellEventArgs e)
        {
            DataGridViewRow slRow = dt.Rows[e.RowIndex];
            string acc = email;
            int wth = width;
            int hght = height;
            TextBox notiTextBox = notificationTextBox;
            bool useSize = useWindowSize;

            await Task.Run(async () =>
            {
                // Thực hiện tìm kiếm trên Mobile
                if (Mb)
                {
                    Thread.Sleep(5000);
                    await sMb(wsMb, email, width, height, notificationTextBox, useWindowSize, dt, e);
                    //slRow.Cells[3].Value = point2;
                }
                // Thực hiện tìm kiếm trên PC
                string point1 = await sPc(wsPc, email, width, height, notificationTextBox, useWindowSize, Mb, dt, e);
                slRow.Cells[3].Value = point1;
                

            });
        }

        public string GetPoints(IWebDriver driver, bool mb)
        {
            try
            {
                if (mb == true)
                {
                    return null;
                }
                driver.Navigate().GoToUrl("https://www.bing.com/search?q=FIFA");
                Thread.Sleep(5000);

                // Tìm thẻ span có ID "id_rc" và lấy nội dung của nó
                IWebElement spanElement = driver.FindElement(By.Id("id_rc"));
                string pointsText = spanElement.Text;

                driver.Navigate().GoToUrl("https://rewards.bing.com/");
                // Trả về nội dung dưới dạng chuỗi
                return pointsText;


            }
            catch
            {
                MessageBox.Show("Lỗi...");
                return null;
            }
        }
        public async Task<List<string>> GetRandomWords(int numberOfWords)
        {
            using (var client = new HttpClient())
            {
                string apiUrl = $"https://random-word.ryanrk.com/api/en/word/random/{numberOfWords}";
                client.BaseAddress = new Uri(apiUrl);
                client.DefaultRequestHeaders.Accept.Clear();
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

                HttpResponseMessage response = await client.GetAsync(apiUrl);

                if (response.IsSuccessStatusCode)
                {
                    string jsonContent = await response.Content.ReadAsStringAsync();
                    List<string> randomWords = JsonConvert.DeserializeObject<List<string>>(jsonContent);
                    return randomWords;
                }
                else
                {
                    return null;
                }
            }
        }

        public class DashboardData
        {
            public Dictionary<string, object> UserStatus { get; set; }
        }



        public static Dictionary<string, object> GetDashboardData(IWebDriver browser)
        {
            var dashboard = FindBetween(browser.FindElement(By.XPath("/html/body")).GetAttribute("innerHTML"), "var dashboard = ", ";\n        appDataModule.constant(\"prefetchedDashboard\", dashboard);");
            var dashboardData = Newtonsoft.Json.JsonConvert.DeserializeObject<Dictionary<string, object>>(dashboard);
            return dashboardData;
        }

        public static int GetAccountPoints(IWebDriver browser)
        {
            var dashboardData = GetDashboardData(browser);
            var userStatus = (Dictionary<string, object>)dashboardData["userStatus"];
            return Convert.ToInt32(userStatus["availablePoints"]);
        }

        public static string FindBetween(string s, string first, string last)
        {
            try
            {
                var start = s.IndexOf(first) + first.Length;
                var end = s.IndexOf(last, start);
                return s.Substring(start, end - start);
            }
            catch (Exception)
            {
                return "";
            }
        }

        public void point(string email, int width, int height, TextBox notificationTextBox, bool useWindowSize)
        {
            ChromeOptions options = webPro(email, width, height, notificationTextBox, useWindowSize);
            // Tạo cấu hình cho ChromeDriverService để ẩn cửa sổ cmd
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            // Khởi tạo trình duyệt Chrome với cấu hình đã tạo
            IWebDriver driver = new ChromeDriver(service, options);
            driver.Navigate().GoToUrl("https://rewards.bing.com/");
            Thread.Sleep(5000);
            // Truyền browser vào hàm để sử dụng
            int accountPoints = GetAccountPoints(driver);
            // Tiếp tục xử lý với accountPoints...
            MessageBox.Show(accountPoints.ToString());
        }
    }
}
