using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Diagnostics;
using System.IO;
using System.Security.Principal;
using OpenQA.Selenium.Remote;
using System.Runtime.InteropServices;
using static System.Net.Mime.MediaTypeNames;
using System.Data.SqlClient;
using OpenQA.Selenium.Support.UI;
using Newtonsoft.Json;
using OpenQA.Selenium.DevTools;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using AutoMS;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using OpenQA.Selenium.DevTools.V113.FedCm;

namespace AutoMS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private CancellationTokenSource cancellationTokenSource = null; // Đối tượng CancellationTokenSource

        SearchHelper searchHelper = new SearchHelper();

        private async void button1_Click(object sender, EventArgs e)
        {
            //try
            //{
            //    Thread thread = new Thread(() =>
            //    {
            //        RunProcess(txtAcc.Text, txtPass.Text, txtNPass.Text, noimage.Checked, headless.Checked);
            //    });
            //    thread.Start();
            //}
            //catch
            //{

            //    int numThreads = (int)numThread.Value; // Số luồng bằng số dòng trong DataGridView
            //    List<Thread> threads = new List<Thread>();

            //    for (int i = 0; i < numThreads; i++)
            //    {
            //        if (dataGridView1.Rows[i].Cells["dtAcc"].Value != null &&
            //            dataGridView1.Rows[i].Cells["dtPass"].Value != null)
            //        {
            //            string acc = dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString();
            //            string pass = dataGridView1.Rows[i].Cells["dtPass"].Value.ToString();

            //            Thread thread = new Thread(() =>
            //            {
            //                RunProcess(acc, pass, txtNPass.Text, noimage.Checked, headless.Checked);
            //            });

            //            thread.Start();
            //        }
            //    }
            //}

            bool hasDataGridViewData = (dataGridView1.Rows.Count > 1);
            if (hasDataGridViewData)
            {
                cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                int numRow = (int)numThread.Value;
                for (int i = 0; i < numRow && i < dataGridView1.Rows.Count; i++)
                {
                    // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                    if (cancellationToken.IsCancellationRequested)
                    {
                        break;
                    }
                    DataGridViewRow row = dataGridView1.Rows[i];
                    if (row.Cells["dtAcc"].Value != null && row.Cells["dtPass"].Value != null)
                    {
                        string acc = row.Cells["dtAcc"].Value.ToString();
                        string pass = row.Cells["dtPass"].Value.ToString();
                        await ProcessDataAsync(acc, pass);
                    }
                }
                MessageBox.Show("Success!");


            }
            else if (!string.IsNullOrEmpty(txtAcc.Text) && !string.IsNullOrEmpty(txtPass.Text))
            {
                Thread thread = new Thread(() =>
                {
                    RunProcess(txtAcc.Text, txtPass.Text, txtNPass.Text, noimage.Checked, headless.Checked);
                });
                thread.Start();
                MessageBox.Show("Success!");
            }
            else
            {
                MessageBox.Show("Vui lòng nhập thông tin tài khoản và mật khẩu", "Lỗi", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }




        }

        private async Task ProcessDataAsync(string acc, string pass)
        {
            await Task.Run(() =>
            {
                RunProcess(acc, pass, txtNPass.Text, noimage.Checked, headless.Checked);
            });
        }

        private async Task CreateProfileDataAsync(string acc, string pass)
        {
            await Task.Run(() =>
            {
                CreatProfile(acc, pass, noimage.Checked, headless.Checked);
            });
        }

        private async Task OpenProfileDataAsync(string acc, int width, int height)
        {
            await Task.Run(() =>
            {
                OpenProifle(acc, (int)numericUpDown1.Value, (int)numericUpDown2.Value);
            });
        }

        private void RunProcess(string Acc, string Pass, string NPass, bool noImage, bool HL)
        {
            Invoke(new Action(() =>
            {
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc)
                    {
                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Running...";
                        break;
                    }
                }
            }));
            string email = Acc;
            int atIndex = email.IndexOf('@');
            string username = email.Substring(0, atIndex).ToLower();
            string numbers = new string(username.Where(char.IsDigit).ToArray());

            string pass = Pass;
            string newPass = NPass + numbers;

            // Tạo cấu hình cho ChromeDriverService để ẩn cửa sổ cmd
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;

            // Tạo một đối tượng ChromeDriver để khởi tạo trình duyệt Chrome
            ChromeOptions options = new ChromeOptions();

            if (noImage)
            {
                options.AddArgument("--blink-settings=imagesEnabled=false");
            }
            if (HL)
            {
                options.AddArgument("--headless"); // Khởi tạo trình duyệt headless
                options.AddArgument("--disable-gpu"); // thêm tùy chọn này để tránh lỗi khi chạy headless
            }
            options.AddArgument("--silent");
            options.AddArgument("--disable-notifications");
            IWebDriver driver = new ChromeDriver(service, options);
            // Navigate đến trang đăng nhập Microsoft
            driver.Navigate().GoToUrl("https://login.live.com/");
            // Lưu ID của tab hiện tại
            string currentTabId = driver.CurrentWindowHandle;


            // Tìm và điền email vào ô nhập liệu Email
            IWebElement emailInput = driver.FindElement(By.Name("loginfmt"));
            emailInput.SendKeys(email + "\r\n");
            Thread.Sleep(3000);
            // Nhấp vào nút Next
            //IWebElement nextButton = driver.FindElement(By.Id("idSIButton9"));
            //nextButton.Click();

            // Tìm và điền password vào ô nhập liệu Password
            IWebElement passwordInput = driver.FindElement(By.Name("passwd"));
            passwordInput.SendKeys(pass);
            Thread.Sleep(2000);
            // Nhấp vào nút Sign In
            IWebElement signInButton = driver.FindElement(By.Id("idSIButton9"));
            signInButton.Click();
            Thread.Sleep(2000);
            try
            {

                if (driver.FindElement(By.Id("passwordError")) != null)
                {
                    driver.Quit();
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Sai Pass"; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Sai Pass";
                                break;
                            }
                        }
                    }));
                    return;


                }
            }
            catch
            {
                try
                {
                    if (driver.FindElement(By.Id("StartHeader")) != null) //thông báo tình trạng acc
                    {
                        IWebElement notification = driver.FindElement(By.Id("StartHeader"));
                        string message = notification.Text;
                        Invoke(new Action(() => { txtNoti.Text = message; }));
                        driver.Quit();
                        return;
                    }
                }
                catch
                {
                    try
                    {
                        if (driver.FindElement(By.Id("iProofEmail")) != null)   //checkpoint security
                        {
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Checkpoint Security"; }));
                            Invoke(new Action(() =>
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc)
                                    {
                                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Checkpoint Security";
                                        break;
                                    }
                                }
                            }));
                            IWebElement unlockmail = driver.FindElement(By.Id("iProofEmail"));
                            unlockmail.SendKeys(username);
                            Thread.Sleep(1000);
                            IWebElement unlocksubmit = driver.FindElement(By.Id("iSelectProofAction"));
                            unlocksubmit.Click();
                            Thread.Sleep(7000);
                            // Mở tab mới và chuyển đổi sang tab mới
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Go to Mailnesia..."; }));
                            Invoke(new Action(() =>
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                        dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                                    {
                                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Go to Mailnesia...";

                                        break;
                                    }
                                }
                            }));
                            IJavaScriptExecutor jsunlock = (IJavaScriptExecutor)driver;
                            jsunlock.ExecuteScript("window.open();");
                            driver.SwitchTo().Window(driver.WindowHandles.Last());

                            // Truy cập trang web mailnesia.com với tên người dùng là alfonzonienow238
                            driver.Navigate().GoToUrl("https://mailnesia.com/mailbox/" + username);
                            Thread.Sleep(3000);
                            // Nhấp vào liên kết email đầu tiên trong danh sách
                            // Tìm phần tử <a> cần nhấn vào
                            IWebElement mailunlock = driver.FindElement(By.CssSelector("a.email[title='Mở thư'][href*='/mailbox/']"));
                            // Nhấp vào liên kết
                            mailunlock.Click();
                            Thread.Sleep(5000);
                            //lấy code
                            // Tìm phần tử <td> chứa mã security code
                            IWebElement lockmail = driver.FindElement(By.Id("i4"));

                            // Lấy toàn bộ nội dung của phần tử <td>
                            string unlmail = lockmail.Text;

                            // Tìm và lấy giá trị security code
                            string unlcode = unlmail.Substring(unlmail.IndexOf("Security code: ") + 15, 7);
                            //txtNoti.Text = unlcode;
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Code: " + unlcode; }));
                            Invoke(new Action(() =>
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc)
                                    {
                                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Code: " + unlcode;
                                        break;
                                    }
                                }
                            }));

                            driver.SwitchTo().Window(currentTabId);

                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Nhập code..."; }));
                            Invoke(new Action(() =>
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                        dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                                    {
                                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Nhập code...";

                                        break;
                                    }
                                }
                            }));
                            IWebElement nhapcode = driver.FindElement(By.Id("iOttText"));
                            nhapcode.SendKeys(unlcode);
                            Thread.Sleep(2000);
                            IWebElement subcode = driver.FindElement(By.Id("iVerifyCodeAction"));
                            subcode.Click();
                            Thread.Sleep(5000);
                            try
                            {
                                Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công, check..."; }));
                                Invoke(new Action(() =>
                                {
                                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                    {
                                        if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                            dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                                        {
                                            dataGridView1.Rows[i].Cells["dtNote"].Value = "Login thành công, check...";

                                            break;
                                        }
                                    }
                                }));
                                Thread.Sleep(5000);
                                IWebElement ctnbtn = driver.FindElement(By.Id("id__0"));
                                ctnbtn.Click();
                                Thread.Sleep(3000);
                                IWebElement noButton = driver.FindElement(By.Id("idBtn_Back"));
                                noButton.Click();
                            }
                            catch
                            {
                                Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công..."; }));
                                Invoke(new Action(() =>
                                {
                                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                    {
                                        if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                            dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                                        {
                                            dataGridView1.Rows[i].Cells["dtNote"].Value = "Login thành công...";

                                            break;
                                        }
                                    }
                                }));
                                //MessageBox.Show("Đăng nhập thành công!");
                                IWebElement noButton = driver.FindElement(By.Id("idBtn_Back"));
                                noButton.Click();
                            }

                        }
                    }
                    catch
                    {
                        try
                        {
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công, check..."; }));
                            Invoke(new Action(() =>
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                        dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                                    {
                                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Login thành công, check...";

                                        break;
                                    }
                                }
                            }));
                            Thread.Sleep(5000);
                            IWebElement ctnbtn = driver.FindElement(By.Id("id__0"));
                            ctnbtn.Click();
                            Thread.Sleep(3000);
                            IWebElement noButton = driver.FindElement(By.Id("idBtn_Back"));
                            noButton.Click();
                        }
                        catch
                        {
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công..."; }));
                            Invoke(new Action(() =>
                            {
                                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                                {
                                    if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                        dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                                    {
                                        dataGridView1.Rows[i].Cells["dtNote"].Value = "Login thành công...";

                                        break;
                                    }
                                }
                            }));
                            //MessageBox.Show("Đăng nhập thành công!");
                            IWebElement noButton = driver.FindElement(By.Id("idBtn_Back"));
                            noButton.Click();
                        }

                    }
                }
                Thread.Sleep(15000);
                try
                {
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Trang quản lý..."; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Trang quản lý...";

                                break;
                            }
                        }
                    }));
                    IWebElement changePass = driver.FindElement(By.Id("home.banner.change-password-column.cta"));
                    changePass.Click();
                    Thread.Sleep(5000);
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Recover Mail..."; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Recover Mail...";

                                break;
                            }
                        }
                    }));
                    IWebElement clMail = driver.FindElement(By.Id("idDiv_SAOTCS_Proofs_Section"));
                    clMail.Click();
                    Thread.Sleep(2000);

                    IWebElement mailrc = driver.FindElement(By.Id("idTxtBx_SAOTCS_ProofConfirmation"));
                    IWebElement submitrc = driver.FindElement(By.Id("idSubmit_SAOTCS_SendCode"));

                    mailrc.SendKeys(username + "@mailnesia.com");
                    Thread.Sleep(1000);
                    submitrc.Click();
                    Thread.Sleep(7000);

                    // Mở tab mới và chuyển đổi sang tab mới
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Lấy code..."; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Lấy code...";

                                break;
                            }
                        }
                    }));
                    IJavaScriptExecutor js = (IJavaScriptExecutor)driver;
                    js.ExecuteScript("window.open();");
                    driver.SwitchTo().Window(driver.WindowHandles.Last());

                    // Truy cập trang web mailnesia.com với tên người dùng là alfonzonienow238
                    driver.Navigate().GoToUrl("https://mailnesia.com/mailbox/" + username);
                    Thread.Sleep(3000);
                    // Nhấp vào liên kết email đầu tiên trong danh sách
                    // Tìm phần tử <a> cần nhấn vào
                    IWebElement emailLink = driver.FindElement(By.CssSelector("a.email[title='Mở thư'][href*='/mailbox/']"));
                    // Nhấp vào liên kết
                    emailLink.Click();
                    Thread.Sleep(5000);
                    //lấy code
                    // Tìm phần tử <td> chứa mã security code
                    IWebElement td = driver.FindElement(By.Id("i4"));

                    // Lấy toàn bộ nội dung của phần tử <td>
                    string tdText = td.Text;

                    // Tìm và lấy giá trị security code
                    string securityCode = tdText.Substring(tdText.IndexOf("Security code: ") + 15, 7);
                    //txtNoti.Text = securityCode;
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Code: " + securityCode; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Code: " + securityCode;
                                break;
                            }
                        }
                    }));
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Đổi pass..."; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Đổi pass...";

                                break;
                            }
                        }
                    }));
                    driver.SwitchTo().Window(currentTabId);
                    IWebElement inputOTP = driver.FindElement(By.Id("idTxtBx_SAOTCC_OTC"));
                    inputOTP.SendKeys(securityCode);
                    Thread.Sleep(2000);
                    IWebElement submitOTP = driver.FindElement(By.Id("idSubmit_SAOTCC_Continue"));
                    submitOTP.Click();
                    Thread.Sleep(2000);
                    IWebElement curPass = driver.FindElement(By.Id("iCurPassword"));
                    curPass.SendKeys(pass);
                    Thread.Sleep(1000);
                    IWebElement chgPass1 = driver.FindElement(By.Id("iPassword"));
                    chgPass1.SendKeys(newPass);
                    Thread.Sleep(1000);
                    IWebElement chgPass2 = driver.FindElement(By.Id("iRetypePassword"));
                    chgPass2.SendKeys(newPass);
                    Thread.Sleep(1000);
                    IWebElement updatePassBtn = driver.FindElement(By.Id("UpdatePasswordAction"));
                    updatePassBtn.Click();
                    Thread.Sleep(10000);
                    driver.Quit();

                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Success!"; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                            {
                                dataGridView1.Rows[i].Cells["dtNPass"].Value = newPass;
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Success!";

                                break;
                            }
                        }
                    }));
                    return;
                }
                catch
                {
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Có lỗi..."; }));
                    Invoke(new Action(() =>
                    {
                        for (int i = 0; i < dataGridView1.Rows.Count; i++)
                        {
                            if (dataGridView1.Rows[i].Cells["dtAcc"].Value.ToString() == Acc &&
                                dataGridView1.Rows[i].Cells["dtPass"].Value.ToString() == pass)
                            {
                                dataGridView1.Rows[i].Cells["dtNote"].Value = "Có lỗi...";
                                break;
                            }
                        }
                    }));
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage1)
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                openFileDialog1.FilterIndex = 0;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;
                    // Đọc dữ liệu từ tệp tin và gán cho DataGridView
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        // Kiểm tra xem dòng có chứa toàn bộ ký tự '|' hay không
                        if (!string.IsNullOrWhiteSpace(line) && !line.All(c => c == '|'))
                        {
                            string[] rowValues = line.Split('|');

                            // Thêm dòng vào DataGridView
                            int rowIndex = dataGridView3.Rows.Add();

                            // Gán giá trị từ mảng rowValues vào các ô của dòng tương ứng
                            for (int i = 0; i < rowValues.Length; i++)
                            {
                                string value = rowValues[i];

                                // Kiểm tra nếu là ô CheckBox
                                if (dataGridView3.Columns[i] is DataGridViewCheckBoxColumn)
                                {
                                    bool checkBoxValue = value == "1";
                                    dataGridView3.Rows[rowIndex].Cells[i].Value = checkBoxValue;
                                }
                                else
                                {
                                    dataGridView3.Rows[rowIndex].Cells[i].Value = value;
                                }
                            }
                        }
                    }

                }


            }
            else if (tabControl1.SelectedTab == tabPage2)
            {

                OpenFileDialog openFileDialog1 = new OpenFileDialog();
                openFileDialog1.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
                openFileDialog1.FilterIndex = 1;

                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog1.FileName;
                    string[] lines = File.ReadAllLines(filePath);
                    foreach (string line in lines)
                    {
                        string[] data = line.Split('|');
                        bool exists = false;
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == data[0])
                            {
                                exists = true;
                                break;
                            }
                        }
                        if (!exists)
                        {
                            try
                            {
                                dataGridView1.Rows.Add(dataGridView1.Rows.Count, data[0], data[1], data[2]);
                            }
                            catch
                            {
                                dataGridView1.Rows.Add(dataGridView1.Rows.Count, data[0], data[1]);
                            }
                        }
                    }
                }
            }


        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = this.dataGridView1.Rows[e.RowIndex];
                if (row.Cells["dtAcc"].Value != null && row.Cells["dtPass"].Value != null)
                {
                    // Lấy giá trị của các ô trong hàng được chọn
                    string account = row.Cells["dtAcc"].Value.ToString();
                    string password = row.Cells["dtPass"].Value.ToString();

                    // Hiển thị giá trị vào các textbox tương ứng
                    txtAcc.Text = account;
                    txtPass.Text = password;
                }
                
            }
        }

        private void dataGridView1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (Path.GetExtension(files[0]).Equals(".txt", StringComparison.OrdinalIgnoreCase))
                {
                    e.Effect = DragDropEffects.Copy;
                    return;
                }
            }
            e.Effect = DragDropEffects.None;
        }

        private void dataGridView1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            string filePath = files[0];
            string[] lines = File.ReadAllLines(filePath);
            foreach (string line in lines)
            {
                string[] data = line.Split('|');
                bool exists = false;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[1].Value != null && row.Cells[1].Value.ToString() == data[0])
                    {
                        exists = true;
                        break;
                    }
                }
                if (!exists)
                {
                    try
                    {
                        dataGridView1.Rows.Add(dataGridView1.Rows.Count, data[0], data[1], data[2]);
                    }
                    catch
                    {
                        dataGridView1.Rows.Add(dataGridView1.Rows.Count, data[0], data[1]);
                    }
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            CreatProfile(txtAcc.Text, txtPass.Text, noimage.Checked, headless.Checked);
            //UpdateDataGridView2();
        }

        public void CreatProfile(string Acc, string Pass, bool noImage, bool HL)
        {
            string email = Acc;
            string pass = Pass;
            int atIndex = email.IndexOf('@');
            // Kiểm tra xem chuỗi email có chứa ký tự '@' hay không
            if (atIndex == -1 || pass == "")
            {
                MessageBox.Show("Error: Invalid email/pass");

                return;
            }
            string username = email.Substring(0, atIndex);

            
            // Tạo cấu hình cho ChromeDriverService để ẩn cửa sổ cmd
            ChromeDriverService service = ChromeDriverService.CreateDefaultService();
            service.HideCommandPromptWindow = true;
            // Khởi tạo trình duyệt Chrome với cấu hình đã tạo
            ChromeOptions options = new ChromeOptions();

            // Đường dẫn đến thư mục chứa các profile trình duyệt đã lưu trước đó
            // Lấy tên người dùng hiện tại của máy tính
            string userNamePc = Environment.UserName;

            // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
            string profilesDirPath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";

            // Tên đăng nhập của bạn

            // Đường dẫn đến thư mục của profile trình duyệt mới
            string profileDirPath = Path.Combine(profilesDirPath, email);

            // Tạo thư mục profile trình duyệt mới nếu nó chưa tồn tại
            if (!Directory.Exists(profileDirPath))
            {
                Directory.CreateDirectory(profileDirPath);
            }
            else
            {
                Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Đã có Profile"; }));
                OpenProifle(email, (int)numericUpDown1.Value, (int)numericUpDown2.Value);
                return;
            }
            options.AddArgument("user-data-dir=" + profileDirPath);
            options.AddArgument("--start-maximized");


            if (noImage)
            {
                options.AddArgument("--blink-settings=imagesEnabled=false");
            }
            if (HL)
            {
                options.AddArgument("--headless"); // Khởi tạo trình duyệt headless
                options.AddArgument("--disable-gpu"); // thêm tùy chọn này để tránh lỗi khi chạy headless
            }

            options.AddArgument("--silent");
            options.AddArgument("--disable-notifications");
            IWebDriver driver = new ChromeDriver(service, options);
            // Navigate đến trang đăng nhập Microsoft
            driver.Navigate().GoToUrl("https://login.live.com/");
            // Lưu ID của tab hiện tại
            string currentTabId = driver.CurrentWindowHandle;


            // Tìm và điền email vào ô nhập liệu Email
            IWebElement emailInput = driver.FindElement(By.Name("loginfmt"));
            emailInput.SendKeys(email + "\r\n");

            // Nhấp vào nút Next
            //IWebElement nextButton = driver.FindElement(By.Id("idSIButton9"));
            //nextButton.Click();

            // Tìm và điền password vào ô nhập liệu Password
            IWebElement passwordInput = driver.FindElement(By.Name("passwd"));
            passwordInput.SendKeys(pass + "\n");
            Thread.Sleep(1000);
            // Nhấp vào nút Sign In
            IWebElement signInButton = driver.FindElement(By.Id("idSIButton9"));
            signInButton.Click();
            try
            {

                if (driver.FindElement(By.Id("passwordError")) != null)
                {
                    driver.Quit();
                    Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Sai Pass"; }));
                    return;


                }
            }
            catch
            {
                try
                {
                    if (driver.FindElement(By.Id("StartHeader")) != null) //thông báo tình trạng acc
                    {
                        IWebElement notification = driver.FindElement(By.Id("StartHeader"));
                        string message = notification.Text;
                        Invoke(new Action(() => { txtNoti.Text = message; }));
                        driver.Quit();
                        return;
                    }
                }
                catch
                {
                    try
                    {
                        if (driver.FindElement(By.Id("iProofEmail")) != null)   //checkpoint security
                        {
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Checkpoint Security"; }));
                            IWebElement unlockmail = driver.FindElement(By.Id("iProofEmail"));
                            unlockmail.SendKeys(username);
                            Thread.Sleep(1000);
                            IWebElement unlocksubmit = driver.FindElement(By.Id("iSelectProofAction"));
                            unlocksubmit.Click();
                            Thread.Sleep(7000);
                            // Mở tab mới và chuyển đổi sang tab mới
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Go to Mailnesia..."; }));
                            IJavaScriptExecutor jsunlock = (IJavaScriptExecutor)driver;
                            jsunlock.ExecuteScript("window.open();");
                            driver.SwitchTo().Window(driver.WindowHandles.Last());

                            // Truy cập trang web mailnesia.com với tên người dùng là alfonzonienow238
                            driver.Navigate().GoToUrl("https://mailnesia.com/mailbox/" + username);
                            Thread.Sleep(3000);
                            // Nhấp vào liên kết email đầu tiên trong danh sách
                            // Tìm phần tử <a> cần nhấn vào
                            IWebElement mailunlock = driver.FindElement(By.CssSelector("a.email[title='Mở thư'][href*='/mailbox/']"));
                            // Nhấp vào liên kết
                            mailunlock.Click();
                            Thread.Sleep(5000);
                            //lấy code
                            // Tìm phần tử <td> chứa mã security code
                            IWebElement lockmail = driver.FindElement(By.Id("i4"));

                            // Lấy toàn bộ nội dung của phần tử <td>
                            string unlmail = lockmail.Text;

                            // Tìm và lấy giá trị security code
                            string unlcode = unlmail.Substring(unlmail.IndexOf("Security code: ") + 15, 7);
                            //txtNoti.Text = unlcode;
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Code: " + unlcode; }));
                            driver.SwitchTo().Window(currentTabId);

                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Nhập code..."; }));
                            IWebElement nhapcode = driver.FindElement(By.Id("iOttText"));
                            nhapcode.SendKeys(unlcode);
                            Thread.Sleep(2000);
                            IWebElement subcode = driver.FindElement(By.Id("iVerifyCodeAction"));
                            subcode.Click();
                            Thread.Sleep(7000);
                            try
                            {
                                Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công, check..."; }));
                                Thread.Sleep(5000);
                                IWebElement ctnbtn = driver.FindElement(By.Id("id__0"));
                                ctnbtn.Click();
                                Thread.Sleep(3000);
                                IWebElement noButton = driver.FindElement(By.Id("idSIButton9"));
                                noButton.Click();
                            }
                            catch
                            {
                                Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công..."; }));
                                //MessageBox.Show("Đăng nhập thành công!");
                                IWebElement noButton = driver.FindElement(By.Id("idSIButton9"));
                                noButton.Click();
                            }

                        }
                    }
                    catch
                    {
                        try
                        {
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công, check..."; }));
                            Thread.Sleep(5000);
                            IWebElement ctnbtn = driver.FindElement(By.Id("id__0"));
                            ctnbtn.Click();
                            Thread.Sleep(3000);
                            IWebElement noButton = driver.FindElement(By.Id("idSIButton9"));
                            noButton.Click();
                        }
                        catch
                        {
                            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Login thành công..."; }));
                            //MessageBox.Show("Đăng nhập thành công!");
                            IWebElement noButton = driver.FindElement(By.Id("idSIButton9"));
                            noButton.Click();
                        }

                    }
                }
            }
            driver.Quit();
            Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Profile success!"; }));
            UpdateDataGridView2();
            return;
        }



        public void OpenProifle(string email, int width, int height)
        {
            // Đường dẫn đến thư mục chứa các profile trình duyệt đã lưu trước đó
            string userNamePc = Environment.UserName;

            // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
            string profilesDirPath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";

            // Đường dẫn đến thư mục của profile trình duyệt
            string profileDirPath = Path.Combine(profilesDirPath, email);

            // Kiểm tra xem profile trình duyệt có tồn tại hay không
            if (!Directory.Exists(profileDirPath))
            {
                MessageBox.Show("Profile not found!");
                return;
            }
            else
            {
                Invoke(new Action(() => { txtNoti.Text += "\r\n" + email + "--> Open..."; }));
                // Tạo cấu hình cho ChromeDriverService để ẩn cửa sổ cmd
                ChromeDriverService service = ChromeDriverService.CreateDefaultService();
                service.HideCommandPromptWindow = true;
                // Tạo cấu hình trình duyệt Chrome với profile trình duyệt đã có sẵn
                ChromeOptions options = new ChromeOptions();
                if (checkBox3.Checked)
                {
                    options.AddArgument($"--window-size={width},{height}");

                }

                options.AddArgument("user-data-dir=" + profileDirPath);
                options.AddArgument("--start-maximized");

                // Khởi tạo trình duyệt Chrome với cấu hình đã tạo
                IWebDriver driver = new ChromeDriver(service, options);

                //searchHelper.GetPoints(driver);


                // Điều hướng đến trang web mặc định của bạn
                //driver.Navigate().GoToUrl("https://rewards.bing.com/");

            }
        }



        private void headless_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (tabControl1.SelectedTab == tabPage1)
            {
                // Tạo OpenFileDialog để chọn nơi lưu file
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "Text Files (*.txt)|*.txt";
                saveFileDialog.Title = "Export to Text File";
                saveFileDialog.ShowDialog();

                if (saveFileDialog.FileName != "")
                {
                    try
                    {
                        // Lưu dữ liệu từ DataGridView vào tệp tin
                        using (StreamWriter writer = new StreamWriter(dataFilePath))
                        {
                            foreach (DataGridViewRow row in dataGridView3.Rows)
                            {
                                string[] values = new string[row.Cells.Count];
                                for (int i = 0; i < row.Cells.Count; i++)
                                {
                                    if (row.Cells[i] is DataGridViewCheckBoxCell)
                                    {
                                        DataGridViewCheckBoxCell checkBoxCell = row.Cells[i] as DataGridViewCheckBoxCell;
                                        bool? checkBoxValue = checkBoxCell?.Value as bool?;
                                        values[i] = checkBoxValue.HasValue ? (checkBoxValue.Value ? "1" : "0") : "";
                                    }
                                    else if (row.Cells[i] is DataGridViewComboBoxCell)
                                    {
                                        DataGridViewComboBoxCell comboBoxCell = (DataGridViewComboBoxCell)row.Cells[i];
                                        string comboBoxValue = comboBoxCell.Value?.ToString() ?? "";
                                        values[i] = comboBoxValue;
                                    }
                                    else
                                    {
                                        values[i] = row.Cells[i].Value?.ToString() ?? "";
                                    }
                                }
                                writer.WriteLine(string.Join("|", values));
                            }
                        }

                        //MessageBox.Show("Export successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            else if (tabControl1.SelectedTab == tabPage2)
            {

            // Tạo OpenFileDialog để chọn nơi lưu file
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt";
            saveFileDialog.Title = "Export to Text File";
            saveFileDialog.ShowDialog();

            if (saveFileDialog.FileName != "")
            {
                try
                {
                    // Tạo StreamWriter để ghi dữ liệu vào file
                    using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName))
                    {
                        foreach (DataGridViewRow row in dataGridView1.Rows)
                        {
                            if (row.Cells["dtNPass"].Value != null)
                            {
                                string acc = row.Cells["dtAcc"].Value.ToString();
                                string pass = row.Cells["dtNPass"].Value.ToString();
                                string line = $"{acc}|{pass}";
                                writer.WriteLine(line);
                            }
                            else
                            {
                                string acc = row.Cells["dtAcc"].Value.ToString();
                                string pass = row.Cells["dtPass"].Value.ToString();
                                string line = $"{acc}|{pass}";
                                writer.WriteLine(line);
                            }
                        }
                    }

                    //MessageBox.Show("Export successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            }
        }

        //private bool isRunning = false; // biến cờ theo dõi trạng thái của các luồng

        private List<Task> runningTasks = new List<Task>(); // danh sách các luồng đang chạy
        
        private void button5_Click(object sender, EventArgs e)
        {
            // Yêu cầu hủy bỏ đối tượng CancellationToken
            if (cancellationTokenSource != null)
            {
                cancellationTokenSource.Cancel();
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            OpenProifle(txtAcc.Text, (int)numericUpDown1.Value, (int)numericUpDown2.Value);
        }

        //load lại dữ liệu datagridview2
        private void UpdateDataGridView2()
        {
            dataGridView2.Rows.Clear();

            string userNamePc = Environment.UserName;

            // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
            string profilePath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";

            // Lấy danh sách các thư mục trong thư mục chứa các profile
            string[] folders = Directory.GetDirectories(profilePath);

            // Duyệt qua từng thư mục để kiểm tra các profile
            foreach (string folder in folders)
            {
                // Kiểm tra tên thư mục có chứa kí tự "@" hay không
                if (folder.Contains("@"))
                {
                    // Nếu có, lấy tên và đường dẫn đến thư mục đó
                    string folderName = new DirectoryInfo(folder).Name;

                    // Thêm thông tin profile vào DataGridView
                    int rowIndex = dataGridView2.Rows.Add();
                    dataGridView2.Rows[rowIndex].Cells["pSTT"].Value = rowIndex + 1; // Gán số thứ tự vào cột pSTT
                    dataGridView2.Rows[rowIndex].Cells["dtPro"].Value = folderName; // Gán giá trị vào cột dtPro
                }
            }
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                LoadDataFromFile();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Đã xảy ra lỗi khi đọc tệp tin: " + ex.Message);
            }
            string userNamePc = Environment.UserName;

            // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
            string profilepath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";            // Lấy danh sách các thư mục trong thư mục chứa các profile
            string[] folders = Directory.GetDirectories(profilepath);

            // Duyệt qua từng thư mục để kiểm tra các profile
            foreach (string folder in folders)
            {
                // Kiểm tra tên thư mục có chứa kí tự "@" hay không
                if (folder.Contains("@"))
                {
                    // Nếu có, lấy tên và đường dẫn đến thư mục đó
                    string folderName = new DirectoryInfo(folder).Name;

                    // Thêm thông tin profile vào DataGridView
                    int rowIndex = dataGridView2.Rows.Add();
                    dataGridView2.Rows[rowIndex].Cells["pSTT"].Value = rowIndex + 1; // Gán số thứ tự vào cột pSTT
                    dataGridView2.Rows[rowIndex].Cells["dtPro"].Value = folderName; // Gán giá trị vào cột dtPro
                }
            }

        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Lấy dòng được chọn
                DataGridViewRow selectedRow = dataGridView1.Rows[e.RowIndex];

                // Lấy giá trị của cột 1 và cột 2 trong dòng được chọn
                string account = selectedRow.Cells[1].Value.ToString();
                string password = selectedRow.Cells[2].Value.ToString();
                try
                {
                    Thread thread = new Thread(() =>
                    {
                        CreatProfile(account, password, noimage.Checked, headless.Checked);
                        UpdateDataGridView2();
                    });
                    thread.Start();
                    
                }
                catch
                {
                    Thread thread = new Thread(() =>
                    {
                        OpenProifle(account, (int)numericUpDown1.Value, (int)numericUpDown2.Value);

                    });
                    thread.Start();
                }
            }
        }

        private void dataGridView2_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Lấy dòng được chọn
                DataGridViewRow selectedRow = dataGridView2.Rows[e.RowIndex];

                // Lấy giá trị của cột 1 và cột 2 trong dòng được chọn
                string account = selectedRow.Cells[1].Value.ToString();
                Thread thread = new Thread(() =>
                {
                    OpenProifle(account, (int)numericUpDown1.Value, (int)numericUpDown2.Value);

                });
                thread.Start();
            }
        }

        private void dataGridView2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Delete)
            {
                var selectedCells = dataGridView2.SelectedCells;

                var confirmResult = MessageBox.Show("Bạn có chắc chắn muốn xoá dữ liệu đã chọn?", "Xác nhận xoá dữ liệu", MessageBoxButtons.YesNo);

                if (confirmResult == DialogResult.Yes)
                {
                    // Sử dụng một danh sách để lưu trữ các dòng đã được xoá để tránh lỗi khi xoá các ô
                    List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                    foreach (DataGridViewCell cell in selectedCells)
                    {
                        DataGridViewRow row = cell.OwningRow;

                        if (!rowsToRemove.Contains(row))
                        {
                            rowsToRemove.Add(row);

                            if (row.Cells["dtPro"].Value != null)
                            {
                                string folder = row.Cells["dtPro"].Value.ToString();
                                string userNamePc = Environment.UserName;

                                // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
                                string profilePath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";
                                try
                                {
                                    // Sử dụng Process để gọi lệnh Command Prompt và sử dụng lệnh rd để xoá thư mục
                                    ProcessStartInfo startInfo = new ProcessStartInfo("cmd.exe", "/c rd /s /q \"" + profilePath + "\"");
                                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(startInfo);
                                }
                                catch (Exception ex)
                                {
                                    string errorMessage = string.Format("Failed to delete profile \"{0}\". Error message: {1}", folder, ex.Message);
                                    MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                    MessageBox.Show("Profile deleted successfully.");

                    // Xoá các dòng đã được lưu trữ trong danh sách
                    foreach (DataGridViewRow row in rowsToRemove)
                    {
                        dataGridView2.Rows.Remove(row);
                    }
                }
            }
        }


        private void dataGridView1_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Lấy các dòng được bôi đen
                DataGridViewSelectedRowCollection selectedRows = dataGridView1.SelectedRows;

                // Nếu có nhiều hơn một dòng được bôi đen, hiển thị ContextMenuStrip
                if (selectedRows.Count >= 1)
                {
                    // Tạo một ContextMenuStrip mới
                    ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

                    // Thêm các ToolStripMenuItem vào ContextMenuStrip
                    ToolStripMenuItem chgPass = new ToolStripMenuItem("Change Pass");
                    ToolStripMenuItem crtProfile = new ToolStripMenuItem("Create Profile");
                    ToolStripMenuItem openPro = new ToolStripMenuItem("Open Profile");
                    contextMenuStrip.Items.Add(chgPass);
                    contextMenuStrip.Items.Add(crtProfile);
                    contextMenuStrip.Items.Add(openPro);

                    // Xác định vị trí của chuột và hiển thị ContextMenuStrip
                    contextMenuStrip.Show(dataGridView1, e.Location);

                    chgPass.Click += async (sender2, e2) =>
                    {
                        cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                        CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                        foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                        {
                            // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                            if (cancellationToken.IsCancellationRequested)
                            {
                                break;
                            }
                            if (row.Cells["dtAcc"].Value != null && row.Cells["dtPass"].Value != null)
                            {
                                string acc = row.Cells["dtAcc"].Value.ToString();
                                string pass = row.Cells["dtPass"].Value.ToString();
                                await ProcessDataAsync(acc, pass);
                            }
                        }
                        MessageBox.Show("Success!");
                    };



                    crtProfile.Click += async (sender2, e2) =>
                    {
                        cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                        CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                        foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                        {
                            // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                            if (cancellationToken.IsCancellationRequested)
                            {
                                break;
                            }
                            if (row.Cells["dtAcc"].Value != null && row.Cells["dtPass"].Value != null)
                            {
                                string acc = row.Cells["dtAcc"].Value.ToString();
                                string pass = row.Cells["dtPass"].Value.ToString();
                                await CreateProfileDataAsync(acc, pass);
                            }
                            UpdateDataGridView2();
                        }
                        MessageBox.Show("Success!");
                    };

                    openPro.Click += async (sender2, e2) =>
                    {
                        cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                        CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                        foreach (DataGridViewRow row in dataGridView1.SelectedRows)
                        {
                            // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                            if (cancellationToken.IsCancellationRequested)
                            {
                                break;
                            }
                            if (row.Cells["dtAcc"].Value != null && row.Cells["dtPass"].Value != null)
                            {
                                string acc = row.Cells["dtAcc"].Value.ToString();
                                await OpenProfileDataAsync(acc, (int)numericUpDown1.Value, (int)numericUpDown2.Value);
                            }
                        }
                        MessageBox.Show("Success!");
                    };
                }
            }
        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.F5)
            {
                UpdateDataGridView2();
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Check if the changed cell is in column 2 and the row is not a new row
            if (e.ColumnIndex == 4 && e.RowIndex >= 0 && e.RowIndex < dataGridView1.Rows.Count - 1)
            {
                // Get the mail and pass values from the corresponding cells
                string mail = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
                string pass = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();

                // Build the path to the data file
                string path = @"C:\Users\Ha Nguyen\Desktop\data.txt";

                // Write the mail and pass to a new line in the file
                using (StreamWriter sw = File.AppendText(path))
                {
                    sw.WriteLine(mail + "|" + pass);
                }
            }
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Delete)
            {
                var selectedCells = dataGridView1.SelectedCells;


                // Sử dụng một danh sách để lưu trữ các dòng đã được xoá để tránh lỗi khi xoá các ô
                List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                foreach (DataGridViewCell cell in selectedCells)
                {
                    DataGridViewRow row = cell.OwningRow;

                    if (!rowsToRemove.Contains(row))
                    {
                        rowsToRemove.Add(row);
                    }
                }
                // Xoá các dòng đã được lưu trữ trong danh sách
                foreach (DataGridViewRow row in rowsToRemove)
                {
                    dataGridView1.Rows.Remove(row);
                }

            }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            // Kích thước trình duyệt mong muốn
            int width = 700;
            int height = 500;

            // Khởi tạo ChromeOptions
            ChromeOptions options = new ChromeOptions();

            // Đặt kích thước cửa sổ trình duyệt
            options.AddArgument($"--window-size={width},{height}");

            // Khởi tạo trình điều khiển Chrome với ChromeOptions
            IWebDriver driver = new ChromeDriver(options);
        }

        private void dataGridView3_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            // Lấy dòng hiện tại
            DataGridViewRow currentRow = dataGridView3.Rows[e.RowIndex];

            // Cập nhật giá trị cho ô số thứ tự
            currentRow.Cells["mSTT"].Value = e.RowIndex + 1;
        }

        private void dataGridView1_llContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

       

        private string dataFilePath = "E:\\lo3.txt"; // Đường dẫn của tệp tin dữ liệu

        // Đọc dữ liệu từ tệp tin và load lên DataGridView
        private void LoadDataFromFile()
        {
            
                // Đọc dữ liệu từ tệp tin và gán cho DataGridView
                string[] lines = File.ReadAllLines(dataFilePath);
                foreach (string line in lines)
                {
                    // Kiểm tra xem dòng có chứa toàn bộ ký tự '|' hay không
                    if (!string.IsNullOrWhiteSpace(line) && !line.All(c => c == '|'))
                    {
                        string[] rowValues = line.Split('|');

                        // Thêm dòng vào DataGridView
                        int rowIndex = dataGridView3.Rows.Add();

                        // Gán giá trị từ mảng rowValues vào các ô của dòng tương ứng
                        for (int i = 0; i < 7; i++)
                        {
                            string value = rowValues[i];

                            // Kiểm tra nếu là ô CheckBox
                            if (dataGridView3.Columns[i] is DataGridViewCheckBoxColumn)
                            {
                                bool checkBoxValue = value == "1";
                                dataGridView3.Rows[rowIndex].Cells[i].Value = checkBoxValue;
                            }
                            else
                            {
                                dataGridView3.Rows[rowIndex].Cells[i].Value = value;
                            }
                        }
                    }
                }

            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Lưu dữ liệu từ DataGridView vào tệp tin
            using (StreamWriter writer = new StreamWriter(dataFilePath))
            {
                foreach (DataGridViewRow row in dataGridView3.Rows)
                {
                    string[] values = new string[row.Cells.Count];
                    for (int i = 0; i < row.Cells.Count; i++)
                    {
                        if (row.Cells[i] is DataGridViewCheckBoxCell)
                        {
                            DataGridViewCheckBoxCell checkBoxCell = row.Cells[i] as DataGridViewCheckBoxCell;
                            bool? checkBoxValue = checkBoxCell?.Value as bool?;
                            values[i] = checkBoxValue.HasValue ? (checkBoxValue.Value ? "1" : "0") : "";
                        }
                        else if (row.Cells[i] is DataGridViewComboBoxCell)
                        {
                            DataGridViewComboBoxCell comboBoxCell = (DataGridViewComboBoxCell)row.Cells[i];
                            string comboBoxValue = comboBoxCell.Value?.ToString() ?? "";
                            values[i] = comboBoxValue;
                        }
                        else
                        {
                            values[i] = row.Cells[i].Value?.ToString() ?? "";
                        }
                    }
                    writer.WriteLine(string.Join("|", values));
                }
            }
        }

        private void dataGridView3_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView3_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Delete)
            {
                var selectedCells = dataGridView3.SelectedCells;

                var confirmResult = MessageBox.Show("Bạn có chắc chắn muốn xoá dữ liệu đã chọn?", "Xác nhận xoá dữ liệu", MessageBoxButtons.YesNo);

                if (confirmResult == DialogResult.Yes)
                {
                    // Sử dụng một danh sách để lưu trữ các dòng đã được xoá để tránh lỗi khi xoá các ô
                    List<DataGridViewRow> rowsToRemove = new List<DataGridViewRow>();

                    foreach (DataGridViewCell cell in selectedCells)
                    {
                        DataGridViewRow row = cell.OwningRow;

                        if (!rowsToRemove.Contains(row))
                        {
                            rowsToRemove.Add(row);

                            if (row.Cells["mAcc"].Value != null)
                            {
                                string folder = row.Cells["mAcc"].Value.ToString();
                                string userNamePc = Environment.UserName;

                                // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
                                string profilePath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";
                                try
                                {
                                    // Sử dụng Process để gọi lệnh Command Prompt và sử dụng lệnh rd để xoá thư mục
                                    ProcessStartInfo startInfo = new ProcessStartInfo("cmd.exe", "/c rd /s /q \"" + profilePath + "\"");
                                    startInfo.WindowStyle = ProcessWindowStyle.Hidden;
                                    Process.Start(startInfo);
                                }
                                catch (Exception ex)
                                {
                                    string errorMessage = string.Format("Failed to delete profile \"{0}\". Error message: {1}", folder, ex.Message);
                                    MessageBox.Show(errorMessage, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                        }
                    }
                    MessageBox.Show("Profile deleted successfully.");

                    // Xoá các dòng đã được lưu trữ trong danh sách
                    foreach (DataGridViewRow row in rowsToRemove)
                    {
                        dataGridView3.Rows.Remove(row);
                    }
                }
            }
        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void dataGridView3_DragDrop(object sender, DragEventArgs e)
        {
            
        }

        private void dataGridView3_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                // Lấy dòng được chọn
                DataGridViewRow selectedRow = dataGridView3.Rows[e.RowIndex];

                // Lấy giá trị của cột 1 và cột 2 trong dòng được chọn
                string account = selectedRow.Cells[1].Value.ToString();
                string password = selectedRow.Cells[2].Value.ToString();
                // Lấy tên người dùng hiện tại của máy tính
                string userNamePc = Environment.UserName;

                // Tạo đường dẫn đến thư mục profile của Chrome dựa trên tên người dùng
                string profilesDirPath = "C:\\Users\\" + userNamePc + "\\AppData\\Local\\Google\\Chrome\\User Data";

                // Tên đăng nhập của bạn

                // Đường dẫn đến thư mục của profile trình duyệt mới
                string profileDirPath = Path.Combine(profilesDirPath, account);

                // Tạo thư mục profile trình duyệt mới nếu nó chưa tồn tại
                if (!Directory.Exists(profileDirPath))
                {
                    Thread thread = new Thread(() =>
                    {
                        CreatProfile(account, password, noimage.Checked, headless.Checked);
                        //UpdateDataGridView2();
                    });
                    thread.Start();
                }
                
               
                else
                {
                    Thread thread = new Thread( () =>
                    {
                        OpenProifle(account, (int)numericUpDown1.Value, (int)numericUpDown2.Value);
                        //searchHelper.Auto((int)numericUpDown3.Value, (int)numericUpDown4.Value, account, (int)numericUpDown1.Value, (int)numericUpDown2.Value, txtNoti, checkBox4.Checked, checkBox3.Checked,dataGridView3, e);
                        //searchHelper.point(account, (int)numericUpDown1.Value, (int)numericUpDown2.Value, txtNoti, checkBox3.Checked);
                    });
                    thread.Start();

                }
            }
        }

        private void dataGridView3_MouseClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                // Lấy các dòng được bôi đen
                DataGridViewSelectedRowCollection selectedRows = dataGridView3.SelectedRows;

                // Nếu có nhiều hơn một dòng được bôi đen, hiển thị ContextMenuStrip
                if (selectedRows.Count >= 1)
                {
                    // Tạo một ContextMenuStrip mới
                    ContextMenuStrip contextMenuStrip = new ContextMenuStrip();

                    // Thêm các ToolStripMenuItem vào ContextMenuStrip
                    ToolStripMenuItem chgPass = new ToolStripMenuItem("Change Pass");
                    ToolStripMenuItem crtProfile = new ToolStripMenuItem("Create Profile");
                    ToolStripMenuItem openPro = new ToolStripMenuItem("Open Profile");
                    ToolStripMenuItem AutoSearch = new ToolStripMenuItem("Auto Search");
                    ToolStripMenuItem setfarmed = new ToolStripMenuItem("Set Farmed");
                    contextMenuStrip.Items.Add(chgPass);
                    contextMenuStrip.Items.Add(crtProfile);
                    contextMenuStrip.Items.Add(openPro);
                    contextMenuStrip.Items.Add(AutoSearch);
                    contextMenuStrip.Items.Add(setfarmed);

                    // Xác định vị trí của chuột và hiển thị ContextMenuStrip
                    contextMenuStrip.Show(dataGridView3, e.Location);

                    chgPass.Click += async (sender2, e2) =>
                    {
                        cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                        CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                        foreach (DataGridViewRow row in dataGridView3.SelectedRows)
                        {
                            // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                            if (cancellationToken.IsCancellationRequested)
                            {
                                break;
                            }
                            if (row.Cells["mAcc"].Value != null && row.Cells["mPass"].Value != null)
                            {
                                string acc = row.Cells["mAcc"].Value.ToString();
                                string pass = row.Cells["mPass"].Value.ToString();
                                await ProcessDataAsync(acc, pass);
                            }
                        }
                        MessageBox.Show("Success!");
                    };



                    crtProfile.Click += async (sender2, e2) =>
                    {
                        cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                        CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                        foreach (DataGridViewRow row in dataGridView3.SelectedRows)
                        {
                            // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                            if (cancellationToken.IsCancellationRequested)
                            {
                                break;
                            }
                            if (row.Cells["mAcc"].Value != null && row.Cells["mPass"].Value != null)
                            {
                                string acc = row.Cells["mAcc"].Value.ToString();
                                string pass = row.Cells["mPass"].Value.ToString();

                                if (!acc.Contains("@") || string.IsNullOrEmpty(acc) || string.IsNullOrEmpty(pass))
                                {
                                    MessageBox.Show("Lỗi dữ liệu đầu vào!");
                                    return;
                                }

                                await CreateProfileDataAsync(acc, pass);
                            }
                            //UpdateDataGridView2();
                        }
                        MessageBox.Show("Success!");
                    };

                    AutoSearch.Click += async (sender2, e2) =>
                    {
                        foreach (DataGridViewRow row in dataGridView3.SelectedRows)
                        {
                            if (row.Cells["mAcc"].Value != null && row.Cells["mPass"].Value != null)
                            {
                                string account = row.Cells["mAcc"].Value.ToString();
                                DataGridViewCellEventArgs args = new DataGridViewCellEventArgs(0, row.Index);
                                await Task.Run(() =>
                                {
                                    searchHelper.Auto((int)numericUpDown3.Value, (int)numericUpDown4.Value, account, (int)numericUpDown1.Value, (int)numericUpDown2.Value, txtNoti, checkBox3.Checked, checkBox4.Checked, dataGridView3, args);
                                });
                                await Task.Delay(1000); // Đợi 1 giây trước khi tiếp tục với hàng tiếp theo
                            }
                        }
                        //MessageBox.Show("Success!");
                    };


                    openPro.Click += async (sender2, e2) =>
                    {
                        cancellationTokenSource = new CancellationTokenSource(); // Tạo mới đối tượng CancellationTokenSource
                        CancellationToken cancellationToken = cancellationTokenSource.Token; // Lấy đối tượng CancellationToken từ CancellationTokenSource

                        foreach (DataGridViewRow row in dataGridView3.SelectedRows)
                        {
                            // Kiểm tra nếu đối tượng CancellationToken được yêu cầu hủy bỏ thì dừng luồng
                            if (cancellationToken.IsCancellationRequested)
                            {
                                break;
                            }
                            if (row.Cells["mAcc"].Value != null && row.Cells["mPass"].Value != null)
                            {
                                string acc = row.Cells[1].Value.ToString();
                                await OpenProfileDataAsync(acc, (int)numericUpDown1.Value, (int)numericUpDown2.Value);
                            }
                        }
                        //MessageBox.Show("Success!");
                    };

                    setfarmed.Click += (sender2, e2) =>
                    {
                        foreach (DataGridViewRow row in dataGridView3.SelectedRows)
                        {
                            DataGridViewCheckBoxCell checkBoxCell = (DataGridViewCheckBoxCell)row.Cells["mFarmed"];
                            checkBoxCell.Value = true;
                        }

                    };
                }
            }
        }

        private Process vpnProcess; // Lưu trữ quá trình NordVPN

        private void ConnectVPN()
        {
            string destination = string.Empty;
            comboBox1.Invoke((MethodInvoker)(() => destination = comboBox1.Text));

            string nordVPNPath = @"C:\Program Files\NordVPN\nordvpn.exe";
            string command = "nordvpn -c -g \"" + destination + "\""; // Kết nối nhanh đến VPN

            ProcessStartInfo processInfo = new ProcessStartInfo(nordVPNPath, command);
            processInfo.CreateNoWindow = true;
            processInfo.UseShellExecute = false;

            vpnProcess = new Process();
            vpnProcess.StartInfo = processInfo;
            vpnProcess.Start();

            // Đợi quá trình kết nối hoàn thành (tuỳ vào tốc độ mạng và thời gian kết nối)
            vpnProcess.WaitForExit();

            // Gửi thông báo kết nối VPN thành công hoặc thất bại tới UI thread
            this.Invoke((MethodInvoker)(() =>
            {
                if (vpnProcess.ExitCode == 0)
                {
                    MessageBox.Show("Connected to VPN");
                }
                else
                {
                    MessageBox.Show("Failed to connect to VPN");
                }
            }));
        }

        private void DisconnectVPN()
        {
            if (vpnProcess != null && !vpnProcess.HasExited)
            {
                vpnProcess.Kill(); // Ngắt kết nối VPN bằng cách kết thúc quá trình NordVPN
                vpnProcess = null;

                MessageBox.Show("Disconnected from VPN");
            }
            else
            {
                MessageBox.Show("No active VPN connection");
            }
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
            Thread thread = new Thread(() =>
            {
                ConnectVPN();
            });
            thread.Start();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            DisconnectVPN();
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_RowEnter(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView2_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

        }

        private void dataGridView2_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Kiểm tra nếu cột thay đổi là cột dtPro
            if (dataGridView2.Columns[e.ColumnIndex].Name == "dtPro")
            {
                // Cập nhật lại số thứ tự cho các hàng
                for (int i = 0; i < dataGridView2.Rows.Count; i++)
                {
                    dataGridView2.Rows[i].Cells["pSTT"].Value = i + 1;
                }
            }
        }



        private async void button9_Click(object sender, EventArgs e)
        {
            SearchHelper searchHelper = new SearchHelper();
            List<string> words = await searchHelper.GetRandomWords(30);
            string wordString = string.Join(", ", words);
            txtNoti.Text = wordString;
        }


        private void button10_Click(object sender, EventArgs e)
        {
            LoadDataFromFile();
        }
    }
}
