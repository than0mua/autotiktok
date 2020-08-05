using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System.IO;
using System.Threading;
using System.Net;
using System.Net.NetworkInformation;
using System.Management;
using cExcel = Microsoft.Office.Interop.Excel;
using System.Threading;
using AutoItX3Lib;
using System.Diagnostics;

namespace Auto_Tiktok
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IWebDriver driver;
        IJavaScriptExecutor js;


        public MainWindow()
        {
            InitializeComponent();
            List<string> sheetName = getAllSheetName();
            var result = sheetName.Select(s => new { value = s }).ToList();
            dataGrid.ItemsSource = result;
        }
        public void Start_Browser()
        {
            /* Local Selenium WebDriver */
            //load profile firefox
            FirefoxProfile profile = new FirefoxProfileManager().GetProfile("Selenium");
            //set useragent de bypass
            var useragent = comboBox.Text;
            profile.SetPreference("general.useragent.override", useragent);
            FirefoxOptions options = new FirefoxOptions();
            options.Profile = profile;
            driver = new FirefoxDriver(options);
            driver.Manage().Window.Maximize();
            //khoi tao bien js de thuc thi js selenium
            js = (IJavaScriptExecutor)driver;
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {

            var dnsaddress = comboBox1.Text;
            //doi DNS
            var test = SetDNS(dnsaddress);
            Console.WriteLine("Test {0}", test);
            //Start trinh duyet
            Start_Browser();
            var tag = textBox.Text;
            //gan url
            driver.Url = "https://www.tiktok.com/tag/"+tag+"?lang=vi";
            //truy cap url
            driver.Navigate();
            //lay id trong the cua tiktok. sau mot time tiktok se thay doi id nay
            var id = textBox1.Text;

            Thread.Sleep(TimeSpan.FromSeconds(3));
            List<string> danhsachtacgia = new List<string>();
            var SCROLL_PAUSE_TIME = 3;
            
            // Get scroll height
            var last_height = js.ExecuteScript("return document.body.scrollHeight");
            var times = 0;
            var crawl_times = textBox2.Text;
            //so lan cuon xuong cuoi trang lay video(moi lan 30 video)
            while (times <= Int16.Parse(crawl_times))
            {
                //lay element tat ca cac video 
                var urlvideos = driver.FindElements(By.XPath("//a[@class='jsx-" + id + " video-feed-item-wrapper']"));
                foreach (var item in urlvideos)
                {
                    //lay link cua video trong element bang href
                    string urlvideo = item.GetAttribute("href");
                    if (urlvideo.Contains('@'))
                    {
                        //add link video vao mot list
                        danhsachtacgia.Add(urlvideo);
                    }
                    
                }
                
                //Scroll down to bottom. cuon xuong cuoi man hinh de load them video
                js.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");
                //Wait to load page
                Thread.Sleep(TimeSpan.FromSeconds(SCROLL_PAUSE_TIME));
                //dem so lan cuon
                times += 1;
            }
            //boc tach danh sach tac gia
            Getlistvideo(danhsachtacgia);
            Thread.Sleep(TimeSpan.FromSeconds(SCROLL_PAUSE_TIME));
            
            Close_Browser();
        }
        private void Getlistvideo(List<string> danhsachvideo)
        {
            // Open a new window
            // This does not change focus to the new window for the driver.
            //js.ExecuteScript ("window.open('');");
            //time.sleep(3)
            // Switch to the new window
            //driver.SwitchTo().Window(driver.WindowHandles[1]);
            //chuan bi file excel de chua du lieu
            // Get fully qualified path for xlsx file
            var spreadsheetLocation = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
            //Mở chương trình Excel:
            cExcel.Application app = new cExcel.Application();
            //Mở File Excel(*.xls) có sẵn:
            object valueMissing = System.Reflection.Missing.Value;
            cExcel.Workbook book = app.Workbooks.Open(spreadsheetLocation, valueMissing,
                false, valueMissing, valueMissing, valueMissing, valueMissing,
                cExcel.XlPlatform.xlWindows, valueMissing, valueMissing,
                valueMissing, valueMissing, valueMissing, valueMissing, valueMissing);
            //Tạo 1 worksheet
            cExcel.Worksheet sheet = (cExcel.Worksheet)book.Worksheets.Add(valueMissing, valueMissing, valueMissing, valueMissing);
            sheet.Name = DateTime.Now.ToString("yyyyMMddHHmmss");
            var count = 1;
            //chuan bi file excel de chua du lieu xong
            foreach (var item in danhsachvideo)
            {
                //mo tung link trong danh sach video
                driver.Url = item;
                driver.Navigate();
                Thread.Sleep(TimeSpan.FromSeconds(1));
                //get link video
                var linkvideo = driver.FindElement(By.XPath("//video['jsx-3900254205  horizontal video-player']"));
                //Get like
                var like = driver.FindElement(By.XPath("//strong[@class='jsx-624911782 like-text']"));
                //get comment
                var comment = driver.FindElement(By.XPath("//strong[@class='jsx-624911782 comment-text']"));
                //get title
                var title = driver.FindElement(By.XPath("//strong[@class='jsx-961678791']"));
                //kiem tra chi lay cac video trieu like va tram nghin comment
                if (like.Text.EndsWith("M")&& comment.Text.EndsWith("K"))
                {
                    string video_url = linkvideo.GetAttribute("src");
                    //ghi danh sach url video vao file excle
                    //lấy vùng cần điền giá trị
                    cExcel.Range rng_name = sheet.get_Range("A"+count.ToString(), "A" + count.ToString());
                    cExcel.Range rng_link = sheet.get_Range("B" + count.ToString(), "B" + count.ToString());
                    cExcel.Range rng_like = sheet.get_Range("C" + count.ToString(), "C" + count.ToString());
                    cExcel.Range rng_comment = sheet.get_Range("D" + count.ToString(), "D" + count.ToString());
                    cExcel.Range rng_title = sheet.get_Range("E" + count.ToString(), "E" + count.ToString());
                    count++;
                    //đưa giá trị vào rng (tức ô A2)
                    rng_name.Value2 = "'"+item.Split('/')[5].ToString();
                    rng_link.Value2 = video_url;
                    rng_like.Value2 = like.Text;
                    rng_comment.Value2 = comment.Text;
                    rng_title.Value2 = title.Text;
                    
                    
                }
                    
            }
            //driver.SwitchTo().Window(driver.WindowHandles[1]).Close();
            //driver.SwitchTo().Window(driver.WindowHandles[0]);
            //Lưu File *.xls:
            book.Save();
            //Đóng File và tắt chương trình Excel:
            book.Close(true, valueMissing, valueMissing);
            app.Quit();

            List<string> sheetName = getAllSheetName();
            var result = sheetName.Select(s => new { value = s }).ToList();
            dataGrid.ItemsSource = result;
        }

        public List<String> getAllSheetName()
        {
            // Get fully qualified path for xlsx file
            var spreadsheetLocation = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
            //Mở chương trình Excel:
            cExcel.Application app = new cExcel.Application();
            //Mở File Excel(*.xls) có sẵn:
            object valueMissing = System.Reflection.Missing.Value;
            cExcel.Workbook book = app.Workbooks.Open(spreadsheetLocation, valueMissing,
                false, valueMissing, valueMissing, valueMissing, valueMissing,
                cExcel.XlPlatform.xlWindows, valueMissing, valueMissing,
                valueMissing, valueMissing, valueMissing, valueMissing, valueMissing);

            List<String> excelSheets = new List<String>();
            int i = 0;
            foreach (Microsoft.Office.Interop.Excel.Worksheet wSheet in book.Worksheets)
            {
                excelSheets.Add( wSheet.Name);
                i++;
            }
            //Đóng File và tắt chương trình Excel:
            book.Close(true, valueMissing, valueMissing);
            app.Quit();
            return excelSheets;
        }
        public void Close_Browser()
        {
            driver.Quit();
        }
        public bool SetDNS(string dns)
        {
            ManagementClass objMC = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection objMOC = objMC.GetInstances();

            foreach (ManagementObject objMO in objMOC)
            {
                if ((bool)objMO["IPEnabled"])
                {

                    string Description = objMO["Description"].ToString();

                    if (Description.Contains("Wireless"))
                        continue;

                    // Set Preferred DNS
                    try
                    {
                        ManagementBaseObject newDNS =
                                             objMO.GetMethodParameters("SetDNSServerSearchOrder");
                        newDNS["DNSServerSearchOrder"] = dns.Split(',');
                        ManagementBaseObject setDNS =
                            objMO.InvokeMethod("SetDNSServerSearchOrder", newDNS, null);

                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Preferred DNS set failed");
                        return false;
                    }
                }
            }

            return true;
        }

        private void Button2_Click(object sender, RoutedEventArgs e)
        {
            FirefoxProfile profile = new FirefoxProfileManager().GetProfile("Selenium");
            FirefoxOptions options = new FirefoxOptions();
            options.Profile = profile;
            IWebDriver driver = new FirefoxDriver(options);
            driver.Url = "https://youtube.com/";
            driver.Navigate();
            driver.Manage().Window.Maximize();
            driver.FindElement(By.XPath("//*[@id='button'][@aria-label='Tạo']")).Click();
            Thread.Sleep(TimeSpan.FromSeconds(3));
            driver.FindElement(By.XPath("//*[@id='label'][text()='Tải video lên']")).Click();
            //driver.FindElement(By.Id("next-button")).Click();
            driver.FindElement(By.Id("select-files-button")).Click();

            Thread.Sleep(TimeSpan.FromSeconds(3));

            // khởi tạo đối tượng autoIT để dùng cho C# -> nhờ nó send key click chuột dùm cái ở ngoài web browser
            AutoItX3 autoIT = new AutoItX3();

            // đưa title của cửa sổ File upload vào chuỗi. 
            // Cửa sổ hiện ra có thể có tiêu đề là File Upload hoặc Tải lên một tập tin
            // lấy ra cửa sổ active có tiêu đề như dưới
            autoIT.WinActivate("Tải lên một tệp tin");

            // file data nằm trong thư mục debug
            // gửi link vào ô đường dẫn

            //autoIT.Send(Application.Current.StartupUri + "//Kteam Data Upload.txt");
            autoIT.Send(@"G:\Linh Tinh\Python\video\793048456765246725.mp4");
            Thread.Sleep(TimeSpan.FromSeconds(1));
            // gửi phím enter sau khi truyền link vào
            autoIT.Send("{ENTER}");
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Process firstProc = new Process();
            firstProc.StartInfo.FileName = @"F:\Program Files\TechSmith\Camtasia Studio 9\CamtasiaStudio.exe";
            firstProc.EnableRaisingEvents = true;

            firstProc.Start();
            AutoItX3 au = new AutoItX3();
            au.WinWaitActive("Camtasia 9");
            //au.MouseMove(650, 300);
            au.MouseClick("LEFT", 650, 300);
            firstProc.WaitForExit();
            
        }

        private void button3_Click(object sender, RoutedEventArgs e)
        {
            // tạo ra danh sách UserInfo rỗng để hứng dữ liệu.
            List<VideoInfo> videoList = new List<VideoInfo>();
            var sheetName = textBox3.Text;
            try
            {
                // Get fully qualified path for xlsx file
                var spreadsheetLocation = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "test.xlsx");
                //Mở chương trình Excel:
                cExcel.Application app = new cExcel.Application();
                // mở file excel
                //Mở File Excel(*.xls) có sẵn:
                object valueMissing = System.Reflection.Missing.Value;
                cExcel.Workbook book = app.Workbooks.Open(spreadsheetLocation, valueMissing,
               false, valueMissing, valueMissing, valueMissing, valueMissing,
               cExcel.XlPlatform.xlWindows, valueMissing, valueMissing,
               valueMissing, valueMissing, valueMissing, valueMissing, valueMissing);
                // lấy ra sheet đầu tiên để thao tác
                
                cExcel.Worksheet sheet = (cExcel.Worksheet)book.Sheets[sheetName];
                sheet.Activate();
                cExcel.Range last = sheet.Cells.SpecialCells(cExcel.XlCellType.xlCellTypeLastCell, Type.Missing);
                cExcel.Range range = sheet.get_Range("A1", last);

                int lastUsedRow = last.Row;
                int lastUsedColumn = last.Column;
                // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file. lưu ý file excel bắt đầu từ số 1 không phải số 0
                for (int i =1; i <= lastUsedRow; i++)
                {
                    try
                    {

                        // lấy ra cột họ tên tương ứng giá trị tại vị trí [i, 1]. i lần đầu là 2
                        // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh
                        var name=sheet.Cells[i, 1].Value.ToString();
                        var url=sheet.Cells[i,2].Value.ToString();
                        // tạo UserInfo từ dữ liệu đã lấy được
                        VideoInfo video = new VideoInfo()
                        {
                            Name = name,
                            Url = url
                        };

                        // add UserInfo vào danh sách userList
                        videoList.Add(video);
                    }
                    catch (Exception)
                    {

                    }
                }
                book.Close(true, valueMissing, valueMissing);
                app.Quit();
            }
            catch (Exception ee)
            {
                MessageBox.Show("Error!");
            }
            var pathString = System.IO.Path.Combine(Directory.GetCurrentDirectory(), sheetName);
            System.IO.Directory.CreateDirectory(pathString);
            foreach (var item in videoList)
            {
                
                //download video
                using (var client = new WebClient())
                {
                    // Get fully qualified path for xlsx file
                    
                    client.DownloadFile(item.Url, pathString+"/"+item.Name+".mp4");
                }
            }

        }

        private void dataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

            var item = dataGrid.SelectedValue.ToString();
            textBox3.Text = item.Split(' ')[3];
        }

        private void dataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {

        }
    }
}
