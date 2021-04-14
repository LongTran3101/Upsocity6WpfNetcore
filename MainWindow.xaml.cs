using OfficeOpenXml;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading;
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

namespace Upsocity6WpfNetcore
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        ChromeDriver driver;
        ChromeOptions options = new ChromeOptions();
        ChromeDriverService driverService = ChromeDriverService.CreateDefaultService();
        public class Image
        {
            public string Foldername { get; set; }
            public string Imagename { get; set; }
            public string Title { get; set; }
            public string Des { get; set; }
            public string Tag { get; set; }
            public string Main { get; set; }
        }
        List<Image> readimage(string nameFile)
        {
            List<Image> imageList = new List<Image>();
            try
            {

                var package = new ExcelPackage(new FileInfo(nameFile));

                // lấy ra sheet đầu tiên để thao tác
                ExcelWorksheet workSheet = package.Workbook.Worksheets[1];

                // duyệt tuần tự từ dòng thứ 2 đến dòng cuối cùng của file. lưu ý file excel bắt đầu từ số 1 không phải số 0
                for (int i = workSheet.Dimension.Start.Row + 1; i <= workSheet.Dimension.End.Row; i++)
                {
                    try
                    {
                        // biến j biểu thị cho một column trong file
                        int j = 1;
                        Image image = new Image();
                        // lấy ra cột họ tên tương ứng giá trị tại vị trí [i, 1]. i lần đầu là 2
                        // tăng j lên 1 đơn vị sau khi thực hiện xong câu lệnh
                        if (workSheet.Cells[i, j].Value == null)
                        {
                            break;
                        }
                        string Foldername = workSheet.Cells[i, j++].Value.ToString();
                        image.Foldername = Foldername;
                        string Imagename = workSheet.Cells[i, j++].Value.ToString();
                        image.Imagename = Imagename;
                        string Title = workSheet.Cells[i, j++].Value.ToString();
                        image.Title = Title;
                        string Des = workSheet.Cells[i, j++].Value.ToString();
                        image.Des = Des;

                        string Tag = workSheet.Cells[i, j++].Value.ToString();
                        image.Tag = Tag;
                        string main = workSheet.Cells[i, j++].Value.ToString();
                        image.Main = main;


                        imageList.Add(image);




                    }
                    catch
                    {

                    }
                }
            }
            catch
            {
                MessageBox.Show("Error read excel!");
            }

            return imageList;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
           
            Thread thread = new Thread(() => {

                options.AddArguments("user-data-dir=ChromeProfile");
                options.AddArguments("--disable-notifications");
                options.AddArguments("start-maximized");
                options.AddExcludedArgument("enable-automation");
                options.AddAdditionalCapability("useAutomationExtension", false);
                options.AddArgument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.183 Safari/537.36");
                driverService.HideCommandPromptWindow = true;
                driver = new ChromeDriver(driverService, options);



            });
            thread.IsBackground = true;
            thread.Start();
        }

        void task(List<Image> imageList, ChromeDriver driver, string username, string passs)
        {

            int timeoutto = 100 * 1000;
            int timeout = 0;
            Boolean oke = false;
            string exeFile = (new System.Uri(Assembly.GetEntryAssembly().CodeBase)).AbsolutePath;
            DirectoryInfo di = new DirectoryInfo(exeFile);
            Console.WriteLine(di.Parent.FullName);
            string fathparen = di.Parent.FullName;
            trangThaiTxb.Dispatcher.Invoke(() => trangThaiTxb.Text = "Bắt đầu");
            driver.Url = "https://fineartamerica.com/controlpanel/updateartwork.html?newartwork=true";
            Thread.Sleep(2000);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));



            oke = false;

            while (!oke && timeout < timeoutto)
            {
                try
                {

                    if (isElementXpath("/html/body/p", driver))
                    {
                        Trangthai = "login";
                        login(driver, username, passs);
                        Thread.Sleep(2000);
                        driver.Url = "https://fineartamerica.com/controlpanel/updateartwork.html?newartwork=true";
                        Thread.Sleep(2000);
                        Trangthai = "login success";
                        oke = true;
                    }
                    oke = true;





                }
                catch
                { Thread.Sleep(100); timeout = timeout + 100; }
            }

            Console.WriteLine("done login");
            //Thread.Sleep(100000);
            //oke = false;

            //while (!oke && timeout < timeoutto)
            //{
            //    try
            //    {
            //        driver.FindElement(By.CssSelector("a#nav-user-sell")).Click();
            //        oke = true;
            //    }
            //    catch
            //    { Thread.Sleep(100); timeout = timeout + 100; }
            //}

            String ScripclickAll = "var items = document.querySelectorAll(\".undefined\");\n"
                   + "for (var i = 0; i < items.length; i++) {\n"
                   + "    \n"
                   + "        items[i].click();\n"
                   + "  \n"
                   + "}";

            using (StreamWriter sw = File.CreateText("log.txt"))
            {
                sw.WriteLine("");

            }
            for (int i = 0; i < imageList.Count(); i++)
            {
                timeoutto = 60 * 1000;
                Console.WriteLine("sau link up");

                //check file tồn tại
                try
                {
                    if (File.Exists(fathparen.Replace("%20", " ") + "\\" + imageList[j].Foldername + "\\" + imageList[j].Imagename))
                    {
                        wirte();
                    }
                    else
                    {

                        Console.WriteLine("lỗi file " + imageList[j].Imagename);
                        j++;
                        continue;
                    }
                }
                catch
                {
                    j++;
                    continue;
                }
                try
                {
                    Trangthai = "bat dau up " + j;
                    trangThaiTxb.Dispatcher.Invoke(() => trangThaiTxb.Text = Trangthai);
                }
                catch
                {


                }

                oke = false;

                while (!oke && timeout < timeoutto)
                {
                    try
                    {

                        IWebElement elem = driver.FindElement(By.XPath("//input[@type='file']"));
                        //string pathfile= imageList[i].Foldername + "/" + imageList[i].Imagename;
                        //Console.WriteLine(fathparen + "\\" + imageList[i].Foldername + "\\" + imageList[i].Imagename);
                        elem.SendKeys(fathparen.Replace("%20", " ") + "\\" + imageList[j].Foldername + "\\" + imageList[j].Imagename);

                        oke = true;
                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }
                try
                {
                    nameimage = imageList[j].Imagename;
                    imagenametxb.Dispatcher.Invoke(() => imagenametxb.Text = nameimage);
                    Console.WriteLine(imageList[j].Imagename);
                    Console.WriteLine("done up file");
                }
                catch
                {


                }




                oke = false;

                while (!oke && timeout < timeoutto)
                {
                    try
                    {
                        driver.FindElement(By.XPath("//*[@id='uploadImageDiv']/a/span")).Click();
                        oke = true;
                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }
                Thread.Sleep(30000);

                try
                {

                    if (driver.FindElement(By.XPath("/html/body/div[2]/div/div[2]/div[1]/p")).Displayed)
                    {
                        Trangthai = "Loi file qua lon tiep tuc up " + j;
                        trangThaiTxb.Dispatcher.Invoke(() => trangThaiTxb.Text = Trangthai);
                        driver.Quit();

                        driver = new ChromeDriver(driverService, options);

                        chromeDrivers.Add(driver);
                        j++;
                        task(imageList, driver, username, passs);
                        break;
                    }

                }
                catch
                {


                }
                oke = false;
                while (!oke && timeout < timeoutto)
                {
                    try
                    {


                        wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//*[@id='imageTitleDiv']/input")));
                        IWebElement title = driver.FindElement(By.XPath("//*[@id='imageTitleDiv']/input"));
                        title.Clear();
                        foreach (char c in imageList[j].Title)
                        {
                            title.SendKeys(c.ToString());
                            Thread.Sleep(300);
                        }

                        oke = true;

                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }

                Thread.Sleep(1000);
                oke = false;
                while (!oke && timeout < timeoutto)
                {
                    try
                    {




                        IWebElement main = driver.FindElement(By.XPath("//*[@id='imageMediumDiv']/input"));
                        main.Clear();
                        foreach (char c in imageList[j].Main)
                        {
                            main.SendKeys(c.ToString());
                            Thread.Sleep(300);
                        }

                        oke = true;


                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }
                Thread.Sleep(1000);
                oke = false;
                while (!oke && timeout < timeoutto)
                {
                    try
                    {


                        IWebElement tag = driver.FindElement(By.CssSelector("textarea#artworkkeywords"));
                        tag.Clear();
                        foreach (char c in imageList[j].Tag)
                        {
                            tag.SendKeys(c.ToString());
                            Thread.Sleep(300);
                        }

                        oke = true;


                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }

                Thread.Sleep(1000);
                oke = false;
                while (!oke && timeout < timeoutto)
                {
                    try
                    {


                        IWebElement des = driver.FindElement(By.XPath("//*[@id='imageDetailsDiv']/div/div[5]/textarea"));
                        des.Clear();
                        foreach (char c in imageList[j].Des)
                        {
                            des.SendKeys(c.ToString());
                            Thread.Sleep(300);
                        }

                        oke = true;


                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }

                Thread.Sleep(1000);
                oke = false;
                while (!oke && timeout < timeoutto)
                {
                    try
                    {


                        IWebElement Depart = driver.FindElement(By.CssSelector("select[name='artworkcategory']"));
                        SelectElement select = new SelectElement(Depart);
                        select.SelectByValue("8000");
                        oke = true;


                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }






                Console.WriteLine("fullthongtin");

                Thread.Sleep(1000);
                oke = false;
                while (!oke && timeout < timeoutto)
                {
                    try
                    {


                        IWebElement Depart = driver.FindElement(By.CssSelector("select[name='artworkcategory']"));
                        SelectElement select = new SelectElement(Depart);
                        select.SelectByValue("8000");
                        oke = true;


                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }

                Thread.Sleep(1000);
                oke = false;

                while (!oke && timeout < timeoutto)
                {
                    try
                    {
                        driver.FindElement(By.CssSelector("a.buttonSubmit")).Click();
                        oke = true;
                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }
                Trangthai = "da submit" + j;
                trangThaiTxb.Dispatcher.Invoke(() => trangThaiTxb.Text = Trangthai);
                Thread.Sleep(10000);

                oke = false;

                while (!oke && timeout < timeoutto)
                {
                    try
                    {

                        driver.Url = "https://fineartamerica.com/controlpanel/updateartwork.html?newartwork=true";
                        oke = true;




                    }
                    catch
                    { Thread.Sleep(100); timeout = timeout + 100; }
                }
                Thread.Sleep(2 * 1000);

                if (timeout >= timeoutto)
                {

                    Trangthai = "Lỗi restart" + j;
                    trangThaiTxb.Dispatcher.Invoke(() => trangThaiTxb.Text = Trangthai);
                    driver.Quit();

                    driver = new ChromeDriver(driverService, options);

                    chromeDrivers.Add(driver);
                    j++;
                    task(imageList, driver, username, passs);

                    break;
                }

                if (j != 0 && j % 50 == 0)
                {

                    Trangthai = "đã đạt 50 ảnh restart chờ 3p" + j;
                    trangThaiTxb.Dispatcher.Invoke(() => trangThaiTxb.Text = Trangthai);
                    driver.Quit();
                    Thread.Sleep(3 * 60 * 1000);
                    driver = new ChromeDriver(driverService, options);

                    chromeDrivers.Add(driver);
                    j++;
                    task(imageList, driver, username, passs);
                    break;
                }
                j++;
                //Console.WriteLine("bat dau click");




            }

            driver.Quit();
        }
    }
}
