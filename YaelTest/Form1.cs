using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager.DriverConfigs.Impl;
using System.Threading;

namespace YaelTest
{
    class WorkDay
    {
        public string date { get; set; }
        public string start { get; set; }
        public string end { get; set; }
    }
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {


            string driverPath = Application.StartupPath; // תיקיית הפרויקט / exe
            ChromeOptions options = new ChromeOptions();
            options.AddArgument("--start-maximized");

            IWebDriver driver = new ChromeDriver(driverPath, options);

            // LOGIN PAGE
            driver.Navigate().GoToUrl("https://yaelsoft.net.hilan.co.il/login");
            driver.FindElement(By.Id("user_nm")).SendKeys("");
            driver.FindElement(By.Id("password_nm")).SendKeys("");
            driver.FindElement(By.XPath("//button[contains(text(),'כניסה')]")).Click();
        }
    }
}
