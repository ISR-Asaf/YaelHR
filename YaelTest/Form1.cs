using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.IO;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using WebDriverManager.DriverConfigs.Impl;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using SeleniumExtras.WaitHelpers;
using System.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TaskbarClock;
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
        private const string CompanyCode = "9750";
        private const string EmployeeCode = "39572";
        private static readonly string UserGridIDPart = CompanyCode + EmployeeCode; // 975039572

        // ⭐ 3. קבועי זיהוי (החלק הסטטי של ה-ID)
        private static readonly string BaseGridID = $"ctl00_mp_RG_Days_{UserGridIDPart}_";
        private const string EntrySuffix = "_cellOf_ManualEntry_EmployeeReports_row_INDEX_0_ManualEntry_EmployeeReports_row_INDEX_0";
        private const string ExitSuffix = "_cellOf_ManualExit_EmployeeReports_row_INDEX_0_ManualExit_EmployeeReports_row_INDEX_0";
        private const string SaveSuffix = "_btnSave";

        // קבועים נוספים
        private const string SelectedDayClassName = "CSD"; // קלאס ליום מסומן
        private const string MonthListId = "ctl00_mp_calendar_monthChanged"; // ID של רשימת החודשים הנפתחת

        // נתוני הדיווח הקבועים
        private const string EntryTime = "08:00";
        private const string ExitTime = "17:00";
        private List<DateTime> datesToSelect = new List<DateTime>
        {
        new DateTime(2025, 12, 1),
        new DateTime(2025, 12, 2),
        new DateTime(2025, 12, 3)
        };

        private string GetMonthYearCode(DateTime date)
        {
            // מחלץ את קוד החודש והשנה בפורמט YYYY_MM (לדוגמה: 2025_11)
            return date.ToString("yyyy_MM");
        }

        private string GetTimeInID(string monthYearCode)
        {
            return BaseGridID + monthYearCode + EntrySuffix;
        }

        private string GetTimeOutID(string monthYearCode)
        {
            return BaseGridID + monthYearCode + ExitSuffix;
        }

        private string GetSaveButtonID(string monthYearCode)
        {
            return BaseGridID + monthYearCode + SaveSuffix;
        }
        public Form1()
        {
            InitializeComponent();
        }


        private void button1_Click(object sender, EventArgs e)
        {
            ChromeOptions opt = new ChromeOptions();
            opt.AddArgument("--start-maximized");

            IWebDriver driver = new ChromeDriver(opt);
            WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(15));

            try
            {
                // 1. כניסה למסך לוגין
                driver.Navigate().GoToUrl("https://yaelsoft.net.hilan.co.il/login");

                // 2. הזנת פרטים
                wait.Until(d => d.FindElement(By.Id("user_nm"))).SendKeys("");
                driver.FindElement(By.Id("password_nm")).SendKeys("");

                // שמירת ה-URL הנוכחי כדי לדעת מתי הוא משתנה
                string loginUrl = driver.Url;

                // 3. לחיצה על כניסה
                driver.FindElement(By.XPath("//button[contains(text(),'כניסה')]")).Click();

                // --- החלק החשוב: המתנה חכמה ---

                // במקום לחפש כפתור, נחכה שה-URL ישתנה (סימן שהלוגין הצליח)
                wait.Until(d => d.Url != loginUrl && !d.Url.Contains("login"));

                // ברגע שהלוגין הצליח - "קפוץ" ישר לדף הדיווח
                Console.WriteLine("התחברות הצליחה, עובר לדף דיווח...");
                driver.Navigate().GoToUrl("https://yaelsoft.net.hilan.co.il/Hilannetv2/Attendance/calendarpage.aspx?isOnSelf=true");

                Console.WriteLine("⬅ נמצא בדף הדיווח");
            }

            catch (Exception ex)
            {
                Console.WriteLine("התרחשה שגיאה: " + ex.Message);
            }

            try
            {
                // 1. המתנה לטעינת לוח השנה
                wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//td[normalize-space(text())='15']")));

                // 2. בחירת החודש הנכון
                string targetMonthText = datesToSelect[0].ToString("MMMM yyyy", new System.Globalization.CultureInfo("he-IL"));
                SelectMonthInCalendar(driver, wait, targetMonthText);

                Console.WriteLine($"החודש {targetMonthText} נבחר בהצלחה.");


                ClearSelectedDays(driver, wait);
                // 3. סימון הימים תוך בדיקת מצבם (משתמש בקלאס 'CSD')
                MarkDaysOnCalendar(driver, wait, datesToSelect);
                var selectedDaysBtnLocator = By.Id("ctl00_mp_RefreshSelectedDays");

                var selectedDaysBtn = wait.Until(ExpectedConditions.ElementToBeClickable(selectedDaysBtnLocator));

                selectedDaysBtn.Click();


                Console.WriteLine("נלחץ כפתור 'ימים נבחרים'.");
                Thread.Sleep(4000);
                FillInHoursAndSave(driver, wait, EntryTime, ExitTime);

                // ⭐ יש להמשיך מכאן לשלב מילוי הטבלה ושמירה
                // RunNextReportingStep(driver, wait); 
            }
            catch (Exception ex)
            {
                Console.WriteLine("שגיאה במהלך סימון התאריכים: " + ex.Message);
            }
        }



        private void FillInHoursAndSave(IWebDriver driver, WebDriverWait wait, string entryTime, string exitTime)
        {
   
            // ⭐ בניית ה-ID's הדינמיים
            string monthYearCode = GetMonthYearCode(datesToSelect.First());
            string timeInID = GetTimeInID(monthYearCode);
            string timeOutID = GetTimeOutID(monthYearCode);
            string saveButtonID = GetSaveButtonID(monthYearCode);

            Console.WriteLine($"מתחילים במילוי שעות: {entryTime} - {exitTime} עבור חודש {monthYearCode}.");

            for (int i = 0; i < datesToSelect.Count; i++)
            {

                try
                {
                    string timeInID_ROW = timeInID.Replace("INDEX", i.ToString());
                    string timeOutID_ROW = timeOutID.Replace("INDEX", i.ToString());
                    // 1. איתור שדות
                    var timeInField = wait.Until(ExpectedConditions.ElementIsVisible(By.Id(timeInID_ROW)));
                    var timeOutField = driver.FindElement(By.Id(timeOutID_ROW));

                    // ⭐ 2. לחיצה וניקוי (חיוני להתחלת הזנת המיסוך)
                    timeInField.Click();


                    // ⭐ 3. שליחת התווים אחד-אחד, ואז מעבר לטאב (כדי לחקות הקלדה אנושית על שדה ממוסך)
                    string[] parts = entryTime.Split(':');

                    entryTime = $"{parts[1]}:{parts[0]}";


                    foreach (char c in entryTime)
                    {
                        timeInField.SendKeys(c.ToString());
                        Thread.Sleep(50); // השהייה קצרה בין הקלדה להקלדה
                    }


                    Thread.Sleep(500);

                    // 4. מילוי שעת יציאה
                    timeOutField.Click();


                    parts = exitTime.Split(':');
                    exitTime = $"{parts[1]}:{parts[0]}";
                    foreach (char c in exitTime)
                    {
                        timeOutField.SendKeys(c.ToString());
                        Thread.Sleep(50);
                    }
                    // מעבר לשדה הבא, כדי שהמערכת תסיים לעבד את השעה
                    timeOutField.SendKeys(OpenQA.Selenium.Keys.Tab);

                    Console.WriteLine("שעות כניסה ויציאה מולאו בהצלחה בשורה הראשונה.");

                }
                catch (Exception ex)
                {
                    Console.WriteLine("שגיאה קריטית במהלך מילוי השעות: " + ex.Message);
                    Console.WriteLine($"בדוק את ה-ID's שנוצרו: כניסה={timeInID}, יציאה={timeOutID}, שמירה={saveButtonID}");
                }
            }
        }
        private void ClearSelectedDays(IWebDriver driver, WebDriverWait wait)
        {
            By selectedDaysLocator = By.XPath($"//td[contains(@class, '{SelectedDayClassName}')]");

            try
            {
                // FindElements מחזירה רשימה ריקה אם לא נמצאו אלמנטים
                var currentlySelectedDays = driver.FindElements(selectedDaysLocator);

                Console.WriteLine($"נמצאו {currentlySelectedDays.Count} ימים מסומנים שיש לבטל.");

                foreach (var dayElement in currentlySelectedDays)
                {
                    // לחיצה על יום מסומן מבטלת את הסימון שלו (CSD מוסר)
                    dayElement.Click();
                    Thread.Sleep(100); // המתנה קצרה ליציבות הממשק
                }

                if (currentlySelectedDays.Count > 0)
                {
                    // המתנה קצרה לוודא שהמערכת עיבדה את כל הביטולים
                    wait.Until(d => d.FindElements(selectedDaysLocator).Count == 0);
                }
            }
            catch (Exception ex)
            {
                // בדרך כלל לא נדרש כאן לכידת שגיאה מכיוון ש-FindElements לא נכשל אם לא נמצאו אלמנטים.
                Console.WriteLine($"שגיאה במהלך ניקוי סימונים קודמים: {ex.Message}");
            }
        }

   

        private void SelectMonthInCalendar(IWebDriver driver, WebDriverWait wait, string targetMonthText)
        {
            // 1. מציאת ה-SPAN המכיל את שם החודש הנוכחי (משמש כמתג לפתיחת הרשימה)
            // הערה: נשתמש ב-XPath חלופי כי ה-ID של ה-span זהה ל-UL, ונשתמש בטקסט כדי להיות בטוחים
            var monthToggleLocator = By.XPath("//span[contains(@class, 'fh-caret-down')]");
            var monthToggle = wait.Until(ExpectedConditions.ElementToBeClickable(monthToggleLocator));

            // בדיקה אם החודש כבר מוצג
            if (monthToggle.Text.Trim() == targetMonthText)
            {
                Console.WriteLine("החודש כבר מוצג, מדלג על בחירת חודש.");
                return;
            }

            // לחיצה על ה-Span לפתיחת רשימת החודשים
            monthToggle.Click();

            // 2. המתנה לפתיחת תפריט החודשים ובחירת החודש המבוקש
            // ⭐ ה-XPath המתוקן: מחפש LI בתוך ה-UL (שמזוהה ע"י ID) שמכיל את הטקסט המבוקש.
            // מכיוון שה-UL קיבל את אותו ID של ה-Span, נשתמש ב-ID של ה-UL כדי לזהות אותו.
            var targetMonthLocator = By.XPath($"//ul[@id='{MonthListId}']//li[normalize-space(text())='{targetMonthText}']");

            // המתנה שהאלמנט יהיה גלוי וניתן ללחיצה
            var targetMonthElement = wait.Until(ExpectedConditions.ElementToBeClickable(targetMonthLocator));

            // ⭐ לחיצה על פריט ה-LI לבחירת החודש
            targetMonthElement.Click();

            // 3. המתנה לטעינת לוח השנה החדש (נחכה שוב שיום 15 יופיע)
            wait.Until(ExpectedConditions.ElementIsVisible(By.XPath("//td[normalize-space(text())='15']")));
        }

        private void MarkDaysOnCalendar(IWebDriver driver, WebDriverWait wait, List<DateTime> dates)
        {
            foreach (var date in dates)
            {
                string day = date.Day.ToString();
                By dayLocator = By.XPath($"//td[normalize-space(text())='{day}']");

                try
                {
                    var dayElement = wait.Until(ExpectedConditions.ElementIsVisible(dayLocator));

                    string elementClasses = dayElement.GetAttribute("class");

                    // ⭐ הבדיקה המדויקת באמצעות הקלאס 'CSD'
                    if (!elementClasses.Contains(SelectedDayClassName))
                    {
                        dayElement.Click();
                        Console.WriteLine($"סומן יום: {day} (כי לא היה מסומן)");
                        Thread.Sleep(200);
                    }
                    else
                    {
                        Console.WriteLine($"יום {day} כבר היה מסומן ({SelectedDayClassName}), מדלג.");
                    }
                }
                catch (Exception ex) when (ex is NoSuchElementException || ex is WebDriverTimeoutException)
                {
                    Console.WriteLine($"יום {day} לא נמצא בלוח השנה המוצג או לא נטען בזמן. שגיאה: {ex.Message}");
                }
            }
        }
    }

}
