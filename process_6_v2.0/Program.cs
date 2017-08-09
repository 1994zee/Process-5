using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using process_5_v2._0.Model;
using process_6_v2._0.Model;
using System.IO;

namespace process_5_v2._0
{

    class Program
    {
        public static void ErrorLogging(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "xeeshan.ah@gmail.com", "Error log", ex.Message, strPath);
        }
        public static void ErrorLogging1(Exception ex)
        {
            string strPath = "Log.txt";
            if (!File.Exists(strPath))
            {
                File.Create(strPath).Dispose();
            }
            using (StreamWriter sw = File.AppendText(strPath))
            {
                sw.WriteLine("=============Error Logging ===========");
                sw.WriteLine("===========Start============= " + DateTime.Now);
                sw.WriteLine("Error Message: " + ex.Message);
                sw.WriteLine("Stack Trace: " + ex.StackTrace);
                sw.WriteLine("===========End============= " + DateTime.Now);
            }
            Emailing.Email.SendEmail("lightbot@lightsourcehr.com", "ba@lightsourcehr.com", "Error log", ex.Message, strPath);
        }

        static void Main(string[] args)
        {
            int tries = 3;
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            List<Record> Records = new List<Record>();//for storing the records from report 1;
            List<Record> withEmployeeID = new List<Record>();
            List<Record> withoutEmployeeID = new List<Record>();
            List<Record> importable = new List<Record>();
            List<Record> cases = new List<Record>();
            ChromeOptions options = new ChromeOptions();
            List<claim2Records> claimRecords = new List<claim2Records>();
            List<JoiningRecords> finalImportData = new List<JoiningRecords>();
            List<JoiningRecords> laterImportData = new List<JoiningRecords>();
            options.AddArgument("--start-maximized");
            wait:
            Console.Write("application on wait !");
            while (true)
            {
                String wait = DateTime.Now.Hour.ToString() + DateTime.Now.Minute.ToString();
                if (wait == "2350")
                {
                    break;
                }
            }
            Records.Clear();
            withEmployeeID.Clear();
            withoutEmployeeID.Clear();
            importable.Clear();
            cases.Clear();
            claimRecords.Clear();
            finalImportData.Clear();
            laterImportData.Clear();    
            IWebDriver gc = new ChromeDriver(options);
            //loggin into client space
            try
            {
                //...loging in 
                gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/Login");
                gc.FindElement(By.Name("LoginID")).SendKeys("lightbot");
                gc.FindElement(By.Name("Password")).SendKeys("RPAuser!");
                gc.FindElement(By.Name("Password")).SendKeys(Keys.Enter);
                //.............
            }
            catch (Exception ex)
            {
                ErrorLogging(ex);
                Exception e = new Exception("Login failed");
                ErrorLogging(e);
            }
            //....Login end

            //opening claim 1 report
            try
            {
                //WcClaim1 report url
                string reporturl = @"https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\WCCLAIM1+Ancillary+Risk+Fees";
                gc.Navigate().GoToUrl(reporturl);
                System.Threading.Thread.Sleep(5000);
            }
            catch (Exception ex)
            {
                Exception e = new Exception("WCC1 report opening failed !");
                ErrorLogging(ex);
                ErrorLogging(e);
            }
            //..opening claim 1 ends here

            // dump section to be removed before deployment
            step1:
            try
            {
                System.Threading.Thread.Sleep(9000);
                gc.FindElement(By.XPath("//*[@id='ndbfc0']")).Click();
                gc.FindElements(By.TagName("option"))[3].Click();
                gc.FindElement(By.Id("updateBtnP")).Click();
                System.Threading.Thread.Sleep(10000);
                //dump section ends
            }
            catch
            {
                goto step1;
            }
            //..scrapping record from report 1
            try
            {

                Console.WriteLine("Scrapping data :");
                //getting reports from alteranating rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text,cols.ElementAt(8).Text,cols.ElementAt(9).Text);
                    Records.Add(r);
                }
                //getting records from normal rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    Record r = new Record(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text,cols.ElementAt(8).Text,cols.ElementAt(9).Text);
                    Records.Add(r);
                }
                if (Records.Count == 0)
                {
                    goto end;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("no data found");
                Exception e = new Exception("WCC1 Scrapping failed !");
                ErrorLogging(ex);
                ErrorLogging(e);
                goto end;
            }
            //scrapping records ends here

            //printing all records
            Console.WriteLine("Number of records Scrapped :" + Records.Count);
            Console.WriteLine("Filtering data over Employee ID...");
            //filtering records on client ID
            try
            {
                foreach (Record r in Records)
                {
                    if (r.EmployeeID == "")
                           withoutEmployeeID.Add(r);
                    else
                        withEmployeeID.Add(r);
                }
            }
            catch
            {

            }
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Records with Employee ID: " + withEmployeeID);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Records without Employee ID: " + withoutEmployeeID);
            Console.ForegroundColor = ConsoleColor.Gray;
            if (withoutEmployeeID.Count > 0)
            {

                Console.WriteLine("Filtering data without Client ID on location and employee-name...");
                //filtering records without client id on base of employee name  comment
                foreach (Record r in withoutEmployeeID)
                {
                    if (r.EmployeeNameComment != "" || r.Location != "")
                        importable.Add(r);
                    else
                        cases.Add(r);
                }
            }
            importable.AddRange(withEmployeeID); //adds all records with client id
            //...

            //creating cases for non importable records
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Number of records in importable :" + importable.Count);
            Console.ForegroundColor = ConsoleColor.Red;
            Console.WriteLine("Number of records in cases : " + cases.Count);
            if (cases.Count > 0)
            {
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine("Creating cases for " + cases.Count + " number of records.");
                Console.ForegroundColor = ConsoleColor.Blue;
                int i = 1;
                foreach (Record record in cases)
                {
                    Console.WriteLine("\nCreating case # " + i);
                    i++;
                    try
                    {
                        gc.Navigate().GoToUrl("https://cwp.clientspace.net/Next/peo/client");
                        gc.FindElement(By.Id("dropdownMenu1")).Click();
                        gc.FindElement(By.Id("All")).Click();
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ContractStage")).SendKeys(Keys.Backspace);
                        gc.FindElement(By.Name("ClientNumber")).SendKeys(record.ClientID.ToString()); // put the client number here.
                        System.Threading.Thread.Sleep(2000);
                        gc.FindElement(By.ClassName("formSearchBtn")).Click();
                        System.Threading.Thread.Sleep(1000);
                        gc.FindElements(By.ClassName("cs-underline"))[1].Click();
                        gc.FindElement(By.XPath("//*[@id='customHeaderHtml']/div[2]/div[6]/div/div[1]/table/tbody/tr/td[1]/span")).Click();
                        gc.FindElement(By.XPath("//*[@id='lstDataform_btnAdd']")).Click();
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys("R");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='crCategory']")).SendKeys(Keys.Tab);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys("M");
                        System.Threading.Thread.Sleep(1500);
                        gc.FindElement(By.XPath("//*[@id='fkCaseTypeID']")).SendKeys(Keys.Enter);
                        gc.FindElement(By.XPath("//*[@id='CallerName']")).SendKeys("Lightbot");
                        gc.FindElement(By.XPath("//*[@id='EmailAddress']")).SendKeys("lightbot@lightsourcehr.com");
                        DateTime dateTime = DateTime.Now;
                        var date = dateTime.AddDays(1).ToString("MM/dd/yyyy");
                        gc.FindElement(By.XPath("//*[@id='DueDate']")).SendKeys(date);

                        System.Threading.Thread.Sleep(10000);
                        if (record.CarrierClaimNumber == "")
                        {
                            gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys(record.CsClaimNumber.ToString());
                        }
                        else
                        {
                            gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys(record.CarrierClaimNumber.ToString());
                        }
                        gc.FindElement(By.XPath("//*[@id='Subject']")).SendKeys(record.ClientID.ToString());

                        gc.FindElement(By.XPath("//*[@id='btnSave']")).Click();
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Result : Case creation Sucessfull.");

                    }
                    catch (Exception ex)
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("Result : Case creation failed..!");
                        Exception e = new Exception("Case Reporting failed !");
                        ErrorLogging(ex);
                        ErrorLogging(e);
                    }
                }
            }
            //Step 2
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Opening WC Claim 2 report...");
            step2:
            try
            {
                string reporturl = @"https://cwp.clientspace.net/BusinessIntelligence/ReportViewer.aspx?rn=LightBot+Admins+Only\WCCLAIM2+Client+Fees";
                gc.Navigate().GoToUrl(reporturl);
                System.Threading.Thread.Sleep(10000);
            }
            catch (Exception ex)
            {
                tries--;
                if(tries>0)
                {
                    goto step2;
                }
                tries = 3;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Opening Report WCCLAIM 2 failed!");
                    
                Exception e = new Exception("WCC2 report opening failed !");
                ErrorLogging(ex);
                ErrorLogging(e);
                goto end;
            }
            //scrapping data from report#2
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Scrapping data from second report...");
            try
            {
                foreach (IWebElement c in gc.FindElements(By.ClassName("AlternatingItem")))
                {
                    //Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    claim2Records r = new claim2Records(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text, cols.ElementAt(8).Text);
                    claimRecords.Add(r);
                }
                //getting records from normal rows
                foreach (IWebElement c in gc.FindElements(By.ClassName("ReportItem")))
                {
                    //Console.WriteLine("Data fount !");
                    ICollection<IWebElement> cols = c.FindElements(By.TagName("td"));
                    claim2Records r = new claim2Records(cols.ElementAt(0).Text, cols.ElementAt(1).Text, cols.ElementAt(2).Text, cols.ElementAt(3).Text, cols.ElementAt(4).Text, cols.ElementAt(5).Text, cols.ElementAt(6).Text, cols.ElementAt(7).Text, cols.ElementAt(8).Text);
                    claimRecords.Add(r);
                }
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Extraction of records from Report#2 successful");
                Console.ForegroundColor = ConsoleColor.Gray;
                Console.WriteLine(claimRecords.Count + " number of records extracted.");


            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Extraction of records from report two unsuccessfull."); Console.ForegroundColor = ConsoleColor.Red;
                Exception e = new Exception("WCC2 Scrapping failed !");
                ErrorLogging(ex);
                ErrorLogging(e);
                if (claimRecords.Count == 0)
                {
                    goto end;
                }

            }

            //joining record
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Joining records of both reports");
            foreach (Record r in importable)
            {
                claim2Records temp = null;
                foreach (claim2Records tem in claimRecords.FindAll(s => s.ClientID == r.ClientID))
                {
                    if (temp == null)
                    {
                        temp = tem;
                    }
                    else
                    {
                        int x = Convert.ToInt32(temp.BillRates.ToString());
                        int y = Convert.ToInt32(tem.BillRates.ToString());
                        if (x > y)
                        {
                            temp = tem;
                        }
                    }
                }
                if (temp != null)
                {
                    try
                    {
                        temp.BillRates = temp.BillRates.Replace(",", "");
                        temp.BillRates = temp.BillRates.Split('.')[0];
                        if (Convert.ToInt32(temp.BillRates.ToString()) >= 2500)
                        {
                            Console.WriteLine("adeed to later import"+r.EmployeeID+" "+temp.BillRates);
                            JoiningRecords rec = new JoiningRecords(r, temp);
                            laterImportData.Add(rec);
                        }
                        else
                        {
                            Console.WriteLine("added to import for now");
                            JoiningRecords rec = new JoiningRecords(r, temp);
                            finalImportData.Add(rec);
                        }
                    }
                    catch
                    {
                        Console.WriteLine(temp.BillRates);
                    }
                }
                temp = null;
            }

            //creating importable report
            Console.WriteLine("Number of records importable : " + importable.Count);
            Console.WriteLine("Number of records later importable : " + laterImportData.Count);
            string delimiter = "\t";
            if (importable.Count > 0)
            {
                try
                {
                    using (var writer = new System.IO.StreamWriter(basePath + "\\data.txt"))
                    {
                        foreach (JoiningRecords i in finalImportData)
                        {
                            if (i.WC1record.EmployeeNameComment == "")
                            {
                                i.WC1record.EmployeeNameComment = i.WC1record.InjuredEmployee;
                            }
                            writer.WriteLine(i.WC1record.BillDate + delimiter + i.WC1record.BillEventCode + delimiter + i.WC2record.BillRates + delimiter + i.WC1record.BillUnits + delimiter + i.WC1record.ClientID + delimiter + i.WC1record.EmployeeID + delimiter + i.WC1record.Location + delimiter + delimiter + i.WC1record.EmployeeNameComment);
                        }
                    }
                }
                catch
                {
                    Console.WriteLine("Import file creation failed.");
                }
            }
            
            try
            {
                if (laterImportData.Count > 0)
                {
                    string filename = DateTime.Now.AddDays(7).ToString("M/d/yyyy");
                    filename = filename.Replace('/', '_');
                    using (var writer = new System.IO.StreamWriter(basePath + "\\" + filename + ".txt"))
                    {
                        foreach (JoiningRecords i in laterImportData)
                        {
                            if(i.WC1record.EmployeeNameComment=="")
                            {
                                i.WC1record.EmployeeNameComment = i.WC1record.InjuredEmployee;
                            }
                            writer.WriteLine(i.WC1record.BillDate + delimiter + i.WC1record.BillEventCode + delimiter + i.WC2record.BillRates + delimiter + i.WC1record.BillUnits + delimiter + i.WC1record.ClientID + delimiter + i.WC1record.EmployeeID + delimiter+i.WC1record.Location+delimiter+delimiter+i.WC1record.EmployeeNameComment);
                        }
                    }
                }
            }
            catch
            {
                Console.WriteLine("Later import failed.");
            }
            if (importable.Count ==0)
            {
                goto end;
            }
            top:
            //importing into prism
            gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
            System.Threading.Thread.Sleep(1000);
            try
            {
                gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
                gc.FindElement(By.XPath("//*[@id='button8v1']")).Click();
                System.Threading.Thread.Sleep(1000);

            }
            catch (Exception ex)
            {
                Exception e = new Exception("Prism Login failed !");
                ErrorLogging(ex);
                ErrorLogging(e);
                goto end;
            }
            try
            {
                gc.FindElement(By.XPath("//*[@id='text31v1']")).Click();
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("C");
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("c");
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("l");

                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("i");

                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("e");
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("n");
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("t");

                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(" bill");
                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);
                System.Threading.Thread.Sleep(1000);

                gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Enter);
                System.Threading.Thread.Sleep(10000);

            }
            catch (Exception ex)
            {

                Exception e = new Exception("Client Bill Pending report opening failed !");
                ErrorLogging(ex);
                ErrorLogging(e);
                goto end;
            }
            System.Threading.Thread.Sleep(1000);
            gc.FindElement(By.XPath("//*[@id='button31v2']")).Click();
            System.Threading.Thread.Sleep(1000);
            String current = gc.CurrentWindowHandle;
            foreach (String winHandle in gc.WindowHandles)
            {
                gc.SwitchTo().Window(winHandle);
            }
            //sometimes the upload window doesn't open
            if (gc.CurrentWindowHandle != current)
            {
                try
                {
                    gc.FindElement(By.XPath("//*[@id='fname']")).SendKeys(basePath + "\\data.txt"); //put the full path of file here

                    gc.FindElement(By.XPath("//*[@id='submit1']")).Click();

                    System.Threading.Thread.Sleep(1000);
                    gc.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();
                    System.Threading.Thread.Sleep(20000);
                    gc.SwitchTo().Window(current);
                }
                catch
                {
                    try
                    {
                        gc.Close();
                    }
                    catch
                    {
                        
                    }
                    try
                    {
                        gc.SwitchTo().Window(current);
                        gc.Close();
                    }
                    catch
                    {

                    }
                    goto end;
                }
                try
                {
                    Exception s = new Exception(gc.FindElement(By.XPath("//*[@id='body_span29v2']")).Text);
                    ErrorLogging1(s);

                }
                catch
                {

                }

                try
                {
                    gc.FindElement(By.XPath("//*[@id='button33v2']")).Click();
                    System.Threading.Thread.Sleep(2000);
                    gc.FindElement(By.XPath("//*[@id='button32v2']")).Click();
                    System.Threading.Thread.Sleep(2000);
                    gc.SwitchTo().Alert().Accept();
                    gc.SwitchTo().Window(current);
                    gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                    gc.SwitchTo().Window(current);
                    gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();

                }
                tryagain:
                try
                {

                    gc.Close();
                    gc.Dispose();
                    Console.WriteLine("Process Complete..!");

                }
                catch (Exception ex)
                {
                    Exception e = new Exception("Chrome closing failed failed !");
                    ErrorLogging(ex);
                    ErrorLogging(e);
                    System.Threading.Thread.Sleep(10000);
                    goto tryagain;
                }

            }
            else
            {
                tryagain:
                try
                {
                    gc.Close();
                    gc.Dispose();
                    goto top;
                }
                catch (Exception ex)
                {
                    Exception e = new Exception("Chrome closing failed failed !");
                    ErrorLogging(ex);
                    ErrorLogging(e);
                    System.Threading.Thread.Sleep(10000);
                    goto tryagain;
                }
            }
            end:
            try
            {
                string filename1 = DateTime.Now.ToString("M/d/yyyy");
                if (File.Exists(basePath + "/" + filename1))
                {
                    Console.WriteLine("Importable file exists...!");
                    Console.WriteLine("trying to import now..");
                    gc.Navigate().GoToUrl("https://ctw.prismhr.com/ctw/dbweb.asp?dbcgm=1");
                    System.Threading.Thread.Sleep(1000);
                    try
                    {
                        gc = new ChromeDriver();
                        gc.FindElement(By.XPath("//*[@id='text4v1']")).SendKeys("lightbot");
                        gc.FindElement(By.XPath("//*[@id='password6v1']")).SendKeys("RPAuser1!");
                        gc.FindElement(By.XPath("//*[@id='button8v1']")).Click();
                        System.Threading.Thread.Sleep(1000);

                    }
                    catch (Exception ex)
                    {
                        Exception e = new Exception("Prism Login failed !");
                        ErrorLogging(ex);
                        ErrorLogging(e);
                        goto end2;
                    }
                    try
                    {
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).Click();
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("C");
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("c");
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("l");

                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("i");

                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("e");
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("n");
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys("t");

                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(" bill");
                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);

                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Backspace);
                        System.Threading.Thread.Sleep(1000);

                        gc.FindElement(By.XPath("//*[@id='text31v1']")).SendKeys(Keys.Enter);
                        System.Threading.Thread.Sleep(10000);

                    }
                    catch (Exception ex)
                    {

                        Exception e = new Exception("Client Bill Pending report opening failed !");
                        ErrorLogging(ex);
                        ErrorLogging(e);
                        goto end;
                    }
                    System.Threading.Thread.Sleep(1000);
                    gc.FindElement(By.XPath("//*[@id='button31v2']")).Click();
                    System.Threading.Thread.Sleep(1000);
                    String current1 = gc.CurrentWindowHandle;
                    foreach (String winHandle in gc.WindowHandles)
                    {
                        gc.SwitchTo().Window(winHandle);
                    }
                    //sometimes the upload window doesn't open
                    if (gc.CurrentWindowHandle != current1)
                    {
                        try
                        {
                            gc.FindElement(By.XPath("//*[@id='fname']")).SendKeys(basePath +"\\"+filename1+ ".txt"); //put the full path of file here

                            gc.FindElement(By.XPath("//*[@id='submit1']")).Click();

                            System.Threading.Thread.Sleep(1000);
                            gc.FindElement(By.XPath("//*[@id='BUTTON1']")).Click();
                            System.Threading.Thread.Sleep(20000);
                            gc.SwitchTo().Window(current1);
                        }
                        catch
                        {
                            try
                            {
                                gc.Close();
                            }
                            catch
                            {

                            }
                            try
                            {
                                gc.SwitchTo().Window(current1);
                                gc.Close();
                            }
                            catch
                            {

                            }
                            goto end2;
                        }
                        try
                        {
                            Exception s = new Exception(gc.FindElement(By.XPath("//*[@id='body_span29v2']")).Text);
                            ErrorLogging1(s);

                        }
                        catch
                        {

                        }

                        try
                        {
                            gc.FindElement(By.XPath("//*[@id='button33v2']")).Click();
                            System.Threading.Thread.Sleep(2000);
                            gc.FindElement(By.XPath("//*[@id='button32v2']")).Click();
                            System.Threading.Thread.Sleep(2000);
                            gc.SwitchTo().Alert().Accept();
                            gc.SwitchTo().Window(current1);
                            gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e);
                            gc.SwitchTo().Window(current1);
                            gc.FindElement(By.XPath("//*[@id='button35v2']")).Click();

                        }
                        tryagain:
                        try
                        {

                            gc.Close();
                            gc.Dispose();
                            Console.WriteLine("Process Complete..!");

                        }
                        catch (Exception ex)
                        {
                            Exception e = new Exception("Chrome closing failed failed !");
                            ErrorLogging(ex);
                            ErrorLogging(e);
                            System.Threading.Thread.Sleep(10000);
                            goto tryagain;
                        }

                    }
                    else
                    {
                        tryagain:
                        try
                        {
                            gc.Close();
                            gc.Dispose();
                            goto top;
                        }
                        catch (Exception ex)
                        {
                            Exception e = new Exception("Chrome closing failed failed !");
                            ErrorLogging(ex);
                            ErrorLogging(e);
                            System.Threading.Thread.Sleep(10000);
                            goto tryagain;
                        }
                    }
                }
            }
            catch
            {

            }
            end2:
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("Process complete");
            goto wait;
        }
    }
}
