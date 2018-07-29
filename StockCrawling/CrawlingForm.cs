using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Threading;

namespace StockCrawling
{
    public partial class CrawlingForm : Form
    {
        IWebDriver driver;
        string savFileName;
        string savInvestOpinionFileName;

        List<string> listCode = new List<string>();

        public CrawlingForm()
        {
            InitializeComponent();

            driver = new ChromeDriver();
            savFileName = Application.StartupPath + @"\CollectStock.xlsx";
            savInvestOpinionFileName = Application.StartupPath + @"\InvestOpinion.xlsx";
        }

        private void CrawlingKOSPI()
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            excelApp = new Excel.Application();

            wb = excelApp.Workbooks.Add();

            // 첫번째 Worksheet
            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

            Excel.Range cells = ws.Cells;
            cells.NumberFormat = "@";




            int row = 2;
            for(int page=1; page<31; ++page)
            {
                driver.Url = "https://finance.naver.com/sise/sise_market_sum.nhn?&page=" + page;

                IReadOnlyCollection<IWebElement> elems = driver.FindElements(By.XPath("//div[@class='box_type_l']//table[@class='type_2']/tbody/tr"));
                List<IWebElement> listElems = elems.ToList<IWebElement>();

                foreach (IWebElement elem in listElems)
                {
                    string value = elem.Text;

                    if (value != "")
                    {
                        string[] words = value.Split(' ');

                        string href = elem.FindElement(By.PartialLinkText(words[1])).GetAttribute("href");
                        string[] hrefs = href.Split('=');

                        // 종목 Code
                        ws.Cells[row, 1] = hrefs[1];

                        // 나머지 데이터
                        int col = 2;
                        foreach (string word in words)
                        {
                            ws.Cells[row, col] = word;
                            col++;
                        }

                        row++;
                    }
                }

                Thread.Sleep(3000);
            }





            // 엑셀파일 저장
            if (File.Exists(savFileName))
            {
                File.Delete(savFileName);
            }

            wb.SaveAs(savFileName);

            wb.Close(true);
            excelApp.Quit();
        }

        private void InvestOpinion()
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            excelApp = new Excel.Application();

            wb = excelApp.Workbooks.Add();

            // 첫번째 Worksheet
            ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

            Excel.Range cells = ws.Cells;
            cells.NumberFormat = "@";






            int row = 2;
            foreach (string code in listCode)
            {
                try
                {
                    driver.Url = "https://finance.naver.com/item/coinfo.nhn?code=" + code;
                    driver.SwitchTo().Frame("coinfo_cp");

                    IReadOnlyCollection<IWebElement> elems = driver.FindElements(By.XPath("//table[@class='gHead all-width']/tbody/tr"));

                    List<IWebElement> listElems = elems.ToList<IWebElement>();

                    if (listElems.Count > 1)
                    {
                        string value = listElems[1].Text;
                        string[] words = value.Split(' ');

                        // 종목 Code
                        ws.Cells[row, 1] = code;

                        // 나머지 데이터
                        int col = 2;
                        foreach (string word in words)
                        {
                            ws.Cells[row, col] = word;
                            col++;
                        }

                        Thread.Sleep(1000);
                    }
                }
                catch (Exception ex)
                {

                }

                row++;
            }






            // 엑셀파일 저장
            if (File.Exists(savInvestOpinionFileName))
            {
                File.Delete(savInvestOpinionFileName);
            }

            wb.SaveAs(savInvestOpinionFileName);

            wb.Close(true);
            excelApp.Quit();

            /*
            // 투자의견 목표주가(원) EPS(원)  PER(배)  추정기관수
            IWebElement e1 = driver.FindElement(By.XPath("//td[@class='noline-bottom line-right center cUp']"));
            IWebElement e2 = driver.FindElement(By.XPath("//td[@class='noline-bottom line-right center']"));
            IWebElement e3 = driver.FindElement(By.XPath("//td[@class='noline-bottom line-right center']"));
            IWebElement e4 = driver.FindElement(By.XPath("//td[@class='noline-bottom line-right center']"));
            IWebElement e5 = driver.FindElement(By.XPath("//td[@class='noline-bottom center']"));
            */

            MessageBox.Show("투자 의견 저장 완료!");
        }

        private void CreateExcelData()
        {
            // 신규 파일 생성시
            List<string> testData = new List<string>()
            { "Excel", "Access", "Word", "OneNote" };

            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();

                wb = excelApp.Workbooks.Add();

                // 첫번째 Worksheet
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                // 데이타 넣기
                int r = 1;
                foreach (var d in testData)
                {
                    ws.Cells[r, 1] = d;
                    r++;
                }

                // 엑셀파일 저장
                if (File.Exists(savFileName))
                {
                    File.Delete(savFileName);
                }

                wb.SaveAs(savFileName);

                wb.Close(true);
                excelApp.Quit();
            }
            finally
            {
                // Clean up
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }

        public void ReadExcelData()
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();

                // 엑셀 파일 열기
                wb = excelApp.Workbooks.Open(savFileName);

                // 첫번째 Worksheet
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                // 현재 Worksheet에서 사용된 Range 전체를 선택
                Excel.Range rng = ws.UsedRange;

                // 현재 Worksheet에서 일부 범위만 선택
                // Excel.Range rng = ws.Range[ws.Cells[2, 1], ws.Cells[5, 3]];

                // Range 데이타를 배열 (One-based array)로
                object[,] data = rng.Value;

                for (int r = 2; r <= data.GetLength(0); r++)
                {
                    try
                    {
                        string code = data[r, 1].ToString();
                        listCode.Add(code);
                    }
                    catch (Exception ex)
                    {
                    }

                    /*
                    for (int c = 1; c <= data.GetLength(1); c++)
                    {
                        //Debug.Write(data[r, c].ToString() + " ");
                    }
                    */
                }

                wb.Close(true);
                excelApp.Quit();
            }
            finally
            {
                // Clean up
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }

            MessageBox.Show("엑셀 읽기 완료!");
        }

        private void ReleaseExcelObject(object obj)
        {
            try
            {
                if (obj != null)
                {
                    Marshal.ReleaseComObject(obj);
                    obj = null;
                }
            }
            catch (Exception ex)
            {
                obj = null;
                throw ex;
            }
            finally

            {
                GC.Collect();
            }
        }

        private void CrawlingForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                driver.Quit();
            }
            catch (Exception ex)
            {

            }
        }

        private void btnCrawlingKOSPI_Click(object sender, EventArgs e)
        {
            CrawlingKOSPI();
        }

        private void btnReadExcelData_Click(object sender, EventArgs e)
        {
            ReadExcelData();
        }

        private void btnReadInvestOpinion_Click(object sender, EventArgs e)
        {
            InvestOpinion();
        }
    }
}
