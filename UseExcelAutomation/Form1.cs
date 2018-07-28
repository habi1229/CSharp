using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;

namespace UseExcelAutomation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }

        private static void WriteExcelData()
        {
            Excel.Application excelApp = null;
            Excel.Workbook wb = null;
            Excel.Worksheet ws = null;

            try
            {
                excelApp = new Excel.Application();

                // 엑셀 파일을 엽니다.
                wb = excelApp.Workbooks.Open(@".\test.xlsx");

                // 첫번째 Worksheet를 선택합니다.
                ws = wb.Worksheets.get_Item(1) as Excel.Worksheet;

                // 해당 Worksheet에서 저장할 범위를 정합니다.
                // 지금은 저장할 행렬의 크기만큼 지정합니다.
                // 예시 Excel.Range rng = ws.Range["B2", "G8"];
                int row = 0;
                int column = 0;
                Excel.Range rng = ws.Range[ws.Cells[1, 1], ws.Cells[row, column]];
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                ReleaseExcelObject(ws);
                ReleaseExcelObject(wb);
                ReleaseExcelObject(excelApp);
            }
        }

        private static void ReleaseExcelObject(object obj)
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
    }
}
