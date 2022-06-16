using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace GPSBIM
{
    public partial class Ribbon1
    {
        Excel.Application excelApp;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            excelApp = Globals.ThisAddIn.Application;
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)//编号排序
        {
            int j = 1;
            int tempNum = 1;
            int modifyNum = 0;
            foreach (Excel.Range rang in excelApp.Selection)
            {
                string cellContent = excelApp.Cells[rang.Row, rang.Column].Value;
                string gpsCode = "01";
                if (cellContent != null && cellContent != "")
                {
                    gpsCode = cellContent.Replace(Regex.Replace(cellContent, "[0-9]", "", RegexOptions.IgnoreCase),"");
                    tempNum=Convert.ToInt32(gpsCode);
                    break;
                }
            }

            foreach (Excel.Range rang in excelApp.Selection)
            {
                string cellContent = excelApp.Cells[rang.Row, rang.Column].Value;
                string modifyCellContent;
                if (cellContent != null && cellContent != "")
                {
                    modifyCellContent = Regex.Replace(cellContent, "[0-9]", "", RegexOptions.IgnoreCase) + tempNum.ToString().PadLeft(2, '0');
                    excelApp.Cells[rang.Row, rang.Column].Value = modifyCellContent;
                    tempNum++;
                    modifyNum++;
                }

                j++;
                if (j == 3500)
                {
                    break;
                }
            }
            MessageBox.Show("成功修改" + modifyNum.ToString() + "处!", "GPSBIM", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button2_Click(object sender, RibbonControlEventArgs e)//子项编号施工图中文
        {
            string cellContent = excelApp.Cells[2, 4].Value;
            string[] sArray = cellContent.Split('-');
            string modifyCellContent = sArray.ElementAt(0);
            // MessageBox.Show(modifyCellContent);        

            int j = 0;
            for (int i = 7; i < 3500; i++)
            {
                string cell = Convert.ToString(excelApp.Cells[i, 1].Value);

                if (cell != null && cell != "" && (FilterCH(cell) == "" || FilterCH(cell) == null))
                {
                    excelApp.Cells[i, 1].Value = modifyCellContent;
                    j++;
                }
            }

            Excel.Worksheet eltSheet = excelApp.ActiveWorkbook.ActiveSheet;
            eltSheet.Name = modifyCellContent;

            string[] pArray = excelApp.Cells[1, 4].Value.Split('-');
            string projectNum = pArray.ElementAt(0);
            eltSheet.PageSetup.LeftFooter = "&\"Arial\"" + "&16" + " " + projectNum + "-" + modifyCellContent + "-" + "WD-ELT";

            excelApp.Cells[3, 4].Value = "施工图";

            MessageBox.Show("成功修改子项车间编号" + j.ToString() + "处!" + "\r" + "\r" + "成功修改页脚！", "GPSBIM", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button3_Click(object sender, RibbonControlEventArgs e)//子项编号投标中文
        {
            string cellContent = excelApp.Cells[2, 4].Value;
            string[] sArray = cellContent.Split('-');
            string modifyCellContent = sArray.ElementAt(0);
            // MessageBox.Show(modifyCellContent);        

            int j = 0;
            for (int i = 7; i < 3500; i++)
            {
                string cell = Convert.ToString(excelApp.Cells[i, 1].Value);

                if (cell != null && cell != "" && (FilterCH(cell) == "" || FilterCH(cell) == null))
                {
                    excelApp.Cells[i, 1].Value = modifyCellContent;
                    j++;
                }
            }

            Excel.Worksheet eltSheet = excelApp.ActiveWorkbook.ActiveSheet;
            eltSheet.Name = modifyCellContent;

            string[] pArray = excelApp.Cells[1, 4].Value.Split('-');
            string projectNum = pArray.ElementAt(0);
            eltSheet.PageSetup.LeftFooter = "&\"Arial\"" + "&16" + " " + modifyCellContent + "-" + "TWD-ELT";

            excelApp.Cells[3, 4].Value = "投标";

            MessageBox.Show("成功修改子项车间编号" + j.ToString() + "处!" + "\r" + "\r" + "成功修改页脚！", "GPSBIM", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void button4_Click(object sender, RibbonControlEventArgs e)//子项编号施工图中英文
        {
            string cellContent = excelApp.Cells[3, 4].Value;
            string[] sArray = cellContent.Split('-');
            string modifyCellContent = sArray.ElementAt(0);
            // MessageBox.Show(modifyCellContent);        

            int j = 0;
            for (int i = 9; i < 3500; i++)
            {
                string cell = Convert.ToString(excelApp.Cells[i, 1].Value);

                if (cell != null && cell != "" && (FilterCH(cell) == "" || FilterCH(cell) == null))
                {
                    excelApp.Cells[i, 1].Value = modifyCellContent;
                    j++;
                }
            }

            Excel.Worksheet eltSheet = excelApp.ActiveWorkbook.ActiveSheet;
            eltSheet.Name = modifyCellContent;

            string[] pArray = excelApp.Cells[1, 4].Value.Split('-');
            string projectNum = pArray.ElementAt(0);
            eltSheet.PageSetup.LeftFooter = "&\"Arial\"" + "&16" + " " + projectNum + "-" + modifyCellContent + "-" + "WD-ELT";

            excelApp.Cells[5, 4].Value = "DETAIL DESIGN";

            MessageBox.Show("成功修改子项车间编号" + j.ToString() + "处!" + "\r" + "\r" + "成功修改页脚！", "GPSBIM", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)//子项编号投标中英文
        {
            string cellContent = excelApp.Cells[3, 4].Value;
            string[] sArray = cellContent.Split('-');
            string modifyCellContent = sArray.ElementAt(0);
            // MessageBox.Show(modifyCellContent);        

            int j = 0;
            for (int i = 9; i < 3500; i++)
            {
                string cell = Convert.ToString(excelApp.Cells[i, 1].Value);

                if (cell != null && cell != "" && (FilterCH(cell) == "" || FilterCH(cell) == null))
                {
                    excelApp.Cells[i, 1].Value = modifyCellContent;
                    j++;
                }
            }

            Excel.Worksheet eltSheet = excelApp.ActiveWorkbook.ActiveSheet;
            eltSheet.Name = modifyCellContent;

            string[] pArray = excelApp.Cells[1, 4].Value.Split('-');
            string projectNum = pArray.ElementAt(0);
            eltSheet.PageSetup.LeftFooter = "&\"Arial\"" + "&16" + " " + modifyCellContent + "-" + "TWD-ELT";

            excelApp.Cells[5, 4].Value = "BID";

            MessageBox.Show("成功修改子项车间编号" + j.ToString() + "处!" + "\r" + "\r" + "成功修改页脚！", "GPSBIM", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        public string FilterEN(string inputValue)
        {
            if (Regex.IsMatch(inputValue, "[A-Za-z0-9\u9fa5-]+"))
            {
                return Regex.Match(inputValue, "[A-Za-z0-9\u9fa5-]+").Value;
            }
            return "";
        }
        public string FilterCH(string inputValue)
        {
            if (Regex.IsMatch(inputValue, "[\u4e00-\u9fa5]+"))
            {
                return Regex.Match(inputValue, "[\u4e00-\u9fa5]+").Value;
            }
            return "";
        }
       
    }
}
