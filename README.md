# data-processing-scripts
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace WindowsFormsApplication8
{
    public partial class Form1 : Form
    {
        Excel.Application xlApp;
        Excel.Workbook xlWorkBook;
        Excel.Worksheet xlWorkSheet;
        object misValue = System.Reflection.Missing.Value;

        Excel.Application xlApp2;
        Excel.Workbook xlWorkBook2;
        Excel.Worksheet xlWorkSheet2;
        object misValue2 = System.Reflection.Missing.Value;
        public Form1()
        {
            InitializeComponent();
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {

            //System.Globalization.CultureInfo oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            //System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-GB");
            int NROW1 = 1; int NROWt = 1; int NROWt1 = 1; int NROW6 = 1; int NROW36 = 1; int NROW46 = 1; int NROW56 = 1; int NROW66 = 1; int NROW76 = 1; int NROW86 = 1; int NCOL2 = 1; DateTime dt; double val = 0; DateTime dt1; double val1 = 0; DateTime dt36; double val36 = 0; DateTime dt2;
            bool tesval36; double mthval = 0; double yerval = 0; double hyerval = 0; double myerval = 0; double mhyerval = 0;



            // try
            // {
            string pat = "";
            string pat2 = "C:\\VC_apl\\Templet.xlsx";

            OpenFileDialog open = new OpenFileDialog();
            open.Filter = "Data Files(*.xls; *.xlsx; *.txt;)|*.xls; *.xlsx; *.txt;";
            if (open.ShowDialog() == DialogResult.OK)
            {
                pat = open.FileName;
            }
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlApp = new Excel.ApplicationClass();
            xlWorkBook = xlApp.Workbooks.Open(pat, 0, false, 1, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, false, 1, 0);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkBook2 = xlApp.Workbooks.Open(pat2, 0, false, 1, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, false, 1, 0);
          
            //============================================================================================================================================================================================================================================
            // xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
            while (!((((Excel.Range)xlWorkSheet.Cells[NROW1, 1]).Value2 == null) && (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
            {
                if (((Excel.Range)xlWorkSheet.Cells[NROW1, 1]).Value2 == null) { NCOL2 = NCOL2 + 1; NROW6 = 3; NROW36 = 3; NROW46 = 3; NROW56 = 3; NROW66 = 3; NROW76 = 3; NROW86 = 3; NROWt = NROW1 + 1; NROWt1 = NROWt + 1; NROW1 = NROWt1 + 1; }// NROW1 = NROW1 + 3; }
                if (NCOL2 > 1)
                {
                    if (NCOL2 == 2)
                    {
                        //--------------------------------------------------------------------------------------------                       

                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1); xlWorkSheet2.Name = "Days";
                        xlWorkSheet2.Cells[NROW6, 1] = xlWorkSheet.Cells[NROW1, 1]; xlWorkSheet2.Cells[NROW6, 2] = xlWorkSheet.Cells[NROW1, 2]; xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = ""; //NROW6 = NROW6 + 1;
                        //--------------------------------------------------------------------------------------------                       
                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(2); xlWorkSheet2.Name = "Days_Date";
                        val = Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 1]).Value2) - 1;
                        dt = Convert.ToDateTime(dateTimePicker1.Value);
                        dt = dt.AddDays(val);
                        xlWorkSheet2.Cells[NROW6, 1] = dt;
                        xlWorkSheet2.Cells[NROW6, 2] = xlWorkSheet.Cells[NROW1, 2]; xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                        //--------------------------------------------------------------------------------------------                                             
                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3); xlWorkSheet2.Name = "Days_Date_Delails";
                        if (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 != null)
                        { val36 = Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2) - 1; tesval36 = true; }
                        else { tesval36 = false; }
                        dt36 = Convert.ToDateTime(dateTimePicker1.Value);
                        val1 = val + 1;
                        if (tesval36 == true)
                        {
                            if (val1 == val36)
                            {
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);
                                dt1 = dt36.AddDays(val);
                                dt2 = dt36.AddDays(val + 1);
                                xlWorkSheet2.Cells[NROW36, 1] = dt1;
                                xlWorkSheet2.Cells[NROW36, 2] = xlWorkSheet.Cells[NROW1, 2];
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4); xlWorkSheet2.Name = "Months";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5); xlWorkSheet2.Name = "Years";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6); xlWorkSheet2.Name = "HidroGiological Years";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7); xlWorkSheet2.Name = "Years Million m3";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [Million m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8); //xlWorkSheet2.Name = "HidroGiological Years Million m3";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [Million m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";

                                if (dt2.Day != 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (dt2.Day == 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                    xlWorkSheet2.Cells[NROW46, 1] = dt2.Month.ToString()+"-" + dt2.Year.ToString();
                                    xlWorkSheet2.Cells[NROW46, 2] = mthval;
                                    mthval = 0;
                                    NROW46 = NROW46 + 1;

                                }//if (dt1.Day != 1)
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                if (!(((dt1.Month == 12) && (dt2.Month == 1))|| (((Excel.Range)xlWorkSheet.Cells[NROW1+1, 1]).Value2 == null)))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);                             
                                    xlWorkSheet2.Cells[NROW56, 1] =  dt2.Year.ToString();
                                    xlWorkSheet2.Cells[NROW56, 2] = yerval;
                                    
                                    NROW56 = NROW56 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                    xlWorkSheet2.Cells[NROW76, 1] = dt2.Year.ToString();
                                    xlWorkSheet2.Cells[NROW76, 2] = yerval/1000000;
                                    NROW76 = NROW76 + 1;
                                    yerval = 0;

                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (((dt1.Month == 9) && (dt2.Month == 10)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                    xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year-1).ToString()+"-"+dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW66, 2] = hyerval;
                                    NROW66 = NROW66 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                    xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW86, 2] = hyerval / 1000000;
                                    NROW86 = NROW86 + 1;
                                    hyerval = 0;
                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                
                                
                                NROW36 = NROW36 + 1;
                            }
                            else
                            {
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);
 
                                dt1 = dt36.AddDays(val);
                                dt2 = dt36.AddDays(val + 1);
                                xlWorkSheet2.Cells[NROW36, 1] = dt1;
                                xlWorkSheet2.Cells[NROW36, 2] = xlWorkSheet.Cells[NROW1, 2];
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                if (dt2.Day != 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (dt2.Day == 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                    xlWorkSheet2.Cells[NROW46, 1] = dt1.Month.ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW46, 2] = mthval;
                                    NROW46 = NROW46 + 1;
                                    mthval = 0;
                                }//if (dt1.Day != 1)
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if ((((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                                    xlWorkSheet2.Cells[NROW56, 1] = dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW56, 2] = yerval;
                                    NROW56 = NROW56 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                    xlWorkSheet2.Cells[NROW76, 1] = dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW76, 2] = yerval / 1000000 ;
                                    NROW76 = NROW76 + 1;
                                    yerval = 0;

                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if ((((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                    xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year-1).ToString()+"-"+dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW66, 2] = hyerval;
                                    NROW66 = NROW66 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                    xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW86, 2] = hyerval / 1000000;
                                    NROW86 = NROW86 + 1;
                                    hyerval = 0;
                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                
                                NROW36 = NROW36 + 1;
                                while (val1 < val36)
                                {
                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);
                                    dt1 = dt36.AddDays(val1);
                                    dt2 = dt36.AddDays(val1 + 1);
                                    xlWorkSheet2.Cells[NROW36, 1] = dt1;
                                    xlWorkSheet2.Cells[NROW36, 2] = xlWorkSheet.Cells[NROW1, 2];

                                    //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                    if (dt2.Day != 1)
                                    {
                                        mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    }
                                    if (dt2.Day == 1)
                                    {
                                        mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                        xlWorkSheet2.Cells[NROW46, 1] = dt1.Month.ToString() + "-" + dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW46, 2] = mthval;
                                        NROW46 = NROW46 + 1;
                                        mthval = 0;

                                    }//if (dt1.Day != 1)
                                    //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                    if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                    {
                                        yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    }
                                    if ((((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                    {
                                        yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                                        xlWorkSheet2.Cells[NROW56, 1] = dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW56, 2] = yerval;
                                        NROW56 = NROW56 + 1;

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                        xlWorkSheet2.Cells[NROW76, 1] = dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW76, 2] = yerval / 1000000;
                                        NROW76 = NROW76 + 1;
                                        yerval = 0;

                                    }
                                    //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                    if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                    {
                                        hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    }
                                    if ((((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                    {
                                        hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                        xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year-1).ToString()+"-"+dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW66, 2] = hyerval;
                                        NROW66 = NROW66 + 1;

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                        xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW86, 2] = hyerval / 1000000;
                                        NROW86 = NROW86 + 1;
                                        hyerval = 0;
                                    }
                                    //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                
                                    val1 = val1 + 1;
                                    NROW36 = NROW36 + 1;
                                }
                            }
                        }
                        else
                        {
 
                            dt1 = dt36.AddDays(val);
                            dt2 = dt36.AddDays(val + 1);
                            xlWorkSheet2.Cells[NROW36, 1] = dt1;
                            xlWorkSheet2.Cells[NROW36, 2] = xlWorkSheet.Cells[NROW1, 2];
                            xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                            //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            if (dt2.Day != 1)
                            {
                                mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                            }
                            if (dt2.Day == 1)
                            {
                                mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                xlWorkSheet2.Cells[NROW46, 1] = dt1.Month.ToString() + "-" + dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW46, 2] = mthval;
                                NROW46 = NROW46 + 1;
                                mthval = 0;

                            }//if (dt1.Day != 1)
                            //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                            {
                                yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                            }
                            if ((((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                            {
                                yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                                xlWorkSheet2.Cells[NROW56, 1] = dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW56, 2] = yerval;
                                NROW56 = NROW56 + 1;

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                xlWorkSheet2.Cells[NROW76, 1] = dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW76, 2] = yerval  / 1000000;
                                NROW76 = NROW76 + 1;
                                yerval = 0;

                            }
                            //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                            if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                            {
                                hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                            }
                            if ((((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                            {
                               // hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year-1).ToString()+"-"+dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW66, 2] = hyerval;
                                NROW66 = NROW66 + 1;

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW86, 2] = hyerval / 1000000;
                                NROW86 = NROW86 + 1;
                                hyerval = 0;
                            }
                            //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                
                        }


                        NROW6 = NROW6 + 1;


                    }//NCOL2==2
                    else
                    {
                        //--------------------------------------------------------------------------------------------                       
                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(1);
                        xlWorkSheet2.Cells[NROW6, NCOL2] = xlWorkSheet.Cells[NROW1, 2]; xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2];
                        //--------------------------------------------------------------------------------------------                       
                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(2);
                        xlWorkSheet2.Cells[NROW6, NCOL2] = xlWorkSheet.Cells[NROW1, 2]; xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2];
                        //--------------------------------------------------------------------------------------------                       
                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);
                        val = Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 1]).Value2) - 1;
                        if (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 != null)
                        { val36 = Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2) - 1; tesval36 = true; }
                        else { tesval36 = false; }
                        dt36 = Convert.ToDateTime(dateTimePicker1.Value);
                        val1 = val + 1;
                        if (tesval36 == true)
                        {
                            if (val1 == val36)
                            {
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);
                                dt1 = dt36.AddDays(val);
                                dt2 = dt36.AddDays(val + 1);
                               // xlWorkSheet2.Cells[NROW36, 1] = dt1;
                                xlWorkSheet2.Cells[NROW36, NCOL2] = xlWorkSheet.Cells[NROW1, 2];
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4); xlWorkSheet2.Name = "Months";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5); xlWorkSheet2.Name = "Years";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6); xlWorkSheet2.Name = "HidroGiological Years";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7); xlWorkSheet2.Name = "Years Million m3";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [Million m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8); //xlWorkSheet2.Name = "HidroGiological Years Million m3";
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = "Rates [Million m^3/Year]"; xlWorkSheet2.Cells[2, 1] = "Time[Year]"; xlWorkSheet2.Cells[2, 1] = "";

                                if (dt2.Day != 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (dt2.Day == 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                  //  xlWorkSheet2.Cells[NROW46, 1] = dt2.Month.ToString() + "-" + dt2.Year.ToString();
                                    xlWorkSheet2.Cells[NROW46, NCOL2] = mthval;
                                    mthval = 0;
                                    NROW46 = NROW46 + 1;

                                }//if (dt1.Day != 1)
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                                   // xlWorkSheet2.Cells[NROW56, 1] = dt2.Year.ToString();
                                    xlWorkSheet2.Cells[NROW56, NCOL2] = yerval;

                                    NROW56 = NROW56 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                   // xlWorkSheet2.Cells[NROW76, 1] = dt2.Year.ToString();
                                    xlWorkSheet2.Cells[NROW76, NCOL2] = yerval / 1000000;
                                    NROW76 = NROW76 + 1;
                                    yerval = 0;

                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (((dt1.Month == 9) && (dt2.Month == 10)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                   // xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW66, NCOL2] = hyerval;
                                    NROW66 = NROW66 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                    //xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW86, NCOL2] = hyerval / 1000000;
                                    NROW86 = NROW86 + 1;
                                    hyerval = 0;
                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY


                                NROW36 = NROW36 + 1;
                            }
                            else
                            {
                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);

                                dt1 = dt36.AddDays(val);
                                dt2 = dt36.AddDays(val + 1);
                               // xlWorkSheet2.Cells[NROW36, 1] = dt1;
                                xlWorkSheet2.Cells[NROW36, NCOL2] = xlWorkSheet.Cells[NROW1, 2];
                                xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                if (dt2.Day != 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if (dt2.Day == 1)
                                {
                                    mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                  //  xlWorkSheet2.Cells[NROW46, 1] = dt1.Month.ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW46, NCOL2] = mthval;
                                    NROW46 = NROW46 + 1;
                                    mthval = 0;
                                }//if (dt1.Day != 1)
                                //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if ((((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                {
                                    yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                                //    xlWorkSheet2.Cells[NROW56, 1] = dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW56, NCOL2] = yerval;
                                    NROW56 = NROW56 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                  //  xlWorkSheet2.Cells[NROW76, 1] = dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW76, NCOL2] = yerval / 1000000;
                                    NROW76 = NROW76 + 1;
                                    yerval = 0;

                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                }
                                if ((((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                {
                                    hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                //    xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW66, NCOL2] = hyerval;
                                    NROW66 = NROW66 + 1;

                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                 //   xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                    xlWorkSheet2.Cells[NROW86, NCOL2] = hyerval / 1000000;
                                    NROW86 = NROW86 + 1;
                                    hyerval = 0;
                                }
                                //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY

                                NROW36 = NROW36 + 1;
                                while (val1 < val36)
                                {
                                    xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(3);
                                    dt1 = dt36.AddDays(val1);
                                    dt2 = dt36.AddDays(val1 + 1);
                                   // xlWorkSheet2.Cells[NROW36, 1] = dt1;
                                    xlWorkSheet2.Cells[NROW36, NCOL2] = xlWorkSheet.Cells[NROW1, 2];

                                    //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                    if (dt2.Day != 1)
                                    {
                                        mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    }
                                    if (dt2.Day == 1)
                                    {
                                        mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                                       // xlWorkSheet2.Cells[NROW46, 1] = dt1.Month.ToString() + "-" + dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW46, NCOL2] = mthval;
                                        NROW46 = NROW46 + 1;
                                        mthval = 0;

                                    }//if (dt1.Day != 1)
                                    //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                                    if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                    {
                                        yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    }
                                    if ((((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                                    {
                                        yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                                      //  xlWorkSheet2.Cells[NROW56, 1] = dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW56, NCOL2] = yerval;
                                        NROW56 = NROW56 + 1;

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                                      //  xlWorkSheet2.Cells[NROW76, 1] = dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW76, NCOL2] = yerval / 1000000;
                                        NROW76 = NROW76 + 1;
                                        yerval = 0;

                                    }
                                    //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                                    if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                    {
                                        hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                    }
                                    if ((((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                                    {
                                        hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                                     //   xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW66, NCOL2] = hyerval;
                                        NROW66 = NROW66 + 1;

                                        xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                                       // xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                        xlWorkSheet2.Cells[NROW86, NCOL2] = hyerval / 1000000;
                                        NROW86 = NROW86 + 1;
                                        hyerval = 0;
                                    }
                                    //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY

                                    val1 = val1 + 1;
                                    NROW36 = NROW36 + 1;
                                }
                            }
                        }
                        else
                        {

                            dt1 = dt36.AddDays(val);
                            dt2 = dt36.AddDays(val + 1);
                          //  xlWorkSheet2.Cells[NROW36, 1] = dt1;
                            xlWorkSheet2.Cells[NROW36, NCOL2] = xlWorkSheet.Cells[NROW1, 2];
                            xlWorkSheet2.Cells[1, NCOL2] = xlWorkSheet.Cells[NROWt, 1]; xlWorkSheet2.Cells[2, NCOL2] = xlWorkSheet.Cells[NROWt1, 2]; xlWorkSheet2.Cells[2, NCOL2 - 1] = xlWorkSheet.Cells[NROWt1, 1]; xlWorkSheet2.Cells[1, 1] = "";
                            //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            if (dt2.Day != 1)
                            {
                                mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                            }
                            if (dt2.Day == 1)
                            {
                                mthval = mthval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(4);
                              //  xlWorkSheet2.Cells[NROW46, 1] = dt1.Month.ToString() + "-" + dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW46, NCOL2] = mthval;
                                NROW46 = NROW46 + 1;
                                mthval = 0;

                            }//if (dt1.Day != 1)
                            //MMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMMM
                            if (!(((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                            {
                                yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                            }
                            if ((((dt1.Month == 12) && (dt2.Month == 1)) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null)))
                            {
                                yerval = yerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(5);
                              //  xlWorkSheet2.Cells[NROW56, 1] = dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW56, NCOL2] = yerval;
                                NROW56 = NROW56 + 1;

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(7);
                               // xlWorkSheet2.Cells[NROW76, 1] = dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW76, NCOL2] = yerval / 1000000;
                                NROW76 = NROW76 + 1;
                                yerval = 0;

                            }
                            //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY
                            if (!(((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                            {
                                hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                            }
                            if ((((dt1.Month == 9) && (dt2.Month == 10))) || (((Excel.Range)xlWorkSheet.Cells[NROW1 + 1, 1]).Value2 == null))
                            {
                                // hyerval = hyerval + Convert.ToDouble(((Excel.Range)xlWorkSheet.Cells[NROW1, 2]).Value2);

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(6);
                               // xlWorkSheet2.Cells[NROW66, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW66, NCOL2] = hyerval;
                                NROW66 = NROW66 + 1;

                                xlWorkSheet2 = (Excel.Worksheet)xlWorkBook2.Worksheets.get_Item(8);
                               // xlWorkSheet2.Cells[NROW86, 1] = (dt1.Year - 1).ToString() + "-" + dt1.Year.ToString();
                                xlWorkSheet2.Cells[NROW86, NCOL2] = hyerval / 1000000;
                                NROW86 = NROW86 + 1;
                                hyerval = 0;
                            }
                            //YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY

                        }


                        NROW6 = NROW6 + 1;
                    }//NCOL2>2
                }
                NROW1 = NROW1 + 1;
            }


            //   MessageBox.Show(((Excel.Range)xlWorkSheet.Cells[NROW1, 1]).Value2.ToString());

            //============================================================================================================================================================================================================================================

            int ind = pat.IndexOf(".");
            xlWorkBook2.SaveCopyAs(pat.Remove(ind) + "_Transform.xlsx");
            // xlWorkBook2.Save();
            xlWorkBook.Close(true, misValue, misValue);
            xlWorkBook2.Close(false, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

        }
    }
}
