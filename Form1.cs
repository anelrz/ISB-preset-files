using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excel_to_ISB_preset_files
{
    public partial class Form1 : Form
    {
        string[] files;
        Excel.Application xlApp;
        Excel.Workbook wb;
        int number;
        NumberFormatInfo nfi;
        StreamWriter sw;

        public Form1()
        {
            InitializeComponent();
            xlApp = new Microsoft.Office.Interop.Excel.Application();
            nfi = new NumberFormatInfo();
            nfi.NumberDecimalSeparator = ".";
            
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            try
            {
                folderBrowserDialog1.ShowDialog();
                files = Directory.GetFiles(folderBrowserDialog1.SelectedPath, "*.xlsx", SearchOption.AllDirectories);
                listBox1.Items.AddRange(files);
            }
            catch (DirectoryNotFoundException) { }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            using (sw = new StreamWriter("myxml.xml"))
            { 
                sw.WriteLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sw.WriteLine("<presets>");

                foreach (string file in files)
                {
                    try
                    {
                        wb = xlApp.Workbooks.Open(Filename: file, IgnoreReadOnlyRecommended: true);
                        Excel.Worksheet ws = wb.Worksheets[1];
                        //Console.WriteLine(ws.Name);

                        int lastRow = ws.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
                    
                        for (int i = 1; i <= lastRow; i++)
                        {
                            try
                            {
                                string cell = (String) ws.Cells[i, 3].Value2;

                                if (cell[0].CompareTo('M') == 0 && Int32.TryParse(cell[1].ToString(), out number))
                                {
                                    string value = ws.Cells[i, 4].Value2;
                                    //Console.WriteLine(cell + "  " + value);

                                    string[] words = value.Split(' ');
                                    double amp;
                                    Double.TryParse(words[0], NumberStyles.Number, nfi, out amp);
                                    string unit = words[1];
                                    Console.Write(".");

                                    switch (unit)
                                    {
                                        case "mV@80Hz":
                                            sw.WriteLine("  <Preset>");
                                            sw.WriteLine("    <name>" + cell + "</name>");
                                            sw.WriteLine("    <frequency>80</frequency>");
                                            sw.WriteLine("    <amplitude>" + amp + "</amplitude>");
                                            sw.WriteLine("    <offset>-8</offset>");
                                            sw.WriteLine("  </Preset>");
                                            break;

                                        case "mV@160Hz":
                                            sw.WriteLine("  <Preset>");
                                            sw.WriteLine("    <name>" + cell + "</name>");
                                            sw.WriteLine("    <frequency>160</frequency>");
                                            sw.WriteLine("    <amplitude>" + amp + "</amplitude>");
                                            sw.WriteLine("    <offset>-8</offset>");
                                            sw.WriteLine("  </Preset>");
                                            break;
                                            
                                        case "Hz":
                                            sw.WriteLine("  <Preset>");
                                            sw.WriteLine("    <name>" + cell + "</name>");
                                            sw.WriteLine("    <frequency>" + amp + "</frequency>");
                                            sw.WriteLine("    <amplitude>" + 1 + "</amplitude>");
                                            sw.WriteLine("    <offset>-8</offset>");
                                            sw.WriteLine("  </Preset>");
                                            break;

                                        case "V":
                                            sw.WriteLine("  <Preset>");
                                            sw.WriteLine("    <name>" + cell + "</name>");
                                            sw.WriteLine("    <frequency>0</frequency>");
                                            sw.WriteLine("    <amplitude>0</amplitude>");
                                            sw.WriteLine("    <offset>" + amp + "</offset>");
                                            sw.WriteLine("  </Preset>");
                                            break;

                                        default:
                                            Console.WriteLine("Unhandled unit: " + unit);
                                            break;

                                    }
                                }
                            }
                            catch (Exception)
                            {
                                //Console.WriteLine(exc.Message);
                            }
                        }
                        wb.Close();
                    }
                    
                    catch (System.Runtime.InteropServices.COMException ex)
                    {
                        Console.WriteLine(ex.Message.ToString());
                    }
                }
                sw.WriteLine("</presets>");
            }
        }
    }
}
