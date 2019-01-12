using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Threading;
namespace AppWordExcel
{
    public partial class Checker : Form
    {
        private Excel.Range excelcells;
        private Excel.Range excelcells2;
        class Settings
        {
            public string SheetName { get; set; }
            public string FirstCell { get; set; }
            public string SecondCell { get; set; }
            public bool Bold { get; set; }
            public bool Italic { get; set; }
        };


        public void XmlParse()
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "XML Files (*.xml*)|*.xml*";
            if (opf.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл настроек!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            filename = opf.FileName;
            XmlDocument xDoc = new XmlDocument();
            try
            {
                xDoc.Load("Settings.xml");
            }
            catch(Exception e)
            {
                MessageBox.Show("Файл не найден или не соответствует формату", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            try
            {
                XmlElement xRoot = xDoc.DocumentElement;
                foreach (XmlNode xnode in xRoot)
                {
                    Settings SheetName = new Settings();

                    if (xnode.Attributes.Count > 0)
                    {
                        XmlNode attr = xnode.Attributes.GetNamedItem("SheetName");
                        WorkSheet.Text = attr.Value;
                    }
                    foreach (XmlNode childnode in xnode.ChildNodes)
                    {
                        if (childnode.Name == "firstCell")
                        {
                            FirstCellBox.Text = childnode.InnerText;
                        }
                        if (childnode.Name == "secondCell")
                        {
                            LastCellBox.Text = childnode.InnerText;
                        }
                        if (childnode.Name == "Bold")
                        {
                            if (childnode.InnerText == "1")
                            {
                                Bold.Checked = true;
                            }
                            else
                            {
                                Bold.Checked = false;
                            }
                        }
                        if (childnode.Name == "Italic")
                        {
                            if (childnode.InnerText == "1")
                            {
                                checkBoxItalic.Checked = true;

                            }
                            else
                            {
                                checkBoxItalic.Checked = false;
                            }
                        }
                    }
                }
                MainTextBox.Text += "Налаштунки вдало імпортовано!" + "\r\n";
            }
            catch ( Exception e)
            {
                MessageBox.Show("Помилка імпорту!", "Увага!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MainTextBox.Text += "Помилка імпорту налаштунків!" + "\r\n";
            }
        }


        public Checker()
        {
            InitializeComponent();
        }
        string PathToSave = "";
        public void ChooseFolder()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                PathToSave = @folderBrowserDialog1.SelectedPath;
            }
        }
        string PathToSaveConfig = "";
        public void ChooseFolderForConfig()
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                PathToSaveConfig = @folderBrowserDialog1.SelectedPath;
            }
        }
        static string NumberToLetters(int number)
        {
            string result;
            if (number > 0)
            {
                int alphabets = (number - 1) / 26;
                int remainder = (number - 1) % 26;
                result = ((char)('A' + remainder)).ToString();
                if (alphabets > 0)
                    result = NumberToLetters(alphabets) + result;
            }
            else
                result = null;
            return result;
        }
        string filename2;
        string filename;
        void FirstFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            if (opf.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали первый Excel файл!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            filename = opf.FileName;
            MainTextBox.Text += "Path to the first book: " + filename + "\r\n";
        }
        void SecondFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf2 = new OpenFileDialog();
            opf2.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            if (opf2.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали второй Excel файл!", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            filename2 = opf2.FileName;
            MainTextBox.Text += "Path to the second book: " + filename2 + "\r\n";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            Excel.Worksheet xlSht;

            Excel.Application xlApp2 = new Excel.Application();
            Excel.Worksheet xlSht2;
            Excel.Workbook xlWorkBook;
            Excel.Workbook xlWorkBook2;

            try
            {
                xlWorkBook = xlApp.Workbooks.Open(filename);
                xlWorkBook2 = xlApp2.Workbooks.Open(filename2);

                xlSht = xlApp.Worksheets[WorkSheet.Text];
                xlSht2 = xlApp2.Worksheets[WorkSheet.Text];

                int arrData3 = xlSht.Range[FirstCellBox.Text + ":" + LastCellBox.Text].Columns.Count;
                int arrData4 = xlSht.Range[FirstCellBox.Text + ":" + LastCellBox.Text].Rows.Count;

                excelcells2 = xlSht2.get_Range("A1", Type.Missing);
                excelcells = xlSht.get_Range("A1", Type.Missing);
                int iRow, iCol;
                int max = arrData4 * arrData3;
                int i = 0;

                string Result = "";
                string bold = "";
                StreamWriter writer = new StreamWriter(PathToSave + "report.txt");
                MainTextBox.Text += "_______________________________________________________________________________\r\n";
                MainTextBox.Text += "Start...\r\n";
                for (iRow = 0; iRow < arrData4; iRow++)
                {
                    for (iCol = 0; iCol < arrData3; iCol++)
                    {
                        excelcells = excelcells.get_Offset(iRow, iCol);
                        excelcells2 = excelcells2.get_Offset(iRow, iCol);
                        if (Bold.Checked)
                        {
                            bool isCell1Bold = excelcells.Font.Bold;
                            bool isCell2Bold = excelcells2.Font.Bold;
                            if (isCell1Bold != isCell2Bold)
                            {
                                writer.WriteLine("Несовпадение формата жирности в клеточке " + NumberToLetters(iCol + 1) + (iRow + 1) + "\n");
                                MainTextBox.Text += "Несовпадение формата жирности в клеточке " + NumberToLetters(iCol + 1) + (iRow + 1) + "\r\n";
                                
                            }

                        }
                        if (checkBoxItalic.Checked)
                        {
                            bool isCell1Italic = excelcells.Font.Italic;
                            bool isCell2Italic = excelcells2.Font.Italic;
                            if (isCell1Italic != isCell2Italic)
                            {
                                writer.WriteLine("Несовпадение формата курсива в клеточке " + NumberToLetters(iCol + 1) + (iRow + 1) + "\n");
                             
                                    MainTextBox.Text += "Несовпадение формата курсива в клеточке " + NumberToLetters(iCol + 1) + (iRow + 1) + "\r\n";
                                
                            }

                        }
                        if (!(excelcells.FormulaLocal == excelcells2.FormulaLocal))
                        {
                            writer.WriteLine("Несовпадение значения в клеточке " + NumberToLetters(iCol + 1) + (iRow + 1) + "\n");
                            
                                MainTextBox.Text += "Несовпадение значения в клеточке " + NumberToLetters(iCol + 1) + (iRow + 1) + "\r\n";
                            

                        }

                        excelcells = excelcells.get_Offset(-iRow, -iCol);
                        excelcells2 = excelcells2.get_Offset(-iRow, -iCol);
                        i++;
                        toolStripProgressBar1.Value = (int)((100 * i) / max);
                        toolStripStatusLabel1.Text = "Checking cell " + NumberToLetters(iCol + 1) + (iRow + 1);
                    }
                }
                toolStripStatusLabel1.Text = "Готово!";
                if (Bold.Checked)
                {
                    Result = Result + bold;
                }
                writer.Close();
                xlApp.Quit();
                xlApp2.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                xlApp = null;
                xlWorkBook = null;
                xlSht = null;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp2);

                xlApp2 = null;
                xlWorkBook2 = null;
                xlSht2 = null;

                System.GC.Collect();
                MainTextBox.Text += "End.\r\n";
                MainTextBox.Text += "_______________________________________________________________________________\r\n";

            }
            catch (Exception Exc)
            {
                MessageBox.Show(Exc.Message);
            }

        }
        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void button4_Click(object sender, EventArgs e)
        {
            ChooseFolder();
        }

       /* private void button5_Click(object sender, EventArgs e)
        {
            OpenFileDialog opf = new OpenFileDialog();
            opf.Filter = "Configurate file (.cfg)|*.cfg|Text file (.txt)|*.txt*";
            if (opf.ShowDialog() != DialogResult.OK)
            {
                MessageBox.Show("Вы не выбрали файл с настройками", "Внимание", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            filename = opf.FileName;
            string[] Mass = File.ReadAllLines(filename, System.Text.Encoding.Default);
            string s = "";
            for(int i = 0; i < Mass.Length; i++)
            {
                s += Mass[i] + "\n";
            }
            string firstcell = "[firstCell]=";
            string seccell = "[secondCell]=";
            string isBold = "[checkForBold]=";
            string isItalic = "[checkForItalic]=";
            try
            {
                int index = s.IndexOf(firstcell);
                int index2 = s.IndexOf(seccell);
                int firstCelllength = index2 - 1 - firstcell.Length;
                int firstcellindex = index2 - firstCelllength - 1;
                string firstCell = s.Substring(firstcellindex, firstCelllength); //First Cell Name
                textBox1.Text = firstCell;
                int index3 = s.IndexOf(seccell) + seccell.Length;
                int index4 = s.IndexOf(isBold) - 1;
                string secondCell = s.Substring(index3, index4 - index3); //Second Cell Name
                int index5 = index4 + isBold.Length + 1;
                textBox2.Text = secondCell;
                bool isBoldCheck = s.Substring(index5, 1) == "1"; //Is Bold must check
                int index6 = index5 + isItalic.Length + 2;
                checkBox1.Checked = isBoldCheck;
                bool isItalicCheck = s.Substring(index6, 1) == "1"; //Is Italic must check
                checkBoxItalic.Checked = isItalicCheck;


                int index7 = s.IndexOf("[SheetName]=") + "[SheetName]=".Length;
                WorkSheet.Text = (s.Substring(index7)).Substring(0, (s.Substring(index7)).Length - 1);
            }catch(Exception Exc)
            {
                MessageBox.Show("Invalid config", "Error");
            }
        }*/
        /*private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                ChooseFolderForConfig();
                string Folder = PathToSaveConfig + "\\config.cfg";
                MessageBox.Show("Config saved to:\n" + PathToSaveConfig);
                string text = "[firstCell]=" + textBox1.Text;
                string sec = "[secondCell]=" + textBox2.Text;
                string bold = "";
                string italic = "";
                if (checkBox1.Checked)
                {
                    bold = "[checkForBold]=1";
                }
                else
                {
                    bold = "[checkForBold]=0";
                }
                if (checkBoxItalic.Checked)
                {
                    italic = "[checkForItalic]=1";
                }
                else
                {
                    italic = "[checkForItalic]=0";
                }
                string sheet = "[SheetName]=" + WorkSheet.Text;
                try
                {
                    using (StreamWriter sw = new StreamWriter(Folder, false, System.Text.Encoding.Default))
                    {
                        sw.WriteLine(text);
                        sw.WriteLine(sec);
                        sw.WriteLine(bold);
                        sw.WriteLine(italic);
                        sw.WriteLine(sheet);
                    }
                }
                catch (Exception Exc)
                {
                    MessageBox.Show(Exc.Message);
                }
            }
            catch (Exception Exc)
            {
                MessageBox.Show("Invalid config", "Error");
            }
        }
        */
        private void button6_Click(object sender, EventArgs e)
        {
            XmlParse();
        }
        private void XmlCreate_Click(object sender, EventArgs e)
        {
            int isBOLD = 1;
            int isITALIC = 1;
            if (Bold.Checked)
            {
                isBOLD = 1;
            }
            else
            {
                isBOLD = 0;
            }
            if (checkBoxItalic.Checked)
            {
                isITALIC = 1;
            }
            else
            {
                isITALIC = 0;
            }
            XDocument xdoc = new XDocument();
            // создаем первый элемент
            XElement settings = new XElement("settings");
            // создаем атрибут
            XAttribute SheetName = new XAttribute("SheetName", WorkSheet.Text);
            XElement firstCell = new XElement("firstCell", FirstCellBox.Text);
            XElement secondCell = new XElement("secondCell", LastCellBox.Text);
            XElement BOLD = new XElement("Bold", isBOLD);
            XElement ITALIC = new XElement("Italic", isITALIC);
            // добавляем атрибут и элементы в первый элемент
            settings.Add(SheetName);
            settings.Add(firstCell);
            settings.Add(secondCell);
            settings.Add(BOLD);
            settings.Add(ITALIC);

            
            // создаем корневой элемент
            XElement setting = new XElement("setting");
            // добавляем в корневой элемент
            setting.Add(settings);
            // добавляем корневой элемент в документ
            xdoc.Add(setting);
            //сохраняем документ
            xdoc.Save("Settings.xml");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            MainTextBox.Text = "";
        }
    }
}