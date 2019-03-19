using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Xml;
using System.IO;

namespace ExcelToEpricaXML
{
   

    public partial class Form1 : Form
    {
        private static Excel.Workbook MyBook = null;
        private static Excel.Application MyApp = null;
        private static Excel.Worksheet ErSheet = null;


        /// <summary>
        /// Функция необходима для того, чтобы убать первую строчку  описания XML. 
        /// ePrica не обрабатывает файл отличный от <XML> </XML>
        /// </summary>
        /// <param name="xmlPath"></param>
        public void removeXMLdeclaration(string xmlPath)
        {
            try
            {
                //Grab file
                StreamReader sr = new StreamReader(xmlPath);

                //Read first line and do nothing (i.e. eliminate XML declaration)
                sr.ReadLine();
                string body = null;
                string line = sr.ReadLine();
                while (line != null) // read file into body string
                {
                    body += line + "\n";
                    line = sr.ReadLine();
                }
                sr.Close(); //close file

                //Write all of the "body" to the same text file
                System.IO.File.WriteAllText(xmlPath, body);
            }
            catch (Exception e3)
            {
                MessageBox.Show(e3.Message);
            }

        }
        public Form1()
        {
            InitializeComponent();
        }

        private void panel1_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.All;
        }

        private void panel1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop,false);
            foreach (string file in files)
            {
                textBox1.Text = file;
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string fname = "";
            fname = textBox1.Text;
            if (fname.Length == 0)
                 return;

            if (!File.Exists(fname))
            {
                MessageBox.Show("Не могу найти файл ");
                return;
            }

            try
            {
                MyApp = new Excel.Application();
                MyApp.Visible = true;


                MyBook = MyApp.Workbooks.Open(fname);

                ErSheet = (Excel.Worksheet)MyBook.Sheets[1];
            }
            catch (Exception e3)
            {
                MessageBox.Show(e3.Message);
                return;
            }
            XmlWriterSettings settingsxml = new XmlWriterSettings();
            String defpath = "defektura.xml";


            settingsxml.Indent = true;
            #region Шапка XML  файла дефектуры пример
            /*
             <HEADER>
            <ROW>
            <ID_DEFECTURA>659fa090-24d4-473f-adc3-ff73c9557e3e</ID_DEFECTURA>
            <DOC_DATE>2019-03-10T17:59:18.597</DOC_DATE>
            <ID_CONTRACTOR_GLOBAL>1fb6f806-63bc-4ef0-afca-736c5dc1e7fa</ID_CONTRACTOR_GLOBAL>
            <CONTRACTOR_NAME>ООО "Волга" ЦО</CONTRACTOR_NAME>
            <CONTRACTOR_CODE>
            </CONTRACTOR_CODE>
            </ROW>
            </HEADER>

            */
            #endregion 
            using (XmlWriter writer = XmlWriter.Create("defektura.xml", settingsxml))
            {
                writer.WriteStartElement("XML");
                writer.WriteStartElement("HEADER");
                writer.WriteStartElement("ROW");
                writer.WriteElementString("ID_DEFECTURA",System.Guid.NewGuid().ToString());
                writer.WriteElementString("DOC_DATE", DateTime.Now.ToString("yyyy-MM-ddThh:mm:ss"));
                writer.WriteElementString("ID_CONTRACTOR_GLOBAL", "1fb6f806-63bc-4ef0-afca-736c5dc1e7fa");
                writer.WriteElementString("CONTRACTOR_NAME","ООО Здоровье");
                writer.WriteElementString("CONTRACTOR_CODE","");
                writer.WriteEndElement();
                writer.WriteEndElement();

                #region Строка дефектуры пример
                /*
                 <ROWS>
                 <ROW>
                   <ID_GOODS_GLOBAL>b0a1680c-501f-4233-97f0-fd18ffb695f9</ID_GOODS_GLOBAL>
                <GOODS_NAME>La Roche-Posay Гиалу В5 крем для глаз 15мл</GOODS_NAME>
                <PRODUCER_NAME>Косметик актив продюксьон</PRODUCER_NAME>
                <QTY_REMAIN>0.0000</QTY_REMAIN>
                <QTY_MIN>0.0000</QTY_MIN>
                  <LAST_PRICE_SAL>1680.0000</LAST_PRICE_SAL>
                <LAST_PRICE_SUP>1345.2000</LAST_PRICE_SUP>
                <LAST_SUPPLIER_NAME>Бьютикс-Восток ООО</LAST_SUPPLIER_NAME>
                </ROW> 
                */
                #endregion
                writer.WriteStartElement("ROWS");
                for (int index = 14; index <= 2500; index++)
                {
                    progressBar1.Value = index;
                    float remain;
                    float min;
                    Excel.Range GoodName = ErSheet.get_Range("A" + index.ToString(), "A" + index.ToString());
                    if (GoodName.Text == ""||GoodName.Text==" ")
                        continue;

                    Excel.Range rngremain = ErSheet.get_Range("C" + index.ToString(), "C" + index.ToString());
                    if (rngremain.Text == "" || rngremain.Text ==" ")
                        remain = 0;
                    else
                        remain = float.Parse(rngremain.Text);


                    Excel.Range rngmin = ErSheet.get_Range("B" + index.ToString(), "B" + index.ToString());
                    if ( rngmin.Text == "" || rngmin.Text == " ")
                        min = 0;
                    else
                        min = float.Parse(rngmin.Text);


                    writer.WriteStartElement("ROW");
                    writer.WriteElementString("ID_GOODS_GLOBAL", System.Guid.NewGuid().ToString());
                    writer.WriteElementString("GOODS_NAME", GoodName.Text);
                    writer.WriteElementString("PRODUCER_NAME", "Производитель");

                    
                    writer.WriteElementString("QTY_REMAIN", remain.ToString("0.0000").Replace(",","."));


                    writer.WriteElementString("QTY_MIN", min.ToString("0.0000").Replace(",", "."));

                    writer.WriteElementString("LAST_PRICE_SAL", "0.000");
                    writer.WriteElementString("LAST_PRICE_SUP", "0.000");
                    writer.WriteElementString("LAST_SUPPLIER_NAME", "Поставщик");
                    writer.WriteEndElement();

                   


                   
                }
                writer.WriteEndElement(); //ROWS
                writer.WriteEndElement(); //XML
                writer.Flush();

               

            }
            removeXMLdeclaration(defpath);
            MessageBox.Show("Дефектура готова");
        }
    }
}
