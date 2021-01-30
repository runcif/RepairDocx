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
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;

namespace RepairDocx
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Microsoft Word (.docx)|*.docx";
            //openFileDialog1.FilterIndex = 1;
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = openFileDialog1.FileName;
                LeggiXML(openFileDialog1.FileName);
            }
        }

        public void LeggiXML(string filepath)
        {
            try
            {
                // Open a WordprocessingDocument based on a filepath.
                using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filepath, false))
                {
                    var paragraphs = wordDocument.MainDocumentPart.RootElement.Descendants<Paragraph>();
                    foreach (var paragraph in paragraphs)
                    {
                        
                            richTextBox1.AppendText(paragraph.InnerText+"\r\n");
                            //outputFile.WriteLine(paragraph.InnerText);
                            //outputFile.Flush();
                            //outputFile.Close();

                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Errore:"+ex);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "txt files (*.txt)|*.txt"; 
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                richTextBox1.SaveFile(saveFileDialog1.FileName, RichTextBoxStreamType.PlainText);
                MessageBox.Show("File salvato con successo!");
            }
        }
    }
}
