using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace Exporter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Word.Application word = new Word.Application();
            word.Visible = true;
            word.WindowState = Word.WdWindowState.wdWindowStateNormal;
            Word.Document doc = word.Documents.Add();
            Word.Paragraph paragraph;
            paragraph = doc.Paragraphs.Add();

            
            paragraph.Range.Text = richTextBox1.Text;
            doc.SaveAs(@"c:\RZ\mydoc.rtf");
            doc.Close();
                word.Quit();
        }

        private void button2_Click(object sender, EventArgs e)
        {

            using (OpenFileDialog ofd = new OpenFileDialog() { ValidateNames = true, Multiselect = false, Filter = "Word Doucment|*.docx|Word 97 - 2003 Document|*.doc" })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    object readOnly = true;
                    object visible = true;
                    object save = false;
                    object fileName = ofd.FileName;
                    object missing = Type.Missing;
                    object newTemplate = false;
                    object docType = 0;
                    Microsoft.Office.Interop.Word._Document oDoc = null;
                    Microsoft.Office.Interop.Word._Application oWord = new Microsoft.Office.Interop.Word.Application() { Visible = false };
                    oDoc = oWord.Documents.Open(
                            ref fileName, ref missing, ref readOnly, ref missing,
                            ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref visible,
                            ref missing, ref missing, ref missing, ref missing);
                    oDoc.ActiveWindow.Selection.WholeStory();
                    oDoc.ActiveWindow.Selection.Copy();
                    IDataObject data = Clipboard.GetDataObject();
                    richTextBox1.Rtf = data.GetData(DataFormats.Rtf).ToString();
                    oWord.Quit(ref missing, ref missing, ref missing);
                }
            }

        }
        public static void QuickReplace(RichTextBox rtb, String word, String word2)
        {
            rtb.Text = rtb.Text.Replace(word, word2);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            richTextBox1.Text = richTextBox1.Text.Replace(direccionTXT.Text, ciudadTXT.Text);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            richTextBox1.SelectAll();

            string[] textArray = richTextBox1.SelectedText.Split(new char[] { '\n' });

            foreach (string strText in textArray)
            {
                if (!string.IsNullOrEmpty(strText))
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**contratoTXT**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaIniTXT**", fechaIniTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaFinTXT**", fechaFinTXT.Text);


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreArrendatarioTXT**", nombreArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idArrendatarioTXT**", idArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoArrendatarioTXT**", telefonoArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularArrendatarioTXT**", celularArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailArrendatarioTXT**", emailArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionArrendatario**", direccionArrendatario.Text);



                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatarioTXT**", nombreCoarrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatarioTXT**", idCoarrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatarioTXT**", idCoarrendatarioTXT.Text);


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario2TXT**", nombreCoarrendatario2TXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario2TXT**", idCoarrendatario2TXT.Text);

                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario3TXT**", nombreCoarrendatario3TXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario3TXT**", idCoarrendatario3TXT.Text);

                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario4TXT**", nombreCoarrendatario4TXT.Text);
                
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**NOM_PROP**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**INMUEBLE**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**DIR_INMUEBLE**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**DE_CIUDAD**. ", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CANON_NUMEROS**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**CONTRATO**", contratoTXT.Text);

            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (printPreviewDialog1.ShowDialog() == DialogResult.OK)
                printDocument1.Print();
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(richTextBox1.Text, richTextBox1.Font, Brushes.Black, 100, 100);

        }

        private void button6_Click(object sender, EventArgs e)
        {
            System.IO.File.WriteAllText(@"C:\RZ\txt.rtf", richTextBox1.Rtf);
        }
    }
}
