using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;

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
            //copiasTXT.SelectedIndex = 3;


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
            //REALIZO REMPLAZO DE VARIABLES EN VISTA PREVIA

            richTextBox1.SelectAll();

            string[] textArray = richTextBox1.SelectedText.Split(new char[] { '\n' });

            foreach (string strText in textArray)
            {
                if (!string.IsNullOrEmpty(strText))
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**contratoTXT**", contratoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaIniTXT**", fechaIniTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaFinTXT**", fechaFinTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**ciudadTXT**", ciudadTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**destinoTXT**", destinoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**inmuebleTXT**", inmuebleTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionTXT**", direccionTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**administracionTXT**", administracionTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**vigenciaTXT**", vigenciaTXT.Text);


                //GENERAMOS VALOR DE CANON EN MONEDA
                string canonTXTFormato = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(canonTXT.Text));
                canonTXTFormato = canonTXTFormato.Substring(0, canonTXTFormato.Length - 3);
                /////////
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**canonTXT**", canonTXTFormato);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**cuantiaTXT**", "( " + canonTXTFormato + " )" + " Valor letras");


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreArrendatarioTXT**", nombreArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idArrendatarioTXT**", idArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoArrendatarioTXT**", telefonoArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularArrendatarioTXT**", celularArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailArrendatarioTXT**", emailArrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionArrendatarioTXT**", direccionArrendatarioTXT.Text);


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombrePropietarioTXT**", nombrePropietarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idPropietarioTXT**", idPropietarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoPropietarioTXT**", telefonoPropietarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularPropietarioTXT**", celularPropietarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailPropietarioTXT**", emailPropietarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionPropietarioTXT**", direccionPropietarioTXT.Text);


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatarioTXT**", nombreCoarrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatarioTXT**", idCoarrendatarioTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatarioTXT**", idCoarrendatarioTXT.Text);


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario2TXT**", nombreCoarrendatario2TXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario2TXT**", idCoarrendatario2TXT.Text);

                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario3TXT**", nombreCoarrendatario3TXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario3TXT**", idCoarrendatario3TXT.Text);

                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario4TXT**", nombreCoarrendatario4TXT.Text);

            }


            //------------------------------------FIN DE GENERAR REMPLAZO VISTA PREVIA
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
          



            //ESCRIBO ARCHIVO EN WORD
            System.IO.File.WriteAllText(@"C:\RZ\txt.rtf", richTextBox1.Rtf);
            //------------------------------------

        }

        private void tabPage1_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void direccionArrendatario_TextChanged(object sender, EventArgs e)
        {

        }

        private void label20_Click(object sender, EventArgs e)
        {

        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void button7_Click(object sender, EventArgs e)
        {

           
        }

        private void canonTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void canonTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
           
            
            
        
        }

        private void button7_Click_1(object sender, EventArgs e)
        {
           
        }

        private void cONTRATODEARRENDAMIENTODEUNBIENINMUEBLESOMETIDOACOPROPIEDADYDESTINADOAVIVIENDAURBANAToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void cONTRATODEARRENDAMIENTODEUNBIENINMUEBLESOMETIDOACOPROPIEDADYDESTINADOAVIVIENDAURBANAPRUEBAToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //TRAIGO EL FORMATO CORRESPONDIENTE
            
            object readOnly = true;
            object visible = true;
            object save = false;
            object fileName = @"\\servidor1\sistemas\PROYECTOS\Exporter\Formatos\formato1.docx";
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

            //-----------------------FIN DE TRAER CONTRATO


            //DESHABILITO LO QUE NO SE NECESITA
            propietarioGroup.Enabled = false;
            //----------------------------
        }
    }
}
