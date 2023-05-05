using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using Humanizer;
using System.Diagnostics;

namespace Exporter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void canonAdminLetrasFn(string numero, string tipo, int vigencia) //FUNCION QUE FORMATEA VALORES DE CANON Y ADMIN, CALCULA CUANTIA Y DEVUELVE VALORES EN LETRAS
        {
            int numeroInteger;
            int cuantiaInteger;

            string numeroFormateado = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(numero));
            numeroFormateado = numeroFormateado.Substring(0, numeroFormateado.Length - 3);

            
            int.TryParse(numero, out numeroInteger);
            
            cuantiaInteger = numeroInteger * vigencia;            
            string cuantiaFormateado = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(cuantiaInteger.ToString()));
            cuantiaFormateado = cuantiaFormateado.Substring(0, cuantiaFormateado.Length - 3);


            switch (tipo)
            {
                case "canon":
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**canonTXT**", "( " + numeroFormateado + " ) " + (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS MCTE");
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**cuantiaTXT**", "( " + cuantiaFormateado + " ) " + (cuantiaInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS MCTE");
                    break;
                case "admin":
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**administracionTXT**", "( " + numeroFormateado + " ) " + (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS MCTE");
                    break;
                          }


           
        }

        public void terceroFn(string[] arr , string tipoTercero)
        {
            //FORMATEO TERCEROS ------------------------------------
            
            int terceroInteger;
            // Loop over strings.
            for (int i = 0; i < arr.Length; i++)
            {

                int.TryParse(arr[i], out terceroInteger);

                string terceroFormateado = terceroInteger.ToString("N", new CultureInfo("es-CL"));
                terceroFormateado = terceroFormateado.Substring(0, terceroFormateado.Length - 3); //<--- QUITAMOS DECIMALES
                Console.WriteLine("Tercero es: " + terceroFormateado);


                switch (i)
                {
                    case 0:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idArrendatarioTXT**", terceroFormateado);
                        break;
                    case 1:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idPropietarioTXT**", terceroFormateado);
                        break;
                    case 2:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatarioTXT**", terceroFormateado);
                        break;
                    case 3:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario2TXT**", terceroFormateado);
                        break;
                    case 4:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario3TXT**", terceroFormateado);
                        break;
                    case 5:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario4TXT**", terceroFormateado);
                        break;
                    
                }




            };

            //------------------------------------------

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
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionTXT**", direccionTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**vigenciaTXT**", vigenciaTXT.Text);


               

                //CONVERTIMOS DE STRING A INT EL VALOR DEL CANON Y EL ADMIN PARA USAR EL HUMANIZER DE NUMEROS A LETRAS
                /*int canonInteger;
                int adminInteger;
                int.TryParse(administracionTXT.Text, out adminInteger);
                int.TryParse(canonTXT.Text, out canonInteger);*/
                //-------------FINALIZA CONVERSION NUMERO A TEXTO


                //GENERAMOS VALOR DE CANON DEL TEXTBOX Y LO PASO A NUMERO MONEDA Y QUITAMOS DECIMALES
                /*
                string canonTXTFormato = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(canonTXT.Text));
                canonTXTFormato = canonTXTFormato.Substring(0, canonTXTFormato.Length - 3);
                */
                ///////////------------



                //CONVERTIMOS DE STRING A INT EL VALOR DEL ADMIN
                


                //-------------FINALIZA CONVERSION NUMERO A TEXTO


                //GENERAMOS VALOR DE ADMIN DEL TEXTBOX Y LO PASO A NUMERO MONEDA Y QUITAMOS DECIMALES

                /*
                string adminTXTFormato = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(administracionTXT.Text));
                adminTXTFormato = adminTXTFormato.Substring(0, adminTXTFormato.Length - 3);
                */

                ///////////------------



                //CALCULAMOS LA CUANTIA = CANON X VIGENCIA
                /*
                int vigenciaInteger;
                int.TryParse(vigenciaTXT.Text, out vigenciaInteger);
                int calculoCuantia = canonInteger * vigenciaInteger;
                string calculoCuantiaStr = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(calculoCuantia.ToString())); //<-- VOLVEMOS EN MILES
                calculoCuantiaStr = calculoCuantiaStr.Substring(0, calculoCuantiaStr.Length - 3); //<--- QUITAMOS DECIMALES
                */

                //-------------FINALIZA CALCULO CUANTIA


                //GENERAMOS VALOR DE CANON DE LOS TERCEROS TEXTBOX Y LO PASO A NUMERO  Y QUITAMOS DECIMALES

                /*
                 * int idArrendatarioInteger;
                int.TryParse(idArrendatarioTXT.Text, out idArrendatarioInteger);
                string usFormated = idArrendatarioInteger.ToString("N", new CultureInfo("es-CL"));
                calculoCuantiaStr = calculoCuantiaStr.Substring(0, calculoCuantiaStr.Length - 3);
                */

                ///////////------------

               // richTextBox1.Rtf = richTextBox1.Rtf.Replace("**cuantiaTXT**", "( " + calculoCuantiaStr + " ) " + (calculoCuantia.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS MCTE") ;
               // richTextBox1.Rtf = richTextBox1.Rtf.Replace("**canonTXT**", "( " + canonTXTFormato + " ) " + (canonInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS MCTE");
                //richTextBox1.Rtf = richTextBox1.Rtf.Replace("**administracionTXT**", "( " + adminTXTFormato + " ) " + (adminInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS MCTE");
                


                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreArrendatarioTXT**", nombreArrendatarioTXT.Text);
                //richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idArrendatarioTXT**", arrendatarioTXTFormato);
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
            var rutaArchivo = @"\\servidor1\Fotos\FOTOS_FIRMA_DE_CONTRATOS\CTO_" + contratoTXT.Text + @"\" + @"DOCUMENTOS\" + contratoTXT.Text + ".rtf";
            System.IO.File.WriteAllText( rutaArchivo, richTextBox1.Rtf);
            //------------------------------------

            //ABRO AUTOMATICAMENTE ARCHIVO GENERADO
            Process.Start(rutaArchivo);
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

            //COPIAS POR DEFECTO
            copiasTXT.SelectedIndex = 3;

            //VIGENCIA POR DEFECTO
            vigenciaTXT.SelectedIndex = 1;

            //CAMBIO DESTINO
            destinoTXT.Text = "Vivienda";

        }




        private void button8_Click(object sender, EventArgs e)
        {



            //LLAMAMOS FUNCION PARA FORMATEAR ID TERCEROS
            string[] arr = new string[6];
            arr[0] = idArrendatarioTXT.Text;
            arr[1] = idPropietarioTXT.Text;
            arr[2] = idCoarrendatarioTXT.Text;
            arr[3] = idCoarrendatario2TXT.Text;
            arr[3] = idCoarrendatario3TXT.Text;
            arr[3] = idCoarrendatario4TXT.Text;
            
            var tercero = idArrendatarioTXT.Text;
            var tipoTercero = "arrendatario";
            terceroFn(arr, tipoTercero);

            //------------------------------------------




           



        }

        private void button9_Click(object sender, EventArgs e)
        {
            canonAdminLetrasFn(canonTXT.Text, "canon", 12);
         

        }
    }
}
