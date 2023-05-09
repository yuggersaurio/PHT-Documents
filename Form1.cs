using System;
using System.Drawing;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Globalization;
using Humanizer;
using AltoHttp;
using Clases.ApiRest;
using Newtonsoft.Json;
using System.Diagnostics;


namespace Exporter
{
    public partial class Form1 : Form
    {

        DBApi dBApi = new DBApi();
        public Form1()
        {
            InitializeComponent();
        }

        private void canonAdminLetrasFn(string numero, string tipo, int vigencia) //FUNCION QUE FORMATEA VALORES DE CANON Y ADMIN, CALCULA CUANTIA Y DEVUELVE VALORES EN LETRAS E IMPRIME RESULTADO EN FORMATO
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
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**canonTXT**", (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS M/CTE " + "( " + numeroFormateado + " ) " );
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**cuantiaTXT**", (cuantiaInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS M/CTE " + "( " + cuantiaFormateado + " ) " );
                    break;
                case "admin":
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**administracionTXT**",  (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + " PESOS M/CTE " + "( " + numeroFormateado + " ) " );
                    break;
                    }


           
        }

        public void terceroFn(string[] arr , string tipoTercero) //FUNCION QUE FORMATEA ID DE TERCEROS E IMPRIME RESULTADOS EN FORMATO
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
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreArrendatarioTXT**", nombreArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoArrendatarioTXT**", telefonoArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularArrendatarioTXT**", celularArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailArrendatarioTXT**", emailArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionArrendatarioTXT**", direccionArrendatarioTXT.Text);

                        break;
                    case 1:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idPropietarioTXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombrePropietarioTXT**", nombrePropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoPropietarioTXT**", telefonoPropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularPropietarioTXT**", celularPropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailPropietarioTXT**", emailPropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionPropietarioTXT**", direccionPropietarioTXT.Text);
                        break;
                    case 2:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario1TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario1TXT**", nombreCoarrendatario1TXT.Text);
                        break;
                    case 3:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario2TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario2TXT**", nombreCoarrendatario2TXT.Text);
                        break;
                    case 4:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario3TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario3TXT**", nombreCoarrendatario3TXT.Text);
                        break;
                    case 5:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario4TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario4TXT**", nombreCoarrendatario4TXT.Text);
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
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**ciudadTXT**", ciudadTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**destinoTXT**", destinoTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionTXT**", direccionTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**vigenciaTXT**", vigenciaTXT.Text);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**copiasTXT**", copiasTXT.Text);


                //LLAMAMOS FUNCION PARA FORMATEAR CANON Y ADMIN Y REGRESAR LETRAS
                canonAdminLetrasFn(canonTXT.Text, "canon", 12);
                canonAdminLetrasFn(administracionTXT.Text, "admin", 12);
                //------------------------------------------

                //LLAMAMOS FUNCION PARA FORMATEAR ID TERCEROS
                string[] arr = new string[6];
                arr[0] = idArrendatarioTXT.Text;
                arr[1] = idPropietarioTXT.Text;
                arr[2] = idCoarrendatario1TXT.Text;
                arr[3] = idCoarrendatario2TXT.Text;
                arr[4] = idCoarrendatario3TXT.Text;
                arr[5] = idCoarrendatario4TXT.Text;

                var tercero = idArrendatarioTXT.Text;
                var tipoTercero = "arrendatario";
                terceroFn(arr, tipoTercero);

                //------------------------------------------


                //FECHAS DE INICIO Y FIN EN LETRAS
                string fechaInicio = fechaIniTXT.Text;
                string fechaFin = fechaFinTXT.Text;
                fechaFinTXT.Value = fechaIniTXT.Value;
                fechaFinTXT.Value = DateTime.Now.AddYears(1);
                fechaFin = fechaFinTXT.Text;
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaIniTXT**", fechaInicio.Humanize(LetterCasing.AllCaps));
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaFinTXT**", fechaFin.Humanize(LetterCasing.AllCaps));

                //------------------------------------------






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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }



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





           



        }

        private void button9_Click(object sender, EventArgs e)
        {
           
         

        }

        private void button10_Click(object sender, EventArgs e)
        {
            
        }

        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {

            //dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=2");
            /* dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=insertar"

                 +"&contratoTXTdb=" + contratoTXT.Text
                 +"&direccionTXTdb=" + direccionTXT.Text
                 +"&ciudadTXTdb=" + ciudadTXT.Text
                 +"&fechaIniTXTdb="+fechaIniTXT.Text
                 +"&fechaFinTXTdb=" + fechaFinTXT.Text
                 +"&canonTXTdb=" + canonTXT.Text
                 +"&administracionTXTdb=" + administracionTXT.Text
                 +"&destinoTXTdb=" + destinoTXT.Text
                 +"&nombrePropietarioTXTdb=" + nombrePropietarioTXT.Text
                 +"&idPropietarioTXTdb=" + idPropietarioTXT.Text
                 +"&telefonoPropietarioTXTdb=" + telefonoPropietarioTXT.Text
                 +"&celularPropietarioTXTdb=" + celularPropietarioTXT.Text
                 +"&emailPropietarioTXTdb=" + emailPropietarioTXT.Text
                 +"&direccionPropietarioTXTdb=" + direccionPropietarioTXT.Text
                 +"&nombreArrendatarioTXTdb=" + nombreArrendatarioTXT.Text
                 +"&idArrendatarioTXTdb=" + idArrendatarioTXT.Text
                 +"&telefonoArrendatarioTXTdb=" + telefonoArrendatarioTXT.Text
                 +"&celularArrendatarioTXTdb=" + celularArrendatarioTXT.Text
                 +"&emailArrendatarioTXTdb=" + emailArrendatarioTXT.Text
                 +"&direccionArrendatarioTXTdb=" + direccionArrendatarioTXT.Text

                 +"&nombreCoarrendatario1TXTdb=" + nombreCoarrendatario1TXT.Text
                 +"&idCoarrendatario1TXTdb=" + idCoarrendatario1TXT.Text
                 +"&telefonoCoarrendatario1TXTdb=" + telefonoCoarrendatario1TXT.Text
                 +"&celularCoarrendatario1TXTdb=" + celularCoarrendatario1TXT.Text
                 +"&direccionCoarrendatario1TXTdb=" + direccionCoarrendatario1TXT.Text
                 +"&emailCoarrendatario1TXTdb=" + emailCoarrendatario1TXT.Text

                 +"&nombreCoarrendatario2TXTdb=" + nombreCoarrendatario2TXT.Text
                 +"&idCoarrendatario2TXTdb=" + idCoarrendatario2TXT.Text
                 +"&telefonoCoarrendatario2TXTdb=" + telefonoCoarrendatario2TXT.Text
                 +"&celularCoarrendatario2TXTdb=" + celularCoarrendatario2TXT.Text
                 +"&direccionCoarrendatario2TXTdb=" + direccionCoarrendatario2TXT.Text
                 +"&emailCoarrendatario2TXTdb=" + emailCoarrendatario2TXT.Text

                 +"&nombreCoarrendatario3TXTdb=" + nombreCoarrendatario3TXT.Text
                 +"&idCoarrendatario3TXTdb=" + idCoarrendatario3TXT.Text
                 +"&telefonoCoarrendatario3TXTdb=" + telefonoCoarrendatario3TXT.Text
                 +"&celularCoarrendatario3TXTdb=" + celularCoarrendatario3TXT.Text
                 +"&direccionCoarrendatario3TXTdb=" + direccionCoarrendatario3TXT.Text
                 +"&emailCoarrendatario3TXTdb=" + emailCoarrendatario3TXT.Text

                 +"&nombreCoarrendatario4TXTdb=" + nombreCoarrendatario4TXT.Text
                 +"&idCoarrendatario4TXTdb=" + idCoarrendatario4TXT.Text
                 +"&telefonoCoarrendatario4TXTdb=" + telefonoCoarrendatario4TXT.Text
                 +"&celularCoarrendatario4TXTdb=" + celularCoarrendatario4TXT.Text
                 +"&direccionCoarrendatario4TXTdb=" + direccionCoarrendatario4TXT.Text
                 +"&emailCoarrendatario4TXTdb=" + emailCoarrendatario4TXT.Text


                 );*/

            Console.WriteLine("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=insertar"

               + "&contratoTXTdb=" + contratoTXT.Text
               + "&direccionTXTdb=" + direccionTXT.Text
               + "&ciudadTXTdb=" + ciudadTXT.Text
               + "&fechaIniTXTdb=" + fechaIniTXT.Text
               + "&fechaFinTXTdb=" + fechaFinTXT.Text
               + "&canonTXTdb=" + canonTXT.Text
               + "&administracionTXTdb=" + administracionTXT.Text
               + "&destinoTXTdb=" + destinoTXT.Text
               + "&nombrePropietarioTXTdb=" + nombrePropietarioTXT.Text
               + "&idPropietarioTXTdb=" + idPropietarioTXT.Text
               + "&telefonoPropietarioTXTdb=" + telefonoPropietarioTXT.Text
               + "&celularPropietarioTXTdb=" + celularPropietarioTXT.Text
               + "&emailPropietarioTXTdb=" + emailPropietarioTXT.Text
               + "&direccionPropietarioTXTdb=" + direccionPropietarioTXT.Text
               + "&nombreArrendatarioTXTdb=" + nombreArrendatarioTXT.Text
               + "&idArrendatarioTXTdb=" + idArrendatarioTXT.Text
               + "&telefonoArrendatarioTXTdb=" + telefonoArrendatarioTXT.Text
               + "&celularArrendatarioTXTdb=" + celularArrendatarioTXT.Text
               + "&emailArrendatarioTXTdb=" + emailArrendatarioTXT.Text
               + "&direccionArrendatarioTXTdb=" + direccionArrendatarioTXT.Text

               + "&nombreCoarrendatario1TXTdb=" + nombreCoarrendatario1TXT.Text
               + "&idCoarrendatario1TXTdb=" + idCoarrendatario1TXT.Text
               + "&telefonoCoarrendatario1TXTdb=" + telefonoCoarrendatario1TXT.Text
               + "&celularCoarrendatario1TXTdb=" + celularCoarrendatario1TXT.Text
               + "&direccionCoarrendatario1TXTdb=" + direccionCoarrendatario1TXT.Text
               + "&emailCoarrendatario1TXTdb=" + emailCoarrendatario1TXT.Text

               + "&nombreCoarrendatario2TXTdb=" + nombreCoarrendatario2TXT.Text
               + "&idCoarrendatario2TXTdb=" + idCoarrendatario2TXT.Text
               + "&telefonoCoarrendatario2TXTdb=" + telefonoCoarrendatario2TXT.Text
               + "&celularCoarrendatario2TXTdb=" + celularCoarrendatario2TXT.Text
               + "&direccionCoarrendatario2TXTdb=" + direccionCoarrendatario2TXT.Text
               + "&emailCoarrendatario2TXTdb=" + emailCoarrendatario2TXT.Text

               + "&nombreCoarrendatario3TXTdb=" + nombreCoarrendatario3TXT.Text
               + "&idCoarrendatario3TXTdb=" + idCoarrendatario3TXT.Text
               + "&telefonoCoarrendatario3TXTdb=" + telefonoCoarrendatario3TXT.Text
               + "&celularCoarrendatario3TXTdb=" + celularCoarrendatario3TXT.Text
               + "&direccionCoarrendatario3TXTdb=" + direccionCoarrendatario3TXT.Text
               + "&emailCoarrendatario3TXTdb=" + emailCoarrendatario3TXT.Text

               + "&nombreCoarrendatario4TXTdb=" + nombreCoarrendatario4TXT.Text
               + "&idCoarrendatario4TXTdb=" + idCoarrendatario4TXT.Text
               + "&telefonoCoarrendatario4TXTdb=" + telefonoCoarrendatario4TXT.Text
               + "&celularCoarrendatario4TXTdb=" + celularCoarrendatario4TXT.Text
               + "&direccionCoarrendatario4TXTdb=" + direccionCoarrendatario4TXT.Text
               + "&emailCoarrendatario4TXTdb=" + emailCoarrendatario4TXT.Text
               );


            //label46.Text = "https://portalhouses.com/administrador/apiPlantas/upload_inv/asesores_documentos/" + respuesta[0].contrato.ToString() + ".pdf";
            /*contratoTXT.Text = respuesta[0].contratoTXTdb.ToString();
            direccionTXT.Text = respuesta[0].direccionTXTdb.ToString();
            ciudadTXT.Text = respuesta[0].ciudadTXTdb.ToString();
            fechaIniTXT.Text = respuesta[0].fechaIniTXTdb.ToString();
            fechaFinTXT.Text = respuesta[0].fechaFinTXTdb.ToString();
            canonTXT.Text = respuesta[0].canonTXTdb.ToString();
            administracionTXT.Text = respuesta[0].administracionTXTdb.ToString();
            destinoTXT.Text = respuesta[0].destinoTXTdb.ToString();
            
            nombrePropietarioTXT.Text = respuesta[0].nombrePropietarioTXTdb.ToString();
            idPropietarioTXT.Text = respuesta[0].idPropietarioTXTdb.ToString();
            telefonoPropietarioTXT.Text = respuesta[0].telefonoPropietarioTXTdb.ToString();
            celularPropietarioTXT.Text = respuesta[0].celularPropietarioTXTdb.ToString();
            emailPropietarioTXT.Text = respuesta[0].emailPropietarioTXTdb.ToString();
            direccionPropietarioTXT.Text = respuesta[0].direccionPropietarioTXTdb.ToString();
            
            nombreArrendatarioTXT.Text = respuesta[0].nombreArrendatarioTXTdb.ToString();
            idArrendatarioTXT.Text = respuesta[0].idArrendatarioTXTdb.ToString();
            telefonoArrendatarioTXT.Text = respuesta[0].telefonoArrendatarioTXTdb.ToString();
            celularArrendatarioTXT.Text = respuesta[0].celularArrendatarioTXTdb.ToString();
            emailArrendatarioTXT.Text = respuesta[0].emailArrendatarioTXTdb.ToString();
            direccionArrendatarioTXT.Text = respuesta[0].direccionArrendatarioTXTdb.ToString();
            */







        }

        private void fechaIniTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            
                
            
        }

        private void fechaIniTXT_ValueChanged(object sender, EventArgs e)
        {
          
        }

        private void fechaFinTXT_ValueChanged(object sender, EventArgs e)
        {

        }

        private void contratoTXT_TextChanged(object sender, EventArgs e)
        {
           


        }

        private void contratoTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void idPropietarioTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void idPropietarioTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void idArrendatarioTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void idArrendatarioTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void administracionTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void administracionTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void idCoarrendatario1TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void idCoarrendatario2TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void idCoarrendatario3TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void idCoarrendatario4TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoPropietarioTXT_TextChanged(object sender, EventArgs e)
        {
        }

        private void telefonoPropietarioTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoArrendatarioTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void telefonoArrendatarioTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularPropietarioTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularArrendatarioTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoCoarrendatario1TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void telefonoCoarrendatario1TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularCoarrendatario1TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoCoarrendatario2TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void telefonoCoarrendatario2TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularCoarrendatario2TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void celularCoarrendatario2TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoCoarrendatario3TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void telefonoCoarrendatario3TXT_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void telefonoCoarrendatario3TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularCoarrendatario3TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void celularCoarrendatario3TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoCoarrendatario4TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularCoarrendatario4TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void contratoGroup_Enter(object sender, EventArgs e)
        {

        }

        private void button7_Click_2(object sender, EventArgs e)
        {
            dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=mostrar");
            fechaIniTXT.Text = respuesta[0].fechaIniTXTdb.ToString();
        }
    }
}
