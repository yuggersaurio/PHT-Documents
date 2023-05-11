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


        public void generarWord()
        {


            //ESCRIBO ARCHIVO EN WORD
            var rutaArchivo = @"\\servidor1\Fotos\FOTOS_FIRMA_DE_CONTRATOS\CTO_" + contratoTXT.Text + @"\" + @"DOCUMENTOS\" + contratoTXT.Text + ".rtf";
            System.IO.File.WriteAllText(rutaArchivo, richTextBox1.Rtf);
            //------------------------------------

            //ABRO AUTOMATICAMENTE ARCHIVO GENERADO
            Process.Start(rutaArchivo);
            //------------------------------------


        }

        public void vistaPrevia(){

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

                fechaFin = fechaFinTXT.Text;
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaIniTXT**", fechaInicio.Humanize(LetterCasing.AllCaps));
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaFinTXT**", fechaFin.Humanize(LetterCasing.AllCaps));

                //------------------------------------------









            }


            //------------------------------------FIN DE GENERAR REMPLAZO VISTA PREVIA

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

        public void cargarFormato(string tipoFormato)
        {
            int i;
           for(i=0 ; i<=1; i++) {
                //TRAIGO EL FORMATO CORRESPONDIENTE

            string URL = "";
            object readOnly = true;
            object visible = true;
            object save = false;
                if (i == 0) { URL = @"\\servidor1\sistemas\PROYECTOS\Exporter\Formatos\formato1Arren.docx"; } else 
                            { URL = @"\\servidor1\sistemas\PROYECTOS\Exporter\Formatos\formato1Prop.docx"; }
            object fileName = URL;
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
                if (i == 0) { richTextBox1.Rtf = data.GetData(DataFormats.Rtf).ToString(); } else { richTextBox2.Rtf = data.GetData(DataFormats.Rtf).ToString(); }
            
            oWord.Quit(ref missing, ref missing, ref missing);


            };

            //-----------------------FIN DE TRAER CONTRATO


            //DESHABILITO LO QUE NO SE NECESITA
            // propietarioGroup.Enabled = false;
            //----------------------------

            //COPIAS POR DEFECTO
            copiasTXT.SelectedIndex = 3;

            //VIGENCIA POR DEFECTO
            vigenciaTXT.SelectedIndex = 1;

            //CAMBIO DESTINO
            destinoTXT.Text = "Vivienda";





        }


        public void limpiarCampos() // LIMPIO LOS TEXTBOX
        {

            TextBox[] arrText = new TextBox[] {    
            contratoTXT,
            direccionTXT,
            ciudadTXT,                      
            canonTXT,
            administracionTXT,
            destinoTXT,
            nombrePropietarioTXT,
            idPropietarioTXT,
            telefonoPropietarioTXT,
            celularPropietarioTXT,
            emailPropietarioTXT,
            direccionPropietarioTXT,
            nombreArrendatarioTXT,
            idArrendatarioTXT,
            telefonoArrendatarioTXT,
            celularArrendatarioTXT,
            emailArrendatarioTXT,
            direccionArrendatarioTXT,
            nombreCoarrendatario1TXT,
            idCoarrendatario1TXT,
            telefonoCoarrendatario1TXT,
            celularCoarrendatario1TXT,
            emailCoarrendatario1TXT,
            direccionCoarrendatario1TXT,
             nombreCoarrendatario2TXT,
            idCoarrendatario2TXT,
            telefonoCoarrendatario2TXT,
            celularCoarrendatario2TXT,
            emailCoarrendatario2TXT,
            direccionCoarrendatario2TXT,
            nombreCoarrendatario3TXT,
            idCoarrendatario3TXT,
            telefonoCoarrendatario3TXT,
            celularCoarrendatario3TXT,
            emailCoarrendatario3TXT,
            direccionCoarrendatario3TXT,
            nombreCoarrendatario4TXT,
            idCoarrendatario4TXT,
            telefonoCoarrendatario4TXT,
            celularCoarrendatario4TXT,
            emailCoarrendatario4TXT,
            direccionCoarrendatario4TXT  };  

            for (int i = 0, len = arrText.Length; i < len; i++)  {             
               arrText[i].Text = "";
                          }


        }


        public void listarContratos()

            
        {
            listaContratosBox.Items.Clear();
            dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=listar");
            // 
            //  Console.WriteLine(respuesta.Count);

            int i = 0;
            while (i < respuesta.Count)
            {
                listaContratosBox.Items.Add(respuesta[i].contratoTXTdb.ToString());
                
                i++;
            }

        }

        public void guardarContrato(string tipoConsultaSQL, string formato)
        {


        
            dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=" + tipoConsultaSQL

                + "&contratoTXTdb=" + contratoTXT.Text
                + "&direccionTXTdb=" + direccionTXT.Text
                + "&ciudadTXTdb=" + ciudadTXT.Text
                + "&fechaIniTXTdb=" + fechaIniTXT.Text
                + "&fechaFinTXTdb=" + fechaFinTXT.Text
                + "&canonTXTdb=" + canonTXT.Text
                + "&administracionTXTdb=" + administracionTXT.Text
                + "&destinoTXTdb=" + destinoTXT.Text
                + "&vigenciaTXTdb=" + vigenciaTXT.Text
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
        }
        public void traerInfoApi(string Contrato)
        {

            dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=mostrar&contratoConsulta="+Contrato);
          
            contratoTXT.Text = respuesta[0].contratoTXTdb.ToString();
            direccionTXT.Text = respuesta[0].direccionTXTdb.ToString();
            ciudadTXT.Text = respuesta[0].ciudadTXTdb.ToString();
            fechaIniTXT.Text = respuesta[0].fechaIniTXTdb.ToString();
            fechaFinTXT.Text = respuesta[0].fechaFinTXTdb.ToString();
            canonTXT.Text = respuesta[0].canonTXTdb.ToString();
            administracionTXT.Text = respuesta[0].administracionTXTdb.ToString();
            destinoTXT.Text = respuesta[0].destinoTXTdb.ToString();
            vigenciaTXT.Text = respuesta[0].vigenciaTXTdb.ToString();

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

            nombreCoarrendatario1TXT.Text = respuesta[0].nombreCoarrendatario1TXTdb.ToString();
            idCoarrendatario1TXT.Text = respuesta[0].idCoarrendatario1TXTdb.ToString();
            telefonoCoarrendatario1TXT.Text = respuesta[0].telefonoCoarrendatario1TXTdb.ToString();
            celularCoarrendatario1TXT.Text = respuesta[0].celularCoarrendatario1TXTdb.ToString();
            emailCoarrendatario1TXT.Text = respuesta[0].emailCoarrendatario1TXTdb.ToString();
            direccionCoarrendatario1TXT.Text = respuesta[0].direccionCoarrendatario1TXTdb.ToString();

            nombreCoarrendatario2TXT.Text = respuesta[0].nombreCoarrendatario2TXTdb.ToString();
            idCoarrendatario2TXT.Text = respuesta[0].idCoarrendatario2TXTdb.ToString();
            telefonoCoarrendatario2TXT.Text = respuesta[0].telefonoCoarrendatario2TXTdb.ToString();
            celularCoarrendatario2TXT.Text = respuesta[0].celularCoarrendatario2TXTdb.ToString();
            emailCoarrendatario2TXT.Text = respuesta[0].emailCoarrendatario2TXTdb.ToString();
            direccionCoarrendatario2TXT.Text = respuesta[0].direccionCoarrendatario2TXTdb.ToString();

            nombreCoarrendatario3TXT.Text = respuesta[0].nombreCoarrendatario3TXTdb.ToString();
            idCoarrendatario3TXT.Text = respuesta[0].idCoarrendatario3TXTdb.ToString();
            telefonoCoarrendatario3TXT.Text = respuesta[0].telefonoCoarrendatario3TXTdb.ToString();
            celularCoarrendatario3TXT.Text = respuesta[0].celularCoarrendatario3TXTdb.ToString();
            emailCoarrendatario3TXT.Text = respuesta[0].emailCoarrendatario3TXTdb.ToString();
            direccionCoarrendatario3TXT.Text = respuesta[0].direccionCoarrendatario3TXTdb.ToString();

            nombreCoarrendatario4TXT.Text = respuesta[0].nombreCoarrendatario4TXTdb.ToString();
            idCoarrendatario4TXT.Text = respuesta[0].idCoarrendatario4TXTdb.ToString();
            telefonoCoarrendatario4TXT.Text = respuesta[0].telefonoCoarrendatario4TXTdb.ToString();
            celularCoarrendatario4TXT.Text = respuesta[0].celularCoarrendatario4TXTdb.ToString();
            emailCoarrendatario4TXT.Text = respuesta[0].emailCoarrendatario4TXTdb.ToString();
            direccionCoarrendatario4TXT.Text = respuesta[0].direccionCoarrendatario4TXTdb.ToString();

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
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario1TXT**", telefonoCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario1TXT**", celularCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario1TXT**", direccionCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario1TXT**", emailCoarrendatario1TXT.Text);
                        break;
                    case 3:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario2TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario2TXT**", nombreCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario2TXT**", telefonoCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario2TXT**", celularCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario2TXT**", direccionCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario2TXT**", emailCoarrendatario2TXT.Text);
                        break;
                    case 4:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario3TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario3TXT**", nombreCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario3TXT**", telefonoCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario3TXT**", celularCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario3TXT**", direccionCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario3TXT**", emailCoarrendatario3TXT.Text);

                        break;
                    case 5:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario4TXT**", terceroFormateado);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario4TXT**", nombreCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario4TXT**", telefonoCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario4TXT**", celularCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario4TXT**", direccionCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario4TXT**", emailCoarrendatario4TXT.Text);

                        break;
                    
                }




            };

            //------------------------------------------

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            listarContratos();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            

            /*Word.Application word = new Word.Application();
            word.Visible = true;
            word.WindowState = Word.WdWindowState.wdWindowStateNormal;
            Word.Document doc = word.Documents.Add();
            Word.Paragraph paragraph;
            paragraph = doc.Paragraphs.Add();

            
            paragraph.Range.Text = richTextBox1.Text;
            doc.SaveAs(@"c:\RZ\mydoc.rtf");
            doc.Close();
                word.Quit();*/
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
            fechaIniTXT.CustomFormat="dd-MMM-yyyy";
        }

        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void button5_Click(object sender, EventArgs e)
        {
            
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            e.Graphics.DrawString(richTextBox1.Text, richTextBox1.Font, Brushes.Black, 100, 100);

        }

        private void button6_Click(object sender, EventArgs e)
        {

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
            cargarFormato("formato1");

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
          



        }

        private void fechaIniTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            
                
            
        }

        private void fechaIniTXT_ValueChanged(object sender, EventArgs e)
        {
            fechaFinTXT.Value = fechaIniTXT.Value;
            fechaFinTXT.Value = DateTime.Now.AddYears(1);
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
     
         
        }

        private void button7_Click_3(object sender, EventArgs e)
        {
            



        }

        private void listaContratosBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Console.WriteLine("seleccionado es " + listaContratosBox.Text);
           // menuStrip1.Items().Enabled = true;
            //guardarBTN.Text = "Duplicar";
            traerInfoApi(listaContratosBox.Text);
            GuardarDuplicarToolStripMenuItem.Text = "Duplicar cómo contrato nuevo";
            GuardartoolStripMenuItem.Enabled = true;
        }

        private void nuevoContratoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            datosContratoTab.Enabled = true;
        }

        private void formatosToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void archivoToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void nuevoContratoToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            datosContratoTab.Enabled = true;
            //guardarBTN.Enabled = true;
            //guardarBTN.Text = "Guardar";
            limpiarCampos();
            GuardarDuplicarToolStripMenuItem.Enabled = true;
            GuardartoolStripMenuItem.Enabled = false;

        }

        private void button1_Click_1(object sender, EventArgs e)
        {
           
           

        }

        private void button1_Click_2(object sender, EventArgs e)
        {
            
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            guardarContrato("actualizar", "formato1");
            listarContratos();
        }

        private void generarVistaPreviaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            vistaPrevia();
        }

        private void generarContratosEnWORDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            generarWord();
        }

        private void duplicarComoContratoNuevoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            guardarContrato("insertar", "formato1");
        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void button1_Click_3(object sender, EventArgs e)
        {
          
        }

        private void propietarioGroup_Enter(object sender, EventArgs e)
        {

        }

        private void arrendatarioGroup_Enter(object sender, EventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void richTextBox3_TextChanged(object sender, EventArgs e)
        {

        }

        private void label65_Click(object sender, EventArgs e)
        {

        }
    }
}
