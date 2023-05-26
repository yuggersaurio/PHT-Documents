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
using System.Threading.Tasks;
using System.Threading;
using System.IO;

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


            //VERIFICO SI DIRECTORIO EXISTE
            var rutaCarpeta = @"\\servidor1\Fotos\FOTOS_FIRMA_DE_CONTRATOS\CTO_" + contratoTXT.Text + @"\" + @"DOCUMENTOS\CONTRATOS PRELIMINARES";

            if (!Directory.Exists(rutaCarpeta))
            {
                Console.WriteLine("Creando el directorio: {0}", rutaCarpeta);
                DirectoryInfo di = Directory.CreateDirectory(rutaCarpeta);
            }

            //--


            //ESCRIBO ARCHIVO EN WORD ARRENDATARIO
            var rutaArchivo = @"\\servidor1\Fotos\FOTOS_FIRMA_DE_CONTRATOS\CTO_" + contratoTXT.Text + @"\" + @"DOCUMENTOS\CONTRATOS PRELIMINARES\" + contratoTXT.Text + " ARRENDATARIO.rtf";

            try
            {
                System.IO.File.WriteAllText(rutaArchivo, richTextBox1.Rtf);
            }
            catch (System.IO.IOException IOEx)
            {
                MessageBox.Show("El contrato de ARRENDATARIO que esta tratando de generar está en uso, cierrelo y presione exportar nuevamente", "Advertencia");

            }
            //------------------------------------

            //ESCRIBO ARCHIVO EN WORD PROPIETARIO
            var rutaArchivo2 = @"\\servidor1\Fotos\FOTOS_FIRMA_DE_CONTRATOS\CTO_" + contratoTXT.Text + @"\" + @"DOCUMENTOS\CONTRATOS PRELIMINARES\" + contratoTXT.Text + " PROPIETARIO.rtf";
            try
            {
                System.IO.File.WriteAllText(rutaArchivo2, richTextBox2.Rtf);
            }
            catch (System.IO.IOException IOEx)
            {
                MessageBox.Show("El contrato de PROPIETARIO que esta tratando de generar está en uso, cierrelo y presione exportar nuevamente", "Advertencia");
            }
            //------------------------------------

            //ABRO AUTOMATICAMENTE ARCHIVOS GENERADO
            Process.Start(rutaArchivo);
            Process.Start(rutaArchivo2);
            //------------------------------------


        }
        public void calcularFechasContrato()
        {
            DateTime now = DateTime.Now;
            var primerDiaMes = new DateTime(now.Year, now.Month, 1);
            primerDiaMes = primerDiaMes.AddMonths(1);
            var ultimoDiaMes = primerDiaMes.AddMonths(12).AddSeconds(-1);
            fechaIniTXT.Value = primerDiaMes;
            fechaFinTXT.Value = ultimoDiaMes;

        }
        public void vistaPrevia(){

            //REALIZO REMPLAZO DE VARIABLES EN VISTA PREVIA

            richTextBox1.SelectAll();

            string[] textArray = richTextBox1.SelectedText.Split(new char[] { '\n' });

            foreach (string strText in textArray)
            {
                if (!string.IsNullOrEmpty(strText)) ; 

                
                
            }

            //---Llenamos formato de arrendatario
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**contratoTXT**", contratoTXT.Text);
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**ciudadTXT**", ciudadTXT.Text);
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**destinoTXT**", destinoTXT.Text);
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionTXT**", direccionTXT.Text);
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**vigenciaTXT**", vigenciaTXT.Text);
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**copiasTXT**", copiasTXT.Text);
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**barrioTXT**", barrioTXT.Text);
            
            //--


            //---Llenamos formato de propietario
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**contratoTXT**", contratoTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**ciudadTXT**", ciudadTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**destinoTXT**", destinoTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**direccionTXT**", direccionTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**vigenciaTXT**", vigenciaTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**copiasTXT**", copiasTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**barrioTXT**", barrioTXT.Text);
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**inmuebleTXT**", inmuebleTXT.Text);
            //--

            //--LLENAMOS QUIEN EXPORTO LOS ARCHIVOS
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**usuarioTXT**", usuarioTXT.Text + " - "+ DateTime.Now);
            //--




            //LLAMAMOS FUNCION PARA FORMATEAR CANON Y ADMIN Y REGRESAR LETRAS
            canonAdminLetrasFn(canonTXT.Text, "canon", 12);
            canonAdminLetrasFn(administracionTXT.Text, "admin", 12);
            int canonadmin = int.Parse(canonTXT.Text) + int.Parse(administracionTXT.Text);
            canonAdminLetrasFn(canonadmin.ToString(), "canonadmin", 12);
            //------------------------------------------

            ;

            //LLAMAMOS FUNCION PARA FORMATEAR ID TERCEROS
            string[] arr = new string[8];
            arr[0] = idArrendatarioTXT.Text;
            arr[1] = idPropietarioTXT.Text;
            arr[2] = idCoarrendatario1TXT.Text;
            arr[3] = idCoarrendatario2TXT.Text;
            arr[4] = idCoarrendatario3TXT.Text;
            arr[5] = idCoarrendatario4TXT.Text;
            arr[6] = idEncargadoTXT.Text;
            arr[7] = idCoarrendatario5TXT.Text;

            var tercero = idArrendatarioTXT.Text;
            var tipoTercero = "arrendatario";
            terceroFn(arr, tipoTercero);

            //------------------------------------------


            //FECHAS DE INICIO Y FIN EN LETRAS
            string fechaInicio = fechaIniTXT.Text;
            string fechaFin = fechaFinTXT.Text;

            fechaFin = fechaFinTXT.Text;
            //LLenamos formato de arrendatario
           
            string porcentajeLetras = "";
            string[] fechaIniDividida = fechaInicio.Split(','); //---LE QUITO EL DIA
            string[] fechaFinDividida = fechaFin.Split(','); //---LE QUITO EL DIA

            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaIniTXT**", fechaIniDividida[1].Humanize(LetterCasing.AllCaps));
            richTextBox1.Rtf = richTextBox1.Rtf.Replace("**fechaFinTXT**", fechaFinDividida[1].Humanize(LetterCasing.AllCaps));
            //--

            //LLenamos formato de propietario
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**fechaIniTXT**", fechaIniDividida[1].Humanize(LetterCasing.AllCaps));
            richTextBox2.Rtf = richTextBox2.Rtf.Replace("**fechaFinTXT**", fechaFinDividida[1].Humanize(LetterCasing.AllCaps));
            //--

            //PORCENTAJES A LETRAS

            porcentajesLetras(clausulaTXT.Text, "clausula");
            porcentajesLetras(servicioCTXT.Text, "servicio");

            //--------------------

     

            generarContratosEnWORDToolStripMenuItem.Enabled = true;
            btnExportar.Enabled = true;
            btnExportarTXT.Enabled = true;
            datosContratoTab.SelectedIndex = 2;
            informacionTXT.Text = @"Vista previa generada, para generar una nueva vista previa debe cargar un formato nuevamente";
            informacionTXT.ForeColor = Color.Chocolate;


            //------------------------------------FIN DE GENERAR REMPLAZO VISTA PREVIA

        }
        private void canonAdminLetrasFn(string numero, string tipo, int vigencia) //FUNCION QUE FORMATEA VALORES DE CANON Y ADMIN, CALCULA CUANTIA Y DEVUELVE VALORES EN LETRAS E IMPRIME RESULTADO EN FORMATO
        {
            int numeroInteger;
            int cuantiaInteger;
            string dePesos = " DE PESOS M/CTE ";
            int numeroSubstraido;

            




            string numeroFormateado = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(numero));
            numeroFormateado = numeroFormateado.Substring(0, numeroFormateado.Length - 3);

            
            int.TryParse(numero, out numeroInteger);
            
            cuantiaInteger = numeroInteger * vigencia;            
            string cuantiaFormateado = string.Format(CultureInfo.CreateSpecificCulture("es-CO"), "{00:C}", double.Parse(cuantiaInteger.ToString()));
            cuantiaFormateado = cuantiaFormateado.Substring(0, cuantiaFormateado.Length - 3);



            //--- AGRAGAMOS PALABRA PESOS/DE PESOS


            string substr1 = numero.Substring(1);

            int.TryParse(substr1, out numeroSubstraido);
            int.TryParse(numero, out numeroInteger);

            if ((numeroSubstraido >= 1) && (numero.Length >= 7))
            {
                dePesos = " PESOS M/CTE ";
            }
            if (numeroInteger < 1000000)
            {
                dePesos = " PESOS M/CTE ";
            }

            //---



            switch (tipo)
            {
                case "canon":
                     

                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**canonTXT**", (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + numeroFormateado + " ) " );
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**cuantiaTXT**", (cuantiaInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + cuantiaFormateado + " ) " );

                    richTextBox2.Rtf = richTextBox2.Rtf.Replace("**canonTXT**", (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + numeroFormateado + " ) ");
                    richTextBox2.Rtf = richTextBox2.Rtf.Replace("**cuantiaTXT**", (cuantiaInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + cuantiaFormateado + " ) ");
                    break;
                case "admin":
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**administracionTXT**",  (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + numeroFormateado + " ) " );
                    richTextBox2.Rtf = richTextBox2.Rtf.Replace("**administracionTXT**", (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + numeroFormateado + " ) ");
                    break;

                case "canonadmin":
                    richTextBox1.Rtf = richTextBox1.Rtf.Replace("**canonadminTXT**", (numeroInteger.ToWords()).Humanize(LetterCasing.AllCaps) + dePesos + "( " + numeroFormateado + " ) ");
                   
                    break;
            }



        }

        public void checkCoarrendatarios(string Checkeado, string numeroCoarrendatario)
        {

            if (Checkeado=="check")
            {
                Console.WriteLine(Checkeado,numeroCoarrendatario);
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**lineaCoarrendatario"+ numeroCoarrendatario + "TXT**", "_____________________________");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloCoarrendatario" + numeroCoarrendatario + "TXT**", "COARRENDATARIO");
                
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloDireccionCoarrendatario" + numeroCoarrendatario + "TXT**", "Dirección: ");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloTelefonoCoarrendatario" + numeroCoarrendatario + "TXT**", "Teléfono: ");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloCelularCoarrendatario" + numeroCoarrendatario + "TXT**", "Celular: ");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloEmailCoarrendatario" + numeroCoarrendatario + "TXT**", "Email: ");
            }
            else
            {
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**lineaCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloDireccionCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloTelefonoCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloCelularCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**tituloEmailCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario" + numeroCoarrendatario + "TXT**", "");
                richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario" + numeroCoarrendatario + "TXT**", "");
            }

        }

        private async 
        Task
cargarFormato(string tipoFormato)
        {
            logoApp.Visible = false;
            informacionTXT.Text = "Cargando formato, por favor espere...";
            informacionTXT.ForeColor = Color.Blue;

            int i;
           for(i=0 ; i<=1; i++) {
                //TRAIGO EL FORMATO CORRESPONDIENTE

            string URL = "";
            object readOnly = true;
            object visible = true;
            object save = false;
                if (i == 0) { URL = @"\\servidor1\sistemas\PROYECTOS\Exporter\Formatos\" + tipoFormato + @"\" + tipoFormato + "Arren.docx"; } else 
                            { URL = @"\\servidor1\sistemas\PROYECTOS\Exporter\Formatos\" + tipoFormato + @"\" + tipoFormato + "Prop.docx"; }
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

           

            //----ACTIVAMOS BOTONES Y TABS
       
            btnGuardar.Enabled = true;
            btnGuardarTXT.Enabled = true;
            btnVista.Enabled = true;
            btnVistaTXT.Enabled = true;
            generarVistaPreviaToolStripMenuItem.Enabled = true;

            ///------------
            

            logoApp.Visible = true;
            informacionTXT.Text = "Formatos cargados actualmente: Inmueble sometido a copropiedad y destinado a vivienda urbana, arrendatario y propietario: ";
            informacionTXT.ForeColor = Color.Green;


        }
        public void porcentajesLetras(string valorPorcentaje, string tipoPorcentaje)
        {
                      
            string[] words = valorPorcentaje.Split('.');
            int validarDecimalServicio;
            int.TryParse(servicioCTXT.Text, out validarDecimalServicio);

            int porcentajeInteger;
            int i = 0;
            string porcentajeLetras = "";


            if (clausulaTXT.Text.Contains("."))
            {
                Console.WriteLine("DECIMAL");

            }


            foreach (var word in words)
            {
                int.TryParse(word, out porcentajeInteger);


                if (i == 0 && valorPorcentaje.Contains(".")) { porcentajeLetras += porcentajeInteger.ToWords().Humanize(LetterCasing.AllCaps) + " PUNTO "; }
                else
                { porcentajeLetras += porcentajeInteger.ToWords().Humanize(LetterCasing.AllCaps); };
                i++;

                ;
            }
            switch (tipoPorcentaje)
            {
                case "clausula":
                    richTextBox2.Rtf = richTextBox2.Rtf.Replace("**clausulaTXT**", porcentajeLetras + " PORCIENTO ( " + valorPorcentaje + " % )");
                    Console.WriteLine(porcentajeLetras);
                    break;
                case "servicio":
                    richTextBox2.Rtf = richTextBox2.Rtf.Replace("**servicioCTXT**", porcentajeLetras + " PORCIENTO ( " + valorPorcentaje + " % )");
                    Console.WriteLine(porcentajeLetras);
                    break;
            }                    
            //richTextBox2.Rtf = richTextBox2.Rtf.Replace("**servicioCTXT**", porcentajeLetras);
            




        }

        public void limpiarCampos() // LIMPIO LOS TEXTBOX
        {

            TextBox[] arrText = new TextBox[] {    
            contratoTXT,
            direccionTXT,
            ciudadTXT,
            barrioTXT,
            inmuebleTXT,
            canonTXT,
            administracionTXT,
            destinoTXT,
            servicioCTXT,
            clausulaTXT,
            nombrePropietarioTXT,
            idPropietarioTXT,
            telefonoPropietarioTXT,
            celularPropietarioTXT,
            emailPropietarioTXT,
            direccionPropietarioTXT,
            nombreEncargadoTXT,
            idEncargadoTXT,
            telefonoEncargadoTXT,
            celularEncargadoTXT,
            emailEncargadoTXT,
            direccionEncargadoTXT,
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

            btnDuplicar.Visible = false;
            btnGuardarTXT.Text = "Guardar";
            btnActualizar.Enabled = false;
            btnActualizarTXT.Enabled = false;
            btnVista.Enabled = false;
            btnVistaTXT.Enabled = false;
            btnExportar.Enabled = false;
            btnExportarTXT.Enabled = false;



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
            btnExportar.Visible = true;

            


        }

        public void guardarContrato(string tipoConsultaSQL, string formato)
        {


        
            dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=" + tipoConsultaSQL

                + "&contratoTXTdb=" + contratoTXT.Text
                + "&direccionTXTdb=" + direccionTXT.Text
                + "&ciudadTXTdb=" + ciudadTXT.Text
                + "&barrioTXTdb=" + barrioTXT.Text
                + "&inmuebleTXTdb=" + inmuebleTXT.Text
                + "&fechaIniTXTdb=" + fechaIniTXT.Text
                + "&fechaFinTXTdb=" + fechaFinTXT.Text
                + "&canonTXTdb=" + canonTXT.Text
                + "&administracionTXTdb=" + administracionTXT.Text
                + "&destinoTXTdb=" + destinoTXT.Text
                + "&vigenciaTXTdb=" + vigenciaTXT.Text
                + "&servicioCTXTdb=" + servicioCTXT.Text
                + "&clausulaTXTdb=" + clausulaTXT.Text

                + "&nombrePropietarioTXTdb=" + nombrePropietarioTXT.Text
                + "&idPropietarioTXTdb=" + idPropietarioTXT.Text
                + "&tipoIdPropietarioTXTdb=" + tipoIdPropietarioTXT.Text
                + "&telefonoPropietarioTXTdb=" + telefonoPropietarioTXT.Text
                + "&celularPropietarioTXTdb=" + celularPropietarioTXT.Text
                + "&emailPropietarioTXTdb=" + emailPropietarioTXT.Text
                + "&direccionPropietarioTXTdb=" + direccionPropietarioTXT.Text

                + "&nombreEncargadoTXTdb=" + nombreEncargadoTXT.Text
                + "&tipoIdEncargadoTXTdb=" + tipoIdEncargadoTXT.Text
                + "&idEncargadoTXTdb=" + idEncargadoTXT.Text
                + "&telefonoEncargadoTXTdb=" + telefonoEncargadoTXT.Text
                + "&celularEncargadoTXTdb=" + celularEncargadoTXT.Text
                + "&emailEncargadoTXTdb=" + emailEncargadoTXT.Text
                + "&direccionEncargadoTXTdb=" + direccionEncargadoTXT.Text

                + "&nombreArrendatarioTXTdb=" + nombreArrendatarioTXT.Text
                + "&tipoIdArrendatarioTXTdb=" + tipoIdArrendatarioTXT.Text
                + "&idArrendatarioTXTdb=" + idArrendatarioTXT.Text
                + "&telefonoArrendatarioTXTdb=" + telefonoArrendatarioTXT.Text
                + "&celularArrendatarioTXTdb=" + celularArrendatarioTXT.Text
                + "&emailArrendatarioTXTdb=" + emailArrendatarioTXT.Text
                + "&direccionArrendatarioTXTdb=" + direccionArrendatarioTXT.Text

                + "&nombreCoarrendatario1TXTdb=" + nombreCoarrendatario1TXT.Text
                + "&tipoIdCoarrendatario1TXTdb=" + tipoIdCoarrendatario1TXT.Text
                + "&idCoarrendatario1TXTdb=" + idCoarrendatario1TXT.Text
                + "&telefonoCoarrendatario1TXTdb=" + telefonoCoarrendatario1TXT.Text
                + "&celularCoarrendatario1TXTdb=" + celularCoarrendatario1TXT.Text
                + "&direccionCoarrendatario1TXTdb=" + direccionCoarrendatario1TXT.Text
                + "&emailCoarrendatario1TXTdb=" + emailCoarrendatario1TXT.Text

                + "&nombreCoarrendatario2TXTdb=" + nombreCoarrendatario2TXT.Text
                + "&tipoIdCoarrendatario2TXTdb=" + tipoIdCoarrendatario2TXT.Text
                + "&idCoarrendatario2TXTdb=" + idCoarrendatario2TXT.Text
                + "&telefonoCoarrendatario2TXTdb=" + telefonoCoarrendatario2TXT.Text
                + "&celularCoarrendatario2TXTdb=" + celularCoarrendatario2TXT.Text
                + "&direccionCoarrendatario2TXTdb=" + direccionCoarrendatario2TXT.Text
                + "&emailCoarrendatario2TXTdb=" + emailCoarrendatario2TXT.Text

                + "&nombreCoarrendatario3TXTdb=" + nombreCoarrendatario3TXT.Text
                + "&tipoIdCoarrendatario3TXTdb=" + tipoIdCoarrendatario3TXT.Text
                + "&idCoarrendatario3TXTdb=" + idCoarrendatario3TXT.Text
                + "&telefonoCoarrendatario3TXTdb=" + telefonoCoarrendatario3TXT.Text
                + "&celularCoarrendatario3TXTdb=" + celularCoarrendatario3TXT.Text
                + "&direccionCoarrendatario3TXTdb=" + direccionCoarrendatario3TXT.Text
                + "&emailCoarrendatario3TXTdb=" + emailCoarrendatario3TXT.Text

                + "&nombreCoarrendatario4TXTdb=" + nombreCoarrendatario4TXT.Text
                + "&tipoIdCoarrendatario4TXTdb=" + tipoIdCoarrendatario4TXT.Text
                + "&idCoarrendatario4TXTdb=" + idCoarrendatario4TXT.Text
                + "&telefonoCoarrendatario4TXTdb=" + telefonoCoarrendatario4TXT.Text
                + "&celularCoarrendatario4TXTdb=" + celularCoarrendatario4TXT.Text
                + "&direccionCoarrendatario4TXTdb=" + direccionCoarrendatario4TXT.Text
                + "&emailCoarrendatario4TXTdb=" + emailCoarrendatario4TXT.Text
                + "&modificacionTXTdb=" + "Ultima modificación por " + usuarioTXT.Text + " " + DateTime.Now


                );
            Console.WriteLine("&servicioCTXTdb=" + servicioCTXT.Text
                + "&clausulaTXTdb=" + clausulaTXT.Text);

            //ACTIVAMOS BOTONES Y TABS

            btnDuplicar.Visible = true;
            btnGuardarTXT.Text = "Duplicar";
            btnActualizar.Enabled = true;
            btnActualizarTXT.Enabled = true;
            GuardarDuplicarToolStripMenuItem.Text = "Duplicar cómo contrato nuevo";
            GuardartoolStripMenuItem.Enabled = true;
            btnVerCarpeta.Enabled = true;
            btnVerCarpetaTXT.Enabled = true;


            //-----

            listarContratos(); //---LISTAMOS CONTRATOS EN COMBOBOX

            modificacionTXT.Text = "Ultima modificación por " + usuarioTXT.Text + " " + DateTime.Now;

        }
        public void traerInfoApi(string Contrato)
        {

            dynamic respuesta = dBApi.Get("https://portalhouses.com/administrador/ApiDocuments/post.php?tipo=mostrar&contratoConsulta="+Contrato);
          
            contratoTXT.Text = respuesta[0].contratoTXTdb.ToString();
            direccionTXT.Text = respuesta[0].direccionTXTdb.ToString();
            ciudadTXT.Text = respuesta[0].ciudadTXTdb.ToString();
            inmuebleTXT.Text = respuesta[0].inmuebleTXTdb.ToString();
            barrioTXT.Text = respuesta[0].barrioTXTdb.ToString();
            fechaIniTXT.Text = respuesta[0].fechaIniTXTdb.ToString();
            fechaFinTXT.Text = respuesta[0].fechaFinTXTdb.ToString();
            canonTXT.Text = respuesta[0].canonTXTdb.ToString();
            administracionTXT.Text = respuesta[0].administracionTXTdb.ToString();

            destinoTXT.Text = respuesta[0].destinoTXTdb.ToString();
            vigenciaTXT.Text = respuesta[0].vigenciaTXTdb.ToString();
            vigenciaTXT.Text = respuesta[0].vigenciaTXTdb.ToString();
            clausulaTXT.Text = respuesta[0].clausulaTXTdb.ToString();
            servicioCTXT.Text = respuesta[0].servicioCTXTdb.ToString();


            nombrePropietarioTXT.Text = respuesta[0].nombrePropietarioTXTdb.ToString();
            tipoIdPropietarioTXT.Text = respuesta[0].tipoIdPropietarioTXTdb.ToString();
            idPropietarioTXT.Text = respuesta[0].idPropietarioTXTdb.ToString();
            telefonoPropietarioTXT.Text = respuesta[0].telefonoPropietarioTXTdb.ToString();
            celularPropietarioTXT.Text = respuesta[0].celularPropietarioTXTdb.ToString();
            emailPropietarioTXT.Text = respuesta[0].emailPropietarioTXTdb.ToString();
            direccionPropietarioTXT.Text = respuesta[0].direccionPropietarioTXTdb.ToString();

            nombreEncargadoTXT.Text = respuesta[0].nombreEncargadoTXTdb.ToString();
            tipoIdEncargadoTXT.Text = respuesta[0].tipoIdEncargadoTXTdb.ToString();
            idEncargadoTXT.Text = respuesta[0].idEncargadoTXTdb.ToString();
            telefonoEncargadoTXT.Text = respuesta[0].telefonoEncargadoTXTdb.ToString();
            celularEncargadoTXT.Text = respuesta[0].celularEncargadoTXTdb.ToString();
            emailEncargadoTXT.Text = respuesta[0].emailEncargadoTXTdb.ToString();
            direccionEncargadoTXT.Text = respuesta[0].direccionEncargadoTXTdb.ToString();

            nombreArrendatarioTXT.Text = respuesta[0].nombreArrendatarioTXTdb.ToString();
            idArrendatarioTXT.Text = respuesta[0].idArrendatarioTXTdb.ToString();
            tipoIdArrendatarioTXT.Text = respuesta[0].tipoIdArrendatarioTXTdb.ToString();
            telefonoArrendatarioTXT.Text = respuesta[0].telefonoArrendatarioTXTdb.ToString();
            celularArrendatarioTXT.Text = respuesta[0].celularArrendatarioTXTdb.ToString();
            emailArrendatarioTXT.Text = respuesta[0].emailArrendatarioTXTdb.ToString();
            direccionArrendatarioTXT.Text = respuesta[0].direccionArrendatarioTXTdb.ToString();

            nombreCoarrendatario1TXT.Text = respuesta[0].nombreCoarrendatario1TXTdb.ToString();
            idCoarrendatario1TXT.Text = respuesta[0].idCoarrendatario1TXTdb.ToString();
            tipoIdCoarrendatario1TXT.Text = respuesta[0].tipoIdCoarrendatario1TXTdb.ToString();
            telefonoCoarrendatario1TXT.Text = respuesta[0].telefonoCoarrendatario1TXTdb.ToString();
            celularCoarrendatario1TXT.Text = respuesta[0].celularCoarrendatario1TXTdb.ToString();
            emailCoarrendatario1TXT.Text = respuesta[0].emailCoarrendatario1TXTdb.ToString();
            direccionCoarrendatario1TXT.Text = respuesta[0].direccionCoarrendatario1TXTdb.ToString();

            nombreCoarrendatario2TXT.Text = respuesta[0].nombreCoarrendatario2TXTdb.ToString();
            idCoarrendatario2TXT.Text = respuesta[0].idCoarrendatario2TXTdb.ToString();
            tipoIdCoarrendatario2TXT.Text = respuesta[0].tipoIdCoarrendatario2TXTdb.ToString();
            telefonoCoarrendatario2TXT.Text = respuesta[0].telefonoCoarrendatario2TXTdb.ToString();
            celularCoarrendatario2TXT.Text = respuesta[0].celularCoarrendatario2TXTdb.ToString();
            emailCoarrendatario2TXT.Text = respuesta[0].emailCoarrendatario2TXTdb.ToString();
            direccionCoarrendatario2TXT.Text = respuesta[0].direccionCoarrendatario2TXTdb.ToString();

            nombreCoarrendatario3TXT.Text = respuesta[0].nombreCoarrendatario3TXTdb.ToString();
            tipoIdCoarrendatario3TXT.Text = respuesta[0].tipoIdCoarrendatario3TXTdb.ToString();
            idCoarrendatario3TXT.Text = respuesta[0].idCoarrendatario3TXTdb.ToString();
            telefonoCoarrendatario3TXT.Text = respuesta[0].telefonoCoarrendatario3TXTdb.ToString();
            celularCoarrendatario3TXT.Text = respuesta[0].celularCoarrendatario3TXTdb.ToString();
            emailCoarrendatario3TXT.Text = respuesta[0].emailCoarrendatario3TXTdb.ToString();
            direccionCoarrendatario3TXT.Text = respuesta[0].direccionCoarrendatario3TXTdb.ToString();

            nombreCoarrendatario4TXT.Text = respuesta[0].nombreCoarrendatario4TXTdb.ToString();
            tipoIdCoarrendatario4TXT.Text = respuesta[0].tipoIdCoarrendatario4TXTdb.ToString();
            idCoarrendatario4TXT.Text = respuesta[0].idCoarrendatario4TXTdb.ToString();
            telefonoCoarrendatario4TXT.Text = respuesta[0].telefonoCoarrendatario4TXTdb.ToString();
            celularCoarrendatario4TXT.Text = respuesta[0].celularCoarrendatario4TXTdb.ToString();
            emailCoarrendatario4TXT.Text = respuesta[0].emailCoarrendatario4TXTdb.ToString();
            direccionCoarrendatario4TXT.Text = respuesta[0].direccionCoarrendatario4TXTdb.ToString();
            modificacionTXT.Text = respuesta[0].modificacionTXTdb.ToString();

            if (respuesta[0].idCoarrendatario1TXTdb.ToString() != "") { checkCoarrendatario1.Checked = true; grupoCoarrendatario1.Enabled = true; }
            if (respuesta[0].idCoarrendatario2TXTdb.ToString() != "") { checkCoarrendatario2.Checked = true; grupoCoarrendatario2.Enabled = true; }
            if (respuesta[0].idCoarrendatario3TXTdb.ToString() != "") { checkCoarrendatario3.Checked = true; grupoCoarrendatario3.Enabled = true; }
            if (respuesta[0].idCoarrendatario4TXTdb.ToString() != "") { checkCoarrendatario4.Checked = true; grupoCoarrendatario4.Enabled = true; }
            if (respuesta[0].idCoarrendatario5TXTdb.ToString() != "") { checkCoarrendatario5.Checked = true; grupoCoarrendatario5.Enabled = true; }



        }

        public void traerInfoApi2(string Contrato)
        {

            string[] arregloTiposConsulta = new string[8];
            string apiURL = "https://portalhouses.com/administrador/ApiDocuments2/post.php?tipo=mostrar&contratoConsulta=";
            arregloTiposConsulta[0] =  apiURL + Contrato + "&tipoTerceroTXTdb=arrendatario";




            dynamic respuesta = dBApi.Get(arregloTiposConsulta[0]);



            string[] arrTercero = respuesta[0].arrendatarioTXTdb.ToString().Split(';');
            nombreArrendatarioTXT.Text = arrTercero[0];
            idArrendatarioTXT.Text = arrTercero[1];
            tipoIdArrendatarioTXT.Text = arrTercero[2];
            ciudadIdArrendatarioTXT.Text = arrTercero[3];
            telefonoArrendatarioTXT.Text = arrTercero[4];
            celularArrendatarioTXT.Text = arrTercero[5];
            emailArrendatarioTXT.Text = arrTercero[6];
            direccionArrendatarioTXT.Text = arrTercero[7];
            empresaArrendatarioTXT.Text = arrTercero[8];
            telEmpresaArrendatarioTXT.Text = arrTercero[9];
            direccionEmpresaArrendatarioTXT.Text = arrTercero[10];
            cargoEmpresaArrendatarioTXT.Text = arrTercero[11];
            cuentaArrendatarioTXT.Text = arrTercero[12];
            tipoCuentaArrendatarioTXT.Text = arrTercero[13];
            bancoArrendatarioTXT.Text = arrTercero[14];

            arrTercero = respuesta[0].propietarioTXTdb.ToString().Split(';');
            nombrePropietarioTXT.Text = arrTercero[0];
            idPropietarioTXT.Text = arrTercero[1];
            tipoIdPropietarioTXT.Text = arrTercero[2];
            ciudadIdPropietarioTXT.Text = arrTercero[3];
            telefonoPropietarioTXT.Text = arrTercero[4];
            celularPropietarioTXT.Text = arrTercero[5];
            emailPropietarioTXT.Text = arrTercero[6];
            direccionPropietarioTXT.Text = arrTercero[7];
            empresaPropietarioTXT.Text = arrTercero[8];
            telEmpresaPropietarioTXT.Text = arrTercero[9];
            direccionEmpresaPropietarioTXT.Text = arrTercero[10];
            cargoEmpresaPropietarioTXT.Text = arrTercero[11];
            cuentaPropietarioTXT.Text = arrTercero[12];
            tipoCuentaPropietarioTXT.Text = arrTercero[13];
            bancoPropietarioTXT.Text = arrTercero[14];

            arrTercero = respuesta[0].encargadoTXTdb.ToString().Split(';');
            nombreEncargadoTXT.Text = arrTercero[0];
            idEncargadoTXT.Text = arrTercero[1];
            tipoIdEncargadoTXT.Text = arrTercero[2];
            ciudadIdEncargadoTXT.Text = arrTercero[3];
            telefonoEncargadoTXT.Text = arrTercero[4];
            celularEncargadoTXT.Text = arrTercero[5];
            emailEncargadoTXT.Text = arrTercero[6];
            direccionEncargadoTXT.Text = arrTercero[7];
            empresaEncargadoTXT.Text = arrTercero[8];
            telEmpresaEncargadoTXT.Text = arrTercero[9];
            direccionEmpresaEncargadoTXT.Text = arrTercero[10];
            cargoEmpresaEncargadoTXT.Text = arrTercero[11];
            cuentaEncargadoTXT.Text = arrTercero[12];
            tipoCuentaEncargadoTXT.Text = arrTercero[13];
            bancoEncargadoTXT.Text = arrTercero[14];


            arrTercero = respuesta[0].coarrendatario1TXTdb.ToString().Split(';');
            nombreCoarrendatario1TXT.Text = arrTercero[0];
            idCoarrendatario1TXT.Text = arrTercero[1];
            tipoIdCoarrendatario1TXT.Text = arrTercero[2];
            ciudadIdCoarrendatario1TXT.Text = arrTercero[3];
            telefonoCoarrendatario1TXT.Text = arrTercero[4];
            celularCoarrendatario1TXT.Text = arrTercero[5];
            emailCoarrendatario1TXT.Text = arrTercero[6];
            direccionCoarrendatario1TXT.Text = arrTercero[7];
            empresaCoarrendatario1TXT.Text = arrTercero[8];
            telEmpresaCoarrendatario1TXT.Text = arrTercero[9];
            cargoEmpresaCoarrendatario1TXT.Text = arrTercero[10];
            direccionEmpresaCoarrendatario1TXT.Text = arrTercero[11];

            arrTercero = respuesta[0].coarrendatario2TXTdb.ToString().Split(';');
            nombreCoarrendatario2TXT.Text = arrTercero[0];
            idCoarrendatario2TXT.Text = arrTercero[1];
            tipoIdCoarrendatario2TXT.Text = arrTercero[2];
            ciudadIdCoarrendatario2TXT.Text = arrTercero[3];
            telefonoCoarrendatario2TXT.Text = arrTercero[4];
            celularCoarrendatario2TXT.Text = arrTercero[5];
            emailCoarrendatario2TXT.Text = arrTercero[6];
            direccionCoarrendatario2TXT.Text = arrTercero[7];
            empresaCoarrendatario2TXT.Text = arrTercero[8];
            telEmpresaCoarrendatario2TXT.Text = arrTercero[9];
            direccionEmpresaCoarrendatario2TXT.Text = arrTercero[10];
            cargoEmpresaCoarrendatario2TXT.Text = arrTercero[11];

            arrTercero = respuesta[0].coarrendatario3TXTdb.ToString().Split(';');
            nombreCoarrendatario3TXT.Text = arrTercero[0];
            idCoarrendatario3TXT.Text = arrTercero[1];
            tipoIdCoarrendatario3TXT.Text = arrTercero[2];
            ciudadIdCoarrendatario3TXT.Text = arrTercero[3];
            telefonoCoarrendatario3TXT.Text = arrTercero[4];
            celularCoarrendatario3TXT.Text = arrTercero[5];
            emailCoarrendatario3TXT.Text = arrTercero[6];
            direccionCoarrendatario3TXT.Text = arrTercero[7];
            empresaCoarrendatario3TXT.Text = arrTercero[8];
            telEmpresaCoarrendatario3TXT.Text = arrTercero[9];
            direccionEmpresaCoarrendatario3TXT.Text = arrTercero[10];
            cargoEmpresaCoarrendatario3TXT.Text = arrTercero[11];

            arrTercero = respuesta[0].coarrendatario4TXTdb.ToString().Split(';');
            nombreCoarrendatario4TXT.Text = arrTercero[0];
            idCoarrendatario4TXT.Text = arrTercero[1];
            tipoIdCoarrendatario4TXT.Text = arrTercero[2];
            ciudadIdCoarrendatario4TXT.Text = arrTercero[3];
            telefonoCoarrendatario4TXT.Text = arrTercero[4];
            celularCoarrendatario4TXT.Text = arrTercero[5];
            emailCoarrendatario4TXT.Text = arrTercero[6];
            direccionCoarrendatario4TXT.Text = arrTercero[7];
            empresaCoarrendatario4TXT.Text = arrTercero[8];
            telEmpresaCoarrendatario4TXT.Text = arrTercero[9];
            direccionEmpresaCoarrendatario4TXT.Text = arrTercero[10];
            cargoEmpresaCoarrendatario4TXT.Text = arrTercero[11];

            arrTercero = respuesta[0].coarrendatario5TXTdb.ToString().Split(';');
            nombreCoarrendatario5TXT.Text = arrTercero[0];
            idCoarrendatario5TXT.Text = arrTercero[1];
            tipoIdCoarrendatario5TXT.Text = arrTercero[2];
            ciudadIdCoarrendatario5TXT.Text = arrTercero[3];
            telefonoCoarrendatario5TXT.Text = arrTercero[4];
            celularCoarrendatario5TXT.Text = arrTercero[5];
            emailCoarrendatario5TXT.Text = arrTercero[6];
            direccionCoarrendatario5TXT.Text = arrTercero[7];
            empresaCoarrendatario5TXT.Text = arrTercero[8];
            telEmpresaCoarrendatario5TXT.Text = arrTercero[9];
            direccionEmpresaCoarrendatario5TXT.Text = arrTercero[10];
            cargoEmpresaCoarrendatario5TXT.Text = arrTercero[11];

            if (idCoarrendatario1TXT.Text != "") { checkCoarrendatario1.Checked = true; grupoCoarrendatario1.Enabled = true; }
            if (idCoarrendatario2TXT.Text != "") { checkCoarrendatario2.Checked = true; grupoCoarrendatario2.Enabled = true; }
            if (idCoarrendatario3TXT.Text != "") { checkCoarrendatario3.Checked = true; grupoCoarrendatario3.Enabled = true; }
            if (idCoarrendatario4TXT.Text != "") { checkCoarrendatario4.Checked = true; grupoCoarrendatario4.Enabled = true; }
            if (idCoarrendatario5TXT.Text != "") { checkCoarrendatario5.Checked = true; grupoCoarrendatario5.Enabled = true; }

        }

        public void verCarpeta()
        {

            var rutaCarpeta = @"\\servidor1\Fotos\FOTOS_FIRMA_DE_CONTRATOS\CTO_" + contratoTXT.Text;
            Process.Start("explorer.exe", rutaCarpeta);
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
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idArrendatarioTXT**", tipoIdArrendatarioTXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreArrendatarioTXT**", nombreArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoArrendatarioTXT**", telefonoArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularArrendatarioTXT**", celularArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailArrendatarioTXT**", emailArrendatarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionArrendatarioTXT**", direccionArrendatarioTXT.Text);

                        break;
                    case 1:
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idPropietarioTXT**", tipoIdPropietarioTXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombrePropietarioTXT**", nombrePropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoPropietarioTXT**", telefonoPropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularPropietarioTXT**", celularPropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailPropietarioTXT**", emailPropietarioTXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionPropietarioTXT**", direccionPropietarioTXT.Text);

                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**idPropietarioTXT**", tipoIdPropietarioTXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**nombrePropietarioTXT**", nombrePropietarioTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**telefonoPropietarioTXT**", telefonoPropietarioTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**celularPropietarioTXT**", celularPropietarioTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**emailPropietarioTXT**", emailPropietarioTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**direccionPropietarioTXT**", direccionPropietarioTXT.Text);

                       
                        break;
                    case 2:
                        if (checkCoarrendatario1.Checked) { checkCoarrendatarios("check", "1");} else { checkCoarrendatarios("uncheck", "1"); }
                            

                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario1TXT**", tipoIdCoarrendatario1TXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario1TXT**", nombreCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario1TXT**", telefonoCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario1TXT**", celularCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario1TXT**", direccionCoarrendatario1TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario1TXT**", emailCoarrendatario1TXT.Text);
                        break;
                    case 3:
                        if (checkCoarrendatario2.Checked) { checkCoarrendatarios("check", "2"); } else { checkCoarrendatarios("uncheck", "2"); }

                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario2TXT**", tipoIdCoarrendatario2TXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario2TXT**", nombreCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario2TXT**", telefonoCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario2TXT**", celularCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario2TXT**", direccionCoarrendatario2TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario2TXT**", emailCoarrendatario2TXT.Text);
                        break;
                    case 4:
                        if (checkCoarrendatario3.Checked) { checkCoarrendatarios("check", "3"); } else { checkCoarrendatarios("uncheck", "3"); }
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario3TXT**", tipoIdCoarrendatario3TXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario3TXT**", nombreCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario3TXT**", telefonoCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario3TXT**", celularCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario3TXT**", direccionCoarrendatario3TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario3TXT**", emailCoarrendatario3TXT.Text);

                        break;
                    case 5:
                        if (checkCoarrendatario4.Checked) { checkCoarrendatarios("check", "4"); } else { checkCoarrendatarios("uncheck", "4"); }
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**idCoarrendatario4TXT**", tipoIdCoarrendatario4TXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**nombreCoarrendatario4TXT**", nombreCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**telefonoCoarrendatario4TXT**", telefonoCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**celularCoarrendatario4TXT**", celularCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**direccionCoarrendatario4TXT**", direccionCoarrendatario4TXT.Text);
                        richTextBox1.Rtf = richTextBox1.Rtf.Replace("**emailCoarrendatario4TXT**", emailCoarrendatario4TXT.Text);

                        break;
                    case 6:
                        if (checkCoarrendatario4.Checked) { checkCoarrendatarios("check", "4"); } else { checkCoarrendatarios("uncheck", "4"); }
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**idEncargadoTXT**", tipoIdEncargadoTXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**nombreEncargadoTXT**", nombreEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**telefonoEncargadoTXT**", telefonoEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**celularEncargadoTXT**", celularEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**emailEncargadoTXT**", emailEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**direccionEncargadoTXT**", direccionEncargadoTXT.Text);

                        break;
                    case 7:
                        if (checkCoarrendatario5.Checked) { checkCoarrendatarios("check", "5"); } else { checkCoarrendatarios("uncheck", "5"); }
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**idEncargadoTXT**", tipoIdCoarrendatario5TXT.Text + ": " + terceroFormateado + " DE CALI");
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**nombreEncargadoTXT**", nombreEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**telefonoEncargadoTXT**", telefonoEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**celularEncargadoTXT**", celularEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**emailEncargadoTXT**", emailEncargadoTXT.Text);
                        richTextBox2.Rtf = richTextBox2.Rtf.Replace("**direccionEncargadoTXT**", direccionEncargadoTXT.Text);

                        break;

                }




            };

            //------------------------------------------

        }
    



        private void Form1_Load(object sender, EventArgs e)
        {
            int MyPropertyInteger;
            int.TryParse(this.MyProperty, out MyPropertyInteger);
            switch (MyProperty) { 
                case "mathias24":
                    usuarioTXT.Text = "Auxadministracion";
                break;
                case "2424":
                    usuarioTXT.Text = "Servicio C." ;
                break;
                case "9999":
                    usuarioTXT.Text = "Sistemas";
                break;
                case "eliza52":
                    usuarioTXT.Text = "Administracion";
                    break;




            }
            listarContratos();
            calcularFechasContrato();
        }
        public string MyProperty { get; set; }

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

            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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

        private  void cONTRATODEARRENDAMIENTODEUNBIENINMUEBLESOMETIDOACOPROPIEDADYDESTINADOAVIVIENDAURBANAPRUEBAToolStripMenuItem_Click(object sender, EventArgs e)
        {
             cargarFormato("formato1");
            informacionTXT.Text = "Formato actual: Inmueble sometido a copropiedad y destinado a vivienda urbana ";
            informacionTXT.ForeColor = Color.Green;
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
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
            btnDuplicar.Visible = true;
            btnGuardarTXT.Text = "Duplicar";
            GuardarDuplicarToolStripMenuItem.Text = "Duplicar contrato"; //---CAMBIAMOS NOMBRE EN EL MENU
            btnActualizar.Enabled = true;
            btnActualizarTXT.Enabled = true;
            GuardartoolStripMenuItem.Enabled = true;
            btnVerCarpeta.Enabled = true;
            btnVerCarpetaTXT.Enabled = true;

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
            generarVistaPreviaToolStripMenuItem.Enabled = false;
            informacionTXT.Text = "No se ha cargado formato de contrato";
            informacionTXT.ForeColor = Color.Red;
            richTextBox1.Text = "";
            richTextBox2.Text = "";
            btnVerCarpeta.Enabled = false;
            btnVerCarpetaTXT.Enabled = false;
            grupoCoarrendatario1.Enabled = false;
            grupoCoarrendatario2.Enabled = false;
            grupoCoarrendatario3.Enabled = false;
            grupoCoarrendatario4.Enabled = false;
            grupoCoarrendatario5.Enabled = false;
            checkCoarrendatario1.Checked = false;
            checkCoarrendatario2.Checked = false;
            checkCoarrendatario3.Checked = false;
            checkCoarrendatario4.Checked = false;
            checkCoarrendatario5.Checked = false;
            clausulaTXT.Text = "3.5";
            servicioCTXT.Text = "10";
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
            if(GuardarDuplicarToolStripMenuItem.Text=="Duplicar contrato") { 
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea duplicar este contrato con el nuevo número: " + contratoTXT.Text + " ?", "Duplicar contrato", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                guardarContrato("insertar", "formato1");
                
                }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            }

            if (GuardarDuplicarToolStripMenuItem.Text == "Guardar contrato nuevo")
            {
                guardarContrato("insertar", "formato1");
            }
          



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

        private void button1_Click_4(object sender, EventArgs e)
        {

        }

        private void button1_Click_5(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_6(object sender, EventArgs e)
        {
         
        }

        private void button1_Click_7(object sender, EventArgs e)
        {
            if (canonTXT.Text == ""){ canonTXT.Text = "0";}
            if (administracionTXT.Text == "") { administracionTXT.Text = "0"; }
            vistaPrevia();



        }

        private void btnExportar_Click(object sender, EventArgs e)
        {
            generarWord();
            
        }

        private void btnDuplicar_Click(object sender, EventArgs e)
        {
            
            DialogResult dialogResult = MessageBox.Show("¿Está seguro que desea duplicar este contrato con el nuevo número: " + contratoTXT.Text + " ?" , "Duplicar contrato", MessageBoxButtons.YesNo);
            if (dialogResult == DialogResult.Yes)
            {
                guardarContrato("insertar", "formato1");
            }
            else if (dialogResult == DialogResult.No)
            {
                //do something else
            }
            listarContratos();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            guardarContrato("insertar", "formato1");
            listarContratos();
        }

        private void btnActualizar_Click(object sender, EventArgs e)
        {
            guardarContrato("actualizar", "formato1");
            listarContratos();
        }

        private  async void button1_Click_8(object sender, EventArgs e)
        {

            
            
          

        }

        private void inmuebleDestinadoAViviendaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cargarFormato("formato2");
            informacionTXT.Text = "Formato actual: Inmueble  destinado a Vivienda urbana: ";
            informacionTXT.ForeColor = Color.Green;
            destinoTXT.Text = "Vivienda";
            administracionTXT.Text = "0";
        }

        private void inmuebleSometidoACopropiedadYDestinadoALocalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cargarFormato("formato3");
            informacionTXT.Text = "Formato actual: Inmueble sometido a copropiedad y destinado a local: ";
            informacionTXT.ForeColor = Color.Green;
            destinoTXT.Text = "Local";
        }

        private void inmuebleDestinadoALocalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            cargarFormato("formato4");
            informacionTXT.Text = "Formato actual: Inmueble destinado local: ";
            informacionTXT.ForeColor = Color.Green;
            destinoTXT.Text = "Local";
            administracionTXT.Text = "0";
        }

        private void button1_Click_9(object sender, EventArgs e)
        {
            
            




        }

        private void servicioCTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void grupoCoarrendatario2_Enter(object sender, EventArgs e)
        {

        }

        private void checkCoarrendatario2_CheckedChanged(object sender, EventArgs e)
        {
            grupoCoarrendatario2.Enabled = true;
        }

        private void checkCoarrendatario1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkCoarrendatario3_CheckedChanged(object sender, EventArgs e)
        {
            grupoCoarrendatario3.Enabled = true;
        }

        private void checkCoarrendatario4_CheckedChanged(object sender, EventArgs e)
        {
            grupoCoarrendatario4.Enabled = true;
        }

        private void checkCoarrendatario5_CheckedChanged(object sender, EventArgs e)
        {
            grupoCoarrendatario5.Enabled = true;
        }

        private void label38_Click(object sender, EventArgs e)
        {

        }

        private void telefonoCoarrendatario4TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void checkCoarrendatario1_Click(object sender, EventArgs e)
        {
            
        }

        private void checkCoarrendatario1_CheckStateChanged(object sender, EventArgs e)
        {
            
        }

        private void checkCoarrendatario1_MouseClick(object sender, MouseEventArgs e)
        {
        
           

            
        }

        private void checkCoarrendatario1_CheckStateChanged_1(object sender, EventArgs e)
        {
            if (checkCoarrendatario1.Checked)
            {
                grupoCoarrendatario1.Enabled = true;
            }
            else { grupoCoarrendatario1.Enabled = false; }


        }


        private void checkCoarrendatario1_EnabledChanged(object sender, EventArgs e)
        {
            
        }

        private void checkCoarrendatario2_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkCoarrendatario2.Checked)
            {
                grupoCoarrendatario2.Enabled = true;
            }
            else { grupoCoarrendatario2.Enabled = false; }

        }

        private void checkCoarrendatario3_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkCoarrendatario3.Checked)
            {
                grupoCoarrendatario3.Enabled = true;
            }
            else { grupoCoarrendatario3.Enabled = false; }

        }

        private void checkCoarrendatario4_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkCoarrendatario4.Checked)
            {
                grupoCoarrendatario4.Enabled = true;
            }
            else { grupoCoarrendatario4.Enabled = false; }

        }

        private void checkCoarrendatario5_CheckedChanged_1(object sender, EventArgs e)
        {

        }

        private void checkCoarrendatario5_CheckStateChanged(object sender, EventArgs e)
        {
            if (checkCoarrendatario5.Checked)
            {
                grupoCoarrendatario5.Enabled = true;
            }
            else { grupoCoarrendatario5.Enabled = false; }

        }

        private void groupBox8_Enter(object sender, EventArgs e)
        {

        }

        private void fechaIniTXT_MouseUp(object sender, MouseEventArgs e)
        {
     
        }

        private void fechaIniTXT_MouseLeave(object sender, EventArgs e)
        {
          
        }

        private void button1_Click_10(object sender, EventArgs e)
        {
        
            
        }

        private void button1_Click_11(object sender, EventArgs e)
        {
           

        }

        private void salirToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click_12(object sender, EventArgs e)
        {
      
        }

        private void button1_Click_13(object sender, EventArgs e)
        {




        }

        private void button1_Click_14(object sender, EventArgs e)
        {
            verCarpeta();
        }

        private void telefonoCoarrendatario5TXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void telefonoCoarrendatario5TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void telefonoEncargadoTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void telefonoEncargadoTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void celularEncargadoTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void logoApp_Click(object sender, EventArgs e)
        {

        }

        private void telefonoArrendatarioTXT_KeyDown(object sender, KeyEventArgs e)
        {

        }

        private void idEncargadoTXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void idEncargadoTXT_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void idCoarrendatario5TXT_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar == '.') || (e.KeyChar == ','))
            {
                e.Handled = true;
            }
        }

        private void button1_Click_15(object sender, EventArgs e)
        {
            generarWord();
        }

        private void button2_Click_2(object sender, EventArgs e)
        {
            generarWord();
        }

        private void button1_Click_16(object sender, EventArgs e)
        {
            
        }

        private void button1_Click_17(object sender, EventArgs e)
        {
     

        }

        private void nombreArrendatarioTXT_TextChanged(object sender, EventArgs e)
        {

        }

        private void label89_Click(object sender, EventArgs e)
        {

        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox5_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox36_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox30_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox39_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox34_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox40_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox20_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox16_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void textBox24_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) && (e.KeyChar != '.'))
            {
                e.Handled = true;
            }
        }

        private void button2_Click_3(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click_18(object sender, EventArgs e)
        {
            traerInfoApi2("9999");
        }
    }
}

