using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
//using Microsoft.Office.Core;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Data.SqlClient;
using System.Configuration;
using iTextSharp.text;
using iTextSharp.text.pdf;


namespace DesempleoApp
{
    
    public partial class Form1 : Form
    {
        private string[] texto = new string[12];
        private string[] textoabuscar = new string[12];
        private Word.Application wordApp;
        private Word.Document aDoc;
        private Excel.Application excelApp;
        private object missing = Type.Missing;//Missing.Value;
        private object filename;
        private object destFile;
        SaveFileDialog saveFileDialog1 = new SaveFileDialog();
        SaveFileDialog saveFileDialog2 = new SaveFileDialog();
        private bool edicion = false;
        private int id_edicion = 0;
        private string stringdeconexion = "";
        int perfil;
        int usr;
        string nombre;
        PictureBox p;
        Usuario u;

        public Form1(int p, int us, ref PictureBox pic)
        {
            InitializeComponent();

            saveFileDialog1.Filter = "Documento pdf|*.pdf";
            saveFileDialog1.Title = "Guardar Documento pdf Adveración";
            saveFileDialog2.Filter = "Documento excel|*.xls";
            saveFileDialog2.Title = "Exportar a Excel todos los registros";
            perfil = p;
            usr = us;
            u = new Usuario(Application.StartupPath);
            nombre = u.nombre_usuario(usr);
            this.p = pic;
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            u.leer_datos(ComboBox1, "Firma",usr);

            u.leer_datos(ComboBox2, "Cargo",usr);

            u.leer_datos(checkedListBox1, "usuarios", usr);

            u.lista_de_registros(ref dataGridView1);

            if (!(u.esta_autorizado(perfil,5))) button4.Visible=false;
            if (!(u.esta_autorizado(perfil, 4))) button3.Visible = false;
            if (!(u.esta_autorizado(perfil, 1))) button5.Visible = false;
            if (!(u.esta_autorizado(perfil, 2))) button1.Visible = false;
            if (!(u.esta_autorizado(perfil, 3))) button2.Visible = false;
            p.Visible = false;

            dataGridView1.Width = this.Width - 50;
            dataGridView1.Height = this.Height - 450;
            button3.Top =  this.Height - 100;
            button4.Top =  this.Height - 100;
        }

        private DataSet leer_datos_principales()
        {
            //leer los datos desde el archivo excel
            string strTabla = string.Empty;
            DataSet oDs;

            //Nombre del rango tal cual lo definimos en el Excel
            strTabla = "Hoja1";
            using (
                OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(
                        "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Path.GetDirectoryName(Application.ExecutablePath) + "\\Entrada.xls; Extended Properties=Excel 8.0;")
                )
            {
                using (OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT * FROM [" + strTabla+"$]", oConn))
                {
                    oConn.Open();
                    oDs = new DataSet();
                    oCmd.Fill(oDs);
                    oConn.Close();
                }

            }
            return oDs;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (TextBox1.Text.Length + TextBox2.Text.Length > 0)
            {
                texto[0] = TextBox1.Text;
                texto[1] = TextBox2.Text;
                texto[2] = TextBox4.Text;
                texto[3] = DateTime.Parse(DateTimePicker2.Text).ToShortDateString();
                texto[4] = ComboBox1.Text;
                texto[5] = ComboBox2.Text;
                texto[6] = u.lista_de_rf(Convert.ToInt32(dataGridView1.CurrentRow.Cells[0].Value));

                textoabuscar[0] = "<NOMBRE>";
                textoabuscar[1] = "<PASAPORTE>";
                textoabuscar[2] = "<SA>";
                textoabuscar[3] = "<FSAL>";
                textoabuscar[4] = "<FIRMA>";
                textoabuscar[5] = "<CARGO>";
                textoabuscar[6] = "<RF>";
                
                button4.Visible = false;
                button1.Visible = false;
                button2.Visible = false;
                button3.Visible = false;
                string path = Directory.GetCurrentDirectory();
                string fileName = "Adveracion documentos.doc";
                string filedest = TextBox1.Text + " " + TextBox2.Text + ".pdf";

                saveFileDialog1.FileName = filedest;

                // If the file name is not an empty string open it for saving.
                if ((saveFileDialog1.ShowDialog() == DialogResult.OK) && (saveFileDialog1.FileName != ""))
                {
                    destFile = saveFileDialog1.FileName;
                    string hoy = DateTime.Today.ToLongDateString().Split(',')[1];
                    generarPDF((string)destFile, texto[0], texto[1], texto[2], texto[3], texto[4], texto[5], texto[6], hoy);

                    string sourceFile = System.IO.Path.Combine(path, fileName);
                    /*destFile = System.IO.Path.Combine(path, filedest);*/

                }
                button1.Visible = true;
                button2.Visible = true;
                button3.Visible = true;
                button4.Visible = true;
            }
            else
            {
                MessageBox.Show("Debe indicar al menos nombre y pasaporte");
            }
        }

        public void SaveDocument()
        {
            try
            {
                aDoc.Save();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error durante el proceso. Descripcion: " + ex.Message, "Error Interno", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CloseDocument()
        {
            object saveChanges = Word.WdSaveOptions.wdSaveChanges;
            wordApp.Quit(ref saveChanges, ref missing, ref missing);
        }

        public void FindAndReplace(object findText, object replaceText)
        {
            this.findAndReplace(wordApp, findText, replaceText);
        }

        public void findAndReplace(Word.Application wordApp, object findText, object replaceText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;

            wordApp.Selection.Find.Execute(ref findText, ref matchCase, ref matchWholeWord,
            ref matchWildCards, ref matchSoundsLike, ref matchAllWordForms,
            ref forward, ref wrap, ref format, ref replaceText, ref replace,
            ref matchKashida, ref matchDiacritics, ref matchAlefHamza, ref matchControl);
        }

        private void generarPDF(string nombreArch,string NOMBRE,
                string PASAPORTE,
                string SA,
                string FSAL,
                string FIRMA,
                string CARGO,
                string RF,
                string FECHAHOY)
        {
            // Creamos el documento con el tamaño de página tradicional
            Document doc = new Document(PageSize.LETTER);
            // Indicamos donde vamos a guardar el documento
            PdfWriter writer = PdfWriter.GetInstance(doc,
                                        new FileStream(nombreArch, FileMode.Create));

            // Le colocamos el título y el autor
            // **Nota: Esto no será visible en el documento
            doc.AddTitle("Adveración");
            doc.AddCreator("Consulado de España");

            // Abrimos el archivo
            doc.Open();

            // Creamos la imagen y le ajustamos el tamaño
            iTextSharp.text.Image imagen = iTextSharp.text.Image.GetInstance(Application.StartupPath + "\\Escudo.png");
            imagen.BorderWidth = 0;
            imagen.Alignment = Element.ALIGN_LEFT;
            /*float percentage = 0.0f;
            percentage = 150 / imagen.Width;*/
            imagen.SetAbsolutePosition(60, doc.PageSize.Height - 90);
            imagen.ScalePercent(50);

            // Insertamos la imagen en el documento
            doc.Add(imagen);

            // Escribimos el encabezamiento en el documento
            /*doc.Add(new Paragraph("Mi primer documento PDF"));
            doc.Add(Chunk.NEWLINE);*/

            iTextSharp.text.Font _standardFont = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontBold = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.HELVETICA, 10, iTextSharp.text.Font.BOLD, BaseColor.BLACK);

            // Creamos una tabla que contendrá el nombre, apellido y país 
            // de nuestros visitante.
            PdfPTable tblPrueba = new PdfPTable(4);
            tblPrueba.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 4, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            float[] medidaCeldas = { 0.8f, 1f, 1.7f, 0.9f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba.SetWidths(medidaCeldas);

            // Configuramos el título de las columnas de la tabla
            PdfPCell clNombre = new PdfPCell(new Phrase("", _standardFont));

            clNombre.BorderWidth = 0;
            clNombre.BorderWidthRight = 0.9f;

            PdfPCell clApellido = new PdfPCell(new Phrase("                                           Consejería de Empleo y Seguridad Social de                       la Embajada de España", _standardFontBold));

            clApellido.BorderWidth = 0;
            clApellido.PaddingLeft = 5f;
            clApellido.BorderWidthLeft = 0.9f;

            PdfPCell clPais = new PdfPCell(new Phrase("Avda. Principal Eugenio Mendoza con Primera Transversal, Edificio Banco Lara, Primer Piso, Urbanización La Castellana -  Caracas – Venezuela", _standardFont));
            //clPais.Padding = 1.5f;
            clPais.PaddingLeft = 5f;
            clPais.BorderWidth = 0;
            clPais.BorderWidthLeft = 0.75f;

            PdfPCell clTelf = new PdfPCell(new Phrase("                                                        Tlfs.:   319.42.30                                                                                                                                                     Fax.:   319.42.42", _standardFont));
            clTelf.PaddingLeft = 5f;
            clTelf.BorderWidth = 0;
            clTelf.BorderWidthLeft = 0.75f;

            // Añadimos las celdas a la tabla
            tblPrueba.AddCell(clNombre);
            tblPrueba.AddCell(clApellido);
            tblPrueba.AddCell(clPais);
            tblPrueba.AddCell(clTelf);

            // Finalmente, añadimos la tabla al documento PDF y cerramos el documento
            doc.Add(tblPrueba);

            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;

            clNombre = new PdfPCell(new Phrase(" ", _standardFont));

            clNombre.BorderWidth = 0;

            tblPrueba.AddCell(clNombre);

            doc.Add(tblPrueba);

            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;

            clNombre = new PdfPCell(new Phrase(" ", _standardFont));

            clNombre.BorderWidth = 0;

            tblPrueba.AddCell(clNombre);

            doc.Add(tblPrueba);


            iTextSharp.text.Font _standardFontbox1 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 8, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontbox2 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 14, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontbox3 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.BOLD, BaseColor.BLACK);
            iTextSharp.text.Font _standardFontbox4 = new iTextSharp.text.Font(iTextSharp.text.Font.FontFamily.TIMES_ROMAN, 12, iTextSharp.text.Font.NORMAL, BaseColor.BLACK);
            PdfPTable tblPrueba1 = new PdfPTable(3);
            tblPrueba1.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 2, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            float[] medidaCeldas1 = { 0.5f, 0.4f, 0.2f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba1.SetWidths(medidaCeldas1);

            // Configuramos el título de las columnas de la tabla
            PdfPCell uno = new PdfPCell(new Phrase("", _standardFont));

            uno.BorderWidth = 0;
            //uno.BorderWidthRight = 0.9f;

            PdfPCell dos = new PdfPCell(new Phrase("           Consejería de Empleo y Seguridad Social ", _standardFontbox1));

            dos.BorderWidth = 0;
            dos.BorderWidthLeft = 0.2f;
            dos.BorderWidthRight = 0.2f;
            dos.BorderWidthTop = 0.9f;
            //dos.BorderWidthLeft = 0.9f;

            PdfPCell tres = new PdfPCell(new Phrase("", _standardFont));

            tres.BorderWidth = 0;

            // Añadimos las celdas a la tabla
            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(dos);
            tblPrueba1.AddCell(tres);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(3);
            tblPrueba1.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 2, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            medidaCeldas1 = new float[] { 0.5f, 0.4f, 0.2f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba1.SetWidths(medidaCeldas1);

            // Configuramos el título de las columnas de la tabla
            uno = new PdfPCell(new Phrase("  ", _standardFont));

            uno.BorderWidth = 0;
            //uno.BorderWidthRight = 0.9f;

            dos = new PdfPCell(new Phrase("             de la Embajada de España - Caracas", _standardFontbox1));

            dos.BorderWidth = 0;
            dos.BorderWidthLeft = 0.2f;
            dos.BorderWidthRight = 0.2f;
            //dos.BorderWidthLeft = 0.9f;

            // Añadimos las celdas a la tabla
            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(dos);
            tblPrueba1.AddCell(tres);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(3);
            tblPrueba1.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 2, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            medidaCeldas1 = new float[] { 0.5f, 0.4f, 0.2f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba1.SetWidths(medidaCeldas1);

            // Configuramos el título de las columnas de la tabla
            uno = new PdfPCell(new Phrase("  ", _standardFont));

            uno.BorderWidth = 0;
            //uno.BorderWidthRight = 0.9f;

            dos = new PdfPCell(new Phrase("  ", _standardFontbox2));

            dos.BorderWidth = 0;
            dos.BorderWidthLeft = 0.2f;
            dos.BorderWidthRight = 0.2f;
            //dos.BorderWidthLeft = 0.9f;

            // Añadimos las celdas a la tabla
            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(dos);
            tblPrueba1.AddCell(tres);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(3);
            tblPrueba1.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 2, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            medidaCeldas1 = new float[] { 0.5f, 0.4f, 0.2f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba1.SetWidths(medidaCeldas1);

            // Configuramos el título de las columnas de la tabla
            uno = new PdfPCell(new Phrase("", _standardFont));

            uno.BorderWidth = 0;
            //uno.BorderWidthRight = 0.9f;

            dos = new PdfPCell(new Phrase("                    SALIDA", _standardFontbox2));

            dos.BorderWidth = 0;
            dos.BorderWidthLeft = 0.2f;
            dos.BorderWidthRight = 0.2f;
            //dos.BorderWidthLeft = 0.9f;


            // Añadimos las celdas a la tabla
            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(dos);
            tblPrueba1.AddCell(tres);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(3);
            tblPrueba1.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 2, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            medidaCeldas1 = new float[] { 0.5f, 0.4f, 0.2f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba1.SetWidths(medidaCeldas1);

            // Configuramos el título de las columnas de la tabla
            uno = new PdfPCell(new Phrase("  ", _standardFont));

            uno.BorderWidth = 0;
            //uno.BorderWidthRight = 0.9f;

            dos = new PdfPCell(new Phrase("  ", _standardFontbox2));

            dos.BorderWidth = 0;
            dos.BorderWidthLeft = 0.2f;
            dos.BorderWidthRight = 0.2f;
            //dos.BorderWidthLeft = 0.9f;

            // Añadimos las celdas a la tabla
            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(dos);
            tblPrueba1.AddCell(tres);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(3);
            tblPrueba1.WidthPercentage = 100;

            // CREO UN ARREGLO QUE CONTIENE LAS MEDIDAS DE CADA UNA DE LAS COLUMNAS
            // EN MI CASO SON 2, (TB PUEDES PASAR EL ARREGLO DIRECTAMENTE)
            medidaCeldas1 = new float[] { 0.5f, 0.4f, 0.2f };

            // ASIGNAS LAS MEDIDAS A LA TABLA (ANCHO)
            tblPrueba1.SetWidths(medidaCeldas1);

            // Configuramos el título de las columnas de la tabla
            uno = new PdfPCell(new Phrase("", _standardFont));

            uno.BorderWidth = 0;
            //uno.BorderWidthRight = 0.9f;

            dos = new PdfPCell(new Phrase("Nº: "+SA+"               Fecha: "+FSAL, _standardFontbox3));

            dos.BorderWidth = 0;
            dos.BorderWidthLeft = 0.2f;
            dos.BorderWidthRight = 0.2f;
            dos.BorderWidthBottom = 0.2f;


            // Añadimos las celdas a la tabla
            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(dos);
            tblPrueba1.AddCell(tres);

            doc.Add(tblPrueba1);



            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;

            clNombre = new PdfPCell(new Phrase(" ", _standardFont));

            clNombre.BorderWidth = 0;

            tblPrueba.AddCell(clNombre);

            doc.Add(tblPrueba);

            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;

            clNombre = new PdfPCell(new Phrase(" ", _standardFont));

            clNombre.BorderWidth = 0;

            tblPrueba.AddCell(clNombre);

            doc.Add(tblPrueba);


            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            medidaCeldas1 = new float[] { 0.15f, 1f };

            tblPrueba1.SetWidths(medidaCeldas1);

            uno = new PdfPCell(new Phrase(" ", _standardFont));

            uno.BorderWidth = 0;

            clNombre = new PdfPCell(new Phrase("  ADVERACIÓN DE DOCUMENTOS RELATIVOS A ACTIVIDAD LABORAL", _standardFontbox3));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);


            clNombre = new PdfPCell(new Phrase("       A EFECTOS DEL CERTIFICADO DE EMIGRANTE RETORNADO", _standardFontbox3));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);


            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);


            clNombre = new PdfPCell(new Phrase("Rfª: "+RF, _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);


            clNombre = new PdfPCell(new Phrase("SOLICITANTE:  "+NOMBRE, _standardFontbox3));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);


            clNombre = new PdfPCell(new Phrase("NUMERO DE INSCRIPCIÓN CONSULAR:    "+PASAPORTE, _standardFontbox3));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);


            clNombre = new PdfPCell(new Phrase("    A los efectos de expedición del certificado de emigrante retornado, esta Consejería", _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);


            clNombre = new PdfPCell(new Phrase("de Empleo y Seguridad Social ADVERA la documentación anexa, presentada por el", _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);

            clNombre = new PdfPCell(new Phrase("solicitante, que  acredita haber desarrollado  en el exterior una actividad laboral por ", _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);

            clNombre = new PdfPCell(new Phrase("cuenta propia o cuenta ajena.", _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);

            clNombre = new PdfPCell(new Phrase("                                  En Caracas, a "+FECHAHOY, _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);

            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);

            clNombre = new PdfPCell(new Phrase("                                             "+CARGO, _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);

            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);

            clNombre = new PdfPCell(new Phrase("                                          "+FIRMA, _standardFontbox4));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);


            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);
            tblPrueba = new PdfPTable(1);
            tblPrueba.WidthPercentage = 100;
            clNombre = new PdfPCell(new Phrase(" ", _standardFont));
            clNombre.BorderWidth = 0;
            tblPrueba.AddCell(clNombre);
            doc.Add(tblPrueba);

            tblPrueba1 = new PdfPTable(2);
            tblPrueba1.WidthPercentage = 100;

            tblPrueba1.SetWidths(medidaCeldas1);

            clNombre = new PdfPCell(new Phrase("Correo electrónico:  venezuela@meyss.es", _standardFontBold));

            clNombre.BorderWidth = 0;

            tblPrueba1.AddCell(uno);
            tblPrueba1.AddCell(clNombre);

            doc.Add(tblPrueba1);

            doc.Close();
            writer.Close();

            MessageBox.Show("¡PDF creado!");
        }

        private void limpiaRF()
        {
            for (int i = 0; i < checkedListBox1.Items.Count; i++)
            {
                checkedListBox1.SetItemChecked(i, false);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            foreach (Control c in this.Controls)
            {
                if (c.GetType().Name.Equals("ComboBox"))
                    c.Text = "";
                if (c.GetType().Name.Equals("TextBox"))
                    c.Text = "";
                DateTimePicker1.Text = DateTime.Now.ToString();
                DateTimePicker2.Text = DateTime.Now.ToString();
            }
            limpiaRF();
            edicion = false;
            TextBox2.Enabled = true;
            u.marcarRF(ref checkedListBox1, nombre);
            if (!(u.esta_autorizado(perfil,2))) button1.Visible=true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!(u.esta_autorizado(perfil, 2))) button1.Visible = false;
            String sql="";

            string fecha= DateTime.Parse(DateTimePicker1.Text).ToShortDateString();
            string fecha2=DateTime.Parse(DateTimePicker2.Text).ToShortDateString();

            if ((edicion)&&(u.esta_autorizado(perfil,2)))
            {
                sql = "UPDATE Desempleo SET Nombre='" + TextBox1.Text + "',";
                sql += "Pasaporte='" + TextBox2.Text + "',Entrada='"+TextBox3.Text+"'";
                sql += ",FEntrada='"+fecha+"',";
                sql += "Salida='" + TextBox4.Text + "',";
                sql += "FSalida='"+fecha2+"',";
                sql += "Firma='" +  ComboBox1.Text + "',";
                sql += "Cargo='" + ComboBox2.Text + "',";
                sql += "RF='";
                sql += "',Observaciones='" + TextBox6.Text + "',";
                sql += "actualizar=" + usr.ToString();
                sql += " where id=" + id_edicion.ToString();
            }
            else
            {
                if (u.esta_autorizado(perfil, 1))
                {
                    sql = "INSERT INTO Desempleo (Nombre,";
                    sql += "Pasaporte,Entrada,FEntrada,";
                    sql += "Salida,FSalida,Firma,";
                    sql += "Cargo,RF,Observaciones,estado,agregar";
                    sql += ") VALUES ('";
                    sql += TextBox1.Text + "','" + TextBox2.Text + "','" + TextBox3.Text + "','"+fecha+"','";
                    sql += TextBox4.Text + "','"+fecha2+"','";
                    sql += ComboBox1.Text + "','" + ComboBox2.Text + "','";
                    sql += "','" + TextBox6.Text + "',1," + usr.ToString() + ")";
                }
            }

            if (u.existe_registro(TextBox2.Text) && (!(edicion)))
            {
                MessageBox.Show("Ya existe un registro con el pasaporte " + TextBox2.Text);
            }
            else {
                if (sql.Length > 5)
                {
                    OleDbConnection conn;
                    conn = new OleDbConnection(u.stringdeconexion);
                    conn.Open();

                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }

                    //conn.Close();
                    //Guardamos RF
                    //int idusuario=-1;
                    if (edicion)
                        try
                        {
                            sql = "delete from RF where iddesempleo=" + id_edicion.ToString();
                            using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                        catch (Exception ex) { }
                    if(checkedListBox1.CheckedItems.Count != 0)
                    {
                        //conn.Open();
                        for(int x = 0; x <= checkedListBox1.CheckedItems.Count - 1 ; x++)
                        {
                            //idusuario = u.id_usuario(checkedListBox1.CheckedItems[x].ToString());
                            sql = "INSERT INTO RF (iddesempleo,idusuario) values";
                            sql += "(" + u.ultimo_id_Desempleo().ToString() + "," + u.id_usuario(checkedListBox1.CheckedItems[x].ToString()) + ")";
                            using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                            {
                                cmd.ExecuteNonQuery();
                            }
                        }
                    }
                    conn.Close();
                    TextBox2.Enabled = false;
                    dataGridView1.Rows.Clear();
                    u.lista_de_registros(ref dataGridView1);
                }
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((dataGridView1.RowCount > 0) && (dataGridView1.SelectedRows.Count>0))
            {

            id_edicion = int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());
            if (MessageBox.Show("El registro que se muestra será eliminado. Está seguro?", "Eliminar registro", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                string sql = "UPDATE Desempleo SET estado=0, eliminar="+usr.ToString()+" WHERE id=" + id_edicion.ToString();

                OleDbConnection conn;
                conn = new OleDbConnection(u.stringdeconexion);
                conn.Open();

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                conn.Close();
                dataGridView1.Rows.Clear();
                u.lista_de_registros(ref dataGridView1);
            }
            }
            else
            {
                MessageBox.Show("No hay registros o no ha seleccionado una fila");
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {


            string path = Directory.GetCurrentDirectory();
            string fileName = "Desempleo.xls";

            if ((saveFileDialog2.ShowDialog() == DialogResult.OK) && (saveFileDialog2.FileName != ""))
            {
                //Copiar el archivo orginal vacio
                destFile = saveFileDialog2.FileName;
                string sourceFile = System.IO.Path.Combine(path, fileName);
                /*destFile = System.IO.Path.Combine(path, filedest);*/
                System.IO.File.Copy(sourceFile, (string)destFile, true);

                //excelApp = new Excel.Application();

                Excel.ApplicationClass excelApp = new Excel.ApplicationClass();
                Excel.Workbook excelBook = excelApp.Workbooks.Open((string)destFile, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing, missing);
                Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Sheets.get_Item(1);

                if (File.Exists((string)destFile))
                {
                    //comenzamos a guardar los registros desde la segunda linea
                    //DataTable dt = dataGridView1.    desempleoDataSet.Tables[0];

                    
                    string nombre = "";
                    string pasaporte = "";
                    string entrada = "";
                    string fentrada = "";
                    string salida = "";
                    string fsalida = "";
                    string firma = "";
                    string cargo = "";
                    string rf = "";
                    string observaciones = "";


                    int fila = 2;
                    int qty = dataGridView1.Rows.Count-1;
                    /*foreach (DataRow row in dt.Rows)*/
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        if (qty > 0)
                        {
                            nombre = row.Cells[1].Value.ToString();
                            pasaporte = row.Cells[2].Value.ToString();
                            entrada = row.Cells[3].Value.ToString();
                            fentrada = row.Cells[4].Value.ToString();
                            salida = row.Cells[5].Value.ToString();
                            fsalida = row.Cells[6].Value.ToString();
                            firma = row.Cells[7].Value.ToString();
                            cargo = row.Cells[8].Value.ToString();
                            rf = row.Cells[9].Value.ToString();
                            observaciones = row.Cells[10].Value.ToString();

                            excelSheet.Cells[fila, 1] = nombre;
                            excelSheet.Cells[fila, 2] = pasaporte;
                            excelSheet.Cells[fila, 3] = entrada;
                            excelSheet.Cells[fila, 4] = fentrada;
                            excelSheet.Cells[fila, 5] = salida;
                            excelSheet.Cells[fila, 6] = fsalida;
                            excelSheet.Cells[fila, 7] = firma;
                            excelSheet.Cells[fila, 8] = cargo;
                            excelSheet.Cells[fila, 9] = rf;
                            excelSheet.Cells[fila, 10] = observaciones;
                            fila++;
                        }
                        qty--;
                    }

                }
                excelSheet.SaveAs((string)destFile, missing, missing, missing, missing, missing, missing, missing, missing, missing);

                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelSheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

                excelBook = null;
                excelSheet = null;
                excelApp = null;
                System.GC.Collect();
                MessageBox.Show("Excel generado");
            }
            else
                MessageBox.Show("Excel no generado. Faltan datos");
        }

        private void dataGridView1_SelectionChanged_1(object sender, EventArgs e)
        {
            if (!(dataGridView1.CurrentRow == null))
            {
                try
                {
                    id_edicion = int.Parse(dataGridView1.CurrentRow.Cells[0].Value.ToString());

                    TextBox1.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                    TextBox2.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                    TextBox3.Text = dataGridView1.CurrentRow.Cells[3].Value.ToString();
                    DateTimePicker1.Value = DateTime.Parse(dataGridView1.CurrentRow.Cells[4].Value.ToString());
                    TextBox4.Text = dataGridView1.CurrentRow.Cells[5].Value.ToString();
                    DateTimePicker2.Value = DateTime.Parse(dataGridView1.CurrentRow.Cells[6].Value.ToString());
                    ComboBox1.Text = dataGridView1.CurrentRow.Cells[7].Value.ToString();
                    ComboBox2.Text = dataGridView1.CurrentRow.Cells[8].Value.ToString();
                    limpiaRF();
                    u.marcarRF(ref checkedListBox1, dataGridView1.CurrentRow.Cells[9].Value.ToString());
                    TextBox6.Text = dataGridView1.CurrentRow.Cells[10].Value.ToString();

                    edicion = true;
                    TextBox2.Enabled = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ocurrio un error inesperado: " + ex.Message);
                }
            }
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            dataGridView1.Width = this.Width - 50;
            dataGridView1.Height = this.Height - 450;
            button3.Top = this.Height - 100;
            button4.Top = this.Height - 100;
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
    }
}
