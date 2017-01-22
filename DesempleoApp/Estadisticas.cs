using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DesempleoApp
{
    public partial class Estadisticas : Form
    {
        Usuario u;
        int usr;
        PictureBox p;
        public Estadisticas(int usr,ref PictureBox pic)
        {
            InitializeComponent();
            u = new Usuario(Application.StartupPath);
            this.usr = usr;
            p = pic;
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {

        }

        private void Estadisticas_Load(object sender, EventArgs e)
        {
            u.leer_datos(comboBox1, "usuarios", usr);
            comboBox1.Items.Add("Todos");
            p.Visible = false;

            dataGridView1.Width = this.Width - 50;
            dataGridView1.Height = this.Height - 300;
            label4.Top = this.Height - 80;
            textBox1.Top = this.Height - 80;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string sql = "";

            int idusr = u.id_usuario(comboBox1.Text);
            if (dateTimePicker1.Value.Date <= dateTimePicker2.Value.Date)
            {
                sql += " AND Desempleo.FEntrada >=#" + dateTimePicker1.Value.Month.ToString() + "/" + dateTimePicker1.Value.Day.ToString() + "/" + dateTimePicker1.Value.Year.ToString() + "#";
                sql += " AND Desempleo.FEntrada <=#" + dateTimePicker2.Value.Month.ToString() + "/" + dateTimePicker2.Value.Day.ToString() + "/" + dateTimePicker2.Value.Year.ToString() + "#";
            }
            else
                MessageBox.Show("La fecha inicial debe ser menor a la fecha final. Filtro de fecha obviado");
            if (idusr>0){
                sql += " AND usuarios.nombre='" + comboBox1.Text + "'";
                
            }
            
            dataGridView1.Rows.Clear();
            u.lista_de_registros(ref dataGridView1, sql);
            textBox1.Text = dataGridView1.Rows.Count.ToString();
        }

        private void Estadisticas_Resize(object sender, EventArgs e)
        {
            dataGridView1.Width = this.Width - 50;
            dataGridView1.Height = this.Height - 300;
            label4.Top = this.Height - 80;
            textBox1.Top = this.Height - 80;
        }
    }
}
