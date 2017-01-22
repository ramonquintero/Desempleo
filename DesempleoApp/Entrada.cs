using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DesempleoApp
{
    public partial class Entrada : Form
    {
        Usuario u;
        int perfil;
        int usr;
        public Entrada()
        {
            InitializeComponent();
            u = new Usuario(Application.StartupPath);
            u.crear_tablas();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (u.tiene_acceso(textBox1.Text, textBox2.Text, out perfil,out usr))
            {
                this.Hide();
                contenedor f = new contenedor(perfil,usr);
                f.ShowDialog();
                this.Close();
            }
            else
            {
                MessageBox.Show("Usuario o clave inválidos.");
            }
        }
    }
}
