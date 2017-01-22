using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DesempleoApp
{
    public partial class contenedor : Form
    {
        Usuario u;
        int perfil;
        int usr;
        public contenedor(int p, int u)
        {
            InitializeComponent();
            perfil = p;
            usr = u;
        }

        private void mantenimientoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }

        private void usuariosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void usuariosToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (perfil == 1)
            {
                pictureBox1.Visible = true;
                Application.DoEvents();
                usuarios f = new usuarios(perfil, ref pictureBox1);
                f.MdiParent = this;
                f.Show();
            }
            else
                MessageBox.Show("Opcion accesible solo por el administrador de la aplicación");
        }

        private void perfilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (perfil == 1)
            {
                pictureBox1.Visible = true;
                Application.DoEvents();
                perfiles f = new perfiles(ref pictureBox1);
                f.MdiParent = this;
                f.Show();
            }
            else
                MessageBox.Show("Opcion accesible solo por el administrador de la aplicación");
        }

        private void accesosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (perfil == 1)
            {
                pictureBox1.Visible = true;
                Application.DoEvents();
                accesos f = new accesos(ref pictureBox1);
                f.MdiParent = this;
                f.Show();
            }
            else
                MessageBox.Show("Opcion accesible solo por el administrador de la aplicación");
        }

        private void registrosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = true;
            Application.DoEvents();
            Form1 f = new Form1(perfil,usr, ref pictureBox1);
            f.MdiParent = this;

            f.Show();
        }

        private void firmasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (perfil == 1)
            {
                pictureBox1.Visible = true;
                Application.DoEvents();
                firma f = new firma(ref pictureBox1);
                f.MdiParent = this;
                f.Show();
            }
            else
                MessageBox.Show("Opcion accesible solo por el administrador de la aplicación");
        }

        private void cargosToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (perfil == 1)
            {
                pictureBox1.Visible = true;
                Application.DoEvents();
                cargo f = new cargo(ref pictureBox1);
                f.MdiParent = this;
                f.Show();
            }
            else
                MessageBox.Show("Opcion accesible solo por el administrador de la aplicación");
        }

        private void estadisticasToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (perfil == 1)
            {
                pictureBox1.Visible = true;
                Application.DoEvents();
                Estadisticas f = new Estadisticas(usr, ref pictureBox1);
                f.MdiParent = this;
                f.Show();
            }
            else
                MessageBox.Show("Opcion accesible solo por el administrador de la aplicación");
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void contenedor_Load(object sender, EventArgs e)
        {
            pictureBox1.Visible=false;
        }
    }
}
