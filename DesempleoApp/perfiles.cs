using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DesempleoApp
{
    public partial class perfiles : Form
    {
        Usuario u;
        Boolean nuevoregistro = false;
        int filanuevoregistro = -1;
        Boolean listo = false;
        PictureBox p;
        public perfiles(ref PictureBox pic)
        {
            InitializeComponent();
            u = new Usuario(Application.StartupPath);
            p = pic;
        }

        private void perfiles_Load(object sender, EventArgs e)
        {
            u.lista_de_perfiles(ref dataGridView1);
            listo = true;
            p.Visible = false;
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            nuevoregistro = true;
        }

        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            if (nuevoregistro)
            {
                filanuevoregistro = dataGridView1.NewRowIndex - 1;
            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (filanuevoregistro >= 0)
            {
                try
                {
                string[] valores = new string[3];
                valores[1] = dataGridView1[1, filanuevoregistro].Value.ToString();
                u.crear_perfil(valores[1]);
                dataGridView1[0, filanuevoregistro].Value = u.ultimo_id_perfil().ToString();
                nuevoregistro = false;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error. Debe llenar todos los datos del nuevo registro");
                }
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (!nuevoregistro && listo)
            {

                try
                {
                u.actualizar_perfil(Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString()),
                    dataGridView1[1, dataGridView1.CurrentRow.Index].Value.ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error actualizando registro");
                }
            }
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            try
            {
                if (MessageBox.Show("El registro actual será eliminado. Está seguro?", "Eliminar registro", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    u.borrar_perfil(Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString()));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error eliminando registro");
            }
        }
    }
}
