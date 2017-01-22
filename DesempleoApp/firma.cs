using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace DesempleoApp
{
    public partial class firma : Form
    {
        Usuario u;
        Boolean nuevoregistro = false;
        int filanuevoregistro = -1;
        Boolean listo = false;
        int perfil;
        PictureBox p;
        public firma(ref PictureBox pic)
        {
            InitializeComponent();
            u = new Usuario(Application.StartupPath);
            p = pic;
        }

        private void firma_Load(object sender, EventArgs e)
        {
            u.lista_de_firmas(ref dataGridView1);
            this.Width = dataGridView1.Width + 300;
            listo = true;
            p.Visible = false;
        }

        private void dataGridView1_UserAddedRow(object sender, DataGridViewRowEventArgs e)
        {
            /*if (u.esta_autorizado(perfil, 6))*/ 
            nuevoregistro = true;
            /*else MessageBox.Show("Su perfil no permite esta operación");*/
        }

        private void dataGridView1_RowLeave(object sender, DataGridViewCellEventArgs e)
        {
            /*if ((u.esta_autorizado(perfil, 6)) && (nuevoregistro))*/
            if ((nuevoregistro))
            {
                filanuevoregistro = dataGridView1.NewRowIndex - 1;
            }
        }

        private void dataGridView1_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            /*if ((u.esta_autorizado(perfil, 6)) && (filanuevoregistro > 0))*/
            if ((filanuevoregistro >= 0))
            {
                try
                {
                    string[] valores = new string[2];
                    valores[0] = dataGridView1[1, filanuevoregistro].Value.ToString();
                    u.crear_firma(valores[0]);
                    dataGridView1[0, filanuevoregistro].Value = u.ultimo_id_firma().ToString();
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
            /*if ((u.esta_autorizado(perfil, 7))&&(!nuevoregistro && listo))*/
            if ((!nuevoregistro && listo))
            {
            try
                {
                    //int perf = u.obtener_perfil(dataGridView1[3, dataGridView1.CurrentRow.Index].Value.ToString());
                    u.actualizar_firma(Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString()),
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
                //if (u.esta_autorizado(perfil, 8))
                if (MessageBox.Show("El registro actual será eliminado. Está seguro?", "Eliminar registro", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    u.borrar_firma(Convert.ToInt32(dataGridView1[0, dataGridView1.CurrentRow.Index].Value.ToString()));
                }
                /*else
                    MessageBox.Show("Su perfil no permite esta operación");*/
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error eliminando registro");
            }
        }
    }
}
