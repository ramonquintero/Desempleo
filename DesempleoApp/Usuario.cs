using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.IO;
//using System.Linq;
using System.Text;
//using System.Threading.Tasks;
using System.Windows.Forms;

namespace DesempleoApp
{
    class Usuario
    {
        public string stringdeconexion = "";
 
        private object missing = Type.Missing;//Missing.Value;

        public Usuario(string path)
        {
            stringdeconexion = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + "\\Desempleo.mdb";
        }

        public void crear_tablas(){
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            //perfiles
            /*try
            {
                sql = "create table perfiles ( ";
                sql += "id counter primary key,";
                sql += "nombre text(50))";

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into perfiles(nombre) values(";
                sql += "'administrador')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }

            //aplicacion
            
            try
            {
                sql = "create table aplicacion ( ";
                sql += "id counter primary key,";
                sql += "modulo text(50))";

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Crear registro')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Modificar registro')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Eliminar registro')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Generar excel')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Generar word')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Crear usuario')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Modificar usuario')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into aplicacion(modulo) values(";
                sql += "'Eliminar usuario')";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }

            //acceso

            try
            {
                sql = "create table acceso ( ";
                sql += "id counter primary key,";
                sql += "idperfil integer,";
                sql += "idaplicacion integer)";

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,1)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,2)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,3)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,4)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,5)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,6)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,7)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into acceso(idperfil,idaplicacion) values(";
                sql += "1,8)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }

            //usuarios
            try
            {
                sql = "create table usuarios ( ";
                sql +="id counter primary key,";
                sql +="login text(50),";
                sql +="passw text(50),";
                sql +="perfil integer)";
            
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "insert into usuarios(login,passw,perfil) values(";
                sql += "'admin','admin',1)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
            */
            //Agregar info de auditoria por registro
            try
            {
                //sql = "alter table Desempleo  ";
                //sql += "ADD COLUMN estado integer";
                

                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "update Desempleo  ";
                //sql += "SET estado=1";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}

                //sql = "alter table Desempleo  ";
                //sql += "ADD COLUMN agregar integer;";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "update Desempleo  ";
                //sql += "SET agregar=-1";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "alter table Desempleo  ";
                //sql += "ADD COLUMN actualizar integer;";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "update Desempleo  ";
                //sql += "SET actualizar=-1";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "alter table Desempleo  ";
                //sql += "ADD COLUMN eliminar integer";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "update Desempleo  ";
                //sql += "SET eliminar=-1";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                //sql = "alter table Usuarios  ";
                //sql += "ADD COLUMN nombre Text(50)";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                
                //Firma               
                //sql = "create table firma ( ";
                //sql += "id counter primary key,";
                //sql += "nombre text(50))";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                ////Cargo               
                //sql = "create table cargo ( ";
                //sql += "id counter primary key,";
                //sql += "nombre text(50))";
                //using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                //{
                //    cmd.ExecuteNonQuery();
                //}
                sql = "create table RF ( ";
                sql += "id counter primary key,";
                sql += "iddesempleo integer,";
                sql += "idusuario integer)";
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
                sql = "select Desempleo.id as desempleo, Desempleo.RF as usuario";
                sql +=" from Desempleo";
                //sql += " from Desempleo,usuarios where Desempleo.RF = usuarios.nombre";
                DataSet oDs;
                OleDbConnection oConn =
                        new System.Data.OleDb.OleDbConnection(stringdeconexion);

                OleDbDataAdapter oCmd = new OleDbDataAdapter(sql, oConn);
                oConn.Open();
                oDs = new DataSet();
                oCmd.Fill(oDs);
                oConn.Close();

                if (oDs.Tables[0].Rows.Count > 0)
                {
                    for (int i = 0; i < oDs.Tables[0].Rows.Count; i++)
                    {
                        sql = "insert into RF (iddesempleo,idusuario) values (";
                        sql += ((int)oDs.Tables[0].Rows[i]["desempleo"]).ToString() + ",";
                        sql += id_usuario(oDs.Tables[0].Rows[i]["usuario"].ToString()).ToString() + ")";
                        using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                        {
                            cmd.ExecuteNonQuery();
                        }
                    }
                }    
            }
            catch (Exception ex)
            {
            }

        }

        public Boolean tiene_acceso(string user, string clave, out int perfil,out int usuario)
        {
            Boolean res = false;
            perfil = -1;
            usuario = -1;
            try
            {
                DataSet oDs;
                
                OleDbConnection oConn =
                        new System.Data.OleDb.OleDbConnection(stringdeconexion);

                OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id,perfil FROM Usuarios where login='" + user + "' and passw='" + clave + "'", oConn);
                oConn.Open();
                oDs = new DataSet();
                oCmd.Fill(oDs);
                oConn.Close();

                if (oDs.Tables[0].Rows.Count > 0)
                {
                    perfil = (int)oDs.Tables[0].Rows[0]["perfil"];
                    usuario = (int)oDs.Tables[0].Rows[0]["id"];
                    res = true;
                }
            }
            catch (Exception ex) { }
            return res;
        }

        public Boolean esta_autorizado(int perfil, int modulo)
        {
            Boolean res = false;

            DataSet oDs;
            //perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT idperfil FROM acceso where idperfil=" + perfil.ToString() + " and idaplicacion=" + modulo.ToString(), oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["idperfil"];
                res = true;
            }

            return res;
        }

        public Boolean existe_perfil(string nombre)
        {
            Boolean res = false;
            int perfil;
            DataSet oDs;
            perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM perfiles where nombre='" + nombre + "'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["id"];
                res = true;
            }

            return res;
        }

        public int obtener_perfil(string nombre)
        {
            int perfil;
            DataSet oDs;
            perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM perfiles where nombre='" + nombre + "'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["id"];
            }

            return perfil;
        }

        public int obtener_modulo(string modulo)
        {
            int perfil;
            DataSet oDs;
            perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM aplicacion where modulo='" + modulo + "'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["id"];
            }

            return perfil;
        }

        public string nombre_usuario(int idusr)
        {
            string nombre="";
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT nombre FROM usuarios where id=" + idusr.ToString() , oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                nombre = oDs.Tables[0].Rows[0]["nombre"].ToString();
            }

            return nombre;
        }

        public int id_usuario(string usr)
        {
            int id=-1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM usuarios where nombre='" + usr+"'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public Boolean existe_usuario(string login)
        {
            Boolean res = false;
            int perfil;
            DataSet oDs;
            perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM usuarios where login='" + login + "'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["id"];
                res = true;
            }

            return res;
        }

        public Boolean existe_firma(string nombre)
        {
            Boolean res = false;
            int perfil;
            DataSet oDs;
            perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM firma where nombre='" + nombre + "'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["id"];
                res = true;
            }

            return res;
        }

        public Boolean existe_cargo(string nombre)
        {
            Boolean res = false;
            int perfil;
            DataSet oDs;
            perfil = -1;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM cargo where nombre='" + nombre + "'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();

            if (oDs.Tables[0].Rows.Count > 0)
            {
                perfil = (int)oDs.Tables[0].Rows[0]["id"];
                res = true;
            }

            return res;
        }

        public bool existe_registro(string pasaporte)
        {
            bool res = true;

            string sql = "select Pasaporte FROM Desempleo WHERE Pasaporte='" + pasaporte + "'";

            OleDbConnection conn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);
            conn.Open();
            OleDbDataReader reader;

            using (OleDbCommand cmd = new OleDbCommand(sql, conn))
            {
                reader = cmd.ExecuteReader();
                res = reader.HasRows;
            }
            conn.Close();

            return res;
        }

        public int ultimo_id_perfil()
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM perfiles order by id desc", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public void crear_perfil(string nombre)
        {
            if (!existe_perfil(nombre))
            {
                OleDbConnection conn;
                conn = new OleDbConnection(stringdeconexion);
                conn.Open();
                String sql;
                //perfiles
                try
                {
                    sql = "insert into perfiles(nombre) values('";
                    sql += nombre+"')";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                }
                conn.Close();
            }
        }

        public void actualizar_perfil(int id, string nombre)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "update perfiles set nombre='" + nombre + "' where id=" + id.ToString();
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void borrar_perfil(int id)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "delete from perfiles where id=" + id.ToString();

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public int ultimo_id_usuario()
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM usuarios order by id desc", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public void crear_usuario(string login,string passw,string nombre,int perfil)
        {
            if (!existe_usuario(login))
            {
                OleDbConnection conn;
                conn = new OleDbConnection(stringdeconexion);
                conn.Open();
                String sql;
                //perfiles
                try
                {
                    sql = "insert into usuarios(login,passw,nombre,perfil) values('";
                    sql += login + "','" + passw + "','" + nombre + "'," + perfil.ToString() + ")";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                }
                conn.Close();
            }
            else
            {
                MessageBox.Show("El login de ese usuario ya existe. No se agregará");
            }
        }

        public void actualizar_usuario(int id,string login, string passw, int perfil)
        {
                OleDbConnection conn;
                conn = new OleDbConnection(stringdeconexion);
                conn.Open();
                String sql;
                try
                {
                    sql = "update usuarios set login='"+login+"',passw='"+passw+"',perfil="+perfil+" where id="+id.ToString();
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                }
        }

        public void borrar_usuario(int id)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "delete from usuarios where id="+id.ToString();
                    
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public int ultimo_id_acceso()
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM acceso order by id desc", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public void crear_acceso(int perfil,int modulo)
        {
                OleDbConnection conn;
                conn = new OleDbConnection(stringdeconexion);
                conn.Open();
                String sql;
                try
                {
                    sql = "insert into acceso(idperfil,idaplicacion) values(";
                    sql += perfil.ToString() + ","+modulo.ToString()+")";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                }
                conn.Close();
        }

        public void actualizar_acceso(int id, int perfil, int modulo)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "update acceso set idperfil=" + perfil.ToString() + ", idaplicacion="+modulo.ToString()+" where id=" + id.ToString();
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void borrar_acceso(int id)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "delete from acceso where id=" + id.ToString();

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public int ultimo_id_firma()
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM firma order by id desc", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public void crear_firma(string nombre)
        {
            if (!existe_firma(nombre))
            {
                OleDbConnection conn;
                conn = new OleDbConnection(stringdeconexion);
                conn.Open();
                String sql;
                //perfiles
                try
                {
                    sql = "insert into firma(nombre) values('";
                    sql += nombre + "')";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                }
                conn.Close();
            }
        }

        public void actualizar_firma(int id, string nombre)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "update firma set nombre='" + nombre + "' where id=" + id.ToString();
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void borrar_firma(int id)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "delete from firma where id=" + id.ToString();

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public int ultimo_id_cargo()
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM cargo order by id desc", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public void crear_cargo(string nombre)
        {
            if (!existe_cargo(nombre))
            {
                OleDbConnection conn;
                conn = new OleDbConnection(stringdeconexion);
                conn.Open();
                String sql;
                try
                {
                    sql = "insert into cargo(nombre) values('";
                    sql += nombre + "')";
                    using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                    {
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                }
                conn.Close();
            }
        }

        public void actualizar_cargo(int id, string nombre)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "update cargo set nombre='" + nombre + "' where id=" + id.ToString();
                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void borrar_cargo(int id)
        {
            OleDbConnection conn;
            conn = new OleDbConnection(stringdeconexion);
            conn.Open();
            String sql;
            try
            {
                sql = "delete from cargo where id=" + id.ToString();

                using (OleDbCommand cmd = new OleDbCommand(sql, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
            catch (Exception ex)
            {
            }
        }

        public void agregar_perfiles(ref DataGridViewComboBoxColumn cmb)
        {
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);
            DataTable dt = new DataTable();
            //cmb.Items.Clear();
            string sql = "SELECT nombre ";
            sql += "FROM perfiles ";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("usuarios");
            da.Fill(dt);
            foreach (DataRow row in dt.Rows)
            {
                cmb.Items.Add(row["nombre"].ToString());
            }
        }

        public void agregar_modulos(ref DataGridViewComboBoxColumn cmb)
        {
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);
            DataTable dt = new DataTable();
            //cmb.Items.Clear();
            string sql = "SELECT modulo ";
            sql += "FROM aplicacion ";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("usuarios");
            da.Fill(dt);
            foreach (DataRow row in dt.Rows)
            {
                cmb.Items.Add(row["modulo"].ToString());
            }
        }

        public void lista_de_usuarios(ref DataGridView grid)
        {
            DataTable dt = new DataTable() ;
            try
            {
                dt.Clear();
            }
            catch (Exception ex)
            {
                dt = new DataTable();
            }
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);


            string sql = "SELECT usuarios.id,usuarios.login as Usuario, usuarios.passw as Clave,usuarios.nombre as nombre, perfiles.nombre as Perfil ";
            sql += "FROM usuarios,perfiles ";
            sql += "WHERE usuarios.perfil=perfiles.id ";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("usuarios");
            da.Fill(dt);
            int qty = dt.Rows.Count;
            grid.ColumnCount = 4;
            grid.Columns[0].Name = "Id";
            grid.Columns[1].Name = "Usuario";
            grid.Columns[2].Name = "Clave";
            grid.Columns[3].Name = "Nombre";
            string[] datos = new String[4];

            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Perfil";
            cmb.Name = "cmb";
            cmb.MaxDropDownItems = 10;
            agregar_perfiles(ref cmb);
            grid.Columns.Add(cmb);
            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                datos[0] = row["Id"].ToString();
                datos[1] = row["Usuario"].ToString();
                datos[2] = row["Clave"].ToString();
                datos[3] = row["nombre"].ToString();
                grid.Rows.Add(datos);
                grid.Rows[i].Cells[4].Value = row["Perfil"].ToString();
                i++;
                
            }
            grid.Columns[0].ReadOnly = true;
            
            
        }

        public void lista_de_perfiles(ref DataGridView grid)
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Clear();
            }
            catch (Exception ex)
            {
                dt = new DataTable();
            }
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);

            string sql = "SELECT id,nombre  ";
            sql += "FROM perfiles";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("perfiles");
            da.Fill(dt);
            int qty = dt.Rows.Count;
            grid.ColumnCount = 2;

            grid.Columns[0].Name = "Id";
            grid.Columns[1].Name = "Nombre";
            
            string[] datos = new String[2];

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                datos[0] = row["id"].ToString();
                datos[1] = row["nombre"].ToString();
                grid.Rows.Add(datos);
                i++;
            }
            grid.Columns[0].ReadOnly = true;
        }

        public void lista_de_accesos(ref DataGridView grid)
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Clear();
            }
            catch (Exception ex)
            {
                dt = new DataTable();
            }
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);

            string sql = "SELECT perfiles.nombre as Perfil,acceso.id as id_modulo,aplicacion.modulo as modulo ";
            sql += "FROM perfiles,aplicacion,acceso ";
            sql += "WHERE acceso.idperfil=perfiles.id and acceso.idaplicacion=aplicacion.id order by perfiles.nombre";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("perfiles");
            da.Fill(dt);
            int qty = dt.Rows.Count;
            grid.ColumnCount = 1;
            
            //grid.Columns[0].Name = "Perfil";
            grid.Columns[0].Name = "Id";
            //grid.Columns[1].Name = "Modulo";

            DataGridViewComboBoxColumn cmbmod = new DataGridViewComboBoxColumn();
            cmbmod.HeaderText = "Modulo";
            cmbmod.Name = "cmbmod";
            cmbmod.MaxDropDownItems = 10;
            agregar_modulos(ref cmbmod);
            grid.Columns.Add(cmbmod);

            DataGridViewComboBoxColumn cmb = new DataGridViewComboBoxColumn();
            cmb.HeaderText = "Perfil";
            cmb.Name = "cmb";
            cmb.MaxDropDownItems = 10;
            agregar_perfiles(ref cmb);
            grid.Columns.Add(cmb);
            string[] datos = new String[2];

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                datos[0] = row["id_modulo"].ToString();
                //datos[1] = row["modulo"].ToString();
                grid.Rows.Add(datos);
                grid.Rows[i].Cells[1].Value = row["modulo"].ToString();
                grid.Rows[i].Cells[2].Value = row["Perfil"].ToString();
                i++;
            }
            grid.Columns[0].ReadOnly = true;
        }

        public void lista_de_firmas(ref DataGridView grid)
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Clear();
            }
            catch (Exception ex)
            {
                dt = new DataTable();
            }
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);

            string sql = "SELECT id,nombre  ";
            sql += "FROM firma ";
            
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("firma");
            da.Fill(dt);
            int qty = dt.Rows.Count;
            grid.ColumnCount = 2;
                        
            grid.Columns[0].Name = "Id";
            grid.Columns[1].Name = "Nombre";

            
            string[] datos = new String[2];

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                datos[0] = row["id"].ToString();
                datos[1] = row["nombre"].ToString();
                grid.Rows.Add(datos);
                i++;
            }
            grid.Columns[0].ReadOnly = true;
        }

        public void lista_de_cargos(ref DataGridView grid)
        {
            DataTable dt = new DataTable();
            try
            {
                dt.Clear();
            }
            catch (Exception ex)
            {
                dt = new DataTable();
            }
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);

            string sql = "SELECT id,nombre  ";
            sql += "FROM cargo ";

            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("cargo");
            da.Fill(dt);
            int qty = dt.Rows.Count;
            grid.ColumnCount = 2;

            grid.Columns[0].Name = "Id";
            grid.Columns[1].Name = "Nombre";


            string[] datos = new String[2];

            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                datos[0] = row["id"].ToString();
                datos[1] = row["nombre"].ToString();
                grid.Rows.Add(datos);
                i++;
            }
            grid.Columns[0].ReadOnly = true;
        }

        public void lista_de_registros(ref DataGridView grid,string condicion="")
        {
            DataTable dt = new DataTable();
            DataTable dt1 = new DataTable();
            try
            {
                dt.Clear();
            }
            catch (Exception ex)
            {
                dt = new DataTable();
            }
            try
            {
                dt1.Clear();
            }
            catch (Exception ex)
            {
                dt1 = new DataTable();
            }
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);

            string sql = "SELECT Desempleo.id,Desempleo.nombre,Desempleo.pasaporte,Desempleo.Entrada";
            sql += ",Desempleo.FEntrada,Desempleo.Salida,Desempleo.FSalida,Desempleo.Firma";
            sql += ",Desempleo.Cargo,Desempleo.Observaciones,usuarios.nombre ";
            sql += "FROM Desempleo, RF, usuarios ";

            //if (condicion.Length > 0)
                sql += " WHERE Desempleo.id=RF.iddesempleo AND usuarios.id=RF.idusuario AND Desempleo.estado=1 "+condicion;

            sql += " ORDER BY Desempleo.id asc ";
            OleDbDataAdapter da = new OleDbDataAdapter(sql, con1);
            dt = new DataTable("desempleo");
            da.Fill(dt);

            int qty = dt.Rows.Count;
            grid.ColumnCount = 11;

            grid.Columns[0].Name = "Id";
            grid.Columns[1].Name = "Nombre";
            grid.Columns[2].Name = "Pasaporte";
            grid.Columns[3].Name = "Entrada";
            grid.Columns[4].Name = "Fec. Entrada";
            grid.Columns[5].Name = "Salida";
            grid.Columns[6].Name = "Fec.Salida";
            grid.Columns[7].Name = "Firma";
            grid.Columns[8].Name = "Cargo";
            grid.Columns[9].Name = "Usuario";
            grid.Columns[10].Name = "Observaciones";

            string[] datos = new String[11];

            int iddesempleo = -1;
            String rf="";
            //OleDbDataAdapter da1;
            int guardado = 0;
            foreach (DataRow row in dt.Rows)
            {
                Application.DoEvents();
                if (iddesempleo == Convert.ToInt32(row[0]) && iddesempleo > 0)
                {
                    guardado = 0;
                    if (rf.Length > 0) rf += ", ";
                    rf += row[10].ToString();

                    datos[0] = row[0].ToString();
                    datos[1] = row[1].ToString();
                    datos[2] = row[2].ToString();
                    datos[3] = row[3].ToString();
                    datos[4] = row[4].ToString();
                    datos[5] = row[5].ToString();
                    datos[6] = row[6].ToString();
                    datos[7] = row[7].ToString();
                    datos[8] = row[8].ToString();
                    datos[9] = rf;
                    datos[10] = row[9].ToString();
                }
                else
                {
                    if (iddesempleo > 0)
                    {
                        grid.Rows.Add(datos);
                        guardado = 1;
                    }
                    guardado = 0;
                    rf = row[10].ToString();   
                    datos[0] = row[0].ToString();
                    datos[1] = row[1].ToString();
                    datos[2] = row[2].ToString();
                    datos[3] = row[3].ToString();
                    datos[4] = row[4].ToString();
                    datos[5] = row[5].ToString();
                    datos[6] = row[6].ToString();
                    datos[7] = row[7].ToString();
                    datos[8] = row[8].ToString();
                    datos[9] = rf;
                    datos[10] = row[9].ToString();
                    iddesempleo = Convert.ToInt32(row[0]);
                    
                }
            }
            if (guardado == 0 && qty>0) grid.Rows.Add(datos);
            grid.Columns[0].ReadOnly = true;
            con1.Close();
        }

        public string lista_de_rf(int iddesempleo)
        {
            DataTable dt1 = new DataTable();
            System.Data.OleDb.OleDbConnection con1 = new System.Data.OleDb.OleDbConnection(stringdeconexion);
            string rf = "";

            string sql = "SELECT usuarios.nombre FROM usuarios,RF ";
            sql += "WHERE RF.iddesempleo= " + iddesempleo.ToString();
            sql += " AND RF.idusuario= usuarios.id";
            OleDbDataAdapter da1 = new OleDbDataAdapter(sql, con1);
            dt1 = new DataTable("rf");
            da1.Fill(dt1);
            con1.Close();
            foreach (DataRow row1 in dt1.Rows)
            {
                Application.DoEvents();
                if (rf.Length > 0) rf += ", ";
                rf += row1["nombre"].ToString();
            }

            return rf;
        }

        public void leer_datos(ComboBox b, String objeto,int usr)
        {
            //leer los datos desde el archivo excel
            string strTabla = string.Empty;
            DataSet oDs;

            //Nombre del rango tal cual lo definimos en el Excel
            strTabla = objeto;
            using (
            OleDbConnection oConn =
                        new System.Data.OleDb.OleDbConnection(stringdeconexion)
                )
            {
                using (OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT * FROM [" + strTabla + "]", oConn))
                {
                    Application.DoEvents();
                    oConn.Open();
                    oDs = new DataSet();
                    oCmd.Fill(oDs);
                    oConn.Close();
                }
            }
            int elem = 1;
            if (objeto.Equals("usuarios"))
                elem = 4;
            for (int fila = 0; fila < (oDs.Tables[0].Rows.Count); fila++)
            {
                Application.DoEvents();
                if (oDs.Tables[0].Rows[fila].ItemArray[elem].ToString().Length > 0)
                    b.Items.Add(oDs.Tables[0].Rows[fila].ItemArray[elem].ToString());
            }
        }

        public void leer_datos(CheckedListBox b, String objeto, int usr)
        {
            //leer los datos desde el archivo excel
            string strTabla = string.Empty;
            DataSet oDs;

            //Nombre del rango tal cual lo definimos en el Excel
            strTabla = objeto;
            using (
            OleDbConnection oConn =
                        new System.Data.OleDb.OleDbConnection(stringdeconexion)
                )
            {
                using (OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT * FROM [" + strTabla + "]", oConn))
                {
                    oConn.Open();
                    oDs = new DataSet();
                    oCmd.Fill(oDs);
                    oConn.Close();
                }
            }
            int elem = 1;
            if (objeto.Equals("usuarios"))
                elem = 4;
            for (int fila = 0; fila < (oDs.Tables[0].Rows.Count); fila++)
            {
                if (oDs.Tables[0].Rows[fila].ItemArray[elem].ToString().Length > 0)
                    b.Items.Add(oDs.Tables[0].Rows[fila].ItemArray[elem].ToString());
            }
        }

        public int ultimo_id_Desempleo()
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM Desempleo order by id desc", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public int id_Desempleo(string pasaporte)
        {
            int id = -1;
            DataSet oDs;
            OleDbConnection oConn =
                    new System.Data.OleDb.OleDbConnection(stringdeconexion);

            OleDbDataAdapter oCmd = new OleDbDataAdapter("SELECT id FROM Desempleo where Pasaporte='"+pasaporte+"'", oConn);
            oConn.Open();
            oDs = new DataSet();
            oCmd.Fill(oDs);
            oConn.Close();
            try
            {
                if (oDs.Tables[0].Rows.Count > 0)
                {
                    id = Convert.ToInt32(oDs.Tables[0].Rows[0]["id"].ToString());
                }
            }
            catch (Exception ex) { }

            return id;
        }

        public void marcarRF(ref CheckedListBox lista, string usuarios)
        {
            string[] users = usuarios.Split(',');
            int i;
            foreach (string usr in users)
            {
                for (i = 0; i < lista.Items.Count;i++)
                {
                    if (usr.Trim().Equals(lista.Items[i].ToString())){
                        lista.SetItemChecked(i, true);
                    }
                }
            }
        }
    }
}
