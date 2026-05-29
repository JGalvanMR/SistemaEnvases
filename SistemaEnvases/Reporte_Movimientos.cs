using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Data.SqlClient;

namespace SistemaEnvases
{
    public partial class Reporte_Movimientos : Form
    {

        public static string cadenaConexion = "Persist Security Info=False;user id=sa; password=Gabira1;Initial Catalog =GAB_Irapuato; server=tcp:192.168.123.6,1433; Connect Timeout = 130";
        //SqlConnection thisConnecionDBGAB = new SqlConnection(Utilerias.Class1.ConnectionStringDBGAB);
        SqlConnection thisConnecion = new SqlConnection(Utilerias.Class1.ConnectionString);
        string usuario_actual = Utilerias.Class1.Usu_login.Trim();
        DataTable Datos = new DataTable();

        public Reporte_Movimientos()
        {
            InitializeComponent();
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void Reporte_Movimientos_Load(object sender, EventArgs e)
        {
            Datos.Columns.Add("FECHA");
            Datos.Columns.Add("NOMBRE COMPUTADORA");
            Datos.Columns.Add("NOMBRE USUARIO");
            Datos.Columns.Add("FOLIO");
            Datos.Columns.Add("DETALLE");


            thisConnecion.Open();
            string queryX = "SELECT * FROM tb_registro_movimientos WHERE op_clave = '2.18' AND tipo_mov = 'ATRASADO' order by fecha DESC";
            SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
            SqlDataReader drX = cmX.ExecuteReader();

            while (drX.Read())
            {
                DataRow rowix = Datos.NewRow();
                rowix["FECHA"] = drX["fecha"];
                rowix["NOMBRE COMPUTADORA"] = drX["nom_compu"];
                rowix["NOMBRE USUARIO"] = drX["nom_usu"];
                rowix["FOLIO"] = drX["folio"];
                rowix["DETALLE"] = drX["detalle"];
                Datos.Rows.Add(rowix);

            }

            thisConnecion.Close();

            dtvinfo.AutoGenerateColumns = true;
            dtvinfo.DataSource = Datos;
            dtvinfo.AutoResizeColumns();

            dtvinfo.Columns["FECHA"].ReadOnly = true;
            dtvinfo.Columns["NOMBRE COMPUTADORA"].ReadOnly = true;
            dtvinfo.Columns["NOMBRE USUARIO"].ReadOnly = true;
            dtvinfo.Columns["FOLIO"].ReadOnly = true;
            dtvinfo.Columns["DETALLE"].ReadOnly = true;

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
