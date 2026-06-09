using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
//using Office = Microsoft.Office.Core;
using Excel = Microsoft.Office.Interop.Excel;

namespace SistemaEnvases
{
    public partial class Form1 : Form
    {
        //public static string cadenaConexion = "Persist Security Info=False;user id=sa; password=Gabira1;Initial Catalog =GAB_Irapuato; server=tcp:192.168.123.6,1433; Connect Timeout = 130";
        //SqlConnection thisConnecionDBGAB = new SqlConnection(Utilerias.Class1.ConnectionStringDBGAB);
        SqlConnection thisConnecion = new SqlConnection(Utilerias.Class1.ConnectionString);
        //SqlConnection thisConnecion = new SqlConnection("Persist Security Info=False;user id=sa; password=Gabira1;Initial Catalog =GAB_Irapuato_Prueba; server=tcp:192.168.123.6,1433; Connect Timeout = 130; MultipleActiveResultSets=True");
        string usuario_actual = Utilerias.Class1.Usu_login.Trim();
        string UsuMail = "ricardo.cortes@mrlucky.com.mx", PwdMail = "rcedillo";
        string UsuMail2 = "sistemas@mrlucky.com.mx", PwdMail2 = "sisgab";
        DataTable Datos = new DataTable();
        DataTable Cat_Cliente = new DataTable();
        DataTable Cat_Prov = new DataTable();
        DataTable Det_PedNal = new DataTable();
        DataTable Det_PedExp = new DataTable();
        DataTable TiemposEmb = new DataTable();
        public DataTable Embarques = new DataTable();
        public DataTable DetEmb = new DataTable();
        public DataTable Proveedor = new DataTable();
        public DataTable Proveedor_Repo = new DataTable();
        public DataTable Proveedor_Repo_Fin = new DataTable();
        public DataTable ProveedorIni = new DataTable();
        public DataTable ProveedorFin = new DataTable();
        public DataTable Rancho = new DataTable();
        public DataTable Tabla = new DataTable();
        public DataTable Producto = new DataTable();
        public DataTable Envase = new DataTable();
        public DataTable ProductosGuardar = new DataTable();
        public DataTable ProductosGuardar_Entrada = new DataTable();

        public DataTable Kardex = new DataTable();
        public DataTable reporteKardex = new DataTable();
        public string[] usuariostotales = {
                "JAVIER",
                "N",
                "ADMINISTRA","RODOLFO"};


        //Datatables de los reportes
        public DataTable inv_x_proveedorini = new DataTable();
        public DataTable inv_x_proveedorini_CAJONES = new DataTable();
        public DataTable inv_x_proveedorini_CAJONES_ESPARRAGO = new DataTable();
        public DataTable inv_x_proveedorini_DET = new DataTable();
        public DataTable inv_x_proveedorini_DET_CAJONES = new DataTable();
        public DataTable inv_x_proveedorini_DET_CAJONES_ESPARRAGO = new DataTable();
        public DataTable inv_x_proveedor = new DataTable();
        public DataTable inv_x_proveedor_CAJONES = new DataTable();
        public DataTable inv_x_proveedor_CAJONES_ESPARRAGO = new DataTable();
        //********************************************************
        public DataTable inv_x_proveedor_Tabla = new DataTable();

        //********************************************************

        public DataTable inv_Ini_dia = new DataTable();

        public DataTable inv_Ini_dia_cajon = new DataTable();

        public DataTable inv_Ini_dia_cajon_ESPARRAGO = new DataTable();

        public DataTable inv_Fis_dia = new DataTable();

        public DataTable Reporte_Entradas = new DataTable();
        public DataTable Reporte_Salidas = new DataTable();
        public DataTable Reporte_Kardex = new DataTable();
        //**************************
        Mensaje Form = new Mensaje();


        public string fecha_hoy_hoy = "";

        public string mes_actual_Maquina = DateTime.Today.ToString("MM");

        BackgroundWorker bg;

        BackgroundWorker bgcajones;

        BackgroundWorker bgESPARRAGO;

        BackgroundWorker bgCorteDiario;

        DataTable Pedidosdia = new DataTable();
        DataTable PedidosCamioneta = new DataTable();
        DataTable CatVehi = new DataTable();
        public static DataTable TOTALINV = new DataTable(); //(PROD_CLAVE C(10), NOMPROD C(50), CANTIDAD N(8), SURTIDO N(8))
        public static string RegistroEmb = "";
        int NReg = 0, TotReg = 0;


        string proveedor = "";
        string rancho = "";
        string tabla = "";
        public string Actualizar = "";
        private bool requiereActualizar = false;

        public Form1()
        {
            InitializeComponent();
            traeultimoidsalida();
            comboenvases();
            comboproovedor();
            comboproducto();


            comboproovedor_reporte();
            comboproovedor_reporte_fin();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            //*** Validacion de campos necesarios para ingresar al datagritview
            if ((clbenv.Text.ToString().Trim().Length == 0) || (Cmb_Envase.Text.ToString().Trim().Length == 0))
            {
                MessageBox.Show("Seleccione o ingrese un Envase válido");
                return;
            }

            if (cant.Text.Trim().Length == 0 || (Convert.ToInt32(cant.Text.Trim()) <= 0))
            {
                MessageBox.Show("La cantidad es incorrecta, Ingrese una cantidad válida");
                return;
            }

            Boolean Existe = false;
            foreach (DataRow row in ProductosGuardar.Select("Cvl_Envase = '" + clbenv.Text + "' and cvl_producto = '" + clbprod.Text + "'"))
                Existe = true;

            if (Existe)
            {
                MessageBox.Show("El Envase y el Producto ya fue ingresado, NO se Puede Agregar 2 Veces", "Ingreso de Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DataRow Contenido = ProductosGuardar.NewRow();
            Contenido["cvl_envase"] = clbenv.Text.ToString().Trim();
            Contenido["nombre_envase"] = Cmb_Envase.Text.ToString().Trim().Replace(" - " + clbenv.Text.ToString().Trim(), ""); ;
            Contenido["Cantidad"] = Convert.ToInt32(cant.Text);
            if ((clbprod.Text.ToString().Trim().Length == 0) || (Cmb_Producto.Text.ToString().Trim().Length == 0))
            {
                Contenido["nombre_producto"] = "";
                Contenido["cvl_producto"] = "";
            }
            else
            {
                Contenido["nombre_producto"] = Cmb_Producto.Text.ToString().Trim().Replace(" - " + clbprod.Text.ToString().Trim(), "");
                Contenido["cvl_producto"] = clbprod.Text.ToString().Trim();
            }


            ProductosGuardar.Rows.Add(Contenido);

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.DataSource = ProductosGuardar;

            clbprod.Text = "";
            Cmb_Producto.SelectedValue = "";
            clbenv.Text = "";
            Cmb_Envase.SelectedValue = "";
            cant.Text = "";


        }


        private void Form1_Load(object sender, EventArgs e)
        {
            thisConnecion.Open();
            string fecha = "SELECT SYSDATETIME()";
            SqlCommand cmd = new SqlCommand(fecha, thisConnecion);
            string fecha_hoy = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();


            #region VALIDA Y FORZA ACTUALIZAR
            if (ValidaActualizacion())
            {
                requiereActualizar = true;
                this.Close(); // Cerramos el formulario para dar paso al actualizador
            }
            #endregion


            ToolTip toolTip1 = new ToolTip();

            // Set up the delays for the ToolTip.
            toolTip1.AutoPopDelay = 5000;
            toolTip1.InitialDelay = 1000;
            toolTip1.ReshowDelay = 500;
            // Force the ToolTip text to be displayed whether or not the form is active.
            toolTip1.ShowAlways = true;

            // Set up the ToolTip text for the Button and Checkbox.
            toolTip1.SetToolTip(this.button18, "Iniciar Cierre diario de Inventario de Envases");
            /*thisConnecion.Open();
            mes_actual = "SELECT  datepart(mm, getdate())";
            SqlCommand cmdmes = new SqlCommand(mes_actual, thisConnecion);
            mes_actual = cmdmes.ExecuteScalar().ToString();
            thisConnecion.Close();*/


            /*thisConnecion.Open();
            string valida = "SELECT fecha FROM tb_registro_movimientos where fecha = '" + fecha_hoy + "' AND detalle = 'CORTE INVENTARIO ENVASES' AND op_clave = '2.18'";
            cmd = new SqlCommand(valida, thisConnecion);
            string existencia = Convert.ToString(cmd.ExecuteScalar());
            thisConnecion.Close();*/


            fecha_hoy_hoy = fecha_hoy;

            dateTimeinventarioinicial.Text = fecha_hoy; // fecha del sistema

            /*if (existencia.Trim().Length == 0 || existencia.Trim() == null) {
                Form.Show();
                thisConnecion.Open();
                string query = "SELECT env_clave FROM  tb_cat_envases WHERE (env_inventario = '1')";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();
                
                valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos ORDER BY ENV_FECHA DESC";
                cmd = new SqlCommand(valida, thisConnecion);
                string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
                

                while (dr.Read())
                {
                    string clave_envase = Convert.ToString(dr["env_clave"]).Trim();

                    //thisConnecion.Open();
                    valida = "SELECT sum(ENV_INV_INI_CANT) FROM TB_MSTR_INV_ENVASES_dos WHERE ENV_CLAVE = '" + clave_envase + "' AND ENV_FECHA = '"+ fecha_ent +"'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND ENT_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND SAL_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    string actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "'";
                    cmd = new SqlCommand(actualizadato, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "'  AND ENV_CLAVE = '" + clave_envase + "'";
                    cmd = new SqlCommand(actualizadato, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();

                    int inv_ini_hoy = inv_ini + entradas - salidas;

                    //thisConnecion.Open();
                    string cadena = "insert into TB_MSTR_INV_ENVASES_dos (ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                    "Values('" + clave_envase + "','" + fecha_hoy + "',' " + inv_ini_hoy + " ','','')";
                    cmd = new SqlCommand(cadena, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();
                }

                //INVENTARIO_CAJONES_PROVEEDOR(fecha_ent);

                thisConnecion.Close();

                thisConnecion.Open();
                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values('" + fecha_hoy + "','" + Environment.MachineName.Trim() + "','','ACTUALIZAR','2.18','', 'CORTE INVENTARIO ENVASES', 'SISGAB', '')";
                cmd = new SqlCommand(cadenaregmov, thisConnecion);
                cmd.ExecuteNonQuery();
                thisConnecion.Close();

                Form.Close();
            }*/


            ProductosGuardar.Columns.Add("cvl_envase");
            ProductosGuardar.Columns.Add("nombre_envase");
            ProductosGuardar.Columns.Add("Cantidad");
            ProductosGuardar.Columns.Add("nombre_producto");
            ProductosGuardar.Columns.Add("cvl_producto");

            ProductosGuardar_Entrada.Columns.Add("cvl_envase");
            ProductosGuardar_Entrada.Columns.Add("nombre_envase");
            ProductosGuardar_Entrada.Columns.Add("Cantidad");
            ProductosGuardar_Entrada.Columns.Add("nombre_producto");
            ProductosGuardar_Entrada.Columns.Add("cvl_producto");


            inv_x_proveedor.Columns.Add("ID");
            inv_x_proveedor.Columns.Add("NOMBRE");
            inv_x_proveedor.Columns.Add("FECHA");
            inv_x_proveedor.Columns.Add("INVENTARIO INICIAL");
            inv_x_proveedor.Columns.Add("ENTRADAS");
            inv_x_proveedor.Columns.Add("SALIDAS");
            inv_x_proveedor.Columns.Add("INVENTARIO FINAL");


            inv_x_proveedor_CAJONES.Columns.Add("ID");
            inv_x_proveedor_CAJONES.Columns.Add("NOMBRE");
            inv_x_proveedor_CAJONES.Columns.Add("ID_PROVEEDOR");
            inv_x_proveedor_CAJONES.Columns.Add("FECHA");
            inv_x_proveedor_CAJONES.Columns.Add("INVENTARIO INICIAL");
            inv_x_proveedor_CAJONES.Columns.Add("ENTRADAS");
            inv_x_proveedor_CAJONES.Columns.Add("SALIDAS");
            inv_x_proveedor_CAJONES.Columns.Add("INVENTARIO FINAL");

            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("ID");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("NOMBRE");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("ID_PROVEEDOR");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("FECHA");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("INVENTARIO INICIAL");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("ENTRADAS");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("SALIDAS");
            inv_x_proveedor_CAJONES_ESPARRAGO.Columns.Add("INVENTARIO FINAL");




            //Reporte Entradas
            Reporte_Entradas.Columns.Add("FOLIO");
            Reporte_Entradas.Columns.Add("FECHA");
            Reporte_Entradas.Columns.Add("PROVEEDOR_CLAVE");
            Reporte_Entradas.Columns.Add("PROVEEDOR_NOMBRE");
            Reporte_Entradas.Columns.Add("RANCHO_CLAVE");
            Reporte_Entradas.Columns.Add("RANCHO_NOMBRE");
            Reporte_Entradas.Columns.Add("TABLA_CLAVE");
            Reporte_Entradas.Columns.Add("TABLA_NOMBRE");
            Reporte_Entradas.Columns.Add("IDENVASE");
            Reporte_Entradas.Columns.Add("NOMBREENVASE");
            Reporte_Entradas.Columns.Add("CANTIDAD");
            Reporte_Entradas.Columns.Add("PROD_CLAVE");
            Reporte_Entradas.Columns.Add("PROD_NOMBRE");
            Reporte_Entradas.Columns.Add("RECIBO_MP");
            Reporte_Entradas.Columns.Add("ESTATUS");

            //REPORTE SALIDAS
            Reporte_Salidas.Columns.Add("FOLIO");
            Reporte_Salidas.Columns.Add("FECHA");
            Reporte_Salidas.Columns.Add("PROVEEDOR_CLAVE");
            Reporte_Salidas.Columns.Add("PROVEEDOR_NOMBRE");
            Reporte_Salidas.Columns.Add("RANCHO_CLAVE");
            Reporte_Salidas.Columns.Add("RANCHO_NOMBRE");
            Reporte_Salidas.Columns.Add("TABLA_CLAVE");
            Reporte_Salidas.Columns.Add("TABLA_NOMBRE");
            Reporte_Salidas.Columns.Add("IDENVASE");
            Reporte_Salidas.Columns.Add("NOMBREENVASE");
            Reporte_Salidas.Columns.Add("CANTIDAD");
            Reporte_Salidas.Columns.Add("PROD_CLAVE");
            Reporte_Salidas.Columns.Add("PROD_NOMBRE");
            Reporte_Salidas.Columns.Add("ESTATUS");

            //KARDEX
            Reporte_Kardex.Columns.Add("NUMERO");
            Reporte_Kardex.Columns.Add("FOLIO");
            Reporte_Kardex.Columns.Add("FECHA");
            Reporte_Kardex.Columns.Add("PROVEEDOR_CLAVE");
            Reporte_Kardex.Columns.Add("PROVEEDOR_NOMBRE");
            Reporte_Kardex.Columns.Add("RANCHO_CLAVE");
            Reporte_Kardex.Columns.Add("RANCHO_NOMBRE");
            Reporte_Kardex.Columns.Add("TABLA_CLAVE");
            Reporte_Kardex.Columns.Add("TABLA_NOMBRE");
            Reporte_Kardex.Columns.Add("IDENVASE");
            Reporte_Kardex.Columns.Add("NOMBREENVASE");
            Reporte_Kardex.Columns.Add("CANTIDAD");
            Reporte_Kardex.Columns.Add("PROD_CLAVE");
            Reporte_Kardex.Columns.Add("PROD_NOMBRE");
            Reporte_Kardex.Columns.Add("RECIBO_MP");
            Reporte_Kardex.Columns.Add("TIPO");
            Reporte_Kardex.Columns.Add("ESTATUS");

            //Inventario Inicial del dia
            inv_Ini_dia.Columns.Add("CLV");
            inv_Ini_dia.Columns.Add("NOMBRE DEL ENVASE");
            inv_Ini_dia.Columns.Add("CANTIDAD");
            inv_Ini_dia.Columns.Add("CANTIDAD MODIFICADA");

            inv_Ini_dia_cajon.Columns.Add("CLV");
            inv_Ini_dia_cajon.Columns.Add("CLV_PROV");
            inv_Ini_dia_cajon.Columns.Add("NOMBRE DEL ENVASE");
            inv_Ini_dia_cajon.Columns.Add("CANTIDAD");
            inv_Ini_dia_cajon.Columns.Add("CANTIDAD MODIFICADA");


            inv_Ini_dia_cajon_ESPARRAGO.Columns.Add("CLV");
            inv_Ini_dia_cajon_ESPARRAGO.Columns.Add("CLV_PROV");
            inv_Ini_dia_cajon_ESPARRAGO.Columns.Add("NOMBRE DEL ENVASE");
            inv_Ini_dia_cajon_ESPARRAGO.Columns.Add("CANTIDAD");
            inv_Ini_dia_cajon_ESPARRAGO.Columns.Add("CANTIDAD MODIFICADA");

            inv_Fis_dia.Columns.Add("CLV");
            inv_Fis_dia.Columns.Add("NOMBRE DEL ENVASE");
            inv_Fis_dia.Columns.Add("CANTIDAD");
            inv_Fis_dia.Columns.Add("ENTRADAS");
            inv_Fis_dia.Columns.Add("SALIDAS");
            inv_Fis_dia.Columns.Add("INVENTARIO TEORICO");
            inv_Fis_dia.Columns.Add("INVENTARIO FISICO");
            inv_Fis_dia.Columns.Add("OBS");
            inv_Fis_dia.Columns.Add("ENCAMPO");
            inv_Fis_dia.Columns.Add("CONPROD");
            inv_Fis_dia.Columns.Add("VACIAS");
            inv_Fis_dia.Columns.Add("CONBASURA");
            inv_Fis_dia.Columns.Add("XREPARAR");
            inv_Fis_dia.Columns.Add("REPARAGUI");


            thisConnecion.Open();
            string queryX = "SELECT * FROM  TB_MSTR_INV_ENVASES_dos INNER JOIN tb_cat_envases AS B ON TB_MSTR_INV_ENVASES_dos.ENV_CLAVE = B.env_clave WHERE (ENV_FECHA = '" + fecha_hoy + "')";
            SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
            SqlDataReader drX = cmX.ExecuteReader();
            Int32 InvF = 0;
            while (drX.Read())
            {
                DataRow rowix = inv_Ini_dia.NewRow();
                rowix["CLV"] = drX["ENV_CLAVE"];
                rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                rowix["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia.Rows.Add(rowix);
                DataRow rowix1 = inv_Fis_dia.NewRow();
                rowix1["CLV"] = drX["ENV_CLAVE"];
                rowix1["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix1["CANTIDAD"] = Convert.ToInt32(drX["ENV_INV_INI_CANT"]).ToString("#,##0");
                rowix1["ENTRADAS"] = Convert.ToInt32(drX["ENV_ENTR_CANT"]).ToString("#,##0");
                rowix1["SALIDAS"] = Convert.ToInt32(drX["ENV_SAL_CANT"]).ToString("#,##0");
                rowix1["INVENTARIO TEORICO"] = (Convert.ToInt32(drX["ENV_INV_INI_CANT"]) + Convert.ToInt32(drX["ENV_ENTR_CANT"]) - Convert.ToInt32(drX["ENV_SAL_CANT"])).ToString("#,##0");
                rowix1["INVENTARIO FISICO"] = (Convert.ToString(drX["ENV_INV_FISICO"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_INV_FISICO"]).ToString("#,##0");
                rowix1["OBS"] = drX["ENV_OBS_DIF"].ToString();
                rowix1["ENCAMPO"] = (Convert.ToString(drX["ENV_ENCAMPO"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_ENCAMPO"]).ToString("#,##0");
                rowix1["CONPROD"] = (Convert.ToString(drX["ENV_CONPROD"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_CONPROD"]).ToString("#,##0");
                rowix1["VACIAS"] = (Convert.ToString(drX["ENV_VACIAS"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_VACIAS"]).ToString("#,##0");
                rowix1["CONBASURA"] = (Convert.ToString(drX["ENV_CONBASURA"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_CONBASURA"]).ToString("#,##0");
                rowix1["XREPARAR"] = (Convert.ToString(drX["ENV_XREPARAR"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_XREPARAR"]).ToString("#,##0");
                rowix1["REPARAGUI"] = (Convert.ToString(drX["ENV_REPARAGUI"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_REPARAGUI"]).ToString("#,##0");
                InvF += Convert.ToInt32(Convert.ToDecimal(rowix1["INVENTARIO FISICO"]));
                inv_Fis_dia.Rows.Add(rowix1);

            }

            thisConnecion.Close();

            dataGridInfoInvenario.AutoGenerateColumns = true;
            dataGridInfoInvenario.AllowUserToAddRows = false;
            dataGridInfoInvenario.DataSource = inv_Ini_dia;
            dataGridInfoInvenario.AutoResizeColumns();

            dataGridInfoInvenario.Columns["CLV"].ReadOnly = true;
            dataGridInfoInvenario.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            dataGridInfoInvenario.Columns["CANTIDAD"].ReadOnly = true;


            dataGridInfoInvenario.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            thisConnecion.Close();
            DGInvFis.DataSource = inv_Fis_dia;

            DGInvFis.AutoGenerateColumns = true;
            DGInvFis.AllowUserToAddRows = false;
            DGInvFis.DataSource = inv_Fis_dia;
            DGInvFis.AutoResizeColumns();

            DGInvFis.Columns["CLV"].ReadOnly = true;
            DGInvFis.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            DGInvFis.Columns["CANTIDAD"].ReadOnly = true;
            DGInvFis.Columns["CANTIDAD"].HeaderText = "INV. INICIAL";
            DGInvFis.Columns["CLV"].ReadOnly = true;
            DGInvFis.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            DGInvFis.Columns["CANTIDAD"].ReadOnly = true;
            DGInvFis.Columns["ENTRADAS"].ReadOnly = true;
            DGInvFis.Columns["SALIDAS"].ReadOnly = true;
            DGInvFis.Columns["INVENTARIO TEORICO"].ReadOnly = true;
            DGInvFis.Columns["INVENTARIO TEORICO"].HeaderText = "INV. TEORICO";
            DGInvFis.Columns["INVENTARIO FISICO"].HeaderText = "INV. FISICO";
            DGInvFis.Columns["ENCAMPO"].ReadOnly = true;
            DGInvFis.Columns["EnCampo"].Visible = false;
            DGInvFis.Columns["CONPROD"].HeaderText = "CON PRODUCTO";
            DGInvFis.Columns["VACIAS"].HeaderText = "VACIAS";
            DGInvFis.Columns["CONBASURA"].HeaderText = "CON BASURA";
            DGInvFis.Columns["XREPARAR"].HeaderText = "X REPARAR";
            DGInvFis.Columns["REPARAGUI"].HeaderText = "ENVASE ROTO";
            DGInvFis.Columns["ENCAMPO"].HeaderText = "EN CAMPO";
            DGInvFis.Columns[0].Width = 40;
            DGInvFis.Columns[1].Width = 220;
            DGInvFis.Columns[2].Width = 55;
            DGInvFis.Columns[3].Width = 55;
            DGInvFis.Columns[4].Width = 55;
            DGInvFis.Columns[5].Width = 55;
            DGInvFis.Columns[6].Width = 60;
            DGInvFis.Columns[7].Width = 200;
            DGInvFis.Columns[8].Width = 55;
            DGInvFis.Columns[9].Width = 55;
            DGInvFis.Columns[10].Width = 55;
            DGInvFis.Columns[11].Width = 55;
            DGInvFis.Columns[12].Width = 55;
            DGInvFis.Columns[13].Width = 55;
            for (int i = 2; i <= 6; i++)
            {
                DGInvFis.Columns[i].DefaultCellStyle.Format = "#,###";
                DGInvFis.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            for (int i = 7; i <= 13; i++)
            {
                DGInvFis.Columns[i].DefaultCellStyle.Format = "#,###";
                DGInvFis.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            }
            BtnSavFis.Enabled = true;
            if (InvF > 0)
                BtnSavFis.Enabled = false;
            DGInvFis.Columns[6].Frozen = true;
            //DGInvFis.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;



            //Inventario Inicial de Cajones por Dia
            thisConnecion.Open();
            queryX = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + fecha_hoy + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
            cmX = new SqlCommand(queryX, thisConnecion);
            drX = cmX.ExecuteReader();
            int cont = 1;
            while (drX.Read())
            {
                if (cont == 1)
                {
                    DataRow rowixcajon = inv_Ini_dia_cajon.NewRow();
                    rowixcajon["CLV"] = "";
                    rowixcajon["CLV_PROV"] = drX["prov_clave"];
                    rowixcajon["NOMBRE DEL ENVASE"] = drX["prov_nombre"];
                    rowixcajon["CANTIDAD"] = "";
                    rowixcajon["CANTIDAD MODIFICADA"] = "";
                    inv_Ini_dia_cajon.Rows.Add(rowixcajon);
                    cont = 0;

                }
                else
                {
                    cont++;

                }

                DataRow rowix = inv_Ini_dia_cajon.NewRow();
                rowix["CLV"] = drX["ENV_CLAVE"];
                rowix["CLV_PROV"] = drX["prov_clave"];
                rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                rowix["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia_cajon.Rows.Add(rowix);

            }

            thisConnecion.Close();


            dataGridCajones.AutoGenerateColumns = true;
            dataGridCajones.AllowUserToAddRows = false;
            dataGridCajones.DataSource = inv_Ini_dia_cajon;
            dataGridCajones.AutoResizeColumns();

            dataGridCajones.Columns["CLV"].ReadOnly = true;
            dataGridCajones.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            dataGridCajones.Columns["CANTIDAD"].ReadOnly = true;


            dataGridCajones.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;




            //Inventario Inicial de Cajones por Dia Esparrago
            thisConnecion.Open();
            queryX = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ_ESPARRAGO AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV_ESPARRAGO AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + fecha_hoy + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
            cmX = new SqlCommand(queryX, thisConnecion);
            drX = cmX.ExecuteReader();
            //int cont = 1;
            while (drX.Read())
            {
                //if (cont == 1)
                //{
                DataRow rowixcajon = inv_Ini_dia_cajon_ESPARRAGO.NewRow();
                rowixcajon["CLV"] = "";
                rowixcajon["CLV_PROV"] = drX["prov_clave"];
                rowixcajon["NOMBRE DEL ENVASE"] = drX["prov_nombre"];
                rowixcajon["CANTIDAD"] = "";
                rowixcajon["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia_cajon_ESPARRAGO.Rows.Add(rowixcajon);
                cont = 0;

                //}
                //else
                //{
                //    cont++;

                //}

                DataRow rowix = inv_Ini_dia_cajon_ESPARRAGO.NewRow();
                rowix["CLV"] = drX["ENV_CLAVE"];
                rowix["CLV_PROV"] = drX["prov_clave"];
                rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                rowix["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia_cajon_ESPARRAGO.Rows.Add(rowix);

            }

            thisConnecion.Close();


            dataGridEsparrago.AutoGenerateColumns = true;
            dataGridEsparrago.AllowUserToAddRows = false;
            dataGridEsparrago.DataSource = inv_Ini_dia_cajon_ESPARRAGO;
            dataGridEsparrago.AutoResizeColumns();

            dataGridEsparrago.Columns["CLV"].ReadOnly = true;
            dataGridEsparrago.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            dataGridEsparrago.Columns["CANTIDAD"].ReadOnly = true;


            dataGridEsparrago.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;



            if (usuariostotales.Contains(usuario_actual) == false)
            {
                tabControl1.TabPages.Remove(tabPage4);
                tabControl1.TabPages.Remove(tabPage5);
                tabControl1.TabPages.Remove(tabPage6);
                fechasal.Enabled = false;
                fecha_entrada.Enabled = false;
            }

        }

        private void INVENTARIO_CAJONES_PROVEEDOR(string fecha_ent)
        {
            string query = "SELECT cve_prov FROM  Tb_ENV_PROV_CAJ WHERE (estatus = '1')";
            SqlCommand cm = new SqlCommand(query, thisConnecion);
            SqlDataReader dr = cm.ExecuteReader();

            while (dr.Read())
            {
                string clave_proveedor = Convert.ToString(dr["cve_prov"]).Trim();



                string[] claves = new string[2] { "02", "80" };
                foreach (string clave_envase in claves)
                {
                    //thisConnecion.Open();
                    string valida = "SELECT ISNULL(sum(ENV_INV_INI_CANT), 0) FROM TB_MSTR_INV_CAJAS_PROV WHERE ENV_CLAVE = '" + clave_envase + "' AND ENV_FECHA = '" + fecha_ent + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                    SqlCommand cmd = new SqlCommand(valida, thisConnecion);
                    int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND ENT_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND SAL_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    string actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "'";
                    cmd = new SqlCommand(actualizadato, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "'  AND ENV_CLAVE = '" + clave_envase + "'";
                    cmd = new SqlCommand(actualizadato, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();

                    int inv_ini_hoy = inv_ini + entradas - salidas;

                    //thisConnecion.Open();
                    string cadena = "insert into TB_MSTR_INV_ENVASES_dos (ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                    "Values('" + clave_envase + "','" + fecha_hoy_hoy + "',' " + inv_ini_hoy + " ','','')";
                    cmd = new SqlCommand(cadena, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();
                }
            }



        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }



        public void comboproovedor()
        {
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT ('100') AS Numero, prov_clave, CONCAT(RTRIM(prov_nombre), ' - ', RTRIM(prov_clave)) AS prov_nombre FROM tb_Cat_Proveedor Order By prov_Nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "prov");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            Proveedor.Load(Info);
            DataColumn column;

            DataRow rowi = Proveedor.NewRow();
            rowi["Numero"] = "0";
            rowi["prov_clave"] = "";
            rowi["prov_nombre"] = "Seleccione una opcion";
            Proveedor.Rows.Add(rowi);

            Proveedor.DefaultView.Sort = "Numero ASC";

            thisConnecion.Close();

            int T = 0;
            if (Proveedor.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();
            AutoCompleteStringCollection stringrep1 = new AutoCompleteStringCollection();
            AutoCompleteStringCollection stringrep2 = new AutoCompleteStringCollection();
            AutoCompleteStringCollection stringent = new AutoCompleteStringCollection();

            foreach (DataRow row in Proveedor.Rows)
            {
                stringCol.Add(Convert.ToString(row["prov_nombre"]));
                stringrep1.Add(Convert.ToString(row["prov_nombre"]));
                stringrep2.Add(Convert.ToString(row["prov_nombre"]));
                stringent.Add(Convert.ToString(row["prov_nombre"]));
            }

            ProveedorFin = Proveedor;


            cmb_proveedor.DataSource = Proveedor;
            cmb_proveedor.DisplayMember = "prov_nombre";
            cmb_proveedor.ValueMember = "prov_clave";

            cmb_proveedor.AutoCompleteCustomSource = stringCol;

            cmb_proveedor.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmb_proveedor.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;

            //Llenado de Combo Proveedor de la entrada ********************************----*******************

            cmb_prov_entr.DataSource = Proveedor;
            cmb_prov_entr.DisplayMember = "prov_nombre";
            cmb_prov_entr.ValueMember = "prov_clave";

            cmb_prov_entr.AutoCompleteCustomSource = stringent;

            cmb_prov_entr.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmb_prov_entr.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;


            //Llenado de combo Proveedor del Reporte
            // Inicio ********************************
            cmbprovinirep.DataSource = ProveedorFin;
            cmbprovinirep.DisplayMember = "prov_nombre";
            cmbprovinirep.ValueMember = "prov_clave";





            //Final **********************************

            cmbprovfinrep.DataSource = ProveedorFin;
            cmbprovfinrep.DisplayMember = "prov_nombre";
            cmbprovfinrep.ValueMember = "prov_clave";

            cmbprovfinrep.AutoCompleteCustomSource = stringrep2;

            cmbprovfinrep.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbprovfinrep.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;

        }



        public void comboproovedor_reporte()
        {
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT ('100') AS Numero, prov_clave AS provclaveini, CONCAT(RTRIM(prov_nombre), ' - ', RTRIM(prov_clave)) AS prov_nombre_ini FROM tb_Cat_Proveedor Order By prov_Nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "prov");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            Proveedor_Repo.Load(Info);

            DataRow rowi = Proveedor_Repo.NewRow();
            rowi["Numero"] = "0";
            rowi["provclaveini"] = "";
            rowi["prov_nombre_ini"] = "Seleccione una opcion";

            Proveedor_Repo.Rows.Add(rowi);

            Proveedor_Repo.DefaultView.Sort = "Numero ASC";

            thisConnecion.Close();

            int T = 0;
            if (Proveedor_Repo.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            AutoCompleteStringCollection stringrep1 = new AutoCompleteStringCollection();

            foreach (DataRow row in Proveedor_Repo.Rows)
            {
                stringrep1.Add(Convert.ToString(row["prov_nombre_ini"]));
            }

            ProveedorIni = Proveedor_Repo;

            //Llenado de combo Proveedor del Reporte
            // Inicio ********************************
            cmbprovinirep.DataSource = Proveedor_Repo;
            cmbprovinirep.DisplayMember = "prov_nombre_ini";
            cmbprovinirep.ValueMember = "provclaveini";

            cmbprovinirep.AutoCompleteCustomSource = stringrep1;

            cmbprovinirep.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbprovinirep.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;
        }

        public void comboproovedor_reporte_fin()
        {
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT ('100') AS Numero, prov_clave AS provclavefin, CONCAT(RTRIM(prov_nombre), ' - ', RTRIM(prov_clave)) AS prov_nombre_fin FROM tb_Cat_Proveedor Order By prov_Nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "prov");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            Proveedor_Repo_Fin.Load(Info);
            DataColumn column;

            DataRow rowi = Proveedor_Repo_Fin.NewRow();
            rowi["Numero"] = "0";
            rowi["provclavefin"] = "";
            rowi["prov_nombre_fin"] = "Seleccione una opcion";

            Proveedor_Repo_Fin.Rows.Add(rowi);

            Proveedor_Repo_Fin.DefaultView.Sort = "Numero ASC";

            thisConnecion.Close();

            int T = 0;
            if (Proveedor_Repo_Fin.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            AutoCompleteStringCollection stringrep2 = new AutoCompleteStringCollection();

            foreach (DataRow row in Proveedor_Repo_Fin.Rows)
            {
                stringrep2.Add(Convert.ToString(row["prov_nombre_fin"]));
            }

            ProveedorFin = Proveedor_Repo_Fin;

            //Llenado de combo Proveedor del Reporte
            // Inicio ********************************
            cmbprovfinrep.DataSource = ProveedorFin;
            cmbprovfinrep.DisplayMember = "prov_nombre_fin";
            cmbprovfinrep.ValueMember = "provclavefin";

            cmbprovfinrep.AutoCompleteCustomSource = stringrep2;

            cmbprovfinrep.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbprovfinrep.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;


        }


        public void comboRancho(string modo)
        {
            string valor = "";
            if (modo == "SALIDA")
            {
                valor = clbprov.Text.ToString().Trim();
            }
            else
            {
                valor = clb_prov_entr.Text.ToString().Trim();
            }
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT '100' AS Numero, rch_clave,rch_nombre  FROM tb_cat_ranchos WHERE prov_clave = '" + valor.ToString().Trim() + "' Order By rch_Nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "ranch");
            Rancho = ds1.Tables["ranch"];
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            DataColumn column;



            DataRow rowi = Rancho.NewRow();
            rowi["Numero"] = "0";
            rowi["rch_clave"] = "";
            if (Rancho.Rows.Count > 0)
            {
                rowi["rch_nombre"] = "Seleccione un Rancho";
            }
            else
            {
                rowi["rch_nombre"] = "Sin Ranchos Disponibles";
            }

            Rancho.Rows.Add(rowi);

            Rancho.DefaultView.Sort = "Numero ASC";

            thisConnecion.Close();


            AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();

            foreach (DataRow row in Rancho.Rows)
            {
                stringCol.Add(Convert.ToString(row["rch_nombre"]));
            }


            if (modo == "SALIDA")
            {
                Cmb_Rancho.DataSource = Rancho;
                Cmb_Rancho.DisplayMember = "rch_nombre";
                Cmb_Rancho.ValueMember = "rch_clave";

                Cmb_Rancho.AutoCompleteCustomSource = stringCol;
                //Cmb_Rancho.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                Cmb_Rancho.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;


                clbrancho.Text = "";

            }
            else
            {

                cmb_rancho_entr.DataSource = Rancho;
                cmb_rancho_entr.DisplayMember = "rch_nombre";
                cmb_rancho_entr.ValueMember = "rch_clave";

                cmb_rancho_entr.AutoCompleteCustomSource = stringCol;
                //cmb_rancho_entr.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cmb_rancho_entr.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;

                clb_rancho_entr.Text = "";
            }



            //Llenado de Combo Proveedor ********************************----*******************

        }


        public void comboTabla(string modo)
        {
            string proveedor = "";
            string rancho = "";
            if (modo == "SALIDA")
            {
                proveedor = clbprov.Text.ToString().Trim();
                rancho = clbrancho.Text.ToString().Trim();
            }
            else
            {
                proveedor = clb_prov_entr.Text.ToString().Trim();
                rancho = clb_rancho_entr.Text.ToString().Trim();
            }
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT '100' AS Numero, tbl_clave,tbl_nombre  FROM tb_cat_tablas WHERE prov_clave = '" + proveedor.ToString().Trim() + "' AND rch_clave = '" + rancho.ToString().Trim() + "' Order By tbl_nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "tbl");
            Tabla = ds1.Tables["tbl"];
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            DataColumn column;



            DataRow rowi = Tabla.NewRow();
            rowi["Numero"] = "0";
            rowi["tbl_clave"] = "";
            if (Tabla.Rows.Count > 0)
            {
                rowi["tbl_nombre"] = "Seleccione una Tabla";
            }
            else
            {
                rowi["tbl_nombre"] = "Sin Tablas Disponibles";
            }

            Tabla.Rows.Add(rowi);

            Tabla.DefaultView.Sort = "Numero ASC";

            thisConnecion.Close();


            AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();

            foreach (DataRow row in Tabla.Rows)
            {
                stringCol.Add(Convert.ToString(row["tbl_nombre"]));
            }

            if (modo == "SALIDA")
            {
                Cmb_Tabla.DataSource = Tabla;
                Cmb_Tabla.DisplayMember = "tbl_nombre";
                Cmb_Tabla.ValueMember = "tbl_clave";

                Cmb_Tabla.AutoCompleteCustomSource = stringCol;
                Cmb_Tabla.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                Cmb_Tabla.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;


                clbtabla.Text = "";

            }
            else
            {

                cmb_tabla_entr.DataSource = Tabla;
                cmb_tabla_entr.DisplayMember = "tbl_nombre";
                cmb_tabla_entr.ValueMember = "tbl_clave";

                cmb_tabla_entr.AutoCompleteCustomSource = stringCol;
                cmb_tabla_entr.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                cmb_tabla_entr.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;


                clb_tabla_entr.Text = "";

            }



            //Llenado de Combo Proveedor ********************************----*******************

        }

        private void Cmb_Rancho_SelectionChangeCommitted(object sender, EventArgs e)
        {
            clbrancho.Text = Cmb_Rancho.SelectedValue.ToString();
            comboTabla("SALIDA");

        }

        private void cmb_proveedor_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmb_proveedor.SelectedValue == null)
            {
                foreach (DataRow row in Proveedor.Rows)
                {
                    if (row["prov_nombre"] == cmb_proveedor.Text)
                    {
                        clbprov.Text = cmb_proveedor.ValueMember.ToString();
                    }
                }
            }
            else
            {
                clbprov.Text = cmb_proveedor.SelectedValue.ToString();
            }


            comboRancho("SALIDA");
        }

        private void Cmb_Rancho_SelectedValueChanged(object sender, EventArgs e)
        {
            if (Cmb_Rancho.SelectedValue == null)
            {
                clbrancho.Text = "";
                MessageBox.Show("El Rancho No existe");
            }
            else
            {
                clbrancho.Text = Cmb_Rancho.SelectedValue.ToString();
                comboTabla("SALIDA");
            }

        }

        private void Cmb_Tabla_SelectedValueChanged(object sender, EventArgs e)
        {
            if (Cmb_Tabla.SelectedValue == null)
            {
                clbtabla.Text = "";
                MessageBox.Show("La Tabla No existe");
            }
            else
            {
                clbtabla.Text = Cmb_Tabla.SelectedValue.ToString();

            }
        }

        public void comboproducto()
        {
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT ('100') AS Numero, prod_clave, CONCAT(RTRIM(prod_nombre), ' - ', RTRIM(prod_clave)) AS prod_nombre, prod_presentacion FROM  tb_cat_producto Order By prod_Nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "prod");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            Producto.Load(Info);
            DataColumn column;

            DataRow rowi = Producto.NewRow();
            rowi["Numero"] = "0";
            rowi["prod_clave"] = "";
            rowi["prod_nombre"] = "Seleccione un Producto";
            Producto.Rows.Add(rowi);

            Producto.DefaultView.Sort = "Numero ASC";

            thisConnecion.Close();

            int T = 0;
            if (Producto.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();

            foreach (DataRow row in Producto.Rows)
            {
                stringCol.Add(Convert.ToString(row["prod_nombre"]));
            }


            Cmb_Producto.DataSource = Producto;
            Cmb_Producto.DisplayMember = "prod_nombre";
            Cmb_Producto.ValueMember = "prod_clave";

            Cmb_Producto.AutoCompleteCustomSource = stringCol;
            Cmb_Producto.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            Cmb_Producto.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;

            //Llenado de Combo pRODUCTO EN ENTRADA ********************************----*******************

            cmb_producto_entr.DataSource = Producto;
            cmb_producto_entr.DisplayMember = "prod_nombre";
            cmb_producto_entr.ValueMember = "prod_clave";

            cmb_producto_entr.AutoCompleteCustomSource = stringCol;
            cmb_producto_entr.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmb_producto_entr.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;

        }

        public void traeultimoidsalida()
        {
            thisConnecion.Open();
            string Cadena = "SELECT TOP (1) FOLIO FROM TB_SALIDAS_ENVASES ORDER BY FOLIO DESC";
            SqlCommand cmd = new SqlCommand(Cadena, thisConnecion);
            int Valor = Convert.ToInt32(cmd.ExecuteScalar()) + 1;
            clb_conse.Text = Valor.ToString();
            thisConnecion.Close();


            thisConnecion.Open();
            Cadena = "SELECT TOP (1) FOLIO FROM TB_ENTRADAS_ENVASES ORDER BY FOLIO DESC";
            cmd = new SqlCommand(Cadena, thisConnecion);
            Valor = Convert.ToInt32(cmd.ExecuteScalar()) + 1;
            conse_entr.Text = Valor.ToString();
            thisConnecion.Close();
        }


        public int traeultimorecibomp()
        {
            thisConnecion.Open();
            string Cadena = "SELECT TOP (1) rmp_recibo FROM rmp_recibo ORDER BY rmp_recibo DESC";
            SqlCommand cmd = new SqlCommand(Cadena, thisConnecion);
            int Valor = Convert.ToInt32(cmd.ExecuteScalar()) + 1;
            thisConnecion.Close();
            return Valor;
        }

        public void comboenvases()
        {
            //Llenado de Combo Proveedor ********************************----*******************
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT ('100') AS Numero, env_clave, CONCAT(RTRIM(env_nombre), ' - ', RTRIM(env_clave)) AS env_nombre FROM  tb_cat_envases Order By env_Nombre";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "env");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            SqlDataReader InfoReporte;
            Info = cmd.ExecuteReader();
            Envase.Load(Info);

            thisConnecion.Close();
            thisConnecion.Open();
            InfoReporte = cmd.ExecuteReader();
            DataTable Envase_Reporte = new DataTable();
            Envase_Reporte.Load(InfoReporte);
            DataColumn column;

            DataRow rowi = Envase.NewRow();
            rowi["Numero"] = "0";
            rowi["env_clave"] = "";
            rowi["env_nombre"] = "Seleccione un Producto";
            Envase.Rows.Add(rowi);

            Envase.DefaultView.Sort = "Numero ASC";


            //Modulo de reportes*******************************
            DataRow rowii = Envase_Reporte.NewRow();
            rowii["Numero"] = "0";
            rowii["env_clave"] = "";
            rowii["env_nombre"] = "Seleccione un Envase";
            Envase_Reporte.Rows.Add(rowii);

            Envase_Reporte.DefaultView.Sort = "Numero ASC";


            thisConnecion.Close();

            int T = 0;
            if (Envase.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            AutoCompleteStringCollection stringCol = new AutoCompleteStringCollection();

            foreach (DataRow row in Envase.Rows)
            {
                stringCol.Add(Convert.ToString(row["env_nombre"]));
            }


            if (Envase_Reporte.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            AutoCompleteStringCollection stringCol2 = new AutoCompleteStringCollection();

            foreach (DataRow row in Envase_Reporte.Rows)
            {
                stringCol2.Add(Convert.ToString(row["env_nombre"]));
            }

            Cmb_Envase.DataSource = Envase;
            Cmb_Envase.DisplayMember = "env_nombre";
            Cmb_Envase.ValueMember = "env_clave";

            Cmb_Envase.AutoCompleteCustomSource = stringCol;
            Cmb_Envase.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            Cmb_Envase.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;

            cmb_envase_entr.DataSource = Envase;
            cmb_envase_entr.DisplayMember = "env_nombre";
            cmb_envase_entr.ValueMember = "env_clave";

            cmb_envase_entr.AutoCompleteCustomSource = stringCol;
            cmb_envase_entr.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmb_envase_entr.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;



            //Llenado combo envase del reporte

            cmbenvrep.DataSource = Envase_Reporte;
            cmbenvrep.DisplayMember = "env_nombre";
            cmbenvrep.ValueMember = "env_clave";

            cmbenvrep.AutoCompleteCustomSource = stringCol2;
            cmbenvrep.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            cmbenvrep.AutoCompleteSource = AutoCompleteSource.FileSystemDirectories;







            //Llenado de Combo Proveedor ********************************----*******************

        }


        private void Cmb_Producto_SelectedValueChanged_1(object sender, EventArgs e)
        {
            clbprod.Text = Cmb_Producto.SelectedValue.ToString();
            DataRow[] foundRows;
            foundRows = Producto.Select("prod_clave = '" + clbprod.Text + "'");
            for (int i = 0; i < foundRows.Length; i++)
            {
                clbenv.Text = foundRows[i][3].ToString();
                Cmb_Envase.SelectedValue = foundRows[i][3].ToString();
            }

        }

        private void Cmb_Envase_SelectedValueChanged(object sender, EventArgs e)
        {
            clbenv.Text = Cmb_Envase.SelectedValue.ToString();
        }

        private void clbprov_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmb_proveedor.SelectedValue = clbprov.Text.ToString();
        }

        private void clbrancho_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Cmb_Rancho.SelectedValue = clbrancho.Text.ToString();
        }


        private void clbprod_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Cmb_Producto.SelectedValue = clbprod.Text.ToString();
        }

        private void clbenv_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Cmb_Envase.SelectedValue = clbenv.Text.ToString();
        }

        private void clbtabla_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                Cmb_Tabla.SelectedValue = clbtabla.Text.ToString();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            clb_conse.ReadOnly = false;
            BtnSplitPed.Enabled = false;
            button2.Enabled = false;
            //button1.Enabled = true;
            //button5.Enabled = true;
            button6.Enabled = false;

            clb_conse.Text = "";
            clb_conse.Focus();

            //Restaurar la informacion de cada dato
            fechasal.Value = DateTime.Now;
            clbprov.Text = "";
            cmb_proveedor.SelectedValue = "";
            clbrancho.Text = "";
            Cmb_Rancho.SelectedIndex = 0;
            clbtabla.Text = "";
            Cmb_Tabla.SelectedValue = "";
            clbprod.Text = "";
            Cmb_Producto.SelectedValue = "";
            clbenv.Text = "";
            Cmb_Envase.SelectedValue = "";
            cant.Text = "";
            Chofer.Text = "";
            Operador.Text = "";
            ProductosGuardar.Clear();

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = ProductosGuardar;
            //******************************************

            //Deshabilitado de los modulos
            fechasal.Enabled = false;
            clbprov.ReadOnly = true;
            cmb_proveedor.Enabled = false;
            clbrancho.ReadOnly = true;
            Cmb_Rancho.Enabled = false;
            clbtabla.ReadOnly = true;
            Cmb_Tabla.Enabled = false;
            clbprod.ReadOnly = true;
            Cmb_Producto.Enabled = false;
            clbenv.ReadOnly = true;
            Cmb_Envase.Enabled = false;
            cant.ReadOnly = true;
            Chofer.ReadOnly = true;
            Operador.ReadOnly = true;

            #region VALIDA Y FORZA ACTUALIZAR
            if (ValidaActualizacion())
            {
                requiereActualizar = true;
                this.Close(); // Cerramos el formulario para dar paso al actualizador
            }
            #endregion

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Limpiar_Salida();


        }

        private void Limpiar_Salida()
        {
            clb_conse.ReadOnly = true;
            BtnSplitPed.Enabled = true;
            button1.Enabled = false;
            button2.Enabled = true;
            button5.Enabled = false;
            button6.Enabled = true;

            //Deshabilitado de los modulos
            fechasal.Enabled = true;
            clbprov.ReadOnly = false;
            cmb_proveedor.Enabled = true;
            clbrancho.ReadOnly = false;
            Cmb_Rancho.Enabled = true;
            clbtabla.ReadOnly = false;
            Cmb_Tabla.Enabled = true;
            clbprod.ReadOnly = false;
            Cmb_Producto.Enabled = true;
            clbenv.ReadOnly = false;
            Cmb_Envase.Enabled = true;
            cant.ReadOnly = false;
            Chofer.ReadOnly = false;
            Operador.ReadOnly = false;

            // Restaurado de Datos de datos
            traeultimoidsalida();

            fechasal.Value = DateTime.Now;
            clbprov.Text = "";
            cmb_proveedor.SelectedValue = "";
            clbrancho.Text = "";
            Cmb_Rancho.SelectedIndex = 0;
            clbtabla.Text = "";
            Cmb_Tabla.SelectedValue = "";
            clbprod.Text = "";
            Cmb_Producto.SelectedValue = "";
            clbenv.Text = "";
            Cmb_Envase.SelectedValue = "";
            cant.Text = "";
            Chofer.Text = "";
            Operador.Text = "";
            ProductosGuardar.Clear();

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = ProductosGuardar;
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            if (usuariostotales.Contains(usuario_actual) == false)
            {
                fechasal.Enabled = false;
                fecha_entrada.Enabled = false;
            }
        }

        private void clb_conse_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                button5.Enabled = true;
                thisConnecion.Open();
                string Cadena = "SELECT * FROM TB_SALIDAS_ENVASES WHERE FOLIO = '" + clb_conse.Text.Trim() + "'";
                DataSet ds1 = new DataSet();
                SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
                da1.Fill(ds1, "SALIDAENVASE");
                SqlCommand cmd;
                cmd = new SqlCommand(Cadena);
                cmd.Connection = thisConnecion;
                SqlDataReader Info;
                Info = cmd.ExecuteReader();
                DataTable salidaenvase = new DataTable();
                salidaenvase.Load(Info);
                thisConnecion.Close();

                if (salidaenvase.Rows.Count > 0)
                {
                    foreach (DataRow row in salidaenvase.Rows)
                    {
                        button1.Enabled = true;
                        //button1.Enabled = true;
                        //button5.Enabled = true;
                        button5.Enabled = true;

                        fechasal.Value = Convert.ToDateTime(row["FECHA"].ToString());
                        clbprov.Text = row["PROV_CLAVE"].ToString();
                        cmb_proveedor.SelectedValue = row["PROV_CLAVE"].ToString();
                        clbrancho.Text = row["RCH_CLAVE"].ToString();
                        Cmb_Rancho.SelectedValue = row["RCH_CLAVE"].ToString();
                        clbtabla.Text = row["TBL_CLAVE"].ToString();
                        Cmb_Tabla.SelectedValue = row["TBL_CLAVE"].ToString();
                        Chofer.Text = row["NOM_CHOFER"].ToString();
                        Operador.Text = row["NOM_OPERADOR"].ToString();

                        proveedor = row["PROV_CLAVE"].ToString().Trim() + " - " + row["PROV_NOMBRE"].ToString().Trim();
                        rancho = row["RCH_CLAVE"].ToString().Trim() + " - " + row["RCH_NOMBRE"].ToString().Trim();
                        tabla = row["TBL_CLAVE"].ToString().Trim() + " - " + row["TBL_NOMBRE"].ToString().Trim();

                        if (row["SAL_STATUS"].ToString().Trim() == "C")
                        {
                            button5.Enabled = false;
                            MessageBox.Show("El Folio Se Ha Cancelado");
                        }
                    }

                    llenadatatable();


                }
                else
                {
                    MessageBox.Show("El Folio No Existe");
                }

            }
            cmb_proveedor.SelectedValue = clbprov.Text.ToString();
        }

        private void llenadatatable()
        {
            ProductosGuardar.Clear();
            button5.Enabled = true;
            thisConnecion.Open();
            string Cadena = "SELECT * FROM  TB_DETSALIDAS_ENVASES WHERE FOLIO = '" + clb_conse.Text.Trim() + "'";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "SALIDAENVASEDETALLE");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            DataTable salidaenvasedetalle = new DataTable();
            salidaenvasedetalle.Load(Info);
            thisConnecion.Close();


            foreach (DataRow row in salidaenvasedetalle.Rows)
            {


                DataRow Contenido = ProductosGuardar.NewRow();
                Contenido["cvl_envase"] = row["ENV_CLAVE"].ToString().Trim();
                Contenido["nombre_envase"] = row["ENV_NOMBRE"].ToString().Trim();
                Contenido["Cantidad"] = Convert.ToInt32(row["CANTIDAD"].ToString().Trim());
                Contenido["nombre_producto"] = row["PROD_NOMBRE"].ToString().Trim();
                Contenido["cvl_producto"] = row["PROD_CLAVE"].ToString().Trim();
                ProductosGuardar.Rows.Add(Contenido);

            }
            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = ProductosGuardar;
        }


        private void llenadatatableEntrada()
        {
            ProductosGuardar_Entrada.Clear();
            cancelar_ent.Enabled = true;
            thisConnecion.Open();
            string Cadena = "SELECT * FROM  TB_DETENTRADAS_ENVASES WHERE FOLIO = '" + conse_entr.Text.Trim() + "'";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "ENTRADAENVASEDETALLE");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            DataTable entradaenvasedetalle = new DataTable();
            entradaenvasedetalle.Load(Info);
            thisConnecion.Close();


            foreach (DataRow row in entradaenvasedetalle.Rows)
            {


                DataRow Contenido = ProductosGuardar_Entrada.NewRow();
                Contenido["cvl_envase"] = row["ENV_CLAVE"].ToString().Trim();
                Contenido["nombre_envase"] = row["ENV_NOMBRE"].ToString().Trim();
                Contenido["Cantidad"] = Convert.ToInt32(row["CANTIDAD"].ToString().Trim());
                Contenido["nombre_producto"] = row["PROD_NOMBRE"].ToString().Trim();
                Contenido["cvl_producto"] = row["PROD_CLAVE"].ToString().Trim();
                ProductosGuardar_Entrada.Rows.Add(Contenido);

            }
            gridview_entr.AutoGenerateColumns = true;
            gridview_entr.DataSource = ProductosGuardar_Entrada;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            thisConnecion.Open();
            string cadena = "UPDATE TB_SALIDAS_ENVASES SET SAL_STATUS = 'C' WHERE FOLIO = '" + clb_conse.Text.Trim() + "'";
            SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
            cmd.ExecuteNonQuery();

            foreach (DataRow row in ProductosGuardar.Rows)
            {
                //row["Cantidad"] + "  |   " + row["cvl_envase"] + "  |  " + row["nombre_envase"] + row["cvl_producto"] + "  |  " + row["nombre_producto"]

                cadena = "UPDATE TB_MSTR_ENVASES SET cant_salidas = (cant_salidas - " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE prov_clave = '" + clbprov.Text.Trim() + "' AND rch_clave = '" + clbrancho.Text.Trim() +
                         "' AND tbl_clave = '" + clbtabla.Text.Trim().Trim() + "' AND env_clave = '" + row["cvl_envase"].ToString().Trim() + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES SET ENV_INV_CANT = (ENV_INV_CANT + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();
            }
            thisConnecion.Close();

            realizar_corte(fechasal.Text, "S", "CANCELACION");

            MessageBox.Show("El Folio Se Ha Cancelado Con Exito");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog1 = new PrintDialog();
            printDialog1.Document = printDocument1;
            DialogResult result = printDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                Imprimir_Salida();
            }
        }

        private void Imprimir_Salida()
        {
            float producto = 152.0F;
            String drawString1 = "\"Comercializadora GAB\"", drawString2 = "Salida de Envases\r\n" + DateTime.Now.ToString("dd/MM/yyyy   hh:mm tt"), drawString3 = "FOLIO: " + clb_conse.Text, drawString4 = "PROV: " + proveedor.Trim(), drawString5 = "RANCHO: " + rancho.Trim(), drawString6 = "TABLA: " + tabla.Trim();
            String envasesprod = "CAN    ENV    DESCRIPCION\r\n";
            foreach (DataRow row in ProductosGuardar.Rows)
            {
                envasesprod = envasesprod + row["Cantidad"] + "  |   " + row["cvl_envase"] + "  |  " + row["nombre_envase"] + "\r\n    ";
                envasesprod = envasesprod + row["cvl_producto"] + "  |  " + row["nombre_producto"] + "\r\n";
                producto = producto + 25;
            }
            String drawLinea = "Prueba7", drawProd = "Prueba8", drawPedi = "Prueba9", drawcajas = "Prueba10", drawtaraprox = "Prueba11";
            String drawstring8 = "Recibio Chofer: " + Chofer.Text.Trim(), drawstring9 = "Entrego Operador: " + Operador.Text.Trim(), drawstring10 = "* " + clb_conse.Text.Trim() + " *", drawstring11 = "__________________________", drawstring12 = "45";

            PrintDocument p = new PrintDocument();
            Font drawFont = new Font("Courier New", 8);//encabezado
            Font drawFont1 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont2 = new Font("Arial", 10, FontStyle.Bold);//encabezado
            Font drawFont3 = new Font("Arial", 10, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont4 = new Font("Arial", 9, FontStyle.Bold);//encabezado
            Font drawFont5 = new Font("Arial", 7, FontStyle.Bold);//encabezado
            Font drawFont8 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont9 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont10 = new Font("PF Barcode 39", 20);//encabezado
            Font drawFont11 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);//encabezado

            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Create point for upper-left corner of drawing. 
            PointF drawPoint1 = new PointF(40.0F, 15.0F);//codigo de barras            
            PointF drawPoint2 = new PointF(7.0f, 45.0f);//linea1
            PointF drawPoint3 = new PointF(7.0f, 78.0f);//datos1
            PointF drawpoint4 = new PointF(7.0F, 95.0F);
            PointF drawpoint5 = new PointF(7.0F, 108.0F);
            PointF drawpoint6 = new PointF(7.0F, 121.0F);
            PointF drawpoint7 = new PointF(7.0F, 140.0F);
            PointF drawpoint8 = new PointF(7.0F, producto);
            producto = producto + 25;
            PointF drawpoint9 = new PointF(7.0F, producto);
            producto = producto + 30;
            PointF drawpoint10 = new PointF(45.0f, producto);//linea2
            producto = producto + 30;
            PointF drawpoint11 = new PointF(7.0f, producto);//linea3
            p.PrintPage += delegate (object sender1, PrintPageEventArgs e1)
            {



                e1.Graphics.DrawString(drawString1, drawFont1, drawBrush, drawPoint1);//encabezado            
                e1.Graphics.DrawString(drawString2, drawFont2, drawBrush, drawPoint2);//linea1
                e1.Graphics.DrawString(drawString3, drawFont3, drawBrush, drawPoint3);//datos1
                e1.Graphics.DrawString(drawString4, drawFont4, drawBrush, drawpoint4);//
                e1.Graphics.DrawString(drawString5, drawFont4, drawBrush, drawpoint5);//
                e1.Graphics.DrawString(drawString6, drawFont4, drawBrush, drawpoint6);//
                e1.Graphics.DrawString(envasesprod, drawFont5, drawBrush, drawpoint7);//
                e1.Graphics.DrawString(drawstring8, drawFont8, drawBrush, drawpoint8);//
                e1.Graphics.DrawString(drawstring9, drawFont9, drawBrush, drawpoint9);//
                e1.Graphics.DrawString(drawstring10, drawFont10, drawBrush, drawpoint10);//
                e1.Graphics.DrawString(drawstring11, drawFont11, drawBrush, drawpoint11);//
                //e.Graphics.DrawString(drawString2, drawFont, drawBrush, drawPoint4);//linea2
                //e1.Graphics.DrawString(drawString2, drawFont3, drawBrush, drawPoint5);//linea3 
            };
            try
            {
                p.Print();
            }
            catch (Exception ex)
            {
                throw new Exception("Exception Occured While Printing", ex);
            }
        }

        private void Imprimir_Salida_Con_Folio(int folio)
        {
            float producto = 152.0F;

            // ✅ Usar el folio recibido como parámetro en lugar de clb_conse.Text
            String drawString1 = "\"Comercializadora GAB\"";
            String drawString2 = "Salida de Envases\r\n" + DateTime.Now.ToString("dd/MM/yyyy   hh:mm tt");
            String drawString3 = "FOLIO: " + folio.ToString();  // ← Cambio aquí
            String drawString4 = "PROV: " + proveedor.Trim();
            String drawString5 = "RANCHO: " + rancho.Trim();
            String drawString6 = "TABLA: " + tabla.Trim();

            String envasesprod = "CAN    ENV    DESCRIPCION\r\n";
            foreach (DataRow row in ProductosGuardar.Rows)
            {
                envasesprod = envasesprod + row["Cantidad"] + "  |   " + row["cvl_envase"] + "  |  " + row["nombre_envase"] + "\r\n    ";
                envasesprod = envasesprod + row["cvl_producto"] + "  |  " + row["nombre_producto"] + "\r\n";
                producto = producto + 25;
            }

            String drawLinea = "Prueba7", drawProd = "Prueba8", drawPedi = "Prueba9", drawcajas = "Prueba10", drawtaraprox = "Prueba11";
            String drawstring8 = "Recibio Chofer: " + Chofer.Text.Trim();
            String drawstring9 = "Entrego Operador: " + Operador.Text.Trim();
            String drawstring10 = "* " + folio.ToString() + " *";  // ← Cambio aquí
            String drawstring11 = "__________________________";
            String drawstring12 = "45";

            PrintDocument p = new PrintDocument();

            // Fuentes
            Font drawFont = new Font("Courier New", 8);
            Font drawFont1 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);
            Font drawFont2 = new Font("Arial", 10, FontStyle.Bold);
            Font drawFont3 = new Font("Arial", 10, FontStyle.Bold | FontStyle.Underline);
            Font drawFont4 = new Font("Arial", 9, FontStyle.Bold);
            Font drawFont5 = new Font("Arial", 7, FontStyle.Bold);
            Font drawFont8 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);
            Font drawFont9 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);
            Font drawFont10 = new Font("PF Barcode 39", 20);
            Font drawFont11 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);

            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Puntos de posición
            PointF drawPoint1 = new PointF(40.0F, 15.0F);
            PointF drawPoint2 = new PointF(7.0f, 45.0f);
            PointF drawPoint3 = new PointF(7.0f, 78.0f);
            PointF drawpoint4 = new PointF(7.0F, 95.0F);
            PointF drawpoint5 = new PointF(7.0F, 108.0F);
            PointF drawpoint6 = new PointF(7.0F, 121.0F);
            PointF drawpoint7 = new PointF(7.0F, 140.0F);
            PointF drawpoint8 = new PointF(7.0F, producto);
            producto = producto + 25;
            PointF drawpoint9 = new PointF(7.0F, producto);
            producto = producto + 30;
            PointF drawpoint10 = new PointF(45.0f, producto);
            producto = producto + 30;
            PointF drawpoint11 = new PointF(7.0f, producto);

            p.PrintPage += delegate (object sender1, PrintPageEventArgs e1)
            {
                e1.Graphics.DrawString(drawString1, drawFont1, drawBrush, drawPoint1);
                e1.Graphics.DrawString(drawString2, drawFont2, drawBrush, drawPoint2);
                e1.Graphics.DrawString(drawString3, drawFont3, drawBrush, drawPoint3);
                e1.Graphics.DrawString(drawString4, drawFont4, drawBrush, drawpoint4);
                e1.Graphics.DrawString(drawString5, drawFont4, drawBrush, drawpoint5);
                e1.Graphics.DrawString(drawString6, drawFont4, drawBrush, drawpoint6);
                e1.Graphics.DrawString(envasesprod, drawFont5, drawBrush, drawpoint7);
                e1.Graphics.DrawString(drawstring8, drawFont8, drawBrush, drawpoint8);
                e1.Graphics.DrawString(drawstring9, drawFont9, drawBrush, drawpoint9);
                e1.Graphics.DrawString(drawstring10, drawFont10, drawBrush, drawpoint10);
                e1.Graphics.DrawString(drawstring11, drawFont11, drawBrush, drawpoint11);
            };

            try
            {
                p.Print();
            }
            catch (Exception ex)
            {
                throw new Exception("Exception Occured While Printing", ex);
            }
        }

        public PrintDocument printDocument1 { get; set; }

        private void BtnSplitPedLEGACY_Click(object sender, EventArgs e)
        {
            thisConnecion.Open();
            string fecha = "SELECT SYSDATETIME()";
            SqlCommand cmdfechoy = new SqlCommand(fecha, thisConnecion);
            fecha_hoy_hoy = Convert.ToDateTime(cmdfechoy.ExecuteScalar()).ToString("dd/MM/yyyy");
            string FecSave = Convert.ToDateTime(cmdfechoy.ExecuteScalar()).ToString("dd/MM/yyyy HH:mm:ss");
            thisConnecion.Close();

            if (usuariostotales.Contains(usuario_actual) == false)
            {
                fechasal.Text = fecha_hoy_hoy;
            }


            //*** Validacion de campos necesarios para ingresar al datagritview
            if ((clbprov.Text.ToString().Trim().Length == 0) || (cmb_proveedor.Text.ToString().Trim().Length == 0))
            {
                MessageBox.Show("Seleccione o ingrese un Proveedor válido");
                return;
            }


            if (Chofer.Text.Trim().Length == 0)
            {
                MessageBox.Show("Ingrese un Nombre para el chofer");
                return;
            }

            if (Operador.Text.Trim().Length == 0)
            {
                MessageBox.Show("Ingrese un Nombre para el chofer");
                return;
            }


            if (dataGridView1.Rows.Count <= 0)
            {
                MessageBox.Show("Ingrese Al menos un Envase para La Salida");
                return;
            }
            if (ProductosGuardar.Rows.Count <= 0)
            {
                MessageBox.Show("Ingrese Al menos un Envase para La Salida");
                return;
            }

            string hora_actual = DateTime.Now.ToString("hh:mm:ss");

            string fecha_actual = Convert.ToDateTime(fechasal.Text).ToString("yyyy-dd-MM");

            string fecha_insert = fecha_actual + " " + hora_actual;



            thisConnecion.Open();

            proveedor = cmb_proveedor.Text.ToString().Trim();
            rancho = Cmb_Rancho.Text.ToString().Trim();
            tabla = Cmb_Tabla.Text.ToString().Trim();
            if (rancho == "Sin Ranchos Disponibles")
            {
                rancho = "";
            }

            if (tabla == "Sin Tablas Disponibles")
            {
                tabla = "";
            }

            string cadena = "insert into  TB_SALIDAS_ENVASES(FECHA, PROV_CLAVE, PROV_NOMBRE, RCH_CLAVE, RCH_NOMBRE, TBL_CLAVE, TBL_NOMBRE, NOM_CHOFER, NOM_OPERADOR, SAL_STATUS, FECHA_GUARDADO) " +
                                "Values('" + Convert.ToDateTime(fechasal.Text).ToString("dd/MM/yyyy") + "','" + clbprov.Text.Trim() + "','" + cmb_proveedor.Text.ToString() +
                                "','" + clbrancho.Text.Trim() + "','" + rancho.ToString().Trim() + "','" + clbtabla.Text.Trim().Trim() + "', '" + tabla.ToString() +
                                "', '" + Chofer.Text.Trim() + "', '" + Operador.Text.Trim() + "', 'T','" + FecSave + "')";
            SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
            cmd.ExecuteNonQuery();


            foreach (DataRow row in ProductosGuardar.Rows)
            {
                //row["Cantidad"] + "  |   " + row["cvl_envase"] + "  |  " + row["nombre_envase"] + row["cvl_producto"] + "  |  " + row["nombre_producto"]
                cadena = "insert into  TB_DETSALIDAS_ENVASES(FOLIO, ENV_CLAVE, ENV_NOMBRE, CANTIDAD, PROD_CLAVE, PROD_NOMBRE)" +
                                "Values('" + clb_conse.Text.Trim() + "','" + row["cvl_envase"].ToString().Trim() + "','" + row["nombre_envase"].ToString().Trim() +
                                "','" + row["Cantidad"].ToString().Trim() + "','" + row["cvl_producto"].ToString().Trim() + "','" + row["nombre_producto"].ToString().Trim() + "')";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();


                cadena = "UPDATE TB_MSTR_ENVASES SET cant_salidas = (cant_salidas + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE prov_clave = '" + clbprov.Text.Trim() + "' AND rch_clave = '" + clbrancho.Text.Trim() +
                         "' AND tbl_clave = '" + clbtabla.Text.Trim().Trim() + "' AND env_clave = '" + row["cvl_envase"].ToString().Trim() + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES SET ENV_INV_CANT = (ENV_INV_CANT - " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES_dos SET ENV_SAL_CANT = (ENV_SAL_CANT + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' and ENV_FECHA = '" + Convert.ToDateTime(fechasal.Text).ToString("dd/MM/yyyy") + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES_sin_corte SET ENV_SAL_CANT = (ENV_SAL_CANT + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' and ENV_FECHA = '" + Convert.ToDateTime(fechasal.Text).ToString("dd/MM/yyyy") + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                if (row["cvl_envase"].ToString().Trim() == "81")
                {
                    validarproveedoresparrago(clbprov.Text.Trim(), Convert.ToDateTime(fechasal.Text).ToString("dd/MM/yyyy"));
                }
            }
            thisConnecion.Close();
            MessageBox.Show("La Salida se ha almacenado Con Exito");

            if (Convert.ToDateTime(fecha_hoy_hoy) > Convert.ToDateTime(fechasal.Text))
            {
                thisConnecion.Open();
                string Cadena = "SELECT TOP (1) FOLIO FROM TB_SALIDAS_ENVASES ORDER BY FOLIO DESC";
                cmd = new SqlCommand(Cadena, thisConnecion);
                int Valor = Convert.ToInt32(cmd.ExecuteScalar());

                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values('" + fecha_hoy_hoy + "','" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ATRASADO','2.18','" + Valor + "', 'INSERCCION DE SALIDA A DESTIEMPO FECHA " + fechasal.Text + "', 'SISGAB', '" + Valor + "')";
                SqlCommand cmdent = new SqlCommand(cadenaregmov, thisConnecion);
                cmdent.ExecuteNonQuery();
                thisConnecion.Close();

                realizar_corte(Convert.ToDateTime(fechasal.Text).ToString("dd/MM/yyyy"), "S", "DESTIEMPO");
            }
            else
            {
                thisConnecion.Open();
                string Cadena = "SELECT TOP (1) FOLIO FROM TB_SALIDAS_ENVASES ORDER BY FOLIO DESC";
                cmd = new SqlCommand(Cadena, thisConnecion);
                int Valor = Convert.ToInt32(cmd.ExecuteScalar());

                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values(GetDate(),'" + Environment.MachineName.Trim() + "','" + usuario_actual + "','SALIDA','2.18','" + Valor + "', 'INSERCCION DE SALIDA FECHA " + fechasal.Text + "', 'SISGAB', '" + Valor + "')";
                SqlCommand cmdent = new SqlCommand(cadenaregmov, thisConnecion);
                cmdent.ExecuteNonQuery();
                thisConnecion.Close();
            }


            PrintDialog printDialog1 = new PrintDialog();
            printDialog1.Document = printDocument1;
            DialogResult result = printDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                Imprimir_Salida();
            }

            Limpiar_Salida();
        }

        private void BtnSplitPed_Click(object sender, EventArgs e)
        {
            string FecSave;
            // ✅ 1. Obtener fecha del servidor (mantener tu lógica original)
            try
            {
                if (thisConnecion.State != ConnectionState.Open) thisConnecion.Open();
                string fecha = "SELECT SYSDATETIME()";
                SqlCommand cmdfechoy = new SqlCommand(fecha, thisConnecion);
                fecha_hoy_hoy = Convert.ToDateTime(cmdfechoy.ExecuteScalar()).ToString("dd/MM/yyyy");
                FecSave = Convert.ToDateTime(cmdfechoy.ExecuteScalar()).ToString("dd/MM/yyyy HH:mm:ss");
                thisConnecion.Close();

                #region VALIDA Y FORZA ACTUALIZAR
                if (ValidaActualizacion())
                {
                    requiereActualizar = true;
                    this.Close(); // Cerramos el formulario para dar paso al actualizador
                }
                #endregion
            }
            catch (Exception ex)
            {
                thisConnecion.Close();
                MessageBox.Show($"Error al obtener fecha del servidor: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // ✅ 2. Validación de fecha para usuarios (mantener tu lógica original)
            if (usuariostotales.Contains(usuario_actual) == false)
            {
                fechasal.Text = fecha_hoy_hoy;
            }

            // ✅ 3. Validaciones de campos (mantener tus validaciones originales)
            if ((clbprov.Text.ToString().Trim().Length == 0) || (cmb_proveedor.Text.ToString().Trim().Length == 0))
            {
                MessageBox.Show("Seleccione o ingrese un Proveedor válido");
                return;
            }

            if (Chofer.Text.Trim().Length == 0)
            {
                MessageBox.Show("Ingrese un Nombre para el chofer");
                return;
            }

            if (Operador.Text.Trim().Length == 0)
            {
                MessageBox.Show("Ingrese un Nombre para el operador");
                return;
            }

            if (dataGridView1.Rows.Count <= 0)
            {
                MessageBox.Show("Ingrese Al menos un Envase para La Salida");
                return;
            }

            if (ProductosGuardar.Rows.Count <= 0)
            {
                MessageBox.Show("Ingrese Al menos un Envase para La Salida");
                return;
            }

            // ✅ 4. Preparar variables necesarias
            proveedor = cmb_proveedor.Text.ToString().Trim();
            rancho = Cmb_Rancho.Text.ToString().Trim();
            tabla = Cmb_Tabla.Text.ToString().Trim();

            if (rancho == "Sin Ranchos Disponibles")
            {
                rancho = "";
            }

            if (tabla == "Sin Tablas Disponibles")
            {
                tabla = "";
            }

            // ✅ 5. Preparar conexión y transacción
            if (thisConnecion.State != ConnectionState.Open) thisConnecion.Open();
            SqlTransaction transaction = thisConnecion.BeginTransaction();

            try
            {
                int folioGenerado = 0;
                SqlCommand cmd;

                // ✅ 6. Insertar cabecera Y capturar el folio con SCOPE_IDENTITY()
                string insertHeader = @"INSERT INTO TB_SALIDAS_ENVASES(
    FECHA, PROV_CLAVE, PROV_NOMBRE, RCH_CLAVE, RCH_NOMBRE, 
    TBL_CLAVE, TBL_NOMBRE, NOM_CHOFER, NOM_OPERADOR, SAL_STATUS, FECHA_GUARDADO) 
    VALUES(@Fecha, @ProvClave, @ProvNombre, @RchClave, @RchNombre, 
           @TblClave, @TblNombre, @Chofer, @Operador, 'T', @FechaGuardado);
    SELECT CAST(SCOPE_IDENTITY() AS INT);";

                cmd = new SqlCommand(insertHeader, thisConnecion, transaction);
                cmd.Parameters.AddWithValue("@Fecha", Convert.ToDateTime(fechasal.Text));
                cmd.Parameters.AddWithValue("@ProvClave", clbprov.Text.Trim());
                cmd.Parameters.AddWithValue("@ProvNombre", cmb_proveedor.Text.Trim());
                cmd.Parameters.AddWithValue("@RchClave", clbrancho.Text.Trim());
                cmd.Parameters.AddWithValue("@RchNombre", rancho.Trim());
                cmd.Parameters.AddWithValue("@TblClave", clbtabla.Text.Trim());
                cmd.Parameters.AddWithValue("@TblNombre", tabla.Trim());
                cmd.Parameters.AddWithValue("@Chofer", Chofer.Text.Trim());
                cmd.Parameters.AddWithValue("@Operador", Operador.Text.Trim());
                cmd.Parameters.AddWithValue("@FechaGuardado", FecSave);

                // ✅ 7. ¡CAPTURAR EL FOLIO REAL GENERADO POR LA BD!
                folioGenerado = Convert.ToInt32(cmd.ExecuteScalar());

                // ✅ 8. Insertar detalles y actualizar inventarios (todo con parámetros y transacción)
                foreach (DataRow row in ProductosGuardar.Rows)
                {
                    int cantidad = Convert.ToInt32(row["Cantidad"].ToString().Trim());
                    string envClave = row["cvl_envase"].ToString().Trim();
                    string envNombre = row["nombre_envase"].ToString().Trim();
                    string prodClave = row["cvl_producto"].ToString().Trim();
                    string prodNombre = row["nombre_producto"].ToString().Trim();
                    DateTime fechaSalida = Convert.ToDateTime(fechasal.Text);

                    // --- Insertar detalle ---
                    string insertDetalle = @"INSERT INTO TB_DETSALIDAS_ENVASES(
        FOLIO, ENV_CLAVE, ENV_NOMBRE, CANTIDAD, PROD_CLAVE, PROD_NOMBRE) 
        VALUES(@Folio, @EnvClave, @EnvNombre, @Cantidad, @ProdClave, @ProdNombre)";

                    // CORREGIDO: Se agregó 'transaction' al constructor
                    cmd = new SqlCommand(insertDetalle, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@Folio", folioGenerado);
                    cmd.Parameters.AddWithValue("@EnvClave", envClave);
                    cmd.Parameters.AddWithValue("@EnvNombre", envNombre);
                    cmd.Parameters.AddWithValue("@Cantidad", cantidad);
                    cmd.Parameters.AddWithValue("@ProdClave", prodClave);
                    cmd.Parameters.AddWithValue("@ProdNombre", prodNombre);
                    cmd.ExecuteNonQuery();

                    // --- UPDATE 1: TB_MSTR_ENVASES ---
                    string updateMstr = @"UPDATE TB_MSTR_ENVASES 
        SET cant_salidas = (cant_salidas + @Cantidad) 
        WHERE prov_clave = @ProvClave 
          AND rch_clave = @RchClave 
          AND tbl_clave = @TblClave 
          AND env_clave = @EnvClave";

                    // CORREGIDO: Se agregó 'transaction' al constructor
                    cmd = new SqlCommand(updateMstr, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@Cantidad", cantidad);
                    cmd.Parameters.AddWithValue("@ProvClave", clbprov.Text.Trim());
                    cmd.Parameters.AddWithValue("@RchClave", clbrancho.Text.Trim());
                    cmd.Parameters.AddWithValue("@TblClave", clbtabla.Text.Trim());
                    cmd.Parameters.AddWithValue("@EnvClave", envClave);
                    cmd.ExecuteNonQuery();

                    // --- UPDATE 2: TB_MSTR_INV_ENVASES ---
                    string updateInv = @"UPDATE TB_MSTR_INV_ENVASES 
        SET ENV_INV_CANT = (ENV_INV_CANT - @Cantidad) 
        WHERE ENV_CLAVE = @EnvClave";

                    // CORREGIDO: Se agregó 'transaction' al constructor
                    cmd = new SqlCommand(updateInv, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@Cantidad", cantidad);
                    cmd.Parameters.AddWithValue("@EnvClave", envClave);
                    cmd.ExecuteNonQuery();

                    // --- UPDATE 3: TB_MSTR_INV_ENVASES_dos ---
                    string updateInvDos = @"UPDATE TB_MSTR_INV_ENVASES_dos 
        SET ENV_SAL_CANT = (ENV_SAL_CANT + @Cantidad) 
        WHERE ENV_CLAVE = @EnvClave 
          AND ENV_FECHA = @FechaEnv";

                    // CORREGIDO: Se agregó 'transaction' al constructor
                    cmd = new SqlCommand(updateInvDos, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@Cantidad", cantidad);
                    cmd.Parameters.AddWithValue("@EnvClave", envClave);
                    cmd.Parameters.AddWithValue("@FechaEnv", fechaSalida);
                    cmd.ExecuteNonQuery();

                    // --- UPDATE 4: TB_MSTR_INV_ENVASES_sin_corte ---
                    string updateInvSinCorte = @"UPDATE TB_MSTR_INV_ENVASES_sin_corte 
        SET ENV_SAL_CANT = (ENV_SAL_CANT + @Cantidad) 
        WHERE ENV_CLAVE = @EnvClave 
          AND ENV_FECHA = @FechaEnv";

                    // CORREGIDO: Se agregó 'transaction' al constructor
                    cmd = new SqlCommand(updateInvSinCorte, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@Cantidad", cantidad);
                    cmd.Parameters.AddWithValue("@EnvClave", envClave);
                    cmd.Parameters.AddWithValue("@FechaEnv", fechaSalida);
                    cmd.ExecuteNonQuery();

                    // --- Validación especial para envase 81 ---
                    if (envClave == "81")
                    {
                        // NOTA: Asegúrate de que este método interno use la misma conexión/transacción si realiza escrituras.
                        validarproveedoresparragoT(clbprov.Text.Trim(), fechaSalida.ToString("dd/MM/yyyy"), transaction);
                    }
                }

                // ✅ 9. Registrar movimiento USANDO el folio capturado
                string tipoMov = Convert.ToDateTime(fecha_hoy_hoy) > Convert.ToDateTime(fechasal.Text) ? "ATRASADO" : "SALIDA";
                string detalleMov = Convert.ToDateTime(fecha_hoy_hoy) > Convert.ToDateTime(fechasal.Text)
                    ? $"INSERCCION DE SALIDA A DESTIEMPO FECHA {fechasal.Text}"
                    : $"INSERCCION DE SALIDA FECHA {fechasal.Text}";

                string insertLog = @"INSERT INTO tb_registro_movimientos 
    (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) 
    VALUES(@FechaLog, @Maquina, @Usuario, @TipoMov, '2.18', @Folio, @Detalle, 'SISGAB', @Folio)";

                // CORREGIDO: Se agregó 'transaction' al constructor
                cmd = new SqlCommand(insertLog, thisConnecion, transaction);
                cmd.Parameters.AddWithValue("@FechaLog", DateTime.Now);
                cmd.Parameters.AddWithValue("@Maquina", Environment.MachineName.Trim());
                cmd.Parameters.AddWithValue("@Usuario", usuario_actual);
                cmd.Parameters.AddWithValue("@TipoMov", tipoMov);
                cmd.Parameters.AddWithValue("@Folio", folioGenerado);
                cmd.Parameters.AddWithValue("@Detalle", detalleMov);
                cmd.ExecuteNonQuery();

                // ✅ 10. Confirmar TODA la transacción
                transaction.Commit();
                thisConnecion.Close();

                // ✅ 11. Mensaje de éxito
                MessageBox.Show("La Salida se ha almacenado Con Exito");

                // ✅ 12. Si es atrasado, realizar corte (fuera de la transacción principal)
                if (Convert.ToDateTime(fecha_hoy_hoy) > Convert.ToDateTime(fechasal.Text))
                {
                    realizar_corte(Convert.ToDateTime(fechasal.Text).ToString("dd/MM/yyyy"), "S", "DESTIEMPO");
                }

                // ✅ 13. Imprimir USANDO el folio capturado
                PrintDialog printDialog1 = new PrintDialog();
                printDialog1.Document = printDocument1;
                DialogResult result = printDialog1.ShowDialog();
                if (result == DialogResult.OK)
                {
                    Imprimir_Salida_Con_Folio(folioGenerado);
                }

                // ✅ 14. Limpiar formulario
                Limpiar_Salida();
            }
            catch (Exception ex)
            {
                // ✅ 15. Revertir todo si algo falla
                try
                {
                    transaction?.Rollback();
                }
                catch { /* Ignorar si la transacción ya se cerró o no se inició */ }

                thisConnecion.Close();
                MessageBox.Show($"Error al guardar la salida: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // ✅ 16. Asegurar que la conexión se cierre siempre
                if (thisConnecion.State == ConnectionState.Open) thisConnecion.Close();
            }
        }

        public void realizar_corte(string fecha_mov, string tipo, string mov)
        {
            DataGridView x = new DataGridView();
            if (tipo == "S")
            {
                x = dataGridView1;
            }
            else
            {
                x = gridview_entr;
            }

            string fecha_edit = Convert.ToDateTime(fecha_mov).ToString("dd/MM/yyyy");

            TimeSpan tspan = Convert.ToDateTime(fecha_hoy_hoy) - Convert.ToDateTime(fecha_mov);

            int dias = tspan.Days;

            dias++;
            int veces = 0;


            thisConnecion.Open();
            string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                            "Values('" + fecha_hoy_hoy + "','" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ACTUALIZAR','2.18','', 'MODIFICACION INVENTARIO INICIAL POR " + mov + " FECHA " + fecha_edit.Trim() + "', 'SIPGAB', '')";
            SqlCommand cmdxs = new SqlCommand(cadenaregmov, thisConnecion);
            cmdxs.ExecuteNonQuery();
            thisConnecion.Close();


            bool ciclo = true;

            while (ciclo == true)
            {
                if (Convert.ToDateTime(fecha_edit) > Convert.ToDateTime(fecha_hoy_hoy))
                {
                    ciclo = false;
                }
                else
                {
                    foreach (DataGridViewRow row in x.Rows)
                    {
                        if (row.Cells["cvl_envase"].Value != null)
                        {
                            thisConnecion.Open();
                            string idenvase = row.Cells["cvl_envase"].Value.ToString();
                            string nomenvase = row.Cells["nombre_envase"].Value.ToString();
                            string cantenvase = row.Cells["cantidad"].Value.ToString();

                            string valida = "SELECT ISNULL(sum(ENV_INV_INI_CANT), 0) FROM TB_MSTR_INV_ENVASES_dos WHERE ENV_CLAVE = '" + idenvase + "' AND ENV_FECHA = '" + fecha_edit + "'";
                            SqlCommand cmd = new SqlCommand(valida, thisConnecion);
                            int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND ENT_STATUS != 'C'";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND SAL_STATUS != 'C'";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            string actualizadato = "";


                            //thisConnecion.Open();
                            actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "'";
                            cmd = new SqlCommand(actualizadato, thisConnecion);
                            cmd.ExecuteNonQuery();
                            //thisConnecion.Close();

                            actualizadato = "UPDATE TB_MSTR_INV_ENVASES_sin_corte SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "'";
                            cmd = new SqlCommand(actualizadato, thisConnecion);
                            cmd.ExecuteNonQuery();

                            int inv_ini_hoy = inv_ini + entradas - salidas;

                            //thisConnecion.Open();
                            actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_INV_INI_CANT = '" + inv_ini_hoy + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy") + "' ";
                            cmd = new SqlCommand(actualizadato, thisConnecion);
                            cmd.ExecuteNonQuery();
                            thisConnecion.Close();

                        }
                    }
                    fecha_edit = Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy");
                }
            }

        }

        private void cmb_producto_entr_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmb_producto_entr.SelectedValue == null)
            {
                foreach (DataRow row in Producto.Rows)
                {
                    if (row["prod_nombre"] == cmb_prov_entr.Text)
                    {
                        clb_producto_entr.Text = cmb_producto_entr.SelectedValue.ToString();
                        DataRow[] foundRows;
                        foundRows = Producto.Select("prod_clave = '" + clb_producto_entr.Text + "'");
                        for (int i = 0; i < foundRows.Length; i++)
                        {
                            clb_envase_entr.Text = foundRows[i][3].ToString();
                            cmb_envase_entr.SelectedValue = foundRows[i][3].ToString();
                        }
                    }
                }
            }
            else
            {
                clb_producto_entr.Text = cmb_producto_entr.SelectedValue.ToString();
            }
        }

        private void cmb_prov_entr_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmb_prov_entr.SelectedValue == null)
            {
                foreach (DataRow row in Proveedor.Rows)
                {
                    if (row["prov_nombre"] == cmb_prov_entr.Text)
                    {
                        clb_prov_entr.Text = cmb_prov_entr.ValueMember.ToString();
                    }
                }
            }
            else
            {
                clb_prov_entr.Text = cmb_prov_entr.SelectedValue.ToString();
            }


            comboRancho("Entrada");
        }

        private void cmb_rancho_entr_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmb_rancho_entr.SelectedValue == null)
            {
                clb_rancho_entr.Text = "";
                MessageBox.Show("El Rancho no Existe");

            }
            else
            {
                clb_rancho_entr.Text = cmb_rancho_entr.SelectedValue.ToString();
                comboTabla("Entrada");
            }

        }

        private void clb_prov_entr_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmb_prov_entr.SelectedValue = clb_prov_entr.Text.ToString();
        }

        private void agregar_ent_Click(object sender, EventArgs e)
        {
            //*** Validacion de campos necesarios para ingresar al datagritview
            if ((clb_envase_entr.Text.ToString().Trim().Length == 0) || (cmb_envase_entr.Text.ToString().Trim().Length == 0))
            {
                MessageBox.Show("Seleccione o ingrese un Envase válido");
                return;
            }



            if (cantidad_entr.Text.Trim().Length == 0 || (Convert.ToInt32(cantidad_entr.Text.Trim()) <= 0))
            {
                MessageBox.Show("La cantidad es incorrecta, Ingrese una cantidad válida");
                return;
            }

            Boolean Existe = false;
            foreach (DataRow row in ProductosGuardar_Entrada.Select("Cvl_Envase = '" + clbenv.Text + "' and cvl_producto = '" + clbprod.Text + "'"))
                Existe = true;

            if (Existe)
            {
                MessageBox.Show("El Envase y el Producto ya fue ingresado, NO se Puede Agregar 2 Veces", "Ingreso de Informacion", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            DataRow Contenido = ProductosGuardar_Entrada.NewRow();
            Contenido["cvl_envase"] = clb_envase_entr.Text.ToString().Trim();
            Contenido["nombre_envase"] = cmb_envase_entr.Text.ToString().Trim().Replace(" - " + clb_envase_entr.Text.ToString().Trim(), ""); ;
            Contenido["Cantidad"] = Convert.ToInt32(cantidad_entr.Text);

            if ((clbprod.Text.ToString().Trim().Length == 0) || (Cmb_Producto.Text.ToString().Trim().Length == 0))
            {
                Contenido["nombre_producto"] = "";
                Contenido["cvl_producto"] = "";
            }
            else
            {
                Contenido["nombre_producto"] = cmb_producto_entr.Text.ToString().Trim().Replace(" - " + clb_producto_entr.Text.ToString().Trim(), "");
                Contenido["cvl_producto"] = clb_producto_entr.Text.ToString().Trim();
            }

            ProductosGuardar_Entrada.Rows.Add(Contenido);

            gridview_entr.AutoGenerateColumns = true;
            gridview_entr.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            gridview_entr.DataSource = ProductosGuardar_Entrada;

            clb_producto_entr.Text = "";
            cmb_producto_entr.SelectedValue = "";
            clb_envase_entr.Text = "";
            cmb_envase_entr.SelectedValue = "";
            cantidad_entr.Text = "";

        }

        private void clb_rancho_entr_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmb_rancho_entr.SelectedValue = clb_rancho_entr.Text.ToString();
        }

        private void clb_tabla_entr_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmb_tabla_entr.SelectedValue = clb_tabla_entr.Text.ToString();
        }

        private void clb_producto_entr_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmb_producto_entr.SelectedValue = clb_producto_entr.Text.ToString();
        }

        private void clb_envase_entr_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmb_envase_entr.SelectedValue = clb_envase_entr.Text.ToString();
        }

        private void buscar_ent_Click(object sender, EventArgs e)
        {
            conse_entr.ReadOnly = false;
            guardar_ent.Enabled = false;
            buscar_ent.Enabled = false;
            //button1.Enabled = true;
            //button5.Enabled = true;
            agregar_ent.Enabled = false;

            conse_entr.Text = "";
            conse_entr.Focus();

            //Restaurar la informacion de cada dato
            fecha_entrada.Value = DateTime.Now;
            clb_prov_entr.Text = "";
            cmb_prov_entr.SelectedValue = "";
            clb_rancho_entr.Text = "";
            cmb_rancho_entr.SelectedIndex = 0;
            clb_tabla_entr.Text = "";
            cmb_tabla_entr.SelectedValue = "";
            clb_producto_entr.Text = "";
            cmb_producto_entr.SelectedValue = "";
            clb_envase_entr.Text = "";
            cmb_envase_entr.SelectedValue = "";
            cantidad_entr.Text = "";
            chofer_ent.Text = "";
            operador_ent.Text = "";
            ProductosGuardar_Entrada.Clear();

            gridview_entr.AutoGenerateColumns = true;
            gridview_entr.DataSource = ProductosGuardar_Entrada;
            //******************************************

            //Deshabilitado de los modulos
            fecha_entrada.Enabled = false;
            clb_prov_entr.ReadOnly = true;
            cmb_prov_entr.Enabled = false;
            clb_rancho_entr.ReadOnly = true;
            cmb_rancho_entr.Enabled = false;
            clb_tabla_entr.ReadOnly = true;
            cmb_tabla_entr.Enabled = false;
            clb_producto_entr.ReadOnly = true;
            cmb_producto_entr.Enabled = false;
            clb_envase_entr.ReadOnly = true;
            cmb_envase_entr.Enabled = false;
            cantidad_entr.ReadOnly = true;
            chofer_ent.ReadOnly = true;
            operador_ent.ReadOnly = true;
        }

        private void cancelar_proceso_ent_Click(object sender, EventArgs e)
        {
            Limpiar_Entrada();
        }

        private void Limpiar_Entrada()
        {
            conse_entr.ReadOnly = true;
            guardar_ent.Enabled = true;
            imprimir_ent.Enabled = false;
            buscar_ent.Enabled = true;
            cancelar_ent.Enabled = false;
            agregar_ent.Enabled = true;

            //Deshabilitado de los modulos
            fecha_entrada.Enabled = true;
            clb_prov_entr.ReadOnly = false;
            cmb_prov_entr.Enabled = true;
            clb_rancho_entr.ReadOnly = false;
            cmb_rancho_entr.Enabled = true;
            clb_tabla_entr.ReadOnly = false;
            cmb_tabla_entr.Enabled = true;
            clb_producto_entr.ReadOnly = false;
            cmb_producto_entr.Enabled = true;
            clb_envase_entr.ReadOnly = false;
            cmb_envase_entr.Enabled = true;
            cantidad_entr.ReadOnly = false;
            chofer_ent.ReadOnly = false;
            operador_ent.ReadOnly = false;

            // Restaurado de Datos de datos
            traeultimoidsalida();

            fecha_entrada.Value = DateTime.Now;
            clb_prov_entr.Text = "";
            cmb_prov_entr.SelectedValue = "";
            clb_rancho_entr.Text = "";
            cmb_rancho_entr.SelectedIndex = 0;
            clb_tabla_entr.Text = "";
            cmb_tabla_entr.SelectedValue = "";
            clb_producto_entr.Text = "";
            cmb_producto_entr.SelectedValue = "";
            clb_envase_entr.Text = "";
            cmb_envase_entr.SelectedValue = "";
            cantidad_entr.Text = "";
            chofer_ent.Text = "";
            operador_ent.Text = "";
            ProductosGuardar_Entrada.Clear();

            gridview_entr.AutoGenerateColumns = true;
            gridview_entr.DataSource = ProductosGuardar_Entrada;
            gridview_entr.DataSource = null;

            if (usuariostotales.Contains(usuario_actual) == false)
            {
                fechasal.Enabled = false;
                fecha_entrada.Enabled = false;
            }

        }

        private void cmb_envase_entr_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmb_envase_entr.SelectedValue == null)
            {
                foreach (DataRow row in Envase.Rows)
                {
                    if (row["env_nombre"] == clb_envase_entr.Text)
                    {
                        clb_envase_entr.Text = cmb_envase_entr.ValueMember.ToString();
                    }
                }
            }
            else
            {
                clb_envase_entr.Text = cmb_envase_entr.SelectedValue.ToString();
            }

        }

        private void cmb_tabla_entr_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmb_tabla_entr.SelectedValue == null)
            {
                clb_tabla_entr.Text = "";
                MessageBox.Show("La Tabla No Existe");
            }
            else
            {
                clb_tabla_entr.Text = cmb_tabla_entr.SelectedValue.ToString();
            }

        }

        private void conse_entr_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                cancelar_ent.Enabled = true;
                thisConnecion.Open();
                string Cadena = "SELECT * FROM TB_ENTRADAS_ENVASES WHERE FOLIO = '" + conse_entr.Text.Trim() + "'";
                DataSet ds1 = new DataSet();
                SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
                da1.Fill(ds1, "ENTRADAENVASE");
                SqlCommand cmd;
                cmd = new SqlCommand(Cadena);
                cmd.Connection = thisConnecion;
                SqlDataReader Info;
                Info = cmd.ExecuteReader();
                DataTable entradaenvase = new DataTable();
                entradaenvase.Load(Info);
                thisConnecion.Close();

                if (entradaenvase.Rows.Count > 0)
                {
                    foreach (DataRow row in entradaenvase.Rows)
                    {
                        imprimir_ent.Enabled = true;
                        //button1.Enabled = true;
                        //button5.Enabled = true;
                        cancelar_ent.Enabled = true;

                        fecha_entrada.Value = Convert.ToDateTime(row["FECHA"].ToString());
                        clb_prov_entr.Text = row["PROV_CLAVE"].ToString();
                        cmb_prov_entr.SelectedValue = row["PROV_CLAVE"].ToString();
                        clb_rancho_entr.Text = row["RCH_CLAVE"].ToString();
                        cmb_rancho_entr.SelectedValue = row["RCH_CLAVE"].ToString();
                        clb_tabla_entr.Text = row["TBL_CLAVE"].ToString();
                        cmb_tabla_entr.SelectedValue = row["TBL_CLAVE"].ToString();
                        chofer_ent.Text = row["NOM_CHOFER"].ToString();
                        operador_ent.Text = row["NOM_OPERADOR"].ToString();

                        proveedor = row["PROV_CLAVE"].ToString().Trim() + " - " + row["PROV_NOMBRE"].ToString().Trim();
                        rancho = row["RCH_CLAVE"].ToString().Trim() + " - " + row["RCH_NOMBRE"].ToString().Trim();
                        tabla = row["TBL_CLAVE"].ToString().Trim() + " - " + row["TBL_NOMBRE"].ToString().Trim();

                        if (row["ENT_STATUS"].ToString().Trim() == "C")
                        {
                            button5.Enabled = false;
                            MessageBox.Show("El Folio Se Ha Cancelado");
                        }
                    }

                    llenadatatableEntrada();


                }
                else
                {
                    MessageBox.Show("El Folio No Existe");
                }

            }
            cmb_proveedor.SelectedValue = clbprov.Text.ToString();
        }

        private void cancelar_ent_Click(object sender, EventArgs e)
        {
            thisConnecion.Open();
            string cadena = "UPDATE TB_ENTRADAS_ENVASES SET ENT_STATUS = 'C' WHERE FOLIO = '" + conse_entr.Text.Trim() + "'";
            SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
            cmd.ExecuteNonQuery();
            thisConnecion.Close();

            realizar_corte(fecha_entrada.Text, "E", "CANCELACION");
            MessageBox.Show("El Folio Se Ha Cancelado Con Exito");
        }

        private void imprimir_ent_Click(object sender, EventArgs e)
        {
            PrintDialog printDialog1 = new PrintDialog();
            printDialog1.Document = printDocument1;
            DialogResult result = printDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                Imprimir_Entrada();
            }
        }

        private void Imprimir_Entrada()
        {
            float producto = 152.0F;
            String drawString1 = "\"Comercializadora GAB\"", drawString2 = "Entrada de Envases\r\n" + DateTime.Now.ToString("dd/MM/yyyy   hh:mm tt"), drawString3 = "FOLIO: " + conse_entr.Text, drawString4 = "PROV: " + proveedor.Trim(), drawString5 = "RANCHO: " + rancho.Trim(), drawString6 = "TABLA: " + tabla.Trim();
            String envasesprod = "CAN    ENV    DESCRIPCION\r\n";
            foreach (DataRow row in ProductosGuardar_Entrada.Rows)
            {
                envasesprod = envasesprod + row["Cantidad"] + "  |   " + row["cvl_envase"] + "  |  " + row["nombre_envase"] + "\r\n    ";
                envasesprod = envasesprod + row["cvl_producto"] + "  |  " + row["nombre_producto"] + "\r\n";
                producto = producto + 25;
            }
            String drawLinea = "Prueba7", drawProd = "Prueba8", drawPedi = "Prueba9", drawcajas = "Prueba10", drawtaraprox = "Prueba11";
            String drawstring8 = "Recibio Chofer: " + chofer_ent.Text.Trim(), drawstring9 = "Entrego Operador: " + operador_ent.Text.Trim(), drawstring10 = "* " + conse_entr.Text.Trim() + " *", drawstring11 = "__________________________", drawstring12 = "45";

            PrintDocument p = new PrintDocument();
            Font drawFont = new Font("Courier New", 8);//encabezado
            Font drawFont1 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont2 = new Font("Arial", 10, FontStyle.Bold);//encabezado
            Font drawFont3 = new Font("Arial", 10, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont4 = new Font("Arial", 9, FontStyle.Bold);//encabezado
            Font drawFont5 = new Font("Arial", 7, FontStyle.Bold);//encabezado
            Font drawFont8 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont9 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);//encabezado
            Font drawFont10 = new Font("PF Barcode 39", 20);//encabezado
            Font drawFont11 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);//encabezado

            SolidBrush drawBrush = new SolidBrush(Color.Black);

            // Create point for upper-left corner of drawing. 
            PointF drawPoint1 = new PointF(40.0F, 15.0F);//codigo de barras            
            PointF drawPoint2 = new PointF(7.0f, 45.0f);//linea1
            PointF drawPoint3 = new PointF(7.0f, 78.0f);//datos1
            PointF drawpoint4 = new PointF(7.0F, 95.0F);
            PointF drawpoint5 = new PointF(7.0F, 108.0F);
            PointF drawpoint6 = new PointF(7.0F, 121.0F);
            PointF drawpoint7 = new PointF(7.0F, 140.0F);
            PointF drawpoint8 = new PointF(7.0F, producto);
            producto = producto + 25;
            PointF drawpoint9 = new PointF(7.0F, producto);
            producto = producto + 30;
            PointF drawpoint10 = new PointF(45.0f, producto);//linea2
            producto = producto + 30;
            PointF drawpoint11 = new PointF(7.0f, producto);//linea3
            p.PrintPage += delegate (object sender1, PrintPageEventArgs e1)
            {



                e1.Graphics.DrawString(drawString1, drawFont1, drawBrush, drawPoint1);//encabezado            
                e1.Graphics.DrawString(drawString2, drawFont2, drawBrush, drawPoint2);//linea1
                e1.Graphics.DrawString(drawString3, drawFont3, drawBrush, drawPoint3);//datos1
                e1.Graphics.DrawString(drawString4, drawFont4, drawBrush, drawpoint4);//
                e1.Graphics.DrawString(drawString5, drawFont4, drawBrush, drawpoint5);//
                e1.Graphics.DrawString(drawString6, drawFont4, drawBrush, drawpoint6);//
                e1.Graphics.DrawString(envasesprod, drawFont5, drawBrush, drawpoint7);//
                e1.Graphics.DrawString(drawstring8, drawFont8, drawBrush, drawpoint8);//
                e1.Graphics.DrawString(drawstring9, drawFont9, drawBrush, drawpoint9);//
                e1.Graphics.DrawString(drawstring10, drawFont10, drawBrush, drawpoint10);//
                e1.Graphics.DrawString(drawstring11, drawFont11, drawBrush, drawpoint11);//
                //e.Graphics.DrawString(drawString2, drawFont, drawBrush, drawPoint4);//linea2
                //e1.Graphics.DrawString(drawString2, drawFont3, drawBrush, drawPoint5);//linea3 
            };
            try
            {
                p.Print();
            }
            catch (Exception ex)
            {
                throw new Exception("Exception Occured While Printing", ex);
            }
        }

        private void guardar_ent_Click(object sender, EventArgs e)
        {
            thisConnecion.Open();
            string fecha = "SELECT SYSDATETIME()";
            SqlCommand cmdfechoy = new SqlCommand(fecha, thisConnecion);
            fecha_hoy_hoy = Convert.ToDateTime(cmdfechoy.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();


            if (usuariostotales.Contains(usuario_actual) == false)
            {
                fecha_entrada.Text = fecha_hoy_hoy;
            }

            //*** Validacion de campos necesarios para ingresar al datagritview
            if ((clb_prov_entr.Text.ToString().Trim().Length == 0) || (cmb_prov_entr.Text.ToString().Trim().Length == 0))
            {
                MessageBox.Show("Seleccione o ingrese un Proveedor válido");
                return;
            }


            if (chofer_ent.Text.Trim().Length == 0)
            {
                MessageBox.Show("Ingrese un Nombre para el chofer");
                return;
            }

            if (operador_ent.Text.Trim().Length == 0)
            {
                MessageBox.Show("Ingrese un Nombre para el chofer");
                return;
            }

            // FROM 


            if (gridview_entr.Rows.Count == 0)
            {
                MessageBox.Show("Ingrese Al menos un Envase Para La Entrada");
                return;
            }


            thisConnecion.Open();

            proveedor = cmb_prov_entr.Text.ToString().Trim();
            rancho = cmb_rancho_entr.Text.ToString().Trim();
            tabla = cmb_tabla_entr.Text.ToString().Trim();

            if (rancho == "Sin Ranchos Disponibles")
            {
                rancho = "";
            }

            if (tabla == "Sin Tablas Disponibles")
            {
                tabla = "";
            }

            string hora_actual = DateTime.Now.ToString("hh:mm:ss");

            string fecha_actual = Convert.ToDateTime(fecha_entrada.Text).ToString("yyyy-dd-MM");

            string fecha_insert = fecha_actual + " " + hora_actual;



            string cadena = "insert into  TB_ENTRADAS_ENVASES(FECHA, PROV_CLAVE, PROV_NOMBRE, RCH_CLAVE, RCH_NOMBRE, TBL_CLAVE, TBL_NOMBRE, NOM_CHOFER, NOM_OPERADOR, RMP_FOLIO, ENT_STATUS) " +
                                "Values(CAST('" + Convert.ToDateTime(fecha_entrada.Text).ToString("dd/MM/yyyy") + "' AS DATETIME),'" + clb_prov_entr.Text.Trim() + "','" + cmb_prov_entr.Text.ToString() +
                                "','" + clb_rancho_entr.Text.Trim() + "','" + rancho.ToString().Trim() + "','" + clb_tabla_entr.Text.Trim().Trim() + "', '" + tabla.ToString() +
                                "', '" + chofer_ent.Text.Trim() + "', '" + operador_ent.Text.Trim() + "', '0', 'T')";
            SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
            cmd.ExecuteNonQuery();

            string Cadenaultimaentrada = "SELECT TOP (1) FOLIO FROM TB_ENTRADAS_ENVASES ORDER BY FOLIO DESC";
            var cmdultiamentrada = new SqlCommand(Cadenaultimaentrada, thisConnecion);
            int UltimaEntrada = Convert.ToInt32(cmdultiamentrada.ExecuteScalar());

            clb_prov_entr.Text = UltimaEntrada.ToString();

            foreach (DataRow row in ProductosGuardar_Entrada.Rows)
            {
                //row["Cantidad"] + "  |   " + row["cvl_envase"] + "  |  " + row["nombre_envase"] + row["cvl_producto"] + "  |  " + row["nombre_producto"]
                cadena = "insert into  TB_DETENTRADAS_ENVASES(FOLIO, ENV_CLAVE, ENV_NOMBRE, CANTIDAD, PROD_CLAVE, PROD_NOMBRE)" +
                                "Values('" + clb_prov_entr.Text.Trim() + "','" + row["cvl_envase"].ToString().Trim() + "','" + row["nombre_envase"].ToString().Trim() +
                                "','" + row["Cantidad"].ToString().Trim() + "','" + row["cvl_producto"].ToString().Trim() + "','" + row["nombre_producto"].ToString().Trim() + "')";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();


                cadena = "UPDATE TB_MSTR_ENVASES SET  cant_entradas = ( cant_entradas + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE prov_clave = '" + clb_prov_entr.Text.Trim() + "' AND rch_clave = '" + clb_rancho_entr.Text.Trim() +
                         "' AND tbl_clave = '" + clb_tabla_entr.Text.Trim().Trim() + "' AND env_clave = '" + row["cvl_envase"].ToString().Trim() + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES SET ENV_INV_CANT = (ENV_INV_CANT + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES_dos SET ENV_ENTR_CANT = (ENV_ENTR_CANT + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' and ENV_FECHA = '" + Convert.ToDateTime(fecha_entrada.Text).ToString("dd/MM/yyyy") + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                cadena = "UPDATE TB_MSTR_INV_ENVASES_sin_corte SET ENV_ENTR_CANT = (ENV_ENTR_CANT + " + Convert.ToInt32(row["Cantidad"].ToString().Trim()) + ") WHERE ENV_CLAVE = '" + row["cvl_envase"].ToString().Trim() + "' and ENV_FECHA = '" + Convert.ToDateTime(fecha_entrada.Text).ToString("dd/MM/yyyy") + "' ";
                cmd = new SqlCommand(cadena, thisConnecion);
                cmd.ExecuteNonQuery();

                if (row["cvl_envase"].ToString().Trim() == "81")
                {
                    validarproveedoresparrago(clb_prov_entr.Text.Trim(), Convert.ToDateTime(fecha_entrada.Text).ToString("dd/MM/yyyy"));
                }

            }
            thisConnecion.Close();

            MessageBox.Show("La Entrada se ha almacenado Con Exito");


            //Registro si es a Destiempo
            if (Convert.ToDateTime(fecha_hoy_hoy) > Convert.ToDateTime(fecha_entrada.Text))
            {
                thisConnecion.Open();
                string Cadena = "SELECT TOP (1) FOLIO FROM TB_ENTRADAS_ENVASES ORDER BY FOLIO DESC";
                cmd = new SqlCommand(Cadena, thisConnecion);
                int Valor = Convert.ToInt32(cmd.ExecuteScalar());
                conse_entr.Text = Valor.ToString();


                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values('" + fecha_hoy_hoy + "','" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ATRASADO','2.18','" + Valor + "', 'INSERCCION DE ENTRADA A DESTIEMPO FECHA " + fecha_entrada.Text + "', 'SISGAB', '" + Valor + "')";
                SqlCommand cmdent = new SqlCommand(cadenaregmov, thisConnecion);
                cmdent.ExecuteNonQuery();
                thisConnecion.Close();

                realizar_corte(Convert.ToDateTime(fecha_entrada.Text).ToString("dd/MM/yyyy"), "E", "DESTIEMPO");
            }
            else
            {
                thisConnecion.Open();
                string Cadena = "SELECT TOP (1) FOLIO FROM TB_ENTRADAS_ENVASES ORDER BY FOLIO DESC";
                cmd = new SqlCommand(Cadena, thisConnecion);
                int Valor = Convert.ToInt32(cmd.ExecuteScalar());
                conse_entr.Text = Valor.ToString();


                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values(GetDate(),'" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ENTRADA','2.18','" + Valor + "', 'INSERCCION DE ENTRADA FECHA " + fecha_entrada.Text + "', 'SISGAB', '" + Valor + "')";
                SqlCommand cmdent = new SqlCommand(cadenaregmov, thisConnecion);
                cmdent.ExecuteNonQuery();
                thisConnecion.Close();

            }


            ///********************************************


            PrintDialog printDialog1 = new PrintDialog();
            printDialog1.Document = printDocument1;
            DialogResult result = printDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                Imprimir_Entrada();
            }

            Limpiar_Entrada();
        }


        public void validarproveedoresparrago(string PROVEEDOR_ID, string fecha_actual)
        {
            //thisConnecion.Open();
            string id_proveedor = "Select cve_prov FROM Tb_ENV_PROV_CAJ_ESPARRAGO WHERE cve_prov = '" + PROVEEDOR_ID + "' AND estatus = '1'";
            SqlCommand cmd = new SqlCommand(id_proveedor, thisConnecion);
            //string proveedor = cmd.ExecuteScalar().ToString();

            object objValue = cmd.ExecuteScalar();
            if (objValue == null)
            {
                if (PROVEEDOR_ID.Trim() != "01" && PROVEEDOR_ID.Trim() != "03" && PROVEEDOR_ID.Trim() != "RO" && PROVEEDOR_ID.Trim() != "212")
                {
                    string cadenaregmov = "insert into Tb_ENV_PROV_CAJ_ESPARRAGO (cve_prov, estatus) " +
                    "Values('" + PROVEEDOR_ID + "','1')";
                    cmd = new SqlCommand(cadenaregmov, thisConnecion);
                    cmd.ExecuteNonQuery();

                    cadenaregmov = "insert into  TB_MSTR_INV_CAJAS_PROV_ESPARRAGO (PROV_CLAVE, ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                    "Values('" + PROVEEDOR_ID + "','81', '" + fecha_actual + "', '0', '0', '0')";
                    cmd = new SqlCommand(cadenaregmov, thisConnecion);
                    cmd.ExecuteNonQuery();
                }

            }
            //thisConnecion.Close();

        }

        public void validarproveedoresparragoT(string PROVEEDOR_ID, string fecha_actual, SqlTransaction transaction)
        {
            // El método asume que thisConnecion está abierta y con la transacción activa.
            // Se crean todos los comandos asociados a la transacción recibida.

            // 1. Consultar si el proveedor ya existe en la tabla de control
            string query = @"SELECT cve_prov 
                     FROM Tb_ENV_PROV_CAJ_ESPARRAGO 
                     WHERE cve_prov = @ProvId AND estatus = '1'";
            SqlCommand cmd = new SqlCommand(query, thisConnecion, transaction);
            cmd.Parameters.AddWithValue("@ProvId", PROVEEDOR_ID);

            object objValue = cmd.ExecuteScalar();

            // 2. Si no existe, y no es uno de los proveedores exentos, insertarlo
            if (objValue == null)
            {
                if (PROVEEDOR_ID.Trim() != "01" && PROVEEDOR_ID.Trim() != "03" &&
                    PROVEEDOR_ID.Trim() != "RO" && PROVEEDOR_ID.Trim() != "212")
                {
                    // Insertar en Tb_ENV_PROV_CAJ_ESPARRAGO
                    string insertProv = @"INSERT INTO Tb_ENV_PROV_CAJ_ESPARRAGO (cve_prov, estatus) 
                                 VALUES (@ProvId, '1')";
                    cmd = new SqlCommand(insertProv, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@ProvId", PROVEEDOR_ID);
                    cmd.ExecuteNonQuery();

                    // Insertar en TB_MSTR_INV_CAJAS_PROV_ESPARRAGO
                    string insertInv = @"INSERT INTO TB_MSTR_INV_CAJAS_PROV_ESPARRAGO 
                                 (PROV_CLAVE, ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) 
                                 VALUES (@ProvId, '81', @Fecha, 0, 0, 0)";
                    cmd = new SqlCommand(insertInv, thisConnecion, transaction);
                    cmd.Parameters.AddWithValue("@ProvId", PROVEEDOR_ID);
                    cmd.Parameters.AddWithValue("@Fecha", fecha_actual);
                    cmd.ExecuteNonQuery();
                }
            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                thisConnecion.Open();
                SqlCommand cmnd2 = thisConnecion.CreateCommand();
                cmnd2.CommandText = "SELECT TOP 1 inicio_sesion, usu_login FROM tb_cat_historial_dia where nombre_maquina = '" + Environment.MachineName + "' and sistema = 'SIPGAB' ORDER BY inicio_sesion desc";
                SqlDataReader reader2 = cmnd2.ExecuteReader();
                while (reader2.Read())
                {
                    Utilerias.Class1.Inicio_sesion = reader2.GetSqlDateTime(0).Value;
                    Utilerias.Class1.Usu_login = reader2.GetSqlString(1).ToString();
                    Utilerias.Class1.Nombre_equipo = Environment.MachineName;
                }
                reader2.Close();

                cmnd2 = thisConnecion.CreateCommand();
                cmnd2.CommandText = "update tb_cat_historial_dia set formulario = ' ' where nombre_maquina ='" + Utilerias.Class1.Nombre_equipo + "' and usu_login = '" + Utilerias.Class1.Usu_login + "' and inicio_sesion = '" + Utilerias.Class1.Inicio_sesion.ToString("s") + "' and sistema = 'SIPGAB'";
                reader2 = cmnd2.ExecuteReader();
                reader2.Close();
                thisConnecion.Close();
                Application.Exit();
            }
            catch (SqlException ex)
            {
                thisConnecion.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex.ToString());
                MessageBox.Show(ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            catch (Exception ex1)
            {
                thisConnecion.Close();
                Utilerias.Class1.SendMail("jbravo@mrlucky.com.mx", "jbravo", "juanjose", ex1.ToString());
                MessageBox.Show(ex1.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT TB_MSTR_INV_ENVASES.ENV_CLAVE, b.env_nombre, TB_MSTR_INV_ENVASES.ENV_FECHA, TB_MSTR_INV_ENVASES.ENV_INV_CANT FROM TB_MSTR_INV_ENVASES INNER JOIN tb_cat_envases AS b ON TB_MSTR_INV_ENVASES.ENV_CLAVE = b.env_clave ORDER BY TB_MSTR_INV_ENVASES.ENV_CLAVE";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "INVTENV");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            inv_x_proveedorini.Load(Info);
            DataColumn column;

            thisConnecion.Close();

            int T = 0;
            if (inv_x_proveedorini.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            foreach (DataRow row in inv_x_proveedorini.Rows)
            {
                int total = 0;

                string Cad = "SELECT COALESCE((SUM(cant_salidas) - SUM(cant_entradas)), 0) AS Expr3 FROM tb_mstr_envases WHERE (env_clave = '" + row["ENV_CLAVE"] + "')";
                SqlCommand cm = new SqlCommand(Cad, thisConnecion);
                thisConnecion.Open();
                int Valor = Convert.ToInt32(cm.ExecuteScalar());
                total = Convert.ToInt32(Valor.ToString());
                thisConnecion.Close();


                /*
                inv_x_proveedor_Tabla.Columns.Add("ENVASE_ID");
                inv_x_proveedor_Tabla.Columns.Add("NOMBRE");
                inv_x_proveedor_Tabla.Columns.Add("CVE");
                inv_x_proveedor_Tabla.Columns.Add("NOMBRE_PROV");
                inv_x_proveedor_Tabla.Columns.Add("CVE_RAN");
                inv_x_proveedor_Tabla.Columns.Add("NOMBRE_RANCH");
                inv_x_proveedor_Tabla.Columns.Add("CVE_TAB");
                inv_x_proveedor_Tabla.Columns.Add("NOMBRE_TABLA");
                inv_x_proveedor_Tabla.Columns.Add("ENTRADAS");
                inv_x_proveedor_Tabla.Columns.Add("SALIDAS");
                inv_x_proveedor_Tabla.Columns.Add("INVENTARIO");


                */


                DataRow rowi = inv_x_proveedor_Tabla.NewRow();
                rowi["ENVASE_ID"] = row["ENV_CLAVE"];
                rowi["NOMBRE"] = row["env_nombre"];
                rowi["CVE"] = "";
                rowi["NOMBRE_PROV"] = "";
                rowi["CVE_RAN"] = "";
                rowi["NOMBRE_RANCH"] = "";
                rowi["CVE_TAB"] = "";
                rowi["NOMBRE_TABLA"] = "";
                rowi["ENTRADAS"] = "";
                rowi["SALIDAS"] = "";
                rowi["INVENTARIO"] = Convert.ToInt32(row["ENV_INV_CANT"]) + total;
                inv_x_proveedor_Tabla.Rows.Add(rowi);

                rowi = inv_x_proveedor_Tabla.NewRow();
                rowi["ENVASE_ID"] = row["ENV_CLAVE"];
                rowi["NOMBRE"] = "";
                rowi["CVE"] = "";
                rowi["NOMBRE_PROV"] = "PLANTA";
                rowi["CVE_RAN"] = "";
                rowi["NOMBRE_RANCH"] = "";
                rowi["CVE_TAB"] = "";
                rowi["NOMBRE_TABLA"] = "";
                rowi["ENTRADAS"] = "";
                rowi["SALIDAS"] = "";
                rowi["INVENTARIO"] = Convert.ToInt32(row["ENV_INV_CANT"]);
                inv_x_proveedor_Tabla.Rows.Add(rowi);

                string CadenaX = "SELECT tb_mstr_envases.prov_clave, B.prov_nombre, tb_mstr_envases.rch_clave, C.rch_nombre, tb_mstr_envases.tbl_clave, D.tbl_nombre, tb_mstr_envases.cant_salidas AS SALIDAS, tb_mstr_envases.cant_entradas AS ENTRADAS," +
                 "tb_mstr_envases.cant_salidas - tb_mstr_envases.cant_entradas AS Inventario " +
                 "FROM tb_mstr_envases LEFT OUTER JOIN " +
                 "tb_cat_proveedor AS B ON tb_mstr_envases.prov_clave = B.prov_clave LEFT OUTER JOIN " +
                 "tb_cat_ranchos AS C ON tb_mstr_envases.prov_clave = C.prov_clave AND tb_mstr_envases.rch_clave = C.rch_clave LEFT OUTER JOIN " +
                 "tb_cat_tablas AS D ON tb_mstr_envases.prov_clave = D.prov_clave AND tb_mstr_envases.rch_clave = D.rch_clave AND tb_mstr_envases.tbl_clave = D.tbl_clave " +
                 "WHERE (tb_mstr_envases.env_clave = '" + row["ENV_CLAVE"] + "') ORDER BY tb_mstr_envases.prov_clave, tb_mstr_envases.rch_clave, tb_mstr_envases.tbl_clave";

                //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
                DataSet ds1X = new DataSet();
                SqlDataAdapter da1X = new SqlDataAdapter(Cadena, thisConnecion);
                da1.Fill(ds1X, "INVTENV");
                SqlCommand cmdX;
                cmdX = new SqlCommand(CadenaX);
                cmdX.Connection = thisConnecion;
                SqlDataReader InfoX;
                thisConnecion.Open();
                InfoX = cmdX.ExecuteReader();
                inv_x_proveedorini_DET.Clear();
                inv_x_proveedorini_DET.Load(InfoX);
                DataColumn columnX;

                thisConnecion.Close();

                foreach (DataRow rowIx in inv_x_proveedorini_DET.Rows)
                {
                    rowi = inv_x_proveedor_Tabla.NewRow();
                    rowi["ENVASE_ID"] = row["ENV_CLAVE"];
                    rowi["NOMBRE"] = "";
                    rowi["CVE"] = rowIx["prov_clave"];
                    rowi["NOMBRE_PROV"] = rowIx["prov_nombre"];
                    rowi["CVE_RAN"] = rowIx["rch_clave"];
                    rowi["NOMBRE_RANCH"] = rowIx["rch_nombre"];
                    rowi["CVE_TAB"] = rowIx["tbl_clave"];
                    rowi["NOMBRE_TABLA"] = rowIx["tbl_nombre"];
                    rowi["ENTRADAS"] = rowIx["ENTRADAS"];
                    rowi["SALIDAS"] = rowIx["SALIDAS"];
                    rowi["INVENTARIO"] = rowIx["INVENTARIO"];
                    inv_x_proveedor_Tabla.Rows.Add(rowi);
                }

            }

            DetalleEnvase ob = new DetalleEnvase(inv_x_proveedor_Tabla, 0);

            ob.Show();
        }

        private void button17_Click(object sender, EventArgs e)
        {



            thisConnecion.Open();
            //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "UNION " +
            //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
            //                "ORDER BY PDN_FOLIO";
            string Cadena = "SELECT TB_MSTR_INV_ENVASES.ENV_CLAVE, b.env_nombre, TB_MSTR_INV_ENVASES.ENV_FECHA, TB_MSTR_INV_ENVASES.ENV_INV_CANT FROM TB_MSTR_INV_ENVASES INNER JOIN tb_cat_envases AS b ON TB_MSTR_INV_ENVASES.ENV_CLAVE = b.env_clave ORDER BY TB_MSTR_INV_ENVASES.ENV_CLAVE";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "INVTENV");
            SqlCommand cmd;
            cmd = new SqlCommand(Cadena);
            cmd.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmd.ExecuteReader();
            inv_x_proveedorini.Load(Info);
            DataColumn column;

            thisConnecion.Close();

            int T = 0;
            if (inv_x_proveedorini.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }


            foreach (DataRow row in inv_x_proveedorini.Rows)
            {
                int total = 0;



                string Cad = "SELECT COALESCE((SUM(cant_salidas) - SUM(cant_entradas)), 0) AS Expr3 FROM tb_mstr_envases WHERE (env_clave = '" + row["ENV_CLAVE"] + "')";
                SqlCommand cm = new SqlCommand(Cad, thisConnecion);
                thisConnecion.Open();
                int Valor = Convert.ToInt32(cm.ExecuteScalar());
                total = Convert.ToInt32(Valor.ToString());
                thisConnecion.Close();

                DataRow rowi = inv_x_proveedor.NewRow();
                rowi["ENVASE_ID"] = row["ENV_CLAVE"];
                rowi["NOMBRE"] = row["env_nombre"];
                rowi["CVE"] = "";
                rowi["NOMBRE_PROV"] = "";
                rowi["ENTRADAS"] = "";
                rowi["SALIDAS"] = "";
                rowi["INVENTARIO"] = Convert.ToInt32(row["ENV_INV_CANT"]) + total;
                inv_x_proveedor.Rows.Add(rowi);

                rowi = inv_x_proveedor.NewRow();
                rowi["ENVASE_ID"] = row["ENV_CLAVE"];
                rowi["NOMBRE"] = "";
                rowi["CVE"] = "";
                rowi["NOMBRE_PROV"] = "PLANTA";
                rowi["ENTRADAS"] = "";
                rowi["SALIDAS"] = "";
                rowi["INVENTARIO"] = Convert.ToInt32(row["ENV_INV_CANT"]);
                inv_x_proveedor.Rows.Add(rowi);



                string CadenaX = "SELECT tb_mstr_envases.prov_clave, B.prov_nombre, SUM(tb_mstr_envases.cant_entradas) AS ENTRADAS, SUM(tb_mstr_envases.cant_salidas) AS SALIDAS, SUM(tb_mstr_envases.cant_salidas)" +
                        " - SUM(tb_mstr_envases.cant_entradas) AS INVENTARIO FROM tb_mstr_envases LEFT OUTER JOIN tb_cat_proveedor AS B ON tb_mstr_envases.prov_clave = B.prov_clave" +
                        " WHERE tb_mstr_envases.env_clave = '" + row["ENV_CLAVE"] + "' GROUP BY tb_mstr_envases.prov_clave, B.prov_nombre ORDER BY tb_mstr_envases.prov_clave";
                //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
                DataSet ds1X = new DataSet();
                SqlDataAdapter da1X = new SqlDataAdapter(Cadena, thisConnecion);
                da1.Fill(ds1X, "INVTENV");
                SqlCommand cmdX;
                cmdX = new SqlCommand(CadenaX);
                cmdX.Connection = thisConnecion;
                SqlDataReader InfoX;
                thisConnecion.Open();
                InfoX = cmdX.ExecuteReader();
                inv_x_proveedorini_DET.Clear();
                inv_x_proveedorini_DET.Load(InfoX);
                DataColumn columnX;

                thisConnecion.Close();

                foreach (DataRow rowIx in inv_x_proveedorini_DET.Rows)
                {
                    rowi = inv_x_proveedor.NewRow();
                    rowi["ENVASE_ID"] = row["ENV_CLAVE"];
                    rowi["NOMBRE"] = "";
                    rowi["CVE"] = rowIx["prov_clave"];
                    rowi["NOMBRE_PROV"] = rowIx["prov_nombre"];
                    rowi["ENTRADAS"] = rowIx["ENTRADAS"];
                    rowi["SALIDAS"] = rowIx["SALIDAS"];
                    rowi["INVENTARIO"] = rowIx["INVENTARIO"];
                    inv_x_proveedor.Rows.Add(rowi);
                }

            }

            DetalleEnvase ob = new DetalleEnvase(inv_x_proveedor, 1);

            ob.Show();

        }

        private void cmbprovinirep_SelectedValueChanged(object sender, EventArgs e)
        {

            if (cmbprovinirep.SelectedValue == null)
            {
                foreach (DataRow row in Proveedor.Rows)
                {
                    if (row["prov_nombre"] == clbprovinirep.Text)
                    {
                        clbprovinirep.Text = cmbprovinirep.ValueMember.ToString();
                    }
                }
            }
            else
            {
                clbprovinirep.Text = cmbprovinirep.SelectedValue.ToString();
            }

        }

        private void cmbprovfinrep_SelectedValueChanged(object sender, EventArgs e)
        {

            if (cmbprovfinrep.SelectedValue == null)
            {
                foreach (DataRow row in Proveedor.Rows)
                {
                    if (row["prov_nombre"] == clbprovfinrep.Text)
                    {
                        clbprovfinrep.Text = cmbprovfinrep.ValueMember.ToString();
                    }
                }
            }
            else
            {
                clbprovfinrep.Text = cmbprovfinrep.SelectedValue.ToString();
            }
        }

        private void clbprovinirep_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmbprovinirep.SelectedValue = clbprovinirep.Text.ToString();
        }


        private void clbprovfinrep_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmbprovfinrep.SelectedValue = clbprovfinrep.Text.ToString();
        }

        private void cmbenvrep_SelectedValueChanged(object sender, EventArgs e)
        {
            if (cmbenvrep.SelectedValue == null)
            {
                foreach (DataRow row in Envase.Rows)
                {
                    if (row["env_nombre"] == clbprovfinrep.Text)
                    {
                        clbenvrep.Text = cmbenvrep.ValueMember.ToString();
                    }
                }
            }
            else
            {
                clbenvrep.Text = cmbenvrep.SelectedValue.ToString();
            }
        }

        private void clbenvrep_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
                cmbenvrep.SelectedValue = clbenvrep.Text.ToString();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (radioButton1.Checked == true)
            {
                thisConnecion.Open();
                //string Cadena = "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
                //                "UNION " +
                //                "SELECT A.PDN_FOLIO,A.PDN_FECHA,A.PLACACAJA,B.prod_clave,B.pdn_num_unidades FROM TB_MSTR_PEDIDOS_NAL A, tb_det_pedidos B WHERE A.PDN_FECHA = '" + Program.MyGlobal.PubFecEmb + "' AND A.placacaja = '" + Program.MyGlobal.PubNoTrailer + "' AND A.PDN_FOLIO = B.PDN_FOLIO AND A.PDN_TIPO = B.PDN_TIPO " +
                //                "ORDER BY PDN_FOLIO";
                string Cadena = "SELECT * FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO" +
                                "WHERE (TB_ENTRADAS_ENVASES.FECHA BETWEEN '" + FechaInirepo.Text + "' AND '" + FechaFinrepo.Text + "') AND (TB_ENTRADAS_ENVASES.PROV_CLAVE BETWEEN '" + clbprovinirep.Text + "' AND '" + clbprovfinrep.Text + "')" +
                                "ORDER BY TB_ENTRADAS_ENVASES.PROV_CLAVE, TB_ENTRADAS_ENVASES.RCH_CLAVE, TB_ENTRADAS_ENVASES.TBL_CLAVE";
                //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
                DataSet ds1 = new DataSet();
                SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
                da1.Fill(ds1, "kardex");
                SqlCommand cmd;
                cmd = new SqlCommand(Cadena);
                cmd.Connection = thisConnecion;
                SqlDataReader Info;
                Info = cmd.ExecuteReader();
                Kardex.Load(Info);
                thisConnecion.Close();

                reporteKardex.Columns.Add("FOLIO");
                reporteKardex.Columns.Add("RECIBO");
                reporteKardex.Columns.Add("FOLIO MP");
                reporteKardex.Columns.Add("FECHA");
                reporteKardex.Columns.Add("CANTIDAD");
                reporteKardex.Columns.Add("ENVASE");
                reporteKardex.Columns.Add("PRODUCTO");
                reporteKardex.Columns.Add("ESTADO");

                DataRow rowi = inv_x_proveedor.NewRow();

                int inicio = 0;

                string antprove = "";
                string antrancho = "";
                string anttabla = "";


                int totalprove = 0;
                int totalrancho = 0;
                int totaltabla = 0;


                foreach (DataRow row in Kardex.Rows)
                {
                    if (antprove != "")
                    {
                        rowi = reporteKardex.NewRow();
                        rowi["FOLIO"] = row["PROV_NOMBRE"] + ", RCH:" + row["RCH_NOMBRE"] + ", TBL:" + row["TBL_NOMBRE"];
                        rowi["RECIBO"] = "";
                        rowi["FOLIO MP"] = "";
                        rowi["FECHA"] = "";
                        rowi["CANTIDAD"] = "";
                        rowi["ENVASE"] = "";
                        rowi["PRODUCTO"] = "";
                        rowi["ESTADO"] = "";
                        reporteKardex.Rows.Add(rowi);
                    }
                    else if (antprove != row["PROV_CLAVE"])
                    {
                        rowi = reporteKardex.NewRow();
                        rowi["FOLIO"] = "TOTAL PROV: " + antprove + " - " + totalprove;
                        rowi["RECIBO"] = "";
                        rowi["FOLIO MP"] = "";
                        rowi["FECHA"] = "";
                        rowi["CANTIDAD"] = "";
                        rowi["ENVASE"] = "";
                        rowi["PRODUCTO"] = "";
                        rowi["ESTADO"] = "";
                        reporteKardex.Rows.Add(rowi);

                        totalprove = 0;
                        totalrancho = 0;
                        totaltabla = 0;

                        rowi = reporteKardex.NewRow();
                        rowi["FOLIO"] = row["PROV_NOMBRE"] + ", RCH:" + row["RCH_NOMBRE"] + ", TBL:" + row["TBL_NOMBRE"];
                        rowi["RECIBO"] = "";
                        rowi["FOLIO MP"] = "";
                        rowi["FECHA"] = "";
                        rowi["CANTIDAD"] = "";
                        rowi["ENVASE"] = "";
                        rowi["PRODUCTO"] = "";
                        rowi["ESTADO"] = "";
                        reporteKardex.Rows.Add(rowi);

                    }
                    else if ((antrancho != row["RCH_CLAVE"] && antrancho != ""))
                    {


                    }
                    else
                    {
                        rowi = reporteKardex.NewRow();
                        rowi["FOLIO"] = row["PROV_NOMBRE"] + ", RCH:" + row["RCH_NOMBRE"] + ", TBL:" + row["TBL_NOMBRE"];
                        rowi["RECIBO"] = "";
                        rowi["FOLIO MP"] = "";
                        rowi["FECHA"] = "";
                        rowi["CANTIDAD"] = "";
                        rowi["ENVASE"] = "";
                        rowi["PRODUCTO"] = "";
                        rowi["ESTADO"] = "";
                        reporteKardex.Rows.Add(rowi);
                    }


                }


            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            int Total_INVENTARIO_Ini = 0;
            int Total_Entradas = 0;
            int Total_Salidas = 0;
            int Total_Inventario_Final = 0;

            thisConnecion.Open();
            string fecha = "SELECT SYSDATETIME()";
            SqlCommand cmd = new SqlCommand(fecha, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();

            /*thisConnecion.Open();
            string valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos where ENV_FECHA NOT IN ('"+ fecha_hoy +"') ORDER BY ENV_FECHA DESC";
            cmd = new SqlCommand(valida, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();*/


            thisConnecion.Open();
            string Cadena = "SELECT TB_MSTR_INV_ENVASES_dos.ENV_CLAVE, b.env_nombre, TB_MSTR_INV_ENVASES_dos.ENV_FECHA, TB_MSTR_INV_ENVASES_dos.ENV_INV_INI_CANT, TB_MSTR_INV_ENVASES_dos.ENV_ENTR_CANT, TB_MSTR_INV_ENVASES_dos.ENV_SAL_CANT FROM TB_MSTR_INV_ENVASES_dos INNER JOIN tb_cat_envases AS b ON TB_MSTR_INV_ENVASES_dos.ENV_CLAVE = b.env_clave where ENV_FECHA = '" + fecha_ent + "' ORDER BY TB_MSTR_INV_ENVASES_dos.ENV_CLAVE";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "INVTENV");
            SqlCommand cmdx = new SqlCommand(Cadena);
            cmdx.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmdx.ExecuteReader();
            inv_x_proveedorini.Clear();
            inv_x_proveedorini.Load(Info);
            DataColumn column;

            thisConnecion.Close();

            int T = 0;
            if (inv_x_proveedorini.Rows.Count == 0)
            {
                MessageBox.Show("CORTE DIARIO NO REALIZADO, NO SE PUEDE MOSTRAR INFORMACION", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            inv_x_proveedor.Clear();

            foreach (DataRow row in inv_x_proveedorini.Rows)
            {
                int total = 0;


                thisConnecion.Open();
                string valida = "SELECT sum(ENV_INV_INI_CANT) FROM TB_MSTR_INV_ENVASES_dos WHERE ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND ENV_FECHA = '" + fecha_ent + "'";
                cmd = new SqlCommand(valida, thisConnecion);
                int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                //thisConnecion.Close();

                //thisConnecion.Open();
                valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND ENT_STATUS != 'C'";
                cmd = new SqlCommand(valida, thisConnecion);
                int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                //thisConnecion.Close();

                //thisConnecion.Open();
                valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND SAL_STATUS != 'C'";
                cmd = new SqlCommand(valida, thisConnecion);
                int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                thisConnecion.Close();




                DataRow rowi = inv_x_proveedor.NewRow();
                rowi["ID"] = row["ENV_CLAVE"];
                rowi["NOMBRE"] = row["env_nombre"];
                rowi["FECHA"] = Convert.ToDateTime(row["ENV_FECHA"]).ToString("dd/MM/yyyy");
                rowi["INVENTARIO INICIAL"] = row["ENV_INV_INI_CANT"];
                rowi["ENTRADAS"] = Convert.ToInt32(entradas);
                rowi["SALIDAS"] = Convert.ToInt32(salidas);
                rowi["INVENTARIO FINAL"] = Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(entradas) - Convert.ToInt32(salidas);
                inv_x_proveedor.Rows.Add(rowi);



                Total_INVENTARIO_Ini = Total_INVENTARIO_Ini + Convert.ToInt16(row["ENV_INV_INI_CANT"]);
                Total_Entradas = Total_Entradas + Convert.ToInt32(entradas);
                Total_Salidas = Total_Salidas + Convert.ToInt32(salidas);
                Total_Inventario_Final = Total_Inventario_Final + Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(salidas) - Convert.ToInt32(entradas);

            }

            DataRow rowixw = inv_x_proveedor.NewRow();
            rowixw["ID"] = "";
            rowixw["NOMBRE"] = "TOTALES";
            rowixw["FECHA"] = "";
            rowixw["INVENTARIO INICIAL"] = Total_INVENTARIO_Ini;
            rowixw["ENTRADAS"] = Total_Entradas;
            rowixw["SALIDAS"] = Total_Salidas;
            rowixw["INVENTARIO FINAL"] = Total_Inventario_Final;
            inv_x_proveedor.Rows.Add(rowixw);

            DetalleEnvase ob = new DetalleEnvase(inv_x_proveedor, 0);

            ob.Show();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            GENERAR FormMENSAJE = new GENERAR();
            Reporte_Entradas.Clear();
            Reporte_Salidas.Clear();
            Reporte_Kardex.Clear();


            if (radioButton1.Checked == true)
            {

                FormMENSAJE.Show();

                thisConnecion.Open();
                string query = "SELECT * FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE IN (SELECT env_clave FROM  tb_cat_envases WHERE (env_inventario = '1')) ORDER BY FECHA, ENV_CLAVE";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                int totalent = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Entradas.NewRow();
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["RECIBO_MP"] = Convert.ToString(dr["RMP_FOLIO"]).Trim();
                    rowi["ESTATUS"] = Convert.ToString(dr["ENT_STATUS"]).Trim();
                    Reporte_Entradas.Rows.Add(rowi);


                    if (Convert.ToString(dr["ENT_STATUS"]).Trim() != "C")
                    {
                        totalent = totalent + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }

                DataRow rowix = Reporte_Entradas.NewRow();
                rowix["FOLIO"] = "";
                rowix["FECHA"] = "";
                rowix["PROVEEDOR_CLAVE"] = "";
                rowix["PROVEEDOR_NOMBRE"] = "";
                rowix["RANCHO_CLAVE"] = "";
                rowix["RANCHO_NOMBRE"] = "";
                rowix["TABLA_CLAVE"] = "";
                rowix["TABLA_NOMBRE"] = "";
                rowix["IDENVASE"] = "";
                rowix["NOMBREENVASE"] = "TOTAL";
                rowix["CANTIDAD"] = totalent;
                rowix["PROD_CLAVE"] = "";
                rowix["PROD_NOMBRE"] = "";
                rowix["RECIBO_MP"] = "";
                rowix["ESTATUS"] = "";
                Reporte_Entradas.Rows.Add(rowix);

                thisConnecion.Close();

                #region fix issue: dll Interoperabilidad de Excel
                // En lugar de:
                // Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                // Usa:
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic aplicacion = Activator.CreateInstance(excelType);

                // A partir de aquí, todo se maneja con dynamic
                dynamic libro = aplicacion.Workbooks.Add();
                dynamic hoja = libro.Worksheets[1];

                // Las llamadas a propiedades y métodos son resueltas en tiempo de ejecución (IDispatch)
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                dynamic rango = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                rango.Font.Bold = true;
                rango.Font.Size = 16;
                rango.MergeCells = true;

                // ... resto del código similar, usando dynamic en lugar de tipos concretos

                // Para hacer visible
                aplicacion.Visible = true;

                // Liberación: no hay interfaz tipada, pero debes liberar los objetos COM igual
                if (aplicacion != null) Marshal.ReleaseComObject(aplicacion);
                #endregion

                //Microsoft.Office.Interop.Excel.Application aplicacion;
                //Microsoft.Office.Interop.Excel.Workbook libro;
                //Microsoft.Office.Interop.Excel.Worksheet hoja;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libro = aplicacion.Workbooks.Add();
                //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range r;
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                r.Font.Bold = true;
                r.Font.Size = 16;
                r.MergeCells = true;

                hoja.Cells[3, 3] = "REPORTE GENERAL DE ENTRADAS DE ENVASES DEL " + FechaInirepo.Text + " AL " + FechaFinrepo.Text;
                r = hoja.Range[hoja.Cells[3, 3], hoja.Cells[3, 8]];
                r.Font.Bold = true;
                r.MergeCells = true;


                string ruta = "c:\\SisGabWeb\\logo.png";
                //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 0, 70, 70);
                hoja.Range[hoja.Cells[1, 2], hoja.Cells[4, 2]].Merge();
                r = hoja.get_Range("D1", "D3");


                hoja.Cells[6, 1] = "Folio";
                hoja.Cells[6, 2] = "Fecha";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Clave Rancho";
                hoja.Cells[6, 6] = "Nombre Rancho";
                hoja.Cells[6, 7] = "Clave Tabla";
                hoja.Cells[6, 8] = "Nombre Tabla";
                hoja.Cells[6, 9] = "Envase Clave";
                hoja.Cells[6, 10] = "Envase Nombre";
                hoja.Cells[6, 11] = "Cantidad";
                hoja.Cells[6, 12] = "Producto Clave";
                hoja.Cells[6, 13] = "producto Nombre";
                hoja.Cells[6, 14] = "Recibo MP";


                r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 14]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;

                r = hoja.Range[hoja.Cells[7, 3], hoja.Cells[Reporte_Entradas.Rows.Count + 10, 14]];
                r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                int filaactual = 7;
                foreach (DataRow row in Reporte_Entradas.Rows)
                {   //Here 2 cell is target value and 1 cell is Volume 

                    hoja.Cells[filaactual, 1] = row[0];
                    hoja.Cells[filaactual, 2] = row[1];
                    hoja.Cells[filaactual, 3] = row[2];
                    hoja.Cells[filaactual, 4] = row[3];
                    hoja.Cells[filaactual, 5] = row[4];
                    hoja.Cells[filaactual, 6] = row[5];
                    hoja.Cells[filaactual, 7] = row[6];
                    hoja.Cells[filaactual, 8] = row[7];
                    hoja.Cells[filaactual, 9] = row[8];
                    hoja.Cells[filaactual, 10] = row[9];
                    hoja.Cells[filaactual, 11] = row[10];
                    hoja.Cells[filaactual, 12] = row[11];
                    hoja.Cells[filaactual, 13] = row[12];
                    hoja.Cells[filaactual, 14] = row[13];

                    filaactual++;
                }


                int rowdatagrid = 7;
                foreach (DataRow row in Reporte_Entradas.Rows)
                {
                    string dato = Convert.ToString(row["ESTATUS"]);
                    if (dato == "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 14]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        r.Font.Bold = true;

                    }

                    string datoenvase = Convert.ToString(row["NOMBREENVASE"]);
                    if (datoenvase == "TOTAL")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 14]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);
                        r.Font.Bold = true;

                    }

                    rowdatagrid++;
                }


                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;
                FormMENSAJE.Close();
                thisConnecion.Close();
            }
            else if (radioButton2.Checked == true)
            {
                FormMENSAJE.Show();
                thisConnecion.Open();
                string query = "SELECT * FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE IN (SELECT env_clave FROM  tb_cat_envases WHERE (env_inventario = '1')) ORDER BY FECHA, ENV_CLAVE";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                int totalsal = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Salidas.NewRow();
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["ESTATUS"] = Convert.ToString(dr["SAL_STATUS"]).Trim();
                    Reporte_Salidas.Rows.Add(rowi);


                    if (Convert.ToString(dr["SAL_STATUS"]).Trim() != "C")
                    {
                        totalsal = totalsal + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }

                DataRow rowix = Reporte_Salidas.NewRow();
                rowix["FOLIO"] = "";
                rowix["FECHA"] = "";
                rowix["PROVEEDOR_CLAVE"] = "";
                rowix["PROVEEDOR_NOMBRE"] = "";
                rowix["RANCHO_CLAVE"] = "";
                rowix["RANCHO_NOMBRE"] = "";
                rowix["TABLA_CLAVE"] = "";
                rowix["TABLA_NOMBRE"] = "";
                rowix["IDENVASE"] = "";
                rowix["NOMBREENVASE"] = "TOTAL";
                rowix["CANTIDAD"] = totalsal;
                rowix["PROD_CLAVE"] = "";
                rowix["PROD_NOMBRE"] = "";
                rowix["ESTATUS"] = "";
                Reporte_Salidas.Rows.Add(rowix);

                thisConnecion.Close();

                #region fix issue: dll Interoperabilidad de Excel
                // En lugar de:
                // Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                // Usa:
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic aplicacion = Activator.CreateInstance(excelType);

                // A partir de aquí, todo se maneja con dynamic
                dynamic libro = aplicacion.Workbooks.Add();
                dynamic hoja = libro.Worksheets[1];

                // Las llamadas a propiedades y métodos son resueltas en tiempo de ejecución (IDispatch)
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                dynamic rango = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                rango.Font.Bold = true;
                rango.Font.Size = 16;
                rango.MergeCells = true;

                // ... resto del código similar, usando dynamic en lugar de tipos concretos

                // Para hacer visible
                aplicacion.Visible = true;

                // Liberación: no hay interfaz tipada, pero debes liberar los objetos COM igual
                if (aplicacion != null) Marshal.ReleaseComObject(aplicacion);
                #endregion

                //Microsoft.Office.Interop.Excel.Application aplicacion;
                //Microsoft.Office.Interop.Excel.Workbook libro;
                //Microsoft.Office.Interop.Excel.Worksheet hoja;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libro = aplicacion.Workbooks.Add();
                //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range r;
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                r.Font.Bold = true;
                r.Font.Size = 16;
                r.MergeCells = true;

                hoja.Cells[3, 3] = "REPORTE GENERAL DE SALIDAS DE ENVASES DEL " + FechaInirepo.Text + " AL " + FechaFinrepo.Text;
                r = hoja.Range[hoja.Cells[3, 3], hoja.Cells[3, 8]];
                r.Font.Bold = true;
                r.MergeCells = true;

                string ruta = "c:\\SisGabWeb\\logo.png";
                //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 0, 70, 70);
                hoja.Range[hoja.Cells[1, 2], hoja.Cells[4, 2]].Merge();
                r = hoja.get_Range("D1", "D3");



                hoja.Cells[6, 1] = "Folio";
                hoja.Cells[6, 2] = "Fecha";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Clave Rancho";
                hoja.Cells[6, 6] = "Nombre Rancho";
                hoja.Cells[6, 7] = "Clave Tabla";
                hoja.Cells[6, 8] = "Nombre Tabla";
                hoja.Cells[6, 9] = "Envase Clave";
                hoja.Cells[6, 10] = "Envase Nombre";
                hoja.Cells[6, 11] = "Cantidad";
                hoja.Cells[6, 12] = "Producto Clave";
                hoja.Cells[6, 13] = "producto Nombre";


                r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 14]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;

                r = hoja.Range[hoja.Cells[7, 3], hoja.Cells[Reporte_Salidas.Rows.Count + 10, 13]];
                r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                int filaactual = 7;
                foreach (DataRow row in Reporte_Salidas.Rows)
                {   //Here 2 cell is target value and 1 cell is Volume 

                    hoja.Cells[filaactual, 1] = row[0];
                    hoja.Cells[filaactual, 2] = row[1];
                    hoja.Cells[filaactual, 3] = row[2];
                    hoja.Cells[filaactual, 4] = row[3];
                    hoja.Cells[filaactual, 5] = row[4];
                    hoja.Cells[filaactual, 6] = row[5];
                    hoja.Cells[filaactual, 7] = row[6];
                    hoja.Cells[filaactual, 8] = row[7];
                    hoja.Cells[filaactual, 9] = row[8];
                    hoja.Cells[filaactual, 10] = row[9];
                    hoja.Cells[filaactual, 11] = row[10];
                    hoja.Cells[filaactual, 12] = row[11];
                    hoja.Cells[filaactual, 13] = row[12];

                    filaactual++;
                }


                int rowdatagrid = 7;
                foreach (DataRow row in Reporte_Salidas.Rows)
                {
                    string dato = Convert.ToString(row["ESTATUS"]);
                    if (dato == "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 13]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        r.Font.Bold = true;

                    }

                    string datoenvase = Convert.ToString(row["NOMBREENVASE"]);
                    if (datoenvase == "TOTAL")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 13]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);
                        r.Font.Bold = true;

                    }

                    rowdatagrid++;
                }

                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;

                FormMENSAJE.Close();



            }
            else
            {

                FormMENSAJE.Show();
                thisConnecion.Open();
                string query = "SELECT * FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE IN (SELECT env_clave FROM  tb_cat_envases WHERE (env_inventario = '1')) ORDER BY FECHA, ENV_CLAVE";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                int totalent = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Kardex.NewRow();
                    rowi["NUMERO"] = 0;
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["RECIBO_MP"] = Convert.ToString(dr["RMP_FOLIO"]).Trim();
                    rowi["TIPO"] = "ENTRADA";
                    rowi["ESTATUS"] = Convert.ToString(dr["ENT_STATUS"]).Trim();
                    Reporte_Kardex.Rows.Add(rowi);


                    if (Convert.ToString(dr["ENT_STATUS"]).Trim() != "C")
                    {
                        totalent = totalent + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }


                thisConnecion.Close();

                //Ingresar Salidas

                thisConnecion.Open();
                query = "SELECT * FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE IN (SELECT env_clave FROM  tb_cat_envases WHERE (env_inventario = '1')) ORDER BY FECHA, ENV_CLAVE";
                cm = new SqlCommand(query, thisConnecion);
                dr = cm.ExecuteReader();

                int totalsal = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Kardex.NewRow();
                    rowi["NUMERO"] = 0;
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["RECIBO_MP"] = "";
                    rowi["TIPO"] = "SALIDA";
                    rowi["ESTATUS"] = Convert.ToString(dr["SAL_STATUS"]).Trim();
                    Reporte_Kardex.Rows.Add(rowi);


                    if (Convert.ToString(dr["SAL_STATUS"]).Trim() != "C")
                    {
                        totalsal = totalsal + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }

                DataRow rowix = Reporte_Kardex.NewRow();
                rowix["NUMERO"] = 1;
                rowix["FOLIO"] = "";
                rowix["FECHA"] = "";
                rowix["PROVEEDOR_CLAVE"] = "";
                rowix["PROVEEDOR_NOMBRE"] = "";
                rowix["RANCHO_CLAVE"] = "TOTALES";
                rowix["RANCHO_NOMBRE"] = "";
                rowix["TABLA_CLAVE"] = "ENTRADAS";
                rowix["TABLA_NOMBRE"] = totalent;
                rowix["IDENVASE"] = "";
                rowix["NOMBREENVASE"] = "SALIDAS";
                rowix["CANTIDAD"] = totalsal;
                rowix["PROD_CLAVE"] = "";
                rowix["PROD_NOMBRE"] = "";
                rowix["RECIBO_MP"] = "";
                rowix["TIPO"] = "";
                rowix["ESTATUS"] = "";
                Reporte_Kardex.Rows.Add(rowix);

                thisConnecion.Close();


                Reporte_Kardex.DefaultView.Sort = "NUMERO, FECHA, Tipo, IDENVASE";
                DataView dv = Reporte_Kardex.DefaultView;

                #region fix issue: dll Interoperabilidad de Excel
                // En lugar de:
                // Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                // Usa:
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic aplicacion = Activator.CreateInstance(excelType);

                // A partir de aquí, todo se maneja con dynamic
                dynamic libro = aplicacion.Workbooks.Add();
                dynamic hoja = libro.Worksheets[1];

                // Las llamadas a propiedades y métodos son resueltas en tiempo de ejecución (IDispatch)
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                dynamic rango = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                rango.Font.Bold = true;
                rango.Font.Size = 16;
                rango.MergeCells = true;

                // ... resto del código similar, usando dynamic en lugar de tipos concretos

                // Para hacer visible
                aplicacion.Visible = true;

                // Liberación: no hay interfaz tipada, pero debes liberar los objetos COM igual
                if (aplicacion != null) Marshal.ReleaseComObject(aplicacion);
                #endregion

                //Microsoft.Office.Interop.Excel.Application aplicacion;
                //Microsoft.Office.Interop.Excel.Workbook libro;
                //Microsoft.Office.Interop.Excel.Worksheet hoja;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libro = aplicacion.Workbooks.Add();
                //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range r;
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                r.Font.Bold = true;
                r.Font.Size = 16;
                r.MergeCells = true;

                hoja.Cells[3, 3] = "REPORTE GENERAL DE ENTRADAS Y SALIDAS DE ENVASES DEL " + FechaInirepo.Text + " AL " + FechaFinrepo.Text;
                r = hoja.Range[hoja.Cells[3, 3], hoja.Cells[3, 8]];
                r.Font.Bold = true;
                r.MergeCells = true;


                string ruta = "c:\\SisGabWeb\\logo.png";
                //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 0, 70, 70);
                hoja.Range[hoja.Cells[1, 2], hoja.Cells[4, 2]].Merge();
                r = hoja.get_Range("D1", "D3");

                hoja.Cells[6, 1] = "Folio";
                hoja.Cells[6, 2] = "Fecha";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Clave Rancho";
                hoja.Cells[6, 6] = "Nombre Rancho";
                hoja.Cells[6, 7] = "Clave Tabla";
                hoja.Cells[6, 8] = "Nombre Tabla";
                hoja.Cells[6, 9] = "Envase Clave";
                hoja.Cells[6, 10] = "Envase Nombre";
                hoja.Cells[6, 11] = "Cantidad";
                hoja.Cells[6, 12] = "Producto Clave";
                hoja.Cells[6, 13] = "producto Nombre";
                hoja.Cells[6, 14] = "Recibo MP";
                hoja.Cells[6, 15] = "TIPO MOVIMIENTO";


                r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 14]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;

                r = hoja.Range[hoja.Cells[7, 3], hoja.Cells[Reporte_Kardex.Rows.Count + 10, 15]];
                r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                int filaactual = 7;
                foreach (DataRowView row in dv)
                {   //Here 2 cell is target value and 1 cell is Volume 

                    hoja.Cells[filaactual, 1] = row[1];
                    hoja.Cells[filaactual, 2] = row[2];
                    hoja.Cells[filaactual, 3] = row[3];
                    hoja.Cells[filaactual, 4] = row[4];
                    hoja.Cells[filaactual, 5] = row[5];
                    hoja.Cells[filaactual, 6] = row[6];
                    hoja.Cells[filaactual, 7] = row[7];
                    hoja.Cells[filaactual, 8] = row[8];
                    hoja.Cells[filaactual, 9] = row[9];
                    hoja.Cells[filaactual, 10] = row[10];
                    hoja.Cells[filaactual, 11] = row[11];
                    hoja.Cells[filaactual, 12] = row[12];
                    hoja.Cells[filaactual, 13] = row[13];
                    hoja.Cells[filaactual, 14] = row[14];
                    hoja.Cells[filaactual, 15] = row[15];

                    filaactual++;
                }


                int rowdatagrid = 7;
                foreach (DataRowView row in dv)
                {
                    string dato = Convert.ToString(row["ESTATUS"]);
                    if (dato == "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        r.Font.Bold = true;

                    }

                    string datotipo = Convert.ToString(row["TIPO"]);
                    if (datotipo == "ENTRADA" && dato != "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                        r.Font.Bold = true;

                    }
                    else if (datotipo == "SALIDA" && dato != "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                        r.Font.Bold = true;

                    }

                    string datoenvase = Convert.ToString(row["RANCHO_CLAVE"]);
                    if (datoenvase == "TOTALES")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);
                        r.Font.Bold = true;

                    }

                    rowdatagrid++;
                }

                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;

                FormMENSAJE.Close();

            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            GENERAR FormMENSAJE = new GENERAR();

            if (clbenvrep.Text.Trim().Length == 0)
            {
                MessageBox.Show("Debe Seleccionar un Envase para Esta Operación");
                return;
            }


            string clave_envase = clbenvrep.Text.Trim();

            Reporte_Entradas.Clear();
            Reporte_Salidas.Clear();
            Reporte_Kardex.Clear();


            if (radioButton1.Checked == true)
            {
                FormMENSAJE.Show();
                thisConnecion.Open();
                string query = "SELECT * FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE = '" + clave_envase + "' ORDER BY FECHA, ENV_CLAVE";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                int totalent = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Entradas.NewRow();
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["RECIBO_MP"] = Convert.ToString(dr["RMP_FOLIO"]).Trim();
                    rowi["ESTATUS"] = Convert.ToString(dr["ENT_STATUS"]).Trim();
                    Reporte_Entradas.Rows.Add(rowi);


                    if (Convert.ToString(dr["ENT_STATUS"]).Trim() != "C")
                    {
                        totalent = totalent + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }

                DataRow rowix = Reporte_Entradas.NewRow();
                rowix["FOLIO"] = "";
                rowix["FECHA"] = "";
                rowix["PROVEEDOR_CLAVE"] = "";
                rowix["PROVEEDOR_NOMBRE"] = "";
                rowix["RANCHO_CLAVE"] = "";
                rowix["RANCHO_NOMBRE"] = "";
                rowix["TABLA_CLAVE"] = "";
                rowix["TABLA_NOMBRE"] = "";
                rowix["IDENVASE"] = "";
                rowix["NOMBREENVASE"] = "TOTAL";
                rowix["CANTIDAD"] = totalent;
                rowix["PROD_CLAVE"] = "";
                rowix["PROD_NOMBRE"] = "";
                rowix["RECIBO_MP"] = "";
                rowix["ESTATUS"] = "";
                Reporte_Entradas.Rows.Add(rowix);

                thisConnecion.Close();

                #region fix issue: dll Interoperabilidad de Excel
                // En lugar de:
                // Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                // Usa:
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic aplicacion = Activator.CreateInstance(excelType);

                // A partir de aquí, todo se maneja con dynamic
                dynamic libro = aplicacion.Workbooks.Add();
                dynamic hoja = libro.Worksheets[1];

                // Las llamadas a propiedades y métodos son resueltas en tiempo de ejecución (IDispatch)
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                dynamic r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                //rango.Font.Bold = true;
                //rango.Font.Size = 16;
                //rango.MergeCells = true;

                // ... resto del código similar, usando dynamic en lugar de tipos concretos

                // Para hacer visible
                //                aplicacion.Visible = true;

                #endregion

                #region issue: dll Interoperabilidad de Excel
                //Microsoft.Office.Interop.Excel.Application aplicacion;
                //Microsoft.Office.Interop.Excel.Workbook libro;
                //Microsoft.Office.Interop.Excel.Worksheet hoja;
                //aplicacion = new Microsoft.Office.Interop.Excel.Application();
                //libro = aplicacion.Workbooks.Add();
                //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
                //hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                //Microsoft.Office.Interop.Excel.Range r;
                #endregion
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                r.Font.Bold = true;
                r.Font.Size = 16;
                r.MergeCells = true;

                hoja.Cells[3, 3] = "REPORTE GENERAL DE ENTRADAS DE ENVASES DEL " + FechaInirepo.Text + " AL " + FechaFinrepo.Text;
                r = hoja.Range[hoja.Cells[3, 3], hoja.Cells[3, 8]];
                r.Font.Bold = true;
                r.MergeCells = true;


                string ruta = "c:\\SisGabWeb\\logo.png";
                //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 0, 70, 70);
                hoja.Range[hoja.Cells[1, 2], hoja.Cells[4, 2]].Merge();
                r = hoja.get_Range("D1", "D3");


                hoja.Cells[6, 1] = "Folio";
                hoja.Cells[6, 2] = "Fecha";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Clave Rancho";
                hoja.Cells[6, 6] = "Nombre Rancho";
                hoja.Cells[6, 7] = "Clave Tabla";
                hoja.Cells[6, 8] = "Nombre Tabla";
                hoja.Cells[6, 9] = "Envase Clave";
                hoja.Cells[6, 10] = "Envase Nombre";
                hoja.Cells[6, 11] = "Cantidad";
                hoja.Cells[6, 12] = "Producto Clave";
                hoja.Cells[6, 13] = "producto Nombre";
                hoja.Cells[6, 14] = "Recibo MP";


                r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 14]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;

                r = hoja.Range[hoja.Cells[7, 3], hoja.Cells[Reporte_Entradas.Rows.Count + 10, 14]];
                r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                int filaactual = 7;
                foreach (DataRow row in Reporte_Entradas.Rows)
                {   //Here 2 cell is target value and 1 cell is Volume 

                    hoja.Cells[filaactual, 1] = row[0];
                    hoja.Cells[filaactual, 2] = row[1];
                    hoja.Cells[filaactual, 3] = row[2];
                    hoja.Cells[filaactual, 4] = row[3];
                    hoja.Cells[filaactual, 5] = row[4];
                    hoja.Cells[filaactual, 6] = row[5];
                    hoja.Cells[filaactual, 7] = row[6];
                    hoja.Cells[filaactual, 8] = row[7];
                    hoja.Cells[filaactual, 9] = row[8];
                    hoja.Cells[filaactual, 10] = row[9];
                    hoja.Cells[filaactual, 11] = row[10];
                    hoja.Cells[filaactual, 12] = row[11];
                    hoja.Cells[filaactual, 13] = row[12];
                    hoja.Cells[filaactual, 14] = row[13];

                    filaactual++;
                }


                int rowdatagrid = 7;
                foreach (DataRow row in Reporte_Entradas.Rows)
                {
                    string dato = Convert.ToString(row["ESTATUS"]);
                    if (dato == "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 14]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        r.Font.Bold = true;

                    }

                    string datoenvase = Convert.ToString(row["NOMBREENVASE"]);
                    if (datoenvase == "TOTAL")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 14]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);
                        r.Font.Bold = true;

                    }

                    rowdatagrid++;
                }
                // Liberación: no hay interfaz tipada, pero debes liberar los objetos COM igual
                if (aplicacion != null) Marshal.ReleaseComObject(aplicacion);
                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;

                FormMENSAJE.Close();
            }
            else if (radioButton2.Checked == true)
            {
                FormMENSAJE.Show();
                thisConnecion.Open();
                string query = "SELECT * FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE = '" + clave_envase + "' ORDER BY FECHA, ENV_CLAVE";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                int totalsal = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Salidas.NewRow();
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["ESTATUS"] = Convert.ToString(dr["SAL_STATUS"]).Trim();
                    Reporte_Salidas.Rows.Add(rowi);


                    if (Convert.ToString(dr["SAL_STATUS"]).Trim() != "C")
                    {
                        totalsal = totalsal + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }

                DataRow rowix = Reporte_Salidas.NewRow();
                rowix["FOLIO"] = "";
                rowix["FECHA"] = "";
                rowix["PROVEEDOR_CLAVE"] = "";
                rowix["PROVEEDOR_NOMBRE"] = "";
                rowix["RANCHO_CLAVE"] = "";
                rowix["RANCHO_NOMBRE"] = "";
                rowix["TABLA_CLAVE"] = "";
                rowix["TABLA_NOMBRE"] = "";
                rowix["IDENVASE"] = "";
                rowix["NOMBREENVASE"] = "TOTAL";
                rowix["CANTIDAD"] = totalsal;
                rowix["PROD_CLAVE"] = "";
                rowix["PROD_NOMBRE"] = "";
                rowix["ESTATUS"] = "";
                Reporte_Salidas.Rows.Add(rowix);

                thisConnecion.Close();


                #region fix issue: dll Interoperabilidad de Excel
                // En lugar de:
                // Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                // Usa:
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic aplicacion = Activator.CreateInstance(excelType);

                // A partir de aquí, todo se maneja con dynamic
                dynamic libro = aplicacion.Workbooks.Add();
                dynamic hoja = libro.Worksheets[1];

                // Las llamadas a propiedades y métodos son resueltas en tiempo de ejecución (IDispatch)
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                dynamic r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                //rango.Font.Bold = true;
                //rango.Font.Size = 16;
                //rango.MergeCells = true;

                // ... resto del código similar, usando dynamic en lugar de tipos concretos

                // Para hacer visible
                aplicacion.Visible = true;

                // Liberación: no hay interfaz tipada, pero debes liberar los objetos COM igual
                if (aplicacion != null) Marshal.ReleaseComObject(aplicacion);
                #endregion


                //Microsoft.Office.Interop.Excel.Application aplicacion;
                //Microsoft.Office.Interop.Excel.Workbook libro;
                //Microsoft.Office.Interop.Excel.Worksheet hoja;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libro = aplicacion.Workbooks.Add();
                //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range r;
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                r.Font.Bold = true;
                r.Font.Size = 16;
                r.MergeCells = true;

                hoja.Cells[3, 3] = "REPORTE GENERAL DE SALIDAS DE ENVASES DEL " + FechaInirepo.Text + " AL " + FechaFinrepo.Text;
                r = hoja.Range[hoja.Cells[3, 3], hoja.Cells[3, 8]];
                r.Font.Bold = true;
                r.MergeCells = true;

                string ruta = "c:\\SisGabWeb\\logo.png";
                //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 0, 70, 70);
                hoja.Range[hoja.Cells[1, 2], hoja.Cells[4, 2]].Merge();
                r = hoja.get_Range("D1", "D3");



                hoja.Cells[6, 1] = "Folio";
                hoja.Cells[6, 2] = "Fecha";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Clave Rancho";
                hoja.Cells[6, 6] = "Nombre Rancho";
                hoja.Cells[6, 7] = "Clave Tabla";
                hoja.Cells[6, 8] = "Nombre Tabla";
                hoja.Cells[6, 9] = "Envase Clave";
                hoja.Cells[6, 10] = "Envase Nombre";
                hoja.Cells[6, 11] = "Cantidad";
                hoja.Cells[6, 12] = "Producto Clave";
                hoja.Cells[6, 13] = "producto Nombre";


                r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 14]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;

                r = hoja.Range[hoja.Cells[7, 3], hoja.Cells[Reporte_Salidas.Rows.Count + 10, 13]];
                r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                int filaactual = 7;
                foreach (DataRow row in Reporte_Salidas.Rows)
                {   //Here 2 cell is target value and 1 cell is Volume 

                    hoja.Cells[filaactual, 1] = row[0];
                    hoja.Cells[filaactual, 2] = row[1];
                    hoja.Cells[filaactual, 3] = row[2];
                    hoja.Cells[filaactual, 4] = row[3];
                    hoja.Cells[filaactual, 5] = row[4];
                    hoja.Cells[filaactual, 6] = row[5];
                    hoja.Cells[filaactual, 7] = row[6];
                    hoja.Cells[filaactual, 8] = row[7];
                    hoja.Cells[filaactual, 9] = row[8];
                    hoja.Cells[filaactual, 10] = row[9];
                    hoja.Cells[filaactual, 11] = row[10];
                    hoja.Cells[filaactual, 12] = row[11];
                    hoja.Cells[filaactual, 13] = row[12];

                    filaactual++;
                }


                int rowdatagrid = 7;
                foreach (DataRow row in Reporte_Salidas.Rows)
                {
                    string dato = Convert.ToString(row["ESTATUS"]);
                    if (dato == "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 13]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        r.Font.Bold = true;

                    }

                    string datoenvase = Convert.ToString(row["NOMBREENVASE"]);
                    if (datoenvase == "TOTAL")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 13]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);
                        r.Font.Bold = true;

                    }

                    rowdatagrid++;
                }

                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;


                FormMENSAJE.Close();


            }
            else
            {
                FormMENSAJE.Show();
                thisConnecion.Open();
                string query = "SELECT * FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE = '" + clave_envase + "' ORDER BY FECHA, ENV_CLAVE";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                int totalent = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Kardex.NewRow();
                    rowi["NUMERO"] = 0;
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["RECIBO_MP"] = Convert.ToString(dr["RMP_FOLIO"]).Trim();
                    rowi["TIPO"] = "ENTRADA";
                    rowi["ESTATUS"] = Convert.ToString(dr["ENT_STATUS"]).Trim();
                    Reporte_Kardex.Rows.Add(rowi);


                    if (Convert.ToString(dr["ENT_STATUS"]).Trim() != "C")
                    {
                        totalent = totalent + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }


                thisConnecion.Close();

                //Ingresar Salidas

                thisConnecion.Open();
                query = "SELECT * FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(FechaInirepo.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(FechaFinrepo.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE = '" + clave_envase + "' ORDER BY FECHA, ENV_CLAVE";
                cm = new SqlCommand(query, thisConnecion);
                dr = cm.ExecuteReader();

                int totalsal = 0;

                while (dr.Read())
                {
                    DataRow rowi = Reporte_Kardex.NewRow();
                    rowi["NUMERO"] = 0;
                    rowi["FOLIO"] = Convert.ToString(dr["FOLIO"]).Trim();
                    rowi["FECHA"] = Convert.ToDateTime(dr["FECHA"]).ToString("dd/MM/yyyy");
                    rowi["PROVEEDOR_CLAVE"] = Convert.ToString(dr["PROV_CLAVE"]).Trim();
                    rowi["PROVEEDOR_NOMBRE"] = Convert.ToString(dr["PROV_NOMBRE"]).Trim();
                    rowi["RANCHO_CLAVE"] = Convert.ToString(dr["RCH_CLAVE"]).Trim();
                    rowi["RANCHO_NOMBRE"] = Convert.ToString(dr["RCH_NOMBRE"]).Trim();
                    rowi["TABLA_CLAVE"] = Convert.ToString(dr["TBL_CLAVE"]).Trim();
                    rowi["TABLA_NOMBRE"] = Convert.ToString(dr["TBL_NOMBRE"]).Trim();
                    rowi["IDENVASE"] = Convert.ToString(dr["ENV_CLAVE"]).Trim();
                    rowi["NOMBREENVASE"] = Convert.ToString(dr["ENV_NOMBRE"]).Trim();
                    rowi["CANTIDAD"] = Convert.ToString(dr["CANTIDAD"]).Trim();
                    rowi["PROD_CLAVE"] = Convert.ToString(dr["PROD_CLAVE"]).Trim();
                    rowi["PROD_NOMBRE"] = Convert.ToString(dr["PROD_NOMBRE"]).Trim();
                    rowi["RECIBO_MP"] = "";
                    rowi["TIPO"] = "SALIDA";
                    rowi["ESTATUS"] = Convert.ToString(dr["SAL_STATUS"]).Trim();
                    Reporte_Kardex.Rows.Add(rowi);


                    if (Convert.ToString(dr["SAL_STATUS"]).Trim() != "C")
                    {
                        totalsal = totalsal + Convert.ToInt32(dr["CANTIDAD"]);
                    }
                }

                DataRow rowix = Reporte_Kardex.NewRow();
                rowix["NUMERO"] = 1;
                rowix["FOLIO"] = "";
                rowix["FECHA"] = "";
                rowix["PROVEEDOR_CLAVE"] = "";
                rowix["PROVEEDOR_NOMBRE"] = "";
                rowix["RANCHO_CLAVE"] = "TOTALES";
                rowix["RANCHO_NOMBRE"] = "";
                rowix["TABLA_CLAVE"] = "ENTRADAS";
                rowix["TABLA_NOMBRE"] = totalent;
                rowix["IDENVASE"] = "";
                rowix["NOMBREENVASE"] = "SALIDAS";
                rowix["CANTIDAD"] = totalsal;
                rowix["PROD_CLAVE"] = "";
                rowix["PROD_NOMBRE"] = "";
                rowix["RECIBO_MP"] = "";
                rowix["TIPO"] = "";
                rowix["ESTATUS"] = "";
                Reporte_Kardex.Rows.Add(rowix);

                thisConnecion.Close();


                Reporte_Kardex.DefaultView.Sort = "NUMERO, FECHA, Tipo, IDENVASE";
                DataView dv = Reporte_Kardex.DefaultView;


                #region fix issue: dll Interoperabilidad de Excel
                // En lugar de:
                // Microsoft.Office.Interop.Excel.Application aplicacion = new Microsoft.Office.Interop.Excel.Application();
                // Usa:
                Type excelType = Type.GetTypeFromProgID("Excel.Application");
                dynamic aplicacion = Activator.CreateInstance(excelType);

                // A partir de aquí, todo se maneja con dynamic
                dynamic libro = aplicacion.Workbooks.Add();
                dynamic hoja = libro.Worksheets[1];

                // Las llamadas a propiedades y métodos son resueltas en tiempo de ejecución (IDispatch)
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                dynamic rango = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                rango.Font.Bold = true;
                rango.Font.Size = 16;
                rango.MergeCells = true;

                // ... resto del código similar, usando dynamic en lugar de tipos concretos

                // Para hacer visible
                aplicacion.Visible = true;

                // Liberación: no hay interfaz tipada, pero debes liberar los objetos COM igual
                if (aplicacion != null) Marshal.ReleaseComObject(aplicacion);
                #endregion


                //Microsoft.Office.Interop.Excel.Application aplicacion;
                //Microsoft.Office.Interop.Excel.Workbook libro;
                //Microsoft.Office.Interop.Excel.Worksheet hoja;
                aplicacion = new Microsoft.Office.Interop.Excel.Application();
                libro = aplicacion.Workbooks.Add();
                //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
                hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

                Microsoft.Office.Interop.Excel.Range r;
                hoja.Cells[2, 3] = "Comercializador GAB, S.A. de C.V.";
                r = hoja.Range[hoja.Cells[2, 3], hoja.Cells[2, 8]];
                r.Font.Bold = true;
                r.Font.Size = 16;
                r.MergeCells = true;

                hoja.Cells[3, 3] = "REPORTE GENERAL DE ENTRADAS Y SALIDAS DE ENVASES DEL " + FechaInirepo.Text + " AL " + FechaFinrepo.Text;
                r = hoja.Range[hoja.Cells[3, 3], hoja.Cells[3, 8]];
                r.Font.Bold = true;
                r.MergeCells = true;


                string ruta = "c:\\SisGabWeb\\logo.png";
                //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 20, 0, 70, 70);
                hoja.Range[hoja.Cells[1, 2], hoja.Cells[4, 2]].Merge();
                r = hoja.get_Range("D1", "D3");

                hoja.Cells[6, 1] = "Folio";
                hoja.Cells[6, 2] = "Fecha";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Clave Rancho";
                hoja.Cells[6, 6] = "Nombre Rancho";
                hoja.Cells[6, 7] = "Clave Tabla";
                hoja.Cells[6, 8] = "Nombre Tabla";
                hoja.Cells[6, 9] = "Envase Clave";
                hoja.Cells[6, 10] = "Envase Nombre";
                hoja.Cells[6, 11] = "Cantidad";
                hoja.Cells[6, 12] = "Producto Clave";
                hoja.Cells[6, 13] = "producto Nombre";
                hoja.Cells[6, 14] = "Recibo MP";
                hoja.Cells[6, 15] = "TIPO MOVIMIENTO";


                r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 15]];
                r.Font.Bold = true;

                Cursor.Current = Cursors.WaitCursor;

                r = hoja.Range[hoja.Cells[7, 3], hoja.Cells[Reporte_Kardex.Rows.Count + 10, 15]];
                r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

                int filaactual = 7;
                foreach (DataRowView row in dv)
                {   //Here 2 cell is target value and 1 cell is Volume 

                    hoja.Cells[filaactual, 1] = row[1];
                    hoja.Cells[filaactual, 2] = row[2];
                    hoja.Cells[filaactual, 3] = row[3];
                    hoja.Cells[filaactual, 4] = row[4];
                    hoja.Cells[filaactual, 5] = row[5];
                    hoja.Cells[filaactual, 6] = row[6];
                    hoja.Cells[filaactual, 7] = row[7];
                    hoja.Cells[filaactual, 8] = row[8];
                    hoja.Cells[filaactual, 9] = row[9];
                    hoja.Cells[filaactual, 10] = row[10];
                    hoja.Cells[filaactual, 11] = row[11];
                    hoja.Cells[filaactual, 12] = row[12];
                    hoja.Cells[filaactual, 13] = row[13];
                    hoja.Cells[filaactual, 14] = row[14];
                    hoja.Cells[filaactual, 15] = row[15];

                    filaactual++;
                }


                int rowdatagrid = 7;
                foreach (DataRowView row in dv)
                {
                    string dato = Convert.ToString(row["ESTATUS"]);
                    if (dato == "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                        r.Font.Bold = true;

                    }

                    string datotipo = Convert.ToString(row["TIPO"]);
                    if (datotipo == "ENTRADA" && dato != "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
                        r.Font.Bold = true;

                    }
                    else if (datotipo == "SALIDA" && dato != "C")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.IndianRed);
                        r.Font.Bold = true;

                    }

                    string datoenvase = Convert.ToString(row["RANCHO_CLAVE"]);
                    if (datoenvase == "TOTALES")// Or your condition 
                    {

                        r = hoja.Range[hoja.Cells[rowdatagrid, 1], hoja.Cells[rowdatagrid, 15]];
                        r.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LawnGreen);
                        r.Font.Bold = true;

                    }

                    rowdatagrid++;
                }

                aplicacion.Columns.AutoFit();
                aplicacion.Rows.AutoFit();
                aplicacion.Visible = true;


                FormMENSAJE.Close();



            }

        }

        private void button8_Click(object sender, EventArgs e)
        {
            bg = new BackgroundWorker();
            bg.WorkerReportsProgress = true;
            bg.ProgressChanged += bg_ProgressChanged;
            bg.DoWork += bg_DoWork;
            bg.RunWorkerCompleted += bg_RunWorkerCompleted;
            bg.RunWorkerAsync();
            recalcular.Visible = true;
            progress.Visible = true;
        }

        private void dateTimeinventarioinicial_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimeinventarioinicial.Text != "")
            {
                string dia_escogido = Convert.ToDateTime(dateTimeinventarioinicial.Text).ToString("dd/MM/yyyy");
                string mes_escogido = Convert.ToDateTime(dateTimeinventarioinicial.Text).ToString("MM");
                //if (mes_actual_Maquina == mes_escogido)
                //{
                thisConnecion.Open();
                string queryX = "SELECT * FROM  TB_MSTR_INV_ENVASES_dos INNER JOIN tb_cat_envases AS B ON TB_MSTR_INV_ENVASES_dos.ENV_CLAVE = B.env_clave WHERE (ENV_FECHA = '" + dia_escogido + "')";
                SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
                SqlDataReader drX = cmX.ExecuteReader();

                inv_Ini_dia.Clear();

                while (drX.Read())
                {
                    DataRow rowix = inv_Ini_dia.NewRow();
                    rowix["CLV"] = drX["ENV_CLAVE"];
                    rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                    rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                    rowix["CANTIDAD MODIFICADA"] = "";
                    inv_Ini_dia.Rows.Add(rowix);

                }

                thisConnecion.Close();

                //}
                //else
                //{
                //    DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                //    string nombreMes = formatoFecha.GetMonthName(Convert.ToInt32(mes_actual_Maquina));
                //    MessageBox.Show("Solo se pueden hacer cambios de inventario inicial del mes actual " + nombreMes);
                //    //dateTimeinventarioinicial.Text = fecha_hoy_hoy;
                //}

            }

        }

        private void bg_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progress.Value = e.ProgressPercentage;
            progress.Step = 1;
            progress.Style = ProgressBarStyle.Continuous;
            progress.Minimum = 0;
            progress.Maximum = 100;

            if (e.ProgressPercentage > 100)
            {
                recalcular.Text = "100%";
                progress.Value = progress.Maximum;
            }
            else
            {
                recalcular.Text = Convert.ToString(e.ProgressPercentage) + "%";
                progress.Value = e.ProgressPercentage;
            }
        }

        private void bg_DoWork(object sender, DoWorkEventArgs e)
        {
            int progreso = 0, porciento = 0, totalEmpleadosProcesar = 0;

            string fecha_edit = Convert.ToDateTime(dateTimeinventarioinicial.Text).ToString("dd/MM/yyyy");

            TimeSpan tspan = Convert.ToDateTime(fecha_hoy_hoy) - Convert.ToDateTime(dateTimeinventarioinicial.Text);

            int dias = tspan.Days;

            dias++;
            int veces = 0;

            foreach (DataGridViewRow row in dataGridInfoInvenario.Rows)
            {
                if (row.Cells["CLV"].Value != null)
                {
                    veces++;
                }
            }

            dias = dias * veces;

            foreach (DataGridViewRow row in dataGridInfoInvenario.Rows)
            {
                if (row.Cells["CLV"].Value != null)
                {
                    string idenvase = row.Cells["CLV"].Value.ToString();
                    string nomenvase = row.Cells["NOMBRE DEL ENVASE"].Value.ToString();
                    string cantenvase = row.Cells["CANTIDAD"].Value.ToString();
                    string newcantenvase = row.Cells["CANTIDAD MODIFICADA"].Value.ToString();
                    if (newcantenvase.Trim().Length > 0)
                    {
                        thisConnecion.Open();
                        string cadena = "UPDATE TB_MSTR_INV_ENVASES_dos SET ENV_INV_INI_CANT = '" + newcantenvase.Trim() + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + fecha_edit.Trim() + "' ";
                        SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
                        cmd.ExecuteNonQuery();
                        thisConnecion.Close();
                    }
                }
            }

            thisConnecion.Open();
            string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                            "Values('" + fecha_hoy_hoy + "','" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ACTUALIZAR','2.18','', 'MODIFICACION INVENTARIO INICIAL FECHA " + fecha_edit.Trim() + "', 'SIPGAB', '')";
            SqlCommand cmdxs = new SqlCommand(cadenaregmov, thisConnecion);
            cmdxs.ExecuteNonQuery();


            bool ciclo = true;

            while (ciclo == true)
            {
                if (Convert.ToDateTime(fecha_edit) > Convert.ToDateTime(fecha_hoy_hoy))
                {
                    ciclo = false;
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridInfoInvenario.Rows)
                    {
                        if (row.Cells["CLV"].Value != null)
                        {
                            string idenvase = row.Cells["CLV"].Value.ToString();
                            string nomenvase = row.Cells["NOMBRE DEL ENVASE"].Value.ToString();
                            string cantenvase = row.Cells["CANTIDAD"].Value.ToString();
                            string newcantenvase = row.Cells["CANTIDAD MODIFICADA"].Value.ToString();

                            string valida = "SELECT ISNULL(sum(ENV_INV_INI_CANT), 0) FROM TB_MSTR_INV_ENVASES_dos WHERE ENV_CLAVE = '" + idenvase + "' AND ENV_FECHA = '" + fecha_edit + "'";
                            SqlCommand cmd = new SqlCommand(valida, thisConnecion);
                            int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND ENT_STATUS != 'C'";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND SAL_STATUS != 'C'";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            string actualizadato = "";

                            if (Convert.ToDateTime(fecha_edit) < Convert.ToDateTime(fecha_hoy_hoy))
                            {
                                //thisConnecion.Open();
                                actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "'";
                                cmd = new SqlCommand(actualizadato, thisConnecion);
                                cmd.ExecuteNonQuery();
                                //thisConnecion.Close();

                            }



                            int inv_ini_hoy = inv_ini + entradas - salidas;

                            //thisConnecion.Open();
                            actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_INV_INI_CANT = '" + inv_ini_hoy + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy") + "' ";
                            cmd = new SqlCommand(actualizadato, thisConnecion);
                            cmd.ExecuteNonQuery();
                            //thisConnecion.Close();

                            totalEmpleadosProcesar = dias;  //Total de Empleados que seleccionó el usuario
                            progreso++; //Aumentando el progreso 
                            porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                            System.Threading.Thread.Sleep(50);
                            bg.ReportProgress(porciento);

                        }



                    }
                    fecha_edit = Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy");
                }
            }
            thisConnecion.Close();

        }

        private void bg_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            thisConnecion.Open();
            inv_Ini_dia.Clear();
            string queryX = "SELECT * FROM  TB_MSTR_INV_ENVASES_dos INNER JOIN tb_cat_envases AS B ON TB_MSTR_INV_ENVASES_dos.ENV_CLAVE = B.env_clave WHERE (ENV_FECHA = '" + Convert.ToDateTime(dateTimeinventarioinicial.Text).ToString("dd/MM/yyyy") + "')";
            SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
            SqlDataReader drX = cmX.ExecuteReader();

            while (drX.Read())
            {
                DataRow rowix = inv_Ini_dia.NewRow();
                rowix["CLV"] = drX["ENV_CLAVE"];
                rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                rowix["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia.Rows.Add(rowix);

            }

            thisConnecion.Close();

            dataGridInfoInvenario.AutoGenerateColumns = true;
            dataGridInfoInvenario.DataSource = inv_Ini_dia;
            dataGridInfoInvenario.AutoResizeColumns();

            dataGridInfoInvenario.Columns["CLV"].ReadOnly = true;
            dataGridInfoInvenario.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            dataGridInfoInvenario.Columns["CANTIDAD"].ReadOnly = true;

            dataGridInfoInvenario.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            MessageBox.Show("INFORMACION ACTUALIZADA CORRECTAMENTE");
            recalcular.Visible = false;
            progress.Visible = false;
        }

        private void Movimientos_Click(object sender, EventArgs e)
        {
            Reporte_Movimientos reporte = new Reporte_Movimientos();
            reporte.Show();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void clbtabla_Leave(object sender, EventArgs e)
        {
            Cmb_Tabla.SelectedValue = clbtabla.Text.ToString();
        }

        private void clb_tabla_entr_Leave(object sender, EventArgs e)
        {
            cmb_tabla_entr.SelectedValue = clb_tabla_entr.Text.ToString();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            int Total_INVENTARIO_Ini = 0;
            int Total_Entradas = 0;
            int Total_Salidas = 0;
            int Total_Inventario_Final = 0;


            thisConnecion.Open();
            string fecha = "SELECT SYSDATETIME()";
            SqlCommand cmd = new SqlCommand(fecha, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();

            /*thisConnecion.Open();
            string valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos where ENV_FECHA NOT IN ('"+ fecha_hoy +"') ORDER BY ENV_FECHA DESC";
            cmd = new SqlCommand(valida, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();*/


            thisConnecion.Open();
            string Cadena = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + fecha_ent + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "INVTENV");
            SqlCommand cmdx = new SqlCommand(Cadena);
            cmdx.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmdx.ExecuteReader();
            inv_x_proveedorini_CAJONES.Clear();

            inv_x_proveedorini_CAJONES.Load(Info);
            DataColumn column;

            thisConnecion.Close();

            int T = 0;
            if (inv_x_proveedorini_CAJONES.Rows.Count == 0)
            {
                MessageBox.Show("CORTE DIARIO NO REALIZADO, NO SE PUEDE MOSTRAR INFORMACION", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            inv_x_proveedor_CAJONES.Clear();
            int total = 1;

            foreach (DataRow row in inv_x_proveedorini_CAJONES.Rows)
            {

                thisConnecion.Open();
                string valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND ENT_STATUS != 'C' AND PROV_CLAVE = '" + row["PROV_CLAVE"] + "'";
                cmd = new SqlCommand(valida, thisConnecion);
                int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                //thisConnecion.Close();

                //thisConnecion.Open();
                valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND SAL_STATUS != 'C' AND PROV_CLAVE = '" + row["PROV_CLAVE"] + "'";
                cmd = new SqlCommand(valida, thisConnecion);
                int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                thisConnecion.Close();

                if (total == 1)
                {
                    DataRow rowix = inv_x_proveedor_CAJONES.NewRow();
                    rowix["ID"] = "";
                    rowix["ID_PROVEEDOR"] = row["PROV_CLAVE"];
                    rowix["NOMBRE"] = row["PROV_NOMBRE"];
                    rowix["FECHA"] = "";
                    rowix["INVENTARIO INICIAL"] = "";
                    rowix["ENTRADAS"] = "";
                    rowix["SALIDAS"] = "";
                    rowix["INVENTARIO FINAL"] = "";
                    inv_x_proveedor_CAJONES.Rows.Add(rowix);
                    total = 0;
                }
                else
                {
                    total++;
                }

                DataRow rowi = inv_x_proveedor_CAJONES.NewRow();
                rowi["ID"] = row["ENV_CLAVE"];
                rowi["ID_PROVEEDOR"] = row["PROV_CLAVE"];
                rowi["NOMBRE"] = row["env_nombre"];
                rowi["FECHA"] = Convert.ToDateTime(row["ENV_FECHA"]).ToString("dd/MM/yyyy");
                rowi["INVENTARIO INICIAL"] = row["ENV_INV_INI_CANT"];
                rowi["ENTRADAS"] = Convert.ToInt32(entradas);
                rowi["SALIDAS"] = Convert.ToInt32(salidas);
                rowi["INVENTARIO FINAL"] = Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(entradas) - Convert.ToInt32(salidas);
                inv_x_proveedor_CAJONES.Rows.Add(rowi);

                Total_INVENTARIO_Ini = Total_INVENTARIO_Ini + Convert.ToInt16(row["ENV_INV_INI_CANT"]);
                Total_Entradas = Total_Entradas + Convert.ToInt32(entradas);
                Total_Salidas = Total_Salidas + Convert.ToInt32(salidas);
                Total_Inventario_Final = Total_Inventario_Final + Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(salidas) - Convert.ToInt32(entradas);


            }


            DataRow rowixw = inv_x_proveedor_CAJONES.NewRow();
            rowixw["ID"] = "";
            rowixw["ID_PROVEEDOR"] = "";
            rowixw["NOMBRE"] = "TOTALES";
            rowixw["FECHA"] = "";
            rowixw["INVENTARIO INICIAL"] = Total_INVENTARIO_Ini;
            rowixw["ENTRADAS"] = Total_Entradas;
            rowixw["SALIDAS"] = Total_Salidas;
            rowixw["INVENTARIO FINAL"] = Total_Inventario_Final;
            inv_x_proveedor_CAJONES.Rows.Add(rowixw);

            CAJONES ob = new CAJONES(inv_x_proveedor_CAJONES, 0);

            ob.Show();
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string valorCelda = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString();

            //nombre columna
            string nombrecolumna = dataGridView1.Columns[e.ColumnIndex].HeaderText;

            //obtienes el valor de la primer columna
            string valorPrimerCelda = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
        }

        private void dataGridCajones_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (this.dataGridCajones.Rows[e.RowIndex].Cells[0].Value.ToString() == "")
                {
                    e.CellStyle.BackColor = Color.LightGreen;
                }

            }
            catch
            {


            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            bgcajones = new BackgroundWorker();
            bgcajones.WorkerReportsProgress = true;
            bgcajones.ProgressChanged += bgcajones_ProgressChanged;
            bgcajones.DoWork += bgcajones_DoWork;
            bgcajones.RunWorkerCompleted += bgcajones_RunWorkerCompleted;
            bgcajones.RunWorkerAsync();
            labelcajones.Visible = true;
            progressBarcajones.Visible = true;
        }


        private void bgcajones_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBarcajones.Value = e.ProgressPercentage;
            progressBarcajones.Step = 1;
            progressBarcajones.Style = ProgressBarStyle.Continuous;
            progressBarcajones.Minimum = 0;
            progressBarcajones.Maximum = 100;

            if (e.ProgressPercentage > 100)
            {
                labelcajones.Text = "100%";
                progressBarcajones.Value = progress.Maximum;
            }
            else
            {
                labelcajones.Text = Convert.ToString(e.ProgressPercentage) + "%";
                progressBarcajones.Value = e.ProgressPercentage;
            }
        }

        private void bgcajones_DoWork(object sender, DoWorkEventArgs e)
        {
            int progreso = 0, porciento = 0, totalEmpleadosProcesar = 0;

            string fecha_edit = Convert.ToDateTime(dateTimecajones.Text).ToString("dd/MM/yyyy");

            TimeSpan tspan = Convert.ToDateTime(fecha_hoy_hoy) - Convert.ToDateTime(dateTimecajones.Text);

            int dias = tspan.Days;

            dias++;
            int veces = 0;

            foreach (DataGridViewRow row in dataGridCajones.Rows)
            {
                if (row.Cells["CLV"].Value != null)
                {
                    veces++;
                }
            }

            dias = dias * veces;

            foreach (DataGridViewRow row in dataGridCajones.Rows)
            {
                if (row.Cells["CLV"].Value != null)
                {
                    string idenvase = row.Cells["CLV"].Value.ToString();
                    string idproveedor = row.Cells["CLV_PROV"].Value.ToString();
                    string nomenvase = row.Cells["NOMBRE DEL ENVASE"].Value.ToString();
                    string cantenvase = row.Cells["CANTIDAD"].Value.ToString();
                    string newcantenvase = row.Cells["CANTIDAD MODIFICADA"].Value.ToString();
                    if (newcantenvase.Trim().Length > 0)
                    {
                        thisConnecion.Open();
                        string cadena = "UPDATE  TB_MSTR_INV_CAJAS_PROV SET ENV_INV_INI_CANT = '" + newcantenvase.Trim() + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + fecha_edit.Trim() + "' AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                        SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
                        cmd.ExecuteNonQuery();
                        thisConnecion.Close();
                    }
                }
            }

            thisConnecion.Open();
            string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                            "Values(GetDate(),'" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ACTUALIZAR','2.18','', 'MODIFICACION INVENTARIO CAJAS INICIAL FECHA " + fecha_edit.Trim() + "', 'SIPGAB', '')";
            SqlCommand cmdxs = new SqlCommand(cadenaregmov, thisConnecion);
            cmdxs.ExecuteNonQuery();


            bool ciclo = true;

            while (ciclo == true)
            {
                if (Convert.ToDateTime(fecha_edit) > Convert.ToDateTime(fecha_hoy_hoy))
                {
                    ciclo = false;
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridCajones.Rows)
                    {
                        if (row.Cells["CLV"].Value != null)
                        {
                            string idenvase = row.Cells["CLV"].Value.ToString();
                            string idproveedor = row.Cells["CLV_PROV"].Value.ToString();
                            string nomenvase = row.Cells["NOMBRE DEL ENVASE"].Value.ToString();
                            string cantenvase = row.Cells["CANTIDAD"].Value.ToString();
                            string newcantenvase = row.Cells["CANTIDAD MODIFICADA"].Value.ToString();

                            string valida = "SELECT ISNULL(sum(ENV_INV_INI_CANT), 0) FROM TB_MSTR_INV_CAJAS_PROV WHERE ENV_CLAVE = '" + idenvase + "' AND ENV_FECHA = '" + fecha_edit + "'  AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                            SqlCommand cmd = new SqlCommand(valida, thisConnecion);
                            int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND ENT_STATUS != 'C' AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND SAL_STATUS != 'C' AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            string actualizadato = "";

                            if (Convert.ToDateTime(fecha_edit) < Convert.ToDateTime(fecha_hoy_hoy))
                            {
                                //thisConnecion.Open();
                                actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "'  AND PROV_CLAVE = '" + idproveedor.Trim() + "'";
                                cmd = new SqlCommand(actualizadato, thisConnecion);
                                cmd.ExecuteNonQuery();
                                //thisConnecion.Close();

                            }



                            int inv_ini_hoy = inv_ini + entradas - salidas;

                            //thisConnecion.Open();
                            actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV SET  ENV_INV_INI_CANT = '" + inv_ini_hoy + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy") + "'  AND PROV_CLAVE = '" + idproveedor.Trim() + "'";
                            cmd = new SqlCommand(actualizadato, thisConnecion);
                            cmd.ExecuteNonQuery();
                            //thisConnecion.Close();

                            totalEmpleadosProcesar = dias;  //Total de Empleados que seleccionó el usuario
                            progreso++; //Aumentando el progreso 
                            porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                            System.Threading.Thread.Sleep(50);
                            bgcajones.ReportProgress(porciento);

                        }



                    }
                    fecha_edit = Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy");
                }
            }
            thisConnecion.Close();

        }

        private void bgcajones_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {

            thisConnecion.Open();
            inv_Ini_dia_cajon.Clear();
            string queryX = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + Convert.ToDateTime(dateTimecajones.Text).ToString("dd/MM/yyyy") + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
            SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
            SqlDataReader drX = cmX.ExecuteReader();
            int cont = 1;
            while (drX.Read())
            {
                if (cont == 1)
                {
                    DataRow rowixcajon = inv_Ini_dia_cajon.NewRow();
                    rowixcajon["CLV"] = "";
                    rowixcajon["CLV_PROV"] = drX["prov_clave"];
                    rowixcajon["NOMBRE DEL ENVASE"] = drX["prov_nombre"];
                    rowixcajon["CANTIDAD"] = "";
                    rowixcajon["CANTIDAD MODIFICADA"] = "";
                    inv_Ini_dia_cajon.Rows.Add(rowixcajon);
                    cont = 0;

                }
                else
                {
                    cont++;

                }

                DataRow rowix = inv_Ini_dia_cajon.NewRow();
                rowix["CLV"] = drX["ENV_CLAVE"];
                rowix["CLV_PROV"] = drX["prov_clave"];
                rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                rowix["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia_cajon.Rows.Add(rowix);

            }

            thisConnecion.Close();

            dataGridCajones.AutoGenerateColumns = true;
            dataGridCajones.DataSource = inv_Ini_dia_cajon;
            dataGridCajones.AutoResizeColumns();

            dataGridCajones.Columns["CLV"].ReadOnly = true;
            dataGridCajones.Columns["CLV_PROV"].ReadOnly = true;
            dataGridCajones.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            dataGridCajones.Columns["CANTIDAD"].ReadOnly = true;

            dataGridCajones.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            MessageBox.Show("INFORMACION ACTUALIZADA CORRECTAMENTE");
            labelcajones.Visible = false;
            progressBarcajones.Visible = false;
        }

        private void dateTimecajones_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimecajones.Text != "")
            {
                string dia_escogido = Convert.ToDateTime(dateTimecajones.Text).ToString("dd/MM/yyyy");
                string mes_escogido = Convert.ToDateTime(dateTimecajones.Text).ToString("MM");
                if (mes_actual_Maquina == mes_escogido)
                {
                    thisConnecion.Open();
                    inv_Ini_dia_cajon.Clear();
                    string queryX = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + Convert.ToDateTime(dateTimecajones.Text).ToString("dd/MM/yyyy") + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
                    SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
                    SqlDataReader drX = cmX.ExecuteReader();
                    int cont = 1;
                    while (drX.Read())
                    {
                        if (cont == 1)
                        {
                            DataRow rowixcajon = inv_Ini_dia_cajon.NewRow();
                            rowixcajon["CLV"] = "";
                            rowixcajon["CLV_PROV"] = drX["prov_clave"];
                            rowixcajon["NOMBRE DEL ENVASE"] = drX["prov_nombre"];
                            rowixcajon["CANTIDAD"] = "";
                            rowixcajon["CANTIDAD MODIFICADA"] = "";
                            inv_Ini_dia_cajon.Rows.Add(rowixcajon);
                            cont = 0;

                        }
                        else
                        {
                            cont++;

                        }

                        DataRow rowix = inv_Ini_dia_cajon.NewRow();
                        rowix["CLV"] = drX["ENV_CLAVE"];
                        rowix["CLV_PROV"] = drX["prov_clave"];
                        rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                        rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                        rowix["CANTIDAD MODIFICADA"] = "";
                        inv_Ini_dia_cajon.Rows.Add(rowix);

                    }

                    thisConnecion.Close();

                }
                else
                {
                    DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                    string nombreMes = formatoFecha.GetMonthName(Convert.ToInt32(mes_actual_Maquina));
                    MessageBox.Show("Solo se pueden hacer cambios de inventario inicial del mes actual " + nombreMes);
                    dateTimeinventarioinicial.Text = fecha_hoy_hoy;
                }

            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            int Total_INVENTARIO_Ini = 0;
            int Total_Entradas = 0;
            int Total_Salidas = 0;
            int Total_Inventario_Final = 0;

            thisConnecion.Open();
            string fecha = "SELECT SYSDATETIME()";
            SqlCommand cmd = new SqlCommand(fecha, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();

            /*thisConnecion.Open();
            string valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos where ENV_FECHA NOT IN ('"+ fecha_hoy +"') ORDER BY ENV_FECHA DESC";
            cmd = new SqlCommand(valida, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();*/


            thisConnecion.Open();
            string Cadena = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ_ESPARRAGO AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV_ESPARRAGO AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + fecha_ent + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
            //string Cadena = "SELECT EMB_FOLIO FROM TB_MSTR_EMBARQUE WHERE HORA_TRAILER = '" + Program.MyGlobal.PubFecEmb + "' AND NO_TRAILER = '" + Program.MyGlobal.PubNoTrailer + "' ORDER BY EMB_FOLIO";
            DataSet ds1 = new DataSet();
            SqlDataAdapter da1 = new SqlDataAdapter(Cadena, thisConnecion);
            da1.Fill(ds1, "INVTENV");
            SqlCommand cmdx = new SqlCommand(Cadena);
            cmdx.Connection = thisConnecion;
            SqlDataReader Info;
            Info = cmdx.ExecuteReader();
            inv_x_proveedorini_CAJONES_ESPARRAGO.Clear();

            inv_x_proveedorini_CAJONES_ESPARRAGO.Load(Info);
            DataColumn column;

            thisConnecion.Close();

            int T = 0;
            if (inv_x_proveedorini_CAJONES_ESPARRAGO.Rows.Count == 0)
            {
                MessageBox.Show("CORTE DIARIO NO REALIZADO, NO SE PUEDE MOSTRAR INFORMACION", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            inv_x_proveedor_CAJONES_ESPARRAGO.Clear();
            int total = 1;

            foreach (DataRow row in inv_x_proveedorini_CAJONES_ESPARRAGO.Rows)
            {

                thisConnecion.Open();
                string valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND ENT_STATUS != 'C' AND PROV_CLAVE = '" + row["PROV_CLAVE"] + "'";
                cmd = new SqlCommand(valida, thisConnecion);
                int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                //thisConnecion.Close();

                //thisConnecion.Open();
                valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND SAL_STATUS != 'C' AND PROV_CLAVE = '" + row["PROV_CLAVE"] + "'";
                cmd = new SqlCommand(valida, thisConnecion);
                int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                thisConnecion.Close();


                DataRow rowix = inv_x_proveedor_CAJONES_ESPARRAGO.NewRow();
                rowix["ID"] = "";
                rowix["ID_PROVEEDOR"] = row["PROV_CLAVE"];
                rowix["NOMBRE"] = row["PROV_NOMBRE"];
                rowix["FECHA"] = "";
                rowix["INVENTARIO INICIAL"] = "";
                rowix["ENTRADAS"] = "";
                rowix["SALIDAS"] = "";
                rowix["INVENTARIO FINAL"] = "";
                inv_x_proveedor_CAJONES_ESPARRAGO.Rows.Add(rowix);


                DataRow rowi = inv_x_proveedor_CAJONES_ESPARRAGO.NewRow();
                rowi["ID"] = row["ENV_CLAVE"];
                rowi["ID_PROVEEDOR"] = row["PROV_CLAVE"];
                rowi["NOMBRE"] = row["env_nombre"];
                rowi["FECHA"] = Convert.ToDateTime(row["ENV_FECHA"]).ToString("dd/MM/yyyy");
                rowi["INVENTARIO INICIAL"] = row["ENV_INV_INI_CANT"];
                rowi["ENTRADAS"] = Convert.ToInt32(entradas);
                rowi["SALIDAS"] = Convert.ToInt32(salidas);
                rowi["INVENTARIO FINAL"] = Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(salidas) - Convert.ToInt32(entradas);
                inv_x_proveedor_CAJONES_ESPARRAGO.Rows.Add(rowi);

                Total_INVENTARIO_Ini = Total_INVENTARIO_Ini + Convert.ToInt16(row["ENV_INV_INI_CANT"]);
                Total_Entradas = Total_Entradas + Convert.ToInt32(entradas);
                Total_Salidas = Total_Salidas + Convert.ToInt32(salidas);
                Total_Inventario_Final = Total_Inventario_Final + Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(salidas) - Convert.ToInt32(entradas);
            }

            DataRow rowixw = inv_x_proveedor_CAJONES_ESPARRAGO.NewRow();
            rowixw["ID"] = "";
            rowixw["ID_PROVEEDOR"] = "";
            rowixw["NOMBRE"] = "TOTALES";
            rowixw["FECHA"] = "";
            rowixw["INVENTARIO INICIAL"] = Total_INVENTARIO_Ini;
            rowixw["ENTRADAS"] = Total_Entradas;
            rowixw["SALIDAS"] = Total_Salidas;
            rowixw["INVENTARIO FINAL"] = Total_Inventario_Final;
            inv_x_proveedor_CAJONES_ESPARRAGO.Rows.Add(rowixw);

            ESPARRAGO ob = new ESPARRAGO(inv_x_proveedor_CAJONES_ESPARRAGO, 0);

            ob.Show();
        }

        private void dateTimePickerEsparrago_ValueChanged(object sender, EventArgs e)
        {
            if (dateTimePickerEsparrago.Text != "")
            {
                string dia_escogido = Convert.ToDateTime(dateTimePickerEsparrago.Text).ToString("dd/MM/yyyy");
                string mes_escogido = Convert.ToDateTime(dateTimePickerEsparrago.Text).ToString("MM");
                if (mes_actual_Maquina == mes_escogido)
                {
                    thisConnecion.Open();
                    inv_Ini_dia_cajon_ESPARRAGO.Clear();
                    string queryX = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ_ESPARRAGO AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV_ESPARRAGO AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + Convert.ToDateTime(dateTimePickerEsparrago.Text).ToString("dd/MM/yyyy") + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
                    SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
                    SqlDataReader drX = cmX.ExecuteReader();
                    int cont = 1;
                    while (drX.Read())
                    {

                        DataRow rowixcajon = inv_Ini_dia_cajon_ESPARRAGO.NewRow();
                        rowixcajon["CLV"] = "";
                        rowixcajon["CLV_PROV"] = drX["prov_clave"];
                        rowixcajon["NOMBRE DEL ENVASE"] = drX["prov_nombre"];
                        rowixcajon["CANTIDAD"] = "";
                        rowixcajon["CANTIDAD MODIFICADA"] = "";
                        inv_Ini_dia_cajon_ESPARRAGO.Rows.Add(rowixcajon);


                        DataRow rowix = inv_Ini_dia_cajon_ESPARRAGO.NewRow();
                        rowix["CLV"] = drX["ENV_CLAVE"];
                        rowix["CLV_PROV"] = drX["prov_clave"];
                        rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                        rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                        rowix["CANTIDAD MODIFICADA"] = "";
                        inv_Ini_dia_cajon_ESPARRAGO.Rows.Add(rowix);

                    }

                    thisConnecion.Close();

                }
                else
                {
                    DateTimeFormatInfo formatoFecha = CultureInfo.CurrentCulture.DateTimeFormat;
                    string nombreMes = formatoFecha.GetMonthName(Convert.ToInt32(mes_actual_Maquina));
                    MessageBox.Show("Solo se pueden hacer cambios de inventario inicial del mes actual " + nombreMes);
                    dateTimeinventarioinicial.Text = fecha_hoy_hoy;
                }

            }
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            bgESPARRAGO = new BackgroundWorker();
            bgESPARRAGO.WorkerReportsProgress = true;
            bgESPARRAGO.ProgressChanged += bgESPARRAGO_ProgressChanged;
            bgESPARRAGO.DoWork += bgESPARRAGO_DoWork;
            bgESPARRAGO.RunWorkerCompleted += bgESPARRAGO_RunWorkerCompleted;
            bgESPARRAGO.RunWorkerAsync();
            labelEsparrago.Visible = true;
            progressBarEsparrago.Visible = true;
        }

        private void bgESPARRAGO_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            progressBarEsparrago.Value = e.ProgressPercentage;
            progressBarEsparrago.Step = 1;
            progressBarEsparrago.Style = ProgressBarStyle.Continuous;
            progressBarEsparrago.Minimum = 0;
            progressBarEsparrago.Maximum = 100;

            if (e.ProgressPercentage > 100)
            {
                labelEsparrago.Text = "100%";
                progressBarEsparrago.Value = progress.Maximum;
            }
            else
            {
                labelEsparrago.Text = Convert.ToString(e.ProgressPercentage) + "%";
                progressBarEsparrago.Value = e.ProgressPercentage;
            }
        }

        private void bgESPARRAGO_DoWork(object sender, DoWorkEventArgs e)
        {
            int progreso = 0, porciento = 0, totalEmpleadosProcesar = 0;

            string fecha_edit = Convert.ToDateTime(dateTimePickerEsparrago.Text).ToString("dd/MM/yyyy");

            TimeSpan tspan = Convert.ToDateTime(fecha_hoy_hoy) - Convert.ToDateTime(dateTimePickerEsparrago.Text);

            int dias = tspan.Days;

            dias++;
            int veces = 0;

            foreach (DataGridViewRow row in dataGridEsparrago.Rows)
            {
                if (row.Cells["CLV"].Value != null)
                {
                    veces++;
                }
            }

            dias = dias * veces;

            foreach (DataGridViewRow row in dataGridEsparrago.Rows)
            {
                if (row.Cells["CLV"].Value != null)
                {
                    string idenvase = row.Cells["CLV"].Value.ToString();
                    string idproveedor = row.Cells["CLV_PROV"].Value.ToString();
                    string nomenvase = row.Cells["NOMBRE DEL ENVASE"].Value.ToString();
                    string cantenvase = row.Cells["CANTIDAD"].Value.ToString();
                    string newcantenvase = row.Cells["CANTIDAD MODIFICADA"].Value.ToString();
                    if (newcantenvase.Trim().Length > 0)
                    {
                        thisConnecion.Open();
                        string cadena = "UPDATE  TB_MSTR_INV_CAJAS_PROV_ESPARRAGO SET ENV_INV_INI_CANT = '" + newcantenvase.Trim() + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + fecha_edit.Trim() + "' AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                        SqlCommand cmd = new SqlCommand(cadena, thisConnecion);
                        cmd.ExecuteNonQuery();
                        thisConnecion.Close();
                    }
                }
            }

            thisConnecion.Open();
            string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                            "Values(GetDate(),'" + Environment.MachineName.Trim() + "','" + usuario_actual + "','ACTUALIZAR','2.18','', 'MODIFICACION INVENTARIO CAJAS ESPARRAGO INICIAL FECHA " + fecha_edit.Trim() + "', 'SIPGAB', '')";
            SqlCommand cmdxs = new SqlCommand(cadenaregmov, thisConnecion);
            cmdxs.ExecuteNonQuery();


            bool ciclo = true;

            while (ciclo == true)
            {
                if (Convert.ToDateTime(fecha_edit) > Convert.ToDateTime(fecha_hoy_hoy))
                {
                    ciclo = false;
                }
                else
                {
                    foreach (DataGridViewRow row in dataGridEsparrago.Rows)
                    {
                        if (row.Cells["CLV"].Value != null)
                        {
                            string idenvase = row.Cells["CLV"].Value.ToString();
                            string idproveedor = row.Cells["CLV_PROV"].Value.ToString();
                            string nomenvase = row.Cells["NOMBRE DEL ENVASE"].Value.ToString();
                            string cantenvase = row.Cells["CANTIDAD"].Value.ToString();
                            string newcantenvase = row.Cells["CANTIDAD MODIFICADA"].Value.ToString();

                            string valida = "SELECT ISNULL(sum(ENV_INV_INI_CANT), 0) FROM TB_MSTR_INV_CAJAS_PROV_ESPARRAGO WHERE ENV_CLAVE = '" + idenvase + "' AND ENV_FECHA = '" + fecha_edit + "'  AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                            SqlCommand cmd = new SqlCommand(valida, thisConnecion);
                            int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND ENT_STATUS != 'C' AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            //thisConnecion.Open();
                            valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "' AND SAL_STATUS != 'C' AND PROV_CLAVE = '" + idproveedor.Trim() + "' ";
                            cmd = new SqlCommand(valida, thisConnecion);
                            int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                            //thisConnecion.Close();

                            string actualizadato = "";

                            if (Convert.ToDateTime(fecha_edit) < Convert.ToDateTime(fecha_hoy_hoy))
                            {
                                //thisConnecion.Open();
                                actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV_ESPARRAGO SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_edit + "' AND ENV_CLAVE = '" + idenvase + "'  AND PROV_CLAVE = '" + idproveedor.Trim() + "'";
                                cmd = new SqlCommand(actualizadato, thisConnecion);
                                cmd.ExecuteNonQuery();
                                //thisConnecion.Close();

                            }


                            //Linea de codigo Original Para El cierre de Canastillas de Esparrago
                            //int inv_ini_hoy = inv_ini + entradas - salidas;
                            //Se invierte posicion debido a que lo solicito Javier Castrejon, las Salidas Ahora seran positivas
                            int inv_ini_hoy = inv_ini + salidas - entradas;

                            //thisConnecion.Open();
                            actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV_ESPARRAGO SET  ENV_INV_INI_CANT = '" + inv_ini_hoy + "' WHERE ENV_CLAVE = '" + idenvase.Trim() + "' AND ENV_FECHA = '" + Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy") + "'  AND PROV_CLAVE = '" + idproveedor.Trim() + "'";
                            cmd = new SqlCommand(actualizadato, thisConnecion);
                            cmd.ExecuteNonQuery();
                            //thisConnecion.Close();

                            totalEmpleadosProcesar = dias;  //Total de Empleados que seleccionó el usuario
                            progreso++; //Aumentando el progreso 
                            porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                            System.Threading.Thread.Sleep(50);
                            bgESPARRAGO.ReportProgress(porciento);

                        }



                    }
                    fecha_edit = Convert.ToDateTime(fecha_edit).AddDays(1).ToString("dd/MM/yyyy");
                }
            }
            thisConnecion.Close();

        }

        private void bgESPARRAGO_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            thisConnecion.Open();
            inv_Ini_dia_cajon_ESPARRAGO.Clear();
            string queryX = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.* FROM tb_cat_proveedor AS A INNER JOIN Tb_ENV_PROV_CAJ_ESPARRAGO AS B ON A.prov_clave = B.cve_prov INNER JOIN TB_MSTR_INV_CAJAS_PROV_ESPARRAGO AS C ON B.cve_prov = C.PROV_CLAVE INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  where ENV_FECHA = '" + Convert.ToDateTime(dateTimePickerEsparrago.Text).ToString("dd/MM/yyyy") + "' ORDER BY A.prov_clave, D.ENV_CLAVE";
            SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
            SqlDataReader drX = cmX.ExecuteReader();
            int cont = 1;
            while (drX.Read())
            {
                DataRow rowixcajon = inv_Ini_dia_cajon_ESPARRAGO.NewRow();
                rowixcajon["CLV"] = "";
                rowixcajon["CLV_PROV"] = drX["prov_clave"];
                rowixcajon["NOMBRE DEL ENVASE"] = drX["prov_nombre"];
                rowixcajon["CANTIDAD"] = "";
                rowixcajon["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia_cajon_ESPARRAGO.Rows.Add(rowixcajon);

                DataRow rowix = inv_Ini_dia_cajon_ESPARRAGO.NewRow();
                rowix["CLV"] = drX["ENV_CLAVE"];
                rowix["CLV_PROV"] = drX["prov_clave"];
                rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                rowix["CANTIDAD"] = drX["ENV_INV_INI_CANT"];
                rowix["CANTIDAD MODIFICADA"] = "";
                inv_Ini_dia_cajon_ESPARRAGO.Rows.Add(rowix);

            }

            thisConnecion.Close();

            dataGridEsparrago.AutoGenerateColumns = true;
            dataGridEsparrago.DataSource = inv_Ini_dia_cajon_ESPARRAGO;
            dataGridEsparrago.AutoResizeColumns();

            dataGridEsparrago.Columns["CLV"].ReadOnly = true;
            dataGridEsparrago.Columns["CLV_PROV"].ReadOnly = true;
            dataGridEsparrago.Columns["NOMBRE DEL ENVASE"].ReadOnly = true;
            dataGridEsparrago.Columns["CANTIDAD"].ReadOnly = true;

            dataGridEsparrago.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            MessageBox.Show("INFORMACION ACTUALIZADA CORRECTAMENTE");
            labelEsparrago.Visible = false;
            progressBarEsparrago.Visible = false;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            bgCorteDiario = new BackgroundWorker();
            bgCorteDiario.WorkerReportsProgress = true;
            bgCorteDiario.ProgressChanged += bg_ProgressChangedcorte;
            bgCorteDiario.DoWork += bg_DoWorkcorte;
            bgCorteDiario.RunWorkerCompleted += bg_RunWorkerCompletedcorte;
            bgCorteDiario.RunWorkerAsync();
            cortemanuallabel.Visible = true;
            cortemanualprogress.Visible = true;
        }

        private void dataGridEsparrago_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (this.dataGridEsparrago.Rows[e.RowIndex].Cells[0].Value.ToString() == "")
                {
                    e.CellStyle.BackColor = Color.LightGreen;
                }

            }
            catch
            {


            }
        }


        private void bg_ProgressChangedcorte(object sender, ProgressChangedEventArgs e)
        {
            // Change the value of the ProgressBar to the BackgroundWorker progress.
            cortemanualprogress.Value = e.ProgressPercentage;
            cortemanualprogress.Step = 1;
            cortemanualprogress.Style = ProgressBarStyle.Continuous;
            cortemanualprogress.Minimum = 0;
            cortemanualprogress.Maximum = 100;

            if (e.ProgressPercentage > 100)
            {
                cortemanuallabel.Text = "100%";
                cortemanualprogress.Value = progress.Maximum;
            }
            else
            {
                cortemanuallabel.Text = Convert.ToString(e.ProgressPercentage) + "%";
                cortemanualprogress.Value = e.ProgressPercentage;
            }
        }

        private void bg_DoWorkcorte(object sender, DoWorkEventArgs e)
        {
            int progreso = 0, porciento = 0, totalEmpleadosProcesar = 0;

            int veces = 0;

            thisConnecion.Open();
            string valida = "SELECT COUNT(env_clave) FROM  tb_cat_envases WHERE (env_inventario = '1')";
            SqlCommand cmd = new SqlCommand(valida, thisConnecion);
            veces = veces + Convert.ToInt32(cmd.ExecuteScalar());

            valida = "SELECT COUNT(cve_prov) FROM Tb_ENV_PROV_CAJ WHERE (estatus = '1')";
            cmd = new SqlCommand(valida, thisConnecion);
            veces = veces + (Convert.ToInt32(cmd.ExecuteScalar()) * 2);

            valida = "SELECT COUNT(cve_prov) FROM Tb_ENV_PROV_CAJ_ESPARRAGO WHERE (estatus = '1')";
            cmd = new SqlCommand(valida, thisConnecion);
            veces = veces + Convert.ToInt32(cmd.ExecuteScalar());


            string fecha = "SELECT SYSDATETIME()";
            cmd = new SqlCommand(fecha, thisConnecion);
            string fecha_hoy = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");
            thisConnecion.Close();

            thisConnecion.Open();
            valida = "SELECT fecha FROM tb_registro_movimientos where fecha = '" + fecha_hoy + "' AND detalle = 'CORTE INVENTARIO ENVASES' AND op_clave = '2.18'";
            cmd = new SqlCommand(valida, thisConnecion);
            string existencia = Convert.ToString(cmd.ExecuteScalar());
            thisConnecion.Close();

            if (existencia.Trim().Length == 0 || existencia.Trim() == null)
            {

                thisConnecion.Open();
                string query = "SELECT env_clave FROM  tb_cat_envases WHERE (env_inventario = '1')";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                /*valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos ORDER BY ENV_FECHA DESC";
                cmd = new SqlCommand(valida, thisConnecion);
                string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");*/

                string fecha_ent = Convert.ToDateTime(fecha_hoy).AddDays(-1).ToString("dd/MM/yyyy");


                while (dr.Read())
                {
                    string clave_envase = Convert.ToString(dr["env_clave"]).Trim();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(ENV_INV_INI_CANT),0) FROM TB_MSTR_INV_ENVASES_dos WHERE ENV_CLAVE = '" + clave_envase + "' AND ENV_FECHA = '" + fecha_ent + "'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND ENT_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND SAL_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    string actualizadato = "UPDATE TB_MSTR_INV_ENVASES_dos SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "'";
                    cmd = new SqlCommand(actualizadato, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();

                    int inv_ini_hoy = (inv_ini + entradas) - salidas;

                    //TB_MSTR_INV_ENVASES_sin_corte

                    //thisConnecion.Open();
                    string cadena = "insert into TB_MSTR_INV_ENVASES_dos (ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                    "Values('" + clave_envase + "','" + fecha_hoy + "',' " + inv_ini_hoy + " ','','')";
                    cmd = new SqlCommand(cadena, thisConnecion);
                    cmd.ExecuteNonQuery();
                    //thisConnecion.Close();

                    cadena = "insert into TB_MSTR_INV_ENVASES_sin_corte (ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                    "Values('" + clave_envase + "','" + fecha_hoy + "',' " + inv_ini_hoy + " ','','')";
                    cmd = new SqlCommand(cadena, thisConnecion);
                    cmd.ExecuteNonQuery();


                    totalEmpleadosProcesar = veces;  //Total de Empleados que seleccionó el usuario
                    progreso++; //Aumentando el progreso 
                    porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                    System.Threading.Thread.Sleep(50);
                    bgCorteDiario.ReportProgress(porciento);
                }


                thisConnecion.Close();

                thisConnecion.Open();
                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values('" + fecha_hoy + "','" + Environment.MachineName.Trim() + "','','CORTEENV','2.18','', 'CORTE INVENTARIO ENVASES', 'SISGAB', '')";
                cmd = new SqlCommand(cadenaregmov, thisConnecion);
                cmd.ExecuteNonQuery();
                thisConnecion.Close();
            }

            string[] valores = { "02", "80" };


            thisConnecion.Open();
            valida = "SELECT fecha FROM tb_registro_movimientos where fecha = '" + fecha_hoy + "' AND detalle = 'CORTE INVENTARIO ENVASES CAJONES PROVEEDOR' AND op_clave = '2.18'";
            cmd = new SqlCommand(valida, thisConnecion);
            existencia = Convert.ToString(cmd.ExecuteScalar());
            thisConnecion.Close();

            if (existencia.Trim().Length == 0 || existencia.Trim() == null)
            {

                thisConnecion.Open();
                string query = "SELECT cve_prov FROM Tb_ENV_PROV_CAJ WHERE (estatus = '1')";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                /*valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos ORDER BY ENV_FECHA DESC";
                cmd = new SqlCommand(valida, thisConnecion);
                string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");*/

                string fecha_ent = Convert.ToDateTime(fecha_hoy).AddDays(-1).ToString("dd/MM/yyyy");


                while (dr.Read())
                {
                    string clave_proveedor = Convert.ToString(dr["cve_prov"]).Trim();

                    foreach (string clave_envase in valores)
                    {
                        //thisConnecion.Open();
                        valida = "SELECT ISNULL(SUM(ENV_INV_INI_CANT),0) FROM TB_MSTR_INV_CAJAS_PROV WHERE ENV_CLAVE = '" + clave_envase + "' AND ENV_FECHA = '" + fecha_ent + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                        cmd = new SqlCommand(valida, thisConnecion);
                        int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                        //thisConnecion.Close();

                        //thisConnecion.Open();
                        valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "' AND ENT_STATUS != 'C'";
                        cmd = new SqlCommand(valida, thisConnecion);
                        int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                        //thisConnecion.Close();

                        //thisConnecion.Open();
                        valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "' AND SAL_STATUS != 'C'";
                        cmd = new SqlCommand(valida, thisConnecion);
                        int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                        //thisConnecion.Close();

                        //thisConnecion.Open();
                        string actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                        cmd = new SqlCommand(actualizadato, thisConnecion);
                        cmd.ExecuteNonQuery();
                        //thisConnecion.Close();

                        int inv_ini_hoy = (inv_ini + entradas) - salidas;

                        //TB_MSTR_INV_ENVASES_sin_corte

                        //thisConnecion.Open();
                        string cadena = "insert into TB_MSTR_INV_CAJAS_PROV (PROV_CLAVE, ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                        "Values('" + clave_proveedor + "', '" + clave_envase + "','" + fecha_hoy + "',' " + inv_ini_hoy + " ','','')";
                        cmd = new SqlCommand(cadena, thisConnecion);
                        cmd.ExecuteNonQuery();
                        //thisConnecion.Close();

                        totalEmpleadosProcesar = veces;  //Total de Empleados que seleccionó el usuario
                        progreso++; //Aumentando el progreso 
                        porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                        System.Threading.Thread.Sleep(50);
                        bgCorteDiario.ReportProgress(porciento);
                    }
                }

                thisConnecion.Close();

                thisConnecion.Open();
                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values('" + fecha_hoy + "','" + Environment.MachineName.Trim() + "','','CORTECAJON','2.18','', 'CORTE INVENTARIO ENVASES CAJONES PROVEEDOR', 'SISGAB', '')";
                cmd = new SqlCommand(cadenaregmov, thisConnecion);
                cmd.ExecuteNonQuery();
                thisConnecion.Close();
            }

            string[] valoresx = { "81" };


            thisConnecion.Open();
            valida = "SELECT fecha FROM tb_registro_movimientos where fecha = '" + fecha_hoy + "' AND detalle = 'CORTE INVENTARIO ENVASES CAJONES ESPARRAGO PROVEEDOR' AND op_clave = '2.18'";
            cmd = new SqlCommand(valida, thisConnecion);
            existencia = Convert.ToString(cmd.ExecuteScalar());
            thisConnecion.Close();

            if (existencia.Trim().Length == 0 || existencia.Trim() == null)
            {

                thisConnecion.Open();
                string query = "SELECT cve_prov FROM Tb_ENV_PROV_CAJ_ESPARRAGO WHERE (estatus = '1')";
                SqlCommand cm = new SqlCommand(query, thisConnecion);
                SqlDataReader dr = cm.ExecuteReader();

                /*valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos ORDER BY ENV_FECHA DESC";
                cmd = new SqlCommand(valida, thisConnecion);
                string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");*/

                string fecha_ent = Convert.ToDateTime(fecha_hoy).AddDays(-1).ToString("dd/MM/yyyy");


                while (dr.Read())
                {
                    string clave_proveedor = Convert.ToString(dr["cve_prov"]).Trim();

                    foreach (string clave_envase in valoresx)
                    {
                        //thisConnecion.Open();
                        valida = "SELECT ISNULL(SUM(ENV_INV_INI_CANT),0) FROM TB_MSTR_INV_CAJAS_PROV_ESPARRAGO WHERE ENV_CLAVE = '" + clave_envase + "' AND ENV_FECHA = '" + fecha_ent + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                        cmd = new SqlCommand(valida, thisConnecion);
                        int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                        //thisConnecion.Close();

                        //thisConnecion.Open();
                        valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "' AND ENT_STATUS != 'C'";
                        cmd = new SqlCommand(valida, thisConnecion);
                        int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                        //thisConnecion.Close();

                        //thisConnecion.Open();
                        valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "' AND SAL_STATUS != 'C'";
                        cmd = new SqlCommand(valida, thisConnecion);
                        int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                        //thisConnecion.Close();

                        //thisConnecion.Open();
                        string actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV_ESPARRAGO SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                        cmd = new SqlCommand(actualizadato, thisConnecion);
                        cmd.ExecuteNonQuery();
                        //thisConnecion.Close();


                        //Linea de codigo Original Para El cierre de Canastillas de Esparrago
                        //int inv_ini_hoy = inv_ini + entradas - salidas;
                        //Se invierte posicion debido a que lo solicito Javier Castrejon, las Salidas Ahora seran positivas
                        int inv_ini_hoy = inv_ini + salidas - entradas;


                        //TB_MSTR_INV_ENVASES_sin_corte

                        //thisConnecion.Open();
                        string cadena = "insert into TB_MSTR_INV_CAJAS_PROV_ESPARRAGO (PROV_CLAVE, ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                        "Values('" + clave_proveedor + "', '" + clave_envase + "','" + fecha_hoy + "',' " + inv_ini_hoy + " ','','')";
                        cmd = new SqlCommand(cadena, thisConnecion);
                        cmd.ExecuteNonQuery();
                        //thisConnecion.Close();
                        totalEmpleadosProcesar = veces;  //Total de Empleados que seleccionó el usuario
                        progreso++; //Aumentando el progreso 
                        porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                        System.Threading.Thread.Sleep(50);
                        bgCorteDiario.ReportProgress(porciento);
                    }
                }

                thisConnecion.Close();

                thisConnecion.Open();
                string cadenaregmov = "insert into tb_registro_movimientos (fecha, nom_compu, nom_usu, tipo_mov, op_clave, folio, detalle, sistema, mov_folio) " +
                                "Values('" + fecha_hoy + "','" + Environment.MachineName.Trim() + "','','CORTEESPAR','2.18','', 'CORTE INVENTARIO ENVASES CAJONES ESPARRAGO PROVEEDOR', 'SISGAB', '')";
                cmd = new SqlCommand(cadenaregmov, thisConnecion);
                cmd.ExecuteNonQuery();
                thisConnecion.Close();
            }

        }

        private void bg_RunWorkerCompletedcorte(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("CORTE DIARIO REALIZADO CORRECTAMENTE");
            cortemanuallabel.Visible = false;
            cortemanualprogress.Visible = false;
        }

        private void BtnActxDia_Click(object sender, EventArgs e)
        {
            string[] valoresx = { "81" };

            string fecha_hoy = FechaInirepo.Value.ToShortDateString();
            SqlCommand cmd;
            string valida = "";

            //thisConnecion.Open();


            thisConnecion.Open();
            string query = "SELECT cve_prov FROM Tb_ENV_PROV_CAJ_ESPARRAGO WHERE (estatus = '1')";
            SqlCommand cm = new SqlCommand(query, thisConnecion);
            SqlDataReader dr = cm.ExecuteReader();

            /*valida = "SELECT TOP(1) ENV_FECHA FROM TB_MSTR_INV_ENVASES_dos ORDER BY ENV_FECHA DESC";
            cmd = new SqlCommand(valida, thisConnecion);
            string fecha_ent = Convert.ToDateTime(cmd.ExecuteScalar()).ToString("dd/MM/yyyy");*/

            string fecha_ent = Convert.ToDateTime(fecha_hoy).AddDays(-1).ToString("dd/MM/yyyy");


            while (dr.Read())
            {
                string clave_proveedor = Convert.ToString(dr["cve_prov"]).Trim();

                foreach (string clave_envase in valoresx)
                {
                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(ENV_INV_INI_CANT),0) FROM TB_MSTR_INV_CAJAS_PROV_ESPARRAGO WHERE ENV_CLAVE = '" + clave_envase + "' AND ENV_FECHA = '" + fecha_ent + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int inv_ini = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "' AND ENT_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int entradas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    valida = "SELECT ISNULL(SUM(cantidad),0) FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "' AND SAL_STATUS != 'C'";
                    cmd = new SqlCommand(valida, thisConnecion);
                    int salidas = Convert.ToInt32(cmd.ExecuteScalar());
                    //thisConnecion.Close();

                    //thisConnecion.Open();
                    string actualizadato = "UPDATE TB_MSTR_INV_CAJAS_PROV_ESPARRAGO SET  ENV_ENTR_CANT = '" + entradas + "', ENV_SAL_CANT = '" + salidas + "' WHERE ENV_FECHA = '" + fecha_ent + "' AND ENV_CLAVE = '" + clave_envase + "' AND PROV_CLAVE = '" + clave_proveedor + "'";
                    cmd = new SqlCommand(actualizadato, thisConnecion);
                    int nRegs = cmd.ExecuteNonQuery();
                    //thisConnecion.Close();


                    //Linea de codigo Original Para El cierre de Canastillas de Esparrago
                    //int inv_ini_hoy = inv_ini + entradas - salidas;
                    //Se invierte posicion debido a que lo solicito Javier Castrejon, las Salidas Ahora seran positivas
                    int inv_ini_hoy = inv_ini + salidas - entradas;


                    //TB_MSTR_INV_ENVASES_sin_corte

                    //thisConnecion.Open();
                    if (nRegs > 0) // no actualizo la informacion
                    {
                        string cadena = "insert into TB_MSTR_INV_CAJAS_PROV_ESPARRAGO (PROV_CLAVE, ENV_CLAVE, ENV_FECHA, ENV_INV_INI_CANT, ENV_ENTR_CANT, ENV_SAL_CANT) " +
                                        "Values('" + clave_proveedor + "', '" + clave_envase + "','" + fecha_hoy + "',' " + inv_ini_hoy + " ','','')";
                        cmd = new SqlCommand(cadena, thisConnecion);
                        cmd.ExecuteNonQuery();
                    }
                    //thisConnecion.Close();
                    //totalEmpleadosProcesar = veces;  //Total de Empleados que seleccionó el usuario
                    //progreso++; //Aumentando el progreso 
                    //porciento = Convert.ToInt16((((double)progreso / (double)totalEmpleadosProcesar) * 100.00)); //Calculo del porcentaje
                    //System.Threading.Thread.Sleep(50);
                    //bgCorteDiario.ReportProgress(porciento);
                }
            }

        }

        private void DtInvFis_ValueChanged(object sender, EventArgs e)
        {
            if (DtInvFis.Text != "")
            {
                string dia_escogido = Convert.ToDateTime(DtInvFis.Text).ToString("dd/MM/yyyy");
                string mes_escogido = Convert.ToDateTime(DtInvFis.Text).ToString("MM");
                //if (mes_actual_Maquina == mes_escogido)
                //{
                thisConnecion.Open();
                string queryX = "SELECT * FROM  TB_MSTR_INV_ENVASES_dos INNER JOIN tb_cat_envases AS B ON TB_MSTR_INV_ENVASES_dos.ENV_CLAVE = B.env_clave WHERE (ENV_FECHA = '" + dia_escogido + "')";
                SqlCommand cmX = new SqlCommand(queryX, thisConnecion);
                SqlDataReader drX = cmX.ExecuteReader();

                inv_Fis_dia.Clear();
                decimal InvF = 0;
                while (drX.Read())
                {
                    DataRow rowix = inv_Fis_dia.NewRow();
                    rowix["CLV"] = drX["ENV_CLAVE"];
                    rowix["NOMBRE DEL ENVASE"] = drX["ENV_NOMBRE"];
                    rowix["CANTIDAD"] = Convert.ToInt32(drX["ENV_INV_INI_CANT"]).ToString("#,##0");
                    rowix["ENTRADAS"] = Convert.ToInt32(drX["ENV_ENTR_CANT"]).ToString("#,##0");
                    rowix["SALIDAS"] = Convert.ToInt32(drX["ENV_SAL_CANT"]).ToString("#,##0");
                    rowix["INVENTARIO TEORICO"] = (Convert.ToInt32(drX["ENV_INV_INI_CANT"]) + Convert.ToInt32(drX["ENV_ENTR_CANT"]) - Convert.ToInt32(drX["ENV_SAL_CANT"])).ToString("#,##0");
                    rowix["INVENTARIO FISICO"] = (Convert.ToString(drX["ENV_INV_FISICO"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_INV_FISICO"]).ToString("#,##0");
                    rowix["OBS"] = drX["ENV_OBS_DIF"].ToString();
                    InvF += Convert.ToDecimal(rowix["INVENTARIO FISICO"]);
                    rowix["ENCAMPO"] = (Convert.ToString(drX["ENV_ENCAMPO"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_ENCAMPO"]).ToString("#,##0");
                    rowix["CONPROD"] = (Convert.ToString(drX["ENV_CONPROD"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_CONPROD"]).ToString("#,##0");
                    rowix["VACIAS"] = (Convert.ToString(drX["ENV_VACIAS"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_VACIAS"]).ToString("#,##0");
                    rowix["CONBASURA"] = (Convert.ToString(drX["ENV_CONBASURA"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_CONBASURA"]).ToString("#,##0");
                    rowix["XREPARAR"] = (Convert.ToString(drX["ENV_XREPARAR"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_XREPARAR"]).ToString("#,##0");
                    rowix["REPARAGUI"] = (Convert.ToString(drX["ENV_REPARAGUI"]).Trim() == "") ? "0" : Convert.ToInt32(drX["ENV_REPARAGUI"]).ToString("#,##0");
                    inv_Fis_dia.Rows.Add(rowix);
                }
                thisConnecion.Close();
                //SendMail(UsuMail.Trim(), UsuMail.Trim().Substring(0, UsuMail.Trim().IndexOf("@")), PwdMail.Trim(), dest3, 3); //;
                BtnSavFis.Enabled = true;
                if (InvF > 0)
                    BtnSavFis.Enabled = false;
            }
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            #region VALIDA Y FORZA ACTUALIZAR
            if (ValidaActualizacion())
            {
                requiereActualizar = true;
                this.Close(); // Cerramos el formulario para dar paso al actualizador
            }
            #endregion
            if (tabControl1.SelectedIndex == 6) // captura de Inventario Fisico 
            {

            }

        }

        void dText_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsDigit(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsControl(e.KeyChar))
            {
                e.Handled = false;
            }
            else if (Char.IsSeparator(e.KeyChar))
            {
                e.Handled = false;
            }
            else
            {
                e.Handled = true;
            }
        }

        private void DGInvFis_EditingControlShowing(object sender, DataGridViewEditingControlShowingEventArgs e)
        {
            DataGridViewTextBoxEditingControl dText = (DataGridViewTextBoxEditingControl)e.Control;
            if (DGInvFis.CurrentCell.ColumnIndex == 6 || DGInvFis.CurrentCell.ColumnIndex > 7) // INVENTARIO FISICO
            {
                dText.KeyPress -= new KeyPressEventHandler(dText_KeyPress);
                dText.KeyPress += new KeyPressEventHandler(dText_KeyPress);
            }
            else
                dText.KeyPress -= new KeyPressEventHandler(dText_KeyPress);
        }

        public void SendMail(string mail, string usuario, string password, string Dest, Int16 Opcion)
        {
            MailMessage msg = new MailMessage();
            MailMessage email = new MailMessage();
            string[] destinatarios = Dest.Split(';');
            string InfoCorreo = "";
            foreach (string destinos in destinatarios)
            {
                email.To.Add(new MailAddress(destinos));
                //email.Bcc.Add("andrea@mrlucky.com.mx");
            }
            email.CC.Add(mail);
            //if (Opcion == 1) // Autorizados
            //{
            //    email.Bcc.Add("musabiaga@mrlucky.com.mx"); //  PARA MANDAR LA COPIA OCULTA 
            //}
            //email.To.Add(new MailAddress(mail));
            //email.To.Add(new MailAddress("gcamacho@mrlucky.com.mx"));
            //string correos = "ricardo.cortes@mrlucky.com.mx; jbravo@mrlucky.com.mx";

            //cap = "";
            //foreach (DataRow fila in captura.Rows)
            //    cap = cap + Convert.ToString(fila["cantidad"].ToString()).Trim() + " - " + Convert.ToString(fila["producto"].ToString()).Trim() + " - " + Convert.ToString(fila["nombre"].ToString()).Trim() + " <br> ";

            // se busca la Informacion del Inventario de Cajones x Proveedor

            string cadena = "SELECT A.prov_clave, A.prov_nombre, D.env_nombre, C.*, (C.ENV_INV_INI_CANT + C.ENV_ENTR_CANT - C.ENV_SAL_CANT ) AS total  " +
                            "FROM tb_cat_proveedor AS A " +
                            "INNER JOIN Tb_ENV_PROV_CAJ AS B ON A.prov_clave = B.cve_prov " +
                            "INNER JOIN TB_MSTR_INV_CAJAS_PROV AS C ON B.cve_prov = C.PROV_CLAVE " +
                            "INNER JOIN tb_cat_envases AS D ON C.ENV_CLAVE = D.env_clave  " +
                            "where ENV_FECHA = '" + DtInvFis.Value.ToShortDateString() + "' " +
                            "ORDER BY A.prov_clave, D.ENV_CLAVE ";
            thisConnecion.Open();
            SqlDataAdapter da = new SqlDataAdapter(cadena, thisConnecion);
            DataSet ds = new DataSet();
            da.Fill(ds, "InvProv");
            DataTable InvProv = ds.Tables["InvProv"];
            thisConnecion.Close();
            int i = 0;
            string cap = "", cap1 = "";
            if (Opcion == 1)
            {
                cap = "<table border=\"1\"> " +
                      "<tr> <TH COLSPAN=14>" + "INVENTARIO FISICO DE CANASTILLAS AL DIA " + DtInvFis.Value.ToString("dd MMM yyyy").ToUpper().Replace(".", "") + "</TH></TR> " +
                      " <tr>   " +
                      " <th scope=\"col\">Clave</strong></th> " +
                      " <th scope=\"col\">Nombre Envse</strong></th> " +
                      " <th scope=\"col\">Inv. Inicial</strong></th> " +
                      " <th scope=\"col\">Entradas</strong></th> " +
                      " <th scope=\"col\">Salidas</strong></th> " +
                      " <th scope=\"col\">Inv. Teorico</strong></th> " +
                      " <th scope=\"col\">Inv. Fisico</strong></th> " +
                      " <th scope=\"col\">Diferencia</strong></th> " +
                      " <th scope=\"col\">Obs</strong></th> " +
                      //" <th scope=\"col\">En Campo</strong></th> " +
                      " <th scope=\"col\">Con Producto</strong></th> " +
                      " <th scope=\"col\">Vacias</strong></th> " +
                      " <th scope=\"col\">Con Basura</strong></th> " +
                      " <th scope=\"col\">X Reparar</strong></th> " +
                      " <th scope=\"col\">Envase Roto</strong></th> " +
                      " </tr> ";
                decimal Correcto = 0;
                foreach (DataGridViewRow Row in DGInvFis.Rows)
                {
                    //cap = cap + Convert.ToString(fila["cantidad"].ToString()).Trim() + " - " + Convert.ToString(fila["producto"].ToString()).Trim() + " - " + Convert.ToString(fila["nombre"].ToString()).Trim() + " <br> ";
                    //if (Convert.ToBoolean(DGDatos.Rows[i].Cells["Autoriza"].Value) == true)
                    //{
                    cap += "<tr> " +
                           "<td ALIGN=LEFT>" + DGInvFis.Rows[i].Cells["CLV"].Value.ToString().Trim() + "</td>" +
                           "<td ALIGN=LEFT>" + DGInvFis.Rows[i].Cells["NOMBRE DEL ENVASE"].Value.ToString().Trim() + "</td>" +
                           "<td ALIGN=CENTER>" + DGInvFis.Rows[i].Cells["CANTIDAD"].Value.ToString().Trim() + "</td>" +
                           "<td ALIGN=CENTER>" + DGInvFis.Rows[i].Cells["ENTRADAS"].Value.ToString().Trim() + "</td>" +
                           "<td ALIGN=CENTER>" + DGInvFis.Rows[i].Cells["SALIDAS"].Value.ToString() + "</td>" +
                           "<td ALIGN=CENTER>" + DGInvFis.Rows[i].Cells["INVENTARIO TEORICO"].Value.ToString() + "</td>" +
                           "<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value).ToString("#,##0") + "</td>" +
                           "<td ALIGN=CENTER>" + (Convert.ToDecimal(DGInvFis.Rows[i].Cells["INVENTARIO TEORICO"].Value) - Convert.ToDecimal(DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value)).ToString("#,##0") + "</td>" +
                           "<td ALIGN=CENTER>" + DGInvFis.Rows[i].Cells["OBS"].Value.ToString() + "</td>" +
                           //"<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["enCampo"].Value).ToString("#,##0")+ "</td>" +
                           "<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["ConProd"].Value).ToString("#,##0") + "</td>" +
                           "<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["Vacias"].Value).ToString("#,##0") + "</td>" +
                           "<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["ConBasura"].Value).ToString("#,##0") + "</td>" +
                           "<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["Xreparar"].Value).ToString("#,##0") + "</td>" +
                           "<td ALIGN=CENTER>" + Convert.ToInt32(DGInvFis.Rows[i].Cells["Reparagui"].Value).ToString("#,##0") + "</td>" +
                           "</tr>";
                    if (Convert.ToDecimal(DGInvFis.Rows[i].Cells["INVENTARIO TEORICO"].Value) == Convert.ToDecimal(DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value))
                        Correcto++;
                    //}
                    i++;
                }
                cap += "<tr> " +
                               "<TH COLSPAN=15>Confiabilidad del " + (Convert.ToDecimal(Correcto / DGInvFis.Rows.Count) * 100).ToString("#,##0") + " %</TH>" +
                               //"<TH>Confiabilidad del " + (DGInvFis.Rows.Count / Correcto).ToString("#,##0") + " %</TH>" +
                               //"<TH COLSPAN=2 ALIGN=RIGTH> </TH>" +
                               "</tr>";
                cap += "</table> ";
                cap += "<br><br><br><br>";
                cap1 = "<table border=\"1\"> " +
                      "<tr> <TH COLSPAN=7>" + "DESGLOSE DE INVENTARIO X PROVEEDOR</TH></TR> " +
                      " <tr>   " +
                      " <th scope=\"col\">Proveedor</strong></th> " +
                      " <th scope=\"col\">Clave Env</strong></th> " +
                      " <th scope=\"col\">Nombre Emvase</strong></th> " +
                      " <th scope=\"col\">Inv. Inicial</strong></th> " +
                      " <th scope=\"col\">Entradas</strong></th> " +
                      " <th scope=\"col\">Salidas</strong></th> " +
                      " <th scope=\"col\">Inv. Teorico</strong></th> " +
                      " </tr> ";
                foreach (DataRow Row in InvProv.Rows)
                {
                    if (Convert.ToDecimal(Row["ENV_INV_INI_CANT"]) != 0 || Convert.ToDecimal(Row["ENV_ENTR_CANT"]) != 0 || Convert.ToDecimal(Row["ENV_SAL_CANT"]) != 0)
                    {
                        //A., D.env_nombre, C.PROV_CLAVE, C., C.ENV_FECHA, C.ENV_INV_INI_CANT, C.ENV_ENTR_CANT, C.ENV_SAL_CANT       
                        cap1 += "<tr> " +
                               "<td ALIGN=LEFT>" + Row["prov_nombre"].ToString().Trim() + "</td>" +
                               "<td ALIGN=LEFT>" + Row["ENV_CLAVE"].ToString().Trim() + "</td>" +
                               "<td ALIGN=CENTER>" + Row["env_nombre"].ToString().Trim() + "</td>" +
                               "<td ALIGN=CENTER>" + Convert.ToInt32(Row["ENV_INV_INI_CANT"].ToString()).ToString("#,##0") + "</td>" +
                               "<td ALIGN=CENTER>" + Convert.ToInt32(Row["ENV_ENTR_CANT"].ToString()).ToString("#,##0") + "</td>" +
                               "<td ALIGN=CENTER>" + Convert.ToInt32(Row["ENV_SAL_CANT"]).ToString("#,##0") + "</td>" +
                               "<td ALIGN=CENTER>" + Convert.ToInt32(Row["TOTAL"]).ToString("#,##0") + "</td>" +
                               "</tr>";
                    }
                }
                cap1 += "</table> ";
                cap = cap + cap1;
            }

            email.From = new MailAddress(mail); //
            email.Subject = "Captura de Inventario Fisico de Canastillas ";//mAsunto; //"Mensaje de Prueba";
            email.Body = "<br><br>" + cap;//mBody;  //"Información de la factura";
            email.IsBodyHtml = true;
            email.Priority = MailPriority.Normal;
            InfoCorreo = "<br><br>" + cap;
            //string archivo = @"C:\Reportes\factura_informativa.txt";
            //if (Archivo.Trim().Length > 0)
            //    if (File.Exists(Archivo))
            //        email.Attachments.Add(new Attachment(Archivo));            
            SmtpClient smtp = new SmtpClient();
            smtp.Host = "mail1.mrlucky.com.mx";
            smtp.Port = 587;
            //smtp.UseDefaultCredentials = false;
            smtp.Credentials = new System.Net.NetworkCredential(usuario, password);
            smtp.EnableSsl = true;
            //smtp.Credentials = new NetworkCredential("ricardo.cortes", "rcedillo");
            //smtp.Credentials = new NetworkCredential(usuario, password);

            try
            {
                smtp.Send(email);
                email.Dispose();
                //MessageBox.Show("correo enviado", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Utilerias.Class1.registrar_movimiento3(DateTime.Now, Environment.MachineName, Utilerias.Class1.Usu_login.Trim(), "A", "2.18", "Inv Fisico de Canastillas", "CAPTURA INV FISICO CANASTILLAS", "SIPGAB", "", "" + InfoCorreo + "");

            }
            catch (Exception ex)
            {
                MessageBox.Show("correo no enviado\r\n" + ex.ToString(), "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnSavFis_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Esta Seguro de Guardar la Informacion? ya no se va a poder MODIFICAR la CAPTURA", "Captura de Inventarios Fisico", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2) == System.Windows.Forms.DialogResult.No)
                return;
            if (!Acumula())
            {
                MessageBox.Show("Revise la Información la Cantidades del detalle estan incorrectas y NO se puede Guardar", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            //SendMail(UsuMail.Trim(), UsuMail.Trim().Substring(0, UsuMail.Trim().IndexOf("@")), PwdMail.Trim(), "ricardo.cortes@mrlucky.com.mx", 1); //;
            thisConnecion.Open();
            string Cadena = "Select Email_dest from tb_mstr_email where cnte_clave = 'CANASTILLA' AND EMAIL_MOV = 'BAS'";
            SqlCommand cmd = new SqlCommand(Cadena, thisConnecion);
            string Dest = Convert.ToString(cmd.ExecuteScalar());
            Int32 i = 0;
            foreach (DataGridViewRow Row in DGInvFis.Rows)
            {
                string Cant = DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value.ToString();
                string Fec = DtInvFis.Value.ToShortDateString();
                string Env = DGInvFis.Rows[i].Cells["CLV"].Value.ToString();
                string Obs = DGInvFis.Rows[i].Cells["OBS"].Value.ToString();
                string A = DGInvFis.Rows[i].Cells["ConProd"].Value.ToString();
                string B = DGInvFis.Rows[i].Cells["Vacias"].Value.ToString();
                string C = DGInvFis.Rows[i].Cells["ConBasura"].Value.ToString();
                string D = DGInvFis.Rows[i].Cells["Xreparar"].Value.ToString();
                string E = DGInvFis.Rows[i].Cells["Reparagui"].Value.ToString();
                string F = DGInvFis.Rows[i].Cells["enCampo"].Value.ToString();
                if (Obs.Trim().Length > 40)
                    Obs = Obs.Substring(0, 40);
                Cadena = "Update TB_MSTR_INV_ENVASES_dos Set ENV_INV_FISICO = '" + Cant + "', ENV_OBS_DIF = '" + Obs + "', " +
                         "Env_EnCampo = '" + F + "', Env_ConProd = '" + A + "', Env_Vacias = '" + B + "', Env_ConBasura = '" + C + "', " +
                         "Env_XReparar = '" + D + "', Env_ReparAgui = '" + E + "' " +
                         "Where ENV_CLAVE = '" + Env + "' AND ENV_FECHA = '" + Fec + "' ";
                cmd = new SqlCommand(Cadena, thisConnecion);
                cmd.ExecuteNonQuery();
                i++;
            }
            thisConnecion.Close();
            MessageBox.Show("Inventario Fisico Actualizado Correctamente!!!", "Inventario Fisico", MessageBoxButtons.OK, MessageBoxIcon.Information);
            BtnSavFis.Enabled = false;
            SendMail(UsuMail.Trim(), UsuMail.Trim().Substring(0, UsuMail.Trim().IndexOf("@")), PwdMail.Trim(), Dest, 1); //;
        }

        private void DGInvFis_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //if (e.ColumnIndex == 7) // Obs
            //    if (e.FormattedValue.ToString().Length > 40)
            //    {
            //        MessageBox.Show("la longitud no puede ser mayor a 40 caracteres");
            //        e.Cancel = false;
            //    }
            //if (e.ColumnIndex == 7)
            //{
            //    string errorMsg;
            //    if (e.FormattedValue.ToString().Trim().Length > 0)
            //        if (ValidaHoraEnt(e.FormattedValue.ToString(), out errorMsg))
            //        {
            //            // Cancel the event and select the text to be corrected by the user.
            //            e.Cancel = true;
            //            MessageBox.Show(errorMsg, "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            //        }
            //        else
            //            e.Cancel = false;
            //}
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {

            #region DESCARGA LA NUEVA VERSION DEL EJECUTABLE
            if (requiereActualizar)
            {
                string updaterPath = @"c:\sisgabweb\DownFile.exe";

                if (File.Exists(updaterPath))
                {
                    try
                    {
                        // Iniciamos el actualizador de manera limpia
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = updaterPath,
                            Arguments = "SistemaEnvases.exe",
                            UseShellExecute = true // Asegura que corra correctamente en el entorno de Windows
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"No se pudo iniciar el actualizador: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("No se encontró el archivo actualizador (DownFile.exe).", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            #endregion

        }

        #region METODO PARA VALIDAR LA ACTUALIZACION DEL EJECUTABLE
        public bool ValidaActualizacion()
        {
            string rutaServerTxt = @"\\gabira1\sisgabweb\Valida.txt";
            string rutaServerExe = @"\\gabira1\sisgabweb\SistemaEnvases.exe";
            string rutaLocalExe = @"c:\sisgabweb\SistemaEnvases.exe";

            try
            {
                // 1. Validamos que el archivo de control en el servidor exista
                if (File.Exists(rutaServerTxt) && File.Exists(rutaServerExe) && File.Exists(rutaLocalExe))
                {
                    DateTime fechaLocal = File.GetLastWriteTime(rutaLocalExe);
                    DateTime fechaServer = File.GetLastWriteTime(rutaServerExe);

                    // 2. Comparamos fechas
                    if (fechaServer > fechaLocal)
                    {
                        MessageBox.Show("Hay una versión más reciente. El sistema se cerrará para actualizarse.\n\nPor favor, vuelva a abrir el programa cuando finalice.",
                                        "Actualización Disponible",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information);
                        return true;
                    }
                }
            }
            catch (IOException ex)
            {
                // Si la red se cae o el archivo está bloqueado, atrapamos el error para que no truene la app
                Console.WriteLine($"Error al validar actualización: {ex.Message}");
            }

            return false;
        }
        #endregion



        private Boolean Acumula()
        {
            Boolean Correcto = true;
            int i = 0;
            decimal SumaInvFis = 0;
            foreach (DataGridViewRow row in DGInvFis.Rows)
            {
                decimal A = Convert.ToDecimal(DGInvFis.Rows[i].Cells["ConProd"].Value);
                decimal B = Convert.ToDecimal(DGInvFis.Rows[i].Cells["Vacias"].Value);
                decimal C = Convert.ToDecimal(DGInvFis.Rows[i].Cells["ConBasura"].Value);
                decimal D = Convert.ToDecimal(DGInvFis.Rows[i].Cells["Xreparar"].Value);
                decimal E = Convert.ToDecimal(DGInvFis.Rows[i].Cells["Reparagui"].Value);
                //decimal F = Convert.ToDecimal(DGInvFis.Rows[i].Cells["enCampo"].Value);
                decimal G = Convert.ToDecimal(DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value);
                decimal Tot = A + B + C + D + E; // +F;
                SumaInvFis += G;
                if (G != Tot) // || F < 0)
                {
                    MessageBox.Show("La Canastilla " + DGInvFis.Rows[i].Cells["NOMBRE DEL ENVASE"].Value + "No Coincide Fisico " + G.ToString("#,###") + " VS Detalle " + Tot.ToString("#,###"), "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                    Correcto = false;
                }
                i++;
            }
            if (SumaInvFis == 0)
                Correcto = false;
            return Correcto;
        }

        private void DGInvFis_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            Suma();
        }

        private void Suma()
        {
            int i = 0;
            decimal SumaInvFis = 0;
            foreach (DataGridViewRow row in DGInvFis.Rows)
            {
                decimal A = Convert.ToDecimal(DGInvFis.Rows[i].Cells["ConProd"].Value);
                decimal B = Convert.ToDecimal(DGInvFis.Rows[i].Cells["Vacias"].Value);
                decimal C = Convert.ToDecimal(DGInvFis.Rows[i].Cells["ConBasura"].Value);
                decimal D = Convert.ToDecimal(DGInvFis.Rows[i].Cells["Xreparar"].Value);
                decimal E = Convert.ToDecimal(DGInvFis.Rows[i].Cells["Reparagui"].Value);
                //decimal F = Convert.ToDecimal(DGInvFis.Rows[i].Cells["enCampo"].Value);
                decimal G = Convert.ToDecimal(DGInvFis.Rows[i].Cells["INVENTARIO FISICO"].Value);
                //decimal F = G - (B + C + D + E + A);
                //DGInvFis.Rows[i].Cells["enCampo"].Value = F;
                decimal Tot = A + B + C + D + E; // +F;
                SumaInvFis += G;
                i++;
            }

        }
    }
}
