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
    
    public partial class DetalleEnvase : Form
    {
        SqlConnection thisConnecion = new SqlConnection(Utilerias.Class1.ConnectionString);
        //SqlConnection thisConnecion = new SqlConnection("Persist Security Info=False;user id=sa; password=Gabira1;Initial Catalog =GAB_Irapuato_Prueba; server=tcp:192.168.123.6,1433; Connect Timeout = 130; MultipleActiveResultSets=True");
        public DataTable inv_x_proveedor_reporte = new DataTable();
        public int reporteTipo = 0;
        public DetalleEnvase(DataTable reporte, int quereporte)
        {
            InitializeComponent();
            inv_x_proveedor_reporte.Clear();
            inv_x_proveedor_reporte = reporte;
            reportegrid.AutoGenerateColumns = true;
            reportegrid.ReadOnly = true;
            reportegrid.AllowUserToAddRows = false;
            reportegrid.DataSource = inv_x_proveedor_reporte;
            reportegrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            reporteTipo = quereporte;

            if (reporteTipo == 0)
            {
                detalle.Visible = true;
                reportegrid.Columns["ENTRADAS"].Visible = false;
                reportegrid.Columns["SALIDAS"].Visible = false;
            }
        }

        private void reportegrid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            try
            {
                if (this.reportegrid.Rows[e.RowIndex].Cells[1].Value.ToString() == "TOTALES")
                {
                    e.CellStyle.BackColor = Color.Aqua;
                }

            }
            catch
            {


            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application aplicacion;
            Microsoft.Office.Interop.Excel.Workbook libro;
            Microsoft.Office.Interop.Excel.Worksheet hoja;
            aplicacion = new Microsoft.Office.Interop.Excel.Application();
            libro = aplicacion.Workbooks.Add();
            //libro = aplicacion.Workbooks.Open(@"C:\\Reportes\Reporte_liquidaciones_esparrago.xls");
            hoja = (Microsoft.Office.Interop.Excel.Worksheet)libro.Worksheets.get_Item(1);

            Microsoft.Office.Interop.Excel.Range r;
            hoja.Cells[2, 2] = "Comercializador GAB, S.A. de C.V.";
            r = hoja.Range[hoja.Cells[2, 2], hoja.Cells[2, 2]];
            r.Font.Bold = true;
            r.Font.Size = 16;

            hoja.Cells[3, 2] = "REPORTE DE INVENTARIO DE EnVASES DIARIO";
            r = hoja.Range[hoja.Cells[3, 2], hoja.Cells[3, 2]];
            r.Font.Bold = true;


            string ruta = "c:\\SisGabWeb\\logo.png";
            //hoja.Shapes.AddPicture(ruta, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoCTrue, 190, 0, 70, 70);
            hoja.Range[hoja.Cells[1, 4], hoja.Cells[4, 4]].Merge();
            r = hoja.get_Range("D1", "D4");


            if (reporteTipo == 1)
            {
                hoja.Cells[6, 1] = "Clave Envase";
                hoja.Cells[6, 2] = "Nombre Envase";
                hoja.Cells[6, 3] = "Clave Proveedor";
                hoja.Cells[6, 4] = "Nombre Proveedor";
                hoja.Cells[6, 5] = "Entradas";
                hoja.Cells[6, 6] = "Salidas";
                hoja.Cells[6, 7] = "Inventario";    
            }
            else {
                if (detalle.Checked == true)
                {

                    hoja.Cells[6, 1] = "Clave Envase";
                    hoja.Cells[6, 2] = "Nombre Envase";
                    hoja.Cells[6, 3] = "Fecha Corte";
                    hoja.Cells[6, 4] = "Inventario Inicial";
                    hoja.Cells[6, 5] = "Entradas";
                    hoja.Cells[6, 6] = "Salidas";
                    hoja.Cells[6, 7] = "Inventario Final";

                }
                else
                {
                    hoja.Cells[6, 1] = "Clave Envase";
                    hoja.Cells[6, 2] = "Nombre Envase";
                    hoja.Cells[6, 3] = "Fecha Corte";
                    hoja.Cells[6, 4] = "Inventario Inicial";
                    hoja.Cells[6, 5] = "Inventario Final";
                }
            }

            
            r = hoja.Range[hoja.Cells[6, 1], hoja.Cells[6, 14]];
            r.Font.Bold = true;

            Cursor.Current = Cursors.WaitCursor;

            r = hoja.Range[hoja.Cells[8, 4], hoja.Cells[reportegrid.Rows.Count + 10, 14]];
            r.NumberFormat = "#,##0";
            r.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignRight;

            int filaactual = 8;
            foreach (DataRow row in inv_x_proveedor_reporte.Rows)
            {   //Here 2 cell is target value and 1 cell is Volume 
                if (reporteTipo == 1)
                {
                    hoja.Cells[filaactual, 1] = row[0];
                    hoja.Cells[filaactual, 2] = row[1];
                    hoja.Cells[filaactual, 3] = row[2];
                    hoja.Cells[filaactual, 4] = row[3];
                    hoja.Cells[filaactual, 5] = row[4];
                    hoja.Cells[filaactual, 6] = row[5];
                    hoja.Cells[filaactual, 7] = row[6];
                }
                else
                {
                    if (detalle.Checked == true)
                    {

                        hoja.Cells[filaactual, 1] = row[0];
                        hoja.Cells[filaactual, 2] = row[1];
                        hoja.Cells[filaactual, 3] = row[2];
                        hoja.Cells[filaactual, 4] = row[3];
                        hoja.Cells[filaactual, 5] = row[4];
                        hoja.Cells[filaactual, 6] = row[5];
                        hoja.Cells[filaactual, 7] = row[6];
                    }
                    else
                    {
                        hoja.Cells[filaactual, 1] = row[0];
                        hoja.Cells[filaactual, 2] = row[1];
                        hoja.Cells[filaactual, 3] = row[2];
                        hoja.Cells[filaactual, 4] = row[3];
                        hoja.Cells[filaactual, 5] = row[6];
                        
                    }
                }
                filaactual++;
            } 


            int rowdatagrid = 8;
            foreach (DataGridViewRow row in reportegrid.Rows)
            {   
                
            } 

            aplicacion.Columns.AutoFit();
            aplicacion.Rows.AutoFit();
            aplicacion.Visible = true;
 
        }

        public void FormattingExcelCells(Microsoft.Office.Interop.Excel.Range range, string HTMLcolorCode, System.Drawing.Color fontColor, bool IsFontbool)
        {
            range.Interior.Color = System.Drawing.ColorTranslator.FromHtml(HTMLcolorCode);
            range.Font.Color = System.Drawing.ColorTranslator.ToOle(fontColor);
            if (IsFontbool == true)
            {
                range.Font.Bold = IsFontbool;
            }
        }

        private void detalle_CheckedChanged(object sender, EventArgs e)
        {
            if (detalle.Checked == true)
            {

                reportegrid.Columns["ENTRADAS"].Visible = true;
                reportegrid.Columns["SALIDAS"].Visible = true;

            }
            else {
                reportegrid.Columns["ENTRADAS"].Visible = false;
                reportegrid.Columns["SALIDAS"].Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            FormCollection formulariosApp = Application.OpenForms;
            foreach (Form f in formulariosApp)
            {
                if (f.Name == "detalle_envase")
                {
                    f.Close();
                    break;
                }
            }
            this.Close();
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            
        }

        private void dateTimeinventarioinicial_ValueChanged(object sender, EventArgs e)
        {

            int Total_INVENTARIO_Ini = 0;
            int Total_Entradas = 0;
            int Total_Salidas = 0;
            int Total_Inventario_Final = 0;


            string fecha_ent = Convert.ToDateTime(Datetimepickerrepor.Text).ToString("dd/MM/yyyy");

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
            DataTable inv_x_proveedorini = new DataTable();
            inv_x_proveedorini.Load(Info);
            DataColumn column;

            thisConnecion.Close();

            int T = 0;
            if (inv_x_proveedorini.Rows.Count == 0)
            {
                MessageBox.Show("NO ENCONTRE INFORMACION PARA ESTE PROVEEDOR", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }

            inv_x_proveedor_reporte.Clear();

            foreach (DataRow row in inv_x_proveedorini.Rows)
            {
                int total = 0;


                thisConnecion.Open();
                string valida = "SELECT sum(ENV_INV_INI_CANT) FROM TB_MSTR_INV_ENVASES_dos WHERE ENV_CLAVE = '" + row["ENV_CLAVE"] + "' AND ENV_FECHA = '" + fecha_ent + "'";
                SqlCommand cmd = new SqlCommand(valida, thisConnecion);
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




                DataRow rowi = inv_x_proveedor_reporte.NewRow();
                rowi["ID"] = row["ENV_CLAVE"];
                rowi["NOMBRE"] = row["env_nombre"];
                rowi["FECHA"] = Convert.ToDateTime(row["ENV_FECHA"]).ToString("dd/MM/yyyy");
                rowi["INVENTARIO INICIAL"] = row["ENV_INV_INI_CANT"];
                rowi["ENTRADAS"] = Convert.ToInt32(entradas);
                rowi["SALIDAS"] = Convert.ToInt32(salidas);
                rowi["INVENTARIO FINAL"] = Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(entradas) - Convert.ToInt32(salidas);
                inv_x_proveedor_reporte.Rows.Add(rowi);


                Total_INVENTARIO_Ini = Total_INVENTARIO_Ini + Convert.ToInt16(row["ENV_INV_INI_CANT"]);
                Total_Entradas = Total_Entradas + Convert.ToInt32(entradas);
                Total_Salidas = Total_Salidas + Convert.ToInt32(salidas);
                Total_Inventario_Final = Total_Inventario_Final + Convert.ToInt32(row["ENV_INV_INI_CANT"]) + Convert.ToInt32(salidas) - Convert.ToInt32(entradas);
				

            }

            DataRow rowixw = inv_x_proveedor_reporte.NewRow();
            rowixw["ID"] = "";
            rowixw["NOMBRE"] = "TOTALES";
            rowixw["FECHA"] = "";
            rowixw["INVENTARIO INICIAL"] = Total_INVENTARIO_Ini;
            rowixw["ENTRADAS"] = Total_Entradas;
            rowixw["SALIDAS"] = Total_Salidas;
            rowixw["INVENTARIO FINAL"] = Total_Inventario_Final;
            inv_x_proveedor_reporte.Rows.Add(rowixw);


            reportegrid.AutoGenerateColumns = true;
            reportegrid.DataSource = inv_x_proveedor_reporte;

        }

        private void reportegrid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            string clave_envase = reportegrid.Rows[e.RowIndex].Cells[0].Value.ToString();
            string NOMBRE_ENVASE = reportegrid.Rows[e.RowIndex].Cells[1].Value.ToString();
            //nombre columna
            string nombrecolumna = reportegrid.Columns[e.ColumnIndex].HeaderText;

            DataTable Reporte_Entradas = new DataTable();
            DataTable Reporte_Salidas = new DataTable();

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
            //Reporte_Kardex.Clear();


            if (nombrecolumna == "ENTRADAS")
            {
                
                thisConnecion.Open();
                string query = "SELECT * FROM TB_ENTRADAS_ENVASES INNER JOIN TB_DETENTRADAS_ENVASES ON TB_ENTRADAS_ENVASES.FOLIO = TB_DETENTRADAS_ENVASES.FOLIO WHERE TB_ENTRADAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(Datetimepickerrepor.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(Datetimepickerrepor.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE = '" + clave_envase + "' ORDER BY FECHA, ENV_CLAVE";
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


                detalle_envase ob = new detalle_envase(Reporte_Entradas, Datetimepickerrepor.Text, "ENTRADA", NOMBRE_ENVASE);

                ob.Show();
            }
            else if (nombrecolumna == "SALIDAS")
            {
                
                thisConnecion.Open();
                string query = "SELECT * FROM TB_SALIDAS_ENVASES INNER JOIN TB_DETSALIDAS_ENVASES ON TB_SALIDAS_ENVASES.FOLIO = TB_DETSALIDAS_ENVASES.FOLIO WHERE TB_SALIDAS_ENVASES.FECHA BETWEEN '" + Convert.ToDateTime(Datetimepickerrepor.Text).ToString("dd/MM/yyyy") + "' AND '" + Convert.ToDateTime(Datetimepickerrepor.Text).ToString("dd/MM/yyyy") + "' AND ENV_CLAVE = '" + clave_envase + "' ORDER BY FECHA, ENV_CLAVE";
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


                detalle_envase ob = new detalle_envase(Reporte_Salidas, Datetimepickerrepor.Text, "SALIDA", NOMBRE_ENVASE);

                ob.Show();


            }

            //obtienes el valor de la primer columna
            
        }

        private void reportegrid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }   
    }
}
