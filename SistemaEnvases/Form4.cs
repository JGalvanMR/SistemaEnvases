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
    public partial class Form4 : Form
    {
        SqlConnection thisConnecion = new SqlConnection(Utilerias.Class1.ConnectionString);
        //SqlConnection thisConnecion = new SqlConnection("Persist Security Info=False;user id=sa; password=Gabira1;Initial Catalog =GAB_Irapuato_Prueba; server=tcp:192.168.123.6,1433; Connect Timeout = 130; MultipleActiveResultSets=True");
        public DataTable inv_x_proveedor_reporte = new DataTable();
        public int reporteTipo = 0;
        public string tippo = "";

        public Form4(DataTable reporte_ESPARRAGO, string fecha, string tipo, string NOMBRE_ENVASE, string NOMBRE_PROVEEDOR)
        {
            InitializeComponent();

            tippo = tipo;

            detenvafech.Text = NOMBRE_PROVEEDOR + "/" + NOMBRE_ENVASE;

            if (tipo == "ENTRADA")
            {
                detalleprovcaj.Text = "Detalle de Entradas Esparrago al " + fecha;

            }
            else
            {
                detalleprovcaj.Text = "Detalle de Salidas Esparrago al " + fecha;
            }

            reportegriddetcajESPARR.AutoGenerateColumns = true;
            reportegriddetcajESPARR.ReadOnly = true;
            reportegriddetcajESPARR.AllowUserToAddRows = false;
            reportegriddetcajESPARR.DataSource = reporte_ESPARRAGO;
            reportegriddetcajESPARR.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }


        private void reportegriddetcajESPARR_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (tippo == "ENTRADA")
            {
                // En entradas, ESTATUS está en Cells[14] (Cells[13] = RECIBO_MP)
                if (this.reportegriddetcajESPARR.Rows[e.RowIndex].Cells[14].Value.ToString() == "C")
                {
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (this.reportegriddetcajESPARR.Rows[e.RowIndex].Cells[9].Value.ToString() == "TOTAL")
                {
                    e.CellStyle.BackColor = Color.LightGreen;
                }
            }
            else
            {
                // En salidas, ESTATUS está en Cells[13] (no hay RECIBO_MP)
                if (this.reportegriddetcajESPARR.Rows[e.RowIndex].Cells[13].Value.ToString() == "C")
                {
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (this.reportegriddetcajESPARR.Rows[e.RowIndex].Cells[9].Value.ToString() == "TOTAL")
                {
                    e.CellStyle.BackColor = Color.LightGreen;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Form4_Load(object sender, EventArgs e)
        {

        }
    }
}
