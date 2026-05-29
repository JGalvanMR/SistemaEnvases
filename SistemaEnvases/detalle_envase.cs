using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SistemaEnvases
{
    public partial class detalle_envase : Form
    {
        public string tippo = "";
        public detalle_envase(DataTable reporte, string fecha, string tipo, string NOMBRE_ENVASE)
        {
            InitializeComponent();

            tippo = tipo;

            label2.Text =  NOMBRE_ENVASE;

            if (tipo == "ENTRADA")
            {
                label1.Text = "Detalle de Entradas al " + fecha;    

            }
            else {
                label1.Text = "Detalle de Salidas al " + fecha;
            }

            GRIDREPORTEENVASEDET.AutoGenerateColumns = true;
            //GRIDDETALLEENVASE.ReadOnly = true;
            //GRIDDETALLEENVASE.AllowUserToAddRows = false;
            GRIDREPORTEENVASEDET.DataSource = reporte;
            GRIDREPORTEENVASEDET.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            funcion();


            
        }

        public void funcion() {
            foreach (DataGridViewRow row in GRIDREPORTEENVASEDET.Rows)
            {
                int rowi = row.Index;
                
            }
        }

        private void GRIDDETALLEENVASE_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (tippo == "ENTRADA")
            {
                if (this.GRIDREPORTEENVASEDET.Rows[e.RowIndex].Cells[14].Value.ToString() == "C")
                {
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (this.GRIDREPORTEENVASEDET.Rows[e.RowIndex].Cells[9].Value.ToString() == "TOTAL")
                {
                    e.CellStyle.BackColor = Color.LightGreen;
                }
            }
            else
            {
                if (this.GRIDREPORTEENVASEDET.Rows[e.RowIndex].Cells[13].Value.ToString() == "C")
                {
                    e.CellStyle.BackColor = Color.Yellow;
                }
                else if (this.GRIDREPORTEENVASEDET.Rows[e.RowIndex].Cells[9].Value.ToString() == "TOTAL")
                {
                    e.CellStyle.BackColor = Color.LightGreen;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
