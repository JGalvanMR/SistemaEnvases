using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace SistemaEnvases
{
    /// <summary>
    /// Pestaña "SALIDAS RFID": consulta las salidas de envases generadas desde la app
    /// RFIDTrackBin (se identifican por la marca "(RFID #n)" en tb_registro_movimientos,
    /// op_clave 2.18) y permite imprimirlas / reimprimirlas con el mismo formato de ticket.
    /// Se agrega por código desde Form1_Load; no modifica el diseñador.
    /// </summary>
    public class SalidasRfidTab
    {
        private readonly DateTimePicker dtDesde = new DateTimePicker();
        private readonly DateTimePicker dtHasta = new DateTimePicker();
        private readonly TextBox txtFolio = new TextBox();
        private readonly Button btnBuscar = new Button();
        private readonly Button btnImprimir = new Button();
        private readonly Label lblInfo = new Label();
        private readonly DataGridView grid = new DataGridView();

        public static SalidasRfidTab Agregar(TabControl tabs)
        {
            SalidasRfidTab t = new SalidasRfidTab();
            t.Construir(tabs);
            return t;
        }

        private void Construir(TabControl tabs)
        {
            TabPage page = new TabPage("SALIDAS RFID");
            page.UseVisualStyleBackColor = true;

            Panel top = new Panel();
            top.Dock = DockStyle.Top;
            top.Height = 64;

            Label l1 = new Label(); l1.Text = "Desde:"; l1.AutoSize = true; l1.Location = new Point(12, 12);
            dtDesde.Format = DateTimePickerFormat.Short; dtDesde.Location = new Point(12, 30); dtDesde.Width = 110;
            Label l2 = new Label(); l2.Text = "Hasta:"; l2.AutoSize = true; l2.Location = new Point(132, 12);
            dtHasta.Format = DateTimePickerFormat.Short; dtHasta.Location = new Point(132, 30); dtHasta.Width = 110;
            Label l3 = new Label(); l3.Text = "Folio (opcional):"; l3.AutoSize = true; l3.Location = new Point(252, 12);
            txtFolio.Location = new Point(252, 30); txtFolio.Width = 110;
            txtFolio.KeyDown += delegate (object s, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; Buscar(); }
            };

            btnBuscar.Text = "Buscar"; btnBuscar.Location = new Point(378, 27); btnBuscar.Size = new Size(90, 26);
            btnBuscar.Click += delegate { Buscar(); };
            btnImprimir.Text = "Imprimir"; btnImprimir.Location = new Point(478, 27); btnImprimir.Size = new Size(110, 26);
            btnImprimir.Enabled = false;
            btnImprimir.Click += delegate { ImprimirSeleccionada(); };

            lblInfo.AutoSize = true; lblInfo.Location = new Point(600, 33);

            top.Controls.AddRange(new Control[] { l1, dtDesde, l2, dtHasta, l3, txtFolio, btnBuscar, btnImprimir, lblInfo });

            grid.Dock = DockStyle.Fill;
            grid.ReadOnly = true;
            grid.AllowUserToAddRows = false;
            grid.AllowUserToDeleteRows = false;
            grid.MultiSelect = false;
            grid.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            grid.RowHeadersVisible = false;
            grid.SelectionChanged += delegate { btnImprimir.Enabled = grid.CurrentRow != null; };
            grid.CellDoubleClick += delegate (object s, DataGridViewCellEventArgs e)
            {
                if (e.RowIndex >= 0) ImprimirSeleccionada();
            };
            grid.DataBindingComplete += delegate
            {
                foreach (DataGridViewRow r in grid.Rows)
                {
                    object est = r.Cells["ESTATUS"].Value;
                    if (est != null && est.ToString().Trim() == "C")
                        r.DefaultCellStyle.BackColor = Color.Yellow; // cancelado
                }
            };

            page.Controls.Add(grid);
            page.Controls.Add(top);
            tabs.TabPages.Add(page);

            // Cargar al abrir la pestaña por primera vez
            bool cargada = false;
            tabs.SelectedIndexChanged += delegate
            {
                if (!cargada && tabs.SelectedTab == page) { cargada = true; Buscar(); }
            };
        }

        private void Buscar()
        {
            int folio = 0;
            bool porFolio = txtFolio.Text.Trim().Length > 0;
            if (porFolio && !int.TryParse(txtFolio.Text.Trim(), out folio))
            {
                MessageBox.Show("El folio debe ser numérico.");
                return;
            }

            string sql =
                "SELECT s.FOLIO, s.FECHA, RTRIM(s.PROV_NOMBRE) AS PROVEEDOR, RTRIM(s.RCH_NOMBRE) AS RANCHO, " +
                "RTRIM(s.TBL_NOMBRE) AS TABLA, RTRIM(s.NOM_CHOFER) AS CHOFER, RTRIM(s.NOM_OPERADOR) AS OPERADOR, " +
                "ISNULL((SELECT SUM(d.CANTIDAD) FROM TB_DETSALIDAS_ENVASES d WHERE d.FOLIO = s.FOLIO), 0) AS CANTIDAD, " +
                "s.SAL_STATUS AS ESTATUS " +
                "FROM TB_SALIDAS_ENVASES s " +
                "WHERE " + (porFolio ? "s.FOLIO = @folio " : "s.FECHA >= @desde AND s.FECHA < DATEADD(day, 1, @hasta) ") +
                "AND EXISTS (SELECT 1 FROM tb_registro_movimientos m WHERE m.op_clave = '2.18' " +
                "AND m.detalle LIKE '%(RFID #%' AND LTRIM(RTRIM(m.folio)) = CAST(s.FOLIO AS varchar(20))) " +
                "ORDER BY s.FOLIO DESC";

            try
            {
                using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
                using (SqlCommand cmd = new SqlCommand(sql, conn))
                {
                    if (porFolio) cmd.Parameters.AddWithValue("@folio", folio);
                    else
                    {
                        cmd.Parameters.AddWithValue("@desde", dtDesde.Value.Date);
                        cmd.Parameters.AddWithValue("@hasta", dtHasta.Value.Date);
                    }
                    DataTable dt = new DataTable();
                    new SqlDataAdapter(cmd).Fill(dt);
                    grid.DataSource = dt;
                    lblInfo.Text = dt.Rows.Count + " salida(s) RFID";
                    if (porFolio && dt.Rows.Count == 0)
                        MessageBox.Show("El folio " + folio + " no corresponde a una salida generada desde la app RFID.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al consultar salidas RFID: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImprimirSeleccionada()
        {
            if (grid.CurrentRow == null) return;
            int folio = Convert.ToInt32(grid.CurrentRow.Cells["FOLIO"].Value);
            string estatus = Convert.ToString(grid.CurrentRow.Cells["ESTATUS"].Value).Trim();
            if (estatus == "C")
            {
                MessageBox.Show("El folio " + folio + " está cancelado y no se puede imprimir.");
                return;
            }

            PrintDialog dlg = new PrintDialog();
            dlg.Document = new PrintDocument();
            if (dlg.ShowDialog() != DialogResult.OK) return;

            try
            {
                Imprimir(folio, dlg.PrinterSettings);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al imprimir: " + ex.Message, "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>Mismo formato de ticket que Form1.Imprimir_Salida_Con_Folio, leyendo de la BD.</summary>
        private void Imprimir(int folio, PrinterSettings impresora)
        {
            string prov, rancho, tabla, chofer, operador, fechaGuardado;
            string envasesprod = "CAN    ENV    DESCRIPCION\r\n";
            float producto = 152.0F;

            using (SqlConnection conn = new SqlConnection(Utilerias.Class1.ConnectionString))
            {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(
                    "SELECT RTRIM(PROV_NOMBRE), RTRIM(RCH_NOMBRE), RTRIM(TBL_NOMBRE), RTRIM(NOM_CHOFER), " +
                    "RTRIM(NOM_OPERADOR), RTRIM(FECHA_GUARDADO) FROM TB_SALIDAS_ENVASES WHERE FOLIO = @f", conn))
                {
                    cmd.Parameters.AddWithValue("@f", folio);
                    using (SqlDataReader rd = cmd.ExecuteReader())
                    {
                        if (!rd.Read()) throw new Exception("No se encontró el folio " + folio);
                        prov = Convert.ToString(rd[0]); rancho = Convert.ToString(rd[1]); tabla = Convert.ToString(rd[2]);
                        chofer = Convert.ToString(rd[3]); operador = Convert.ToString(rd[4]); fechaGuardado = Convert.ToString(rd[5]);
                    }
                }
                using (SqlCommand cmd = new SqlCommand(
                    "SELECT CANTIDAD, RTRIM(ENV_CLAVE), RTRIM(ENV_NOMBRE), RTRIM(PROD_CLAVE), RTRIM(PROD_NOMBRE) " +
                    "FROM TB_DETSALIDAS_ENVASES WHERE FOLIO = @f", conn))
                {
                    cmd.Parameters.AddWithValue("@f", folio);
                    using (SqlDataReader rd = cmd.ExecuteReader())
                    {
                        while (rd.Read())
                        {
                            envasesprod += rd[0] + "  |   " + rd[1] + "  |  " + rd[2] + "\r\n    ";
                            envasesprod += rd[3] + "  |  " + rd[4] + "\r\n";
                            producto += 25;
                        }
                    }
                }
            }

            string drawString1 = "\"Comercializadora GAB\"";
            string drawString2 = "Salida de Envases (REIMPRESION)\r\n" + fechaGuardado;
            string drawString3 = "FOLIO: " + folio;
            string drawString4 = "PROV: " + prov;
            string drawString5 = "RANCHO: " + rancho;
            string drawString6 = "TABLA: " + tabla;
            string drawstring8 = "Recibio Chofer: " + chofer;
            string drawstring9 = "Entrego Operador: " + operador;
            string drawstring10 = "* " + folio + " *";
            string drawstring11 = "__________________________";

            Font f1 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);
            Font f2 = new Font("Arial", 10, FontStyle.Bold);
            Font f3 = new Font("Arial", 10, FontStyle.Bold | FontStyle.Underline);
            Font f4 = new Font("Arial", 9, FontStyle.Bold);
            Font f5 = new Font("Arial", 7, FontStyle.Bold);
            Font f8 = new Font("Arial", 8, FontStyle.Bold | FontStyle.Underline);
            Font f10 = new Font("PF Barcode 39", 20);
            Font f11 = new Font("Arial", 12, FontStyle.Bold | FontStyle.Underline);
            SolidBrush brush = new SolidBrush(Color.Black);

            PointF p1 = new PointF(40.0F, 15.0F);
            PointF p2 = new PointF(7.0f, 45.0f);
            PointF p3 = new PointF(7.0f, 78.0f);
            PointF p4 = new PointF(7.0F, 95.0F);
            PointF p5 = new PointF(7.0F, 108.0F);
            PointF p6 = new PointF(7.0F, 121.0F);
            PointF p7 = new PointF(7.0F, 140.0F);
            PointF p8 = new PointF(7.0F, producto);
            producto += 25;
            PointF p9 = new PointF(7.0F, producto);
            producto += 30;
            PointF p10 = new PointF(45.0f, producto);
            producto += 30;
            PointF p11 = new PointF(7.0f, producto);

            PrintDocument doc = new PrintDocument();
            doc.PrinterSettings = impresora;
            doc.PrintPage += delegate (object s, PrintPageEventArgs e)
            {
                e.Graphics.DrawString(drawString1, f1, brush, p1);
                e.Graphics.DrawString(drawString2, f2, brush, p2);
                e.Graphics.DrawString(drawString3, f3, brush, p3);
                e.Graphics.DrawString(drawString4, f4, brush, p4);
                e.Graphics.DrawString(drawString5, f4, brush, p5);
                e.Graphics.DrawString(drawString6, f4, brush, p6);
                e.Graphics.DrawString(envasesprod, f5, brush, p7);
                e.Graphics.DrawString(drawstring8, f8, brush, p8);
                e.Graphics.DrawString(drawstring9, f8, brush, p9);
                e.Graphics.DrawString(drawstring10, f10, brush, p10);
                e.Graphics.DrawString(drawstring11, f11, brush, p11);
            };
            doc.Print();
        }
    }
}
