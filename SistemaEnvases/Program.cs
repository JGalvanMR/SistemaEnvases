using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Diagnostics;

namespace SistemaEnvases
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Utilerias.Class1.validar_ip();
            validar_login();
            if (Utilerias.Class1.Login == true)
                Application.Run(new Form1());

            if (Utilerias.Class1.Login == false)
                Application.Exit();
        }

        public static bool validar_login()
        {
            Utilerias.Class1.validar_ip();
            SqlConnection thisConnection = new SqlConnection(Utilerias.Class1.ConnectionString);
            SqlCommand cmnd2;
            SqlDataReader reader2;

            //bool esta = false;
            //Utilerias.Login = false;
            //foreach (Process proceso in Process.GetProcesses())
            //{
            //    if (proceso.ProcessName == "SIPGAB")
            //    {
            //        esta = true;
            //    }
            //}
            //args = new string[1] { "N" };
            //if (esta == false)
            //{
            thisConnection.Open();
            cmnd2 = thisConnection.CreateCommand();
            cmnd2.CommandText = "select TOP 1 usu_login, idioma from tb_cat_historial_dia where nombre_maquina ='" + Environment.MachineName + "' AND fin_sesion IS NULL and sistema ='SIPGAB' ORDER BY inicio_sesion DESC";
            reader2 = cmnd2.ExecuteReader();
            while (reader2.Read())
            {
                Utilerias.Class1.Usu_login = reader2.GetSqlString(0).ToString();
                Utilerias.Class1.Login = true;
                Utilerias.Class1.Idioma = reader2.GetString(1).ToString();
            }
            if (reader2.HasRows == false)
            {
                MessageBox.Show("Error: Favor de iniciar sesion", "AVISO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //System.Diagnostics.Process MiProceso = new System.Diagnostics.Process();
                //MiProceso.StartInfo.WorkingDirectory = "C:\\SisGabWeb"; // ubicacion donde esta el ejecutable
                //MiProceso.StartInfo.FileName = "SIPGAB.exe"; // nombre del archivo a ejecutar con su extension
                ////MiProceso.StartInfo.Arguments = Utilerias.Grupo.ToString(); // esto es opcional en caso que el ejecutablee reciba parametros
                //MiProceso.Start(); // inicia el ejecutable     
                ProcessStartInfo proces = new ProcessStartInfo(@"C:\\SisGabWeb\SIPGAB.exe");
                Process.Start(proces);
                Utilerias.Class1.Login = false;
                Application.Exit();
                //Application.Run(new OrdVenNal());                   
            }
            //}
            thisConnection.Close();
            return Utilerias.Class1.Login;
        }
    }
}
