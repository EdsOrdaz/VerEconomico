using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Management;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace VerEconomico
{
    public partial class Form1 : Form
    {
        private String versiontext = "11.1";
        private String version = "1ce61e3ed2e368a4f295a36748276df98b06b331";
        public static String conexionsqllast = "server=148.223.153.37,5314; database=InfEq;User ID=eordazs;Password=Corpame*2013; integrated security = false ; MultipleActiveResultSets=True";

        public static String servivor = "148.223.153.43\\MSSQLSERVER1";
        public static String basededatos = "bd_SiRAc";
        public static String usuariobd = "sa";
        public static String passbd = "At3n4";
        public static string nsql = "server=" + servivor + "; database=" + basededatos + " ;User ID=" + usuariobd + ";Password=" + passbd + "; integrated security = false ; MultipleActiveResultSets=True";

        public Form1()
        {
            InitializeComponent();
        }

        public static string funcion_wmi(String database, String dato)
        {
            ConnectionOptions options = new ConnectionOptions();
            options.Impersonation = System.Management.ImpersonationLevel.Impersonate;


            ManagementScope scope = new ManagementScope("\\\\" + Environment.MachineName.ToString() + "\\root\\cimv2", options);
            scope.Connect();

            ObjectQuery query = new ObjectQuery("SELECT * FROM " + database);
            ManagementObjectSearcher searcher = new ManagementObjectSearcher(scope, query);

            String cadena = "";
            ManagementObjectCollection queryCollection = searcher.Get();
            foreach (ManagementObject m in queryCollection)
            {
                cadena = m[dato].ToString();
            }
            dato = null;
            database = null;
            if (cadena.Equals(null))
            {
                return "Error";
            }
            return cadena;
        }

        public static String NumSerie()
        {
            String[] data = new string[] {
                   "Win32_BIOS",
                    "SerialNumber",
                    "Sin Resultado"
                };
            return (funcion_wmi(data[0], data[1]) == "") ? data[2] : funcion_wmi(data[0], data[1]);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            //Validar Version
            try
            {
                using (SqlConnection conexion2 = new SqlConnection(conexionsqllast))
                {
                    conexion2.Open();
                    String sql2 = "SELECT (SELECT valor FROM Configuracion WHERE nombre='VerEconomico_Version') as version,valor FROM Configuracion WHERE nombre='VerEconomico_hash'";
                    SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                    SqlDataReader nwReader2 = comm2.ExecuteReader();
                    if (nwReader2.Read())
                    {
                        if (nwReader2["valor"].ToString() != version)
                        {
                            MessageBox.Show("No se puede inciar porque hay una nueva version.\n\nNueva Version: " + nwReader2["version"].ToString() + "\nVersion actual: " + versiontext + "\n\nEl programa se cerrara.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Application.Exit();
                            return;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en validar la version del programa\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }



            try
            {
                using (SqlConnection conexion2 = new SqlConnection(nsql))
                {
                    conexion2.Open();
                    String sql2 = "SELECT * FROM [bd_SiRAc].[dbo].[LNJ_Equipos_Gilberto] where [No. Serie]='"+NumSerie()+"'";
                    SqlCommand comm2 = new SqlCommand(sql2, conexion2);
                    SqlDataReader nwReader2 = comm2.ExecuteReader();

                    if (nwReader2.Read())
                    {
                        dataGridView1.Rows.Add("Número economico:", nwReader2["No. Activo"].ToString());
                        dataGridView1.Rows.Add("Fecha de Alta:", nwReader2["Fecha SIRAC"].ToString());
                        dataGridView1.Rows.Add("Marca:", nwReader2["marca"].ToString());
                        dataGridView1.Rows.Add("Modelo:", nwReader2["modelo"].ToString());
                        dataGridView1.Rows.Add("Número de Serie:", NumSerie());
                        dataGridView1.Rows.Add("Resguardado a:", nwReader2["Nombre de Resguardatario"].ToString());
                        dataGridView1.Rows.Add("Centro de costos:", nwReader2["Centro de Costo"].ToString());
                        dataGridView1.Rows.Add("Base:", nwReader2["Base"].ToString());
                        dataGridView1.Rows.Add("Fecha de Asignación:", nwReader2["Fecha Asignación"].ToString());

                    }
                    System.Windows.Forms.DataGridViewCellStyle boldStyle = new System.Windows.Forms.DataGridViewCellStyle();
                    boldStyle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold);

                    for (int i = 0; i <= 8; i++)
                    {
                        dataGridView1.Rows[i].Cells[0].Style = boldStyle;
                        dataGridView1.Rows[i].Cells[0].Style.BackColor = Color.Silver;
                    }
                    this.dataGridView1.ClearSelection();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error en validar la version del programa\n\nMensaje: " + ex.Message, "Información del Equipo", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
        }
    }
}
