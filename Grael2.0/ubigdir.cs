using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.Drawing;

namespace Grael2
{
    public partial class ubigdir : Form
    {
        public string para1 = "";       // valor a calcular cambio
        public string para2 = "";       // codigo moneda a efectar el cambio
        public string para3 = "";       // fecha del cambio
        //
        DataTable dt = new DataTable();
        // string de conexion
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        //
        libreria lib = new libreria();

        AutoCompleteStringCollection departamentos = new AutoCompleteStringCollection();// autocompletado departamentos
        AutoCompleteStringCollection provincias = new AutoCompleteStringCollection();   // autocompletado provincias
        AutoCompleteStringCollection distritos = new AutoCompleteStringCollection();    // autocompletado distritos
        DataTable dataUbig = (DataTable)CacheManager.GetItem("ubigeos");

        public ubigdir(string param1, string param2, string param3)
        {
            para1 = param1;             // valor original en moneda local se supone
            para2 = param2;             // codigo moneda deseada a cambiar
            para3 = param3;             // fecha del cambio
            InitializeComponent();
        }
        private void ubigdir_Load(object sender, EventArgs e)
        {
            if (dataUbig == null)
            {
                DataTable dataUbig = (DataTable)CacheManager.GetItem("ubigeos");
            }
            // autocompletados
            tx_dptoRtt.AutoCompleteMode = AutoCompleteMode.Suggest;           // departamentos
            tx_dptoRtt.AutoCompleteSource = AutoCompleteSource.CustomSource;  // departamentos
            tx_dptoRtt.AutoCompleteCustomSource = departamentos;              // departamentos
            tx_provRtt.AutoCompleteMode = AutoCompleteMode.Suggest;           // provincias
            tx_provRtt.AutoCompleteSource = AutoCompleteSource.CustomSource;  // provincias
            tx_provRtt.AutoCompleteCustomSource = provincias;                 // provincias
            tx_distRtt.AutoCompleteMode = AutoCompleteMode.Suggest;           // distritos
            tx_distRtt.AutoCompleteSource = AutoCompleteSource.CustomSource;  // distritos
            tx_distRtt.AutoCompleteCustomSource = distritos;                  // distritos

            autodepa();                                     // autocompleta departamentos

            Image salir = Image.FromFile("recursos/Close_32.png");
            button3.Image = salir;
            button3.ImageAlign = ContentAlignment.MiddleCenter;
        }
        private void ubigdir_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SendKeys.Send("{TAB}");
            }
        }   
        public string ReturnValue1 { get; set; }        // valor cambiado a la moneda deseada
        public string ReturnValue2 { get; set; }        // valor en moneda local
        public string ReturnValue3 { get; set; }        // tipo de cambio de la operacion

        private void button1_Click(object sender, EventArgs e)
        {
            ReturnValue1 = tx_ubigRtt.Text;
            this.Close();
        }
        private void button3_Click(object sender, EventArgs e)
        {
            ReturnValue1 = para1;
            ReturnValue2 = para1;
            ReturnValue3 = "0";
            this.Close();
        }

        #region autocompletados
        private void autodepa()                             // se jala en el load
        {
            DataRow[] depar = dataUbig.Select("depart<>'00' and provin='00' and distri='00'");
            departamentos.Clear();
            foreach (DataRow row in depar)
            {
                departamentos.Add(row["nombre"].ToString());
            }
        }
        private void autoprov()                 // se jala despues de ingresado el departamento
        {
            if (tx_dptoRtt.Text.Trim() != "")
            {
                DataRow[] provi = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and provin<>'00' and distri='00'");
                provincias.Clear();
                foreach (DataRow row in provi)
                {
                    provincias.Add(row["nombre"].ToString());
                }
            }
        }
        private void autodist()                 // se jala despues de ingresado la provincia
        {
            if (tx_ubigRtt.Text.Trim() != "" && tx_provRtt.Text.Trim() != "")
            {
                DataRow[] distr = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and provin='" + tx_ubigRtt.Text.Substring(2, 2) + "' and distri<>'00'");
                distritos.Clear();
                foreach (DataRow row in distr)
                {
                    distritos.Add(row["nombre"].ToString());
                }
            }
        }
        #endregion autocompletados

        private void tx_dptoRtt_Leave(object sender, EventArgs e)
        {
            if (tx_dptoRtt.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("nombre='" + tx_dptoRtt.Text.Trim() + "' and provin='00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_ubigRtt.Text = row[0].ItemArray[1].ToString();
                    autoprov();
                }
                else tx_dptoRtt.Text = "";
            }
        }
        private void tx_provRtt_Leave(object sender, EventArgs e)
        {
            if (tx_provRtt.Text != "" && tx_dptoRtt.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and nombre='" + tx_provRtt.Text.Trim() + "' and provin<>'00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_ubigRtt.Text = tx_ubigRtt.Text.Trim().Substring(0, 2) + row[0].ItemArray[2].ToString();
                    autodist();
                }
                else tx_provRtt.Text = "";
            }
        }
        private void tx_distRtt_Leave(object sender, EventArgs e)
        {
            if (tx_distRtt.Text.Trim() != "" && tx_provRtt.Text.Trim() != "" && tx_dptoRtt.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and provin='" + tx_ubigRtt.Text.Substring(2, 2) + "' and nombre='" + tx_distRtt.Text.Trim() + "'");
                if (row.Length > 0)
                {
                    tx_ubigRtt.Text = tx_ubigRtt.Text.Trim().Substring(0, 4) + row[0].ItemArray[3].ToString();
                }
                else tx_distRtt.Text = "";
            }
        }
    }
}
