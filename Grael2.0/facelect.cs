using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using CrystalDecisions.CrystalReports.Engine;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Gma.QrCodeNet.Encoding;
using Gma.QrCodeNet.Encoding.Windows.Render;
using System.Drawing.Imaging;

namespace Grael2
{
    public partial class facelect : Form
    {
        static string nomform = "facelect";             // nombre del formulario
        string colback = Grael2.Program.colbac;   // color de fondo
        string colpage = Grael2.Program.colpag;   // color de los pageframes
        string colgrid = Grael2.Program.colgri;   // color de las grillas
        string colfogr = Grael2.Program.colfog;   // color fondo con grillas
        string colsfon = Grael2.Program.colsbg;   // color fondo seleccion
        string colsfgr = Grael2.Program.colsfc;   // color seleccion grilla
        string colstrp = Grael2.Program.colstr;   // color del strip
        bool conectS = Grael2.Program.vg_conSol;    // usa conector solorsoft? true=si; false=no
        static string nomtab = "madocvtas";         // maestra de comprobantes del ERP_Grael 1.0 

        #region variables
        string img_btN = "";
        string img_btE = "";
        string img_btA = "";            // anula = bloquea
        string img_btP = "";            // imprime
        string img_btV = "";            // visualiza
        string img_bti = "";            // imagen boton inicio
        string img_bts = "";            // imagen boton siguiente
        string img_btr = "";            // imagen boton regresa
        string img_btf = "";            // imagen boton final
        string img_btq = "";
        string img_grab = "";
        string img_anul = "";
        string img_ver = "";
        string vtc_dni = "";            // variable tipo cliente natural
        string vtc_ruc = "";            // variable tipo cliente empresa
        string vtc_ext = "";            // variable tipo cliente extranjero
        string codAnul = "";            // codigo de documento anulado
        string codGene = "";            // codigo documento nuevo generado
        string codCanc = "";            // codigo documento cancelado (pagado 100%)
        string MonDeft = "";            // moneda por defecto
        string v_clu = "";              // codigo del local del usuario
        string v_slu = "";              // serie del local del usuario
        string v_nbu = "";              // nombre del usuario
        string vi_formato = "";         // formato de impresion del documento
        string vi_copias = "";          // cant copias impresion
        //string v_impA5 = "";            // nombre de la impresora matricial
        string v_impTK = "";            // nombre de la ticketera
        //string v_cid = "";              // codigo interno de tipo de documento
        string v_fra2 = "";             // frase que va en obs de cobranza cuando se cancela desde el doc.vta.
        string v_sanu = "";             // serie anulacion interna ANU
        string v_mpag = "";             // medio de pago credito
        string v_codcob = "";           // codigo del documento cobranza
        string v_CR_gr_ind = "";        // nombre del formato FT/BV en CR
        string v_mfildet = "";          // maximo numero de filas en el detalle, coord. con el formato ..... facturas corporativos max 99 o mas filas
        string v_mfdetn = "";           // maximo numero de filas en el detalle de facturas normales
        string vint_A0 = "";            // variable codigo anulacion interna por BD
        string v_codidv = "";           // variable codifo interno de documento de venta en vista TDV
        string codfact = "";            // idcodice de factura
        string codbole = "";            // codigo de boleta electronica
        string v_igv = "";              // valor igv %
        string v_sercob = "";           // serie de cobranza del local
        string v_estcaj = "";           // estado de la caja
        string v_idcaj = "";            // id de la caja actual
        string codAbie = "";            // codigo estado de caja abierta
        string logoclt = "";            // ruta y nombre archivo logo
        string fshoy = "";              // fecha hoy del servidor en formato ansi
        string codppc = "";             // codigo del plazo de pago por defecto para fact a crédito
        string codsuser_cu = "";        // usuarios autorizados a crear Ft de cargas unicas
        int v_cdpa = 0;                 // cantidad de días despues de emitida la fact. en que un usuario normal puede anular
        string vint_gg = "";            // glosa del detalle inicial de la guía "sin verificar contenido"
        decimal limbolsd = 0;           // limite en soles para boletas sin direccion
        string mpdef = "";              // medio de pago por defecto, efectivo
        string texpagC = "";            // texto para el pago GLOSA, CONTADO
        string texpagD = "";            // texto para el pago GLOSA, CREDITO
        string texpag2 = "";            // texto para las cuotas
        string texpag3 = "";            // texto para la fecha vencimiento
        string glocopa = "";            // glosa de condicion de pago
        string ppauto = "SI";           // SI o NO, si el comprobante es credito, el plazo de pago es automatico?
        string ufcorp = "";             // usuarios que pueden generar comprobantes + de 3 guias - corporativos
        int v_iabol = 0;                // contador inicial para anulaciones de boletas .. RC
        //
        string leyg_sg = "";            // leyenda para la guia de la fact
        string rutatxt = "";            // ruta de los txt para la fact. electronica
        string tipdo = "";              // CODIGO SUNAT tipo de documento de venta
        string tipoDocEmi = "";         // CODIGO SUNAT tipo de documento RUC/DNI
        string tipoMoneda = "";         // CODIGO SUNAT tipo de moneda
        string codleyt = "";            // codigo leyenda detraccion
        string leydet1 = "";            // glosa para las operaciones con detraccion 1
        string leydet2 = "";            // glosa para las operaciones con detraccion 2
        string glosser = "";            // glosa que va en el detalle del doc. de venta
        string glosser2 = "";           // glosa 2 que va despues de la glosa principal
        string restexto = "xxx";        // texto resolucion sunat autorizando prov. fact electronica
        string autoriz_bizlinks = "";   // numero resolucion sunat autorizando prov. fact electronica
        string cdrdef = "";             // cdr por defecto
        string despe2 = "";             // texto para mensajes al cliente al final de la impresión del doc.vta. 
        string provee = "";             // direccion web del ose o pse para la descarga del 
        string correo_gen = "";         // correo generico del emisor cuando el cliente no tiene correo propio
        string codusanu = "";           // usuarios que pueden anular fuera de plazo
        string cusdscto = "";           // usuarios que pueden hacer descuentos
        string otro = "";               // ruta y nombre del png código QR
        string caractNo = "";           // caracter prohibido en campos texto, caracter delimitador para los TXT
        string nipfe = "";              // nombre identificador del proveedor de fact electronica
        string glosaAnul = "";          // texto motivo de baja/anulacion en los TXT para el pse/ose
        string tipdocAnu = "";          // Tipos de documentos que se pueden dar de baja
        string tdocsBol = "";           // tipos de documentos de clientes que permiten boletas
        string tdocsFac = "";           // tipos de documentos de clientes que permiten facturas
        string texmotran = "";          // texto modalidad de transporte
        string codtxmotran = "";        // codigo motivo de traslado de bienes
        //
        static libreria lib = new libreria();   // libreria de procedimientos
        publico lp = new publico();             // libreria de clase

        string verapp = System.Diagnostics.FileVersionInfo.GetVersionInfo(Application.ExecutablePath).FileVersion;
        string nomclie = Program.cliente;           // cliente usuario del sistema
        string rucclie = Program.ruc;               // ruc del cliente usuario del sistema
        string ubiclie = Program.ubidirfis;         // ubigeo direc fiscal
        string asd = Grael2.Program.vg_user;    // usuario conectado al sistema
        string dirloc = Grael2.Program.vg_duse; // direccion completa del local usuario conectado
        string ubiloc = Grael2.Program.vg_uuse; // ubigeo local del usuario conectado
        #endregion
        
        AutoCompleteStringCollection departamentos = new AutoCompleteStringCollection();// autocompletado departamentos
        AutoCompleteStringCollection provincias = new AutoCompleteStringCollection();   // autocompletado provincias
        AutoCompleteStringCollection distritos = new AutoCompleteStringCollection();    // autocompletado distritos
        DataTable dataUbig = (DataTable)CacheManager.GetItem("ubigeos");

        // string de conexion
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        string db_conn_grael = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.dataG + ";";

        DataTable dttd0 = new DataTable();
        DataTable dttd1 = new DataTable();
        DataTable dtm = new DataTable();        // moneda
        DataTable dtp = new DataTable();        // plazo de credito 
        DataTable tcfe = new DataTable();       // facturacion electronica - cabecera
        DataTable tdfe = new DataTable();       // facturacion electronica -detalle
        DataTable mpa = new DataTable();        // pago mpa - contado efectivo, credito.
        DataTable dtru = new DataTable();       // rutas cargas unicas detraccion
        public string script = "";              // script de conexion a Bizlinks
        NumLetra nl = new NumLetra();
        string[] datcltsR = { "", "", "", "", "", "", "", "", "", "" };
        string[] datcltsD = { "", "", "", "", "", "", "", "", "", "" };
        string[] datguias = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "","", "" }; // 20
        string[] datcargu = { "", "", "", "", "", "", "", "", "", "", "", "", "", "" };    // 14
        public facelect()
        {
            InitializeComponent();
        }
        private void facelect_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N) Bt_add.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E) Bt_edit.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.A) Bt_anul.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O) Bt_ver.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P) Bt_print.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S) Bt_close.PerformClick();
        }
        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)    // F1
        {
            string para1 = "";
            string para2 = "";
            string para3 = "";
            if (keyData == Keys.F1 && tx_e_ubiori.Focused == true && Tx_modo.Text == "NUEVO")
            {
                para1 = "";
                para2 = "";
                para3 = "";
                ubigdir ayu3 = new ubigdir(para1, para2, para3);
                var result = ayu3.ShowDialog();
                if (result == DialogResult.Cancel)
                {
                    tx_e_ubiori.Text = ayu3.ReturnValue1;
                }
                return true;    // indicate that you handled this keystroke
            }
            if (keyData == Keys.F1 && tx_e_ubides.Focused == true && Tx_modo.Text == "NUEVO")
            {
                para1 = "";
                para2 = "";
                para3 = "";
                ubigdir ayu3 = new ubigdir(para1, para2, para3);
                var result = ayu3.ShowDialog();
                if (result == DialogResult.Cancel)  // deberia ser OK, pero que chuuu
                {
                    tx_e_ubides.Text = ayu3.ReturnValue1;
                }
                return true;    // indicate that you handled this keystroke
            }
            // Call the base class
            return base.ProcessCmdKey(ref msg, keyData);
        }
        private void facelect_Load(object sender, EventArgs e)
        {
            /*
            ToolTip toolTipNombre = new ToolTip();           // Create the ToolTip and associate with the Form container.
            // Set up the delays for the ToolTip.
            toolTipNombre.AutoPopDelay = 5000;
            toolTipNombre.InitialDelay = 1000;
            toolTipNombre.ReshowDelay = 500;
            toolTipNombre.ShowAlways = true;                 // Force the ToolTip text to be displayed whether or not the form is active.
            toolTipNombre.SetToolTip(toolStrip1, nomform);   // Set up the ToolTip text for the object
            */
            this.Focus();
            jalainfo();
            init();
            dataload();
            toolboton();
            this.KeyPreview = true;
            autodepa();                                     // autocompleta departamentos
            if (valiVars() == false)
            {
                Application.Exit();
                return;
            }
        }
        private void init()
        {
            this.Width = 683;
            this.BackColor = Color.FromName(colback);
            toolStrip1.BackColor = Color.FromName(colstrp);
            dataGridView1.DefaultCellStyle.BackColor = Color.FromName(colgrid);
            //dataGridView1.DefaultCellStyle.BackColor = Color.FromName(colgrid);
            //dataGridView1.DefaultCellStyle.ForeColor = Color.FromName(colfogr);
            //dataGridView1.DefaultCellStyle.SelectionBackColor = Color.FromName(colsfon);
            //dataGridView1.DefaultCellStyle.SelectionForeColor = Color.FromName(colsfgr);
            //
            tx_user.Text += asd;
            tx_nomuser.Text = Grael2.Program.vg_nuse;   // lib.nomuser(asd);
            //tx_locuser.Text = Grael2.Program.vg_luse;  // lib.locuser(asd);
            tx_locuser.Text = tx_locuser.Text + " " + Grael2.Program.vg_nlus;
            tx_fechact.Text = DateTime.Today.ToString();
            //
            Bt_add.Image = Image.FromFile(img_btN);
            Bt_edit.Image = Image.FromFile(img_btE);
            Bt_anul.Image = Image.FromFile(img_btA);
            Bt_ver.Image = Image.FromFile(img_btV);
            Bt_print.Image = Image.FromFile(img_btP);
            Bt_close.Image = Image.FromFile(img_btq);
            Bt_ini.Image = Image.FromFile(img_bti);
            Bt_sig.Image = Image.FromFile(img_bts);
            Bt_ret.Image = Image.FromFile(img_btr);
            Bt_fin.Image = Image.FromFile(img_btf);
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
            // longitudes maximas de campos
            tx_serie.MaxLength = 3;             // serie doc vta
            tx_numero.MaxLength = 7;            // numero doc vta
            tx_serGR.MaxLength = 3;             // serie guia
            tx_numGR.MaxLength = 7;             // numero guia
            tx_numDocRem.MaxLength = 11;        // ruc o dni cliente
            tx_dirRem.MaxLength = 100;
            tx_nomRem.MaxLength = 100;          // nombre remitente
            tx_distRtt.MaxLength = 25;
            tx_provRtt.MaxLength = 25;
            tx_dptoRtt.MaxLength = 25;
            tx_obser1.MaxLength = 150;
            tx_telc1.MaxLength = 12;
            tx_telc2.MaxLength = 12;
            tx_fletLetras.MaxLength = 249;
            //
            tx_e_aut.MaxLength = 15;
            tx_e_dirlle.MaxLength = 150;
            tx_e_dirpar.MaxLength = 150;
            tx_e_dnicho.MaxLength = 8;
            tx_e_glos1.MaxLength = 245;
            tx_e_glos2.MaxLength = 245;
            tx_e_glos3.MaxLength = 245;
            tx_e_nfv.MaxLength = 10;
            tx_e_ntrans.MaxLength = 150;
            tx_e_placa.MaxLength = 10;
            tx_e_ruct.MaxLength = 11;
            // grilla
            dataGridView1.ReadOnly = true;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            // todo desabilidado
            sololee();
            // jalamos datos de fact electrónica
            OSE_Fact_Elect();
        }
        private void initIngreso()
        {
            limpiar();
            limpia_chk();
            limpia_otros();
            limpia_combos();
            cmb_tdv.SelectedIndex = -1;
            dataGridView1.Rows.Clear();
            dataGridView1.ReadOnly = true;
            tx_igv.Text = "";
            tx_subt.Text = "";
            tx_flete.Text = "";
            tx_pagado.Text = "";
            tx_salxcob.Text = "";
            tx_numero.Text = "";
            tx_serie.Text = v_slu;
            tx_numero.ReadOnly = true;
            tx_dat_mone.Text = MonDeft;
            cmb_mon.SelectedValue = tx_dat_mone.Text;
            tx_fechope.Text = DateTime.Today.ToString("dd/MM/yyyy");
            tx_digit.Text = v_nbu;
            tx_dat_estad.Text = codGene;
            tx_estado.Text = lib.nomstat(tx_dat_estad.Text);
            tx_idcaja.ReadOnly = true;
            tx_idcaja.Text = "";
            tx_fletLetras.ReadOnly = true;
            if (Tx_modo.Text == "NUEVO" && v_estcaj == codAbie)      // caja esta abierta?
            {
                if (fshoy != Grael2.Program.vg_fcaj)  // fecha de la caja vs fecha de hoy
                {
                    MessageBox.Show("Las fechas no coinciden" + Environment.NewLine +
                        "Fecha de caja vs Fecha actual", "Caja fuera de fecha", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    //return;
                }
                else
                {
                    tx_idcaja.Text = v_idcaj;
                }
            }
            if (Tx_modo.Text == "NUEVO")
            {
                tx_tfil.Text = "0";
                rb_si.Enabled = true;
                rb_no.Enabled = true;
                if (cusdscto.Contains(asd)) tx_flete.ReadOnly = false;
                else tx_flete.ReadOnly = true;
                rb_no.PerformClick();
                // 
                chk_consol.Enabled = false;
                if (ufcorp.Contains(asd)) chk_consol.Enabled = true;
            }
            tx_dat_nombd.Text = "Bultos";
            tx_dat_nombd.ReadOnly = true;

            rb_fnor.Checked = true;
            inivarGR();
        }
        private void jalainfo()                  // obtiene datos de imagenes y variables
        {
            try
            {
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                string consulta = "select formulario,campo,param,valor from enlaces where formulario in (@nofo,@nfin,@nofa,@nofi,@noca,@noco,@nocg)";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                micon.Parameters.AddWithValue("@nofo", "main");
                micon.Parameters.AddWithValue("@nfin", "interno");
                micon.Parameters.AddWithValue("@nofi", "clients");
                micon.Parameters.AddWithValue("@noco", "cobranzas");
                micon.Parameters.AddWithValue("@noca", "ayccaja");
                micon.Parameters.AddWithValue("@nocg", "guiati_a");
                micon.Parameters.AddWithValue("@nofa", nomform);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                for (int t = 0; t < dt.Rows.Count; t++)
                {
                    DataRow row = dt.Rows[t];
                    if (row["formulario"].ToString() == "main")
                    {
                        if (row["campo"].ToString() == "imagenes")
                        {
                            if (row["param"].ToString() == "img_btN") img_btN = row["valor"].ToString().Trim();         // imagen del boton de accion NUEVO
                            if (row["param"].ToString() == "img_btE") img_btE = row["valor"].ToString().Trim();         // imagen del boton de accion EDITAR
                            if (row["param"].ToString() == "img_btA") img_btA = row["valor"].ToString().Trim();         // imagen del boton de accion ANULAR/BORRAR
                            if (row["param"].ToString() == "img_btQ") img_btq = row["valor"].ToString().Trim();         // imagen del boton de accion SALIR
                            if (row["param"].ToString() == "img_btP") img_btP = row["valor"].ToString().Trim();         // imagen del boton de accion IMPRIMIR
                            if (row["param"].ToString() == "img_btV") img_btV = row["valor"].ToString().Trim();         // imagen del boton de accion visualizar
                            if (row["param"].ToString() == "img_bti") img_bti = row["valor"].ToString().Trim();         // imagen del boton de accion IR AL INICIO
                            if (row["param"].ToString() == "img_bts") img_bts = row["valor"].ToString().Trim();         // imagen del boton de accion SIGUIENTE
                            if (row["param"].ToString() == "img_btr") img_btr = row["valor"].ToString().Trim();         // imagen del boton de accion RETROCEDE
                            if (row["param"].ToString() == "img_btf") img_btf = row["valor"].ToString().Trim();         // imagen del boton de accion IR AL FINAL
                            if (row["param"].ToString() == "img_gra") img_grab = row["valor"].ToString().Trim();         // imagen del boton grabar nuevo
                            if (row["param"].ToString() == "img_anu") img_anul = row["valor"].ToString().Trim();         // imagen del boton grabar anular
                            if (row["param"].ToString() == "img_preview") img_ver = row["valor"].ToString().Trim();      // imagen del boton grabar visualizar
                            if (row["param"].ToString() == "logoPrin") logoclt = row["valor"].ToString().Trim();         // logo emisor
                        }
                        if (row["campo"].ToString() == "estado")
                        {
                            if (row["param"].ToString() == "anulado") codAnul = row["valor"].ToString().Trim();         // codigo doc anulado
                            if (row["param"].ToString() == "generado") codGene = row["valor"].ToString().Trim();        // codigo doc generado
                            if (row["param"].ToString() == "cancelado") codCanc = row["valor"].ToString().Trim();        // codigo doc cancelado
                        }
                        if (row["campo"].ToString() == "rutas")
                        {
                            if (row["param"].ToString() == "fe_txt") rutatxt = row["valor"].ToString().Trim();         // ruta de los txt para la fact. electronica
                        }
                    }
                    if (row["formulario"].ToString() == "clients" && row["campo"].ToString() == "documento")
                    {
                        if (row["param"].ToString() == "dni") vtc_dni = row["valor"].ToString().Trim();
                        if (row["param"].ToString() == "ruc") vtc_ruc = row["valor"].ToString().Trim();
                        if (row["param"].ToString() == "ext") vtc_ext = row["valor"].ToString().Trim();
                    }
                    if (row["formulario"].ToString() == "cobranzas" && row["campo"].ToString() == "documento")
                    {
                        if (row["param"].ToString() == "codigo") v_codcob = row["valor"].ToString().Trim();
                    }
                    if (row["formulario"].ToString() == nomform)
                    {
                        if (row["campo"].ToString() == "documento")
                        {
                            if (row["param"].ToString() == "frase2") v_fra2 = row["valor"].ToString().Trim();               // frase cuando se cancela el doc.vta.
                            if (row["param"].ToString() == "serieAnu") v_sanu = row["valor"].ToString().Trim();               // serie anulacion interna
                            if (row["param"].ToString() == "mpacre") v_mpag = row["valor"].ToString().Trim();               // medio de pago credito
                            if (row["param"].ToString() == "factura") codfact = row["valor"].ToString().Trim();               // codigo doc.venta factura
                            if (row["param"].ToString() == "boleta") codbole = row["valor"].ToString().Trim();               // codigo doc.venta boleta
                            if (row["param"].ToString() == "plazocred") codppc = row["valor"].ToString().Trim();               // codigo plazo de pago x defecto para fact. a CREDITO
                            if (row["param"].ToString() == "autPago") ppauto = row["valor"].ToString().Trim();               // SI o NO, seleccion plazo automatico o no
                            if (row["param"].ToString() == "usercar_unic") codsuser_cu = row["valor"].ToString().Trim();    // usuarios autorizados a crear Ft de cargas unicas
                            if (row["param"].ToString() == "diasanul") v_cdpa = int.Parse(row["valor"].ToString());         // cant dias en que usuario normal puede anular 
                            if (row["param"].ToString() == "useranul") codusanu = row["valor"].ToString();                  // usuarios autorizados a anular fuera de plazo 
                            if (row["param"].ToString() == "userdscto") cusdscto = row["valor"].ToString();                 // usuarios que pueden hacer descuentos
                            if (row["param"].ToString() == "cltesBol") tdocsBol = row["valor"].ToString();                  // tipos de documento de clientes para boletas
                            if (row["param"].ToString() == "cltesFac") tdocsFac = row["valor"].ToString();                  // tipo de documentos para facturas
                            if (row["param"].ToString() == "limbolsd") limbolsd = decimal.Parse(row["valor"].ToString());   // limite soles para boletas sin direccion
                            if (row["param"].ToString() == "ufcorp") ufcorp = row["valor"].ToString();                      // usuarios que emiten FT consolidados (corporativos 100 gr)
                        }
                        if (row["campo"].ToString() == "impresion")
                        {
                            if (row["param"].ToString() == "formato") vi_formato = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "filasDet") v_mfildet = row["valor"].ToString().Trim();       // maxima cant de filas de detalle en fact consolidadas
                            if (row["param"].ToString() == "mfdetn") v_mfdetn = row["valor"].ToString().Trim();          // maxima cant de filas de detalle en fact normales
                            if (row["param"].ToString() == "copias") vi_copias = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "impTK") v_impTK = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "nomfor_cr") v_CR_gr_ind = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "acceso") despe2 = row["valor"].ToString();
                            if (row["param"].ToString() == "glosTkC1") texpagC = row["valor"].ToString();            // texto para el pago GLOSA, CONTADO 
                            if (row["param"].ToString() == "glosTkD1") texpagD = row["valor"].ToString();            // texto para el pago GLOSA, CREDITO
                            if (row["param"].ToString() == "glosTkp2") texpag2 = row["valor"].ToString();            // texto para las cuotas
                            if (row["param"].ToString() == "glosTkp3") texpag3 = row["valor"].ToString();            // texto para el vencimiento
                            if (row["param"].ToString() == "gloconpa") glocopa = row["valor"].ToString();            // glosa para la condicion de pago
                        }
                        if (row["campo"].ToString() == "moneda" && row["param"].ToString() == "default") MonDeft = row["valor"].ToString().Trim();      // moneda por defecto soles
                        if (row["campo"].ToString() == "detraccion" && row["param"].ToString() == "glosa") leydet1 = row["valor"].ToString().Trim();    // glosa detraccion
                        if (row["campo"].ToString() == "detraccion" && row["param"].ToString() == "leyenda2") leydet2 = row["valor"].ToString().Trim();
                        if (row["campo"].ToString() == "detraccion" && row["param"].ToString() == "codley") codleyt = row["valor"].ToString().Trim();
                        if (row["campo"].ToString() == "factelect")
                        {
                            /*
                            if (row["param"].ToString() == "textaut") restexto = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "autoriz") autoriz_OSE_PSE = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "despedi") despedida = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "webose") webose = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "caracterNo") caractNo = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "ose-pse") nipfe = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "motivoBaja") glosaAnul = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "tipsDocbaja") tipdocAnu = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "modTran") texmotran = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "codmotTran") codtxmotran = row["valor"].ToString().Trim();
                            */
                            if (row["param"].ToString() == "correo_c1") correo_gen = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "autoriz") autoriz_bizlinks = row["valor"].ToString();                          // resolución de autorizacion de bizlinks                            
                            if (row["param"].ToString() == "webose") provee = row["valor"].ToString();
                            if (row["param"].ToString() == "textaut") restexto = row["valor"].ToString();
                            if (row["campo"].ToString() == "cdr" && row["param"].ToString() == "Default")
                            {
                                cdrdef = row["valor"].ToString();       // cdr por defecto
                            }
                            if (row["campo"].ToString() == "leyguia" && row["param"].ToString() == "sg")
                            {
                                leyg_sg = row["valor"].ToString();      // leyenda para impresion "S/G :"
                            }
                            if (row["param"].ToString() == "inibolanu") v_iabol = int.Parse(row["valor"].ToString());
                        }
                    }
                    if (row["formulario"].ToString() == "ayccaja" && row["campo"].ToString() == "estado")
                    {
                        if (row["param"].ToString() == "abierto") codAbie = row["valor"].ToString().Trim();             // codigo caja abierta
                        //if (row["param"].ToString() == "cerrado") codCier = row["valor"].ToString().Trim();             // codigo caja cerrada
                    }
                    if (row["formulario"].ToString() == "interno")              // codigo enlace interno de anulacion del cliente con en BD A0
                    {
                        if (row["campo"].ToString() == "anulado" && row["param"].ToString() == "A0") vint_A0 = row["valor"].ToString().Trim();
                        if (row["campo"].ToString() == "codinDV" && row["param"].ToString() == "DV") v_codidv = row["valor"].ToString().Trim();           // codigo de dov.vta en tabla TDV
                        if (row["campo"].ToString() == "igv" && row["param"].ToString() == "%") v_igv = row["valor"].ToString().Trim();
                    }
                    if (row["formulario"].ToString() == "guiati_a")
                    {
                        if (row["campo"].ToString() == "detalle" && row["param"].ToString() == "glosa") vint_gg = row["valor"].ToString().Trim();
                    } 
                }
                da.Dispose();
                dt.Dispose();
                // jalamos datos del usuario y local
                v_clu = Grael2.Program.vg_luse;                // codigo local usuario
                v_slu = lib.serlocs(v_clu);                    // serie local usuario
                v_nbu = Grael2.Program.vg_nuse;                // nombre del usuario
                conn.Close();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show(ex.Message, "Error de conexión");
                Application.Exit();
                return;
            }
        }
        private void jalaoc(string campo)        // jala doc venta
        {
            rb_fnor.Checked = true;
            //try
            {
                string parte = "";
                if (campo == "tx_idr")
                {
                    parte = "where a.id=@ida";
                }
                if (campo == "sernum")
                {
                    parte = "where a.docvta=@tdv and a.servta=@ser and a.corvta=@num and a.mfe=@mfe";
                }
                MySqlConnection conn = new MySqlConnection(db_conn_grael);
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    string consulta = "select a.id,a.fechope,a.mfe,a.tipcam,a.servta,a.corvta,a.rd3,a.doccli,a.numdcli,a.nomclie,a.direc,a.dpto,a.prov,a.dist,a.ubiclte,a.email,d.telef1," +
                        "a.local,'' as dirorig,'' as ubiorig,a.observ,0 as canfidt,0 as canbudt,a.moneda,a.tipcam,a.subtot,a.igv,a.pigv,a.doctot,0 as totpags,a.saldo,a.status,'' as frase01,'S' as impreso," +
                        "'' as tipoclt,'' as m1clien,a.frecepf,a.userc,a.fechc,a.userm,a.fechm,b.descrizionerid as nomest,ifnull(c.id, '') as cobra,0 as idcaja," +
                        "0 as porcendscto,a.dscto,a.docvta,a.tippago,a.plazocred,a.condpag," +
                        "f.tipoAd,f.placa,f.placa2,f.confv,f.autoriz,f.cargaEf,f.cargaUt,f.rucTrans,f.nomTrans,f.fecIniTras,f.dirPartida,f.ubiPartida,f.dirDestin,f.ubiDestin," +
                        "f.dniChof,f.brevete,f.valRefViaje,f.valRefVehic,f.valRefTon,f.precioTN,f.pesoTN,f.glosa1,f.glosa2,f.glosa3,f.detMon1,f.detMon2,f.detMon3," +
                        "f.ruta,f.valruta,f.detrac,f.valrefcu,f.cargaUt,f.valRefViaje " +
                        "from madocvtas a left join desc_sit b on b.idcodice = a.status " +
                        "left join macobran c on c.docvta = a.docvta and c.servta = a.servta and c.corvta = a.corvta and c.status<> @coda " +
                        "left join anag_cli d on d.ruc = a.numdcli and d.docu = a.doccli " +
                        "LEFT JOIN adifactu f ON f.idc=a.id " 
                        + parte;
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@coda", codAnul);
                    if (campo == "tx_idr")
                    {
                        micon.Parameters.AddWithValue("@ida", tx_idr.Text);
                    }
                    if (campo == "sernum")
                    {
                        micon.Parameters.AddWithValue("@tdv", tx_dat_tdv.Text);
                        micon.Parameters.AddWithValue("@ser", tx_serie.Text);
                        micon.Parameters.AddWithValue("@num", tx_numero.Text);
                        micon.Parameters.AddWithValue("@mfe", cmb_tdv.Text.Substring(0,1));
                    }
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr != null)
                    {
                        if (dr.Read())
                        {
                            tx_idr.Text = dr.GetString("id");
                            tx_idcaja.Text = dr.GetString("idcaja");
                            tx_fechope.Text = dr.GetString("fechope").Substring(0, 10);
                            //.Text = dr.GetString("martdve");
                            tx_dat_tdv.Text = dr.GetString("docvta");
                            tx_serie.Text = dr.GetString("servta");
                            tx_numero.Text = dr.GetString("corvta");
                            rb_remGR.Checked = (dr.GetString("rd3") == "1")? true : false;
                            rb_desGR.Checked = (dr.GetString("rd3") == "2") ? true : false;
                            rb_otro.Checked = (dr.GetString("rd3") == "3") ? true : false;
                            tx_dat_tdRem.Text = dr.GetString("doccli");
                            tx_numDocRem.Text = dr.GetString("numdcli");
                            tx_nomRem.Text = dr.GetString("nomclie");
                            tx_dirRem.Text = dr.GetString("direc");
                            tx_dptoRtt.Text = dr.GetString("dpto");
                            tx_provRtt.Text = dr.GetString("prov");
                            tx_distRtt.Text = dr.GetString("dist");
                            tx_ubigRtt.Text = dr.GetString("ubiclte");
                            tx_email.Text = dr.GetString("email");
                            tx_telc1.Text = dr.GetString("telef1");
                            tx_dat_loca.Text = dr.GetString("local");
                            //locorig,dirorig,ubiorig
                            tx_obser1.Text = dr.GetString("observ");
                            tx_tfil.Text = dr.GetString("canfidt");
                            tx_totcant.Text = dr.GetString("canbudt");  // total bultos
                            tx_dat_mone.Text = dr.GetString("moneda");
                            tx_tipcam.Text = dr.GetString("tipcam");
                            tx_subt.Text = Math.Round(dr.GetDecimal("subtot"),2).ToString();
                            tx_igv.Text = Math.Round(dr.GetDecimal("igv"), 2).ToString();
                            //,,,porcigv
                            tx_flete.Text = Math.Round(dr.GetDecimal("doctot"),2).ToString();           // total inc. igv
                            tx_pagado.Text = dr.GetString("totpags");
                            tx_salxcob.Text = dr.GetString("saldo");
                            tx_dat_estad.Text = dr.GetString("status");        // estado
                            tx_dat_tcr.Text = dr.GetString("tipoclt");          // tipo de cliente credito o contado
                            tx_dat_m1clte.Text = dr.GetString("m1clien");
                            tx_impreso.Text = dr.GetString("impreso");
                            tx_idcob.Text = dr.GetString("cobra");              // id de cobranza
                            //
                            cmb_tdv.SelectedValue = tx_dat_tdv.Text;
                            cmb_tdv_SelectedIndexChanged(null, null);
                            tx_numero.Text = dr.GetString("corvta");       // al cambiar el indice en el combox se borra numero, por eso lo volvemos a jalar
                            cmb_docRem.SelectedValue = tx_dat_tdRem.Text;
                            cmb_mon.SelectedValue = tx_dat_mone.Text;
                            tx_estado.Text = dr.GetString("nomest");   // lib.nomstat(tx_dat_estad.Text);
                            if (dr.GetString("userm") == "") tx_digit.Text = lib.nomuser(dr.GetString("userc"));
                            else tx_digit.Text = lib.nomuser(dr.GetString("userm"));
                            if (dr.GetString("condpag") == "E") rb_si.Checked = true;
                            if (dr.GetString("condpag") == "R") rb_no.Checked = true;
                            if (dr.GetString("condpag") == "C") rb_cre.Checked = true;
                            if (dr.GetString("tippago") == "")
                            {
                                // es anterior al tiempo que se ponía contado o credito
                                // MessageBox.Show("Que hago aca?");
                            }
                            tx_dat_tipag.Text = dr.GetString("tippago");
                            cmb_tipop.SelectedValue = tx_dat_tipag.Text;
                            tx_dat_dpla.Text = dr.GetString("plazocred");
                            cmb_plazoc.SelectedValue = tx_dat_dpla.Text;
                            // campos de carga unica
                            tx_valdscto.Text = dr.GetString("dscto");
                            tx_dat_porcDscto.Text = dr.GetString("porcendscto");
                            if (dr["tipoAd"].ToString() != "")   // campos de factura especial
                            {
                                rb_fesp.Checked = true;
                                tx_e_aut.Text = dr.GetString("autoriz");
                                tx_e_cant.Text = dr.GetString("pesoTN");
                                tx_e_dirlle.Text = dr.GetString("dirDestin");
                                tx_e_dirpar.Text = dr.GetString("dirPartida");
                                tx_e_dnicho.Text = dr.GetString("dniChof");
                                tx_e_ftras.Text = dr.GetString("fecIniTras");
                                tx_e_glos1.Text = dr.GetString("glosa1");
                                tx_e_glos2.Text = dr.GetString("glosa2");
                                tx_e_glos3.Text = dr.GetString("glosa3");
                                tx_e_nfv.Text = dr.GetString("confv");
                                tx_e_ntrans.Text = dr.GetString("nomTrans");
                                tx_e_placa.Text = dr.GetString("placa");
                                tx_e_prec.Text = dr.GetString("precioTN");
                                tx_e_ruct.Text = dr.GetString("rucTrans");
                                tx_e_ubides.Text = dr.GetString("ubiDestin");
                                tx_e_ubiori.Text = dr.GetString("ubiPartida");
                                tx_dat_ruta.Text = dr.GetString("ruta");
                                tx_sxtm.Text = dr.GetString("valruta");
                                tx_e_valref.Text = dr.GetString("valRefViaje");
                                tx_detrac.Text = dr.GetString("detrac");
                                tx_e_carut.Text = dr.GetString("valrefcu");
                                tx_e_tn.Text = dr.GetString("cargaUt");
                                cmb_ruta.SelectedValue = tx_dat_ruta.Text;
                            }
                        }
                        else
                        {
                            MessageBox.Show("No existe el número del documento de venta!", "Atención - dato incorrecto",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                            tx_numero.Text = "";
                            tx_numero.Focus();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No existe el número buscado!", "Atención - dato incorrecto",
                            MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    }
                    dr.Dispose();
                    micon.Dispose();
                    //
                    if (decimal.Parse(tx_valdscto.Text) > 0)
                    {
                        lin_dscto.Visible = true;
                        lb_dscto.Visible = true;
                        tx_valdscto.Visible = true;
                    }
                    else
                    {
                        lin_dscto.Visible = false;
                        lb_dscto.Visible = false;
                        tx_valdscto.Visible = false;
                    }
                    NumLetra numLetra = new NumLetra();
                    DataRow[] row = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");
                    tx_fletLetras.Text = numLetra.Convertir(tx_flete.Text, true) + row[0][3].ToString().Trim();
                }
                conn.Close();
            }
        }
        private void jaladet(string idr)         // jala el detalle
        {
            /*
            string jalad = "select filadet,codgror,cantbul,unimedp,descpro,pesogro,codmogr,totalgr " +
                "from detfactu where idc=@idr";
            */
            string jalad = "select a.idc,a.docvta,a.servta,a.corvta,a.sergr,a.corgr,a.moneda,a.valor,a.ruta,a.glosa,a.status," +
                        "a.userc,a.fechc,a.docremi,a.bultos,a.monrefd1,a.monrefd2,a.monrefd3,c.moneda as nomMon,c.fechope,c.docremi," +
                        "concat(lo.descrizionerid,'-',ld.descrizionerid),c.saldo,SUBSTRING_INDEX(SUBSTRING_INDEX(a.bultos, ' ', 2), ' ', -1) AS unidad,a.valorel " +
                        "from detavtas a left join desc_mon b on b.idcodice=a.moneda " +
                        "left join magrem c on c.sergre=a.sergr and c.corgre=a.corgr " +
                        "left join desc_sds lo on lo.idcodice = c.origen " +
                        "left join desc_sds ld on ld.idcodice = c.destino " +
                        "where a.idc=@idr";
            using (MySqlConnection conn = new MySqlConnection(db_conn_grael))
            {
                conn.Open();
                using (MySqlCommand micon = new MySqlCommand(jalad, conn))
                {
                    micon.Parameters.AddWithValue("@idr", idr);
                    using (MySqlDataAdapter da = new MySqlDataAdapter(micon))
                    {
                        DataTable dt = new DataTable();
                        da.Fill(dt);
                        foreach (DataRow row in dt.Rows)
                        {
                            dataGridView1.Rows.Add(
                                row[4].ToString() + "-" + row[5].ToString(),
                                row[9].ToString(),
                                row[14].ToString(),
                                row[6].ToString(),
                                row[7].ToString(),
                                "",
                                row[18].ToString(),
                                row[19].ToString().Substring(0,10),
                                row[20].ToString(),
                                "",
                                row[21].ToString(),
                                row[22].ToString(),
                                row[23].ToString(),
                                row[24].ToString());
                        }
                        dt.Dispose();
                    }
                }
            }
        }
        public void dataload()                  // jala datos para los combos 
        {
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                while (true)
                {
                    try
                    {
                        conn.Open();
                        break;
                    }
                    catch (MySqlException ex)
                    {
                        var aa = MessageBox.Show(ex.Message + Environment.NewLine + "No se pudo conectar con el servidor" + Environment.NewLine +
                            "Desea volver a intentarlo?", "Error de conexión", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (aa == DialogResult.No)
                        {
                            Application.Exit();
                            return;
                        }
                    }
                }
                // datos para el combobox documento de venta
                cmb_tdv.Items.Clear();
                string consu = "select distinct a.idcodice,a.descrizionerid,a.enlace1,a.codsunat,b.glosaser,b.serie,a.deta1 " +
                    "from desc_tdo a LEFT JOIN series b ON b.tipdoc = a.IDCodice where a.numero=@bloq and a.codigo=@codv and b.sede=@loca";
                using (MySqlCommand cdv = new MySqlCommand(consu, conn))
                {
                    cdv.Parameters.AddWithValue("@bloq", 1);
                    cdv.Parameters.AddWithValue("@codv", v_codidv);
                    cdv.Parameters.AddWithValue("@loca", v_clu);
                    using (MySqlDataAdapter datv = new MySqlDataAdapter(cdv))
                    {
                        dttd1.Clear();
                        datv.Fill(dttd1);
                        cmb_tdv.DataSource = dttd1;
                        cmb_tdv.DisplayMember = "descrizionerid";
                        cmb_tdv.ValueMember = "idcodice";
                    }
                }
                //  datos para los combobox de tipo de documento
                cmb_docRem.Items.Clear();
                using (MySqlCommand cdu = new MySqlCommand("select idcodice,descrizionerid,marca1 as codigo,codsunat from desc_doc where numero=@bloq", conn))
                {
                    cdu.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter datd = new MySqlDataAdapter(cdu))
                    {
                        dttd0.Clear();
                        datd.Fill(dttd0);
                        cmb_docRem.DataSource = dttd0;
                        cmb_docRem.DisplayMember = "descrizionerid";
                        cmb_docRem.ValueMember = "idcodice";
                    }
                }
                // datos para el combo de moneda
                cmb_mon.Items.Clear();
                using (MySqlCommand cmo = new MySqlCommand("select idcodice,descrizionerid,codsunat,deta1 from desc_mon where numero=@bloq", conn))
                {
                    cmo.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter dacu = new MySqlDataAdapter(cmo))
                    {
                        dtm.Clear();
                        dacu.Fill(dtm);
                        cmb_mon.DataSource = dtm;
                        cmb_mon.DisplayMember = "descrizionerid";
                        cmb_mon.ValueMember = "idcodice";
                    }
                }
                // datos del combo plazo de pago creditos
                using (MySqlCommand compla = new MySqlCommand("select idcodice,descrizionerid,codsunat,marca1 from desc_tpa where numero=@bloq", conn))
                {
                    compla.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter dapla = new MySqlDataAdapter(compla))
                    {
                        dtp.Clear();
                        dapla.Fill(dtp);
                        cmb_plazoc.DataSource = dtp;
                        cmb_plazoc.DisplayMember = "descrizionerid";
                        cmb_plazoc.ValueMember = "idcodice";
                    }
                }
                // DATOS DE PAGO - MPA
                using (MySqlCommand qmpa = new MySqlCommand("select idcodice,descrizionerid from desc_mpa where numero=@bloq", conn))
                {
                    qmpa.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter dapla = new MySqlDataAdapter(qmpa))
                    {
                        mpa.Clear();
                        dapla.Fill(mpa);
                        mpdef = mpa.Rows[0].ItemArray[0].ToString();    // primera fila siempre x defecto
                        cmb_tipop.DataSource = mpa;
                        cmb_tipop.DisplayMember = "descrizionerid";
                        cmb_tipop.ValueMember = "idcodice";
                    }
                }
                // datos de ruta para cargas únicas (detracción)
                using (MySqlCommand conrut = new MySqlCommand("select idcodice,descrizione,marca2 from desc_rut where numero=@bloq", conn))
                {
                    conrut.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter darut = new MySqlDataAdapter(conrut))
                    {
                        dtru.Clear();
                        darut.Fill(dtru);
                        cmb_ruta.DataSource = dtru;
                        cmb_ruta.DisplayMember = "descrizione";
                        cmb_ruta.ValueMember = "idcodice";
                    }
                }
                // jalamos la caja
                using (MySqlConnection cong = new MySqlConnection(db_conn_grael))
                {
                    cong.Open();
                    if (cong.State == ConnectionState.Open)
                    {
                        using (MySqlCommand micon = new MySqlCommand("SELECT a.id,a.fecha,a.status,b.flag1 from macajas a LEFT JOIN desc_sds b ON b.idcodice=a.local where LOCAL=@luc order by id desc LIMIT 1", cong))
                        {   // "select id,fechope,statusc from cabccaja where loccaja=@luc order by id desc limit 1"
                            micon.Parameters.AddWithValue("@luc", v_clu);
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    v_estcaj = dr.GetString("status");
                                    v_idcaj = dr.GetString("id");
                                    v_sercob = dr.GetString("flag1");
                                }
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("No hay conexión a la base de Grael 1.0", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }
            }
        }
        private bool valiVars()                 // valida existencia de datos en variables del form
        {
            bool retorna = true;
            if (vtc_dni == "")           // variable tipo cliente natural
            {
                lib.messagebox("Tipo de cliente Natural");
                retorna = false;
            }
            if (vtc_ruc == "")          // variable tipo cliente empresa
            {
                lib.messagebox("Tipo de cliente Empresa");
                retorna = false;
            }
            if (vtc_ext == "")          // variable tipo cliente extranjero
            {
                lib.messagebox("Tipo de cliente Extranjero");
                retorna = false;
            }
            if (codAnul == "")          // codigo de documento anulado
            {
                lib.messagebox("Código de Doc.Venta ANULADA");
                retorna = false;
            }
            if (codGene == "")          // codigo documento nuevo generado
            {
                lib.messagebox("Código de Doc.Venta GENERADA/NUEVA");
                retorna = false;
            }
            if (MonDeft == "")          // moneda por defecto
            {
                lib.messagebox("Moneda por defecto");
                retorna = false;
            }
            if (v_slu == "")            // serie del local del usuario
            {
                lib.messagebox("Serie general local del usuario");
                retorna = false;
            }
            if (vi_formato == "")       // formato de impresion del documento
            {
                lib.messagebox("formato de impresion del Doc.Venta");
                retorna = false;
            }
            if (vi_copias == "")        // cant copias impresion
            {
                lib.messagebox("# copias impresas del Doc.Venta");
                retorna = false;
            }
            if (v_impTK == "")           // nombre de la ticketera
            {
                lib.messagebox("Nombre de impresora de Tickets");
                retorna = false;
            }
            if (v_sanu == "")           // serie de anulacion del documento
            {
                lib.messagebox("Serie de Anulación interna");
                retorna = false;
            }
            if (v_CR_gr_ind == "")
            {
                lib.messagebox("Nombre formato Doc.Venta en CR");
                retorna = false;
            }
            if (v_mfildet == "")
            {
                lib.messagebox("Max. filas de detalle");
                retorna = false;
            }
            if (vint_A0 == "")
            {
                lib.messagebox("Código interno enlace anulación BD - A0");
                retorna = false;
            }
            // aca falta agregar resto  ...........
            return retorna;
        }
        private void inivarGR()                 // inicializa vars de la guia que se van ingresando
        {
            datcltsR[0] = "";
            datcltsR[1] = "";
            datcltsR[2] = "";
            datcltsR[3] = "";
            datcltsR[4] = "";
            datcltsR[5] = "";
            datcltsR[6] = "";
            datcltsR[7] = "";
            datcltsR[8] = "";
            datcltsR[9] = "";
            //
            datcltsD[0] = "";
            datcltsD[1] = "";
            datcltsD[2] = "";
            datcltsD[3] = "";
            datcltsD[4] = "";
            datcltsD[5] = "";
            datcltsD[6] = "";
            datcltsD[7] = "";
            datcltsD[8] = "";
            datcltsD[9] = "";
            //
            datguias[0] = "";   // num GR
            datguias[1] = "";   // descrip
            datguias[2] = "";   // cant bultos
            datguias[3] = "";   // nombre de la moneda de la GR
            datguias[4] = "";   // valor de la guía en su moneda
            datguias[5] = "";   // valor en moneda local
            datguias[6] = "";   // codigo moneda local
            datguias[7] = "";   // codigo moneda de la guia
            datguias[8] = "";   // tipo de cambio
            datguias[9] = "";   // fecha de la GR
            datguias[10] = "";  // guia del cliente, sustento del cliente
            datguias[11] = "";   // placa
            datguias[12] = "";   // carreta
            datguias[13] = "";   // autoriz circulacion
            datguias[14] = "";   // conf. vehicular
            datguias[15] = "";  // local origen-destino
            datguias[16] = "";  // saldo de la GR
            datguias[17] = "";  // unidad de medida 
            datguias[18] = "";  // direccion local grael partida 
            datguias[19] = "";  // ubigeo local grael partida

            datcargu[0] = "";   // direc. partida
            datcargu[1] = "";   // depart. pto. partida
            datcargu[2] = "";   // provin. pto. partida
            datcargu[3] = "";   // distri. pto. partida
            datcargu[4] = "";   // ubigeo punto partida
            datcargu[5] = "";   // direc. llegada
            datcargu[6] = "";   // depart. pto. llegada
            datcargu[7] = "";   // provin. pto. llegada
            datcargu[8] = "";   // distri. pto. llegada
            datcargu[9] = "";   // ubigeo punto llegada
            datcargu[10] = "";  // ruc del camion
            datcargu[11] = "";  // razon social del ruc
            datcargu[12] = "";  // fecha inicio del traslado
        }   
        private bool validGR(string serie, string corre)    // validamos y devolvemos datos
        {
            bool retorna = false;
            if (serie != "" && corre != "")
            {
                // validamos que la GR: 1.exista, 2.No este facturada, 3.No este anulada
                // y devolvemos una fila con los datos del remitente y otra fila los datos del destinatario
                string hay = "no";
                using (MySqlConnection conn = new MySqlConnection(db_conn_grael))
                {
                    lib.procConn(conn);
                    //string cons = "select fecguitra,totguitra,estadoser,fecdocvta,tipdocvta,serdocvta,numdocvta,codmonvta,totdocvta,saldofina " +
                    //    "from controlg where serguitra=@ser and numguitra=@num";
                    string cons = "select fecgr,totnot,status,fecdv,tipdv,serdv,cordv,mondv,totdv,saldo " +
                        "from mactacte where sergr=@ser and corgr=@num";
                    using (MySqlCommand mic1 = new MySqlCommand(cons, conn))
                    {
                        mic1.Parameters.AddWithValue("@ser", serie);
                        mic1.Parameters.AddWithValue("@num", corre);
                        using (MySqlDataReader dr = mic1.ExecuteReader())
                        {
                            if (dr.HasRows)
                            {
                                if (dr.Read())
                                {
                                    if (dr.GetString("cordv").Trim() != "") hay = "sif"; // si hay guía pero ya esta facturado
                                    else hay = "sin";    // si hay guía y no tiene factura
                                    if (dr.GetString("saldo") != dr.GetString("totnot") && dr.GetDecimal("saldo") > 0)
                                    {
                                        MessageBox.Show("No esta permitido generar un documento" + Environment.NewLine + 
                                            "de venta de una guía que tiene pago parcial","Atención - no puede continuar");
                                        hay = "no";
                                    }
                                }
                            }
                            else
                            {
                                hay = "no"; // no existe la guía
                            }
                        }
                    }
                    if (hay == "sin")
                    {
                        string consulta = "SELECT a.docrem,a.numdre,b1.nombre as nombregri,b1.direc as direregri,b1.ubigeo as ubigregri,ifnull(b1.email,'') as emailR,ifnull(b1.telef1,'') as numtel1R," +
                            "ifnull(b1.telef2, '') as numtel2R,a.docdes,a.numdes,b2.nombre as nombdegri,b2.direc as diredegri,b2.ubigeo as ubigdegri,ifnull(b2.email, '') as emailD," +
                            "ifnull(b2.telef1, '') as numtel1D,ifnull(b2.telef2, '') as numtel2D,a.moneda,a.doctot,a.saldo as saldoG,SUM(d.cantid) AS bultos, date(a.fechope) as fechopegr,a.tipcam," +
                            "max(d.descrip) AS descrip,max(d.unidad) as unidad,ifnull(m.descrizionerid, '') as mon,a.doctot as totgrMN,a.moneda as codMN,c.fecdv,' ' as tipsrem,' ' as tipsdes,a.docremi," +
                            "a.placa,a.carreta,a.cerinsc,a.nfv,concat(lo.descrizionerid, ' - ', ld.descrizionerid) as orides,c.saldo,a.dirorig1 as dirpartida," +
                            "' ' as ubigpartida,a.dirdest1 as dirllegada,' ' as ubigllegada,ifnull(c.fecma, '') as fechplani,a.ruc,ifnull(p.nombre, '') as RazonSocial,dr.flag1 as dr,dd.flag1 as dd," +
                            "a.origen AS codOriG,lo.ubigeo AS ubiOriG,CONCAT(lo.deta1,' ',lo.deta2,' ',lo.deta3,' ',lo.deta4) AS diroriG " +
                            "from magrem a left join detagrem d on d.idc = a.id " +
                            "LEFT JOIN mactacte c ON c.sergr = a.sergre AND c.corgr = a.corgre " +
                            "left join anag_for p on p.ruc = a.ruc " +
                            "left join anag_cli b1 on b1.docu = a.docrem and b1.ruc = a.numdre " +
                            "left join anag_cli b2 on b2.docu = a.docdes and b2.ruc = a.numdes " +
                            "left join desc_mon m on m.idcodice = a.moneda " +
                            "left join desc_sds lo on lo.idcodice = a.origen " +
                            "left join desc_sds ld on ld.idcodice = a.destino " +
                            "left join desc_doc dr on dr.idcodice = a.docrem " +
                            "left join desc_doc dd on dd.idcodice = a.docdes " +
                            "WHERE a.sergre = @ser AND a.corgre = @num AND a.status not IN(@est) AND c.fecdv IS NULL";
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.Parameters.AddWithValue("@ser", serie);
                            micon.Parameters.AddWithValue("@num", corre);
                            micon.Parameters.AddWithValue("@est", codAnul);
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    if (!dr.IsDBNull(0))    //  && dr[24] == DBNull.Value
                                    {
                                        if (int.Parse(tx_tfil.Text) > 0)
                                        {
                                            if ((tx_dat_tdRem.Text + tx_numDocRem != dr.GetString("docrem") + dr.GetString("numdre")) && (tx_dat_tdRem.Text + tx_numDocRem != dr.GetString("docdes") + dr.GetString("numdes")))
                                            {
                                                var aaa = MessageBox.Show("El remitente y destinatario de la GR" + Environment.NewLine +
                                                    "no coinciden con el cliente del comprobante" + Environment.NewLine +
                                                    "Desea continuar?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                                if (aaa == DialogResult.No)
                                                {
                                                    //inivarGR();
                                                    tx_numGR.Text = "";
                                                    tx_numGR.Focus();
                                                    return retorna;
                                                }
                                            }
                                        }
                                        inivarGR();
                                        datcltsR[0] = dr.GetString("docrem");        // datos del remitente de la GR
                                        datcltsR[1] = dr.GetString("numdre");
                                        datcltsR[2] = dr.GetString("nombregri");
                                        datcltsR[3] = dr.GetString("direregri");
                                        datcltsR[4] = dr.GetString("ubigregri");
                                        datcltsR[5] = dr.GetString("emailR");
                                        datcltsR[6] = dr.GetString("numtel1R");
                                        datcltsR[7] = dr.GetString("numtel2R");
                                        datcltsR[8] = dr.GetString("tipsrem");
                                        datcltsR[9] = dr.GetString("dr");
                                        //
                                        datcltsD[0] = dr.GetString("docdes");        // datos del destinatario de la GR
                                        datcltsD[1] = dr.GetString("numdes");
                                        datcltsD[2] = dr.GetString("nombdegri");
                                        datcltsD[3] = dr.GetString("diredegri");
                                        datcltsD[4] = dr.GetString("ubigdegri");
                                        datcltsD[5] = dr.GetString("emailD");
                                        datcltsD[6] = dr.GetString("numtel1D");
                                        datcltsD[7] = dr.GetString("numtel2D");
                                        datcltsD[8] = dr.GetString("tipsdes");
                                        datcltsD[9] = dr.GetString("dd");
                                        //
                                        datguias[0] = serie + "-" + corre;                 // GR
                                        datguias[1] = (dr.IsDBNull(20)) ? "" : dr.GetString("descrip");         // descrip
                                        datguias[2] = (dr.IsDBNull(19)) ? "0" : dr.GetString("bultos");          // cant bultos
                                        datguias[3] = dr.GetString("mon");             // nombre moneda de la GR
                                        datguias[4] = dr.GetString("doctot");          // valor GR en su moneda
                                        datguias[5] = dr.GetString("totgrMN");         // valor GR en moneda local
                                        datguias[6] = dr.GetString("codMN");            // codigo moneda local
                                        datguias[7] = dr.GetString("moneda");           // codigo moneda de la guía
                                        datguias[8] = dr.GetString("tipcam");           // tipo de cambio de la GR
                                        datguias[17] = dr.GetString("unidad");           // unidad de medida
                                        var a = dr.GetString("fechopegr").Substring(0, 10);
                                        datguias[9] = a.Substring(6,4) + "-" + a.Substring(3,2) + "-" + a.Substring(0,2);     // fecha de la GR
                                        datguias[10] = dr.GetString("docremi");
                                        datguias[11] = dr.GetString("placa"); 
                                        datguias[12] = dr.GetString("carreta");
                                        datguias[13] = dr.GetString("cerinsc");
                                        datguias[14] = dr.GetString("nfv");
                                        datguias[15] = dr.GetString("orides");
                                        datguias[16] = dr.GetString("saldo");
                                        datguias[18] = dr.GetString("diroriG");     // direc local partida grael
                                        datguias[19] = dr.GetString("ubiOriG");     // ubigeo local partida grael 
                                        //
                                        datcargu[0] = dr.GetString("dirpartida");
                                        datcargu[4] = dr.GetString("ubigpartida");   // ubigeo punto partida
                                        string[] aa = lib.retDPDubigeo(datcargu[4]);
                                        datcargu[1] = aa[0];   // depart. pto. partida
                                        datcargu[2] = aa[1];   // provin. pto. partida
                                        datcargu[3] = aa[2];   // distri. pto. partida
                                        datcargu[5] = dr.GetString("dirllegada");   // direc. llegada
                                        datcargu[9] = dr.GetString("ubigllegada");   // ubigeo punto llegada
                                        aa = lib.retDPDubigeo(datcargu[9]);
                                        datcargu[6] = aa[0];   // depart. pto. llegada
                                        datcargu[7] = aa[1];   // provin. pto. llegada
                                        datcargu[8] = aa[2];   // distri. pto. llegada
                                        datcargu[10] = dr.GetString("ruc");  // ruc del camion
                                        datcargu[11] = dr.GetString("RazonSocial");  // razon social del ruc
                                        datcargu[12] = dr.GetString("fechplani");    // fecha inicio traslado
                                        //
                                        tx_dat_saldoGR.Text = dr.GetString("saldo");
                                        // pasamos los datos al panel de cargas única si: es factura especial
                                        if (rb_fesp.Checked == true)
                                        {
                                            datPanelCargasU();   
                                        }
                                        retorna = true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return retorna;
        }
        private void datPanelCargasU()          // carga datos panel cargas unicas
        {
            tx_e_placa.Text = datguias[11];
            tx_e_aut.Text = datguias[13];
            tx_e_ruct.Text = datcargu[10];
            tx_e_ntrans.Text = datcargu[11];
            tx_e_ftras.Text = datcargu[12];
            tx_e_nfv.Text = datguias[14];
            //tx_e_dnicho.Text = 
            tx_e_dirpar.Text = datcargu[0] + datcargu[1] + " - " + datcargu[2] + " - " + datcargu[3];
            tx_e_dirlle.Text = datcargu[5] + datcargu[6] + " - " + datcargu[7] + " - " + datcargu[8];
            tx_e_ubiori.Text = datcargu[4];
            tx_e_ubides.Text = datcargu[9];
        }
        private void tipcambio(string codmod)               // funcion para calculos con el tipo de cambio
        {
            decimal totflet = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    totflet = totflet + decimal.Parse(dataGridView1.Rows[i].Cells[5].Value.ToString()); // VALOR DE LA GR EN MONEDA LOCAL
                }
            }
            // si codmod es moneda local, suma campos totales de moneda local y retorna valor
            if (codmod == MonDeft)
            {
                tx_flete.Text = totflet.ToString("#0.00");
            }
            else
            {
                if (codmod != "")
                {
                    vtipcam vtipcam = new vtipcam(tx_tfmn.Text, codmod, DateTime.Now.Date.ToString());
                    var result = vtipcam.ShowDialog();
                    tx_flete.Text = vtipcam.ReturnValue1;
                    tx_fletMN.Text = vtipcam.ReturnValue2;
                    tx_tipcam.Text = vtipcam.ReturnValue3;
                    tx_flete_Leave(null, null);
                }
            }
        }
        private void calculos(decimal totDoc)
        {
            decimal tigv = 0;
            decimal tsub = 0;
            if (totDoc > 0)
            {
                tsub = Math.Round(totDoc / (1 + decimal.Parse(v_igv) / 100), 2);
                tigv = Math.Round(totDoc - tsub, 2);
                
            }
            tx_igv.Text = tigv.ToString("#0.00");
            tx_subt.Text = tsub.ToString("#0.00");
        }
        private void totalizaG()                            // calcula totales de la grilla
        {
            int totfil = 0;
            int totcant = 0;
            decimal totflet = 0;    // acumulador en moneda de la GR 
            decimal totflMN = 0;
            decimal totsal = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1.Rows[i].Cells[0].Value != null)
                {
                    totcant = totcant + int.Parse(dataGridView1.Rows[i].Cells[2].Value.ToString());
                    totfil += 1;
                    if (tx_dat_mone.Text != MonDeft)
                    {
                        totflet = totflet + decimal.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString()); // VALOR de la GR
                        totflMN = totflMN + decimal.Parse(dataGridView1.Rows[i].Cells[5].Value.ToString()); // VALOR DE LA GR EN MONEDA LOCAL
                        totsal = totsal + decimal.Parse(dataGridView1.Rows[i].Cells[11].Value.ToString());
                    }
                    else
                    {
                        totflet = totflet + decimal.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString()); // VALOR DE LA GR EN SU MONEDA
                        totflMN = totflMN + decimal.Parse(dataGridView1.Rows[i].Cells[5].Value.ToString()); // VALOR DE LA GR EN MONEDA LOCAL
                        totsal = totsal + decimal.Parse(dataGridView1.Rows[i].Cells[11].Value.ToString());
                    }
                }
            }
            tx_tfmn.Text = totflMN.ToString("#0.00");
            tx_totcant.Text = totcant.ToString();
            tx_tfil.Text = totfil.ToString();
            tx_flete.Text = totflet.ToString("#0.00");
            tx_fletMN.Text = totflMN.ToString("#0.00"); // Math.Round(decimal.Parse(tx_flete.Text) * decimal.Parse(tx_tipcam.Text), 2).ToString();
            tx_saldoT.Text = totsal.ToString("#0.00");
        }
        int CentimeterToPixel(double Centimeter)
        {
            double pixel = -1;
            using (Graphics g = this.CreateGraphics())
            {
                pixel = Centimeter * g.DpiY / 2.54d;
            }
            return (int)pixel;
        }
        private void validaclt()
        {
            if (tx_numDocRem.Text.Trim().Length != Int16.Parse(tx_mld.Text))
            {
                MessageBox.Show("El número de caracteres para" + Environment.NewLine +
                    "su tipo de documento debe ser: " + tx_mld.Text, "Error de validación", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                tx_numDocRem.Focus();
                return;
            }
            if (tx_dat_tdRem.Text == vtc_ruc && lib.valiruc(tx_numDocRem.Text, vtc_ruc) == false)
            {
                MessageBox.Show("Número de RUC inválido", "Atención - revise", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tx_numDocRem.Focus();
                return;
            }
            string encuentra = "no";
            if (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDITAR")
            {
                string[] datos = lib.datossn("CLI", tx_dat_tdRem.Text.Trim(), tx_numDocRem.Text.Trim());
                if (datos[0] != "")  // datos.Length > 0
                {
                    tx_nomRem.Text = datos[0];
                    tx_nomRem.Select(0, 0);
                    tx_dirRem.Text = datos[1];
                    tx_dirRem.Select(0, 0);
                    tx_dptoRtt.Text = datos[2];
                    tx_dptoRtt.Select(0, 0);
                    tx_provRtt.Text = datos[3];
                    tx_provRtt.Select(0, 0);
                    tx_distRtt.Text = datos[4];
                    tx_distRtt.Select(0, 0);
                    tx_ubigRtt.Text = datos[5];
                    tx_ubigRtt.Select(0, 0);
                    tx_email.Text = datos[7];
                    tx_email.Select(0, 0);
                    tx_telc1.Text = datos[6];
                    tx_telc1.Select(0, 0);
                    encuentra = "si";
                    tx_dat_m1clte.Text = "E";
                }
                if (tx_dat_tdRem.Text == vtc_ruc)
                {
                    if (encuentra == "no")
                    {
                        if (Grael2.Program.vg_conSol == true) // conector solorsoft para ruc
                        {
                            string[] rl = lib.conectorSolorsoft("RUC", tx_numDocRem.Text);
                            tx_nomRem.Text = rl[0];      // razon social
                            tx_ubigRtt.Text = rl[1];     // ubigeo
                            tx_dirRem.Text = rl[2];      // direccion
                            tx_dptoRtt.Text = rl[3];      // departamento
                            tx_provRtt.Text = rl[4];      // provincia
                            tx_distRtt.Text = rl[5];      // distrito
                            tx_dat_m1clte.Text = "N";
                        }
                    }
                }
                if (tx_dat_tdRem.Text == vtc_dni)
                {
                    if (encuentra == "no")
                    {
                        if (Grael2.Program.vg_conSol == true) // conector solorsoft para dni
                        {
                            string[] rl = lib.conectorSolorsoft("DNI", tx_numDocRem.Text);
                            tx_nomRem.Text = rl[0];      // nombre
                                                         //tx_numDocRem.Text = rl[1];     // num dni
                            tx_dat_m1clte.Text = "N";
                        }
                        if (rb_remGR.Checked == true)
                        {
                            tx_dirRem.Text = datcltsR[3].ToString();    // datcltsR[3] = dr.GetString("direregri");
                            tx_ubigRtt.Text = datcltsR[4].ToString();    // datcltsR[4] = dr.GetString("ubigregri");
                        }
                        if (rb_desGR.Checked == true)
                        {
                            tx_dirRem.Text = datcltsD[3].ToString();
                            tx_ubigRtt.Text = datcltsD[4].ToString();
                        }
                        string[] aa = lib.retDPDubigeo(tx_ubigRtt.Text);
                        tx_dptoRtt.Text = aa[0].ToString();
                        tx_provRtt.Text = aa[1].ToString();
                        tx_distRtt.Text = aa[2].ToString();
                    }
                }
            }
        }

        #region facturacion electronica
        private void OSE_Fact_Elect()                       // string de conexion a la base de datos mssql de bizlinks
        {
            MySqlConnection conn = new MySqlConnection(db_conn_grael);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                string consulta = "select param,value from confmod where used='R'";
                MySqlCommand micon = new MySqlCommand(consulta, conn);
                MySqlDataAdapter da = new MySqlDataAdapter(micon);
                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["param"].ToString() == "serv") script = "Data Source=" + row["value"].ToString() + ";";              // "ServerName;"
                        //if(row["param"].ToString() == "port") script = script + "Port=" + row["value"].ToString() + ";";            // puerto
                        if (row["param"].ToString() == "data") script = script + "Initial Catalog=" + row["value"].ToString() + ";";  // database name
                        if (row["param"].ToString() == "user") script = script + "User id=" + row["value"].ToString() + ";";          // usuario
                        if (row["param"].ToString() == "pass") script = script + "Password=" + row["value"].ToString() + ";";        // password
                    }
                }
                else
                {
                    da.Dispose();
                    conn.Close();
                    MessageBox.Show("No se encuentra configuración para acceso" + Environment.NewLine +
                        "a base de datos del OSE de Fact. Elect.", "Error de configuración", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    Application.Exit();
                    return;
                }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
                return;
            }
            conn.Close();
        }
        private bool valiconsql(string strcon)              // valida conexion al servidor sql de bizlinks
        {
            bool retorna = false;
            using (SqlConnection conms = new SqlConnection(strcon))
            {
                conms.Open();
                if (conms.State == ConnectionState.Open)
                {
                    retorna = true;
                }
            }
            return retorna;
        }
        private void grabfactelec()                         // graba en la tabla de fact. electrónicas
        {                                                   // OSE BIZLINKS 10-10-2018  / actualizacion 14/08/2021 / actualización 23/03/2022
            string tipo = tx_dat_tdv.Text;
            string serie = tx_serie.Text;
            string corre = "0" + tx_numero.Text;
            string sernum = cmb_tdv.Text.Substring(0, 1) + tx_serie.Text + "-" + corre;              // v serieNumero 
            string fecemi = tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) +
                "-" + tx_fechope.Text.Substring(0, 2);                                                 // v 
            DataRow[] row = dttd1.Select("idcodice='" + tx_dat_tdv.Text + "'");                     // tipo de documento venta
            string tipdoc = row[0][3].ToString();                                                 // v tipoDocumento
            glosser = row[0][4].ToString();
            DataRow[] rowm = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");                     // tipo de moneda
            string tipmon = rowm[0][2].ToString().Trim();                                           // v tipoMoneda
            string nudoem = Program.ruc;                                                            // v numeroDocumentoEmisor
            string tidoem = "6";                                                                    // v tipoDocumentoEmisor
            string nocoem = "-";                                                                    // v nombreComercialEmisor
            string rasoem = Program.cliente.Trim();                                                 // v razonSocialEmisor
            string coremi = Program.mailclte.Trim();                                               // v correoEmisor
            string coloem = Program.codlocsunat;                                                   // v codigoLocalAnexoEmisor
            string ubiemi = Program.ubidirfis;                                                      // v ubigeoEmisor
            string diremi = Program.dirfisc;                                                       // v direccionEmisor
            string provemi = Program.provfis;                                                     // v provinciaEmisor
            string depaemi = Program.depfisc;                                                     // v departamentoEmisor
            string distemi = Program.distfis;                                                     // v distritoEmisor
            string urbemis = "-";                                                                   // urbanizacion (parte de la direccion) del emisor
            string pasiemi = "PE";                                                                  // v paisEmisor
            DataRow[] rowc = dttd0.Select("idcodice='" + tx_dat_tdRem.Text + "'");
            string tidoad = rowc[0][3].ToString().Trim();                                           // v tipoDocumentoAdquiriente
            string nudoad = tx_numDocRem.Text;                                                      // v numeroDocumentoAdquiriente
            string rasoad = tx_nomRem.Text;                                                          // v razonSocialAdquiriente
            string coradq = tx_email.Text;                                                         // v correoAdquiriente
            decimal totimp = Math.Round(decimal.Parse(tx_igv.Text), 2);                              // v totalImpuestos
            decimal totigv = Math.Round(decimal.Parse(tx_igv.Text), 2);                              // v totalIgv
            decimal totvta = Math.Round(decimal.Parse(tx_flete.Text), 2);                            // v totalVenta
            decimal tovane = Math.Round(decimal.Parse(tx_subt.Text), 2);                            // v totalValorVentaNetoOpGravadas
            // todas las guias, tanto del transportista como de el cliente
            string gpgrael = "";
            string gpadqui = dataGridView1.Rows[0].Cells[3].Value.ToString();                         // guias del adquiriente
            for (int j = 1; j < dataGridView1.Rows.Count - 1; j++)
            {
                gpadqui = gpadqui + " - " + dataGridView1.Rows[j].Cells[3].Value.ToString();
            }
            string codaux40_1 = "9011";                                                             // v codigoAuxiliar40_1
            string etiaux40_1 = "18%";                                                              // v textoAuxiliar40_1
            string tipope = "0101"; // segun rudver, poner esto en una config                       // v tipoOperacion
            string estreg = "A";                                                                    // bl_estadoRegistro
            string coley1 = "1000";                                                                 // v codigoLeyenda_1
            string teley1 = "SON: " + tx_fletLetras.Text; // nl.Convertir(tx_total.Text, true) + tx_dat_dmon.Text;         // v textoLeyenda_1
            string tiref1 = "";     // detalle
            string nudor1 = "";     // detalle
            string tiref2 = "";     // detalle
            string nudor2 = "";     // detalle
            string tiref3 = "";     // detalle
            string nudor3 = "";     // detalle
            string Ctiref1 = "";     // cabecera
            string Cnudor1 = "";     // cabecera
            string Ctiref2 = "";     // cabecera
            string Cnudor2 = "";     // cabecera
            string Ctiref3 = "";     // cabecera
            string Cnudor3 = "";     // cabecera
            double totdet = 0;
            string leydet = leydet1 + " " + leydet2 + " " + Program.ctadetra;                   // textoLeyenda_2
            string cauxdet1 = "6665";   // codigo para adicional observacion detracción en caso de fact de cargas unicas
            //string tauxdet1 = "OPERACION SUJETA A SPOT BANCO DE LA NACION CTA.CTE.MN " + Program.ctadetra +
            //    " VALOR REFERENCIAL S/ " + tx_e_valref.Text + " DETRACCION S/ " + tx_detrac.Text;     // texto de la observacion de la detraccion en fact especiales de cargas unicas
            string tauxdet1 = leydet1 + " " + leydet2 + Program.ctadetra +
                " VALOR REFERENCIAL S/ " + tx_e_valref.Text + " DETRACCION S/ " + tx_detrac.Text;
            if (tx_dat_mone.Text != MonDeft) tauxdet1 = tauxdet1 + " AL T.C. " + tx_tipcam.Text;
            SqlConnection conms = new SqlConnection(script);
            conms.Open();
            if (conms.State == ConnectionState.Open)
            {
                string insertadet = "";                         // segmento detalle
                try
                {
                    string ubiPtoOri = "";                                                    // ubigeoPtoOrigen
                    string dirPtoOri = "";                                                    // direccionCompletaPtoOrigen
                    string ubiPtoDes = "";                                                    // ubigeoPtoDestino
                    string dirPtoDes = "";                                                    // direccionCompletaPtoDestino
                    string detViaje = "";                                                  // detalleViaje
                    string monRefSer = "";                                                    // montoRefServicioTransporte
                    string monRefCar = "";                                                    // montoRefCargaEfectiva
                    string monRefUti = "";                                                     // montoRefCargaUtilNominal

                    for (int q = 0; q < dataGridView1.Rows.Count - 1; q++)
                    {
                        glosser2 = dataGridView1.Rows[q].Cells[10].Value.ToString();
                        string nuori1 = (q + 1).ToString();                                                                     // numeroOrdenItem
                        string codprd1 = "-";                                                                                   // codigoProducto
                        string coprsu1 = "78101802";                                                                            // codigoProductoSunat
                        string descr1 = "";
                        decimal canti1 = Math.Round(decimal.Parse("1"), 2);
                        string unime1 = "ZZ";                                                                                   // unidadMedida
                        decimal psi1, igv1;                                                                                     // calculos de precios x item sin y con impuestos
                        double inuns1 = 0;                                                                                      // importeUnitarioSinImpuesto
                        decimal inunc1 = Math.Round(decimal.Parse(dataGridView1.Rows[q].Cells["valor"].Value.ToString()), 2);   // importeUnitarioConImpuesto
                        string coimu1 = "01";                                                                   // codigoImporteUnitarioConImpues
                        string imtoi1 = "";

                        if (rb_fesp.Checked == true)
                        {
                            unime1 = "TNE";                                                                                     // unidadMedida
                            descr1 = tx_e_glos1.Text.Trim() + " " + tx_e_glos2.Text.Trim() + " " + tx_e_glos3.Text.Trim();      // descripcion cargas únicas 
                            descr1 = descr1.Replace(System.Environment.NewLine, " ");
                            canti1 = Math.Round(decimal.Parse(tx_e_cant.Text), 2);
                            //unime1 = "TN";
                            inuns1 = Math.Round(double.Parse(tx_e_prec.Text), 6);                                               // importeUnitarioSinImpuesto
                            inunc1 = Math.Round(decimal.Parse(tx_e_prec.Text) * (1 + (decimal.Parse(v_igv)/100)), 6);           // importeUnitarioConImpuesto
                            imtoi1 = Math.Round(inunc1 - Convert.ToDecimal(inuns1), 2).ToString();
                             ubiPtoOri = tx_e_ubiori.Text;                                                    // ubigeoPtoOrigen
                             dirPtoOri = tx_e_dirpar.Text;                                                    // direccionCompletaPtoOrigen
                             ubiPtoDes = tx_e_ubides.Text;                                                    // ubigeoPtoDestino
                             dirPtoDes = tx_e_dirlle.Text;                                                    // direccionCompletaPtoDestino
                             detViaje = "Detalle del viaje";                                                  // detalleViaje
                             monRefSer = tx_e_valref.Text;                                                    // montoRefServicioTransporte
                             monRefCar = tx_e_valref.Text;                                                    // montoRefCargaEfectiva
                             monRefUti = tx_e_carut.Text;                                                     // montoRefCargaUtilNominal
                        }
                        else
                        {
                            descr1 = glosser + " " + glosser2 + " " +
                                vint_gg + " " + dataGridView1.Rows[q].Cells["Descrip"].Value.ToString();                        // descripcion estandar
                            descr1 = descr1.Replace(System.Environment.NewLine, " ");
                            if (decimal.TryParse(dataGridView1.Rows[q].Cells["valor"].Value.ToString(), out psi1))
                            {
                                inuns1 = Math.Round(((double)psi1 / ((double)decimal.Parse(v_igv) / 100 + 1)), 2);                  // importeUnitarioSinImpuesto
                            }
                            else { inuns1 = 0; }                                                                                    // importeUnitarioSinImpuesto
                            if (decimal.TryParse(dataGridView1.Rows[q].Cells["valor"].Value.ToString(), out igv1))
                            {
                                imtoi1 = Math.Round(((double)igv1 - ((double)igv1 / ((double)decimal.Parse(v_igv) / 100 + 1))), 2).ToString();
                            }
                            else { imtoi1 = "0.00"; }                                                               // importeTotalImpuestos 
                            if (rb_fnor.Checked == true && double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                            {
                                ubiPtoOri = datguias[19];                                        // ubigeoPtoOrigen
                                dirPtoOri = datguias[18];                                        // direccionCompletaPtoOrigen
                                ubiPtoDes = tx_ubigRtt.Text;                                     // ubigeoPtoDestino
                                dirPtoDes = tx_dirRem.Text;                                      // direccionCompletaPtoDestino
                                detViaje = "Detalle del viaje";                                  // detalleViaje
                                monRefSer = "0.00";                                                    // montoRefServicioTransporte
                                monRefCar = "0.00";                                                    // montoRefCargaEfectiva
                                monRefUti = "0.00";                                                     // montoRefCargaUtilNominal
                            }
                        }
                        double mobai1 = Math.Round(inuns1,2);                                                     // montoBaseIgv
                        string taigv1 = ((decimal.Parse(v_igv))).ToString();                                    // tasaIgv
                        string imigv1 = imtoi1;                                                                 // importeIgv
                        string corae1 = "10";   // grabado operacion onerosa                                    // codigoRazonExoneracion
                        double intos1 = Math.Round(inuns1,2);                                                   // importeTotalSinImpuesto

                        insertadet = "insert into SPE_EINVOICEDETAIL (tipoDocumentoEmisor,numeroDocumentoEmisor,tipoDocumento,serieNumero," +
                            "numeroOrdenItem,codigoProducto,codigoProductoSunat,descripcion,cantidad,unidadMedida,importeTotalSinImpuesto," +
                            "importeUnitarioSinImpuesto,importeUnitarioConImpuesto,codigoImporteUnitarioConImpues,montoBaseIgv,tasaIgv," +
                            "importeIgv,importeTotalImpuestos,codigoRazonExoneracion,codigoAuxiliar40_1,textoAuxiliar40_1,codigoAuxiliar500_1,textoAuxiliar500_1";
                        if (!string.IsNullOrEmpty(tiref2) && !string.IsNullOrWhiteSpace(tiref2)) insertadet = insertadet + ",codigoAuxiliar500_2,textoAuxiliar500_2";
                        if (!string.IsNullOrEmpty(tiref3) && !string.IsNullOrWhiteSpace(tiref3)) insertadet = insertadet + ",codigoAuxiliar500_3,textoAuxiliar500_3";
                        if (rb_fesp.Checked == true ||
                            (rb_fnor.Checked == true && double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact))
                        {
                            insertadet = insertadet + ",ubigeoPtoOrigen,direccionCompletaPtoOrigen,ubigeoPtoDestino,direccionCompletaPtoDestino,detalleViaje," +
                                "montoRefServicioTransporte,montoRefCargaEfectiva,montoRefCargaUtilNominal";
                        }
                        insertadet = insertadet + ") ";
                        insertadet = insertadet + "values (@tidoem,@nudoem,@tipdoc,@sernum," +
                            "@nuori1,@codprd1,@coprsu1,@descr1,@canti1,@unime1,@intos1," +
                            "@inuns1,@inunc1,@coimu1,@mobai1,@taigv1," +
                            "@imigv1,@imtoi1,@corae1,@tiref1,@nudor1,@coaux1,@teaux1";
                        if (!string.IsNullOrEmpty(tiref2) && !string.IsNullOrWhiteSpace(tiref2)) insertadet = insertadet + ",@coaux2,@teaux2";
                        if (!string.IsNullOrEmpty(tiref2) && !string.IsNullOrWhiteSpace(tiref2)) insertadet = insertadet + ",@coaux3,@teaux3";
                        if (rb_fesp.Checked == true ||
                            (rb_fnor.Checked == true && double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact))
                        {
                            insertadet = insertadet + ",@ubiPtoOri,@dirPtoOri,@ubiPtoDes,@dirPtoDes,@detViaje," +
                            "@monRefSer,@monRefCar,@monRefUti";
                        }
                        insertadet = insertadet + ") ";
                        {
                            SqlCommand indet = new SqlCommand(insertadet, conms);
                            indet.Parameters.AddWithValue("@tidoem", tidoem);       // tipo documento
                            indet.Parameters.AddWithValue("@nudoem", nudoem);       // numero documento
                            indet.Parameters.AddWithValue("@tipdoc", tipdoc);       // tipo documento
                            indet.Parameters.AddWithValue("@sernum", sernum);       // serieNumero
                            indet.Parameters.AddWithValue("@nuori1", nuori1);       // numero orden
                            indet.Parameters.AddWithValue("@codprd1", codprd1);     // codigo producto
                            indet.Parameters.AddWithValue("@coprsu1", coprsu1);     // codigo producto SUNAT
                            indet.Parameters.AddWithValue("@descr1", descr1);       // descripcion 
                            indet.Parameters.AddWithValue("@canti1", canti1.ToString("###.00"));       // cantidad
                            indet.Parameters.AddWithValue("@unime1", unime1);                           // unidad medida
                            indet.Parameters.AddWithValue("@intos1", intos1);                           // ImporteTotalSinImpuesto
                            indet.Parameters.AddWithValue("@inuns1", inuns1);                           // ImporteUnitarioSinImpuesto
                            indet.Parameters.AddWithValue("@inunc1", inunc1);                       // ImporteUnitarioConImpuesto
                            indet.Parameters.AddWithValue("@coimu1", coimu1);                       // codifoImporteUnitarioConImpuesto
                            indet.Parameters.AddWithValue("@mobai1", mobai1);                       // montoBaseIgv
                            indet.Parameters.AddWithValue("@taigv1", taigv1);                       // tasaIgv
                            indet.Parameters.AddWithValue("@imigv1", imigv1);                       // importeIgv
                            indet.Parameters.AddWithValue("@imtoi1", imtoi1);                       // importeTotalImpuestos
                            indet.Parameters.AddWithValue("@corae1", corae1);                       // codigoRazonExo
                            indet.Parameters.AddWithValue("@coaux1", "9840");  // tiref1                       // codigoAuxiliar500_1
                            indet.Parameters.AddWithValue("@teaux1", dataGridView1.Rows[q].Cells[8].Value.ToString().Trim());  //nudor1                       // textoAuxiliar500_1
                            indet.Parameters.AddWithValue("@tiref1", "8054"); // Ctiref1                     // codigoAuxiliar40_1
                            indet.Parameters.AddWithValue("@nudor1", dataGridView1.Rows[q].Cells[0].Value.ToString()); //Cnudor1                      // textoAuxiliar40_1
                            if (!string.IsNullOrEmpty(tiref2) && !string.IsNullOrWhiteSpace(tiref2))
                            {
                                indet.Parameters.AddWithValue("@coaux2", tiref2);                       // codigoAuxiliar500_2
                                indet.Parameters.AddWithValue("@teaux2", nudor2);                       // textoAuxiliar500_2
                            }
                            if (!string.IsNullOrEmpty(tiref3) && !string.IsNullOrWhiteSpace(tiref3))
                            {
                                indet.Parameters.AddWithValue("@coaux3", tiref3);                       // codigoAuxiliar500_3
                                indet.Parameters.AddWithValue("@teaux3", nudor3);                       // textoAuxiliar500_3
                            }
                            if (rb_fesp.Checked == true)
                            {
                                indet.Parameters.AddWithValue("@ubiPtoOri", ubiPtoOri);                   // ubigeoPtoOrigen
                                indet.Parameters.AddWithValue("@dirPtoOri", dirPtoOri);                   // direccionCompletaPtoOrigen
                                indet.Parameters.AddWithValue("@ubiPtoDes", ubiPtoDes);                   // ubigeoPtoDestino
                                indet.Parameters.AddWithValue("@dirPtoDes", dirPtoDes);                   // direccionCompletaPtoDestino
                                indet.Parameters.AddWithValue("@detViaje", detViaje);                     // detalleViaje
                                indet.Parameters.AddWithValue("@monRefSer", monRefSer);                   // montoRefServicioTransporte
                                indet.Parameters.AddWithValue("@monRefCar", monRefCar);                   // montoRefCargaEfectiva
                                indet.Parameters.AddWithValue("@monRefUti", monRefUti);                   // montoRefCargaUtilNominal
                            }
                            if (rb_fnor.Checked == true && double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@ubiPtoOri", ubiPtoOri);                   // ubigeoPtoOrigen
                                indet.Parameters.AddWithValue("@dirPtoOri", dirPtoOri);                   // direccionCompletaPtoOrigen
                                indet.Parameters.AddWithValue("@ubiPtoDes", ubiPtoDes);                   // ubigeoPtoDestino
                                indet.Parameters.AddWithValue("@dirPtoDes", dirPtoDes);                   // direccionCompletaPtoDestino
                                indet.Parameters.AddWithValue("@detViaje", detViaje);                     // detalleViaje
                                indet.Parameters.AddWithValue("@monRefSer", monRefSer);                   // montoRefServicioTransporte
                                indet.Parameters.AddWithValue("@monRefCar", monRefCar);                   // montoRefCargaEfectiva
                                indet.Parameters.AddWithValue("@monRefUti", monRefUti);                   // montoRefCargaUtilNominal
                            }
                            indet.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error en insertar detalle bizlinks");
                    Application.Exit();
                    return;
                }
                string insadd = "";                             // segmento adicionales cabecera     SPE_EINVOICEHEADER_ADD
                try
                {
                    string diradq = (tx_dirRem.Text.Trim() == "") ? "-" : tx_dirRem.Text.Trim();
                    string ubiadq = tx_ubigRtt.Text;
                    string urbadq = "-";
                    string proadq = (tx_provRtt.Text.Trim() == "") ? "-" : tx_provRtt.Text.Trim();
                    string depadq = (tx_dptoRtt.Text.Trim() == "") ? "-" : tx_dptoRtt.Text.Trim();
                    string disadq = (tx_distRtt.Text.Trim() == "") ? "-" : tx_distRtt.Text.Trim();
                    string paiadq = "PE";
                    string formpa = (rb_cre.Checked == true) ? "1" : "0";
                    string fecpag = "";
                    decimal tvc = totvta;
                    insadd = "insert into SPE_EINVOICEHEADER_ADD (" +
                            "clave,numeroDocumentoEmisor,serieNumero,tipoDocumento,tipoDocumentoEmisor,valor) values ";
                    insadd = insadd + "('direccionAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@diradq)," +
                                "('provinciaAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@proadq)," +
                                "('departamentoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@depadq)," +
                                "('distritoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@disadq)," +
                                "('paisAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@paiadq)";
                    if (tx_dat_tdv.Text == codfact)                                                         // esto aplica solo para facturas
                    {
                        if (tx_dat_mone.Text == MonDeft)
                        {
                            if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra))
                            {
                                if (rb_fesp.Checked == true) tvc = tvc - decimal.Parse(tx_detrac.Text);
                                else tvc = tvc - Math.Round(decimal.Parse(tx_flete.Text) * decimal.Parse(Program.pordetra) / 100, 2);    // totalDetraccion
                            }
                        }
                        else
                        {
                            if (double.Parse(tx_flete.Text) * double.Parse(tx_tipcam.Text) > double.Parse(Program.valdetra) * double.Parse(tx_tipcam.Text))
                            {
                                if (rb_fesp.Checked == true) tvc = tvc - (decimal.Parse(tx_detrac.Text) / decimal.Parse(tx_tipcam.Text));
                                else tvc = tvc - (decimal.Parse(tx_flete.Text) * decimal.Parse(Program.pordetra) / 100);
                            }
                        }
                        if (rb_cre.Checked == true)
                        {
                            DateTime fecpd = Convert.ToDateTime(tx_fechope.Text);
                            fecpag = fecpd.AddDays(int.Parse(tx_dat_diasp.Text)).ToString("yyyy'-'MM'-'dd");  // ToString("dd'/'MM'/'yyyy")
                        }
                        insadd = insadd + ",('formaPago',@nudoem,@sernum,@tipdoc,@tidoem,@fpagoD)";
                        if (formpa == "0") insadd = insadd + ",('formaPagoNegociable',@nudoem,@sernum,@tipdoc,@tidoem,@formpa)";
                        else
                        {
                            insadd = insadd + ",('formaPagoNegociable',@nudoem,@sernum,@tipdoc,@tidoem,@formpa)," +
                                "('montoNetoPendiente',@nudoem,@sernum,@tipdoc,@tidoem,@totvta)," +
                                "('montoPagoCuota1',@nudoem,@sernum,@tipdoc,@tidoem,@totvta)," +
                                "('fechaPagoCuota1',@nudoem,@sernum,@tipdoc,@tidoem,@fecpag)";
                        }
                    }
                    //  "('ubigeoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@ubiadq)," +   // no va por el momento hasta tenerlo
                    SqlCommand insertadd = new SqlCommand(insadd, conms);
                    insertadd.Parameters.AddWithValue("@nudoem", nudoem);
                    insertadd.Parameters.AddWithValue("@sernum", sernum);
                    insertadd.Parameters.AddWithValue("@tipdoc", tipdoc);
                    insertadd.Parameters.AddWithValue("@tidoem", tidoem);
                    insertadd.Parameters.AddWithValue("@diradq", diradq);
                    //insertadd.Parameters.AddWithValue("@ubiadq", ubiadq);
                    insertadd.Parameters.AddWithValue("@proadq", proadq);
                    insertadd.Parameters.AddWithValue("@depadq", depadq);
                    insertadd.Parameters.AddWithValue("@disadq", disadq);
                    insertadd.Parameters.AddWithValue("@paiadq", paiadq);
                    if (tx_dat_tdv.Text == codfact)  // formpa == "1"
                    {
                        insertadd.Parameters.AddWithValue("@formpa", formpa);
                        insertadd.Parameters.AddWithValue("@fpagoD", "009");
                        insertadd.Parameters.AddWithValue("@totvta", tvc.ToString("#0.00")); // totvta.ToString("#0.00")
                        insertadd.Parameters.AddWithValue("@fecpag", fecpag);
                    }
                    insertadd.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en insertar adicional cabecera");
                    Application.Exit();
                    return;
                }
                string insertcab = "";                          // segmento cabecera
                try
                {
                    insertcab = "insert into SPE_EINVOICEHEADER (serieNumero,fechaEmision,tipoDocumento,tipoMoneda," +
                        "numeroDocumentoEmisor,tipoDocumentoEmisor,nombreComercialEmisor,razonSocialEmisor,correoEmisor,codigoLocalAnexoEmisor," +
                        "ubigeoEmisor,direccionEmisor,provinciaEmisor,departamentoEmisor,distritoEmisor,urbanizacion,paisEmisor,codigoAuxiliar40_1,textoAuxiliar40_1," +
                        "tipoDocumentoAdquiriente,numeroDocumentoAdquiriente,razonSocialAdquiriente,correoAdquiriente,totalImpuestos," +
                        "totalValorVentaNetoOpGravadas,codigoLeyenda_1,textoLeyenda_1,bl_estadoRegistro," +
                        "totalIgv,totalVenta,tipoOperacion,totalValorVenta,totalPrecioVenta,codigoAuxiliar500_1,textoAuxiliar500_1";
                    if (!string.IsNullOrEmpty(nudor2) && !string.IsNullOrWhiteSpace(nudor2))
                    {
                        insertcab = insertcab + ",codigoAuxiliar500_2,textoAuxiliar500_2";
                    }
                    if (!string.IsNullOrEmpty(nudor3) && !string.IsNullOrWhiteSpace(nudor3))
                    {
                        insertcab = insertcab + ",codigoAuxiliar500_3,textoAuxiliar500_3";
                    }
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                                                                        // tipo de operacion con detraccion
                        if (rb_fesp.Checked == true)
                        {
                            if (Math.Round(double.Parse(tx_detrac.Text), 2) >= (Math.Round(double.Parse(tx_fletMN.Text) * double.Parse(Program.pordetra) / 100, 2))) { 
                                totdet = double.Parse(tx_detrac.Text); 
                            }
                            else { 
                                totdet = Math.Round(double.Parse(tx_fletMN.Text) * double.Parse(Program.pordetra) / 100, 2); 
                            }
                            insertcab = insertcab + ",codigoDetraccion,totalDetraccion,porcentajeDetraccion,numeroCtaBancoNacion,codigoLeyenda_2,textoLeyenda_2"; // codigoAuxiliar500_2,textoAuxiliar500_2
                        }
                        else
                        {
                            totdet = Math.Round(double.Parse(tx_fletMN.Text) * double.Parse(Program.pordetra) / 100, 2);    // totalDetraccion
                            insertcab = insertcab + ",codigoDetraccion,totalDetraccion,porcentajeDetraccion,numeroCtaBancoNacion,codigoLeyenda_2,textoLeyenda_2";
                        }
                    }
                    insertcab = insertcab + ") " +
                        "values (@sernum,@fecemi,@tipdoc,@tipmon," +
                        "@nudoem,@tidoem,@nocoem,@rasoem,@coremi,@coloem," +
                        "@ubiemi,@diremi,@provemi,@depaemi,@distemi,@urbemi,@pasiemi,@codaux40_1,@etiaux40_1," +
                        "@tidoad,@nudoad,@rasoad,@coradq,@totimp," +
                        "@tovane,@coley1,@teley1,@estreg," +
                        "@totigv,@totvta,@tipope,@tovane,@totvta,@tiref1,@nudor1";
                    if (!string.IsNullOrEmpty(nudor2) && !string.IsNullOrWhiteSpace(nudor2)) insertcab = insertcab + ",@tiref2,@nudor2";
                    if (!string.IsNullOrEmpty(nudor3) && !string.IsNullOrWhiteSpace(nudor3)) insertcab = insertcab + ",@tiref3,@nudor3";
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        if (rb_fesp.Checked == true)
                        {
                            insertcab = insertcab + ",@coddetra,@totdet,@pordetra,@ctadetra,@codleyt,@tauxdet1";   // @cauxdet1,
                        }
                        else
                        {
                            insertcab = insertcab + ",@coddetra,@totdet,@pordetra,@ctadetra,@codleyt,@leydet";
                        }
                    }
                    insertcab = insertcab + ")";
                    //
                    SqlCommand inserta = new SqlCommand(insertcab, conms);
                    inserta.Parameters.AddWithValue("@sernum", sernum);
                    inserta.Parameters.AddWithValue("@fecemi", fecemi);
                    inserta.Parameters.AddWithValue("@tipdoc", tipdoc);
                    inserta.Parameters.AddWithValue("@tipmon", tipmon);
                    inserta.Parameters.AddWithValue("@nudoem", nudoem);
                    inserta.Parameters.AddWithValue("@tidoem", tidoem);
                    inserta.Parameters.AddWithValue("@nocoem", nocoem);
                    inserta.Parameters.AddWithValue("@rasoem", rasoem);
                    inserta.Parameters.AddWithValue("@coremi", coremi);
                    inserta.Parameters.AddWithValue("@coloem", coloem);
                    inserta.Parameters.AddWithValue("@ubiemi", ubiemi);                                                       // v ubigeoEmisor
                    inserta.Parameters.AddWithValue("@diremi", diremi);                                                       // v direccionEmisor
                    inserta.Parameters.AddWithValue("@provemi", provemi);                                                     // v provinciaEmisor
                    inserta.Parameters.AddWithValue("@depaemi", depaemi);                                                     // v departamentoEmisor
                    inserta.Parameters.AddWithValue("@distemi", distemi);                                                     // v distritoEmisor
                    inserta.Parameters.AddWithValue("@urbemi", urbemis);                                                      // v urbanizacion
                    inserta.Parameters.AddWithValue("@pasiemi", pasiemi);                                                     // v paisEmisor
                    inserta.Parameters.AddWithValue("@codaux40_1", codaux40_1);                                               // v codigoAuxiliar40_1,
                    inserta.Parameters.AddWithValue("@etiaux40_1", etiaux40_1);                                               // v textoAuxiliar40_1
                    inserta.Parameters.AddWithValue("@tidoad", tidoad);
                    inserta.Parameters.AddWithValue("@nudoad", nudoad);
                    inserta.Parameters.AddWithValue("@rasoad", rasoad);
                    inserta.Parameters.AddWithValue("@coradq", coradq);
                    inserta.Parameters.AddWithValue("@totimp", totimp);
                    inserta.Parameters.AddWithValue("@tovane", tovane);
                    inserta.Parameters.AddWithValue("@totigv", totigv);
                    inserta.Parameters.AddWithValue("@totvta", totvta);
                    if (string.IsNullOrEmpty(tiref1) == true) inserta.Parameters.AddWithValue("@tiref1", DBNull.Value);
                    else inserta.Parameters.AddWithValue("@tiref1", Ctiref1);
                    if (string.IsNullOrEmpty(nudor1) == true) inserta.Parameters.AddWithValue("@nudor1", DBNull.Value);
                    else inserta.Parameters.AddWithValue("@nudor1", Cnudor1);
                    inserta.Parameters.AddWithValue("@tipope", tipope);
                    inserta.Parameters.AddWithValue("@coley1", coley1);
                    inserta.Parameters.AddWithValue("@teley1", teley1);
                    inserta.Parameters.AddWithValue("@estreg", estreg);
                    if (Ctiref2 != "")
                    {
                        inserta.Parameters.AddWithValue("@tiref2", Ctiref2);
                        inserta.Parameters.AddWithValue("@nudor2", Cnudor2);
                    }
                    if (Ctiref3 != "")
                    {
                        inserta.Parameters.AddWithValue("@tiref3", Ctiref3);
                        inserta.Parameters.AddWithValue("@nudor3", Cnudor3);
                    }
                    inserta.Parameters.AddWithValue("@coddetra", Program.coddetra);
                    inserta.Parameters.AddWithValue("@totdet", totdet);
                    inserta.Parameters.AddWithValue("@pordetra", Program.pordetra);
                    inserta.Parameters.AddWithValue("@ctadetra", Program.ctadetra);
                    inserta.Parameters.AddWithValue("@codleyt", codleyt);
                    inserta.Parameters.AddWithValue("@leydet", leydet);
                    inserta.Parameters.AddWithValue("@cauxdet1", cauxdet1);
                    inserta.Parameters.AddWithValue("@tauxdet1", tauxdet1);
                    inserta.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(insertcab + Environment.NewLine + ex.Message);
                    Application.Exit();
                    return;
                }
                conms.Close();
            }
            else
            {
                MessageBox.Show("No se puede conectar al servidor" + Environment.NewLine +
                "de Facturación Electrónica", "Error de conexión para escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
        }
        private void anulfactelec(int cta)                  // anula (baja) de comprobante electrónico
        {
            using (SqlConnection conn = new SqlConnection(script))
            {
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    string nudoem = Program.ruc;                                                            // v numeroDocumentoEmisor
                    string tidoem = "6";                                                                    // v tipoDocumentoEmisor
                    string restip = (tx_dat_tdv.Text == codfact) ? "RA" : "RC";
                    string resuid = restip + "-" + tx_fechact.Text.Substring(6, 4) + tx_fechact.Text.Substring(3, 2) + tx_fechact.Text.Substring(0, 2) + "-" + lib.Right("00"+cta,3);    // resumen id
                    string fecomp = tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2);
                    string febaja = tx_fechact.Text.Substring(6, 4) + "-" + tx_fechact.Text.Substring(3, 2) + "-" + tx_fechact.Text.Substring(0, 2);
                    string rasoem = Program.cliente.Trim();                                                 // v razonSocialEmisor
                    string coremi = Program.mailclte.Trim();                                               // v correoEmisor
                    string estreg = "A";                                                                    // bl_estadoRegistro
                    //
                    DataRow[] row = dttd1.Select("idcodice='" + tx_dat_tdv.Text + "'");                     // tipo de documento venta
                    string tipdoc = row[0][3].ToString();                                                   // v tipoDocumento
                    string serdoc = cmb_tdv.Text.Substring(0, 1) + tx_serie.Text;                           // v serieNumero
                    string numdoc = lib.Right("0000000" + tx_numero.Text,8);                                // numero 
                    string motivo = "Anulación por error en emisión";                                       // motivo de la baja
                    string estite = "3";                                                                    // estado item = 3 = anulado
                    row = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");
                    string tipmon = row[0][2].ToString();                                                   // tipo moneda sunat
                    string numcor = serdoc + "-" + numdoc;                                                  // numero correlativo
                    row = dttd0.Select("idcodice='" + tx_dat_tdRem.Text + "'");
                    string tidoad = row[0][3].ToString().Trim();                                            // v tipoDocumentoAdquiriente
                    string nudoad = tx_numDocRem.Text;                                                      // v numeroDocumentoAdquiriente
                    string subtot = tx_subt.Text;                                                           // totalValorVentaOpGravadasConIgv
                    string totvta = tx_flete.Text;                                                          // totalVenta
                    string totisc = "0.00";                                                                 // totalIsc
                    string icbpe = "0.00";                                                                  // totalICBPER
                    string totigv = tx_igv.Text;                                                            // total Igv
                    // 
                    if (tx_dat_tdv.Text == codfact)                             // anulación facturas
                    {
                        string insertcab = "insert into SPE_CANCELHEADER (numeroDocumentoEmisor,tipoDocumentoEmisor," +
                            "resumenId,fechaEmisionComprobante,fechaGeneracionResumen,razonSocialEmisor,correoEmisor,resumenTipo,inhabilitado,bl_estadoRegistro) " +
                            "values (@nudoem,@tidoem,@resuid,@fecomp,@febaja,@rasoem,@coremi,@restip,1,@estreg)";
                        SqlCommand sqlc = new SqlCommand(insertcab, conn);
                        sqlc.Parameters.AddWithValue("@nudoem", nudoem);
                        sqlc.Parameters.AddWithValue("@tidoem", tidoem);
                        sqlc.Parameters.AddWithValue("@resuid", resuid);
                        sqlc.Parameters.AddWithValue("@fecomp", fecomp);
                        sqlc.Parameters.AddWithValue("@febaja", febaja);
                        sqlc.Parameters.AddWithValue("@rasoem", rasoem);
                        sqlc.Parameters.AddWithValue("@coremi", coremi);
                        sqlc.Parameters.AddWithValue("@restip", restip);
                        sqlc.Parameters.AddWithValue("@estreg", estreg);
                        sqlc.ExecuteNonQuery();
                        // 
                        string insertdet = "insert into SPE_CANCELDETAIL (" +
                            "tipoDocumentoEmisor,numeroDocumentoEmisor,resumenId,numeroFila,tipoDocumento,serieDocumentoBaja,numeroDocumentoBaja,motivoBaja) " +
                            "values (@tidoem,@nudoem,@resuid,1,@tipdoc,@serdoc,@numdoc,@motivo)";
                        sqlc = new SqlCommand(insertdet, conn);
                        sqlc.Parameters.AddWithValue("@tidoem", tidoem);
                        sqlc.Parameters.AddWithValue("@nudoem", nudoem);
                        sqlc.Parameters.AddWithValue("@resuid", resuid);
                        sqlc.Parameters.AddWithValue("@tipdoc", tipdoc);
                        sqlc.Parameters.AddWithValue("@serdoc", serdoc);
                        sqlc.Parameters.AddWithValue("@numdoc", numdoc);
                        sqlc.Parameters.AddWithValue("@motivo", motivo);
                        sqlc.ExecuteNonQuery();
                    }
                    if (tx_dat_tdv.Text == codbole)                             // anulación boletas
                    {
                        string incabbol = "insert into SPE_SUMMARYHEADER (tipoDocumentoEmisor,numeroDocumentoEmisor,resumenId," +
                            "fechaEmisionComprobante,fechaGeneracionResumen,razonSocialEmisor,correoEmisor,resumenTipo,inhabilitado,bl_estadoRegistro) " +
                            "values (@tidoem,@nudoem,@resuid,@fecomp,@febaja,@rasoem,@coremi,@restip,1,@estreg)";
                        SqlCommand sqlb = new SqlCommand(incabbol, conn);
                        sqlb.Parameters.AddWithValue("@tidoem", tidoem);
                        sqlb.Parameters.AddWithValue("@nudoem", nudoem);
                        sqlb.Parameters.AddWithValue("@resuid", resuid);
                        sqlb.Parameters.AddWithValue("@fecomp", fecomp);
                        sqlb.Parameters.AddWithValue("@febaja", febaja);
                        sqlb.Parameters.AddWithValue("@rasoem", rasoem);
                        sqlb.Parameters.AddWithValue("@coremi", coremi);
                        sqlb.Parameters.AddWithValue("@restip", restip);
                        sqlb.Parameters.AddWithValue("@estreg", estreg);
                        sqlb.ExecuteNonQuery();
                        //
                        string indetbol = "insert into SPE_SUMMARY_ITEM (tipoDocumentoEmisor,numeroDocumentoEmisor,resumenId,numeroFila,tipoDocumento,tipoMoneda," +
                            "numeroCorrelativo,tipoDocumentoAdquiriente,numeroDocumentoAdquiriente,estadoItem,totalValorVentaOpGravadaConIgv,totalVenta,totalIsc,totalIgv,totalICBPER) " +
                            "values (@tidoem,@nudoem,@resuid,1,@tipdoc,@tipmon,@numcor,@tidoad,@nudoad,@estite,@subtot,@totvta,@totisc,@totigv,@icbpe)";
                        sqlb = new SqlCommand(indetbol, conn);
                        sqlb.Parameters.AddWithValue("@tidoem", tidoem);
                        sqlb.Parameters.AddWithValue("@nudoem", nudoem);
                        sqlb.Parameters.AddWithValue("@resuid", resuid);
                        sqlb.Parameters.AddWithValue("@tipdoc", tipdoc);    // tipoDocumento
                        sqlb.Parameters.AddWithValue("@tipmon", tipmon);    // tipoMoneda
                        sqlb.Parameters.AddWithValue("@numcor", numcor);    // numeroCorrelativo
                        sqlb.Parameters.AddWithValue("@tidoad", tidoad);    // tipoDocumentoAdquiriente
                        sqlb.Parameters.AddWithValue("@nudoad", nudoad);    // numeroDocumentoAdquiriente
                        sqlb.Parameters.AddWithValue("@estite", estite);    // estado item = 3 = anulado
                        sqlb.Parameters.AddWithValue("@subtot", subtot);    // subtotal = totalValorVentaOpGravadasConIgv
                        sqlb.Parameters.AddWithValue("@totvta", totvta);    // @totvta
                        sqlb.Parameters.AddWithValue("@totisc", totisc);    // @totisc
                        sqlb.Parameters.AddWithValue("@totigv", totigv);    // @totigv
                        sqlb.Parameters.AddWithValue("@icbpe", icbpe);      // total icbpe
                        sqlb.ExecuteNonQuery();
                    }
                }
                else
                {
                    MessageBox.Show("Se perdió conexión con el servidor" + Environment.NewLine +
                        "de facturación electrónica" + Environment.NewLine +
                        "Notifique a Contabilidad y sistemas!","Transacción electrónica NO HECHA",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
        }
        #endregion

        #region autocompletados
        private void autodepa()                             // se jala en el load
        {
            if (dataUbig == null)
            {
                DataTable dataUbig = (DataTable)CacheManager.GetItem("ubigeos");
            }
            DataRow[] depar = dataUbig.Select("depart<>'00' and provin='00' and distri='00'");
            departamentos.Clear();
            foreach (DataRow row in depar)
            {
                departamentos.Add(row["nombre"].ToString());
            }
        }
        private void autoprov(string donde)                 // se jala despues de ingresado el departamento
        {
            switch(donde)
            {
                case "cliente":
                    if (tx_dptoRtt.Text.Trim() != "")
                    {
                        DataRow[] provi = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and provin<>'00' and distri='00'");
                        provincias.Clear();
                        foreach (DataRow row in provi)
                        {
                            provincias.Add(row["nombre"].ToString());
                        }
                    }
                    break;
            } 
        }
        private void autodist(string donde)                 // se jala despues de ingresado la provincia
        {
            switch (donde)
            {
                case "cliente":
                    if (tx_ubigRtt.Text.Trim() != "" && tx_provRtt.Text.Trim() != "")
                    {
                        DataRow[] distr = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and provin='" + tx_ubigRtt.Text.Substring(2, 2) + "' and distri<>'00'");
                        distritos.Clear();
                        foreach (DataRow row in distr)
                        {
                            distritos.Add(row["nombre"].ToString());
                        }
                    }
                    break;
                case "partida":
                    if (tx_dat_upo.Text.Trim() != "")
                    {
                        DataRow[] distr = dataUbig.Select("depart='" + tx_dat_upo.Text.Substring(0, 2) + "' and provin='" + tx_dat_upo.Text.Substring(2, 2) + "' and distri<>'00'");
                        distritos.Clear();
                        foreach (DataRow row in distr)
                        {
                            distritos.Add(row["nombre"].ToString());
                        }
                    }
                    break;
                case "llegada":
                    if (tx_dat_upd.Text.Trim() != "")
                    {
                        DataRow[] distr = dataUbig.Select("depart='" + tx_dat_upd.Text.Substring(0, 2) + "' and provin='" + tx_dat_upd.Text.Substring(2, 2) + "' and distri<>'00'");
                        distritos.Clear();
                        foreach (DataRow row in distr)
                        {
                            distritos.Add(row["nombre"].ToString());
                        }
                    }
                    break;
            }
        }
        #endregion autocompletados

        #region limpiadores_modos
        private void sololee()
        {
            lp.sololee(this);
            lp.sololeePanel(panel2);
            lp.sololeePanel(panel1);
            rb_fesp.Enabled = false;
            rb_fnor.Enabled = false;
        }
        private void escribe()
        {
            lp.escribe(this);
            lp.escribePanel(panel2);
            lp.escribePanel(panel1);
            tx_nomRem.ReadOnly = true;
            rb_fnor.Enabled = true;
            rb_fesp.Enabled = true;
            //tx_dirRem.ReadOnly = true;
            //tx_dptoRtt.ReadOnly = true;
            //tx_provRtt.ReadOnly = true;
            //tx_distRtt.ReadOnly = true;
        }
        private void limpiar()
        {
            lp.limpiar(this);
            lp.limpiapanel(panel2);
        }
        private void limpia_chk()    
        {
            lp.limpia_chk(this);
            rb_si.Checked = false;
            rb_no.Checked = false;
            rb_cre.Checked = false;
        }
        private void limpia_otros()
        {
            //
        }
        private void limpia_combos()
        {
            lp.limpia_cmb(this);
            cmb_plazoc.SelectedIndex = -1;
            cmb_tipop.SelectedIndex = -1;
        }
        #endregion limpiadores_modos;

        #region boton_form GRABA EDITA ANULA
        private void bt_agr_Click(object sender, EventArgs e)
        {
            if (tx_serGR.Text.Trim() != "" && tx_numGR.Text.Trim() != "" && Tx_modo.Text == "NUEVO")
            {
                // validamos que no se repita la GR
                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    if (dataGridView1.Rows[i].Cells[0].Value != null)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString().Trim() == (tx_serGR.Text.Trim() + "-" + tx_numGR.Text.Trim()))
                        {
                            MessageBox.Show("Esta repitiendo la Guía!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            tx_numGR.Text = "";
                            tx_numGR.Focus();
                            return;
                        }
                    }
                }
                // validamos que no sea mas guias de las posibles
                if (chk_consol.Checked == true && dataGridView1.Rows.Count > int.Parse(v_mfildet))
                {
                    MessageBox.Show("Limite máximo de guías alcanzado!", "Atención - modo consolidado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tx_numGR.Text = "";
                    tx_numGR.Focus();
                    return;
                }
                if (chk_consol.Checked == false && dataGridView1.Rows.Count > int.Parse(v_mfdetn))
                {
                    MessageBox.Show("Limite máximo de guías alcanzado!", "Atención - modo simple", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tx_numGR.Text = "";
                    tx_numGR.Focus();
                    return;
                }
                // validamos que la GR: 1.exista, 2.No este facturada, 3.No este anulada
                if (validGR(tx_serGR.Text, tx_numGR.Text) == false)
                {
                    MessageBox.Show("La GR no existe, esta anulada o ya esta facturada", "Error en Guía", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tx_numGR.Text = "";
                    tx_numGR.Focus();
                    return;
                }
                // validamos que alguno de los clientes de la GR sea igual al existente en el formulario
                if (int.Parse(tx_tfil.Text) > 0)
                {
                    /*
                    if ((tx_dat_tdRem.Text + tx_numDocRem != datcltsR[0].ToString() + datcltsR[1]) && (tx_dat_tdRem.Text + tx_numDocRem != datcltsD[0].ToString() + datcltsD[1]))
                    {
                        var aa = MessageBox.Show("El remitente y destinatario de la GR" + Environment.NewLine +
                            "no coinciden con el cliente del comprobante" + Environment.NewLine +
                            "Desea continuar?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (aa == DialogResult.No)
                        {
                            inivarGR();
                            tx_numGR.Text = "";
                            tx_numGR.Focus();
                            return;
                        }
                    }
                    */
                }
                else
                {
                    rb_desGR.PerformClick();
                }
                /* validamos que las guias esten con saldo o no
                if (rb_si.Checked == false && rb_no.Checked == false)
                {
                    // estamos ante la primera guia
                    if (double.Parse(datguias[16].ToString()) > 0)
                    {
                        rb_no.Checked = true;
                        rb_no.PerformClick();
                    }
                    else
                    {
                        rb_si.Checked = true;
                        rb_si.PerformClick();
                    }
                }
                else
                {
                    // NO USAMOS ESTA PARTE PORQUE NO GENERAMOS COBRANZAS AUTOMATICAS DESDE ESTE FORM ... 15/03/2022 || SI COBRAMOS EN AUTOMATICO 21/03/2022
                    if (rb_si.Checked == true && double.Parse(datguias[16].ToString()) > 0)
                    {
                        MessageBox.Show("Las Guías deben ser todas con o sin saldo","Error en ingreso",MessageBoxButtons.OK,MessageBoxIcon.Error);
                        tx_numGR.Text = "";
                        tx_numGR.Focus();
                        return;
                    }
                    if (rb_no.Checked == true && double.Parse(datguias[16].ToString()) <= 0)
                    {
                        MessageBox.Show("Las Guías deben ser todas con o sin saldo", "Error en ingreso", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tx_numGR.Text = "";
                        tx_numGR.Focus();
                        return;
                    }
                }
                */
                if (dataGridView1.Rows.Count >= 2)
                {
                    if ((double.Parse(datguias[16].ToString()) > 0 && double.Parse(tx_salxcob.Text) <= 0) ||
                        (double.Parse(datguias[16].ToString()) <= 0 && double.Parse(tx_salxcob.Text) > 0))
                    {
                        MessageBox.Show("Las Guías deben ser todas con o sin saldo", "Error en ingreso", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        tx_numGR.Text = "";
                        tx_numGR.Focus();
                        return;
                    }
                }
                //dataGridView1.Rows.Clear(); nooooo, se puede hacer una fact de varias guias, n guias
                dataGridView1.Rows.Add(datguias[0], datguias[1], datguias[2], datguias[3], datguias[4], datguias[5], datguias[6], datguias[9], datguias[10], datguias[7], datguias[15],datguias[16],datguias[17]);     // insertamos en la grilla los datos de la GR
                totalizaG();
                tx_dat_mone.Text = datguias[7].ToString();
                cmb_mon.SelectedValue = datguias[7].ToString();
                if (tx_dat_mone.Text != MonDeft && datguias[9].ToString().Substring(0,10) != tx_fechope.Text)
                {
                    // llamanos a tipo de cambio
                    vtipcam vtipcam = new vtipcam("", tx_dat_mone.Text, DateTime.Now.Date.ToString());
                    var result = vtipcam.ShowDialog();
                    //tx_flete.Text = vtipcam.ReturnValue1;
                    //tx_fletMN.Text = vtipcam.ReturnValue2;
                    tx_tipcam.Text = (vtipcam.ReturnValue3 == null)? "0" : vtipcam.ReturnValue3;
                    tx_fletMN.Text = Math.Round(decimal.Parse(tx_flete.Text) * decimal.Parse(tx_tipcam.Text), 2).ToString();
                }
                else
                {
                    tx_tipcam.Text = datguias[8].ToString();
                }
                if (int.Parse(tx_tfil.Text) == int.Parse(v_mfildet))
                {
                    MessageBox.Show("Número máximo de filas de detalle", "El formato no permite mas", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    dataGridView1.AllowUserToAddRows = false;
                }
                else
                {
                    dataGridView1.AllowUserToAddRows = true;
                }
                rb_no.Enabled = true;
                // comprobación de filas de guias, pagos y saldos, si hay + de 1 fila y alguna esta pagada => no se permite cobrar automatico
                if (dataGridView1.Rows.Count >= 3 && (decimal.Parse(tx_dat_saldoGR.Text) <= 0 && decimal.Parse(dataGridView1.Rows[0].Cells[11].Value.ToString()) > 0 
                    || decimal.Parse(tx_dat_saldoGR.Text) > 0 && decimal.Parse(dataGridView1.Rows[0].Cells[11].Value.ToString()) <= 0))
                {
                    /*
                    MessageBox.Show("El presente comprobante no se " + Environment.NewLine +
                         "puede cobrar en automático", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    rb_si.Checked = false;
                    rb_si.Enabled = false;
                    tx_salxcob.Text = tx_flete.Text;
                    tx_pagado.Text = "0";
                    */
                    MessageBox.Show("El presente comprobante tiene guias" + Environment.NewLine +
                         "con y sin saldo, no se permite continuar", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Bt_add_Click(null, null);
                    return;
                }
                if (dataGridView1.Rows.Count <= 2 && decimal.Parse(tx_dat_saldoGR.Text) <= 0)
                {
                    MessageBox.Show("La GR esta cancelada, el documento de venta"+ Environment.NewLine +
                         "se creará con el estado cancelado","Atención verifique",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    //rb_si.PerformClick();
                    rb_no.Enabled = false;
                    rb_cre.Enabled = false;
                    rb_si.Enabled = true;
                    rb_si.Checked = true;
                    tx_salxcob.Text = "0";
                    tx_pagado.Text = tx_flete.Text;
                }
                else
                {
                    //tx_flete.ReadOnly = true;
                    if (cusdscto.Contains(asd)) tx_flete.ReadOnly = false;
                    else tx_flete.ReadOnly = true;
                }
                tx_flete_Leave(null, null);
                //rb_si.Checked = false;
                //rb_no.Checked = false;   // true
                rb_no.PerformClick();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            #region validaciones
            if (tx_serie.Text.Trim() == "")
            {
                tx_serie.Focus();
                return;
            }
            if (tx_dat_mone.Text.Trim() == "")
            {
                MessageBox.Show("Seleccione el tipo de moneda", " Atención ");
                cmb_mon.Focus();
                return;
            }
            if (tx_flete.Text.Trim() == "" || tx_flete.Text.Trim() == "0")
            {
                MessageBox.Show("No existe valor del documento", " Atención ");
                tx_flete.Focus();
                return;
            }
            if (tx_tfil.Text.Trim() == "")
            {
                MessageBox.Show("Ingrese el detalle del documento de venta", "Faltan ingresar guías");
                tx_serGR.Focus();
                return;
            }
            if (tx_dat_tdRem.Text.Trim() == "")
            {
                MessageBox.Show("Seleccione el documento de cliente", " Error en Cliente ");
                tx_dat_tdRem.Focus();
                return;
            }
            if (tx_numDocRem.Text.Trim() == "")
            {
                MessageBox.Show("Ingrese el número de documento", " Error en Cliente ");
                tx_numDocRem.Focus();
                return;
            }
            if (tx_nomRem.Text.Trim() == "")
            {
                MessageBox.Show("Ingrese el nombre o razón social", " Error en Cliente ");
                tx_nomRem.Focus();
                return;
            }
            if (tx_email.Text.Trim() == "")
            {
                MessageBox.Show("Debe ingresar un correo electrónico", " Error en Cliente ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                tx_email.Focus();
                return;
            }
            if (tx_dat_tdv.Text == codfact || (tx_dat_tdv.Text == codbole && decimal.Parse(tx_fletMN.Text) > limbolsd))
            {
                if (tx_dirRem.Text.Trim() == "")
                {
                    MessageBox.Show("Ingrese la dirección", " Error en Remitente ");
                    tx_dirRem.Focus();
                    return;
                }
                if (tx_dptoRtt.Text.Trim() == "" || tx_provRtt.Text.Trim() == "" || tx_distRtt.Text.Trim() == "")
                {
                    MessageBox.Show("Ingrese departamento, provincia y distrito", "Dirección incompleta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tx_dptoRtt.Focus();
                    return;
                }
            }
            /*
            if (tx_dat_tdec.Text != tx_dat_tdRem.Text)
            {
                MessageBox.Show("Asegurese que el tipo de documento de venta" + Environment.NewLine +
                    "sean coincidente con el tipo de cliente", "Error de tipos", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                cmb_docRem.Focus();
                return;
            }*/
            if (tx_dat_tdec.Text != tx_dat_tdRem.Text)
            {
                // aca validamos que el tipo de doc de venta se corresponda con el documento del cliente
                if (tx_dat_tdv.Text != codfact)
                {
                    if (!tdocsBol.Contains(tx_dat_tdRem.Text))
                    {
                        MessageBox.Show("Asegurese que el tipo de documento de venta" + Environment.NewLine +
                            "sean coincidente con el tipo de cliente", "Error de tipo Boleta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cmb_docRem.Focus();
                        return;
                    }
                }
                else
                {
                    if (!tdocsFac.Contains(tx_dat_tdRem.Text))
                    {
                        MessageBox.Show("Asegurese que el tipo de documento de venta" + Environment.NewLine +
                            "sean coincidente con el tipo de cliente", "Error de tipo Factura", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        cmb_docRem.Focus();
                        return;
                    }
                }
            }
            #endregion
            // grabamos, actualizamos, etc
            string modo = Tx_modo.Text;
            string iserror = "no";
            if (modo == "NUEVO")
            {
                // valida pago y calcula
                if (rb_si.Checked == false && rb_no.Checked == false && rb_cre.Checked == false)
                {
                    MessageBox.Show("Seleccione si se cancela la factura o no","Atención - Confirme",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    rb_si.Focus();
                    return;
                }
                if (rb_si.Checked == true && decimal.Parse(tx_saldoT.Text) <= 0)
                {
                    MessageBox.Show("No hay saldo que cobrar!","Error en condición de pago",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    //return;
                    rb_no.Checked = true;
                    return;
                } 
                if (tx_dat_diasp.Text.Trim() == "" && rb_cre.Checked == true)
                {
                    MessageBox.Show("Seleccione el plazo de credito", "Atención - Confirme", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    cmb_plazoc.Focus();
                    return;
                }
                if (tx_pagado.Text.Trim() == "" && tx_salxcob.Text.Trim() == "")
                {
                    MessageBox.Show("Seleccione si se cancela la factura o no", "Atención - Confirme", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    rb_si.Focus();
                    return;
                }
                if (tx_dat_mone.Text != MonDeft && tx_tipcam.Text == "" || tx_tipcam.Text == "0")
                {
                    MessageBox.Show("Problemas con el tipo de cambio","Atención",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    cmb_mon.Focus();
                    return;
                }
                if (tx_dat_mone.Text != MonDeft && decimal.Parse(tx_tipcam.Text) > 1)
                {
                    if (Math.Round(decimal.Parse(tx_tfmn.Text), 1) != Math.Round(decimal.Parse(tx_fletMN.Text), 1))
                    {
                        MessageBox.Show("El valor a facturar no puede ser diferente al valor de la(s) GR");
                        tx_flete.Focus();
                        return;
                    }
                }
                // validaciones si el comprobante es "especial"
                if (rb_fesp.Checked == true)
                {
                    if (double.Parse(tx_fletMN.Text) < double.Parse(Program.valdetra))
                    {
                        MessageBox.Show("El valor del comprobante debe estar en el rango de detracción" + Environment.NewLine + 
                            "Los comprobantes especiales tienen que tener detracción","Tipo de comprobante erróneo",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        return;
                    }
                    if (tx_e_aut.Text.Trim() == "")
                    {
                        MessageBox.Show("Falta Autorización de circulación","Config. Fact. Especial",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        tx_e_aut.Focus();
                        return;
                    }
                    if (tx_e_cant.Text == "" || tx_e_cant.Text == "0")
                    {
                        MessageBox.Show("Falta cantidad en toneladas", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_cant.Focus();
                        return;
                    }
                    if (tx_e_dirlle.Text.Trim() == "")
                    {
                        MessageBox.Show("Falta dirección de llegada", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_dirlle.Focus();
                        return;
                    }
                    if (tx_e_dirpar.Text.Trim() == "")
                    {
                        MessageBox.Show("Falta dirección de partida", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_dirpar.Focus();
                        return;
                    }
                    if (tx_e_dnicho.Text.Trim() == "")
                    {
                        MessageBox.Show("Falta el DNI del chofer", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_dnicho.Focus();
                        return;
                    }
                    if (tx_e_ftras.Text == "")
                    {
                        MessageBox.Show("Falta la fecha inicial del traslado", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_ftras.Focus();
                        return;
                    }
                    if (tx_e_ntrans.Text == "")
                    {
                        MessageBox.Show("Falta el nombre del transportista", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_ntrans.Focus();
                        return;
                    }
                    if (tx_e_placa.Text.Trim() == "")
                    {
                        MessageBox.Show("Falta la placa del vehículo", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_placa.Focus();
                        return;
                    }
                    if (tx_e_prec.Text == "")
                    {
                        MessageBox.Show("Falta el precio por tonelada", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_prec.Focus();
                        return;
                    }
                    if (tx_e_ruct.Text.Trim() == "")
                    {
                        MessageBox.Show("Falta el ruc del transportista", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_ruct.Focus();
                        return;
                    }
                    if (tx_dat_ruta.Text.Trim() == "")
                    {
                        cmb_ruta.Focus();
                        MessageBox.Show("Seleccione la ruta", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if (tx_e_tn.Text.Trim() == "")
                    {
                        MessageBox.Show("Ingrese la carga util en TN", "Config. Fact. Especial", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_tn.Focus();
                        return;
                    }
                    if (tx_e_ubiori.Text.Trim().Length != 6)
                    {
                        MessageBox.Show("Debe ingresar el código de ubigeo" + Environment.NewLine +
                            "Presione F1", "Falta código de origen!",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        tx_e_ubiori.Focus();
                        return;
                    }
                    if (tx_e_ubides.Text.Trim().Length != 6)
                    {
                        MessageBox.Show("Debe ingresar el código de ubigeo" + Environment.NewLine +
                            "Presione F1", "Falta código de destino!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        tx_e_ubides.Focus();
                        return;
                    }
                    if (Math.Round(double.Parse(tx_e_cant.Text) * double.Parse(tx_e_prec.Text),1) != Math.Round(double.Parse(tx_subt.Text),1))
                    {
                        MessageBox.Show("Carga en TN por el precio debe ser igual" + Environment.NewLine +
                            "al valor del subtotal del comprobante","Error en importes",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                        tx_e_prec.Focus();
                        return;
                    }
                }
                if (tx_idr.Text.Trim() == "")
                {
                    var aa = MessageBox.Show("Confirma que desea crear el documento?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        if (valiconsql(script) == true)  // confirmas que se pueda abrir conexión con la base bizlinks
                        {
                            if (graba() == true)
                            {
                                grabfactelec();
                                if (rb_fnor.Checked == true)
                                {
                                    var bb = MessageBox.Show("Desea imprimir el documento?" + Environment.NewLine +
                                        "El formato actual es " + vi_formato, "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (bb == DialogResult.Yes)
                                    {
                                        Bt_print.PerformClick();
                                    }
                                }
                                // actualizamos la tabla seguimiento de usuarios
                                string resulta = lib.ult_mov(nomform, nomtab, asd);
                                if (resulta != "OK")
                                {
                                    MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }
                            else
                            {
                                MessageBox.Show("No se puede grabar el documento de venta electrónico" + Environment.NewLine +
                                    "Problemas en la conexión o en los datos del documento", "Error en conexión");
                                iserror = "si";
                            }
                        }
                        else
                        {
                            MessageBox.Show("No existe ruta, no es valida o no hay conexión" + Environment.NewLine +
                                        "para generar el comprobante electrónico", "Ruta para Fact.Electrónica", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                            iserror = "si";
                        }
                    }
                    else
                    {
                        tx_numDocRem.Focus();
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Los datos no son nuevos en doc.venta", "Verifique duplicidad", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                    return;
                }
            }
            if (modo == "EDITAR")
            {
                if (tx_numero.Text.Trim() == "")
                {
                    tx_numero.Focus();
                    MessageBox.Show("Ingrese el número del documento", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (tx_dat_estad.Text == codAnul)
                {
                    MessageBox.Show("El documento esta ANULADO", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    tx_numero.Focus();
                    return;
                }
                if (true)
                {
                    if (tx_idr.Text.Trim() != "")
                    {
                        var aa = MessageBox.Show("Confirma que desea modificar el documento?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (aa == DialogResult.Yes)
                        {
                            edita();    // modificacion solo observaciones
                            // actualizamos la tabla seguimiento de usuarios
                            string resulta = lib.ult_mov(nomform, nomtab, asd);
                            if (resulta != "OK")
                            {
                                MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else
                        {
                            tx_serie.Focus();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("El documento ya debe existir para editar", "Debe ser edición", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        return;
                    }
                }
            }
            if (modo == "ANULAR")
            {
                if (tx_numero.Text.Trim() == "")
                {
                    tx_numero.Focus();
                    MessageBox.Show("Ingrese el número del documento", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    return;
                }
                if (tx_dat_estad.Text == codAnul)
                {
                    MessageBox.Show("El documento esta ANULADO", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    tx_numero.Focus();
                    return;
                }
                if (tx_idcob.Text != "")
                {
                    MessageBox.Show("El documento de venta tiene Cobranza activa" + Environment.NewLine +
                        "La cobranza permanece sin cambios", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //return;
                }
                // validaciones de fecha para poder anular
                DateTime fedv = DateTime.Parse(tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                TimeSpan span = DateTime.Parse(lib.fechaServ("ansi")) - fedv;
                if ((span.Days > v_cdpa) || v_clu != tx_dat_loca.Text)
                {
                    // no se puede anular ... a menos que sea un usuario autorizado
                    if (codusanu.Contains(asd))
                    {
                        // SOLO USUARIOS AUTORIZADOS DEBEN ACCEDER A ESTA OPCIÓN
                        // SE ANULA EL DOCUMENTO Y SE HACEN LOS MOVIMIENTOS INTERNOS
                        // LA ANULACION EN EL PROVEEDOR DE FACT. ELECTRONICA SE HACE A MANO POR EL ENCARGADO ... 28/10/2020 ya no al 09/01/2021
                        // la anulacion debe generar un TXT de comunicacion de baja y guardarse en el directorio del prov. de fact. electronica 09/01/2021
                        if (tx_idr.Text.Trim() != "")
                        {
                            var aa = MessageBox.Show("Confirma que desea ANULAR el documento?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (aa == DialogResult.Yes)
                            {
                                if (valiconsql(script) == true)
                                {
                                    int cta = anula("FIS");                             // cantidad de doc.vtas anuladas en la fecha
                                    var xx = MessageBox.Show("Genera la baja electrónica por anulación?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (xx == DialogResult.Yes)
                                    {
                                        anulfactelec(cta);                                  // anulacion electrónica
                                        string resulta = lib.ult_mov(nomform, nomtab, asd); // actualizamos la tabla seguimiento de usuarios
                                        if (resulta != "OK")
                                        {
                                            MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("No existe ruta o no es valida para" + Environment.NewLine +
                                        "generar la anulación electrónica","Ruta para Fact.Electrónica",MessageBoxButtons.OK,MessageBoxIcon.Hand);
                                    iserror = "si";
                                }
                            }
                            else
                            {
                                tx_serie.Focus();
                                return;
                            }
                        }
                        else
                        {
                            MessageBox.Show("El documento ya debe existir para anular", "No esta el Id del registro", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("No se puede anular por estar fuera de plazo o sede","Usuario no permitido",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                    }
                }
                else
                {
                    if (tx_idr.Text.Trim() != "")
                    {
                        var aa = MessageBox.Show("Confirma que desea ANULAR el documento?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (aa == DialogResult.Yes)
                        {
                            if (valiconsql(script) == true)
                            {
                                int cta = anula("FIS");                             // cantidad de doc.vtas anuladas en la fecha
                                var xx = MessageBox.Show("Genera la baja electrónica por anulación?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (xx == DialogResult.Yes)
                                {
                                    anulfactelec(cta);                                  // anulación electrónica
                                    string resulta = lib.ult_mov(nomform, nomtab, asd); // actualizamos la tabla seguimiento de usuarios
                                    if (resulta != "OK")
                                    {
                                        MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                            }
                            else
                            {
                                MessageBox.Show("No existe ruta o no es valida para" + Environment.NewLine +
                                        "generar la anulación electrónica", "Ruta para Fact.Electrónica", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                                iserror = "si";
                            }
                        }
                        else
                        {
                            tx_serie.Focus();
                            return;
                        }
                    }
                    else
                    {
                        MessageBox.Show("El documento ya debe existir para anular", "No esta el Id del registro", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        return;
                    }
                }
            }
            if (iserror == "no")
            {
                string resulta = lib.ult_mov(nomform, nomtab, asd);
                if (resulta != "OK")                                        // actualizamos la tabla usuarios
                {
                    MessageBox.Show(resulta, "Error en actualización de tabla usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                // debe limpiar los campos y actualizar la grilla
            }
            initIngreso();          // limpiamos todo para volver a empesar
        }
        private bool graba()
        {
            bool retorna = false;
            MySqlConnection conn = new MySqlConnection(db_conn_grael);
            conn.Open();
            if(conn.State == ConnectionState.Open)
            {
                if (tx_tipcam.Text == "") tx_tipcam.Text = "0";
                decimal fletMN = 0;
                decimal subtMN = 0;
                decimal igvtMN = 0;
                if (tx_dat_mone.Text != MonDeft)
                {
                    if (tx_tipcam.Text == "0" || tx_fletMN.Text == "")
                    {
                        MessageBox.Show("Error con el tipo de cambio", "Error interno", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return retorna;
                    }
                    else
                    {
                        fletMN = Math.Round(decimal.Parse(tx_fletMN.Text), 3);
                        subtMN = Math.Round(fletMN / (1 + decimal.Parse(v_igv)/100), 3);
                        igvtMN = Math.Round(fletMN - subtMN, 3);
                    }
                }
                else
                {
                    fletMN = Math.Round(decimal.Parse(tx_flete.Text), 3);
                    subtMN = Math.Round(decimal.Parse(tx_subt.Text), 3);
                    igvtMN = Math.Round(decimal.Parse(tx_igv.Text), 3);
                }
                // comprobamos si los datos del cliente tienen cambios
                if (rb_remGR.Checked == true)
                {
                    if (datcltsR[3].ToString().Trim() != tx_dirRem.Text.Trim() ||
                        datcltsR[6].ToString().Trim() != tx_telc1.Text.Trim() ||
                        datcltsR[5].ToString().Trim() != tx_email.Text.Trim() ||
                        datcltsR[4].ToString().Trim() != tx_ubigRtt.Text.Trim())
                    {
                        tx_dat_m1clte.Text = "E";
                    }
                }
                if (rb_desGR.Checked == true)
                {
                    if (datcltsD[3].ToString().Trim() != tx_dirRem.Text.Trim() ||
                        datcltsD[6].ToString().Trim() != tx_telc1.Text.Trim() ||
                        datcltsD[5].ToString().Trim() != tx_email.Text.Trim() ||
                        datcltsD[4].ToString().Trim() != tx_ubigRtt.Text.Trim())
                    {
                        tx_dat_m1clte.Text = "E";
                    }
                }
                string todo = "corre_serie";
                using (MySqlCommand micon = new MySqlCommand(todo, conn))
                {
                    micon.CommandType = CommandType.StoredProcedure;
                    micon.Parameters.AddWithValue("td", cmb_tdv.Text.Trim());  // descrizionerid .. FT, BV
                    micon.Parameters.AddWithValue("ser", tx_serie.Text);
                    using (MySqlDataReader dr0 = micon.ExecuteReader())
                    {
                        if (dr0.Read())
                        {
                            if (dr0[0] != null && dr0.GetString(0) != "")
                            {
                                tx_numero.Text = lib.Right("0000000" + dr0.GetString(0), 7);
                            }
                        }
                    }
                }
                string inserta = "insert into madocvtas (fechope,tipcam,docvta,servta,corvta,doccli,numdcli,direc,nomclie," +
                    "observ,moneda,aumigv,subtot,igv,doctot,status,pigv,userc,fechc,condpag," +
                    "local,rd3,dist,prov,dpto,saldo,cdr,mfe,email,ubiclte,canfild,tippago,plazocred,pagauto," +
                    "prop01,prop02,prop03,prop04,prop05,prop06,prop07,prop08,prop09,prop10) " +
                    "values (" +
                    "@fechop,@tcoper,@ctdvta,@serdv,@numdv,@tdcrem,@ndcrem,@dircre,@nomrem," +
                    "@obsprg,@monppr,@aig,@subpgr,@igvpgr,@totpgr,@estpgr,@porcigv,@asd,now(),@copag," +
                    "@ldcpgr,@tcdvta,@distcl,@provcl,@dptocl,@salxpa,@cdr,@mtdvta,@mailcl,@ubicre,@canfil,@tippa,@plzoc,@pagau," +
                    "@pr01,@pr02,@pr03,@pr04,@pr05,@pr06,@pr07,@pr08,@pr09,@pr10)";
                using (MySqlCommand micon = new MySqlCommand(inserta, conn))
                {
                    micon.Parameters.AddWithValue("@fechop", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                    micon.Parameters.AddWithValue("@tcoper", tx_tipcam.Text);                   // TIPO DE CAMBIO
                    micon.Parameters.AddWithValue("@ctdvta", tx_dat_tdv.Text);
                    micon.Parameters.AddWithValue("@serdv", tx_serie.Text);
                    micon.Parameters.AddWithValue("@numdv", tx_numero.Text);
                    micon.Parameters.AddWithValue("@tdcrem", tx_dat_tdRem.Text);
                    micon.Parameters.AddWithValue("@ndcrem", tx_numDocRem.Text);
                    micon.Parameters.AddWithValue("@dircre", tx_dirRem.Text);
                    micon.Parameters.AddWithValue("@nomrem", tx_nomRem.Text);
                    micon.Parameters.AddWithValue("@obsprg", tx_obser1.Text);
                    micon.Parameters.AddWithValue("@monppr", tx_dat_mone.Text);
                    micon.Parameters.AddWithValue("@aig", "0");  // aumenta igv ??, aca no hay esa opcion
                    micon.Parameters.AddWithValue("@subpgr", tx_subt.Text);                     // sub total
                    micon.Parameters.AddWithValue("@igvpgr", tx_igv.Text);                      // igv
                    micon.Parameters.AddWithValue("@totpgr", tx_flete.Text);                    // total inc. igv
                    micon.Parameters.AddWithValue("@estpgr", (rb_si.Checked == true && tx_idcaja.Text.Trim() != "") ? codCanc : tx_dat_estad.Text); // estado;
                    micon.Parameters.AddWithValue("@porcigv", v_igv);                           // porcentaje en numeros de IGV
                    micon.Parameters.AddWithValue("@asd", asd);
                    micon.Parameters.AddWithValue("@copag", (rb_si.Checked == true)? "E" : (rb_no.Checked == true)? "R" : "C");    // condicion de pago E=contado efectivo, R=contado en reparto, C=crédito
                    micon.Parameters.AddWithValue("@ldcpgr", Grael2.Program.almuser);         // local origen
                    micon.Parameters.AddWithValue("@tcdvta", (rb_remGR.Checked == true) ? "1" : (rb_desGR.Checked == true) ? "2" : "3");
                    micon.Parameters.AddWithValue("@distcl", tx_distRtt.Text);
                    micon.Parameters.AddWithValue("@provcl", tx_provRtt.Text);
                    micon.Parameters.AddWithValue("@dptocl", tx_dptoRtt.Text);
                    micon.Parameters.AddWithValue("@salxpa", (tx_salxcob.Text == "") ? "0" : tx_salxcob.Text);
                    micon.Parameters.AddWithValue("@cdr", "9");  // cdr en la base actual tiene valor 9
                    micon.Parameters.AddWithValue("@mtdvta", cmb_tdv.Text.Substring(0,1));
                    micon.Parameters.AddWithValue("@mailcl", tx_email.Text);
                    micon.Parameters.AddWithValue("@ubicre", tx_ubigRtt.Text);
                    micon.Parameters.AddWithValue("@canfil", tx_tfil.Text);
                    micon.Parameters.AddWithValue("@tippa", tx_dat_plazo.Text);
                    micon.Parameters.AddWithValue("@plzoc", tx_dat_dpla.Text);
                    micon.Parameters.AddWithValue("@pagau", (rb_si.Checked == true)? "S" : "N");    // cobranza automatica en efectivo
                    micon.Parameters.AddWithValue("@pr01", "0");
                    micon.Parameters.AddWithValue("@pr02", "0");
                    micon.Parameters.AddWithValue("@pr03", "0");
                    micon.Parameters.AddWithValue("@pr04", "0");
                    micon.Parameters.AddWithValue("@pr05", "0");
                    micon.Parameters.AddWithValue("@pr06", "0");
                    micon.Parameters.AddWithValue("@pr07", "0");
                    micon.Parameters.AddWithValue("@pr08", "0");
                    micon.Parameters.AddWithValue("@pr09", "0");
                    micon.Parameters.AddWithValue("@pr10", "0");
                    micon.ExecuteNonQuery();
                }
                using (MySqlCommand micon = new MySqlCommand("select last_insert_id()", conn))
                {
                    using (MySqlDataReader dr = micon.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            tx_idr.Text = dr.GetString(0);
                        }
                    }
                }
                // si es factura especial graba en la tb adicionales
                if (rb_fesp.Checked == true)
                {
                    if (tx_idr.Text.Trim() == "")
                    {
                        MessageBox.Show("No existe idr, el registro en adifactu no tendrá enlace" + Environment.NewLine + 
                            "Notifique al área de sistemas inmediatamente","Error!",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    }
                    string insesp = "insert into adifactu (idc,tipoAd,placa,placa2,confv,autoriz,cargaEf,cargaUt,rucTrans,nomTrans,fecIniTras," +
                                    "dirPartida,ubiPartida,dirDestin,ubiDestin,dniChof,brevete,valRefViaje,valRefVehic,valRefTon,precioTN," +
                                    "pesoTN,glosa1,glosa2,glosa3,detMon1,detMon2,detMon3,ruta,valruta,detrac,valrefcu) " +
                                    "values (@idc,@tipo,@pla1,@pla2,@conf,@auto,@carE,@carU,@rucT,@nomT,@fecI," +
                                    "@dirP,@ubiP,@dirD,@ubiD,@dniC,@brev,@vaRV,@vaVe,@valR,@prec," +
                                    "@peTN,@glo1,@glo2,@glo3,@deM1,@deM2,@deM3,@rut,@varu,@detra,@vrcu)";
                    using (MySqlCommand micon = new MySqlCommand(insesp, conn))
                    {
                        micon.Parameters.AddWithValue("@idc", tx_idr.Text);             // id de la cabecera de factura
                        micon.Parameters.AddWithValue("@tipo", 1);                      // 1=carga unica
                        micon.Parameters.AddWithValue("@pla1", tx_e_placa.Text);        // placa del camion
                        micon.Parameters.AddWithValue("@pla2", "");                     // placa de la carreta
                        micon.Parameters.AddWithValue("@conf", tx_e_nfv.Text);          // configuración vehicular
                        micon.Parameters.AddWithValue("@auto", tx_e_aut.Text);          // autorizacion circulación
                        micon.Parameters.AddWithValue("@carE", 0);                      // carga efectiva
                        micon.Parameters.AddWithValue("@carU", tx_e_tn.Text);           // carga util del vehículo
                        micon.Parameters.AddWithValue("@rucT", tx_e_ruct.Text);         // ruc del transportista
                        micon.Parameters.AddWithValue("@nomT", tx_e_ntrans.Text);       // nombre razon social
                        micon.Parameters.AddWithValue("@fecI", tx_e_ftras.Text);        // fecha traslado inicial
                        micon.Parameters.AddWithValue("@dirP", tx_e_dirpar.Text);       // direccion partida
                        micon.Parameters.AddWithValue("@ubiP", tx_e_ubiori.Text);       // ubigeo origen
                        micon.Parameters.AddWithValue("@dirD", tx_e_dirlle.Text);       // dirección de llegada
                        micon.Parameters.AddWithValue("@ubiD", tx_e_ubides.Text);       // ubigeo destino
                        micon.Parameters.AddWithValue("@dniC", tx_e_dnicho.Text);       // dni chofer
                        micon.Parameters.AddWithValue("@brev", "");                     // brevete chofer
                        micon.Parameters.AddWithValue("@vaRV", tx_e_valref.Text);       // valor referencial del viaje
                        micon.Parameters.AddWithValue("@vaVe", 0);                      // valor referencial del vehículo
                        micon.Parameters.AddWithValue("@valR", 0);                      // valor referencial en T.M.
                        micon.Parameters.AddWithValue("@prec", tx_e_prec.Text);         // precio por tonelada
                        micon.Parameters.AddWithValue("@peTN", tx_e_cant.Text);         // peso en toneladas
                        micon.Parameters.AddWithValue("@glo1", tx_e_glos1.Text);        // glosa 1
                        micon.Parameters.AddWithValue("@glo2", tx_e_glos2.Text);        // glosa 2
                        micon.Parameters.AddWithValue("@glo3", tx_e_glos3.Text);        // glosa 3
                        micon.Parameters.AddWithValue("@deM1", 0);                      // monto referencial del servicio de transp
                        micon.Parameters.AddWithValue("@deM2", 0);                      // monto referencial de la carga efectiva
                        micon.Parameters.AddWithValue("@deM3", 0);                      // monto referencial de la carga nominal
                        micon.Parameters.AddWithValue("@rut", tx_dat_ruta.Text);        // ruta del viaje
                        micon.Parameters.AddWithValue("@varu", tx_sxtm.Text);           // valor S/ de la ruta/viaje, se multiplica por TN
                        micon.Parameters.AddWithValue("@detra", tx_detrac.Text);        // detraccion resultante
                        micon.Parameters.AddWithValue("@vrcu", tx_e_carut.Text);        // valor referencial de la carga util
                        micon.ExecuteNonQuery();
                    }
                }
                // detalle
                if (dataGridView1.Rows.Count > 0)
                {
                    int fila = 1;
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString().Trim() != "")
                        {
                            string inserd2 = "insert into detavtas (idc,docvta,servta,corvta,sergr,corgr,moneda," +
                                "valor,ruta,glosa,status,userc,fechc,docremi,bultos,monrefd1,monrefd2,monrefd3,mfe,fecdoc,totaldoc) " +
                                "values (@idr,@doc,@svt,@cvt,@sgui,@cgui,@codm," +
                                "@pret,@ruta,@desc,@sta,@asd,now(),@dre,@bult,@mrd1,@mrd2,@mrd3,@mtdvta,@fechop,@totpgr)";
                            using (MySqlCommand micon = new MySqlCommand(inserd2, conn))
                            {
                                micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                                micon.Parameters.AddWithValue("@doc", tx_dat_tdv.Text);
                                micon.Parameters.AddWithValue("@svt", tx_serie.Text);
                                micon.Parameters.AddWithValue("@cvt", tx_numero.Text);
                                micon.Parameters.AddWithValue("@sgui", dataGridView1.Rows[i].Cells[0].Value.ToString().Trim().Substring(0,3));
                                micon.Parameters.AddWithValue("@cgui", dataGridView1.Rows[i].Cells[0].Value.ToString().Trim().Substring(4,7));
                                micon.Parameters.AddWithValue("@codm", dataGridView1.Rows[i].Cells[3].Value.ToString());
                                micon.Parameters.AddWithValue("@pret", dataGridView1.Rows[i].Cells[4].Value.ToString());
                                micon.Parameters.AddWithValue("@ruta", dataGridView1.Rows[i].Cells[10].Value.ToString());
                                micon.Parameters.AddWithValue("@desc", vint_gg + " " + dataGridView1.Rows[i].Cells[1].Value.ToString());
                                micon.Parameters.AddWithValue("@sta", (rb_si.Checked == true && tx_idcaja.Text.Trim() != "") ? codCanc : tx_dat_estad.Text); // estado;
                                micon.Parameters.AddWithValue("@asd", asd);
                                micon.Parameters.AddWithValue("@dre", dataGridView1.Rows[i].Cells[8].Value.ToString());
                                micon.Parameters.AddWithValue("@bult", dataGridView1.Rows[i].Cells[2].Value.ToString() + " " + dataGridView1.Rows[i].Cells[12].Value.ToString());
                                micon.Parameters.AddWithValue("@mrd1", "0.00"); // (tx_dref1.Text == "") ? "0.00" : tx_dref1.Text
                                micon.Parameters.AddWithValue("@mrd2", "0.00"); // (tx_dcar1.Text == "") ? "0.00" : tx_dcar1.Text
                                micon.Parameters.AddWithValue("@mrd3", "0.00");   // (tx_dnom1.Text == "") ? "0.00" : tx_dnom1.Text
                                micon.Parameters.AddWithValue("@mtdvta", cmb_tdv.Text.Substring(0, 1));
                                micon.Parameters.AddWithValue("@fechop", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                                micon.Parameters.AddWithValue("@totpgr", tx_flete.Text);                    // total inc. igv
                                //micon.Parameters.AddWithValue("@fila", fila);
                                /*
                                micon.Parameters.AddWithValue("@unim", "");
                                micon.Parameters.AddWithValue("@cmnn", dataGridView1.Rows[i].Cells[6].Value.ToString());
                                micon.Parameters.AddWithValue("@tgrmn", dataGridView1.Rows[i].Cells[5].Value.ToString());
                                micon.Parameters.AddWithValue("@pagaut", (rb_si.Checked == true) ? "S" : "N");
                                */
                                micon.ExecuteNonQuery();
                                fila += 1;
                                retorna = true;         // no hubo errores!
                            }
                            // actualizaciones de magrem, manoen y mactacte LOS HACE EL TRIGGER DETAVTAS 
                        }
                    }
                }
                // cobranza automática en efectivo?
                if (rb_si.Checked == true && tx_idcaja.Text.Trim() != "" && double.Parse(tx_salxcob.Text) <= 0)
                {
                    string consulta = "select id,fecha,status from macajas where local=@loc order by id desc limit 1";
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@loc", Grael2.Program.almuser);
                    MySqlDataReader dr = micon.ExecuteReader();
                    if (dr.Read())
                    {
                        if (dr.GetString(0) != tx_idcaja.Text)
                        {
                            tx_idcaja.Text = "";
                            MessageBox.Show("No esta abierta la caja del local!","No se puede cobrar",MessageBoxButtons.OK,MessageBoxIcon.Information);
                        }
                        if (dr.GetDateTime(1).ToString("dd/MM/yyyy") != tx_fechope.Text)
                        {
                            tx_idcaja.Text = "";
                            MessageBox.Show("La fecha del documento no corresponde a la caja", "Caja abierta con fecha anterior", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                        }
                    }
                    else
                    {
                        tx_idcaja.Text = "";
                        MessageBox.Show("Error interno en cobranzas", "No se puede cobrar", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    dr.Close();
                    //
                    if (tx_idcaja.Text != "")
                    {
                        string inscob = "insert into macobran (" +
                            "fechope,tipcam,sercob,local,status,docvta,servta,corvta,doccli,numcli," +
                            "moneda,totdoc,saldo,empcob,fechcob,observ,userc,fechc,mpago,valortc," +
                            "pagado,idcaja,mfe" +
                            ") values (" +
                            "@fec,@tca,@sco,@loc,@sta,@doc,@svt,@cvt,@dcl,@ndc," +
                            "@mon,@tot,@sal,@eco,@fco,@obs,@asd,now(),@mpa,@vtc," +
                            "@pag,@caj,@mfe)";
                        // sernot,cornot
                        // @sno,@cno
                        using (MySqlCommand minsert = new MySqlCommand(inscob, conn))
                        {
                            minsert.Parameters.AddWithValue("@fec", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                            minsert.Parameters.AddWithValue("@tca", tx_tipcam.Text);
                            minsert.Parameters.AddWithValue("@sco", v_sercob);
                            minsert.Parameters.AddWithValue("@loc", v_clu);
                            minsert.Parameters.AddWithValue("@sta", tx_dat_estad.Text);
                            minsert.Parameters.AddWithValue("@doc", tx_dat_tdv.Text);
                            minsert.Parameters.AddWithValue("@svt", tx_serie.Text);
                            minsert.Parameters.AddWithValue("@cvt", tx_numero.Text);
                            minsert.Parameters.AddWithValue("@dcl", tx_dat_tdRem.Text);
                            minsert.Parameters.AddWithValue("@ndc", tx_numDocRem.Text);
                            minsert.Parameters.AddWithValue("@mon", tx_dat_mone.Text);
                            minsert.Parameters.AddWithValue("@tot", tx_flete.Text);
                            minsert.Parameters.AddWithValue("@sal", 0);
                            minsert.Parameters.AddWithValue("@eco", Grael2.Program.codempc);  // lib.codemple(asd)
                            minsert.Parameters.AddWithValue("@fco", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                            minsert.Parameters.AddWithValue("@obs", tx_obser1.Text);
                            minsert.Parameters.AddWithValue("@asd", asd);
                            minsert.Parameters.AddWithValue("@mpa", tx_dat_tipag.Text);
                            if (tx_tipcam.Text != "") minsert.Parameters.AddWithValue("@vtc", double.Parse(tx_flete.Text) * double.Parse(tx_tipcam.Text));
                            else minsert.Parameters.AddWithValue("@vtc", tx_flete.Text);
                            minsert.Parameters.AddWithValue("@pag", tx_flete.Text);
                            minsert.Parameters.AddWithValue("@caj", tx_idcaja.Text);
                            minsert.Parameters.AddWithValue("@mfe", tx_cfe.Text);
                            //minsert.Parameters.AddWithValue("@sno", dataGridView1.Rows[i].Cells[0].Value.ToString().Trim().Substring(0, 3));
                            //minsert.Parameters.AddWithValue("@cno", dataGridView1.Rows[i].Cells[14].Value.ToString());
                            minsert.ExecuteNonQuery();
                            //
                            consulta = "select last_insert_id()";
                            micon = new MySqlCommand(consulta, conn);
                            MySqlDataReader dri = micon.ExecuteReader();
                            if (dri.Read())
                            {
                                tx_idcob.Text = dri.GetString(0);
                            }
                            dri.Close();
                        }
                        /*
                        for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString().Trim() != "")
                            {
                                //
                                consulta = "insert into detacobnot (sercob,corcob,sernot,cornot,valnot,salnot,pagado,userc,fechc,idc) " +
                                    "values (@sco,@cco,@sernot,@cornot,@totnot,0,@totnot,@asd,now(),@idc)";
                                using (micon = new MySqlCommand(consulta, conn))
                                {
                                    micon.Parameters.AddWithValue("@sco", v_sercob);
                                    micon.Parameters.AddWithValue("@cco", lib.Right(("000000" + tx_idcob.Text), 7));
                                    micon.Parameters.AddWithValue("@sernot", dataGridView1.Rows[i].Cells[0].Value.ToString().Trim().Substring(0, 3));
                                    micon.Parameters.AddWithValue("@cornot", dataGridView1.Rows[i].Cells[14].Value.ToString());
                                    micon.Parameters.AddWithValue("@totnot", dataGridView1.Rows[i].Cells[5].Value.ToString());
                                    micon.Parameters.AddWithValue("@asd", asd);
                                    micon.Parameters.AddWithValue("@idc", tx_idcob.Text);
                                    micon.ExecuteNonQuery();
                                }
                            }
                        }
                        */
                        consulta = "insert into detacobnot (sercob,corcob,sernot,cornot,valnot,salnot,pagado,userc,fechc,idc) " +
                            "select @sco,@cco,sernot,cornot,totnot,0,totnot,@asd,now(),@idc from mactacte " +
                            "where tipdv=@tdv and serdv=@svt and cordv=@cvt and mfe=@mfe";
                        using (micon = new MySqlCommand(consulta, conn))
                        {
                            micon.Parameters.AddWithValue("@sco", v_sercob);
                            micon.Parameters.AddWithValue("@cco", lib.Right(("000000" + tx_idcob.Text), 7));
                            micon.Parameters.AddWithValue("@asd", asd);
                            micon.Parameters.AddWithValue("@tdv", tx_dat_tdv.Text);
                            micon.Parameters.AddWithValue("@svt", tx_serie.Text);
                            micon.Parameters.AddWithValue("@cvt", tx_numero.Text);
                            micon.Parameters.AddWithValue("@idc", tx_idcob.Text);
                            micon.Parameters.AddWithValue("@mfe", tx_cfe.Text);
                            micon.ExecuteNonQuery();
                        }
                        //
                        consulta = "update mactacte, manoen set mactacte.status=@nst,mactacte.pagado=totnot,mactacte.saldo=0,mactacte.fecpa=@fpa," +
                            "manoen.status=@nst,manoen.parcial=manoen.doctot,manoen.saldo=0 " +
                            "WHERE mactacte.tipdv = manoen.tdvfac and mactacte.serdv = manoen.serfac and mactacte.cordv = manoen.corfac and mactacte.mfe = manoen.mfe AND " +
                            "mactacte.tipdv = @tdv and mactacte.serdv = @sdv and mactacte.cordv = @cdv and mactacte.mfe = @mfe";
                        using (micon = new MySqlCommand(consulta, conn))
                        {
                            micon.Parameters.AddWithValue("@nst", codCanc);
                            micon.Parameters.AddWithValue("@fpa", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                            micon.Parameters.AddWithValue("@tdv", tx_dat_tdv.Text);
                            micon.Parameters.AddWithValue("@sdv", tx_serie.Text);
                            micon.Parameters.AddWithValue("@cdv", tx_numero.Text);
                            micon.Parameters.AddWithValue("@mfe", tx_cfe.Text);
                            micon.ExecuteNonQuery();
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("No fue posible conectarse al servidor de datos");
                Application.Exit();
                return retorna;
            }
            conn.Close();
            return retorna;
        }
        private void edita()
        {
            MySqlConnection conn = new MySqlConnection(db_conn_grael);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    if (true)     // EDICION DE CABECERA
                    {
                        string actua = "update madocvtas a set observ=@obsprg," +
                            "a.userm=@asd,a.fechm=now() " +
                            "where a.id=@idr";
                        MySqlCommand micon = new MySqlCommand(actua, conn);
                        micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                        micon.Parameters.AddWithValue("@obsprg", tx_obser1.Text);
                        micon.Parameters.AddWithValue("@asd", asd);
                        micon.ExecuteNonQuery();
                        //
                        // EDICION DEL DETALLE .... no hay 
                        micon.Dispose();
                    }
                    conn.Close();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en modificar el documento");
                    Application.Exit();
                    return;
                }
            }
            else
            {
                MessageBox.Show("No fue posible conectarse al servidor de datos");
                Application.Exit();
                return;
            }
        }
        private int anula(string tipo)  // SOLO DEBEN TENER PERMISO LOS USUARIOS DE CONTABILIDAD
        {
            int ctanul = 0;
            // en el caso de documentos de venta HAY 1: ANULACION FISICA 
            // tambien podría haber ANULACION interna con la serie ANU1 .... NO, NO DEBE HABER PORQUE EL CORRELATIVO QUEDARA MAL, HABRA HUECOS EN EL REG. DE VENTAS 
            // Anulacion fisica se "anula" el numero del documento en sistema y en fisico se tacha y en prov. fact.electronica se da baja de numeracion
            // se borran todos los enlaces mediante triggers en la B.D.
            if (tipo == "FIS")
            {
                using (MySqlConnection conn = new MySqlConnection(db_conn_grael))
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        string canul = "update madocvtas a, detavtas b set a.status=@estser,a.observ=@obse,a.usera=@asd,a.fecha=now()," +
                            "b.usera=@asd,b.fecha=now(),b.status=@estser " +
                            "where a.id = b.idc and a.id = @idr";
                        using (MySqlCommand micon = new MySqlCommand(canul, conn))
                        {
                            micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                            micon.Parameters.AddWithValue("@estser", codAnul);
                            micon.Parameters.AddWithValue("@obse", tx_obser1.Text);
                            micon.Parameters.AddWithValue("@asd", asd);
                            micon.ExecuteNonQuery();
                        }
                        //
                        string consul = "select count(id) from madocvtas where date(fecha)=@fech and status=@estser and docvta=@tdv";
                        using (MySqlCommand micon = new MySqlCommand(consul, conn))
                        {
                            micon.Parameters.AddWithValue("@fech", tx_fechact.Text.Substring(6, 4) + "-" + tx_fechact.Text.Substring(3, 2) + "-" + tx_fechact.Text.Substring(0, 2));
                            micon.Parameters.AddWithValue("@estser", codAnul);
                            micon.Parameters.AddWithValue("@tdv", tx_dat_tdv.Text);
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    if (tx_dat_tdv.Text == codbole) ctanul = dr.GetInt32(0) + v_iabol;
                                    else ctanul = dr.GetInt32(0);
                                }
                            }
                        }
                    }
                }
            }
            if (tipo == "INT")      /// esta no se usa 25/08/2021
            {
                using (MySqlConnection conn = new MySqlConnection(db_conn_grael))
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        /*
                        string canul = "update madocvtas set servta=@sain,status=@estser,observ=@obse,usera=@asd,fecha=now() " +
                            "where id=@idr";
                        using (MySqlCommand micon = new MySqlCommand(canul, conn))
                        {
                            micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                            micon.Parameters.AddWithValue("@sain", v_sanu);
                            micon.Parameters.AddWithValue("@estser", codAnul);
                            micon.Parameters.AddWithValue("@obse", tx_obser1.Text);
                            micon.Parameters.AddWithValue("@asd", asd);
                            micon.ExecuteNonQuery();
                        }
                        */
                        // falta actualizar detavtas y con su TRIGGER actualizar mactacte,magrem y manoen
                        /*
                        string updser = "update series set actual=actual-1 where tipdoc=@tipd AND serie=@serd";
                        using (MySqlCommand micon = new MySqlCommand(updser, conn))
                        {
                            micon.Parameters.AddWithValue("@tipd", tx_dat_tdv.Text);
                            micon.Parameters.AddWithValue("@serd", tx_serie.Text);
                            micon.ExecuteNonQuery();
                        }
                        */
                    }
                }
            }
            return ctanul;
        }
        #endregion boton_form;

        #region leaves y checks
        private void tx_idr_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "NUEVO" && tx_idr.Text != "")
            {
                dataGridView1.Rows.Clear();
                string antes = tx_idr.Text;
                initIngreso();
                tx_idr.Text = antes;
                jalaoc("tx_idr");
                jaladet(tx_idr.Text);
            }
        }
        private void tx_nomRem_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_nomRem);
        }
        private void tx_dirRem_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_dirRem);
        }
        private void textBox7_Leave(object sender, EventArgs e)         // departamento del remitente, jala provincia
        {
            if(tx_dptoRtt.Text.Trim() != "")    //  && Grael2.Program.vg_conSol == false
            {
                DataRow[] row = dataUbig.Select("nombre='" + tx_dptoRtt.Text.Trim() + "' and provin='00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_ubigRtt.Text = row[0].ItemArray[1].ToString(); // lib.retCodubigeo(tx_dptoRtt.Text.Trim(),"","");
                    autoprov("cliente");
                }
                else tx_dptoRtt.Text = "";
            }
        }
        private void textBox8_Leave(object sender, EventArgs e)         // provincia del remitente
        {
            if(tx_provRtt.Text != "" && tx_dptoRtt.Text.Trim() != "")   // && Grael2.Program.vg_conSol == false
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and nombre='" + tx_provRtt.Text.Trim() + "' and provin<>'00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_ubigRtt.Text = tx_ubigRtt.Text.Trim().Substring(0, 2) + row[0].ItemArray[2].ToString();
                    autodist("cliente");
                }
                else tx_provRtt.Text = "";
            }
        }
        private void textBox9_Leave(object sender, EventArgs e)         // distrito del remitente
        {
            if(tx_distRtt.Text.Trim() != "" && tx_provRtt.Text.Trim() != "" && tx_dptoRtt.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_ubigRtt.Text.Substring(0, 2) + "' and provin='" + tx_ubigRtt.Text.Substring(2, 2) + "' and nombre='" + tx_distRtt.Text.Trim() + "'");
                if (row.Length > 0)
                {
                    tx_ubigRtt.Text = tx_ubigRtt.Text.Trim().Substring(0, 4) + row[0].ItemArray[3].ToString(); // lib.retCodubigeo(tx_distRtt.Text.Trim(),"",tx_ubigRtt.Text.Trim());
                }
                else tx_distRtt.Text = "";
            }
        }
        private void textBox13_Leave(object sender, EventArgs e)        // ubigeo del remitente
        {
            if(tx_ubigRtt.Text.Trim() != "" && tx_ubigRtt.Text.Length == 6)
            {
                string[] du_remit = lib.retDPDubigeo(tx_ubigRtt.Text);
                tx_dptoRtt.Text = du_remit[0];
                tx_provRtt.Text = du_remit[1];
                tx_distRtt.Text = du_remit[2];
            }
        }
        private void textBox3_Leave(object sender, EventArgs e)         // número de documento del cliente
        {
            if (tx_numDocRem.Text.Trim() != "") //  && tx_mld.Text.Trim() != ""
            {
                validaclt();
            }
            if (tx_numDocRem.Text.Trim() != "" && tx_mld.Text.Trim() == "")
            {
                cmb_docRem.Focus();
            }
        }
        private void tx_numero_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text != "NUEVO" && tx_numero.Text.Trim() != "")
            {
                // en el caso de las pre guias el numero es el mismo que el ID del registro
                rb_si.Checked = false;
                rb_no.Checked = false;
                tx_numero.Text = lib.Right("0000000" + tx_numero.Text, 7);
                if (tx_dat_tdv.Text == "") 
                {
                    cmb_tdv.Focus();
                    return;
                }
                if (tx_serie.Text == "")
                {
                    tx_serie.Focus();
                    return;
                }
                dataGridView1.Rows.Clear();
                jalaoc("sernum");
                jaladet(tx_idr.Text);
            }
        }
        private void tx_serie_Leave(object sender, EventArgs e)
        {
            tx_serie.Text = lib.Right("000" + tx_serie.Text, 3);
            if (Tx_modo.Text == "NUEVO") tx_serGR.Focus();
        }
        private void tx_flete_Leave(object sender, EventArgs e)
        {
            if (tx_flete.Text.Trim() != "" && Tx_modo.Text == "NUEVO")
            {
                tx_flete.Text = Math.Round(decimal.Parse(tx_flete.Text), 2).ToString("#0.00");
                calculos(decimal.Parse((tx_flete.Text.Trim() != "") ? tx_flete.Text : "0"));
                //
                if (tx_dat_mone.Text != MonDeft)
                {
                    if (tx_tipcam.Text == "" || tx_tipcam.Text.Trim() == "0")
                    {
                        MessageBox.Show("Se requiere tipo de cambio");
                        tx_flete.Text = "";
                        tx_flete.Focus();
                        return;
                    }
                    else
                    {
                        tx_fletMN.Text = Math.Round(decimal.Parse(tx_flete.Text) * decimal.Parse(tx_tipcam.Text), 2).ToString();
                        if (Math.Round(decimal.Parse(tx_tfmn.Text),1) != Math.Round(decimal.Parse(tx_fletMN.Text),1))   // OJO, no hacemos dscto en moneda diferente al nacional
                        {
                            MessageBox.Show("No coinciden los valores!","Error en calculo",MessageBoxButtons.OK,MessageBoxIcon.Error);
                            tx_flete.Text = "";
                            tx_flete.Focus();
                            return;
                        }
                    }
                }
                else
                {
                    // si el valor del flete es menor al valor de tx_tfmn ===> tiene descuento
                    // si tiene descuento, visibiliza campo descuento y calcula monto y %
                    if (Math.Round(decimal.Parse(tx_flete.Text), 1) < Math.Round(decimal.Parse(tx_tfmn.Text), 1))
                    {
                        lin_dscto.Visible = true;
                        lb_dscto.Visible = true;
                        tx_valdscto.Visible = true;
                        // calculos
                        tx_valdscto.Text = (Math.Round(decimal.Parse(tx_tfmn.Text), 1) - Math.Round(decimal.Parse(tx_flete.Text), 1)).ToString("#0.0");
                        tx_dat_porcDscto.Text = ((Math.Round(decimal.Parse(tx_flete.Text), 1) * 100) / Math.Round(decimal.Parse(tx_tfmn.Text), 1)).ToString("#0.00");
                    }
                    else
                    {
                        if (Math.Round(decimal.Parse(tx_flete.Text), 1) > Math.Round(decimal.Parse(tx_tfmn.Text), 1))
                        {
                            MessageBox.Show("No se permite facturar montos de las guías","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                            tx_flete.Text = tx_tfmn.Text;
                        }
                        lin_dscto.Visible = false;
                        lb_dscto.Visible = false;
                        tx_valdscto.Visible = false;
                    }
                }
                DataRow[] row = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");
                NumLetra numLetra = new NumLetra();
                tx_fletLetras.Text = numLetra.Convertir(tx_flete.Text,true) + row[0][3].ToString().Trim();
                button1.Focus();
            }
        }
        private void tx_serGR_Leave(object sender, EventArgs e)
        {
            tx_serGR.Text = lib.Right("000" + tx_serGR.Text, 3);
        }
        private void tx_numGR_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" && tx_serGR.Text.Trim() != "" && tx_numGR.Text.Trim() != "")
            {
                tx_numGR.Text = lib.Right("0000000" + tx_numGR.Text, 7);
            }
        }
        private void rb_remGR_Click(object sender, EventArgs e)         // datos del remitente de la GR
        {
            tx_dat_tdRem.Text = datcltsR[0];
            tx_mld.Text = datcltsR[9];
            cmb_docRem.SelectedValue = tx_dat_tdRem.Text;
            tx_numDocRem.Text = datcltsR[1];
            tx_nomRem.Text = datcltsR[2];
            tx_dirRem.Text = datcltsR[3];
            tx_dptoRtt.Text = "";
            tx_provRtt.Text = "";
            tx_distRtt.Text = "";
            if (datcltsR[4].ToString().Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + datcltsR[4].Substring(0, 2) + "' and provin='00' and distri='00'");
                tx_dptoRtt.Text = row[0].ItemArray[4].ToString();
                row = dataUbig.Select("depart='" + datcltsR[4].Substring(0, 2) + "' and provin ='" + datcltsR[4].Substring(2, 2) + "' and distri='00'");
                tx_provRtt.Text = row[0].ItemArray[4].ToString();
                row = dataUbig.Select("depart='" + datcltsR[4].Substring(0, 2) + "' and provin ='" + datcltsR[4].Substring(2, 2) + "' and distri='" + datcltsR[4].Substring(4, 2) + "'");
                tx_distRtt.Text = row[0].ItemArray[4].ToString();
                //
                tx_email.Text = datcltsR[5];
                tx_telc1.Text = datcltsR[6];
                tx_telc2.Text = datcltsR[7];
                tx_ubigRtt.Text = datcltsR[4];
            }
            cmb_docRem.Enabled = false;
            tx_numDocRem.ReadOnly = true;
            tx_nomRem.ReadOnly = true;
            // accionar el tipo de doc de venta acorde al tipo de documento del cliente
            if (tdocsFac.Contains(tx_dat_tdRem.Text)) cmb_tdv.SelectedValue = codfact;
            else cmb_tdv.SelectedValue = codbole;
            validaclt();
        }
        private void rb_desGR_Click(object sender, EventArgs e)         // datos del destinatario de la GR
        {
            tx_dat_tdRem.Text = datcltsD[0];
            tx_mld.Text = datcltsD[9];
            cmb_docRem.SelectedValue = tx_dat_tdRem.Text;
            tx_numDocRem.Text = datcltsD[1];
            tx_nomRem.Text = datcltsD[2];
            tx_dirRem.Text = datcltsD[3];
            tx_dptoRtt.Text = "";
            tx_provRtt.Text = "";
            tx_distRtt.Text = "";
            try
            {
                if (datcltsD[4].ToString().Trim() != "")
                {
                    DataRow[] row = dataUbig.Select("depart='" + datcltsD[4].Substring(0, 2) + "' and provin='00' and distri='00'");
                    tx_dptoRtt.Text = row[0].ItemArray[4].ToString();
                    row = dataUbig.Select("depart='" + datcltsD[4].Substring(0, 2) + "' and provin ='" + datcltsD[4].Substring(2, 2) + "' and distri='00'");
                    tx_provRtt.Text = row[0].ItemArray[4].ToString();
                    row = dataUbig.Select("depart='" + datcltsD[4].Substring(0, 2) + "' and provin ='" + datcltsD[4].Substring(2, 2) + "' and distri='" + datcltsD[4].Substring(4, 2) + "'");
                    tx_distRtt.Text = row[0].ItemArray[4].ToString();
                    //
                    tx_email.Text = datcltsD[5];
                    tx_telc1.Text = datcltsD[6];
                    tx_telc2.Text = datcltsD[7];
                    tx_ubigRtt.Text = datcltsD[4];
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show("Error en datos del Destinatario " + Environment.NewLine + ex.Message, "Error interno", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            //
            cmb_docRem.Enabled = false;
            tx_numDocRem.ReadOnly = true;
            tx_nomRem.ReadOnly = true;
            // accionar el tipo de doc de venta acorde al tipo de documento del cliente
            if (tdocsFac.Contains(tx_dat_tdRem.Text)) cmb_tdv.SelectedValue = codfact;
            else cmb_tdv.SelectedValue = codbole;
            validaclt();
        }
        private void rb_otro_Click(object sender, EventArgs e)
        {
            cmb_docRem.Enabled = true;
            tx_numDocRem.ReadOnly = false;
            tx_nomRem.ReadOnly = false;
            //
            tx_numDocRem.Text = "";
            tx_nomRem.Text = "";
            tx_dirRem.Text = "";
            tx_dptoRtt.Text = "";
            tx_provRtt.Text = "";
            tx_distRtt.Text = "";
            tx_email.Text = "";
            tx_telc1.Text = "";
            tx_telc2.Text = "";
            cmb_docRem.SelectedIndex = 0;
            tx_dat_tdRem.Text = cmb_docRem.SelectedValue.ToString();
            DataRow[] fila = dttd0.Select("idcodice='" + tx_dat_tdRem.Text + "'");
            foreach (DataRow row in fila)
            {
                tx_mld.Text = row[2].ToString();
            }
            cmb_docRem.Focus();
        }
        private void tx_email_Leave(object sender, EventArgs e)
        {
            if (tx_email.Text.Trim() != "")
            {
                if (lib.email_bien_escrito(tx_email.Text.Trim()) == false)
                {
                    MessageBox.Show("El correo electrónico esta mal", "Por favor corrija", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    tx_email.Focus();
                    return;
                }
                if (tx_dat_m1clte.Text != "N") tx_dat_m1clte.Text = "E";
            }
        }
        private void tx_telc1_Leave(object sender, EventArgs e)
        {
            if (tx_telc1.Text.Trim() != "") //  && (Tx_modo.Text == "NUEVO" || Tx_modo.Text == "EDITAR")
            {
                val_NoCaracteres(tx_telc1);
                if (tx_dat_m1clte.Text != "N") tx_dat_m1clte.Text = "E";
            }
        }
        private void tx_obser1_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_obser1);
        }
        private void rb_si_Click(object sender, EventArgs e)            // contado furioso, SE COBRA entra la plata a caja de frente
        {
            cmb_tipop.Enabled = true;
            cmb_tipop.SelectedIndex = 0;        // el primer registro de tipo de pago es el predefinido
            cmb_tipop.Enabled = false;
            //
            tx_pagado.Text = "0.00";
            double once = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[11].Value.ToString())) dataGridView1.Rows[i].Cells[11].Value = "0";
                once = once + double.Parse(dataGridView1.Rows[i].Cells[11].Value.ToString());
            }
            tx_salxcob.Text = once.ToString("#0.00"); // tx_flete.Text;
            tx_salxcob.BackColor = Color.Red;
            //
            cmb_plazoc.SelectedIndex = -1;
            cmb_plazoc.Enabled = false;
            //
            if (tx_idcaja.Text != "" && Tx_modo.Text == "NUEVO")
            {
                // validamos la fecha de la caja
                string fhoy = lib.fechaServ("ansi");
                if (fhoy != Grael2.Program.vg_fcaj)  // ambas fecahs formato yyyy-mm-dd
                {
                    MessageBox.Show("Debe cerrar la caja anterior!", "Caja fuera de fecha", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    rb_si.Checked = false;
                    rb_no.PerformClick();
                    return;
                }
                else
                {
                    if (tx_dat_saldoGR.Text.Trim() != "")
                    {
                        if (tx_dat_mone.Text == MonDeft)
                        {
                            if (decimal.Parse(tx_dat_saldoGR.Text) > 0)
                            {
                                tx_pagado.Text = tx_fletMN.Text;     // tx_flete.Text;
                                tx_salxcob.Text = "0.00";
                                tx_salxcob.BackColor = Color.Green;
                            }
                            else
                            {
                                tx_salxcob.Text = "0.00";
                                tx_dat_plazo.Text = "";
                                cmb_plazoc.SelectedIndex = -1;
                            }
                        }
                        else
                        {
                            MessageBox.Show("Solo puede cancelar en moneda local","Atención",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                            rb_no.PerformClick();
                            return;
                        }
                    }
                    //
                    if (ppauto == "SI")           // automatico o no?
                    {
                        /*
                        cmb_plazoc.Enabled = true;
                        cmb_plazoc.SelectedValue = codppc;
                        DataRow[] dias = dtp.Select("idcodice='" + codppc + "'");
                        tx_dat_diasp.Text = dias[0][3].ToString();
                        */
                    }
                }
            }
            else
            {
                MessageBox.Show("No existe caja abierta!" + Environment.NewLine +
                    "No puede cobrar hasta aperturar caja", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                rb_si.Checked = false;
                rb_no.PerformClick();
            }
        }
        private void rb_no_Click(object sender, EventArgs e)            // contado reparto, se genera contado PERO NO SE COBRA!
        {
            cmb_tipop.SelectedIndex = -1;
            cmb_tipop.Enabled = false;
            //
            tx_pagado.Text = "0.00";
            double once = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[11].Value.ToString())) dataGridView1.Rows[i].Cells[11].Value = "0";
                once = once + double.Parse(dataGridView1.Rows[i].Cells[11].Value.ToString());
            }
            tx_salxcob.Text = once.ToString("#0.00"); // tx_flete.Text;
            tx_salxcob.BackColor = Color.Red;
            //
            cmb_plazoc.SelectedIndex = -1;
            cmb_plazoc.Enabled = false;
        }
        private void rb_cre_Click(object sender, EventArgs e)           // crédito, se cobra cuando paguen, se genera como crédito
        {
            cmb_tipop.SelectedIndex = -1;
            cmb_tipop.Enabled = false;
            cmb_plazoc.Enabled = true;
            tx_dat_plazo.Text = v_mpag;
            //
            double once = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[11].Value.ToString())) dataGridView1.Rows[i].Cells[11].Value = "0";
                once = once + double.Parse(dataGridView1.Rows[i].Cells[11].Value.ToString());
            }
            tx_pagado.Text = "0.00";
            tx_salxcob.Text = once.ToString("#0.00"); // tx_flete.Text;
            tx_salxcob.BackColor = Color.Red;
        }
        private void chk_sinco_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_sinco.Checked == true)
            {
                if (tx_email.Text.Trim() != "") chk_sinco.Checked = false;
                else tx_email.Text = correo_gen;
            }
            else
            {
                if (tx_email.Text.Trim() != "") tx_email.Text = "";
                //else 
            }
        }
        private void val_NoCaracteres(TextBox textBox)
        {
            if (caractNo != "")
            {
                int index = textBox.Text.IndexOf(caractNo);
                if (index > -1)
                {
                    char cno = caractNo.ToCharArray()[0];
                    textBox.Text = textBox.Text.Replace(cno, ' ');
                }
            }
        }
        private void chk_consol_CheckStateChanged(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO")
            {
                if (chk_consol.Checked == false)
                {
                    if (tx_tfil.Text.Trim() != "")
                    {
                        if (int.Parse(tx_tfil.Text) > 1)
                        {
                            var aa = MessageBox.Show("Desea pasar a modo simple?" + Environment.NewLine +
                                "Perdera las guías registradas", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (aa == DialogResult.Yes)
                            {
                                initIngreso();
                                tx_serGR.Focus();
                                // modo simple, maximo 3 guias
                            }
                        }
                    }
                }
                else
                {
                    // modo consolidado, maximo 200 guias
                }
            }
        }
        private void rb_fesp_CheckedChanged(object sender, EventArgs e)                 // habilita campos para facturas especiales de cargas unicas
        {
            if (rb_fesp.Checked == true)
            {
                this.Width = 880;
                panel2.Visible = true;
                panel2.Width = 197;
                panel2.BackColor = Color.FloralWhite;
                // SI ES NUEVO y ya esta "jalada" la guía => pasamos los datos del camión al panel de cargas únicas
                if (Tx_modo.Text == "NUEVO" && dataGridView1.Rows.Count > 1) datPanelCargasU();
                tx_dat_tdv.Text = codfact;
                cmb_tdv.SelectedValue = tx_dat_tdv.Text;
            }
        }
        private void rb_fnor_CheckedChanged(object sender, EventArgs e)
        {
            if (rb_fnor.Checked == true)
            {
                this.Width = 683;
                toolStrip1.Width = 668;
                pn_usloc.Width = 769;
                panel2.Visible = false;
            }

        }
        private void tx_e_tn_Leave(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" && rb_fesp.Checked == true && tx_sxtm.Text.Trim() != "" && tx_e_tn.Text.Trim() != "")
            {
                tx_e_valref.Text = (double.Parse(tx_sxtm.Text) * double.Parse(tx_e_cant.Text)).ToString("#0.00");
                tx_detrac.Text = (double.Parse(tx_e_valref.Text) * double.Parse(Program.pordetra) / 100).ToString("#0.00");
                tx_e_carut.Text = (double.Parse(tx_sxtm.Text) * double.Parse(tx_e_tn.Text)).ToString("#0.00");
            }
        }
        #endregion

        #region botones_de_comando
        public void toolboton()
        {
            DataTable mdtb = new DataTable();
            const string consbot = "select * from permisos where formulario=@nomform and usuario=@use";
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    MySqlCommand consulb = new MySqlCommand(consbot, conn);
                    consulb.Parameters.AddWithValue("@nomform", nomform);
                    consulb.Parameters.AddWithValue("@use", asd);
                    MySqlDataAdapter mab = new MySqlDataAdapter(consulb);
                    mab.Fill(mdtb);
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, " Error ");
                    return;
                }
                finally { conn.Close(); }
            }
            else
            {
                MessageBox.Show("No se pudo conectar con el servidor", "Error de conexión");
                Application.Exit();
                return;
            }
            if (mdtb.Rows.Count > 0)
            {
                DataRow row = mdtb.Rows[0];
                if (Convert.ToString(row["btn1"]) == "S")
                {
                    this.Bt_add.Visible = true;
                }
                else { this.Bt_add.Visible = false; }
                if (Convert.ToString(row["btn2"]) == "S")
                {
                    this.Bt_edit.Visible = true;
                }
                else { this.Bt_edit.Visible = false; }
                //if (Convert.ToString(row["btn5"]) == "S")
                //{
                //    this.Bt_print.Visible = true;
                //}
                //else { this.Bt_print.Visible = false; }
                if (Convert.ToString(row["btn3"]) == "S")
                {
                    this.Bt_anul.Visible = true;
                }
                else { this.Bt_anul.Visible = false; }
                //if (Convert.ToString(row["btn4"]) == "S")
                //{
                //    this.Bt_ver.Visible = true;
                //}
                //else { this.Bt_ver.Visible = false; }
                if (Convert.ToString(row["btn6"]) == "S")
                {
                    this.Bt_close.Visible = true;
                }
                else { this.Bt_close.Visible = false; }
            }
        }
        #region botones
        private void Bt_add_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "NUEVO";
            button1.Image = Image.FromFile(img_grab);
            escribe();
            // 
            Bt_ini.Enabled = false;
            Bt_sig.Enabled = false;
            Bt_ret.Enabled = false;
            Bt_fin.Enabled = false;
            tx_salxcob.BackColor = Color.White;
            // validamos la fecha de la caja
            fshoy = lib.fechaServ("ansi");
            //
            tx_flete.ReadOnly = true;
            initIngreso();
            tx_numero.ReadOnly = true;
            cmb_tdv_SelectedIndexChanged(null, null);
            if (codsuser_cu.Contains(asd)) rb_fesp.Enabled = true;
            else rb_fesp.Enabled = false;
            tx_serGR.Focus();
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            sololee();          
            Tx_modo.Text = "EDITAR";
            button1.Image = Image.FromFile(img_grab);
            tx_flete.ReadOnly = true;
            initIngreso();
            tx_obser1.Enabled = true;
            tx_obser1.ReadOnly = false;
            tx_numero.Text = "";
            tx_numero.ReadOnly = false;
            tx_serie.Focus();
            //
            Bt_ini.Enabled = true;
            Bt_sig.Enabled = true;
            Bt_ret.Enabled = true;
            Bt_fin.Enabled = true;
            tx_salxcob.BackColor = Color.White;
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            if (tx_serie.Text.Trim() != "" && tx_numero.Text.Trim() != "")
            {
                if (tx_impreso.Text == "S") // Impresion ó Re-impresion ??
                {
                    var aa = MessageBox.Show("Desea re imprimir el documento?", "Confirme por favor",
                        MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        if (vi_formato == "A4")            // Seleccion de formato ... A4
                        {
                            if (imprimeA4() == true) updateprint("S");
                        }
                        if (vi_formato == "A5")            // Seleccion de formato ... A5
                        {
                            if (imprimeA5() == true) updateprint("S");
                        }
                        if (vi_formato == "TK")            // Seleccion de formato ... Ticket
                        {
                            if (rb_fnor.Checked == true) if (imprimeTK() == true) updateprint("S");
                            else
                                {
                                    MessageBox.Show("Las facturas especiales no tienen" + Environment.NewLine + 
                                        "formato de impresión en ticket!","Atención",MessageBoxButtons.OK,MessageBoxIcon.Exclamation);
                                    return;
                                }
                        }
                    }
                }
                else
                {
                    if (vi_formato == "A4")            // Seleccion de formato ... A4
                    {
                        if (imprimeA4() == true) updateprint("S");
                    }
                    if (vi_formato == "A5")
                    {
                        if (imprimeA5() == true) updateprint("S");
                    }
                    if (vi_formato == "TK")
                    {
                        if (imprimeTK() == true) updateprint("S");
                    }
                }
            }
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {
            sololee();
            Tx_modo.Text = "ANULAR";
            button1.Image = Image.FromFile(img_anul);
            initIngreso();
            gbox_serie.Enabled = true;
            tx_serie.ReadOnly = false;
            tx_numero.ReadOnly = false;
            tx_obser1.Enabled = true;
            tx_obser1.ReadOnly = false;
            cmb_tdv.Focus();
            //
            Bt_ini.Enabled = true;
            Bt_sig.Enabled = true;
            Bt_ret.Enabled = true;
            Bt_fin.Enabled = true;
        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            sololee();
            Tx_modo.Text = "VISUALIZAR";
            button1.Image = Image.FromFile(img_ver);
            initIngreso();
            gbox_serie.Enabled = true;
            tx_serie.ReadOnly = false;
            tx_numero.ReadOnly = false;
            tx_serie.Focus();
            //
            Bt_ini.Enabled = true;
            Bt_sig.Enabled = true;
            Bt_ret.Enabled = true;
            Bt_fin.Enabled = true;
        }
        private void Bt_first_Click(object sender, EventArgs e)
        {
            limpiar();
            limpia_chk();
            limpia_combos();
            limpia_otros();
            limpia_chk();
            tx_idr.Text = lib.gofirts(nomtab);
            tx_idr_Leave(null, null);
        }
        private void Bt_back_Click(object sender, EventArgs e)
        {
            if(tx_idr.Text.Trim() != "")
            {
                int aca = int.Parse(tx_idr.Text) - 1;
                limpiar();
                limpia_chk();
                limpia_combos();
                limpia_otros();
                tx_idr.Text = aca.ToString();
                tx_idr_Leave(null, null);
            }
        }
        private void Bt_next_Click(object sender, EventArgs e)
        {
            int aca = int.Parse(tx_idr.Text) + 1;
            limpiar();
            limpia_chk();
            limpia_combos();
            limpia_otros();
            tx_idr.Text = aca.ToString();
            tx_idr_Leave(null, null);
        }
        private void Bt_last_Click(object sender, EventArgs e)
        {
            limpiar();
            limpia_chk();
            limpia_combos();
            limpia_otros();
            tx_idr.Text = lib.golast(nomtab);
            tx_idr_Leave(null, null);
        }
        #endregion botones;
        // proveed para habilitar los botones de comando
        #endregion botones_de_comando  ;

        #region comboboxes
        private void cmb_docRem_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (cmb_docRem.SelectedIndex > -1)
            {
                tx_dat_tdRem.Text = cmb_docRem.SelectedValue.ToString();
                DataRow[] fila = dttd0.Select("idcodice='" + tx_dat_tdRem.Text + "'");
                foreach (DataRow row in fila)
                {
                    tx_mld.Text = row[2].ToString();
                }
            }
        }
        private void cmb_tdv_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_tdv.SelectedIndex > -1)
            {
                DataRow[] row = dttd1.Select("idcodice='" + cmb_tdv.SelectedValue.ToString() + "'");
                if (row.Length > 0)
                {
                    tx_dat_tdv.Text = row[0].ItemArray[0].ToString();
                    tx_dat_tdec.Text = row[0].ItemArray[2].ToString();
                    glosser = row[0].ItemArray[4].ToString();
                    tx_cfe.Text = row[0].ItemArray[6].ToString();
                    if (Tx_modo.Text == "NUEVO") tx_serie.Text = row[0].ItemArray[5].ToString();
                    tx_numero.Text = "";
                }
                if (tx_dat_tdv.Text == codbole && rb_fesp.Checked == true && Tx_modo.Text == "NUEVO")
                {
                    MessageBox.Show("Solo las Facturas pueden ser especiales","Atención",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    rb_fnor.Checked = true;
                    return;
                }
            }
        }
        private void cmb_mon_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" && int.Parse(tx_totcant.Text) > 0)    //  || Tx_modo.Text == "EDITAR"
            {   // lo de totcant es para accionar solo cuando el detalle de la GR se haya cargado
                if (cmb_mon.SelectedValue.ToString() != tx_dat_mone.Text) // cmb_mon.SelectedIndex > -1
                {
                    tx_dat_mone.Text = cmb_mon.SelectedValue.ToString();
                    DataRow[] row = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");
                    tx_dat_monsunat.Text = row[0][2].ToString();
                    tipcambio(tx_dat_mone.Text);
                    // hasta aca
                    if (tx_flete.Text != "" && tx_flete.Text != "0.00") calculos(decimal.Parse(tx_flete.Text));
                    if (rb_no.Checked == true) rb_no_Click(null, null);
                    if (rb_si.Checked == true) rb_si_Click(null, null);
                    if (tx_dat_mone.Text != MonDeft)
                    {
                        tx_flete.ReadOnly = false;
                        tx_flete.Focus();
                    }
                    else
                    {
                        if (decimal.Parse(tx_dat_saldoGR.Text) <= 0)
                        {
                            if (cusdscto.Contains(asd)) tx_flete.ReadOnly = false;
                            else tx_flete.ReadOnly = true;
                        }
                        else
                        {
                            tx_flete.ReadOnly = true;
                        }
                    }
                }
            }
        }
        private void cmb_plazoc_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_plazoc.SelectedIndex > -1)
            {
                //tx_dat_plazo.Text = cmb_plazoc.SelectedValue.ToString();  
                //foreach (DataRow row in dias)
                {
                    tx_dat_dpla.Text = dtp.Rows[cmb_plazoc.SelectedIndex].ItemArray[0].ToString();
                    tx_dat_diasp.Text = dtp.Rows[cmb_plazoc.SelectedIndex].ItemArray[3].ToString();
                    //tx_dat_dpla.Text = cmb_plazoc.SelectedValue.ToString();
                    //DataRow[] dias = dtp.Select("idcodice='" + tx_dat_dpla.Text + "'");
                    //tx_dat_diasp.Text  = dias[0][3].ToString();
                }
            }
            else
            {
                //tx_dat_plazo.Text = "";
                tx_dat_dpla.Text = "";
                tx_dat_diasp.Text = "";
            }
        }
        private void cmb_tipop_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_tipop.SelectedIndex > -1)
            {
                tx_dat_tipag.Text = mpa.Rows[cmb_tipop.SelectedIndex].ItemArray[0].ToString();

            }
            else
            {
                tx_dat_tipag.Text = "";
            }
        }
        private void cmb_ruta_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_ruta.SelectedIndex > -1)
            {
                tx_dat_ruta.Text = dtru.Rows[cmb_ruta.SelectedIndex].ItemArray[0].ToString();
                tx_sxtm.Text = dtru.Rows[cmb_ruta.SelectedIndex].ItemArray[2].ToString();
                if (tx_sxtm.Text != "" && tx_e_cant.Text != "")
                {
                    tx_e_valref.Text = (double.Parse(tx_sxtm.Text) * double.Parse(tx_e_cant.Text)).ToString("#0.00");
                    tx_detrac.Text = (double.Parse(tx_e_valref.Text) * double.Parse(Program.pordetra)/100).ToString("#0.00");
                }
                if (tx_e_tn.Text.Trim() == "")
                {
                    tx_e_tn.Focus();
                    return;
                }
                else
                {
                    tx_e_carut.Text = (double.Parse(tx_sxtm.Text) * double.Parse(tx_e_tn.Text)).ToString("#0.00");
                }
            }
            else
            {
                tx_dat_ruta.Text = "";
                tx_sxtm.Text = "";
                tx_detrac.Text = "";
            }
        }
        #endregion comboboxes

        #region impresion
        private bool imprimeA4()
        {
            bool retorna = false;

            return retorna;
        }
        private bool imprimeA5()
        {
            bool retorna = false;
            //llenaDataSet();                         // metemos los datos al dataset de la impresion
            return retorna;
        }
        private bool imprimeTK()
        {
            bool retorna = false;
            //try
            {
                printDocument1.PrinterSettings.PrinterName = v_impTK;
                printDocument1.Print();
                retorna = true;
            }
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message,"Error en imprimir TK");
            //    retorna = false;
            //}
            return retorna;
        }
        private void printDoc_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            if (vi_formato == "A4")
            {
                imprime_A4(sender, e);
            }
            if (vi_formato == "A5")
            {
                imprime_A5(sender, e);
            }
            if (vi_formato == "TK")
            {
                imprime_TK(sender, e);
                if (File.Exists(@otro))
                {
                    //File.Delete(@"C:\test.txt");
                    File.Delete(@otro);
                }

            }
        }
        private void imprime_A4(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {

        }
        private void imprime_A5(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            float alfi = 20.0F;     // alto de cada fila
            float alin = 50.0F;     // alto inicial
            float posi = 80.0F;     // posición de impresión
            float coli = 20.0F;     // columna mas a la izquierda
            float cold = 80.0F;
            Font lt_tit = new Font("Arial", 11);
            Font lt_titB = new Font("Arial", 11, FontStyle.Bold);
            PointF puntoF = new PointF(coli, alin);
            e.Graphics.DrawString(nomclie, lt_titB, Brushes.Black, puntoF, StringFormat.GenericTypographic);                      // titulo del reporte
            posi = posi + alfi;
            string numguia = "GR " + tx_serie.Text + "-" + tx_numero.Text;
            float lt = (lp.CentimeterToPixel(this,21F) - e.Graphics.MeasureString(numguia, lt_titB).Width) / 2;
            puntoF = new PointF(lt, posi);
            e.Graphics.DrawString(numguia, lt_titB, Brushes.Black, puntoF, StringFormat.GenericTypographic);                      // titulo del reporte
            posi = posi + alfi*2;
            PointF ptoimp = new PointF(coli, posi);                     // fecha de emision
            e.Graphics.DrawString("EMITIDO: " + tx_fechope.Text.Substring(0,10), lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
            posi = posi + alfi + 30.0F;                                         // avance de fila
            ptoimp = new PointF(coli, posi);                               // direccion partida
            e.Graphics.DrawString("REMITENTE: " + tx_nomRem.Text.Trim(), lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
            posi = posi + alfi;
            ptoimp = new PointF(coli, posi);                       // destinatario
            posi = posi + alfi * 2;
            /*
            // seleccion de impresion en ruc u otro tipo
            ptoimp = new PointF(coli + 50.0F, posi);
            e.Graphics.DrawString(tx_numDocRem.Text, lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
            ptoimp = new PointF(colm + 185.0F, posi);
            e.Graphics.DrawString(tx_numDocDes.Text, lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
            posi = 330.0F;             // avance de fila
            */
            // detalle de la pre guia
            for (int fila = 0; fila < dataGridView1.Rows.Count - 1; fila++)
            {
                ptoimp = new PointF(coli + 20.0F, posi);
                e.Graphics.DrawString(dataGridView1.Rows[fila].Cells[0].Value.ToString(), lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
                ptoimp = new PointF(cold, posi);
                e.Graphics.DrawString(dataGridView1.Rows[fila].Cells[1].Value.ToString(), lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
                ptoimp = new PointF(cold + 80.0F, posi);
                e.Graphics.DrawString(dataGridView1.Rows[fila].Cells[2].Value.ToString(), lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
                ptoimp = new PointF(cold + 400.0F, posi);
                e.Graphics.DrawString("KGs." + dataGridView1.Rows[fila].Cells[3].Value.ToString(), lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
                posi = posi + alfi;             // avance de fila
            }
            // guias del cliente
            posi = posi + alfi;
            ptoimp = new PointF(coli, posi);
            e.Graphics.DrawString("Docs. de remisión: ", lt_tit, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
            // imprime el flete
            posi = posi + alfi * 2;
            string gtotal = "FLETE " + cmb_mon.Text + " " + tx_flete.Text;
            lt = (lp.CentimeterToPixel(this,21F) - e.Graphics.MeasureString(gtotal, lt_titB).Width) / 2;
            ptoimp = new PointF(lt, posi);
            e.Graphics.DrawString(gtotal, lt_titB, Brushes.Black, ptoimp, StringFormat.GenericTypographic);
            posi = posi + alfi;

        }
        private void imprime_TK(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            {
                // DATOS PARA EL TICKET
                string nomclie = Program.nomclie;
                string rasclie = Program.cliente;
                string rucclie = Program.ruc;
                string dirclie = Program.dirfisc;
                // TIPOS DE LETRA PARA EL DOCUMENTO FORMATO TICKET
                Font lt_gra = new Font("Arial", 11);                // grande
                Font lt_tit = new Font("Lucida Console", 10);       // mediano
                Font lt_med = new Font("Arial", 9);                // normal textos
                Font lt_peq = new Font("Arial", 8);                 // pequeño
                                                                    //
                float anchTik = 7.8F;                               // ancho del TK en centimetros
                int coli = 5;                                      // columna inicial
                float posi = 20;                                    // posicion x,y inicial
                int alfi = 15;                                      // alto de cada fila
                float ancho = 360.0F;                                // ancho de la impresion
                int copias = 1;                                     // cantidad de copias del ticket
                Image photo = Image.FromFile(logoclt);
                for (int i = 1; i <= copias; i++)
                {
                    PointF puntoF = new PointF(coli, posi);
                    // imprimimos el logo o el nombre comercial del emisor
                    if (logoclt != "")
                    {
                        SizeF cuadLogo = new SizeF(CentimeterToPixel(anchTik) - 20.0F, alfi * 6);
                        RectangleF reclogo = new RectangleF(puntoF, cuadLogo);
                        e.Graphics.DrawImage(photo, reclogo);
                    }
                    else
                    {
                        e.Graphics.DrawString(nomclie, lt_gra, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // nombre comercial
                    }
                    float lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(nomclie, lt_gra).Width) / 2;
                    posi = posi + alfi * 7;
                    lt = (ancho - e.Graphics.MeasureString(rasclie, lt_gra).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(rasclie, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // razon social
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Dom.Fiscal", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // direccion emisor
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    SizeF cuad = new SizeF(CentimeterToPixel(anchTik) - (coli + 70), alfi * 2);
                    RectangleF recdom = new RectangleF(puntoF, cuad);
                    e.Graphics.DrawString(dirclie, lt_med, Brushes.Black, recdom, StringFormat.GenericTypographic);     // direccion emisor
                    posi = posi + alfi * 2;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Sucursal", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // direccion emisor
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    cuad = new SizeF(CentimeterToPixel(anchTik) - (coli + 70), alfi * 2);
                    recdom = new RectangleF(puntoF, cuad);
                    e.Graphics.DrawString(dirloc, lt_med, Brushes.Black, recdom, StringFormat.GenericTypographic);     // direccion emisor
                    posi = posi + alfi * 2;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("RUC ", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // ruc de emisor
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    e.Graphics.DrawString(rucclie, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);     // ruc de emisor
                    //string tipdo = cmb_tdv.Text;                                  // tipo de documento
                    string serie = cmb_tdv.Text.Substring(0, 1).ToUpper() + lib.Right(tx_serie.Text,3);                    // serie electrónica
                    string corre = tx_numero.Text;                                // numero del documento electrónico
                    //string nota = tipdo + "-" + serie + "-" + corre;
                    string titdoc = "";
                    if (tx_dat_tdv.Text != codfact) titdoc = "Boleta de Venta Electrónica";
                    if (tx_dat_tdv.Text == codfact) titdoc = "Factura Electrónica";
                    posi = posi + alfi + 8;
                    lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(titdoc, lt_gra).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(titdoc, lt_gra, Brushes.Black, puntoF, StringFormat.GenericTypographic);                  // tipo de documento
                    posi = posi + alfi + 8;
                    string titnum = serie + " - " + corre;
                    lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(titnum, lt_gra).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(titnum, lt_gra, Brushes.Black, puntoF, StringFormat.GenericTypographic);   // serie y numero
                    posi = posi + alfi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("F. Emisión", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic); // fecha y hora emision
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    e.Graphics.DrawString(tx_fechope.Text, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic); // fecha y hora emision
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Cliente", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);                  // DNI/RUC cliente
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    if (tx_nomRem.Text.Trim().Length > 39) cuad = new SizeF(CentimeterToPixel(anchTik) - (coli + 70), alfi * 2);
                    else cuad = new SizeF(CentimeterToPixel(anchTik) - (coli + 70), alfi * 1);
                    recdom = new RectangleF(puntoF, cuad);
                    e.Graphics.DrawString(tx_nomRem.Text.Trim(), lt_peq, Brushes.Black, recdom, StringFormat.GenericTypographic);                  // DNI/RUC cliente
                    if (tx_nomRem.Text.Trim().Length > 39) posi = posi + alfi + alfi;
                    else posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    if (tx_dat_tdv.Text == codfact)
                    {
                        e.Graphics.DrawString("RUC", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);    // nombre del cliente
                    }
                    else
                    {
                        if (tx_dat_tdRem.Text == vtc_dni)
                        {
                            e.Graphics.DrawString("DNI", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);    // nombre del cliente
                        }
                        else
                        {
                            e.Graphics.DrawString("OTROS", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);    // nombre del cliente
                        }
                    }
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    e.Graphics.DrawString(tx_numDocRem.Text, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);    // ruc/dni del cliente
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Dirección", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);  // direccion
                    puntoF = new PointF(coli + 65, posi);
                    e.Graphics.DrawString(":", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    puntoF = new PointF(coli + 70, posi);
                    string dipa = tx_dirRem.Text.Trim() + Environment.NewLine + tx_distRtt.Text.Trim() + " - " + tx_provRtt.Text.Trim() + " - " + tx_dptoRtt.Text.Trim();
                    if (dipa.Length < 60) cuad = new SizeF(CentimeterToPixel(anchTik) - (coli + 70), alfi * 2);
                    else cuad = new SizeF(CentimeterToPixel(anchTik) - (coli + 70), alfi * 3);
                    RectangleF recdir = new RectangleF(puntoF, cuad);
                    e.Graphics.DrawString(tx_dirRem.Text.Trim() + Environment.NewLine +
                        tx_distRtt.Text.Trim() + " - " + tx_provRtt.Text.Trim() + " - " + tx_dptoRtt.Text.Trim(),
                        lt_peq, Brushes.Black, recdir, StringFormat.GenericTypographic);  // direccion
                    if (dipa.Length < 60) posi = posi + alfi + alfi;
                    else posi = posi + alfi + alfi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString(" ", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    // **************** detalle del documento ****************//
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("--------------------------------------------------------------------------", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    StringFormat alder = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                    SizeF siz = new SizeF(70, 15);
                    RectangleF recto = new RectangleF(puntoF, siz);
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Descripción", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Cantidad                    Precio                          Importe", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("--------------------------------------------------------------------------", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    for (int l = 0; l < dataGridView1.Rows.Count - 1; l++)
                    {
                        if (!string.IsNullOrEmpty(dataGridView1.Rows[l].Cells[0].Value.ToString()))
                        {
                            glosser2 = dataGridView1.Rows[l].Cells[10].Value.ToString();
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString(glosser, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString(glosser2, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            //recto = new RectangleF(puntoF, siz);
                            if (Tx_modo.Text == "NUEVO") e.Graphics.DrawString(vint_gg + " " + dataGridView1.Rows[l].Cells[1].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            else e.Graphics.DrawString(dataGridView1.Rows[l].Cells[1].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString("GR Transportista: " + dataGridView1.Rows[l].Cells[0].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            if (dataGridView1.Rows[l].Cells[8].Value.ToString().Trim().Length > 30)
                            {
                                e.Graphics.DrawString("Según doc.cliente: ", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                int alrect = (dataGridView1.Rows[l].Cells[8].Value.ToString().Trim().Length / 30) + 1;
                                siz = new SizeF(180, 15 * alrect);
                                puntoF = new PointF(coli + 100.0F, posi);
                                recto = new RectangleF(puntoF, siz);
                                e.Graphics.DrawString(dataGridView1.Rows[l].Cells[8].Value.ToString(), lt_peq, Brushes.Black, recto, StringFormat.GenericTypographic);
                                posi = posi + alfi * alrect;
                            }
                            else
                            {
                                e.Graphics.DrawString("Según doc.cliente: " + dataGridView1.Rows[l].Cells[8].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                posi = posi + alfi;
                            }
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString("  1", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            puntoF = new PointF(coli + 90.0F, posi);
                            if (tx_dat_tdv.Text == codfact)
                            {
                                e.Graphics.DrawString(dataGridView1.Rows[l].Cells[3].Value.ToString() + " " + dataGridView1.Rows[l].Cells[4].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                siz = new SizeF(70, 15);
                                puntoF = new PointF(coli + 190, posi);
                                recto = new RectangleF(puntoF, siz);
                                e.Graphics.DrawString(dataGridView1.Rows[l].Cells[3].Value.ToString() + " " + dataGridView1.Rows[l].Cells[4].Value.ToString(), lt_peq, Brushes.Black, recto, alder);
                                posi = posi + alfi;
                            }
                            else
                            {
                                if (double.Parse(dataGridView1.Rows[l].Cells[4].Value.ToString()) > double.Parse(tx_flete.Text))
                                {
                                    e.Graphics.DrawString(dataGridView1.Rows[l].Cells[3].Value.ToString() + " " + dataGridView1.Rows[l].Cells[13].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                    siz = new SizeF(70, 15);
                                    puntoF = new PointF(coli + 190, posi);
                                    recto = new RectangleF(puntoF, siz);
                                    e.Graphics.DrawString(dataGridView1.Rows[l].Cells[3].Value.ToString() + " " + dataGridView1.Rows[l].Cells[13].Value.ToString(), lt_peq, Brushes.Black, recto, alder);
                                    posi = posi + alfi;
                                }
                            }
                        }
                    }
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("--------------------------------------------------------------------------", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);

                    // pie del documento ;
                    if (tx_dat_tdv.Text != codfact)
                    {
                        siz = new SizeF(70, 15);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("OP. GRAVADA", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recst = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_subt.Text, lt_peq, Brushes.Black, recst, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("OP. INAFECTA", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recig = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString("0.00", lt_peq, Brushes.Black, recig, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("OP. EXONERADA", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recex = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString("0.00", lt_peq, Brushes.Black, recex, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("IGV", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recgv = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_igv.Text, lt_peq, Brushes.Black, recgv, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("IMPORTE TOTAL " + cmb_mon.Text, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        siz = new SizeF(70, 15);
                        recto = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_flete.Text, lt_peq, Brushes.Black, recto, alder);
                    }
                    if (tx_dat_tdv.Text == codfact)
                    {
                        siz = new SizeF(70, 15);
                        //StringFormat alder = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("OP. GRAVADA", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recst = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_subt.Text, lt_peq, Brushes.Black, recst, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("OP. INAFECTA", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recig = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString("0.00", lt_peq, Brushes.Black, recig, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("OP. EXONERADA", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recex = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString("0.00", lt_peq, Brushes.Black, recex, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("IGV", lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        RectangleF recgv = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_igv.Text, lt_peq, Brushes.Black, recgv, alder);
                        posi = posi + alfi;
                        puntoF = new PointF(coli, posi);
                        e.Graphics.DrawString("IMPORTE TOTAL " + cmb_mon.Text, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        puntoF = new PointF(coli + 190, posi);
                        siz = new SizeF(70, 15);
                        recto = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_flete.Text, lt_peq, Brushes.Black, recto, alder);
                    }
                    posi = posi + alfi * 2;
                    puntoF = new PointF(coli, posi);
                    NumLetra nl = new NumLetra();
                    string monlet = "SON: " + tx_fletLetras.Text;
                    if (monlet.Length <= 30) siz = new SizeF(CentimeterToPixel(anchTik), alfi);
                    else siz = new SizeF(CentimeterToPixel(anchTik), alfi * 2);
                    recto = new RectangleF(puntoF, siz);
                    e.Graphics.DrawString(monlet, lt_peq, Brushes.Black, recto, StringFormat.GenericTypographic);
                    if (monlet.Length <= 30) posi = posi + alfi;
                    else posi = posi + alfi + alfi;
                    if (tx_dat_tdv.Text == codfact)
                    {
                        if (rb_si.Checked == true || rb_no.Checked == true)
                        {
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString(glocopa + " " + texpagC + " " + tx_flete.Text, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                        }
                        else
                        {
                            if (rb_cre.Checked == true)
                            {
                                puntoF = new PointF(coli, posi);
                                e.Graphics.DrawString(glocopa, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                posi = posi + alfi;
                                puntoF = new PointF(coli, posi);
                                DateTime fecpd = Convert.ToDateTime(tx_fechope.Text);
                                string fecpag = fecpd.AddDays(int.Parse(tx_dat_diasp.Text)).ToString("dd'/'MM'/'yyyy");  // ToString("yyyy'-'MM'-'dd")
                                if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra))
                                {
                                    string aaa = (double.Parse(tx_flete.Text) - (double.Parse(tx_flete.Text) * double.Parse(Program.pordetra) / 100)).ToString("#0.00");
                                    e.Graphics.DrawString(texpagD + " " + aaa + " " + texpag2 + " " + texpag3 + fecpag, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                }
                                else
                                {
                                    e.Graphics.DrawString(texpagD + " " + tx_flete.Text + " " + texpag2 + " " + texpag3 + fecpag, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                                }
                            }
                        }
                        if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra))
                        {
                            posi = posi + alfi * 1.5F;
                            siz = new SizeF(CentimeterToPixel(anchTik), 15 * 4);        // detraccion solo facturas .. Richard 07/09/2021
                            puntoF = new PointF(coli, posi);
                            recto = new RectangleF(puntoF, siz);
                            e.Graphics.DrawString(leydet1.Trim() + " "  + leydet2 + " " + Program.ctadetra.Trim(), lt_peq, Brushes.Black, recto, StringFormat.GenericTypographic);
                            posi = posi + alfi * 3;
                        }
                        else
                        {
                            posi = posi + alfi;
                        }
                    }
                    puntoF = new PointF(coli, posi);
                    string repre = "Representación impresa de la";
                    lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(repre, lt_med).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(repre, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    string previo = "";
                    if (tx_dat_tdv.Text != codfact) previo = "boleta de venta electrónica";
                    if (tx_dat_tdv.Text == codfact) previo = "factura electrónica";
                    lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(previo, lt_med).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(previo, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    //posi = posi + alfi;
                    string separ = "|";
                    string codigo = rucclie + separ + tipdo + separ +
                        serie + separ + tx_numero.Text + separ +
                        tx_igv.Text + separ + tx_flete.Text + separ +
                        tx_fechope.Text.Substring(6,4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2) + separ + tipoDocEmi + separ +
                        tx_numDocRem.Text + separ;  // string.Format("{0:yyyy-MM-dd}", tx_fechope.Text)
                    //
                    var rnd = Path.GetRandomFileName();
                    otro = Path.GetFileNameWithoutExtension(rnd);
                    otro = otro + ".png";
                    //
                    var qrEncoder = new QrEncoder(ErrorCorrectionLevel.H);
                    var qrCode = qrEncoder.Encode(codigo);
                    var renderer = new GraphicsRenderer(new FixedModuleSize(5, QuietZoneModules.Two), Brushes.Black, Brushes.White);
                    using (var stream = new FileStream(otro, FileMode.Create))
                        renderer.WriteToStream(qrCode.Matrix, ImageFormat.Png, stream);
                    Bitmap png = new Bitmap(otro);
                    posi = posi + alfi + 7;
                    lt = (CentimeterToPixel(anchTik) - lib.CentimeterToPixel(3)) / 2;
                    puntoF = new PointF(lt, posi);
                    SizeF cuadro = new SizeF(lib.CentimeterToPixel(3), lib.CentimeterToPixel(3));    // 5x5 cm
                    RectangleF rec = new RectangleF(puntoF, cuadro);
                    e.Graphics.DrawImage(png, rec);
                    png.Dispose();
                    // leyenda 2
                    posi = posi + lib.CentimeterToPixel(3);
                    lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(restexto, lt_med).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(restexto, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi;
                    lt = (CentimeterToPixel(anchTik) - e.Graphics.MeasureString(autoriz_bizlinks, lt_med).Width) / 2;
                    puntoF = new PointF(lt, posi);
                    e.Graphics.DrawString(autoriz_bizlinks, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    // centrado en rectangulo   *********************
                    StringFormat sf = new StringFormat();       //  *
                    sf.Alignment = StringAlignment.Center;      //  *
                    posi = posi + alfi + 5;
                    SizeF leyen = new SizeF(CentimeterToPixel(anchTik) - 20, alfi * 3);
                    puntoF = new PointF(coli, posi);
                    leyen = new SizeF(CentimeterToPixel(anchTik) - 20, alfi * 2);
                    RectangleF recley5 = new RectangleF(puntoF, leyen);
                    e.Graphics.DrawString(provee, lt_med, Brushes.Black, recley5, sf);
                    posi = posi + alfi * 3;
                    string locyus = tx_locuser.Text + " - " + tx_user.Text;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString(locyus, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);                  // tienda y vendedor
                    posi = posi + alfi;
                    puntoF = new PointF(coli, posi);
                    e.Graphics.DrawString("Imp. " + DateTime.Now, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi + alfi;
                    puntoF = new PointF((CentimeterToPixel(anchTik) - e.Graphics.MeasureString(despe2, lt_med).Width) / 2, posi);
                    e.Graphics.DrawString(despe2, lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                    posi = posi + alfi + alfi;
                    //puntoF = new PointF(coli, posi);
                    //e.Graphics.DrawString(".", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                }
            }
        }
        private void updateprint(string sn)  // actualiza el campo impreso de la GR = S
        {   // S=si impreso || N=no impreso
            /*
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
            {
                conn.Open();
                string consulta = "update cabfactu set impreso=@sn where id=@idr";
                using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                {
                    micon.Parameters.AddWithValue("@sn", sn);
                    micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                    micon.ExecuteNonQuery();
                }
            }
            */
        }
        #endregion

        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            totalizaG();
        }
    }
}
