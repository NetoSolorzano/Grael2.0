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
        static string nomtab = "cabfactu";              // ????

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
        string v_mpag = "";             // medio de pago automatico x defecto para las cobranzas
        string v_codcob = "";           // codigo del documento cobranza
        string v_CR_gr_ind = "";        // nombre del formato FT/BV en CR
        string v_mfildet = "";          // maximo numero de filas en el detalle, coord. con el formato
        string vint_A0 = "";            // variable codigo anulacion interna por BD
        string v_codidv = "";           // variable codifo interno de documento de venta en vista TDV
        string codfact = "";            // idcodice de factura
        string codbole = "";            // codigo de boleta electronica
        string v_igv = "";              // valor igv %
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
        public string script = "";              // script de conexion a Bizlinks
        NumLetra nl = new NumLetra();
        string[] datcltsR = { "", "", "", "", "", "", "", "", "", "" };
        string[] datcltsD = { "", "", "", "", "", "", "", "", "", "" };
        string[] datguias = { "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" }; // 17
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
            armacfe();
            armadfe();
            if (valiVars() == false)
            {
                Application.Exit();
                return;
            }
        }
        private void init()
        {
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
            tx_dp_dep.AutoCompleteMode = AutoCompleteMode.Suggest;            // punto de partida - cargas unicas
            tx_dp_dep.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tx_dp_dep.AutoCompleteCustomSource = departamentos;
            tx_dp_pro.AutoCompleteMode = AutoCompleteMode.Suggest;
            tx_dp_pro.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tx_dp_pro.AutoCompleteCustomSource = provincias;
            tx_dp_dis.AutoCompleteMode = AutoCompleteMode.Suggest;
            tx_dp_dis.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tx_dp_dis.AutoCompleteCustomSource = distritos;
            tx_dd_dep.AutoCompleteMode = AutoCompleteMode.Suggest;          // punto llegada - cargas unicas
            tx_dd_dep.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tx_dd_dep.AutoCompleteCustomSource = departamentos;
            tx_dd_pro.AutoCompleteMode = AutoCompleteMode.Suggest;
            tx_dd_pro.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tx_dd_pro.AutoCompleteCustomSource = provincias;
            tx_dd_dis.AutoCompleteMode = AutoCompleteMode.Suggest;
            tx_dd_dis.AutoCompleteSource = AutoCompleteSource.CustomSource;
            tx_dd_dis.AutoCompleteCustomSource = distritos;
            // longitudes maximas de campos
            tx_serie.MaxLength = 3;         // serie doc vta
            tx_numero.MaxLength = 7;        // numero doc vta
            tx_serGR.MaxLength = 3;         // serie guia
            tx_numGR.MaxLength = 7;         // numero guia
            tx_numDocRem.MaxLength = 11;    // ruc o dni cliente
            tx_dirRem.MaxLength = 100;
            tx_nomRem.MaxLength = 100;           // nombre remitente
            tx_distRtt.MaxLength = 25;
            tx_provRtt.MaxLength = 25;
            tx_dptoRtt.MaxLength = 25;
            tx_obser1.MaxLength = 150;
            tx_telc1.MaxLength = 12;
            tx_telc2.MaxLength = 12;
            tx_fletLetras.MaxLength = 249;
            tx_dat_dpo.MaxLength = 100;
            tx_dat_dpd.MaxLength = 100;
            tx_pla_placa.MaxLength = 7;
            tx_pla_confv.MaxLength = 15;
            tx_pla_autor.MaxLength = 15;
            // grilla
            dataGridView1.ReadOnly = true;
            dataGridView1.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[3].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            // todo desabilidado
            sololee();
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
                /* if (fshoy != Grael2.Program.vg_fcaj)  // fecha de la caja vs fecha de hoy
                {
                    MessageBox.Show("Las fechas no coinciden" + Environment.NewLine +
                        "Fecha de caja vs Fecha actual", "Caja fuera de fecha", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                    //return;
                }
                else
                {
                    tx_idcaja.Text = v_idcaj;
                } */
            }
            if (Tx_modo.Text == "NUEVO")
            {
                //rb_si.Enabled = true;
                rb_no.Enabled = true;
                if (codsuser_cu.Contains(asd)) chk_cunica.Enabled = true;
                else chk_cunica.Enabled = false;
                if (cusdscto.Contains(asd)) tx_flete.ReadOnly = false;
                else tx_flete.ReadOnly = true;
            }
            tx_dat_nombd.Text = "Bultos";
            tx_dat_nombd.ReadOnly = true;
            cargaunica();
            //pan_pago.Enabled = false;
            rb_no.PerformClick();
        }
        private void jalainfo()                 // obtiene datos de imagenes y variables
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
                            if (row["param"].ToString() == "mpagdef") v_mpag = row["valor"].ToString().Trim();               // medio de pago x defecto para cobranzas
                            if (row["param"].ToString() == "factura") codfact = row["valor"].ToString().Trim();               // codigo doc.venta factura
                            if (row["param"].ToString() == "boleta") codbole = row["valor"].ToString().Trim();               // codigo doc.venta boleta
                            if (row["param"].ToString() == "plazocred") codppc = row["valor"].ToString().Trim();               // codigo plazo de pago x defecto para fact. a CREDITO
                            if (row["param"].ToString() == "usercar_unic") codsuser_cu = row["valor"].ToString().Trim();       // usuarios autorizados a crear Ft de cargas unicas
                            if (row["param"].ToString() == "diasanul") v_cdpa = int.Parse(row["valor"].ToString());            // cant dias en que usuario normal puede anular 
                            if (row["param"].ToString() == "useranul") codusanu = row["valor"].ToString();                      // usuarios autorizados a anular fuera de plazo 
                            if (row["param"].ToString() == "userdscto") cusdscto = row["valor"].ToString();                 // usuarios que pueden hacer descuentos
                            if (row["param"].ToString() == "cltesBol") tdocsBol = row["valor"].ToString();                  // tipos de documento de clientes para boletas
                            if (row["param"].ToString() == "cltesFac") tdocsFac = row["valor"].ToString();                  // tipo de documentos para facturas
                            if (row["param"].ToString() == "limbolsd") limbolsd = decimal.Parse(row["valor"].ToString());                  // limite soles para boletas sin direccion
                        }
                        if (row["campo"].ToString() == "impresion")
                        {
                            if (row["param"].ToString() == "formato") vi_formato = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "filasDet") v_mfildet = row["valor"].ToString().Trim();       // maxima cant de filas de detalle
                            if (row["param"].ToString() == "copias") vi_copias = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "impTK") v_impTK = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "nomfor_cr") v_CR_gr_ind = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "acceso") despe2 = row["valor"].ToString();
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
                            if (row["param"].ToString() == "correo_c1") correo_gen = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "caracterNo") caractNo = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "ose-pse") nipfe = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "motivoBaja") glosaAnul = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "tipsDocbaja") tipdocAnu = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "modTran") texmotran = row["valor"].ToString().Trim();
                            if (row["param"].ToString() == "codmotTran") codtxmotran = row["valor"].ToString().Trim();
                            */
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
                v_slu = lib.serlocs(v_clu);                        // serie local usuario
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
            //try
            {
                string parte = "";
                if (campo == "tx_idr")
                {
                    parte = "where a.id=@ida";
                }
                if (campo == "sernum")
                {
                    parte = "where a.tipdvta=@tdv and a.serdvta=@ser and a.numdvta=@num";
                }
                MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
                conn.Open();
                if (conn.State == ConnectionState.Open)
                {
                    string consulta = "select a.id,a.fechope,a.martdve,a.tipdvta,a.serdvta,a.numdvta,a.ticltgr,a.tidoclt,a.nudoclt,a.nombclt,a.direclt,a.dptoclt,a.provclt,a.distclt,a.ubigclt,a.corrclt,a.teleclt," +
                        "a.locorig,a.dirorig,a.ubiorig,a.obsdvta,a.canfidt,a.canbudt,a.mondvta,a.tcadvta,a.subtota,a.igvtota,a.porcigv,a.totdvta,a.totpags,a.saldvta,a.estdvta,a.frase01,a.impreso," +
                        "a.tipoclt,a.m1clien,a.tippago,a.ferecep,a.userc,a.fechc,a.userm,a.fechm,b.descrizionerid as nomest,ifnull(c.id,'') as cobra,a.idcaja," +
                        "a.cargaunica,a.placa,a.confveh,a.autoriz,a.detPeso,a.detputil,a.detMon1,a.detMon2,a.detMon3,a.dirporig,a.ubiporig,a.dirpdest,a.ubipdest,a.porcendscto,a.valordscto " +
                        "from cabfactu a left join desc_est b on b.idcodice=a.estdvta " +
                        "left join cabcobran c on c.tipdoco=a.tipdvta and c.serdoco=a.serdvta and c.numdoco=a.numdvta and c.estdcob<>@coda "
                        + parte;
                    MySqlCommand micon = new MySqlCommand(consulta, conn);
                    micon.Parameters.AddWithValue("@tdep", vtc_ruc);
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
                            tx_dat_tdv.Text = dr.GetString("tipdvta");
                            tx_serie.Text = dr.GetString("serdvta");
                            tx_numero.Text = dr.GetString("numdvta");
                            rb_remGR.Checked = (dr.GetString("ticltgr") == "1")? true : false;
                            rb_desGR.Checked = (dr.GetString("ticltgr") == "2") ? true : false;
                            rb_otro.Checked = (dr.GetString("ticltgr") == "3") ? true : false;
                            tx_dat_tdRem.Text = dr.GetString("tidoclt");
                            tx_numDocRem.Text = dr.GetString("nudoclt");
                            tx_nomRem.Text = dr.GetString("nombclt");
                            tx_dirRem.Text = dr.GetString("direclt");
                            tx_dptoRtt.Text = dr.GetString("dptoclt");
                            tx_provRtt.Text = dr.GetString("provclt");
                            tx_distRtt.Text = dr.GetString("distclt");
                            tx_ubigRtt.Text = dr.GetString("ubigclt");
                            tx_email.Text = dr.GetString("corrclt");
                            tx_telc1.Text = dr.GetString("teleclt");
                            //locorig,dirorig,ubiorig
                            tx_obser1.Text = dr.GetString("obsdvta");
                            tx_tfil.Text = dr.GetString("canfidt");
                            tx_totcant.Text = dr.GetString("canbudt");  // total bultos
                            tx_dat_mone.Text = dr.GetString("mondvta");
                            tx_tipcam.Text = dr.GetString("tcadvta");
                            tx_subt.Text = Math.Round(dr.GetDecimal("subtota"),2).ToString();
                            tx_igv.Text = Math.Round(dr.GetDecimal("igvtota"), 2).ToString();
                            //,,,porcigv
                            tx_flete.Text = Math.Round(dr.GetDecimal("totdvta"),2).ToString();           // total inc. igv
                            tx_pagado.Text = dr.GetString("totpags");
                            tx_salxcob.Text = dr.GetString("saldvta");
                            tx_dat_estad.Text = dr.GetString("estdvta");        // estado
                            tx_dat_tcr.Text = dr.GetString("tipoclt");          // tipo de cliente credito o contado
                            tx_dat_m1clte.Text = dr.GetString("m1clien");
                            tx_impreso.Text = dr.GetString("impreso");
                            tx_idcob.Text = dr.GetString("cobra");              // id de cobranza
                            //
                            cmb_tdv.SelectedValue = tx_dat_tdv.Text;
                            cmb_tdv_SelectedIndexChanged(null, null);
                            tx_numero.Text = dr.GetString("numdvta");       // al cambiar el indice en el combox se borra numero, por eso lo volvemos a jalar
                            cmb_docRem.SelectedValue = tx_dat_tdRem.Text;
                            cmb_mon.SelectedValue = tx_dat_mone.Text;
                            tx_estado.Text = dr.GetString("nomest");   // lib.nomstat(tx_dat_estad.Text);
                            if (dr.GetString("userm") == "") tx_digit.Text = lib.nomuser(dr.GetString("userc"));
                            else tx_digit.Text = lib.nomuser(dr.GetString("userm"));
                            if (decimal.Parse(tx_salxcob.Text) == decimal.Parse(tx_flete.Text)) rb_no.Checked = true;
                            else rb_si.Checked = true;
                            // campos de carga unica
                            tx_dat_upd.Text = dr.GetString("ubipdest");
                            tx_dat_upo.Text = dr.GetString("ubiporig");
                            tx_dat_dpd.Text = dr.GetString("dirpdest");
                            tx_dat_dpo.Text = dr.GetString("dirporig");
                            tx_valref3.Text = dr.GetString("detMon3");
                            tx_valref2.Text = dr.GetString("detMon2");
                            tx_valref1.Text = dr.GetString("detMon1");
                            tx_cutm.Text = dr.GetString("detputil");
                            tx_cetm.Text = dr.GetString("detPeso");
                            tx_pla_autor.Text = dr.GetString("autoriz");
                            tx_pla_confv.Text = dr.GetString("confveh");
                            tx_pla_placa.Text = dr.GetString("placa");
                            if (dr.GetInt16("cargaunica") == 1) chk_cunica.Checked = true;
                            tx_valdscto.Text = dr.GetString("valordscto");
                            tx_dat_porcDscto.Text = dr.GetString("porcendscto");
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
                }
                conn.Close();
            }
        }
        private void jaladet(string idr)         // jala el detalle
        {
            string jalad = "select filadet,codgror,cantbul,unimedp,descpro,pesogro,codmogr,totalgr " +
                "from detfactu where idc=@idr";
            using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
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
                                row[1].ToString(),
                                row[4].ToString(),
                                row[2].ToString(),
                                row[6].ToString(),
                                row[7].ToString(),
                                "",
                                "",
                                "",
                                "",
                                "");
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
                string consu = "select distinct a.idcodice,a.descrizionerid,a.enlace1,a.codsunat,b.glosaser,b.serie " +
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
                /*      / jalamos la caja
                using (MySqlCommand micon = new MySqlCommand("select id,fechope,statusc from cabccaja where loccaja=@luc order by id desc limit 1", conn))
                {
                    micon.Parameters.AddWithValue("@luc", v_clu);
                    using (MySqlDataReader dr = micon.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            v_estcaj = dr.GetString("statusc");
                            v_idcaj = dr.GetString("id");
                        }
                    }
                }
                */
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
        private bool validGR(string serie, string corre)    // validamos y devolvemos datos
        {
            bool retorna = false;
            if (serie != "" && corre != "")
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
                //
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

                // validamos que la GR: 1.exista, 2.No este facturada, 3.No este anulada
                // y devolvemos una fila con los datos del remitente y otra fila los datos del destinatario
                string hay = "no";
                using (MySqlConnection conn = new MySqlConnection(db_conn_grael))   // DB_CONN_STR
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
                            "ifnull(b2.telef1, '') as numtel1D,ifnull(b2.telef2, '') as numtel2D,a.moneda,a.doctot,a.saldo,SUM(d.cantid) AS bultos, date(a.fechope) as fechopegr,a.tipcam," +
                            "max(d.descrip) AS descrip, ifnull(m.descrizionerid, '') as mon,a.doctot as totgrMN,a.moneda as codMN,c.fecdv,' ' as tipsrem,' ' as tipsdes,a.docremi," +
                            "a.placa,a.carreta,a.cerinsc,a.nfv,concat(lo.descrizionerid, ' - ', ld.descrizionerid) as orides,c.saldo,a.dirorig1 as dirpartida," +
                            "' ' as ubigpartida,a.dirdest1 as dirllegada,' ' as ubigllegada,ifnull(c.fecma, '') as fechplani,a.ruc,ifnull(p.nombre, '') as RazonSocial,dr.flag1 as dr,dd.flag1 as dd " +
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
                                        var a = dr.GetString("fechopegr").Substring(0, 10);
                                        datguias[9] = a.Substring(6,4) + "-" + a.Substring(3,2) + "-" + a.Substring(0,2);     // fecha de la GR
                                        datguias[10] = dr.GetString("docremi");
                                        datguias[11] = dr.GetString("placa"); 
                                        datguias[12] = dr.GetString("carreta");
                                        datguias[13] = dr.GetString("cerinsc");
                                        datguias[14] = dr.GetString("nfv");
                                        datguias[15] = dr.GetString("orides");
                                        datguias[16] = dr.GetString("saldo");
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
        private void tipcambio(string codmod)                // funcion para calculos con el tipo de cambio
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
        int CentimeterToPixel(double Centimeter)
        {
            double pixel = -1;
            using (Graphics g = this.CreateGraphics())
            {
                pixel = Centimeter * g.DpiY / 2.54d;
            }
            return (int)pixel;
        }
        private void cargaunica()               // campos de carga unica
        {
            if (Tx_modo.Text == "NUEVO")
            {
                if (chk_cunica.Checked == true) // .CheckState.ToString() == "True"
                {
                    panel2.Enabled = true;
                    tx_dat_dpo.Enabled = true;
                    tx_dat_dpd.Enabled = true;
                    if (dataGridView1.Rows[0].Cells[0].Value != null)
                    {
                        tx_pla_placa.Text = datguias[11].ToString();
                        tx_pla_confv.Text = datguias[14].ToString();
                        tx_pla_autor.Text = datguias[13].ToString();
                        tx_rucT.Text = datcargu[10].ToString();
                        tx_razonS.Text = datcargu[11].ToString();
                        tx_fecini.Text = (datcargu[12].ToString().Length < 10) ? "" : datcargu[12].ToString().Substring(0, 10);
                        tx_cetm.Text = "";
                        tx_cutm.Text = "";
                        tx_valref1.Text = "";
                        tx_valref2.Text = "";
                        tx_valref3.Text = "";
                        tx_dat_dpo.Text = datcargu[0].ToString();       // datcltsR[3].ToString();
                        tx_dp_dep.Text = datcargu[1].ToString();
                        tx_dp_pro.Text = datcargu[2].ToString();
                        tx_dp_dis.Text = datcargu[3].ToString();
                        tx_dat_upo.Text = datcargu[4].ToString();     // datcltsR[4].ToString();
                        tx_dat_dpd.Text = datcargu[5].ToString();       // datcltsD[3].ToString();
                        tx_dd_dep.Text = datcargu[6].ToString();
                        tx_dd_pro.Text = datcargu[7].ToString();
                        tx_dd_dis.Text = datcargu[8].ToString();
                        tx_dat_upd.Text = datcargu[9].ToString();     // datcltsD[4].ToString();
                    }
                    tx_dat_nombd.ReadOnly = false;
                }
                else
                {
                    panel2.Enabled = false;
                    tx_dat_dpo.Enabled = false;
                    tx_dat_dpd.Enabled = false;
                    //
                    tx_pla_placa.Text = "";
                    tx_pla_confv.Text = "";
                    tx_pla_autor.Text = "";
                    tx_rucT.Text = "";
                    tx_razonS.Text = "";
                    tx_fecini.Text = "";
                    tx_cetm.Text = "";
                    tx_cutm.Text = "";
                    tx_valref1.Text = "";
                    tx_valref2.Text = "";
                    tx_valref3.Text = "";
                    tx_dat_dpo.Text = "";
                    tx_dp_dep.Text = "";
                    tx_dp_pro.Text = "";
                    tx_dp_dis.Text = "";
                    tx_dat_upo.Text = "";
                    tx_dat_dpd.Text = "";
                    tx_dd_dep.Text = "";
                    tx_dd_pro.Text = "";
                    tx_dd_dis.Text = "";
                    tx_dat_upd.Text = "";
                    tx_dat_nombd.Text = "Bultos";
                    tx_dat_nombd.ReadOnly = true;
                }
            }
        }
        private void armacfe()                  // arma cabecera de fact elect.
        {
            tcfe.Clear();
            tcfe.Columns.Add("_fecemi");    // fecha de emision   yyyy-mm-dd
            tcfe.Columns.Add("Prazsoc");    // razon social del emisor
            tcfe.Columns.Add("Pnomcom");    // nombre comercial del emisor
            tcfe.Columns.Add("ubigEmi");    // UBIGEO DOMICILIO FISCAL
            tcfe.Columns.Add("Pdf_dir");    // DOMICILIO FISCAL - direccion
            tcfe.Columns.Add("Pdf_urb");    // DOMICILIO FISCAL - Urbanizacion
            tcfe.Columns.Add("Pdf_pro");    // DOMICILIO FISCAL - provincia
            tcfe.Columns.Add("Pdf_dep");    // DOMICILIO FISCAL - departamento
            tcfe.Columns.Add("Pdf_dis");    // DOMICILIO FISCAL - distrito
            tcfe.Columns.Add("paisEmi");    // DOMICILIO FISCAL - código de país
            tcfe.Columns.Add("Ptelef1");    // teléfono del emisor
            tcfe.Columns.Add("Pweb1");      // página web del emisor
            tcfe.Columns.Add("Prucpro");    // Ruc del emisor
            tcfe.Columns.Add("Pcrupro");    // codigo Ruc emisor
            tcfe.Columns.Add("_tipdoc");    // Tipo de documento de venta - 1 car
            tcfe.Columns.Add("_moneda");    // Moneda del doc. de venta - 3 car
            tcfe.Columns.Add("_sercor");    // Serie y correlat concatenado F001-00000001 - 13 car
            tcfe.Columns.Add("Cnumdoc");    // numero de doc. del cliente - 15 car
            tcfe.Columns.Add("Ctipdoc");    // tipo de doc. del cliente - 1 car
            tcfe.Columns.Add("Cnomcli");    // nombre del cliente - 100 car
            tcfe.Columns.Add("ubigAdq");    // ubigeo del adquiriente - 6 car
            tcfe.Columns.Add("dir1Adq");    // direccion del adquiriente 1
            tcfe.Columns.Add("dir2Adq");    // direccion del adquiriente 2
            tcfe.Columns.Add("provAdq");    // provincia del adquiriente
            tcfe.Columns.Add("depaAdq");    // departamento del adquiriente
            tcfe.Columns.Add("distAdq");    // distrito del adquiriente
            tcfe.Columns.Add("paisAdq");    // pais del adquiriente
            tcfe.Columns.Add("_totoin");    // total operaciones inafectas
            tcfe.Columns.Add("_totoex");    // total operaciones exoneradas
            tcfe.Columns.Add("_toisc");     // total impuesto selectivo consumo
            tcfe.Columns.Add("_totogr");    // Total valor venta operaciones grabadas n(12,2)  15
            tcfe.Columns.Add("_totven");    // Importe total de la venta n(12,2)             15
            tcfe.Columns.Add("tipOper");    // tipo de operacion - 4 car
            tcfe.Columns.Add("codLocE");    // codigo local emisor
            tcfe.Columns.Add("conPago");    // condicion de pago
            tcfe.Columns.Add("plaPago");    // plazo de pago en días
            tcfe.Columns.Add("fvencto");    // fecha de vencimiento de la fact credito yyyy-mm-dd
            tcfe.Columns.Add("_codgui");    // Código de la guia de remision TRANSPORTISTA
            tcfe.Columns.Add("_scotro");    // serie y numero concatenado de la guia
            tcfe.Columns.Add("obser1");     // observacion del documento
            tcfe.Columns.Add("obser2");     // mas observaciones
            tcfe.Columns.Add("maiAdq");     // correo del adquiriente
            tcfe.Columns.Add("teladq");     // telefono del adquiriente
            tcfe.Columns.Add("totImp");     // total impuestos del documento
            tcfe.Columns.Add("codImp");     // codigo impuesto
            tcfe.Columns.Add("nomImp");     // nombre del tipo de impuesto
            tcfe.Columns.Add("tipTri");     // tipo de tributo
            tcfe.Columns.Add("monLet");     // monto en letras
            tcfe.Columns.Add("_horemi");    // hora de emision del doc.venta
            tcfe.Columns.Add("_fvcmto");    // fecha de vencimiento del doc.venta
            tcfe.Columns.Add("corclie");    // correo del emisor
            tcfe.Columns.Add("_morefD");    // moneda de refencia para el tipo de cambio
            tcfe.Columns.Add("_monobj");    // moneda objetivo del tipo de cambio
            tcfe.Columns.Add("_tipcam");    // tipo de cambio con 3 decimales
            tcfe.Columns.Add("_fechca");    // fecha del tipo de cambio
            tcfe.Columns.Add("d_conpa");    // condicion de pago
            tcfe.Columns.Add("d_valre");    // valor referencial
            tcfe.Columns.Add("d_numre");    // numero registro mtc del camion
            tcfe.Columns.Add("d_confv");    // config. vehicular del camion
            tcfe.Columns.Add("d_ptori");    // Pto de origen
            tcfe.Columns.Add("d_ptode");    // Pto de destino
            tcfe.Columns.Add("d_vrepr");    // valor referencial preliminar
            tcfe.Columns.Add("codleyt");    // codigoLeyenda 1 - valor en letras
            tcfe.Columns.Add("codobs");     // codigo del ose para las observaciones, caso carrion documentos origen del remitente
            tcfe.Columns.Add("_forpa");     // glosa de forma de pago SUNAT
            tcfe.Columns.Add("_valcr");     // valor credito
            tcfe.Columns.Add("_fechc");     // fecha programada del pago credito
            // detraccion
            tcfe.Columns.Add("d_porde");                    // 2 Porcentaje de detracción
            tcfe.Columns.Add("d_valde");                    // 3 Monto de la detracción
            tcfe.Columns.Add("d_codse");                    // 4 Código del Bien o Servicio Sujeto a Detracción
            tcfe.Columns.Add("d_ctade");                    // 5 Número del cta en el bco de la nación
            tcfe.Columns.Add("d_medpa");                    // 6 medio de pago de la detraccion (001 = deposito en cuenta)
            tcfe.Columns.Add("glosdet");                    // 7 Leyenda: Detracción        300
            tcfe.Columns.Add("totdet", typeof(double));     // total detraccion
            tcfe.Columns.Add("codleyd");                    // codigo leyenda detraccion
            tcfe.Columns.Add("d_monde");                    // moneda de la detraccion
            // carga unica - traslado de bienes
            tcfe.Columns.Add("cu_cpapp");                   // 02    codigo pais de origen ... osea PE
            tcfe.Columns.Add("cu_ubipp");                   // 03    Ubigeo del punto de partida 
            tcfe.Columns.Add("cu_deppp");                   // 04    Departamento del punto de partida
            tcfe.Columns.Add("cu_propp");                   // 05    Provincia del punto de partida
            tcfe.Columns.Add("cu_dispp");                   // 06    Distrito del punto de partida
            tcfe.Columns.Add("cu_urbpp");                   // 07    Urbanización del punto de partida
            tcfe.Columns.Add("cu_dirpp");                   // 08    Dirección detallada del punto de partida
            tcfe.Columns.Add("cu_cppll");                   // 09    Código país del punto de llegada
            tcfe.Columns.Add("cu_ubpll");                   // 10    Ubigeo del punto de llegada
            tcfe.Columns.Add("cu_depll");                   // 11    Departamento del punto de llegada
            tcfe.Columns.Add("cu_prpll");                   // 12    Provincia del punto de llegada
            tcfe.Columns.Add("cu_dipll");                   // 13    Distrito del punto de llegada
            tcfe.Columns.Add("cu_urpll");                   // 14    Urbanización del punto de llegada
            tcfe.Columns.Add("cu_ddpll");                   // 15    Dirección detallada del punto de llegada
            tcfe.Columns.Add("cu_placa");                   // 16    Placa del Vehículo
            tcfe.Columns.Add("cu_coins");                   // 17    Constancia de inscripción del vehículo o certificado de habilitación vehicular
            tcfe.Columns.Add("cu_marca");                   // 18    Marca del Vehículo
            tcfe.Columns.Add("cu_breve");                   // 19    Nro.de licencia de conducir
            tcfe.Columns.Add("cu_ructr");                   // 20    RUC del transportista
            tcfe.Columns.Add("cu_nomtr");                   // 21    Razón social del Transportista
            tcfe.Columns.Add("cu_modtr");                   // 22    Modalidad de Transporte
            tcfe.Columns.Add("cu_pesbr");                   // 23    Total Peso Bruto
            tcfe.Columns.Add("cu_motra");                   // 24    Código de Motivo de Traslado
            tcfe.Columns.Add("cu_fechi");                   // 25    Fecha de Inicio de Traslado
            tcfe.Columns.Add("cu_remtc");                   // 26    Registro MTC
            tcfe.Columns.Add("cu_nudch");                   // 27    Nro.Documento del conductor
            tcfe.Columns.Add("cu_tidch");                   // 28    Tipo de Documento del conductor
            tcfe.Columns.Add("cu_plac2");                   // 29    Placa del Vehículo secundario
            tcfe.Columns.Add("cu_insub");                   // 30   Indicador de subcontratación
        }
        private void armadfe()                  // arma detalle de fact elect.
        {
            tdfe.Clear();
            tdfe.Columns.Add("Inumord");                    // 2 numero de orden del item           
            tdfe.Columns.Add("Idatper");                    // 3 Datos personilazados del item      
            tdfe.Columns.Add("Iumeded");                    // 4 Unidad de medida                    3
            tdfe.Columns.Add("Icantid");                    // 5 Cantidad de items             n(12,2)
            tdfe.Columns.Add("Idescri");                    // 6 Descripcion                       500
            tdfe.Columns.Add("Idesglo");                    // 7 descricion de la glosa del item   250
            tdfe.Columns.Add("Icodprd");                    // 8 codigo del producto del cliente    30
            tdfe.Columns.Add("Icodpro");                    // 9 codigo del producto SUNAT           8
            tdfe.Columns.Add("Icodgs1");                    // 10 codigo del producto GS1           14
            tdfe.Columns.Add("Icogtin");                    // 11 tipo de producto GTIN             14
            tdfe.Columns.Add("Inplaca");                    // 12 numero placa de vehiculo
            tdfe.Columns.Add("Ivaluni");                    // 13 Valor unitario del item SIN IMPUESTO 
            tdfe.Columns.Add("Ipreuni");                    // 14 Precio de venta unitario CON IGV
            tdfe.Columns.Add("Ivalref");                    // 15 valor referencial del item cuando la venta es gratuita
            tdfe.Columns.Add("_msigv", typeof(double));     // 16 monto igv
            tdfe.Columns.Add("Icatigv");                    // 17 tipo/codigo de afectacion igv
            tdfe.Columns.Add("Itasigv");                    // 18 tasa del igv
            tdfe.Columns.Add("Iigvite");                    // 19 monto IGV del item
            tdfe.Columns.Add("Icodtri");                    // 20 codigo del tributo por item
            tdfe.Columns.Add("Iiscmba");                    // 21 ISC monto base
            tdfe.Columns.Add("Iisctas");                    // 22 ISC tasa del tributo
            tdfe.Columns.Add("Iisctip");                    // 23 ISC tipo de afectacion
            tdfe.Columns.Add("Iiscmon");                    // 24 ISC monto del tributo
            tdfe.Columns.Add("Icbper1");                    // 25 indicador de afecto a ICBPER
            tdfe.Columns.Add("Icbper2");                    // 26 monto unitario de ICBPER
            tdfe.Columns.Add("Icbper3");                    // 27 monto total ICBPER del item
            tdfe.Columns.Add("Iotrtri");                    // 28 otros tributos monto base
            tdfe.Columns.Add("Iotrtas");                    // 29 otros tributos tasa del tributo
            tdfe.Columns.Add("Iotrlin");                    // 30 otros tributos monto unitario
            tdfe.Columns.Add("Itdscto");                    // 31 Descuentos por ítem
            tdfe.Columns.Add("Iincard");                    // 32 indicador de cargo/descuento
            tdfe.Columns.Add("Icodcde");                    // 33 codigo de cargo/descuento
            tdfe.Columns.Add("Ifcades");                    // 34 Factor de cargo/descuento
            tdfe.Columns.Add("Imoncde");                    // 35 Monto de cargo/descuento
            tdfe.Columns.Add("Imobacd");                    // 36 Monto base del cargo/descuento
            tdfe.Columns.Add("Ivalvta");                    // 37 Valor de venta del ítem

            //tdfe.Columns.Add("Iotrsis");                    // otros tributos tipo de sistema
            //tdfe.Columns.Add("Imonbas");                    // monto base (valor sin igv * cantidad)
            //tdfe.Columns.Add("Isumigv");                    // Sumatoria de igv
            //tdfe.Columns.Add("Iindgra");                    // indicador de gratuito
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
        private void grabfactelec()                         // graba en la tabla de fact. electrónicas
        {                               // facturacion electrónica con OSE BIZLINKS 10-10-2018 
            string tipo = tx_dat_tdv.Text;
            string serie = tx_serie.Text;
            string corre = "0" + tx_numero.Text;
            //
            DataRow[] row = dttd1.Select("idcodice='" + tx_dat_tdv.Text + "'");             // tipo de documento venta
            tipdo = row[0][3].ToString();
            //string serie = row[0][1].ToString().Substring(0, 1) + lib.Right(tx_serie.Text, 3);
            DataRow[] rowd = dttd0.Select("idcodice='" + tx_dat_tdRem.Text + "'");          // tipo de documento del cliente
            tipoDocEmi = rowd[0][3].ToString().Trim();
            DataRow[] rowm = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");         // tipo de moneda
            tipoMoneda = rowm[0][2].ToString().Trim();
            // insertamos la cabecera en la tabla del temporal bizlinks
            SqlConnection conms = new SqlConnection(script);
            conms.Open();
            if (conms.State == ConnectionState.Open)
            {
                string sernum = cmb_tdv.Text.Substring(0,1) + tx_serie.Text + "-" + corre;                                     // v serieNumero 
                string fecemi = tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2); // v
                string tipdoc = tipdo;                                                                  // v tipoDocumento
                string tipmon = tipoMoneda;                                                             // v tipoMoneda
                string nudoem = Program.ruc;                                                            // v numeroDocumentoEmisor
                string tidoem = "6";                                                                    // v tipoDocumentoEmisor
                string nocoem = "";                                                                    // v nombreComercialEmisor
                string rasoem = Program.cliente.Trim();                                                 // v razonSocialEmisor
                string coremi = Program.mailclte.Trim();                                               // v correoEmisor
                string coloem = Program.codlocsunat;                                                    // v codigoLocalAnexoEmisor
                string ubiemi = Program.ubidirfis;                                                      // v ubigeoEmisor
                string diremi = Program.dirfisc;                                                        // v direccionEmisor
                string provemi = Program.provfis;                                                     // v provinciaEmisor
                string depaemi = Program.depfisc;                                                     // v departamentoEmisor
                string distemi = Program.distfis;                                                     // v distritoEmisor
                string pasiemi = "PE";                                                                  // v paisEmisor
                string tidoad = tipoDocEmi;                                                             // v tipoDocumentoAdquiriente
                string nudoad = tx_numDocRem.Text;                                                      // v numeroDocumentoAdquiriente
                string rasoad = tx_nomRem.Text;                                                         // v razonSocialAdquiriente
                string coradq = tx_email.Text;                                                         // v correoAdquiriente
                decimal totimp = Math.Round(decimal.Parse(tx_igv.Text), 2);                              // v totalImpuestos
                decimal totigv = Math.Round(decimal.Parse(tx_igv.Text), 2);                              // v totalIgv
                decimal totvta = Math.Round(decimal.Parse(tx_flete.Text), 2);                            // v totalVenta
                decimal tovane = Math.Round(decimal.Parse(tx_subt.Text), 2);                            // v totalValorVentaNetoOpGravadas
                //string tovano = "0";                                                                    // totalValorVentaNetoOpNoGravada
                //string tovaex = "0";                                                                    // totalValorVentaNetoOpExoneradas
                //string tovagr = "0";                                                                    // totalValorVentaNetoOpGratuitas
                string codaux40_1 = "9011";                                                             // v codigoAuxiliar40_1
                string etiaux40_1 = "18%";                                                              // v textoAuxiliar40_1
                string tipope = "0101"; // segun rudver, poner esto en una config                       // v tipoOperacion
                //
                string tiref1 = "";                                                                     // v tipoReferencia_1
                string nudor1 = "";                                                                     // v numeroDocumentoReferencia_1
                string estreg = "A";                                                                    // bl_estadoRegistro
                string tiref2 = "";
                string nudor2 = "";
                string tiref3 = "";
                string nudor3 = "";
                string tiref4 = "";
                string nudor4 = "";
                for (int z = 0; z < dataGridView1.Rows.Count - 1; z++)
                {
                    switch (z)
                    {
                        case 0:
                            tiref1 = "31";
                            nudor1 = dataGridView1.Rows[z].Cells[0].Value.ToString(); // tx_grs1.Text + "-" + tx_grc1.Text;
                            break;
                        case 1:
                            tiref2 = "31";                                                                      // tipoReferencia_2
                            nudor2 = dataGridView1.Rows[z].Cells[0].Value.ToString();                           // numeroDocumentoReferencia_2
                            break;
                        case 2:
                            tiref3 = "31";                                                                      // tipoReferencia_3
                            nudor3 = dataGridView1.Rows[z].Cells[0].Value.ToString();                           // numeroDocumentoReferencia_3
                            break;
                        case 3:
                            tiref4 = "31";                                                                      // tipoReferencia_4
                            nudor4 = dataGridView1.Rows[z].Cells[0].Value.ToString();                           // numeroDocumentoReferencia_4
                            break;
                    }
                }
                string coley1 = "1000";                                                                 // v codigoLeyenda_1
                string teley1 = "SON: " + tx_fletLetras.Text;
                //string teley1 = "SON: " + nl.Convertir(tx_flete.Text, true) + tx_dat_dmon.Text.Trim();  // v textoLeyenda_1
                //string canglo = tx_cant.Text;                                                           // cantidad que se imprime en la fact
                //decimal preglo = Math.Round(decimal.Parse(tx_precio.Text), 2);                          // precio que se imprime en la fact
                //string tx1glo = tx_glosa1.Text;                                                         // detalle 1 de la fact
                //string tx2glo = tx_glosa2.Text;                                                         // detalle 2 de la fact
                //string tx3glo = tx_glosa3.Text;                                                         // detalle 3 de la fact
                // ***************** EN OTROS CASOS VAN LAS GUIAS DEL CLIENTE
                string insertcab = "insert into SPE_EINVOICEHEADER (serieNumero,fechaEmision,tipoDocumento,tipoMoneda," +
                    "numeroDocumentoEmisor,tipoDocumentoEmisor,nombreComercialEmisor,razonSocialEmisor,correoEmisor,codigoLocalAnexoEmisor," +
                    "ubigeoEmisor,direccionEmisor,provinciaEmisor,departamentoEmisor,distritoEmisor,paisEmisor,codigoAuxiliar40_1,textoAuxiliar40_1," +
                    "tipoDocumentoAdquiriente,numeroDocumentoAdquiriente,razonSocialAdquiriente,correoAdquiriente,totalImpuestos," +
                    "totalValorVentaNetoOpGravadas,codigoLeyenda_1,textoLeyenda_1,bl_estadoRegistro," +
                    "totalIgv,totalVenta,tipoOperacion,tipoReferencia_1,numeroDocumentoReferencia_1";
                if (!string.IsNullOrEmpty(nudor2) && !string.IsNullOrWhiteSpace(nudor2))
                {
                    insertcab = insertcab + ",tipoReferencia_2,numeroDocumentoReferencia_2";
                }
                if (!string.IsNullOrEmpty(nudor3) && !string.IsNullOrWhiteSpace(nudor3))
                {
                    insertcab = insertcab + ",tipoReferencia_3,numeroDocumentoReferencia_3";
                }
                if (!string.IsNullOrEmpty(nudor4) && !string.IsNullOrWhiteSpace(nudor4))
                {
                    insertcab = insertcab + ",tipoReferencia_4,numeroDocumentoReferencia_4";
                }
                // *********************   calculo y campos de detracciones   ******************************
                // Están sujetos a las detracciones los servicios de transporte de bienes por vía terrestre gravado con el IGV, 
                // siempre que el importe de la operación o el valor referencial, según corresponda, sea mayor a 
                // S/ 400.00 o su equivalente en dólares ........ DICE SUNAT
                // ctadetra            // cuenta detraccion del emisor
                // valdetra            // monto detraccion
                double totdet = 0;
                string leydet = leydet1 + " " + leydet2 + " " + Program.ctadetra;                   // textoLeyenda_2
                // codleyt                                                                  // codigoLeyenda_2
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)    // soles
                    {
                        // ctadetra;                                                            // numeroCtaBancoNacion
                        // valdetra;                                                            // monto a partir del cual tiene detraccion la operacion
                        // coddetra;                                                            // codigoDetraccion
                        // pordetra;                                                            // porcentajeDetraccion
                        totdet = double.Parse(tx_flete.Text) * double.Parse(Program.pordetra) / 100;                  // totalDetraccion
                        insertcab = insertcab + ",codigoDetraccion,totalDetraccion,porcentajeDetraccion,numeroCtaBancoNacion,codigoLeyenda_2,textoLeyenda_2";
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)    // dolares
                    {
                        // ctadetra;                                                            // numeroCtaBancoNacion
                        // valdetra;                                                            // monto a partir del cual tiene detraccion la operacion
                        // coddetra;                                                            // codigoDetraccion
                        // pordetra;                                                            // porcentajeDetraccion
                        totdet = double.Parse(tx_flete.Text) * double.Parse(Program.pordetra) / 100;                  // totalDetraccion
                        insertcab = insertcab + ",codigoDetraccion,totalDetraccion,porcentajeDetraccion,numeroCtaBancoNacion,codigoLeyenda_2,textoLeyenda_2";
                    }
                }
                insertcab = insertcab + ") " +
                    "values (@sernum,@fecemi,@tipdoc,@tipmon," +
                    "@nudoem,@tidoem,@nocoem,@rasoem,@coremi,@coloem," +
                    "@ubiemi,@diremi,@provemi,@depaemi,@distemi,@pasiemi,@codaux40_1,@etiaux40_1," +
                    "@tidoad,@nudoad,@rasoad,@coradq,@totimp," +
                    "@tovane,@coley1,@teley1,@estreg," +
                    "@totigv,@totvta,@tipope,@tiref1,@nudor1";
                if (!string.IsNullOrEmpty(nudor2) && !string.IsNullOrWhiteSpace(nudor2)) insertcab = insertcab + ",@tiref2,@nudor2";
                if (!string.IsNullOrEmpty(nudor3) && !string.IsNullOrWhiteSpace(nudor3)) insertcab = insertcab + ",@tiref3,@nudor3";
                if (!string.IsNullOrEmpty(nudor4) && !string.IsNullOrWhiteSpace(nudor4)) insertcab = insertcab + ",@tiref4,@nudor4";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        insertcab = insertcab + ",@coddetra,@totdet,@pordetra,@ctadetra,@codleyt,@leydet";
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        insertcab = insertcab + ",@coddetra,@totdet,@pordetra,@ctadetra,@codleyt,@leydet";
                    }
                }
                insertcab = insertcab + ")";
                // primero insertamos el detalle, luego la cabecera
                // ....
                // *****************************   detalle   *****************************
                string ubiPtoO1 = "";
                string dirPtoO1 = "";
                string ubiPtoD1 = "";
                string dirPtoD1 = "";
                string detViaj1 = "";
                string pesoD1 = "";
                string monRefS1 = "1.00";
                string monRefC1 = "1.00";
                string monRefU1 = "1.00";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO1 = tx_dser1.Text;           // ubigeoPtoOrigen
                        dirPtoO1 = tx_dcor1.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD1 = tx_ubdes1.Text;          // ubigeoPtoDestino
                        dirPtoD1 = tx_drut1.Text;           // direccionCompletaPtoDestino
                        detViaj1 = tx_dd1.Text;             // detalleViaje
                        pesoD1 = tx_peso1.Text;             // 
                        if (tx_dref1.Text.Trim() == "" || tx_dref1.Text == "0" || tx_dref1.Text == "0.00")
                        {
                            monRefS1 = "1.00";           // montoRefServicioTransporte
                            monRefC1 = "1.00";           // montoRefCargaEfectiva
                            monRefU1 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS1 = tx_dref1.Text;           // montoRefServicioTransporte
                            monRefC1 = tx_dcar1.Text;           // montoRefCargaEfectiva
                            monRefU1 = tx_dnom1.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO1 = tx_dser1.Text;           // ubigeoPtoOrigen
                        dirPtoO1 = tx_dcor1.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD1 = tx_ubdes1.Text;          // ubigeoPtoDestino
                        dirPtoD1 = tx_drut1.Text;           // direccionCompletaPtoDestino
                        detViaj1 = tx_dd1.Text;             // detalleViaje
                        pesoD1 = tx_peso1.Text;             // 
                        if (tx_dref1.Text.Trim() == "" || tx_dref1.Text == "0" || tx_dref1.Text == "0.00")
                        {
                            monRefS1 = "1.00";           // montoRefServicioTransporte
                            monRefC1 = "1.00";           // montoRefCargaEfectiva
                            monRefU1 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS1 = tx_dref1.Text;           // montoRefServicioTransporte
                            monRefC1 = tx_dcar1.Text;           // montoRefCargaEfectiva
                            monRefU1 = tx_dnom1.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                string nuori1 = "1";                                                                    // numeroOrdenItem
                string codprd1 = "-";                                                                   // codigoProducto
                string coprsu1 = "78101802";                                                            // codigoProductoSunat
                string descr1 = glosser + " / " + dataGridView1.Rows[0].Cells[10].Value.ToString().Trim() + " / " +
                    dataGridView1.Rows[0].Cells[1].Value.ToString().Trim() + " / " + dataGridView1.Rows[0].Cells[2].Value.ToString().Trim() + 
                    " S/G: " + dataGridView1.Rows[0].Cells[8].Value.ToString().Trim();         // descripcion
                decimal canti1 = Math.Round(decimal.Parse("1"), 2);
                string unime1 = "ZZ";                                                                   // unidadMedida
                decimal psi1, igv1;                                                                     // calculos de precios x item sin y con impuestos
                double inuns1 = 0;
                if (decimal.TryParse(dataGridView1.Rows[0].Cells[4].Value.ToString(), out psi1))
                {
                    inuns1 = Math.Round(((double)psi1 / ((double)decimal.Parse(v_igv) / 100 + 1)), 2);   // importeUnitarioSinImpuesto
                }
                else { inuns1 = 0; }                                                                    // importeUnitarioSinImpuesto
                decimal inunc1 = Math.Round(decimal.Parse(dataGridView1.Rows[0].Cells[4].Value.ToString()), 2);                              // importeUnitarioConImpuesto
                string coimu1 = "01";                                                                   // codigoImporteUnitarioConImpues
                string imtoi1 = "";
                if (decimal.TryParse(dataGridView1.Rows[0].Cells[4].Value.ToString(), out igv1)) { imtoi1 = Math.Round(((double)igv1 - ((double)igv1 / ((double)decimal.Parse(v_igv) / 100 + 1))), 2).ToString(); }
                else { imtoi1 = "0.00"; }                                                               // importeTotalImpuestos 
                double mobai1 = inuns1;                                                                 // montoBaseIgv
                string taigv1 = ((decimal.Parse(v_igv))).ToString();                                    // tasaIgv
                string imigv1 = imtoi1;                                                                 // importeIgv
                string corae1 = "10";   // grabado operacion onerosa                                    // codigoRazonExoneracion
                double intos1 = inuns1;                                                                 // importeTotalSinImpuesto
                //
                string ubiPtoO2 = "";
                string dirPtoO2 = "";
                string ubiPtoD2 = "";
                string dirPtoD2 = "";
                string detViaj2 = "";
                string pesoD2 = "";
                string monRefS2 = "1.00";
                string monRefC2 = "1.00";
                string monRefU2 = "1.00";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO2 = tx_dser2.Text;           // ubigeoPtoOrigen
                        dirPtoO2 = tx_dcor2.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD2 = tx_ubdes2.Text;          // ubigeoPtoDestino
                        dirPtoD2 = tx_drut2.Text;           // direccionCompletaPtoDestino
                        detViaj2 = tx_dd2.Text;             // detalleViaje
                        pesoD2 = tx_peso2.Text;             // 
                        if (tx_dref2.Text.Trim() == "" || tx_dref2.Text == "0" || tx_dref2.Text == "0.00")
                        {
                            monRefS2 = "1.00";           // montoRefServicioTransporte
                            monRefC2 = "1.00";           // montoRefCargaEfectiva
                            monRefU2 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS2 = tx_dref2.Text;           // montoRefServicioTransporte
                            monRefC2 = tx_dcar2.Text;           // montoRefCargaEfectiva
                            monRefU2 = tx_dnom2.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO2 = tx_dser2.Text;           // ubigeoPtoOrigen
                        dirPtoO2 = tx_dcor2.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD2 = tx_ubdes2.Text;          // ubigeoPtoDestino
                        dirPtoD2 = tx_drut2.Text;           // direccionCompletaPtoDestino
                        detViaj2 = tx_dd2.Text;             // detalleViaje
                        pesoD2 = tx_peso2.Text;             // 
                        if (tx_dref2.Text.Trim() == "" || tx_dref2.Text == "0" || tx_dref2.Text == "0.00")
                        {
                            monRefS2 = "1.00";           // montoRefServicioTransporte
                            monRefC2 = "1.00";           // montoRefCargaEfectiva
                            monRefU2 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS2 = tx_dref2.Text;           // montoRefServicioTransporte
                            monRefC2 = tx_dcar2.Text;           // montoRefCargaEfectiva
                            monRefU2 = tx_dnom2.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                string nuori2 = "";
                string codprd2 = "";
                string coprsu2 = "";
                string descr2 = "";
                decimal canti2 = 0;
                string unime2 = "";
                double intos2 = 0;
                decimal inunc2 = 0;
                string coimu2 = "";
                double mobai2 = 0;
                string taigv2 = "";
                string imigv2 = "";
                string imtoi2 = "";
                string corae2 = "";
                double inuns2 = 0;
                if (dataGridView1.Rows[1].Cells[0].Value != null && dataGridView1.Rows[1].Cells[0].Value.ToString().Trim() != "")
                {
                    nuori2 = "2";                                                                    // numeroOrdenItem
                    codprd2 = "-";                                                                   // codigoProducto
                    coprsu2 = "78101802";                                                            // codigoProductoSunat
                    descr2 = glosser + " / " + dataGridView1.Rows[1].Cells[10].Value.ToString().Trim() + " / " +
                        dataGridView1.Rows[1].Cells[1].Value.ToString().Trim() + " / " + dataGridView1.Rows[1].Cells[2].Value.ToString().Trim() + 
                        " S/G: " + dataGridView1.Rows[1].Cells[8].Value.ToString().Trim();           // descripcion
                    canti2 = Math.Round(decimal.Parse("1"), 2);
                    unime2 = "ZZ";                                                                   // unidadMedida
                    decimal psi2, igv2;                                                              // calculos de precios x item sin y con impuestos
                    inuns2 = 0;
                    if (decimal.TryParse(dataGridView1.Rows[1].Cells[4].Value.ToString(), out psi2))
                    {
                        inuns2 = Math.Round(((double)psi2 / ((double)decimal.Parse(v_igv) / 100 + 1)), 2);   // importeUnitarioSinImpuesto
                    }
                    else { inuns2 = 0; }                                                                     // importeUnitarioSinImpuesto
                    inunc2 = Math.Round(decimal.Parse(dataGridView1.Rows[1].Cells[4].Value.ToString()), 2);                               // importeUnitarioConImpuesto
                    coimu2 = "01";                                                                   // codigoImporteUnitarioConImpues
                    imtoi2 = "";
                    if (decimal.TryParse(dataGridView1.Rows[1].Cells[4].Value.ToString(), out igv2)) { imtoi2 = Math.Round(((double)igv2 - ((double)igv2 / ((double)decimal.Parse(v_igv) / 100 + 1))), 2).ToString(); }
                    else { imtoi2 = "0.00"; }                                                               // importeTotalImpuestos 
                    mobai2 = inuns2;                                                                 // montoBaseIgv
                    taigv2 = ((decimal.Parse(v_igv))).ToString();                              // tasaIgv
                    imigv2 = imtoi2;                                                                 // importeIgv
                    corae2 = "10";   // grabado operacion onerosa                                    // codigoRazonExoneracion
                    intos2 = inuns2;                                                                 // importeTotalSinImpuest
                }
                //
                string ubiPtoO3 = "";
                string dirPtoO3 = "";
                string ubiPtoD3 = "";
                string dirPtoD3 = "";
                string detViaj3 = "";
                string pesoD3 = "";
                string monRefS3 = "1.00";
                string monRefC3 = "1.00";
                string monRefU3 = "1.00";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO3 = tx_dser3.Text;           // ubigeoPtoOrigen
                        dirPtoO3 = tx_dcor3.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD3 = tx_ubdes3.Text;          // ubigeoPtoDestino
                        dirPtoD3 = tx_drut3.Text;           // direccionCompletaPtoDestino
                        detViaj3 = tx_dd3.Text;             // detalleViaje
                        pesoD3 = tx_peso3.Text;             // 
                        if (tx_dref3.Text.Trim() == "" || tx_dref3.Text == "0" || tx_dref3.Text == "0.00")
                        {
                            monRefS3 = "1.00";           // montoRefServicioTransporte
                            monRefC3 = "1.00";           // montoRefCargaEfectiva
                            monRefU3 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS3 = tx_dref3.Text;           // montoRefServicioTransporte
                            monRefC3 = tx_dcar3.Text;           // montoRefCargaEfectiva
                            monRefU3 = tx_dnom3.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO3 = tx_dser3.Text;           // ubigeoPtoOrigen
                        dirPtoO3 = tx_dcor3.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD3 = tx_ubdes3.Text;          // ubigeoPtoDestino
                        dirPtoD3 = tx_drut3.Text;           // direccionCompletaPtoDestino
                        detViaj3 = tx_dd3.Text;             // detalleViaje
                        pesoD3 = tx_peso3.Text;             // 
                        if (tx_dref3.Text.Trim() == "" || tx_dref3.Text == "0" || tx_dref3.Text == "0.00")
                        {
                            monRefS3 = "1.00";           // montoRefServicioTransporte
                            monRefC3 = "1.00";           // montoRefCargaEfectiva
                            monRefU3 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS3 = tx_dref3.Text;           // montoRefServicioTransporte
                            monRefC3 = tx_dcar3.Text;           // montoRefCargaEfectiva
                            monRefU3 = tx_dnom3.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                string nuori3 = "";
                string codprd3 = "";
                string coprsu3 = "";
                string descr3 = "";
                decimal canti3 = 0;
                string unime3 = "";
                double intos3 = 0;
                decimal inunc3 = 0;
                string coimu3 = "";
                double mobai3 = 0;
                string taigv3 = "";
                string imigv3 = "";
                string imtoi3 = "";
                string corae3 = "";
                double inuns3 = 0;
                if (!string.IsNullOrEmpty(dataGridView1.Rows[2].Cells[0].Value.ToString()))
                {
                    nuori3 = "3";                                                                    // numeroOrdenItem
                    codprd3 = "-";                                                                   // codigoProducto
                    coprsu3 = "78101802";                                                            // codigoProductoSunat
                    descr3 = glosser + " / " + dataGridView1.Rows[2].Cells[10].Value.ToString().Trim() + " / " +
                        dataGridView1.Rows[2].Cells[1].Value.ToString().Trim() + " / " + dataGridView1.Rows[2].Cells[2].Value.ToString().Trim() + 
                        " S/G: " + dataGridView1.Rows[2].Cells[8].Value.ToString();  // descripcion
                    canti3 = Math.Round(decimal.Parse("1"), 2);
                    unime3 = "ZZ";                                                                   // unidadMedida
                    decimal psi3, igv3;             // calculos de precios x item sin y con impuestos
                    inuns3 = 0;
                    if (decimal.TryParse(dataGridView1.Rows[2].Cells[4].Value.ToString(), out psi3))
                    {
                        inuns3 = Math.Round(((double)psi3 / ((double)decimal.Parse(v_igv) / 100 + 1)), 2);   // importeUnitarioSinImpuesto
                    }
                    else { inuns3 = 0; }                                                                  // importeUnitarioSinImpuesto
                    inunc3 = Math.Round(decimal.Parse(dataGridView1.Rows[2].Cells[4].Value.ToString()), 2);                            // importeUnitarioConImpuesto
                    coimu3 = "01";                                                                   // codigoImporteUnitarioConImpues
                    imtoi3 = "";
                    if (decimal.TryParse(dataGridView1.Rows[2].Cells[4].Value.ToString(), out igv3)) { imtoi3 = ((double)igv3 - ((double)igv3 / ((double)decimal.Parse(v_igv) / 100 + 1))).ToString(); }
                    else { imtoi3 = "0.00"; }                                                               // importeTotalImpuestos 
                    mobai3 = inuns3;                                                                 // montoBaseIgv
                    taigv3 = ((decimal.Parse(v_igv))).ToString();                                    // tasaIgv
                    imigv3 = imtoi3;                                                                 // importeIgv
                    corae3 = "10";   // grabado operacion onerosa                                    // codigoRazonExoneracion
                    intos3 = inuns3;                                                                 // importeTotalSinImpuesto
                }
                //
                string ubiPtoO4 = "";
                string dirPtoO4 = "";
                string ubiPtoD4 = "";
                string dirPtoD4 = "";
                string detViaj4 = "";
                string pesoD4 = "";
                string monRefS4 = "1.00";
                string monRefC4 = "1.00";
                string monRefU4 = "1.00";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO4 = tx_dser4.Text;           // ubigeoPtoOrigen
                        dirPtoO4 = tx_dcor4.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD4 = tx_ubdes4.Text;          // ubigeoPtoDestino
                        dirPtoD4 = tx_drut4.Text;           // direccionCompletaPtoDestino
                        detViaj4 = tx_dd4.Text;             // detalleViaje
                        pesoD4 = tx_peso4.Text;             // 
                        if (tx_dref4.Text.Trim() == "" || tx_dref4.Text == "0" || tx_dref4.Text == "0.00")
                        {
                            monRefS4 = "1.00";           // montoRefServicioTransporte
                            monRefC4 = "1.00";           // montoRefCargaEfectiva
                            monRefU4 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS4 = tx_dref4.Text;           // montoRefServicioTransporte
                            monRefC4 = tx_dcar4.Text;           // montoRefCargaEfectiva
                            monRefU4 = tx_dnom4.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        tipope = "1004";                    // tipo de operacion cuando tiene detraccion
                        /*
                        ubiPtoO4 = tx_dser4.Text;           // ubigeoPtoOrigen
                        dirPtoO4 = tx_dcor4.Text;           // direccionCompletaPtoOrigen
                        ubiPtoD4 = tx_ubdes4.Text;          // ubigeoPtoDestino
                        dirPtoD4 = tx_drut4.Text;           // direccionCompletaPtoDestino
                        detViaj4 = tx_dd4.Text;             // detalleViaje
                        pesoD4 = tx_peso4.Text;             // 
                        if (tx_dref4.Text.Trim() == "" || tx_dref4.Text == "0" || tx_dref4.Text == "0.00")
                        {
                            monRefS4 = "1.00";           // montoRefServicioTransporte
                            monRefC4 = "1.00";           // montoRefCargaEfectiva
                            monRefU4 = "1.00";           // montoRefCargaUtilNominal
                        }
                        else
                        {
                            monRefS4 = tx_dref4.Text;           // montoRefServicioTransporte
                            monRefC4 = tx_dcar4.Text;           // montoRefCargaEfectiva
                            monRefU4 = tx_dnom4.Text;           // montoRefCargaUtilNominal
                        }
                        */
                    }
                }
                string nuori4 = "";
                string codprd4 = "";
                string coprsu4 = "";
                string descr4 = "";
                decimal canti4 = 0;
                string unime4 = "";
                double intos4 = 0;
                decimal inunc4 = 0;
                string coimu4 = "";
                double mobai4 = 0;
                string taigv4 = "";
                string imigv4 = "";
                string imtoi4 = "";
                string corae4 = "";
                double inuns4 = 0;
                if (!string.IsNullOrEmpty(dataGridView1.Rows[3].Cells[0].Value.ToString()))
                {
                    nuori4 = "4";                                                                    // numeroOrdenItem
                    codprd4 = "-";                                                                   // codigoProducto
                    coprsu4 = "78101802";                                                            // codigoProductoSunat
                    descr4 = glosser + " / " + dataGridView1.Rows[3].Cells[10].Value.ToString().Trim() + " / " +
                        dataGridView1.Rows[3].Cells[1].Value.ToString().Trim() + " / " + dataGridView1.Rows[3].Cells[2].Value.ToString().Trim() + 
                        " S/G: " + dataGridView1.Rows[3].Cells[8].Value.ToString().Trim();                              // descripcion
                    canti4 = Math.Round(decimal.Parse("1"), 2);
                    unime4 = "ZZ";                                                                   // unidadMedida
                    decimal psi4, igv4;                                                              // calculos de precios x item sin y con impuestos
                    inuns4 = 0;
                    if (decimal.TryParse(dataGridView1.Rows[3].Cells[4].Value.ToString(), out psi4))
                    {
                        inuns4 = Math.Round(((double)psi4 / ((double)decimal.Parse(v_igv) / 100 + 1)), 2);   // importeUnitarioSinImpuesto
                    }
                    else { inuns4 = 0; }                                                             // importeUnitarioSinImpuesto
                    inunc4 = Math.Round(decimal.Parse(dataGridView1.Rows[3].Cells[4].Value.ToString()), 2);                              // importeUnitarioConImpuesto
                    coimu4 = "01";                                                                   // codigoImporteUnitarioConImpues
                    imtoi4 = "";
                    if (decimal.TryParse(dataGridView1.Rows[3].Cells[4].Value.ToString(), out igv4)) { imtoi4 = ((double)igv4 - ((double)igv4 / ((double)decimal.Parse(v_igv) / 100 + 1))).ToString(); }
                    else { imtoi4 = "0.00"; }                                                        // importeTotalImpuestos 
                    mobai4 = inuns4;                                                                 // montoBaseIgv
                    taigv4 = ((decimal.Parse(v_igv))).ToString();                                    // tasaIgv
                    imigv4 = imtoi4;                                                                 // importeIgv
                    corae4 = "10";   // grabado operacion onerosa                                    // codigoRazonExoneracion
                    intos4 = inuns4;                                                                 // importeTotalSinImpuesto
                }
                // ******************************************************************************************************************************************************
                string insertadet = "insert into SPE_EINVOICEDETAIL (tipoDocumentoEmisor,numeroDocumentoEmisor,tipoDocumento,serieNumero," +
                    "numeroOrdenItem,codigoProducto,codigoProductoSunat,descripcion,cantidad,unidadMedida,importeTotalSinImpuesto," +
                    "importeUnitarioSinImpuesto,importeUnitarioConImpuesto,codigoImporteUnitarioConImpues,montoBaseIgv,tasaIgv," +
                    "importeIgv,importeTotalImpuestos,codigoRazonExoneracion";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        insertadet = insertadet + ",ubigeoPtoOrigen,direccionCompletaPtoOrigen,ubigeoPtoDestino,direccionCompletaPtoDestino," +
                                "detalleViaje,montoRefServicioTransporte,montoRefCargaEfectiva,montoRefCargaUtilNominal";
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        insertadet = insertadet + ",ubigeoPtoOrigen,direccionCompletaPtoOrigen,ubigeoPtoDestino,direccionCompletaPtoDestino," +
                                "detalleViaje,montoRefServicioTransporte,montoRefCargaEfectiva,montoRefCargaUtilNominal";
                    }
                }
                insertadet = insertadet + ") " +
                    "values (@tidoem,@nudoem,@tipdoc,@sernum," +
                    "@nuori1,@codprd1,@coprsu1,@descr1,@canti1,@unime1,@intos1," +
                    "@inuns1,@inunc1,@coimu1,@mobai1,@taigv1," +
                    "@imigv1,@imtoi1,@corae1";
                if (tx_dat_mone.Text == MonDeft)
                {
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                    {
                        insertadet = insertadet + ",@ubiPtoO1,@dirPtoO1,@ubiPtoD1,@dirPtoD1," +
                                "@detViaj1,@monRefS1,@monRefC1,@monRefU1"; // pesoD1,
                    }
                }
                else
                {
                    if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                    {
                        insertadet = insertadet + ",@ubiPtoO1,@dirPtoO1,@ubiPtoD1,@dirPtoD1," +
                                "@detViaj1,@monRefS1,@monRefC1,@monRefU1"; // pesoD1,
                    }
                }
                insertadet = insertadet + ")";
                // debemos insertar las 4 lineas de un solo porrazo, con if vamos viendo si agregamos la linea
                if (dataGridView1.Rows[1].Cells[0].Value.ToString() != "")
                {
                    insertadet = insertadet + ",(@t2idoem,@n2udoem,@t2ipdoc,@s2ernum," +
                        "@n2uori2,@c2odprd2,@c2oprsu2,@d2escr2,@c2anti2,@u2nime2,@i2ntos2," +
                        "@i2nuns2,@i2nunc2,@c2oimu2,@m2obai2,@t2aigv2," +
                        "@i2migv2,@i2mtoi2,@c2orae2";
                    if (tx_dat_mone.Text == MonDeft)
                    {
                        if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                        {
                            insertadet = insertadet + ",@u2biPtoO2,@d2irPtoO2,@u2biPtoD2,@d2irPtoD2," +
                                "@d2etViaj2,@m2onRefS2,@m2onRefC2,@m2onRefU2"; // pesoD2,
                        }
                    }
                    else
                    {
                        if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                        {
                            insertadet = insertadet + ",@u2biPtoO2,@d2irPtoO2,@u2biPtoD2,@d2irPtoD2," +
                                "@d2etViaj2,@m2onRefS2,@m2onRefC2,@m2onRefU2"; // pesoD2,
                        }
                    }
                    insertadet = insertadet + ")";
                }
                if (dataGridView1.Rows[2].Cells[0].Value.ToString() != "")
                {
                    insertadet = insertadet + ",(@t3idoem,@n3udoem,@t3ipdoc,@s3ernum," +
                        "@n3uori3,@c3odprd3,@c3oprsu3,@d3escr3,@c3anti3,@u3nime3,@i3ntos3," +
                        "@i3nuns3,@i3nunc3,@c3oimu3,@m3obai3,@t3aigv3," +
                        "@i3migv3,@i3mtoi3,@c3orae3";
                    if (tx_dat_mone.Text == MonDeft)
                    {
                        if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                        {
                            insertadet = insertadet + ",@u3biPtoO3,@d3irPtoO3,@u3biPtoD3,@d3irPtoD3," +
                                "@d3etViaj3,@m3onRefS3,@m3onRefC3,@m3onRefU3"; // pesoD3,
                        }
                    }
                    else
                    {
                        if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                        {
                            insertadet = insertadet + ",@u3biPtoO3,@d3irPtoO3,@u3biPtoD3,@d3irPtoD3," +
                                "@d3etViaj3,@m3onRefS3,@m3onRefC3,@m3onRefU3"; // pesoD3,
                        }
                    }
                    insertadet = insertadet + ")";
                }
                if (dataGridView1.Rows[3].Cells[0].Value.ToString() != "")
                {
                    insertadet = insertadet + ",(@t4idoem,@n4udoem,@t4ipdoc,@s4ernum," +
                        "@n4uori4,@c4odprd4,@c4oprsu4,@d4escr4,@c4anti4,@u4nime4,@i4ntos4," +
                        "@i4nuns4,@i4nunc4,@c4oimu4,@m4obai4,@t4aigv4," +
                        "@i4migv4,@i4mtoi4,@c4orae4";
                    if (tx_dat_mone.Text == MonDeft)
                    {
                        if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                        {
                            insertadet = insertadet + ",@u4biPtoO4,@d4irPtoO4,@u4biPtoD4,@d4irPtoD4," +
                                "@d4etViaj4,@m4onRefS4,@m4onRefC4,@m4onRefU4"; // pesoD4,
                        }
                    }
                    else
                    {
                        if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                        {
                            insertadet = insertadet + ",@u4biPtoO4,@d4irPtoO4,@u4biPtoD4,@d4irPtoD4," +
                                "@d4etViaj4,@m4onRefS4,@m4onRefC4,@m4onRefU4"; // pesoD4,
                        }
                    }
                    insertadet = insertadet + ")";
                }
                try
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
                    indet.Parameters.AddWithValue("@unime1", unime1);       // unidad medida
                    indet.Parameters.AddWithValue("@intos1", intos1);       // ImporteTotalSinImpuesto
                    indet.Parameters.AddWithValue("@inuns1", inuns1);       // ImporteUnitarioSinImpuesto
                    indet.Parameters.AddWithValue("@inunc1", inunc1);       // ImporteUnitarioConImpuesto
                    indet.Parameters.AddWithValue("@coimu1", coimu1);       // codifoImporteUnitarioConImpuesto
                    indet.Parameters.AddWithValue("@mobai1", mobai1);       // montoBaseIgv
                    indet.Parameters.AddWithValue("@taigv1", taigv1);       // tasaIgv
                    indet.Parameters.AddWithValue("@imigv1", imigv1);       // importeIgv
                    indet.Parameters.AddWithValue("@imtoi1", imtoi1);       // importeTotalImpuestos
                    indet.Parameters.AddWithValue("@corae1", corae1);       // codigoRazonExo
                    if (tx_dat_mone.Text == MonDeft)
                    {
                        if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                        {
                            indet.Parameters.AddWithValue("@ubiPtoO1", ubiPtoO1);
                            indet.Parameters.AddWithValue("@dirPtoO1", dirPtoO1);
                            indet.Parameters.AddWithValue("@ubiPtoD1", ubiPtoD1);
                            indet.Parameters.AddWithValue("@dirPtoD1", dirPtoD1);
                            indet.Parameters.AddWithValue("@detViaj1", detViaj1);
                            indet.Parameters.AddWithValue("@monRefS1", monRefS1);
                            indet.Parameters.AddWithValue("@monRefC1", monRefC1);
                            indet.Parameters.AddWithValue("@monRefU1", monRefU1);
                        }
                    }
                    else
                    {
                        if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                        {
                            indet.Parameters.AddWithValue("@ubiPtoO1", ubiPtoO1);
                            indet.Parameters.AddWithValue("@dirPtoO1", dirPtoO1);
                            indet.Parameters.AddWithValue("@ubiPtoD1", ubiPtoD1);
                            indet.Parameters.AddWithValue("@dirPtoD1", dirPtoD1);
                            indet.Parameters.AddWithValue("@detViaj1", detViaj1);
                            indet.Parameters.AddWithValue("@monRefS1", monRefS1);
                            indet.Parameters.AddWithValue("@monRefC1", monRefC1);
                            indet.Parameters.AddWithValue("@monRefU1", monRefU1);
                        }
                    }
                    if (dataGridView1.Rows[1].Cells[0].Value.ToString() != "")                                  // tiene 2 lineas de detalle
                    {
                        indet.Parameters.AddWithValue("@t2idoem", tidoem);
                        indet.Parameters.AddWithValue("@n2udoem", nudoem);
                        indet.Parameters.AddWithValue("@t2ipdoc", tipdoc);
                        indet.Parameters.AddWithValue("@s2ernum", sernum);
                        indet.Parameters.AddWithValue("@n2uori2", nuori2);
                        indet.Parameters.AddWithValue("@c2odprd2", codprd2);     // codigo producto
                        indet.Parameters.AddWithValue("@c2oprsu2", coprsu2);     // codigo producto SUNAT
                        indet.Parameters.AddWithValue("@d2escr2", descr2);
                        indet.Parameters.AddWithValue("@c2anti2", canti2);
                        indet.Parameters.AddWithValue("@u2nime2", unime2);
                        indet.Parameters.AddWithValue("@i2ntos2", intos2);
                        indet.Parameters.AddWithValue("@i2nuns2", inuns2);
                        indet.Parameters.AddWithValue("@i2nunc2", inunc2);       // 
                        indet.Parameters.AddWithValue("@c2oimu2", coimu2);       // 
                        indet.Parameters.AddWithValue("@m2obai2", mobai2);
                        indet.Parameters.AddWithValue("@t2aigv2", taigv2);
                        indet.Parameters.AddWithValue("@i2migv2", imigv2);       // importeIgv
                        indet.Parameters.AddWithValue("@i2mtoi2", imtoi2);
                        indet.Parameters.AddWithValue("@c2orae2", corae2);
                        if (tx_dat_mone.Text == MonDeft)
                        {
                            if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@u2biPtoO2", ubiPtoO2);
                                indet.Parameters.AddWithValue("@d2irPtoO2", dirPtoO2);
                                indet.Parameters.AddWithValue("@u2biPtoD2", ubiPtoD2);
                                indet.Parameters.AddWithValue("@d2irPtoD2", dirPtoD2);
                                indet.Parameters.AddWithValue("@d2etViaj2", detViaj2);
                                indet.Parameters.AddWithValue("@m2onRefS2", monRefS2);
                                indet.Parameters.AddWithValue("@m2onRefC2", monRefC2);
                                indet.Parameters.AddWithValue("@m2onRefU2", monRefU2);
                            }
                        }
                        else
                        {
                            if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@u2biPtoO2", ubiPtoO2);
                                indet.Parameters.AddWithValue("@d2irPtoO2", dirPtoO2);
                                indet.Parameters.AddWithValue("@u2biPtoD2", ubiPtoD2);
                                indet.Parameters.AddWithValue("@d2irPtoD2", dirPtoD2);
                                indet.Parameters.AddWithValue("@d2etViaj2", detViaj2);
                                indet.Parameters.AddWithValue("@m2onRefS2", monRefS2);
                                indet.Parameters.AddWithValue("@m2onRefC2", monRefC2);
                                indet.Parameters.AddWithValue("@m2onRefU2", monRefU2);
                            }
                        }
                    }
                    if (dataGridView1.Rows[2].Cells[0].Value.ToString() != "")                                  // tiene 3 lineas de detalle
                    {
                        indet.Parameters.AddWithValue("@t3idoem", tidoem);
                        indet.Parameters.AddWithValue("@n3udoem", nudoem);
                        indet.Parameters.AddWithValue("@t3ipdoc", tipdoc);
                        indet.Parameters.AddWithValue("@s3ernum", sernum);
                        indet.Parameters.AddWithValue("@n3uori3", nuori3);
                        indet.Parameters.AddWithValue("@c3odprd3", codprd3);     // codigo producto
                        indet.Parameters.AddWithValue("@c3oprsu3", coprsu3);     // codigo producto SUNAT
                        indet.Parameters.AddWithValue("@d3escr3", descr3);
                        indet.Parameters.AddWithValue("@c3anti3", canti3);
                        indet.Parameters.AddWithValue("@u3nime3", unime3);
                        indet.Parameters.AddWithValue("@i3ntos3", intos3);
                        indet.Parameters.AddWithValue("@i3nuns3", inuns3);
                        indet.Parameters.AddWithValue("@i3nunc3", inunc3);       // 
                        indet.Parameters.AddWithValue("@c3oimu3", coimu3);       // 
                        indet.Parameters.AddWithValue("@m3obai3", mobai3);
                        indet.Parameters.AddWithValue("@t3aigv3", taigv3);
                        indet.Parameters.AddWithValue("@i3migv3", imigv3);
                        indet.Parameters.AddWithValue("@i3mtoi3", imtoi3);
                        indet.Parameters.AddWithValue("@c3orae3", corae3);
                        if (tx_dat_mone.Text == MonDeft)
                        {
                            if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@u3biPtoO3", ubiPtoO3);
                                indet.Parameters.AddWithValue("@d3irPtoO3", dirPtoO3);
                                indet.Parameters.AddWithValue("@u3biPtoD3", ubiPtoD3);
                                indet.Parameters.AddWithValue("@d3irPtoD3", dirPtoD3);
                                indet.Parameters.AddWithValue("@d3etViaj3", detViaj3);
                                indet.Parameters.AddWithValue("@m3onRefS3", monRefS3);
                                indet.Parameters.AddWithValue("@m3onRefC3", monRefC3);
                                indet.Parameters.AddWithValue("@m3onRefU3", monRefU3);
                            }
                        }
                        else
                        {
                            if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@u3biPtoO3", ubiPtoO3);
                                indet.Parameters.AddWithValue("@d3irPtoO3", dirPtoO3);
                                indet.Parameters.AddWithValue("@u3biPtoD3", ubiPtoD3);
                                indet.Parameters.AddWithValue("@d3irPtoD3", dirPtoD3);
                                indet.Parameters.AddWithValue("@d3etViaj3", detViaj3);
                                indet.Parameters.AddWithValue("@m3onRefS3", monRefS3);
                                indet.Parameters.AddWithValue("@m3onRefC3", monRefC3);
                                indet.Parameters.AddWithValue("@m3onRefU3", monRefU3);
                            }
                        }
                    }
                    if (dataGridView1.Rows[3].Cells[0].Value.ToString() != "")                                  // tiene 4 lineas de detalle
                    {
                        indet.Parameters.AddWithValue("@t4idoem", tidoem);
                        indet.Parameters.AddWithValue("@n4udoem", nudoem);
                        indet.Parameters.AddWithValue("@t4ipdoc", tipdoc);
                        indet.Parameters.AddWithValue("@s4ernum", sernum);
                        indet.Parameters.AddWithValue("@n4uori4", nuori4);
                        indet.Parameters.AddWithValue("@c4odprd4", codprd4);     // codigo producto
                        indet.Parameters.AddWithValue("@c4oprsu4", coprsu4);     // codigo producto SUNAT
                        indet.Parameters.AddWithValue("@d4escr4", descr4);
                        indet.Parameters.AddWithValue("@c4anti4", canti4);
                        indet.Parameters.AddWithValue("@u4nime4", unime4);
                        indet.Parameters.AddWithValue("@i4ntos4", intos4);
                        indet.Parameters.AddWithValue("@i4nuns4", inuns4);
                        indet.Parameters.AddWithValue("@i4nunc4", inunc4);       // 
                        indet.Parameters.AddWithValue("@c4oimu4", coimu4);       // 
                        indet.Parameters.AddWithValue("@m4obai4", mobai4);
                        indet.Parameters.AddWithValue("@t4aigv4", taigv4);
                        indet.Parameters.AddWithValue("@i4migv4", imigv4);
                        indet.Parameters.AddWithValue("@i4mtoi4", imtoi4);
                        indet.Parameters.AddWithValue("@c4orae4", corae4);
                        if (tx_dat_mone.Text == MonDeft)
                        {
                            if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@u4biPtoO4", ubiPtoO4);
                                indet.Parameters.AddWithValue("@d4irPtoO4", dirPtoO4);
                                indet.Parameters.AddWithValue("@u4biPtoD4", ubiPtoD4);
                                indet.Parameters.AddWithValue("@d4irPtoD4", dirPtoD4);
                                indet.Parameters.AddWithValue("@d4etViaj4", detViaj4);
                                indet.Parameters.AddWithValue("@m4onRefS4", monRefS4);
                                indet.Parameters.AddWithValue("@m4onRefC4", monRefC4);
                                indet.Parameters.AddWithValue("@m4onRefU4", monRefU4);
                            }
                        }
                        else
                        {
                            if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                            {
                                indet.Parameters.AddWithValue("@u4biPtoO4", ubiPtoO4);
                                indet.Parameters.AddWithValue("@d4irPtoO4", dirPtoO4);
                                indet.Parameters.AddWithValue("@u4biPtoD4", ubiPtoD4);
                                indet.Parameters.AddWithValue("@d4irPtoD4", dirPtoD4);
                                indet.Parameters.AddWithValue("@d4etViaj4", detViaj4);
                                indet.Parameters.AddWithValue("@m4onRefS4", monRefS4);
                                indet.Parameters.AddWithValue("@m4onRefC4", monRefC4);
                                indet.Parameters.AddWithValue("@m4onRefU4", monRefU4);
                            }
                        }
                    }
                    indet.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en insertar detalle bizlinks");
                    Application.Exit();
                    return;
                }
                // cabecera ******************************************************************************************************************************************
                try
                {
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
                    inserta.Parameters.AddWithValue("@pasiemi", pasiemi);                                                     // v paisEmisor
                    inserta.Parameters.AddWithValue("@codaux40_1", codaux40_1);                                               // v codigoAuxiliar40_1,
                    inserta.Parameters.AddWithValue("@etiaux40_1", etiaux40_1);                                               // v textoAuxiliar40_1
                    inserta.Parameters.AddWithValue("@tidoad", tidoad);
                    inserta.Parameters.AddWithValue("@nudoad", nudoad);
                    inserta.Parameters.AddWithValue("@rasoad", rasoad);
                    inserta.Parameters.AddWithValue("@coradq", coradq);
                    inserta.Parameters.AddWithValue("@totimp", totimp);
                    inserta.Parameters.AddWithValue("@tovane", tovane);
                    //inserta.Parameters.AddWithValue("@tovano", tovano);
                    //inserta.Parameters.AddWithValue("@tovaex", tovaex);
                    //inserta.Parameters.AddWithValue("@tovagr", tovagr);
                    inserta.Parameters.AddWithValue("@totigv", totigv);
                    inserta.Parameters.AddWithValue("@totvta", totvta);
                    inserta.Parameters.AddWithValue("@tipope", tipope);
                    inserta.Parameters.AddWithValue("@coley1", coley1);
                    inserta.Parameters.AddWithValue("@teley1", teley1);
                    inserta.Parameters.AddWithValue("@estreg", estreg);
                    inserta.Parameters.AddWithValue("@tiref1", tiref1);
                    inserta.Parameters.AddWithValue("@nudor1", nudor1);
                    if (tiref2 != "")
                    {
                        inserta.Parameters.AddWithValue("@tiref2", tiref2);
                        inserta.Parameters.AddWithValue("@nudor2", nudor2);
                    }
                    if (tiref3 != "")
                    {
                        inserta.Parameters.AddWithValue("@tiref3", tiref3);
                        inserta.Parameters.AddWithValue("@nudor3", nudor3);
                    }
                    if (tiref4 != "")
                    {
                        inserta.Parameters.AddWithValue("@tiref4", tiref4);
                        inserta.Parameters.AddWithValue("@nudor4", nudor4);
                    }
                    if (tx_dat_mone.Text == MonDeft)
                    {
                        if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra) && tx_dat_tdv.Text == codfact)
                        {
                            inserta.Parameters.AddWithValue("@coddetra", Program.coddetra);
                            inserta.Parameters.AddWithValue("@totdet", totdet);
                            inserta.Parameters.AddWithValue("@pordetra", Program.pordetra);
                            inserta.Parameters.AddWithValue("@ctadetra", Program.ctadetra);
                            inserta.Parameters.AddWithValue("@codleyt", codleyt);
                            inserta.Parameters.AddWithValue("@leydet", leydet);
                        }
                    }
                    else
                    {
                        if (double.Parse(tx_flete.Text) > (double.Parse(Program.valdetra) / double.Parse(tx_tipcam.Text)) && tx_dat_tdv.Text == codfact)
                        {
                            inserta.Parameters.AddWithValue("@coddetra", Program.coddetra);
                            inserta.Parameters.AddWithValue("@totdet", totdet);
                            inserta.Parameters.AddWithValue("@pordetra", Program.pordetra);
                            inserta.Parameters.AddWithValue("@ctadetra", Program.ctadetra);
                            inserta.Parameters.AddWithValue("@codleyt", codleyt);
                            inserta.Parameters.AddWithValue("@leydet", leydet);
                        }
                    }
                    inserta.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en insertar cabecera bizlinks");
                    Application.Exit();
                    return;
                }
                // adicionales cabecera     SPE_EINVOICEHEADER_ADD
                string diradq = tx_dirRem.Text.Trim();
                string ubiadq = tx_ubigRtt.Text; // "150101";
                string urbadq = "";
                string proadq = tx_provRtt.Text.Trim();       // 
                string depadq = tx_dptoRtt.Text.Trim();
                string disadq = tx_distRtt.Text.Trim();
                string paiadq = "PE";                       // prueba
                string formpa = "009";                      // forma de pago efectivo
                string insadd = "insert into SPE_EINVOICEHEADER_ADD (" +
                    "clave,numeroDocumentoEmisor,serieNumero,tipoDocumento,tipoDocumentoEmisor,valor) values " +
                    "('direccionAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@diradq)," +
                    "('provinciaAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@proadq)," +
                    "('departamentoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@depadq)," +
                    "('distritoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@disadq)," +
                    "('paisAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@paiadq)," +
                    "('formaPago',@nudoem,@sernum,@tipdoc,@tidoem,@formpa)";
                //  "('ubigeoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@ubiadq)," +   // no va por el momento hasta tenerlo
                SqlCommand insertadd = new SqlCommand(insadd, conms);
                insertadd.Parameters.AddWithValue("@nudoem", nudoem);
                insertadd.Parameters.AddWithValue("@sernum", sernum);
                insertadd.Parameters.AddWithValue("@tipdoc", tipdoc);
                insertadd.Parameters.AddWithValue("@tidoem", tidoem);
                insertadd.Parameters.AddWithValue("@diradq", diradq);
                insertadd.Parameters.AddWithValue("@ubiadq", ubiadq);
                insertadd.Parameters.AddWithValue("@proadq", proadq);
                insertadd.Parameters.AddWithValue("@depadq", depadq);
                insertadd.Parameters.AddWithValue("@disadq", disadq);
                insertadd.Parameters.AddWithValue("@paiadq", paiadq);
                insertadd.Parameters.AddWithValue("@formpa", formpa);
                try
                {
                    insertadd.ExecuteNonQuery();
                }
                catch (MySqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en insertar adicional cabecera");
                    Application.Exit();
                    return;
                }
            }
            else
            {
                MessageBox.Show("No se puede conectar al servidor" + Environment.NewLine +
                "de Facturación Electrónica", "Error de conexión para escritura", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.Exit();
                return;
            }
        }
        #endregion

        #region autocompletados
        private void autodepa()                 // se jala en el load
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
                case "partida":
                    if (tx_dp_dep.Text.Trim() != "")
                    {
                        DataRow[] provi = dataUbig.Select("depart='" + tx_dat_upo.Text.Substring(0, 2) + "' and provin<>'00' and distri='00'");
                        provincias.Clear();
                        foreach (DataRow row in provi)
                        {
                            provincias.Add(row["nombre"].ToString());
                        }
                    }
                    break;
                case "llegada":
                    if (tx_dd_dep.Text.Trim() != "")
                    {
                        DataRow[] provi = dataUbig.Select("depart='" + tx_dat_upd.Text.Substring(0, 2) + "' and provin<>'00' and distri='00'");
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
                    if (tx_dat_upo.Text.Trim() != "" && tx_dp_pro.Text.Trim() != "")
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
                    if (tx_dat_upd.Text.Trim() != "" && tx_dd_pro.Text.Trim() != "")
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
        }
        private void escribe()
        {
            lp.escribe(this);
            tx_nomRem.ReadOnly = true;
            //tx_dirRem.ReadOnly = true;
            //tx_dptoRtt.ReadOnly = true;
            //tx_provRtt.ReadOnly = true;
            //tx_distRtt.ReadOnly = true;
        }
        private void limpiar()
        {
            lp.limpiar(this);
        }
        private void limpia_chk()    
        {
            lp.limpia_chk(this);
        }
        private void limpia_otros()
        {
            //
        }
        private void limpia_combos()
        {
            lp.limpia_cmb(this);
            cmb_plazoc.SelectedIndex = -1;
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
                // validamos que la GR: 1.exista, 2.No este facturada, 3.No este anulada
                if (validGR(tx_serGR.Text, tx_numGR.Text) == false)
                {
                    MessageBox.Show("La GR no existe, esta anulada o ya esta facturada", "Error en Guía", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    tx_numGR.Text = "";
                    tx_numGR.Focus();
                    return;
                }
                else
                {
                    rb_desGR.PerformClick();
                }
                //dataGridView1.Rows.Clear(); nooooo, se puede hacer una fact de varias guias, n guias
                dataGridView1.Rows.Add(datguias[0], datguias[1], datguias[2], datguias[3], datguias[4], datguias[5], datguias[6], datguias[9], datguias[10], datguias[7], datguias[15],datguias[16]);     // insertamos en la grilla los datos de la GR
                int totfil = 0;
                int totcant = 0;
                decimal totflet = 0;    // acumulador en moneda de la GR 
                decimal totflMN = 0;
                tx_dat_mone.Text = datguias[7].ToString();
                cmb_mon.SelectedValue = datguias[7].ToString();
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
                        }
                        else
                        {
                            totflet = totflet + decimal.Parse(dataGridView1.Rows[i].Cells[4].Value.ToString()); // VALOR DE LA GR EN SU MONEDA
                            totflMN = totflMN + decimal.Parse(dataGridView1.Rows[i].Cells[5].Value.ToString()); // VALOR DE LA GR EN MONEDA LOCAL
                        }
                    }
                }
                tx_tfmn.Text = totflMN.ToString("#0.00");
                tx_totcant.Text = totcant.ToString();
                tx_tfil.Text = totfil.ToString();
                tx_flete.Text = totflet.ToString("#0.00");
                tx_fletMN.Text = totflMN.ToString("#0.00"); // Math.Round(decimal.Parse(tx_flete.Text) * decimal.Parse(tx_tipcam.Text), 2).ToString();
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
                if (dataGridView1.Rows.Count >= 3 && decimal.Parse(tx_dat_saldoGR.Text) <= 0)
                {
                    MessageBox.Show("El presente comprobante no se " + Environment.NewLine +
                         "puede cobrar en automático", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    rb_si.Checked = false;
                    rb_si.Enabled = false;
                    tx_salxcob.Text = tx_flete.Text;
                    tx_pagado.Text = "0";
                }
                if (dataGridView1.Rows.Count <= 2 && decimal.Parse(tx_dat_saldoGR.Text) <= 0)
                {
                    MessageBox.Show("La GR esta cancelada, el documento de venta"+ Environment.NewLine +
                         "se creará con el estado cancelado","Atención verifique",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    //rb_si.PerformClick();
                    rb_no.Enabled = false;
                    rb_si.Enabled = false;
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
                rb_si.Checked = false;
                rb_no.Checked = false;   // true
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
                if (rb_si.Checked == false && rb_no.Checked == false && rb_no.Enabled == true)
                {
                    MessageBox.Show("Seleccione si se cancela la factura o no","Atención - Confirme",MessageBoxButtons.OK,MessageBoxIcon.Warning);
                    rb_si.Focus();
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
                if (tx_idr.Text.Trim() == "")
                {
                    var aa = MessageBox.Show("Confirma que desea crear el documento?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        if (true)  // lib.DirectoryVisible(rutatxt) == true
                        {
                            if (graba() == true)
                            {
                                OSE_Fact_Elect();
                                grabfactelec();
                                if (true)       // factElec(nipfe, "txt", "alta", 0) == true
                                {
                                    /* actualizamos la tabla seguimiento de usuarios
                                    string resulta = lib.ult_mov(nomform, nomtab, asd);
                                    if (resulta != "OK")
                                    {
                                        MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    */
                                    var bb = MessageBox.Show("Desea imprimir el documento?" + Environment.NewLine +
                                        "El formato actual es " + vi_formato, "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                    if (bb == DialogResult.Yes)
                                    {
                                        Bt_print.PerformClick();
                                    }
                                }
                                else
                                {
                                    MessageBox.Show("No se puede generar el documento de venta electrónico" + Environment.NewLine +
                                        "Se generó una anulación interna para el presente documento", "Error en proveedor de Fact.Electrónica");
                                    iserror = "si";
                                    anula("INT");
                                }
                            }
                            else
                            {
                                MessageBox.Show("No se puede grabar el documento de venta electrónico", "Error en conexión");
                                iserror = "si";
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
                            edita();    // modificacion total
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
                    //tx_numero.Focus();
                    //return;
                }
                // validaciones de fecha para poder anular
                DateTime fedv = DateTime.Parse(tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                TimeSpan span = DateTime.Parse(lib.fechaServ("ansi")) - fedv;
                if (span.Days > v_cdpa)
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
                                if (lib.DirectoryVisible(rutatxt) == true)
                                {
                                    /*
                                    int cta = anula("FIS");      // cantidad de doc.vtas anuladas en la fecha
                                    if (factElec(nipfe, "txt", "baja", cta) == true)
                                    {
                                        string resulta = lib.ult_mov(nomform, nomtab, asd);
                                        if (resulta != "OK")
                                        {
                                            MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                        }
                                    }
                                    */
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
                        MessageBox.Show("No se puede anular por estar fuera de plazo","Usuario no permito",MessageBoxButtons.OK,MessageBoxIcon.Asterisk);
                    }
                }
                else
                {
                    if (tx_idr.Text.Trim() != "")
                    {
                        var aa = MessageBox.Show("Confirma que desea ANULAR el documento?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (aa == DialogResult.Yes)
                        {
                            if (lib.DirectoryVisible(rutatxt) == true)
                            {
                                /*
                                int cta = anula("FIS");      // cantidad de doc.vtas anuladas en la fecha
                                if (factElec(nipfe, "txt", "baja", cta) == true)
                                {
                                    string resulta = lib.ult_mov(nomform, nomtab, asd);
                                    if (resulta != "OK")
                                    {
                                        MessageBox.Show(resulta, "Error en actualización de seguimiento", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                }
                                */
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
            MySqlConnection conn = new MySqlConnection(db_conn_grael);   // DB_CONN_STR
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
                /*
                string inserta = "insert into cabfactu (" +
                    "fechope,martdve,tipdvta,serdvta,numdvta,ticltgr,tidoclt,nudoclt,nombclt,direclt,dptoclt,provclt,distclt,ubigclt,corrclt,teleclt," +
                    "locorig,dirorig,ubiorig,obsdvta,canfidt,canbudt,mondvta,tcadvta,subtota,igvtota,porcigv,totdvta,totpags,saldvta,estdvta,frase01," +
                    "tipoclt,m1clien,tippago,ferecep,impreso,codMN,subtMN,igvtMN,totdvMN,pagauto,tipdcob,idcaja,plazocred,porcendscto,valordscto," +
                    "cargaunica,placa,confveh,autoriz,detPeso,detputil,detMon1,detMon2,detMon3,dirporig,ubiporig,dirpdest,ubipdest," +
                    "verApp,userc,fechc,diriplan4,diripwan4,netbname) values (" +
                    "@fechop,@mtdvta,@ctdvta,@serdv,@numdv,@tcdvta,@tdcrem,@ndcrem,@nomrem,@dircre,@dptocl,@provcl,@distcl,@ubicre,@mailcl,@telecl," +
                    "@ldcpgr,@didegr,@ubdegr,@obsprg,@canfil,@totcpr,@monppr,@tcoper,@subpgr,@igvpgr,@porcigv,@totpgr,@pagpgr,@salxpa,@estpgr,@frase1," +
                    "@ticlre,@m1clte,@tipacc,@feredv,@impSN,@codMN,@subMN,@igvMN,@totMN,@pagaut,@tipdco,@idcaj,@plazc,@pordesc,@valdesc," +
                    "@caruni,@placa,@confv,@autor,@dPeso,@dputil,@dMon1,@dMon2,@dMon3,@dporig,@uporig,@dpdest,@updest," +
                    "@verApp,@asd,now(),@iplan,@ipwan,@nbnam)";
                */
                string inserta = "insert into madocvtas (fechope,tipcam,docvta,servta,corvta,doccli,numdcli,direc,nomclie," +
                    "observ,moneda,aumigv,subtot,igv,doctot,status,pigv,userc,fechc," +
                    "local,rd3,dist,prov,dpto,saldo,cdr,mfe,email,ubiclte,canfild," +
                    "prop01,prop02,prop03,prop04,prop05,prop06,prop07,prop08,prop09,prop10) " +
                    "values (" +
                    "@fechop,@tcoper,@ctdvta,@serdv,@numdv,@tdcrem,@ndcrem,@dircre,@nomrem," +
                    "@obsprg,@monppr,@aig,@subpgr,@igvpgr,@totpgr,@estpgr,@porcigv,@asd,now()," +
                    "@ldcpgr,@tcdvta,@distcl,@provcl,@dptocl,@salxpa,@cdr,@mtdvta,@mailcl,@ubicre,@canfil," +
                    "@pr01,@pr02,@pr03,@pr04,@pr05,@pr06,@pr07,@pr08,@pr09,@pr10)";

                //                    "prop01,prop02,prop03,prop04,prop05,prop06,prop07,prop08,prop09,prop10) " +
                //                     "@pr01,@pr02,@pr03,@pr04,@pr05,@pr06,@pr07,@pr08,@pr09,@pr10)";
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
                    micon.Parameters.AddWithValue("@estpgr", (tx_pagado.Text == "" || tx_pagado.Text == "0.00") ? tx_dat_estad.Text : codCanc); // estado
                    micon.Parameters.AddWithValue("@porcigv", v_igv);                           // porcentaje en numeros de IGV
                    micon.Parameters.AddWithValue("@asd", asd);
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
                    /*
                    micon.Parameters.AddWithValue("@telecl", tx_telc1.Text);
                    micon.Parameters.AddWithValue("@didegr", dirloc);                             // direccion origen
                    micon.Parameters.AddWithValue("@ubdegr", ubiloc);                             // ubigeo origen
                    micon.Parameters.AddWithValue("@totcpr", tx_totcant.Text);  // total bultos
                    micon.Parameters.AddWithValue("@pagpgr", (rb_si.Checked == true) ? tx_fletMN.Text : "0");  // (tx_pagado.Text == "") ? "0" : tx_pagado.Text);
                    micon.Parameters.AddWithValue("@frase1", "");                   // no hay nada que poner 19/11/2020
                    micon.Parameters.AddWithValue("@ticlre", tx_dat_tcr.Text);      // tipo de cliente credito o contado
                    micon.Parameters.AddWithValue("@m1clte", tx_dat_m1clte.Text);
                    micon.Parameters.AddWithValue("@tipacc", v_mpag);                   // pago del documento x defecto si nace la fact pagada
                    micon.Parameters.AddWithValue("@feredv", DBNull.Value);         // si es pago contado la fecha de recep del doc. es la misma fecha
                    micon.Parameters.AddWithValue("@impSN", "N");
                    micon.Parameters.AddWithValue("@codMN", MonDeft);               // codigo moneda local
                    micon.Parameters.AddWithValue("@subMN", subtMN);
                    micon.Parameters.AddWithValue("@igvMN", igvtMN);
                    micon.Parameters.AddWithValue("@totMN", fletMN);
                    micon.Parameters.AddWithValue("@pagaut", (rb_si.Checked == true)? "S" : "N");
                    micon.Parameters.AddWithValue("@tipdco", (rb_si.Checked == true)? v_codcob : "");
                    micon.Parameters.AddWithValue("@idcaj", (rb_si.Checked == true)? tx_idcaja.Text : "0");
                    micon.Parameters.AddWithValue("@plazc", (rb_no.Checked == true)? codppc : "");
                    micon.Parameters.AddWithValue("@pordesc", (tx_dat_porcDscto.Text.Trim() == "") ? "0" : tx_dat_porcDscto.Text);
                    micon.Parameters.AddWithValue("@valdesc", (tx_valdscto.Text.Trim() == "") ? "0" : tx_valdscto.Text);
                    micon.Parameters.AddWithValue("@caruni", (chk_cunica.Checked == true)? 1 : 0);
                    micon.Parameters.AddWithValue("@placa", tx_pla_placa.Text);
                    micon.Parameters.AddWithValue("@confv", tx_pla_confv.Text);
                    micon.Parameters.AddWithValue("@autor", tx_pla_autor.Text);
                    micon.Parameters.AddWithValue("@dPeso", (tx_cetm.Text.Trim() == "")? "0" : tx_cetm.Text);
                    micon.Parameters.AddWithValue("@dputil", (tx_cutm.Text.Trim() == "")? "0" : tx_cutm.Text);
                    micon.Parameters.AddWithValue("@dMon1", (tx_valref1.Text.Trim() == "")? "0" : tx_valref1.Text);
                    micon.Parameters.AddWithValue("@dMon2", (tx_valref2.Text.Trim() == "")? "0" : tx_valref2.Text);
                    micon.Parameters.AddWithValue("@dMon3", (tx_valref3.Text.Trim() == "")? "0" : tx_valref3.Text);
                    micon.Parameters.AddWithValue("@dporig", tx_dat_dpo.Text);
                    micon.Parameters.AddWithValue("@uporig", tx_dat_upo.Text);
                    micon.Parameters.AddWithValue("@dpdest", tx_dat_dpd.Text);
                    micon.Parameters.AddWithValue("@updest", tx_dat_upd.Text);
                    micon.Parameters.AddWithValue("@verApp", verapp);
                    micon.Parameters.AddWithValue("@iplan", lib.iplan());
                    micon.Parameters.AddWithValue("@ipwan", Grael2.Program.vg_ipwan);
                    micon.Parameters.AddWithValue("@nbnam", Environment.MachineName);
                    */
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
                // detalle
                if (dataGridView1.Rows.Count > 0)
                {
                    int fila = 1;
                    for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
                    {
                        if (dataGridView1.Rows[i].Cells[0].Value.ToString().Trim() != "")
                        {
                            /*
                             * string inserd2 = "update detfactu set " +
                                "codgror=@guia,cantbul=@bult,unimedp=@unim,descpro=@desc,pesogro=@peso,codmogr=@codm,totalgr=@pret,codMN=@cmnn," +
                                "totalgrMN=@tgrmn,pagauto=@pagaut " +
                                "where idc=@idr and filadet=@fila";
                            
                            string inserd2 = "update detavtas set " +
                                "sergr=@sgui,corgr=@cgui,moneda=@codm,valor=@pret,ruta=@ruta,glosa=@desc,docremi=@dsus,bultos=@bult,valorel=@pret " +
                                "where idc=@idr and filadet=@fila";
                            */
                            string inserd2 = "insert into detavtas (idc,docvta,servta,corvta,sergr,corgr,moneda," +
                        "valor,ruta,glosa,status,userc,fechc,docremi,bultos,monrefd1,monrefd2,monrefd3) " +
                        "values (@idr,@doc,@svt,@cvt,@sgui,@cgui,@codm," +
                        "@pret,@ruta,@desc,@sta,@asd,now(),@dre,@bult,@mrd1,@mrd2,@mrd3)";
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
                                micon.Parameters.AddWithValue("@desc", dataGridView1.Rows[i].Cells[1].Value.ToString());
                                micon.Parameters.AddWithValue("@sta", (tx_pagado.Text == "" || tx_pagado.Text == "0.00") ? tx_dat_estad.Text : codCanc); // estado;
                                micon.Parameters.AddWithValue("@asd", asd);
                                micon.Parameters.AddWithValue("@dre", dataGridView1.Rows[i].Cells[8].Value.ToString());
                                micon.Parameters.AddWithValue("@bult", dataGridView1.Rows[i].Cells[2].Value.ToString());
                                micon.Parameters.AddWithValue("@mrd1", "0.00"); // (tx_dref1.Text == "") ? "0.00" : tx_dref1.Text
                                micon.Parameters.AddWithValue("@mrd2", "0.00"); // (tx_dcar1.Text == "") ? "0.00" : tx_dcar1.Text
                                micon.Parameters.AddWithValue("@mrd3", "0.00");   // (tx_dnom1.Text == "") ? "0.00" : tx_dnom1.Text
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
            MySqlConnection conn = new MySqlConnection(DB_CONN_STR);
            conn.Open();
            if (conn.State == ConnectionState.Open)
            {
                try
                {
                    if (true)     // EDICION DE CABECERA
                    {
                        string actua = "update cabfactu a set obsdvta=@obsprg," +
                            "a.verApp=@verApp,a.userm=@asd,a.fechm=now(),a.diriplan4=@iplan,a.diripwan4=@ipwan,a.netbname=@nbnam " +
                            "where a.id=@idr";
                        MySqlCommand micon = new MySqlCommand(actua, conn);
                        micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                        micon.Parameters.AddWithValue("@obsprg", tx_obser1.Text);
                        micon.Parameters.AddWithValue("@verApp", verapp);
                        micon.Parameters.AddWithValue("@asd", asd);
                        micon.Parameters.AddWithValue("@iplan", lib.iplan());
                        micon.Parameters.AddWithValue("@ipwan", Grael2.Program.vg_ipwan);
                        micon.Parameters.AddWithValue("@nbnam", Environment.MachineName);
                        micon.ExecuteNonQuery();
                        //
                        // EDICION DEL DETALLE .... no hay 28/10/2020
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
        private int anula(string tipo)
        {
            int ctanul = 0;
            // en el caso de documentos de venta HAY 1: ANULACION FISICA ... 28/10/2020
            // tambien podría haber ANULACION interna con la serie ANU1 .... 19/11/2020
            // Anulacion fisica se "anula" el numero del documento en sistema y en fisico se tacha y en prov. fact.electronica se da baja de numeracion
            // se borran todos los enlaces mediante triggers en la B.D.
            if (tipo == "FIS")
            {
                using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        string canul = "update cabfactu set estdvta=@estser,obsdvta=@obse,usera=@asd,fecha=now()," +
                            "verApp=@veap,diriplan4=@dil4,diripwan4=@diw4,netbname=@nbnp,estintreg=@eiar " +
                            "where id=@idr";
                        using (MySqlCommand micon = new MySqlCommand(canul, conn))
                        {
                            micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                            micon.Parameters.AddWithValue("@estser", codAnul);
                            micon.Parameters.AddWithValue("@obse", tx_obser1.Text);
                            micon.Parameters.AddWithValue("@asd", asd);
                            micon.Parameters.AddWithValue("@dil4", lib.iplan());
                            micon.Parameters.AddWithValue("@diw4", Grael2.Program.vg_ipwan);
                            micon.Parameters.AddWithValue("@nbnp", Environment.MachineName);
                            micon.Parameters.AddWithValue("@veap", verapp);
                            micon.Parameters.AddWithValue("@eiar", (vint_A0 == codAnul) ? "A0" : "");  // codigo anulacion interna en DB A0
                            micon.ExecuteNonQuery();
                        }
                        string consul = "select count(id) from cabfactu where date(fecha)=@fech and estdvta=@estser";
                        using (MySqlCommand micon = new MySqlCommand(consul, conn))
                        {
                            micon.Parameters.AddWithValue("@fech", tx_fechact.Text.Substring(6, 4) + "-" + tx_fechact.Text.Substring(3, 2) + "-" + tx_fechact.Text.Substring(0, 2));
                            micon.Parameters.AddWithValue("@estser", codAnul);
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                if (dr.Read())
                                {
                                    ctanul = dr.GetInt32(0);
                                }
                            }
                        }
                    }
                }
            }
            if (tipo == "INT")
            {
                using (MySqlConnection conn = new MySqlConnection(DB_CONN_STR))
                {
                    conn.Open();
                    if (conn.State == ConnectionState.Open)
                    {
                        string canul = "update cabfactu set serdvta=@sain,estdvta=@estser,obsdvta=@obse,usera=@asd,fecha=now()," +
                            "verApp=@veap,diriplan4=@dil4,diripwan4=@diw4,netbname=@nbnp,estintreg=@eiar " +
                            "where id=@idr";
                        using (MySqlCommand micon = new MySqlCommand(canul, conn))
                        {
                            micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                            micon.Parameters.AddWithValue("@sain", v_sanu);
                            micon.Parameters.AddWithValue("@estser", codAnul);
                            micon.Parameters.AddWithValue("@obse", tx_obser1.Text);
                            micon.Parameters.AddWithValue("@asd", asd);
                            micon.Parameters.AddWithValue("@dil4", lib.iplan());
                            micon.Parameters.AddWithValue("@diw4", Grael2.Program.vg_ipwan);
                            micon.Parameters.AddWithValue("@nbnp", Environment.MachineName);
                            micon.Parameters.AddWithValue("@veap", verapp);
                            micon.Parameters.AddWithValue("@eiar", (vint_A0 == codAnul) ? "A0" : "");  // codigo anulacion interna en DB A0
                            micon.ExecuteNonQuery();
                        }
                        string updser = "update series set actual=actual-1 where tipdoc=@tipd AND serie=@serd";
                        using (MySqlCommand micon = new MySqlCommand(updser, conn))
                        {
                            micon.Parameters.AddWithValue("@tipd", tx_dat_tdv.Text);
                            micon.Parameters.AddWithValue("@serd", tx_serie.Text);
                            micon.ExecuteNonQuery();
                        }
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
        private void tx_dp_dep_Leave(object sender, EventArgs e)        // departamento del punto de partida
        {
            if (tx_dp_dep.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("nombre='" + tx_dp_dep.Text.Trim() + "' and provin='00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_dat_upo.Text = row[0].ItemArray[1].ToString();
                    autoprov("partida");
                }
                else tx_dp_dep.Text = "";
            }
        }
        private void tx_dp_pro_Leave(object sender, EventArgs e)        // provincia del punto de partida
        {
            if (tx_dp_pro.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_dat_upo.Text.Substring(0, 2) + "' and nombre='" + tx_dp_pro.Text.Trim() + "' and provin<>'00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_dat_upo.Text = tx_dat_upo.Text.Trim().Substring(0, 2) + row[0].ItemArray[2].ToString();
                    autodist("partida");
                }
                else tx_dp_pro.Text = "";
            }
        }
        private void tx_dp_dis_Leave(object sender, EventArgs e)        // distrito del punto de partida
        {
            if (tx_dp_dis.Text.Trim() != "" && tx_dp_pro.Text.Trim() != "" && tx_dp_dep.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_dat_upo.Text.Substring(0, 2) + "' and provin='" + tx_dat_upo.Text.Substring(2, 2) + "' and nombre='" + tx_dp_dis.Text.Trim() + "'");
                if (row.Length > 0)
                {
                    tx_dat_upo.Text = tx_dat_upo.Text.Trim().Substring(0, 4) + row[0].ItemArray[3].ToString();
                }
                else tx_dp_dis.Text = "";
            }
        }
        private void tx_dd_dep_Leave(object sender, EventArgs e)        // departamento del punto de llegada
        {
            if (tx_dd_dep.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("nombre='" + tx_dd_dep.Text.Trim() + "' and provin='00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_dat_upd.Text = row[0].ItemArray[1].ToString();
                    autoprov("llegada");
                }
                else tx_dd_dep.Text = "";
            }
        }
        private void tx_dd_pro_Leave(object sender, EventArgs e)        // provincia del punto de llegada
        {
            if (tx_dd_pro.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_dat_upd.Text.Substring(0, 2) + "' and nombre='" + tx_dd_pro.Text.Trim() + "' and provin<>'00' and distri='00'");
                if (row.Length > 0)
                {
                    tx_dat_upd.Text = tx_dat_upd.Text.Trim().Substring(0, 2) + row[0].ItemArray[2].ToString();
                    autodist("llegada");
                }
                else tx_dd_pro.Text = "";
            }
        }
        private void tx_dd_dis_Leave(object sender, EventArgs e)        // distrito del punto de llegada
        {
            if (tx_dd_dis.Text.Trim() != "" && tx_dd_pro.Text.Trim() != "" && tx_dd_dep.Text.Trim() != "")
            {
                DataRow[] row = dataUbig.Select("depart='" + tx_dat_upd.Text.Substring(0, 2) + "' and provin='" + tx_dat_upd.Text.Substring(2, 2) + "' and nombre='" + tx_dd_dis.Text.Trim() + "'");
                if (row.Length > 0)
                {
                    tx_dat_upd.Text = tx_dat_upd.Text.Trim().Substring(0, 4) + row[0].ItemArray[3].ToString();
                }
                else tx_dd_dis.Text = "";
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
                tx_numero.Text = lib.Right("0000000" + tx_numero.Text, 7);
                //tx_idr.Text = tx_numero.Text;
                jalaoc("sernum");
                dataGridView1.Rows.Clear();
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
        private void tx_pla_placa_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_pla_placa);
        }
        private void tx_pla_confv_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_pla_confv);
        }
        private void tx_pla_autor_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_pla_autor);
        }
        private void tx_obser1_Leave(object sender, EventArgs e)
        {
            val_NoCaracteres(tx_obser1);
        }
        private void rb_si_Click(object sender, EventArgs e)
        {
            if (tx_idcaja.Text != "")
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
        private void rb_no_Click(object sender, EventArgs e)
        {
            double once = 0;
            for (int i = 0; i<dataGridView1.Rows.Count - 1; i++)
            {
                if (string.IsNullOrEmpty(dataGridView1.Rows[i].Cells[11].Value.ToString())) dataGridView1.Rows[i].Cells[11].Value = "0";
                once = once + double.Parse(dataGridView1.Rows[i].Cells[11].Value.ToString());
            }
            tx_pagado.Text = "0.00";
            tx_salxcob.Text = once.ToString("#0.00"); // tx_flete.Text;
            tx_salxcob.BackColor = Color.Red;
            cmb_plazoc.Enabled = true;
            cmb_plazoc.SelectedValue = codppc;
            tx_dat_plazo.Text = codppc;
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
        private void chk_cunica_CheckedChanged(object sender, EventArgs e)
        {
            cargaunica();
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
            tx_serGR.Focus();
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            sololee();          
            Tx_modo.Text = "EDITAR";                    // solo puede editarse la observacion 28/10/2020
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
            // Impresion ó Re-impresion ??
            if (tx_impreso.Text == "S")
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
                        if (imprimeTK() == true) updateprint("S");
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
            // Cantidad de copias
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
            tx_serie.Focus();
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
        private void cmb_mon_SelectedIndexChanged(object sender, EventArgs e)
        {
            /*if (Tx_modo.Text == "NUEVO" && tx_totcant.Text != "")    //  || Tx_modo.Text == "EDITAR"
            {   // lo de totcant es para accionar solo cuando el detalle de la GR se haya cargado
                if (cmb_mon.SelectedIndex > -1)
                {
                    tx_dat_mone.Text = cmb_mon.SelectedValue.ToString();
                    DataRow[] row = dtm.Select("idcodice='"+ tx_dat_mone.Text+"'");
                    tx_dat_monsunat.Text = row[0][2].ToString();
                    tipcambio(tx_dat_mone.Text);
                    if (tx_flete.Text != "" && tx_flete.Text != "0.00") calculos(decimal.Parse(tx_flete.Text));
                    if (rb_no.Checked == true) rb_no_Click(null,null);
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
            } */
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
                    if (Tx_modo.Text == "NUEVO") tx_serie.Text = row[0].ItemArray[5].ToString();
                    tx_numero.Text = "";
                }
            }
        }
        private void cmb_mon_SelectionChangeCommitted(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO" && tx_totcant.Text != "")    //  || Tx_modo.Text == "EDITAR"
            {   // lo de totcant es para accionar solo cuando el detalle de la GR se haya cargado
                if (cmb_mon.SelectedValue.ToString() != tx_dat_mone.Text) // cmb_mon.SelectedIndex > -1
                {
                    tx_dat_mone.Text = cmb_mon.SelectedValue.ToString();
                    DataRow[] row = dtm.Select("idcodice='" + tx_dat_mone.Text + "'");
                    tx_dat_monsunat.Text = row[0][2].ToString();
                    tipcambio(tx_dat_mone.Text);
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
                tx_dat_plazo.Text = cmb_plazoc.SelectedValue.ToString();
                DataRow[] dias = dtp.Select("idcodice='" + tx_dat_plazo.Text + "'");
                foreach (DataRow row in dias)
                {
                    tx_dat_dpla.Text = row[3].ToString();
                }
            }
            else
            {
                tx_dat_plazo.Text = "";
                tx_dat_dpla.Text = "";
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
                string nomclie = Program.cliente;
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
                    e.Graphics.DrawString("RUC", lt_med, Brushes.Black, puntoF, StringFormat.GenericTypographic);    // nombre del cliente
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
                    StringFormat alder = new StringFormat(StringFormatFlags.DirectionRightToLeft);
                    SizeF siz = new SizeF(70, 15);
                    RectangleF recto = new RectangleF(puntoF, siz);
                    for (int l = 0; l < dataGridView1.Rows.Count - 1; l++)
                    {
                        if (!string.IsNullOrEmpty(dataGridView1.Rows[l].Cells[0].Value.ToString()))
                        {
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString(glosser, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString(glosser2, lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            //recto = new RectangleF(puntoF, siz);
                            e.Graphics.DrawString(dataGridView1.Rows[l].Cells[0].Value.ToString() + " " + dataGridView1.Rows[l].Cells[1].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                            puntoF = new PointF(coli, posi);
                            e.Graphics.DrawString("Según doc.cliente: " + dataGridView1.Rows[l].Cells[8].Value.ToString(), lt_peq, Brushes.Black, puntoF, StringFormat.GenericTypographic);
                            posi = posi + alfi;
                        }
                    }
                    // pie del documento ;
                    if (tx_dat_tdv.Text != codfact)
                    {
                        //SizeF siz = new SizeF(70, 15);
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
                        recto = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(tx_flete.Text, lt_peq, Brushes.Black, recto, alder);
                    }
                    if (tx_dat_tdv.Text == codfact)
                    {
                        //SizeF siz = new SizeF(70, 15);
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
                    if (double.Parse(tx_flete.Text) > double.Parse(Program.valdetra))                // leyenda de detracción
                    {
                        siz = new SizeF(CentimeterToPixel(anchTik), 15 * 3);
                        puntoF = new PointF(coli, posi);
                        recto = new RectangleF(puntoF, siz);
                        e.Graphics.DrawString(leydet1.Trim() + " " + Program.ctadetra.Trim(), lt_peq, Brushes.Black, recto, StringFormat.GenericTypographic);
                        posi = posi + alfi * 3;
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
        }
        #endregion

    }
}
