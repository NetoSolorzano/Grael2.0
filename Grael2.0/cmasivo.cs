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
    public partial class cmasivo : Form
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
        static string nomtab = "";                  // maestra de comprobantes del ERP_Grael 1.0 

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
        string codAnu1 = "";            // codigos de anulacion en el erp_grael 1.0
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
        string v_mpag = "";             // medio de pago automatico x defecto CONTADO
        string v_codcob = "";           // codigo del documento cobranza
        string v_CR_gr_ind = "";        // nombre del formato FT/BV en CR
        string v_mfildet = "";          // maximo numero de filas en el detalle, coord. con el formato ..... 4 filas detalle máximo
        string vint_A0 = "";            // variable codigo anulacion interna por BD
        string v_codidv = "";           // variable codifo interno de documento de venta en vista TDV
        string codfact = "";            // idcodice de factura
        string codbole = "";            // codigo de boleta electronica
        string v_igv = "";              // valor igv %
        string v_estcaj = "";           // estado de la caja
        string v_idcaj = "";            // id de la caja actual
        string logoclt = "";            // ruta y nombre archivo logo
        string fshoy = "";              // fecha hoy del servidor en formato ansi
        string codppc = "";             // codigo del plazo de pago por defecto para fact a crédito
        string codsuser_cu = "";        // usuarios autorizados a crear Ft de cargas unicas
        int v_cdpa = 0;                 // cantidad de días despues de emitida la fact. en que un usuario normal puede anular
        string vint_gg = "";            // glosa del detalle inicial de la guía "sin verificar contenido"
        decimal limbolsd = 0;           // limite en soles para boletas sin direccion
        int valini = 0;                 // valor inicial del rango nuevo valor
        int valfin = 0;                 // valor final del rango nuevo valor
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
        
        DataTable dataUbig = (DataTable)CacheManager.GetItem("ubigeos");

        // string de conexion
        string DB_CONN_STR = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.data + ";";
        string db_conn_grael = "server=" + login.serv + ";uid=" + login.usua + ";pwd=" + login.cont + ";database=" + login.dataG + ";";

        DataTable dttd0 = new DataTable();        // tipo doc remitente
        DataTable dttd1 = new DataTable();        // locales
        DataTable dtmon = new DataTable();        // moneda
        DataTable dtdvt = new DataTable();        // tipo doc venta
        //DataTable dtp = new DataTable();        // plazo de credito 
        //DataTable tcfe = new DataTable();       // facturacion electronica - cabecera
        //DataTable tdfe = new DataTable();       // facturacion electronica -detalle
        public string script = "";              // script de conexion a Bizlinks
        string tcbul = "0";                     // total bultos             -|
        string tdesc = "-";                     // descripcion de la carga   |-> datos de linea de detalle de cada fila
        string tunid = "-";                     // unidad de medida         -|
        NumLetra nl = new NumLetra();
        public cmasivo()
        {
            InitializeComponent();
        }
        private void cmasivo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter) SendKeys.Send("{TAB}");
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.N) Bt_add.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.E) Bt_edit.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.A) Bt_anul.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.O) Bt_ver.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.P) Bt_print.PerformClick();
            if (Control.ModifierKeys == Keys.Control && e.KeyCode == Keys.S) Bt_close.PerformClick();
        }
        private void cmasivo_Load(object sender, EventArgs e)
        {
            this.Focus();
            jalainfo();
            init();
            dataload();
            toolboton();
            this.KeyPreview = true;
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
            // todo desabilidado
            sololee();
        }
        private void grilla()
        {
            for (int i=1; i<dataGridView1.Columns.Count; i++)
            {
                dataGridView1.Columns[i].ReadOnly = true;
            }
        }
        private void initIngreso()
        {
            limpiar();
            limpia_chk();
            limpia_otros();
            limpia_combos();
            dataGridView1.Rows.Clear();
            //dataGridView1.ReadOnly = true;
            if (Tx_modo.Text == "NUEVO")
            {
                tx_fechope.Text = tx_fechact.Text.Substring(0, 10);
            }
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
                micon.Parameters.AddWithValue("@noca", "cmasivo");
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
                            if (row["param"].ToString() == "mpacon") v_mpag = row["valor"].ToString().Trim();               // medio de pago x defecto CONTADO
                            if (row["param"].ToString() == "factura") codfact = row["valor"].ToString().Trim();               // codigo doc.venta factura
                            if (row["param"].ToString() == "boleta") codbole = row["valor"].ToString().Trim();               // codigo doc.venta boleta
                            if (row["param"].ToString() == "plazocred") codppc = row["valor"].ToString().Trim();               // codigo plazo de pago x defecto para fact. a CREDITO
                            if (row["param"].ToString() == "usercar_unic") codsuser_cu = row["valor"].ToString().Trim();       // usuarios autorizados a crear Ft de cargas unicas
                            if (row["param"].ToString() == "diasanul") v_cdpa = int.Parse(row["valor"].ToString());            // cant dias en que usuario normal puede anular 
                            if (row["param"].ToString() == "useranul") codusanu = row["valor"].ToString();                      // usuarios autorizados a anular fuera de plazo 
                            if (row["param"].ToString() == "userdscto") cusdscto = row["valor"].ToString();                 // usuarios que pueden hacer descuentos
                            if (row["param"].ToString() == "cltesBol") tdocsBol = row["valor"].ToString();                  // tipos de documento de clientes para boletas
                            if (row["param"].ToString() == "cltesFac") tdocsFac = row["valor"].ToString();                  // tipo de documentos para facturas
                            if (row["param"].ToString() == "limbolsd") limbolsd = decimal.Parse(row["valor"].ToString());   // limite soles para boletas sin direccion
                            if (row["param"].ToString() == "codAnu1.0") codAnu1 = row["valor"].ToString();                  // codigos anulacion erp_grael 1.0
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
                        }
                    }
                    if (row["formulario"].ToString() == "cmasivo") 
                    {
                        if (row["campo"].ToString() == "valores" && row["param"].ToString() == "valin") valini = int.Parse(row["valor"].ToString());            // valor ini del rango
                        if (row["campo"].ToString() == "valores" && row["param"].ToString() == "valfi") valfin = int.Parse(row["valor"].ToString());            // valor fin del rango
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
                // combo de sede
                string consu = "select a.idcodice,a.descrizionerid from desc_sds a where a.numero=@bloq";
                using (MySqlCommand cdv = new MySqlCommand(consu, conn))
                {
                    cdv.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter datv = new MySqlDataAdapter(cdv))
                    {
                        dttd1.Clear();
                        datv.Fill(dttd1);
                        cmb_loca.DataSource = dttd1;
                        cmb_loca.DisplayMember = "descrizionerid";
                        cmb_loca.ValueMember = "idcodice";
                    }
                }
                // datos para el combobox documento de venta
                consu = "select distinct a.idcodice,a.descrizionerid,a.enlace1,a.codsunat,b.glosaser,b.serie " +
                    "from desc_tdo a LEFT JOIN series b ON b.tipdoc = a.IDCodice where a.numero=@bloq and a.codigo=@codv and a.idcodice in (@tipb,@tipf)";
                using (MySqlCommand cdv = new MySqlCommand(consu, conn))
                {
                    cdv.Parameters.AddWithValue("@bloq", 1);
                    cdv.Parameters.AddWithValue("@codv", "venta");
                    cdv.Parameters.AddWithValue("@tipb", codbole);
                    cdv.Parameters.AddWithValue("@tipf", codfact);
                    using (MySqlDataAdapter datv = new MySqlDataAdapter(cdv))
                    {
                        dtdvt.Clear();
                        datv.Fill(dtdvt);
                    }
                }
                //  datos para los combobox de tipo de documento
                using (MySqlCommand cdu = new MySqlCommand("select idcodice,descrizionerid,marca1 as codigo,codsunat from desc_doc where numero=@bloq", conn))
                {
                    cdu.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter datd = new MySqlDataAdapter(cdu))
                    {
                        dttd0.Clear();
                        datd.Fill(dttd0);
                    }
                }
                // datos para el combo de moneda
                using (MySqlCommand cmo = new MySqlCommand("select idcodice,descrizionerid,codsunat,deta1 from desc_mon where numero=@bloq", conn))
                {
                    cmo.Parameters.AddWithValue("@bloq", 1);
                    using (MySqlDataAdapter dacu = new MySqlDataAdapter(cmo))
                    {
                        dtmon.Clear();
                        dacu.Fill(dtmon);
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
            return retorna;
        }
        private bool validGR()                  // validamos y devolvemos datos
        {
            bool retorna = false;
            if (true)
            {
                // validamos que las GRs: No este facturada, No este anulada, destinatario con DNI, cobrada al 100% ..... ya no es valido desde 17/03/2022
                // validamos que las GRs: No este facturada, No este anulada, cobrada al 100% ... 30/04/2022
                using (MySqlConnection conn = new MySqlConnection(db_conn_grael))
                {
                    lib.procConn(conn);
                    if (true)
                    {
                        string parte = "";
                        if (tx_dat_loca.Text != "")
                        {
                            parte = " and a.destino=@dest";
                        }
                        string consulta =
                            "select c.descrizionerid as sede,a.fechope,concat(a.sergre,'-',a.corgre) as guia,d.descrizionerid as DOC,a.numdes as numdoc," +
                            "concat(trim(b.nombre), ' ', trim(b.nombre2)) as cliente,a.doctot as totalGR,a.id,a.docdes,a.moneda," +
                            "a.destino,a.docremi,e.serie as serie,concat(o.descrizionerid, ' - ', c.descrizionerid) as ruta,m.descrizionerid," +
                            "if (b.direc = '-' OR b.direc = ' ',a.dirdest2,b.direc) AS direc," +
                            "if (b.dist = '-' OR b.dist = ' ',a.dirdest2,b.dist) AS dist," +
                            "if (b.provin = '-' OR b.provin = ' ',a.dirdest3,b.provin) AS provin," +
                            "if (b.dpto = '-' OR b.dpto = ' ',a.dirdest4,b.dpto) AS dpto, b.ubigeo " +
                            "from magrem a left join anag_cli b on b.docu = a.docdes and b.ruc = a.numdes " +
                            "LEFT JOIN mactacte t ON t.sergr = a.sergre AND t.corgr = a.corgre " +
                            "left join desc_sds c on c.idcodice = a.destino " +
                            "left join desc_doc d on d.idcodice = a.docdes " +
                            "left join desc_sds o on o.idcodice = a.origen " +
                            "left join desc_mon m on m.idcodice = a.moneda " +
                            "left join series e on e.sede = a.destino and e.tipdoc = @tdbo " +
                            "where a.fechope between @fini and @fina and a.status NOT IN(@canu) and TRIM(t.serdv)= '' and " +
                            "t.totnot > 0 and t.saldo <= 0 and a.moneda = @mone" + parte;
                        using (MySqlCommand micon = new MySqlCommand(consulta, conn))
                        {
                            micon.Parameters.AddWithValue("@fini", dtp_ini.Value.Date.ToString("yyyy-MM-dd"));
                            micon.Parameters.AddWithValue("@fina", dtp_fin.Value.Date.ToString("yyyy-MM-dd"));
                            micon.Parameters.AddWithValue("@canu", codAnu1);    // codigos de anulacion validos
                            //micon.Parameters.AddWithValue("@cdes", vtc_ruc);  // documentos distintos a RUC
                            micon.Parameters.AddWithValue("@mone", MonDeft);    // moneda solo soles
                            micon.Parameters.AddWithValue("@tdbo", "BV");       // serie de BV es igual al de FT
                            if (parte != "") micon.Parameters.AddWithValue("@dest", tx_dat_loca.Text);
                            using (MySqlDataReader dr = micon.ExecuteReader())
                            {
                                while (dr.Read())
                                {
                                    dataGridView1.Rows.Add(
                                        1,
                                        dr.GetString(0),
                                        dr.GetString(1).Substring(0, 10),
                                        dr.GetString(2),
                                        dr.GetString(3),
                                        dr.GetString(4),
                                        dr.GetString(5),
                                        dr.GetString(6),
                                        "",
                                        "", 
                                        dr.GetString(7),
                                        dr.GetString(8),
                                        dr.GetString(9),
                                        dr.GetString(10),
                                        dr.GetString(11),
                                        dr.GetString(12),
                                        dr.GetString(13),
                                        dr.GetString(14),
                                        dr.GetString(15),
                                        dr.GetString(16),
                                        dr.GetString(17),
                                        dr.GetString(18),
                                        dr.GetString(19)
                                        );
                                    // sede,a.fechope,guia,DOC,numdoc,cliente,totalGR,a.id,a.docdes,a.moneda,a.destino,a.docremi,serie,ruta,m.descrizionerid,b.direc,b.dist,b.provin,b.dpto,b.ubigeo
                                    retorna = true;
                                }
                            }
                            totalizaG();
                        }
                    }
                }
            }
            return retorna;
        }
        private void totalizaG()                // calcula totales de la grilla
        {
            int totfil = 0, tfsele = 0;
            double totfle = 0;
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                totfil += 1;
                //MessageBox.Show(dataGridView1.Rows[i].Cells[0].Value.ToString(),"fila: " + i.ToString());
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "1" || dataGridView1.Rows[i].Cells[0].Value.ToString() == "True")
                {
                    totfle += double.Parse(dataGridView1.Rows[i].Cells[7].Value.ToString());
                    tfsele += 1;
                }
            }
            tx_tfil.Text = totfil.ToString();
            tx_flete.Text = totfle.ToString("#0.00");
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
        private void grabfactelec(int ind)                         // graba en la tabla de fact. electrónicas
        {                               // facturacion electrónica con OSE BIZLINKS 10-10-2018  / actualizacion 14/08/2021
            string tipo = "";           //  codbole;   // tx_dat_tdv.Text;
            double ntot = double.Parse(dataGridView1.Rows[ind].Cells[9].Value.ToString());
            double igv = 1.00 + (double.Parse(v_igv) / 100);
            double subt = ntot / igv;   // 1.18
            double migv = ntot - subt;
            // insertamos la cabecera en la tabla del temporal bizlinks
            SqlConnection conms = new SqlConnection(script);
            conms.Open();
            if (conms.State == ConnectionState.Open)
            {
                if (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC") tipo = codfact;
                else tipo = codbole;
                string sernum = dataGridView1.Rows[ind].Cells[8].Value.ToString();                        // v serieNumero 
                string fecemi = tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) +
                    "-" + tx_fechope.Text.Substring(0, 2);                                                 // v 
                DataRow[] row = dtdvt.Select("idcodice='" + tipo + "'");
                string tipdoc = row[0][3].ToString();                                                 // v tipoDocumento
                glosser = row[0][4].ToString();
                DataRow[] rowm = dtmon.Select("idcodice='" + MonDeft + "'");
                string tipmon = rowm[0][2].ToString();                                                 // v tipoMoneda
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
                DataRow[] rowc = dttd0.Select("idcodice='" + dataGridView1.Rows[ind].Cells[11].Value.ToString() + "'");
                string tidoad = rowc[0][3].ToString().Trim();                                           // v tipoDocumentoAdquiriente
                string nudoad = dataGridView1.Rows[ind].Cells[5].Value.ToString();                       // v numeroDocumentoAdquiriente
                string rasoad = dataGridView1.Rows[ind].Cells[6].Value.ToString();                       // v razonSocialAdquiriente
                string coradq = correo_gen;                                                             // v correoAdquiriente
                decimal totimp = Math.Round((decimal)migv, 2);                                            // v totalImpuestos
                decimal totigv = Math.Round((decimal)migv, 2);                                            // v totalIgv
                decimal totvta = Math.Round((decimal)ntot, 2);                                          // v totalVenta
                decimal tovane = Math.Round((decimal)subt, 2);                                          // v totalValorVentaNetoOpGravadas
                // todas las guias, tanto del transportista como de el cliente
                string gpgrael = "";
                string gpadqui = dataGridView1.Rows[ind].Cells[14].Value.ToString();                     // guias del adquiriente
                string codaux40_1 = "9011";                                                             // v codigoAuxiliar40_1
                string etiaux40_1 = "18%";                                                              // v textoAuxiliar40_1
                string tipope = "0101"; // segun rudver, poner esto en una config                       // v tipoOperacion
                string estreg = "A";                                                                    // bl_estadoRegistro
                string coley1 = "1000";                                                                 // v codigoLeyenda_1
                string teley1 = "SON: " + nl.Convertir(ntot.ToString(), true) + row[3].ToString();         // v textoLeyenda_1
                string tiref1 = "";     // detalle
                string nudor1 = "";     // detalle
                string Ctiref1 = "";     // cabecera
                string Cnudor1 = "";     // cabecera
                // ***************** EN OTROS CASOS VAN LAS GUIAS DEL CLIENTE
                string insertcab = "insert into SPE_EINVOICEHEADER (serieNumero,fechaEmision,tipoDocumento,tipoMoneda," +
                    "numeroDocumentoEmisor,tipoDocumentoEmisor,nombreComercialEmisor,razonSocialEmisor,correoEmisor,codigoLocalAnexoEmisor," +
                    "ubigeoEmisor,direccionEmisor,provinciaEmisor,departamentoEmisor,distritoEmisor,urbanizacion,paisEmisor,codigoAuxiliar40_1,textoAuxiliar40_1," +
                    "tipoDocumentoAdquiriente,numeroDocumentoAdquiriente,razonSocialAdquiriente,correoAdquiriente,totalImpuestos," +
                    "totalValorVentaNetoOpGravadas,codigoLeyenda_1,textoLeyenda_1,bl_estadoRegistro," +
                    "totalIgv,totalVenta,tipoOperacion,totalValorVenta,totalPrecioVenta";

                double totdet = 0;
                string leydet = leydet1 + " " + leydet2 + " " + Program.ctadetra;                   // textoLeyenda_2
                insertcab = insertcab + ") " +
                    "values (@sernum,@fecemi,@tipdoc,@tipmon," +
                    "@nudoem,@tidoem,@nocoem,@rasoem,@coremi,@coloem," +
                    "@ubiemi,@diremi,@provemi,@depaemi,@distemi,@urbemi,@pasiemi,@codaux40_1,@etiaux40_1," +
                    "@tidoad,@nudoad,@rasoad,@coradq,@totimp," +
                    "@tovane,@coley1,@teley1,@estreg," +
                    "@totigv,@totvta,@tipope,@tovane,@totvta";
                insertcab = insertcab + ")";
                // primero insertamos el detalle, luego la cabecera
                // *****************************   detalle   *****************************
                {              //  ********************   sin detraccion   ************************** 
                    {
                        glosser2 = dataGridView1.Rows[ind].Cells[16].Value.ToString();
                        string nuori1 = "1";                                                                   // numeroOrdenItem
                        string codprd1 = "-";                                                                   // codigoProducto
                        string coprsu1 = "78101802";                                                            // codigoProductoSunat
                        string descr1 = glosser + " " + glosser2 + " " + tdesc;                                // descripcion
                        decimal canti1 = Math.Round(decimal.Parse("1"), 2);
                        string unime1 = "ZZ";                                                                   // unidadMedida
                        decimal psi1, igv1;                                                                     // calculos de precios x item sin y con impuestos
                        double inuns1 = 0;                                                                      // importeUnitarioSinImpuesto
                        if (decimal.TryParse(dataGridView1.Rows[ind].Cells[9].Value.ToString(), out psi1))
                        {
                            inuns1 = Math.Round(((double)psi1 / ((double)decimal.Parse(v_igv) / 100 + 1)), 2);   // importeUnitarioSinImpuesto
                        }
                        else { inuns1 = 0; }                                                                    // importeUnitarioSinImpuesto
                        decimal inunc1 = Math.Round(decimal.Parse(dataGridView1.Rows[ind].Cells[9].Value.ToString()), 2); // importeUnitarioConImpuesto
                        string coimu1 = "01";                                                                   // codigoImporteUnitarioConImpues
                        string imtoi1 = "";
                        if (decimal.TryParse(dataGridView1.Rows[ind].Cells[9].Value.ToString(), out igv1))
                        {
                            imtoi1 = Math.Round(((double)igv1 - ((double)igv1 / ((double)decimal.Parse(v_igv) / 100 + 1))), 2).ToString();
                        }
                        else { imtoi1 = "0.00"; }                                                               // importeTotalImpuestos 
                        double mobai1 = inuns1;                                                                 // montoBaseIgv
                        string taigv1 = ((decimal.Parse(v_igv))).ToString();                                    // tasaIgv
                        string imigv1 = imtoi1;                                                                 // importeIgv
                        string corae1 = "10";   // grabado operacion onerosa                                    // codigoRazonExoneracion
                        double intos1 = inuns1;                                                                 // importeTotalSinImpuesto
                        gpadqui = dataGridView1.Rows[ind].Cells[14].Value.ToString().Trim();
                        tiref1 = "9840";   // G.R. remitente                                                // codigoAuxiliar500_1
                        nudor1 = gpadqui;
                        gpgrael = dataGridView1.Rows[ind].Cells[3].Value.ToString().Trim();
                        Ctiref1 = "4067";     // G.R. transportista de grael                                  // codigoAuxiliar40_1 
                        Cnudor1 = gpgrael;
                        string ubiPtoOri = "";              // ubigeoPtoOrigen
                        string dirPtoOri = "";              // direccionCompletaPtoOrigen
                        string ubiPtoDes = "";              // ubigeoPtoDestino
                        string dirPtoDes = "";              // direccionCompletaPtoDestino
                        string detViaje = "";               // detalleViaje
                        string monRefSer = "";
                        string monRefCar = "";
                        string monRefUti = "";
                        //
                        string insertadet = "insert into SPE_EINVOICEDETAIL (tipoDocumentoEmisor,numeroDocumentoEmisor,tipoDocumento,serieNumero," +
                            "numeroOrdenItem,codigoProducto,codigoProductoSunat,descripcion,cantidad,unidadMedida,importeTotalSinImpuesto," +
                            "importeUnitarioSinImpuesto,importeUnitarioConImpuesto,codigoImporteUnitarioConImpues,montoBaseIgv,tasaIgv," +
                            "importeIgv,importeTotalImpuestos,codigoRazonExoneracion,codigoAuxiliar40_1,textoAuxiliar40_1,codigoAuxiliar500_1,textoAuxiliar500_1";
                        insertadet = insertadet + ",ubigeoPtoOrigen,direccionCompletaPtoOrigen,ubigeoPtoDestino,direccionCompletaPtoDestino," +
                            "detalleViaje,montoRefServicioTransporte,montoRefCargaEfectiva,montoRefCargaUtilNominal) ";
                        //
                        insertadet = insertadet + "values (@tidoem,@nudoem,@tipdoc,@sernum," +
                            "@nuori1,@codprd1,@coprsu1,@descr1,@canti1,@unime1,@intos1," +
                            "@inuns1,@inunc1,@coimu1,@mobai1,@taigv1," +
                            "@imigv1,@imtoi1,@corae1,@tiref1,@nudor1,@coaux1,@teaux1";
                        insertadet = insertadet + ",@ubiPtoOri,@dirPtoOri,@ubiPtoDes,@dirPtoDes," +
                            "@detViaje,@monRefSer,@monRefCar,@monRefUti)";
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
                            indet.Parameters.AddWithValue("@coaux1", tiref1);                       // codigoAuxiliar500_1
                            indet.Parameters.AddWithValue("@teaux1", nudor1);                       // textoAuxiliar500_1
                            indet.Parameters.AddWithValue("@tiref1", Ctiref1);                      // codigoAuxiliar40_1
                            indet.Parameters.AddWithValue("@nudor1", Cnudor1);                      // textoAuxiliar40_1
                            indet.Parameters.AddWithValue("@ubiPtoOri", ubiPtoOri);                   // ubigeoPtoOrigen
                            indet.Parameters.AddWithValue("@dirPtoOri", dirPtoOri);                   // direccionCompletaPtoOrigen
                            indet.Parameters.AddWithValue("@ubiPtoDes", ubiPtoDes);                   // ubigeoPtoDestino
                            indet.Parameters.AddWithValue("@dirPtoDes", dirPtoDes);                   // direccionCompletaPtoDestino
                            indet.Parameters.AddWithValue("@detViaje", detViaje);                     // detalleViaje
                            indet.Parameters.AddWithValue("@monRefSer", monRefSer);                   // montoRefServicioTransporte
                            indet.Parameters.AddWithValue("@monRefCar", monRefCar);                   // montoRefCargaEfectiva
                            indet.Parameters.AddWithValue("@monRefUti", monRefUti);                   // montoRefCargaUtilNominal
                            //
                            indet.ExecuteNonQuery();
                        }
                        catch (SqlException ex)
                        {
                            MessageBox.Show(ex.Message, "Error en insertar detalle bizlinks");
                            Application.Exit();
                            return;
                        }
                    }
                }
                // insertamos la cabecera
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
                    inserta.Parameters.AddWithValue("@tipope", tipope);
                    inserta.Parameters.AddWithValue("@coley1", coley1);
                    inserta.Parameters.AddWithValue("@teley1", teley1);
                    inserta.Parameters.AddWithValue("@estreg", estreg);
                    inserta.Parameters.AddWithValue("@coddetra", Program.coddetra);
                    inserta.Parameters.AddWithValue("@totdet", totdet);
                    inserta.Parameters.AddWithValue("@pordetra", Program.pordetra);
                    inserta.Parameters.AddWithValue("@ctadetra", Program.ctadetra);
                    inserta.Parameters.AddWithValue("@codleyt", codleyt);
                    inserta.Parameters.AddWithValue("@leydet", leydet);
                    inserta.ExecuteNonQuery();
                }
                catch (SqlException ex)
                {
                    MessageBox.Show(ex.Message, "Error en insertar cabecera bizlinks");
                    Application.Exit();
                    return;
                }
                // adicionales cabecera     SPE_EINVOICEHEADER_ADD
                string diradq = dataGridView1.Rows[ind].Cells[18].Value.ToString().Trim();
                string ubiadq = dataGridView1.Rows[ind].Cells[22].Value.ToString().Trim();
                string urbadq = "-";
                string proadq = dataGridView1.Rows[ind].Cells[20].Value.ToString().Trim();
                string depadq = dataGridView1.Rows[ind].Cells[21].Value.ToString().Trim();
                string disadq = dataGridView1.Rows[ind].Cells[19].Value.ToString().Trim();
                string paiadq = "PE";
                string formpa = "0";        // TODAS CONTADO
                string insadd = "insert into SPE_EINVOICEHEADER_ADD (" +
                "clave,numeroDocumentoEmisor,serieNumero,tipoDocumento,tipoDocumentoEmisor,valor) values ";
                insadd = insadd + "('direccionAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@diradq)," +
                        "('provinciaAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@proadq)," +
                        "('departamentoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@depadq)," +
                        "('distritoAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@disadq)," +
                        "('paisAdquiriente',@nudoem,@sernum,@tipdoc,@tidoem,@paiadq)";
                if (tipo == codfact)                                                         // esto aplica solo para
                {
                    insadd = insadd + ",('formaPago',@nudoem,@sernum,@tipdoc,@tidoem,@fpagoD)";
                    if (formpa == "0") insadd = insadd + ",('formaPagoNegociable',@nudoem,@sernum,@tipdoc,@tidoem,@formpa)";
                }
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
                if (tipo == codfact)
                {
                    insertadd.Parameters.AddWithValue("@formpa", formpa);
                    insertadd.Parameters.AddWithValue("@fpagoD", "009");
                }
                insertadd.ExecuteNonQuery();
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

        #region limpiadores_modos
        private void sololee()
        {
            lp.sololee(this);
        }
        private void escribe()
        {
            lp.escribe(this);
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
        }
        #endregion limpiadores_modos;

        #region boton_form GRABA EDITA ANULA
        private void bt_agr_Click(object sender, EventArgs e)
        {
            if (Tx_modo.Text == "NUEVO")
            {
                // fecha final sea > fecha inicial
                if (dtp_ini.Value > dtp_fin.Value)
                {
                    MessageBox.Show("La fecha inicial debe ser menor", "Error en fechas", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    dtp_ini.Focus();
                    return;
                }
                dataGridView1.Rows.Clear();
                grilla();
                if (validGR() == false)
                {
                    return;
                }
                if (int.Parse(tx_tfil.Text) > 0)
                {
                    // si hay guias por boletear, avisamos
                }
                else
                {
                    MessageBox.Show("No existen documentos pendientes","Atención",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    cmb_loca.Focus();
                    return;
                }
                dataGridView1.Focus();
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            #region validaciones
            totalizaG();
            for (int i = 0; i < dataGridView1.Rows.Count - 1; i++)
            {
                if (dataGridView1.Rows[i].Cells[8].Value.ToString().Trim() != "")
                {
                    MessageBox.Show("Debe cerrar y volver a abrir el formulario");
                    this.Close();
                    return;
                }
            }
            #endregion
            // grabamos, actualizamos, etc
            string modo = Tx_modo.Text;
            string iserror = "no";
            if (modo == "NUEVO")
            {
                // valida pago y calcula
                if (true)
                {
                    var aa = MessageBox.Show("Confirma que desea crear los documentos?", "Confirme por favor", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (aa == DialogResult.Yes)
                    {
                        for (int i=0; i<dataGridView1.Rows.Count - 1; i++)
                        {
                            if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "1" || dataGridView1.Rows[i].Cells[0].Value.ToString() == "True")
                            {
                                if (graba(i) == true)
                                {
                                    OSE_Fact_Elect();
                                    grabfactelec(i);
                                }
                                else
                                {
                                    MessageBox.Show("No se puede grabar el documento de venta electrónico", "Error en conexión");
                                    iserror = "si";
                                }
                            }
                        }
                    }
                    else
                    {
                        return;
                    }
                }
            }
            if (modo == "EDITAR")
            {

            }
            if (modo == "ANULAR")
            {

            }
            if (iserror == "no")
            {
                /*
                string resulta = lib.ult_mov(nomform, nomtab, asd);
                if (resulta != "OK")                                        // actualizamos la tabla usuarios
                {
                    MessageBox.Show(resulta, "Error en actualización de tabla usuarios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                */
            }
            //initIngreso();          // limpiamos todo para volver a empesar
        }
        private bool graba(int ind)
        {
            bool retorna = false;
            MySqlConnection conn = new MySqlConnection(db_conn_grael);   // DB_CONN_STR
            conn.Open();
            if(conn.State == ConnectionState.Open)
            {
                string todo = "corre_serie";
                using (MySqlCommand micon = new MySqlCommand(todo, conn))
                {
                    micon.CommandType = CommandType.StoredProcedure;
                    micon.Parameters.AddWithValue("td", (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC") ? "FT" : "BV"); 
                    micon.Parameters.AddWithValue("ser", dataGridView1.Rows[ind].Cells[15].Value.ToString()); // serie del local de destino
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
                // calculos 
                Random rand = new Random();
                int ntot = rand.Next(valini, valfin); // valor entre 5 y 8    
                double igv = 1.00 + (double.Parse(v_igv) / 100);
                double subt = ntot / igv;   // 1.18
                double migv = ntot - subt;
                string inserta = "insert into madocvtas (fechope,tipcam,docvta,servta,corvta,doccli,numdcli,direc,nomclie," +
                    "observ,moneda,aumigv,subtot,igv,doctot,status,pigv,userc,fechc,condpag," +                     //      tippago,
                    "local,rd3,dist,prov,dpto,saldo,cdr,mfe,email,ubiclte,canfild,tippago,plazocred,pagauto," +     // 
                    "prop01,prop02,prop03,prop04,prop05,prop06,prop07,prop08,prop09,prop10) " +
                    "values (" +
                    "@fechop,@tcoper,@ctdvta,@serdv,@numdv,@tdcrem,@ndcrem,@dircre,@nomrem," +
                    "@obsprg,@monppr,@aig,@subpgr,@igvpgr,@totpgr,@estpgr,@porcigv,@asd,now(),@copag," +            //      @tipap,
                    "@ldcpgr,@tcdvta,@distcl,@provcl,@dptocl,@salxpa,@cdr,@mtdvta,@mailcl,@ubicre,@canfil,@tippa,@plzoc,@pagau," +       // 
                    "@pr01,@pr02,@pr03,@pr04,@pr05,@pr06,@pr07,@pr08,@pr09,@pr10)";
                using (MySqlCommand micon = new MySqlCommand(inserta, conn))
                {
                    micon.Parameters.AddWithValue("@fechop", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                    micon.Parameters.AddWithValue("@tcoper", 0);                   // TIPO DE CAMBIO
                    micon.Parameters.AddWithValue("@ctdvta", (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC") ? codfact : codbole);
                    micon.Parameters.AddWithValue("@serdv", dataGridView1.Rows[ind].Cells[15].Value.ToString());
                    micon.Parameters.AddWithValue("@numdv", tx_numero.Text);
                    micon.Parameters.AddWithValue("@tdcrem", dataGridView1.Rows[ind].Cells[11].Value.ToString());
                    micon.Parameters.AddWithValue("@ndcrem", dataGridView1.Rows[ind].Cells[5].Value.ToString());
                    micon.Parameters.AddWithValue("@dircre", "-");
                    micon.Parameters.AddWithValue("@nomrem", dataGridView1.Rows[ind].Cells[6].Value.ToString());
                    micon.Parameters.AddWithValue("@obsprg", "masivo");
                    micon.Parameters.AddWithValue("@monppr", dataGridView1.Rows[ind].Cells[12].Value.ToString());
                    micon.Parameters.AddWithValue("@aig", "0");             // aumenta igv ??, aca no hay esa opcion
                    micon.Parameters.AddWithValue("@subpgr", subt.ToString("#0.##"));                     // sub total
                    micon.Parameters.AddWithValue("@igvpgr", migv.ToString("#0.##"));                      // igv
                    micon.Parameters.AddWithValue("@totpgr", ntot.ToString("#0.##"));                    // total inc. igv
                    micon.Parameters.AddWithValue("@estpgr", codCanc);      // estado
                    micon.Parameters.AddWithValue("@porcigv", v_igv);                           // porcentaje en numeros de IGV
                    micon.Parameters.AddWithValue("@asd", asd);
                    //micon.Parameters.AddWithValue("@tipap", v_mpag);             
                    micon.Parameters.AddWithValue("@copag", "R");              // (rb_si.Checked == true) ? "E" : (rb_no.Checked == true) ? "R" : "C"
                    micon.Parameters.AddWithValue("@ldcpgr", dataGridView1.Rows[ind].Cells[13].Value.ToString());         // local origen
                    micon.Parameters.AddWithValue("@tcdvta", "2");          // destinatario
                    micon.Parameters.AddWithValue("@distcl", "-");
                    micon.Parameters.AddWithValue("@provcl", "-");
                    micon.Parameters.AddWithValue("@dptocl", "-");
                    micon.Parameters.AddWithValue("@salxpa", "0");
                    micon.Parameters.AddWithValue("@cdr", "9");             // cdr en la base actual tiene valor 9
                    micon.Parameters.AddWithValue("@mtdvta", (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC")? "F" : "B");
                    micon.Parameters.AddWithValue("@mailcl", correo_gen);         // "tx_email.Text"
                    micon.Parameters.AddWithValue("@ubicre", "");
                    micon.Parameters.AddWithValue("@canfil", 1);            // cant filas de detalle, como son boletas masivas les ponemos 1 a todos
                    micon.Parameters.AddWithValue("@tippa", "");
                    micon.Parameters.AddWithValue("@plzoc", "");
                    micon.Parameters.AddWithValue("@pagau", "N");    // cobranza automatica en efectivo NO porque ya estan cobradas
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
                // detalle
                tcbul = "0";
                tdesc = "-";
                tunid = "-";

                string condeta = "select codprd,descrip,unidad,sum(cantid) as cant from erp_grael.detagrem where idc=@idr group by idc";
                using (MySqlCommand micon = new MySqlCommand(condeta, conn))
                {
                    micon.Parameters.AddWithValue("@idr", dataGridView1.Rows[ind].Cells[10].Value.ToString());
                    using (MySqlDataReader dr = micon.ExecuteReader())
                    {
                        if (dr.Read())
                        {
                            tcbul = dr.GetString(3);
                            tdesc = (dr.GetString(0) + " " + dr.GetString(1)).Trim();
                            tunid = dr.GetString(2);
                        }
                    }
                }

                string inserd2 = "insert into detavtas (idc,docvta,servta,corvta,sergr,corgr,moneda,valor,ruta," +
                    "glosa,status,userc,fechc,docremi,bultos,monrefd1,monrefd2,monrefd3,mfe,fecdoc,totaldoc,valorel) " +
                    "values (@idr,@doc,@svt,@cvt,@sgui,@cgui,@codm,@pret,@ruta," +
                    "@desc,@sta,@asd,now(),@dre,@bult,@mrd1,@mrd2,@mrd3,@mtdvta,@fechop,@totpgr,@velfi)";
                using (MySqlCommand micon = new MySqlCommand(inserd2, conn))
                {
                    micon.Parameters.AddWithValue("@idr", tx_idr.Text);
                    micon.Parameters.AddWithValue("@doc", (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC") ? codfact : codbole);    // codbole
                    micon.Parameters.AddWithValue("@svt", dataGridView1.Rows[ind].Cells[15].Value.ToString());
                    micon.Parameters.AddWithValue("@cvt", tx_numero.Text);
                    micon.Parameters.AddWithValue("@sgui", dataGridView1.Rows[ind].Cells[3].Value.ToString().Trim().Substring(0,3));
                    micon.Parameters.AddWithValue("@cgui", dataGridView1.Rows[ind].Cells[3].Value.ToString().Trim().Substring(4,7));
                    micon.Parameters.AddWithValue("@codm", dataGridView1.Rows[ind].Cells[17].Value.ToString());   // dataGridView1.Rows[ind].Cells[12].Value.ToString()
                    micon.Parameters.AddWithValue("@pret", dataGridView1.Rows[ind].Cells[7].Value.ToString());  // flete de la gr
                    micon.Parameters.AddWithValue("@ruta", dataGridView1.Rows[ind].Cells[16].Value.ToString());
                    micon.Parameters.AddWithValue("@desc", tdesc);
                    micon.Parameters.AddWithValue("@sta", codCanc); // estado;
                    micon.Parameters.AddWithValue("@asd", asd);
                    micon.Parameters.AddWithValue("@dre", dataGridView1.Rows[ind].Cells[14].Value.ToString());
                    micon.Parameters.AddWithValue("@bult", tcbul + " " + tunid);
                    micon.Parameters.AddWithValue("@mrd1", "0.00"); // (tx_dref1.Text == "") ? "0.00" : tx_dref1.Text
                    micon.Parameters.AddWithValue("@mrd2", "0.00"); // (tx_dcar1.Text == "") ? "0.00" : tx_dcar1.Text
                    micon.Parameters.AddWithValue("@mrd3", "0.00");   // (tx_dnom1.Text == "") ? "0.00" : tx_dnom1.Text
                    micon.Parameters.AddWithValue("@mtdvta", (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC") ? "F" : "B");          // todas son B de boleta
                    micon.Parameters.AddWithValue("@fechop", tx_fechope.Text.Substring(6, 4) + "-" + tx_fechope.Text.Substring(3, 2) + "-" + tx_fechope.Text.Substring(0, 2));
                    micon.Parameters.AddWithValue("@totpgr", ntot.ToString("#0.##"));                    // total inc. igv
                    micon.Parameters.AddWithValue("@velfi", ntot.ToString("#0.##"));  // valor fila fact electronica
                    //micon.Parameters.AddWithValue("@fila", fila);
                    /*
                    micon.Parameters.AddWithValue("@unim", "");
                    micon.Parameters.AddWithValue("@cmnn", dataGridView1.Rows[i].Cells[6].Value.ToString());
                    micon.Parameters.AddWithValue("@tgrmn", dataGridView1.Rows[i].Cells[5].Value.ToString());
                    micon.Parameters.AddWithValue("@pagaut", (rb_si.Checked == true) ? "S" : "N");
                    */
                    micon.ExecuteNonQuery();
                    //fila = 1;
                    if (dataGridView1.Rows[ind].Cells[4].Value.ToString() == "RUC") dataGridView1.Rows[ind].Cells[8].Value = "F" + dataGridView1.Rows[ind].Cells[15].Value.ToString() + "-" + "0" + tx_numero.Text;
                    else dataGridView1.Rows[ind].Cells[8].Value = "B" + dataGridView1.Rows[ind].Cells[15].Value.ToString() + "-" + "0" + tx_numero.Text;
                    dataGridView1.Rows[ind].Cells[9].Value = ntot.ToString("#0.##");

                    retorna = true;         // no hubo errores!
                    //
                    // OJO, las actualizaciones en las tablas magrem, manoen y mactacte, las hace el TRIGGER de la tabla detvtas
                    // 
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
        #endregion boton_form;

        #region leaves y checks

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
                    consulb.Parameters.AddWithValue("@nomform", "cmasivo");
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
            Bt_add.Visible = false;
            Bt_edit.Visible = false;
            Bt_anul.Visible = false;
            Bt_ver.Visible = false;
            Bt_print.Visible = false;

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
            // botones inicio, fin, siguiente y regresa .... no permitidos aqui
            Bt_fin.Visible = false;
            Bt_ini.Visible = false;
            Bt_sig.Visible = false;
            Bt_ret.Visible = false;
        }
        #region botones
        private void Bt_add_Click(object sender, EventArgs e)
        {
            Tx_modo.Text = "NUEVO";
            escribe();
            // 
            Bt_ini.Enabled = false;
            Bt_sig.Enabled = false;
            Bt_ret.Enabled = false;
            Bt_fin.Enabled = false;
            initIngreso();
        }
        private void Bt_edit_Click(object sender, EventArgs e)
        {
            sololee();          
            Tx_modo.Text = "EDITAR";
            initIngreso();
            //
            Bt_ini.Enabled = true;
            Bt_sig.Enabled = true;
            Bt_ret.Enabled = true;
            Bt_fin.Enabled = true;
        }
        private void Bt_close_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void Bt_print_Click(object sender, EventArgs e)
        {
            // Impresion ó Re-impresion ??
        }
        private void Bt_anul_Click(object sender, EventArgs e)
        {

        }
        private void Bt_ver_Click(object sender, EventArgs e)
        {
            sololee();
            Tx_modo.Text = "VISUALIZAR";
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
        }
        private void Bt_last_Click(object sender, EventArgs e)
        {
            limpiar();
            limpia_chk();
            limpia_combos();
            limpia_otros();
            tx_idr.Text = lib.golast(nomtab);
        }
        #endregion botones;
        // proveed para habilitar los botones de comando
        #endregion botones_de_comando  ;

        #region comboboxes
        private void cmb_loca_SelectionChangeCommitted(object sender, EventArgs e)
        {
            tx_dat_loca.Text = cmb_loca.SelectedValue.ToString();
        }
        private void cmb_loca_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                cmb_loca.SelectedIndex = -1;
                tx_dat_loca.Text = "";
            }
        }

        #endregion comboboxes

        #region datagrid
        private void dataGridView1_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            totalizaG();
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            /*if (e.ColumnIndex == 0)
            {
                if (dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString() == "1")
                {
                    //dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.White;
                    
                }
                else
                {
                    dataGridView1.Rows[e.RowIndex].DefaultCellStyle.BackColor = Color.Red;
                }
            }*/
            //totalizaG();
        }

        #endregion

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void dataGridView1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            //MessageBox.Show("validated");
            totalizaG();
        }

        private void dataGridView1_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            //MessageBox.Show("validating");
            //totalizaG();
        }
    }
}
