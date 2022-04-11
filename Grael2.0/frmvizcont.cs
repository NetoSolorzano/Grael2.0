using System;
using CrystalDecisions.CrystalReports.Engine;
using System.Windows.Forms;

namespace Grael2
{
    public partial class frmvizcont : Form
    {
        conClie _datosReporte;

        private frmvizcont()
        {
            InitializeComponent();
        }

        public frmvizcont(conClie datos): this()
        {
            _datosReporte = datos;
        }

        private void frmvizcont_Load(object sender, EventArgs e)
        {
            /*
            if (_datosReporte.cuadreCaja_cab.Rows.Count > 0)
            {
                string nf = _datosReporte.cuadreCaja_cab.Rows[0].ItemArray[0].ToString();
                ReportDocument rpt = new ReportDocument();
                rpt.Load(nf);   // rpt.Load("formatos/cuadreCaja1.rpt");
                rpt.SetDataSource(_datosReporte);
                crystalReportViewer1.ReportSource = rpt;
            }
            */
            if (_datosReporte.ventasCab.Rows.Count > 0)
            {
                string nf = _datosReporte.ventasCab.Rows[0].ItemArray[1].ToString();
                ReportDocument rpt = new ReportDocument();
                rpt.Load(nf);
                rpt.SetDataSource(_datosReporte);
                crystalReportViewer1.ReportSource = rpt;
            }
        }
    }
}
