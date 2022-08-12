using System;
using System.Windows.Forms;
using System.Diagnostics;
using Reportes.MisClases;


namespace Reportes
{
    public partial class Form1 : Form
    {
        string Direccion = Application.StartupPath;
        Configuracion Config = new Configuracion();
        Servicios EjecutarServicios = new Servicios();
        PINES EjecutarPines = new PINES();
        //Recargas EjecutarRecargas = new Recargas();
        //RecargaDOS EjecutarRecargas = new RecargaDOS();
        RecargasTRES EjecutarRecargas = new RecargasTRES();

        Exportar exp = new Exportar();
        ExportarRecargas expR = new ExportarRecargas();
        /**********************************************************************************************************************************************/
        string[] tienda;
        string[] fecha;
        string[] hora;
        string[] vendedor;
        string[] clasificacion;
        string[] concepto;
        string[] referencia;
        string[] importe;
        /**********************************************************************************************************************************************/
        public string[,] CatalogoCajas;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try{
                Config.CargarConfiguracion(TxtP_Tienda, TxtP_Fecha, TxtP_Hora, TxtP_Vendedor, TxtP_Clasificacion, TxtP_Concepto, TxtP_Referencia, TxtP_Importe, TxtP_Conceptos, TxtM_Conceptos, Direccion, TxtM_Tienda, TxtM_Fecha, TxtM_Hora, TxtM_Vendedor, TxtM_Clasificacion, TxtM_Concepto, TxtM_Referencia, TxtM_Importe);
                Config.CargarGV(DGVCajas, Direccion + "//Catalogo.xlsx", "Cajas");
                Config.CargarGV(DGVCatalogo, Direccion + "//Catalogo.xlsx", "Catalogo");
                Config.CargarGV(DGVPines, Direccion + "//Catalogo.xlsx", "Pines");
                Config.CargarGV(DGVRecargas, Direccion + "//Catalogo.xlsx", "Recargas");
                Config.CargarGV(DGVSupervisores, Direccion + "//Catalogo.xlsx", "Supervisores");                               
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja", ex.Message);
            }
        }
        public void Configuracione() {
            tienda = TxtM_Tienda.Text.Split('*');
            fecha = TxtM_Fecha.Text.Split('*');
            hora = TxtM_Hora.Text.Split('*');
            vendedor = TxtM_Vendedor.Text.Split('*');
            clasificacion = TxtM_Clasificacion.Text.Split('*');
            concepto = TxtM_Concepto.Text.Split('*');
            referencia = TxtM_Referencia.Text.Split('*');
            importe = TxtM_Importe.Text.Split('*');
        }

        private void BtnIniciar_Click(object sender, EventArgs e)
        {
            Configuracione();
            string ColumnasPlataforma = TxtP_Tienda.Text + "/" + TxtP_Fecha.Text + "/" + TxtP_Hora.Text + "/" + TxtP_Vendedor.Text + "/" + TxtP_Clasificacion.Text + "/" + TxtP_Concepto.Text + "/" + TxtP_Referencia.Text + "/" + TxtP_Importe.Text;
            string ColumnasManejador = tienda[0].ToString() + "/" + fecha[0].ToString() + "/" + hora[0].ToString() + "/" + vendedor[0].ToString() + "/" + clasificacion[0].ToString() + "/" + concepto[0].ToString() + "/" + referencia[0].ToString() + "/" + importe[0].ToString();
            string[] M_Concetpos = TxtM_Conceptos.Text.Split('/');
            EjecutarServicios.ConciliarServicios(TxtP_Conceptos.Text,ColumnasPlataforma, M_Concetpos[1].ToString(), ColumnasManejador,  PBar, DGVManager, DGVCatalogo, DGVMTCenter, DGVCajas, DGVMatch, DGVNoMatch);
            btnExportar.Enabled = true;
            BtnIniciar.Enabled = false;
        }
        private void BTNMatchPines_Click(object sender, EventArgs e)
        {
            Configuracione();
            string ColumnasPlataforma = TxtP_Tienda.Text + "/" + TxtP_Fecha.Text + "/" + TxtP_Hora.Text + "/" + TxtP_Vendedor.Text + "/" + TxtP_Clasificacion.Text + "/" + TxtP_Concepto.Text + "/" + TxtP_Referencia.Text + "/" + TxtP_Importe.Text;
            string ColumnasManejador = tienda[0].ToString() + "/" + fecha[0].ToString() + "/" + hora[0].ToString() + "/" + vendedor[0].ToString() + "/" + clasificacion[0].ToString() + "/" + concepto[0].ToString() + "/" + referencia[0].ToString() + "/" + importe[0].ToString();
            string[] M_Concetpos = TxtM_Conceptos.Text.Split('/');
            EjecutarPines.ConciliarPines(TxtP_Conceptos.Text, ColumnasPlataforma, M_Concetpos[2].ToString(), ColumnasManejador, PBar, DGVManager, DGVPines, DGVMTCenter, DGVCajas, DGVPinesMatch, DGVPinesNoMatch);
        }
        private void btnExportar_Click(object sender, EventArgs e)
        {            
            exp.ExportarDataGridViewExcel(DGVNoMatch, "NO Empatados ");
            exp.ExportarDataGridViewExcel(DGVMatch, "Empatados ");   
        }
        private void BTNMatchRecargas_Click(object sender, EventArgs e)
        {
            Configuracione();
            string ColumnasPlataforma = TxtP_Tienda.Text + "/" + TxtP_Fecha.Text + "/" + TxtP_Hora.Text + "/" + TxtP_Vendedor.Text + "/" + TxtP_Clasificacion.Text + "/" + TxtP_Concepto.Text + "/" + TxtP_Referencia.Text + "/" + TxtP_Importe.Text;
            string ColumnasManejador = tienda[1].ToString() + "/" + fecha[1].ToString() + "/" + hora[1].ToString() + "/" + vendedor[1].ToString() + "/" + clasificacion[1].ToString() + "/" + concepto[1].ToString() + "/" + referencia[1].ToString() + "/" + importe[1].ToString();
            string[] M_Concetpos = TxtM_Conceptos.Text.Split('/');
            EjecutarRecargas.ConciliarRecargas(PBar, ColumnasManejador, DGVManager, DGVTicketsFound, DGVTicketsNotFound, DGVEmpates, DGVMatchRecargas, DGVNoMachtRecargas, DGVTempRECARGAS, DGVCajas, DGVRecargas, DGVTickets, DGVReporte, DGVMTCenter);
            BtnMATCHRecargas.Enabled = false;
            BtnExportarRecargas.Enabled = true;
        }
        private void BtnExportarRecargas_Click(object sender, EventArgs e)
        {
            expR.ExportarDataGridViewExcel(DGVNoMachtRecargas, "NO Empatados ");
            expR.ExportarDataGridViewExcel(DGVMatchRecargas, "Empatados ");
        }

        private void btnExpRecargas_Click(object sender, EventArgs e)
        {            
            expR.ExportarDataGridViewExcel(DGVNoMachtRecargas, "NO Empatados ");
            expR.ExportarDataGridViewExcel(DGVMatchRecargas, "Empatados ");
        }
        private void BtnComisiones_Click(object sender, EventArgs e)
        {
           // EjecutarComisionesIUSA.ConciliarComisiones(PBar, DGVComisionesMNGR, DGVIUSAR, DGVNoMAchtIUSA, DGVMatchComisiones, DGVCajas, DGVComiOtros);
            //EjecutarComisiones.ConciliarComisiones(PBar, DGVComisionesMNGR, DGVIUSAR, DGVNoMAchtIUSA, DGVMatchComisiones, DGVCajas, DGVComiOtros);
        }
        private void button1_Click(object sender, EventArgs e)
        {
            expR.ExportarDataGridViewExcel(DGVNoMAchtIUSA, "NO Empatados ");
            expR.ExportarDataGridViewExcel(DGVMatchComisiones, "Empatados ");
        }
        private void BtnAbriCat_Click(object sender, EventArgs e)
        {
            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.FileName = "EXCEL.EXE";
            startInfo.Arguments = @Direccion + "\\Catalogo.xlsx";
            Process.Start(startInfo);
        }
        
        private void button2_Click(object sender, EventArgs e)
        {
            expR.ExportarDataGridViewExcel(DGVPinesNoMatch, "NO Empatados ");
            expR.ExportarDataGridViewExcel(DGVPinesMatch, "Empatados ");
        }
        private void BtnGuardarConfig_Click(object sender, EventArgs e)
        {
            string ColumnasPlataforma = TxtP_Tienda.Text + "/" + TxtP_Fecha.Text + "/" + TxtP_Hora.Text + "/" + TxtP_Vendedor.Text + "/" + TxtP_Clasificacion.Text + "/" + TxtP_Concepto.Text + "/" + TxtP_Referencia.Text + "/" + TxtP_Importe.Text;
            string ColumnasManegador = TxtM_Tienda.Text + "/" + TxtM_Fecha.Text + "/" + TxtM_Hora.Text + "/" + TxtM_Vendedor.Text + "/" + TxtM_Clasificacion.Text + "/" + TxtM_Concepto.Text + "/" + TxtM_Referencia.Text + "/" + TxtM_Importe.Text;
            string Conceptos = TxtP_Conceptos.Text;
            string ConceptosManejador = TxtM_Conceptos.Text;
            Config.SaveConfig(ColumnasPlataforma, ColumnasManegador, Conceptos, ConceptosManejador, TxtP_Tienda, TxtP_Fecha, TxtP_Hora, TxtP_Vendedor, TxtP_Clasificacion, TxtP_Concepto, TxtP_Referencia, TxtP_Importe, TxtP_Conceptos, TxtM_Conceptos,Direccion, TxtM_Tienda, TxtM_Fecha, TxtM_Hora, TxtM_Vendedor, TxtM_Clasificacion, TxtM_Concepto, TxtM_Referencia, TxtM_Importe);
        }

        private void BtnAdjuntarCat_Click(object sender, EventArgs e)
        {

        }

        
    }
}
