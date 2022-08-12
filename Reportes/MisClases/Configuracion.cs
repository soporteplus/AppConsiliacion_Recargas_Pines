using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class Configuracion
    {
        //identifica la ruta desde donde se esta ejecutando el archivo

        string ColumnasPlataforma = "tienda/fechaHora/fechaHora/cajero/clasificacion/producto/referencia1/entrada";
        string ColumnasManegador = "Caja*Área de Negocio Factura/Fecha*Fecha/Hora*Hora/Vendedor*Nombre Vendedor/Clasificacion*Clasificacion/Comentario*Descripción/Concepto*Concepto/*";
        string Conceptos = "TAE/Pago Servicio/Pines";
        string ConceptosManejador = "RECARGA*TELCEL/RECIBO*PAGO*ELIJA UNA OPCION/PIN*ELIJA UNA OPCION";
        //tienda        = Caja
        //fechaHora     = Fecha + Hora
        //cajero        = Vendedor
        //clasificacion = *para saber si es Tiempo Aire E(TAE), Pines o Pago Servicio
        //producto      = Concepto
        //referencia1   = Comentario
        //entrada       = IMPORTE; 

        public void CargarConfiguracion(TextBox TxtP_Tienda, TextBox TxtP_Fecha, TextBox TxtP_Hora, TextBox TxtP_Vendedor, TextBox TxtP_Clasificacion, TextBox TxtP_Concepto, TextBox TxtP_Referencia, TextBox TxtP_Importe, TextBox TxtP_Conceptos, TextBox TxtM_Conceptos, string Dir, TextBox TxtM_Tienda, TextBox TxtM_Fecha, TextBox TxtM_Hora, TextBox TxtM_Vendedor, TextBox TxtM_Clasificacion, TextBox TxtM_Concepto, TextBox TxtM_Referencia, TextBox TxtM_Importe)
        {
            if (File.Exists(Dir + "\\ConfigReportes.ini"))
            {
                LoadConfig(TxtP_Tienda, TxtP_Fecha, TxtP_Hora, TxtP_Vendedor, TxtP_Clasificacion, TxtP_Concepto, TxtP_Referencia, TxtP_Importe, TxtP_Conceptos, TxtM_Conceptos, Dir, TxtM_Tienda, TxtM_Fecha, TxtM_Hora, TxtM_Vendedor, TxtM_Clasificacion, TxtM_Concepto, TxtM_Referencia, TxtM_Importe);
            }
            else
            {
                SaveConfig(ColumnasPlataforma, ColumnasManegador, Conceptos, ConceptosManejador, TxtP_Tienda, TxtP_Fecha, TxtP_Hora, TxtP_Vendedor, TxtP_Clasificacion, TxtP_Concepto, TxtP_Referencia, TxtP_Importe, TxtP_Conceptos, TxtM_Conceptos, Dir, TxtM_Tienda, TxtM_Fecha, TxtM_Hora, TxtM_Vendedor, TxtM_Clasificacion, TxtM_Concepto, TxtM_Referencia, TxtM_Importe); 
            }
        }
        public void LoadConfig(TextBox TxtP_Tienda, TextBox TxtP_Fecha, TextBox TxtP_Hora, TextBox TxtP_Vendedor, TextBox TxtP_Clasificacion, TextBox TxtP_Concepto, TextBox TxtP_Referencia, TextBox TxtP_Importe, TextBox TxtP_Conceptos, TextBox TxtM_Conceptos, string Dir, TextBox TxtM_Tienda, TextBox TxtM_Fecha, TextBox TxtM_Hora, TextBox TxtM_Vendedor, TextBox TxtM_Clasificacion, TextBox TxtM_Concepto, TextBox TxtM_Referencia, TextBox TxtM_Importe)
        {
            string[] lines = File.ReadAllLines(Dir + "\\ConfigReportes.ini");
            string Plataforma = lines[0].ToString();
            string[] ColPlataforma = Plataforma.Split('/');
            TxtP_Tienda.Text = ColPlataforma[0].ToString();
            TxtP_Fecha.Text = ColPlataforma[1].ToString();
            TxtP_Hora.Text = ColPlataforma[2].ToString();
            TxtP_Vendedor.Text = ColPlataforma[3].ToString();
            TxtP_Clasificacion.Text = ColPlataforma[4].ToString();
            TxtP_Concepto.Text = ColPlataforma[5].ToString();
            TxtP_Referencia.Text = ColPlataforma[6].ToString();
            TxtP_Importe.Text = ColPlataforma[7].ToString();

            string Manejador = lines[1].ToString();
            string[] ColManejador = Manejador.Split('/');
            TxtM_Tienda.Text = ColManejador[0].ToString();
            TxtM_Fecha.Text = ColManejador[1].ToString();
            TxtM_Hora.Text = ColManejador[2].ToString();
            TxtM_Vendedor.Text = ColManejador[3].ToString();
            TxtM_Clasificacion.Text = ColManejador[4].ToString();
            TxtM_Concepto.Text = ColManejador[5].ToString();
            TxtM_Referencia.Text = ColManejador[6].ToString();
            TxtM_Importe.Text = ColManejador[7].ToString();

            string ConceptosPlataforma = lines[2].ToString();
            TxtP_Conceptos.Text = ConceptosPlataforma;
            TxtM_Conceptos.Text = lines[3].ToString();
        }
        public void SaveConfig(string Plataforma, string Manejador, string Conceptos, string ConceptosManejador, TextBox TxtP_Tienda, TextBox TxtP_Fecha, TextBox TxtP_Hora, TextBox TxtP_Vendedor, TextBox TxtP_Clasificacion, TextBox TxtP_Concepto, TextBox TxtP_Referencia, TextBox TxtP_Importe,TextBox TxtP_Conceptos,TextBox TxtM_Conceptos, string Dir, TextBox TxtM_Tienda, TextBox TxtM_Fecha, TextBox TxtM_Hora, TextBox TxtM_Vendedor, TextBox TxtM_Clasificacion, TextBox TxtM_Concepto, TextBox TxtM_Referencia, TextBox TxtM_Importe)
        {
            string[] lines = {  Plataforma, Manejador,Conceptos, ConceptosManejador};
            File.WriteAllLines(Dir + "\\ConfigReportes.ini", lines);
            MessageBox.Show("Se realizaron los cambios con exito");
            LoadConfig(TxtP_Tienda, TxtP_Fecha, TxtP_Hora, TxtP_Vendedor, TxtP_Clasificacion, TxtP_Concepto, TxtP_Referencia, TxtP_Importe, TxtP_Conceptos, TxtM_Conceptos, Dir, TxtM_Tienda, TxtM_Fecha, TxtM_Hora, TxtM_Vendedor, TxtM_Clasificacion, TxtM_Concepto, TxtM_Referencia, TxtM_Importe);
        }

        public void CargarGV(DataGridView GridV, string LibroExcel, string HojaExcel)
        {
            OleDbConnection ConexionMNGR = null;
            DataSet dataSetMNGR = null;
            OleDbDataAdapter dataAdapterMNGR = null;
            string consultaExcelMNGR = "Select * from [" + HojaExcel + "$]";
            string ConexionAExcelMNGR = "provider=Microsoft.ACE.OLEDB.12.0;Data Source='" + LibroExcel + "';Extended Properties=Excel 12.0;";
            if (string.IsNullOrEmpty(HojaExcel))
            {
                MessageBox.Show("No hay una hoja para leer");
            }
            else
            {
                //try
                //{
                    ConexionMNGR = new OleDbConnection(ConexionAExcelMNGR);
                    ConexionMNGR.Open();
                    dataAdapterMNGR = new OleDbDataAdapter(consultaExcelMNGR, ConexionMNGR);
                    dataSetMNGR = new DataSet();
                    dataAdapterMNGR.Fill(dataSetMNGR, HojaExcel);
                    GridV.DataSource = dataSetMNGR.Tables[0];
                    ConexionMNGR.Close();
                    GridV.AllowUserToAddRows = false;
                    if (GridV.Name == "GVRecargasMNGR")
                    {
                        GridV.Sort(GridV.Columns["F5"], ListSortDirection.Ascending);
                        GridV.Sort(GridV.Columns["F6"], ListSortDirection.Ascending);
                        // GridV.Sort(GridV.Columns["F7"], ListSortDirection.Ascending);
                    }
                    if (GridV.Name == "DGVComisionesMNGR")
                    {
                        GridV.Sort(GridV.Columns["F7"], ListSortDirection.Ascending);
                    }
                //}
                //catch (Exception ex)
                //{
                //    MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja", ex.Message);
                //}
            }
        }        
    }
}