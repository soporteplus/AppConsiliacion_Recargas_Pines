using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class RecargaDOS
    {
        Configuracion LLEnarGV = new Configuracion();

        string tienda;

        int NoTickets = 0;
        int NoRecargas = 0;
        bool tICKETS;

        int RReportadas = 0;

        bool REPORTE;
        int NoFTotla = 0;

        string[,] Tickets;

        string[,] RecargasMNGR;
        string[,] RecargasMTC;
        string[,] Reportadas;

        string[,] CatalogoCajas;
        string[,] CatalogoRecargasMNGR;
        string[,] CatalogoRecargasMTCe;

      
        public void CargarTickets(string ExcelName, DataGridView DGVTickets)
        {
            try
            {
                LLEnarGV.CargarGV(DGVTickets, ExcelName, "Tickets");
                Tickets = new string[DGVTickets.Rows.Count, 3];//0 -> Serie, 1 -> NoTicket, 2 -> Estado(True o False) se se encontro o no
                for (int fila = 0; fila < DGVTickets.Rows.Count; fila++)
                {
                    if (DGVTickets.Rows[fila].Cells[0].Value.ToString() != "" && DGVTickets.Rows[fila].Cells[1].Value.ToString() != "")
                    {
                        Tickets[fila, 0] = DGVTickets.Rows[fila].Cells[0].Value.ToString();
                        Tickets[fila, 1] = DGVTickets.Rows[fila].Cells[1].Value.ToString();
                        Tickets[fila, 2] = "";
                        NoTickets = NoTickets + 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No hay Tickets que procesar.", ex.Message);
            }
        }
        public void ReporteAnt(string ExcelName, DataGridView DGVReporte)
        {
            try
            {
                LLEnarGV.CargarGV(DGVReporte, ExcelName, "Reporte");
                int FilaReporte = 0;
                for (int fila = 2; fila < DGVReporte.Rows.Count; fila++)
                {
                    if (DGVReporte.Rows[fila].Cells[1].Value.ToString() != "")
                    {
                        NoFTotla++;
                    }
                }
                Reportadas = new string[NoFTotla, 5];//0 -> Serie, 1 -> NoTicket, 2 -> Estado(True o False) se se encontro o no
                int Nplus = 0;
                for (int fila = 2; fila < DGVReporte.Rows.Count; fila++)
                {
                    if (DGVReporte.Rows[fila].Cells[1].Value.ToString() != "")
                    {
                        if (DGVReporte.Rows[fila].Cells[0].Value.ToString() != "")
                        {
                            string Caja = DGVReporte.Rows[fila].Cells[0].Value.ToString().Replace("PLUS ", "");
                            Nplus = Convert.ToInt16(Caja);
                        }
                        Reportadas[FilaReporte, 0] = Nplus.ToString();//DGVReporte.Rows[fila].Cells[0].Value.ToString();//Plus
                        Reportadas[FilaReporte, 1] = DGVReporte.Rows[fila].Cells[1].Value.ToString();//Fecha y Hora
                        Reportadas[FilaReporte, 2] = DGVReporte.Rows[fila].Cells[2].Value.ToString();//Vendedor
                        Reportadas[FilaReporte, 3] = DGVReporte.Rows[fila].Cells[3].Value.ToString();//Recarga
                        Reportadas[FilaReporte, 4] = "";//Conciliacion
                        FilaReporte++;
                        RReportadas = RReportadas + 1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("No hay Recargas que procesar.", ex.Message);
            }
        }
        public void ConciliarRecargas(ProgressBar PBar,  DataGridView DGVManager, DataGridView DGVTicketsFound, DataGridView DGVTicketsNotFound, DataGridView DGVEmpates, DataGridView DGVMatchRecargas, DataGridView DGVNoMachtRecargas, DataGridView DGVTempRECARGAS,  DataGridView DGVCajas, DataGridView DGVRecargas, DataGridView DGVTickets, DataGridView DGVReporte, DataGridView DGVRecargasMTC)
        {
            
            OpenFileDialog Carpeta = new OpenFileDialog();
            Carpeta.ValidateNames = false;
            Carpeta.CheckFileExists = false;
            Carpeta.CheckPathExists = true;
            Carpeta.FileName = "Seleccione...";
            if (Carpeta.ShowDialog() == DialogResult.OK)
            {
                string DireccionCarpeta = Path.GetDirectoryName(Carpeta.FileName);
                //ReM.Renombrar(DireccionCarpeta + "\\MTC\\");
                tICKETS = File.Exists(DireccionCarpeta + "\\Tickets.xls");
                REPORTE = File.Exists(DireccionCarpeta + "\\RESUMEN_RECARGAS.xls");
                if (tICKETS == true)
                {
                    if (REPORTE == true)
                    {
                        ReporteAnt(DireccionCarpeta + "\\RESUMEN_RECARGAS.xls", DGVReporte);
                        CargarTickets(DireccionCarpeta + "\\Tickets.xls", DGVTickets);
                    }
                    else
                    {
                        MessageBox.Show("No Esta el reporte de Recargas de la semana Anterior");
                        CargarTickets(DireccionCarpeta + "\\Tickets.xls", DGVTickets);
                    }
                }
                else
                {
                    //Log.Text = "No hay Ticktes por Procesar";
                }
                DatosMNGRRecargas(DireccionCarpeta + "\\00.xls", DGVManager, DGVTicketsFound, DGVTicketsNotFound, DGVEmpates, DGVRecargas);
                Recargas_MTCMNGR(PBar, DireccionCarpeta + "\\MTC\\",  DGVMatchRecargas,  DGVNoMachtRecargas,  DGVTempRECARGAS,  DGVCajas,  DGVRecargas,  DGVRecargasMTC);
            }
            PBar.Minimum = 0;
        }
        public void DatosMNGRRecargas(string ExcelName,  DataGridView DGVManager, DataGridView DGVTicketsFound, DataGridView DGVTicketsNotFound, DataGridView DGVEmpates, DataGridView DGVRecargas)
        {
            try
            {
                LLEnarGV.CargarGV(DGVManager, ExcelName, "00");
                int NReg = 0;
                if (tICKETS == true)
                {
                    for (int fila = 0; fila < DGVManager.Rows.Count - 1; fila++)
                    {
                        for (int FilaT = 0; FilaT < NoTickets; FilaT++)
                        {
                            string Serie = DGVManager.Rows[fila].Cells[0].Value.ToString();
                            string NoTicket = DGVManager.Rows[fila].Cells[3].Value.ToString();
                            if (Serie == Tickets[FilaT, 0].ToString() && NoTicket == Tickets[FilaT, 1].ToString())
                            {
                                int unidades = Convert.ToInt16(DGVManager.Rows[fila].Cells[6].Value.ToString());
                                if (unidades > 1)
                                {
                                    for (int t = 0; t < unidades; t++)
                                    {
                                        DGVTicketsFound.Rows.Add(DGVManager.Rows[fila].Cells[0].Value.ToString(), DGVManager.Rows[fila].Cells[1].Value.ToString().Replace(" 12:00:00 a. m.", "") + " " + DGVManager.Rows[fila].Cells[2].Value.ToString().Replace("30/12/1899 ", ""), DGVManager.Rows[fila].Cells[3].Value.ToString(), DGVManager.Rows[fila].Cells[4].Value.ToString(), DGVManager.Rows[fila].Cells[5].Value.ToString());
                                    }
                                }
                                else
                                {
                                    DGVTicketsFound.Rows.Add(DGVManager.Rows[fila].Cells[0].Value.ToString(), DGVManager.Rows[fila].Cells[1].Value.ToString().Replace(" 12:00:00 a. m.", "") + " " + DGVManager.Rows[fila].Cells[2].Value.ToString().Replace("30/12/1899 ", ""), DGVManager.Rows[fila].Cells[3].Value.ToString(), DGVManager.Rows[fila].Cells[4].Value.ToString(), DGVManager.Rows[fila].Cells[5].Value.ToString());
                                }
                                DGVManager.Rows[fila].Cells[0].Value = "";
                                DGVManager.Rows[fila].Cells[1].Value = "";
                                DGVManager.Rows[fila].Cells[2].Value = "";
                                DGVManager.Rows[fila].Cells[3].Value = 0;
                                DGVManager.Rows[fila].Cells[4].Value = "";
                                DGVManager.Rows[fila].Cells[5].Value = "";
                                DGVManager.Rows[fila].Cells[6].Value = 0;

                                Tickets[FilaT, 2] = "true";
                                break;
                            }
                        }
                    }
                    //Mostrasr los tickets no encontradoas
                    for (int FilaT = 0; FilaT < NoTickets; FilaT++)
                    {
                        if (Tickets[FilaT, 2].ToString() != "true")
                        {
                            DGVTicketsNotFound.Rows.Add(Tickets[FilaT, 0].ToString(), Tickets[FilaT, 1].ToString());
                        }
                    }
                }
                for (int FUnidades = 1; FUnidades < DGVManager.Rows.Count - 1; FUnidades++)
                {
                    //El número 10 es por la columna donde se localizan las unidades, es como en los arreglos la numeración empieza en cero

                    int val = Convert.ToInt32(DGVManager.Rows[FUnidades].Cells[6].Value.ToString());
                    if (val < 0)
                    {
                        val = val * -1;
                    }
                    NoRecargas = NoRecargas + val;
                }
                RecargasMNGR = new string[NoRecargas, 5];
                NReg = DGVManager.Rows.Count - 2;
                long NDato = 0;
                for (int fila = 1; fila < DGVManager.Rows.Count - 1; fila++)
                {
                    int unidades = Convert.ToInt32(DGVManager.Rows[fila].Cells[6].Value.ToString());
                    if (unidades != 0 && DGVManager.Rows[fila].Cells[6].Value.ToString() != "")
                    {
                        if (unidades < 0)
                        {
                            unidades = unidades * -1;
                        }
                        if (unidades > 1)
                        {
                            for (int unit = 0; unit < unidades; unit++)
                            {
                                //Limpiar Caja
                                
                                RecargasMNGR[NDato, 0] = DGVManager.Rows[fila].Cells[0].Value.ToString().Replace("1T", "").Replace("2T", "");//Es la caja
                                RecargasMNGR[NDato, 1] = DGVManager.Rows[fila].Cells[1].Value.ToString().Replace(" 12:00:00 a. m.", "") + " " + DGVManager.Rows[fila].Cells[2].Value.ToString().Replace("30/12/1899 ", "");//Fecha Hora
                                RecargasMNGR[NDato, 2] = DGVManager.Rows[fila].Cells[4].Value.ToString();//Usuario
                                RecargasMNGR[NDato, 3] = Concepto(DGVManager.Rows[fila].Cells[5].Value.ToString(), DGVRecargas);//Recarga
                                RecargasMNGR[NDato, 4] = "";
                                NDato++;
                            }
                        }
                        else
                        {
                            RecargasMNGR[NDato, 0] = DGVManager.Rows[fila].Cells[0].Value.ToString().Replace("1T", "").Replace("2T", "");//Es la caja
                            RecargasMNGR[NDato, 1] = DGVManager.Rows[fila].Cells[1].Value.ToString().Replace(" 12:00:00 a. m.", "") + " " + DGVManager.Rows[fila].Cells[2].Value.ToString().Replace("30/12/1899 ", "");//Fecha Hora
                            RecargasMNGR[NDato, 2] = DGVManager.Rows[fila].Cells[4].Value.ToString();//Usuario
                            RecargasMNGR[NDato, 3] = Concepto(DGVManager.Rows[fila].Cells[5].Value.ToString(), DGVRecargas);//Recarga                                                                                                           
                            RecargasMNGR[NDato, 4] = "";
                            NDato++;
                        }
                    }

                }
                DGVTicketsFound.Sort(DGVTicketsFound.Columns[0], ListSortDirection.Ascending);
                //LimiarMNGR(DGVManager, DDGVManager, Log);                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja datos MANAGER Recargas", ex.Message);
            }
        }
        public string Concepto(string Cadena, DataGridView DGVRecargas)
        {
            string ConceptoComun = "";
            
            for (int CCSRC = 1; CCSRC < DGVRecargas.Rows.Count; CCSRC++)
            {
                string RecargaCatalogo = DGVRecargas.Rows[CCSRC].Cells[2].Value.ToString();

                if (Cadena == RecargaCatalogo)
                {
                    ConceptoComun = DGVRecargas.Rows[CCSRC].Cells[1].Value.ToString();
                    break;
                }
            }
            return ConceptoComun;
        }
        public void Recargas_MTCMNGR(ProgressBar PBar, string ruta, DataGridView DGVMatchRecargas, DataGridView DGVNoMachtRecargas,  DataGridView DGVTempRECARGAS, DataGridView DGVCajas, DataGridView DGVRecargas, DataGridView DGVRecargasMTC)
        {
            DirectoryInfo di = new DirectoryInfo(@ruta);
            int cont = 1;
            //CALCULAR VALOR MAXIMO DEL PROGRES
            int ValMaxProgress = 0;
            foreach (var fi in di.GetFiles())
            {
                ValMaxProgress++;
            }
            PBar.Maximum = ValMaxProgress;
            foreach (var fi in di.GetFiles())
            {
                LLEnarGV.CargarGV(DGVTempRECARGAS, ruta + "reporteVentas(" + cont + ").xls", "reporteVentas(" + cont + ")");
                Recargas_MTC(DGVRecargasMTC, DGVTempRECARGAS, DGVCajas, DGVRecargas);
                Diferncia_Tiempo(DGVMatchRecargas, -05, 05);
                //Diferncia_Tiempo(DGVMatchRecargas, -10, 10);
                //Diferncia_Tiempo(DGVMatchRecargas, -20, 20);
                //Diferncia_Tiempo(DGVMatchRecargas, -30, 30);
                //MatchRecargas(DGVMatchRecargas);
                //DGVMatchRecargas.Rows.Add("", "", "", "", "");
                //MatchRecargasFechasVendedor(DGVNoMachtRecargas);
                //MatchRecargasFechas(DGVNoMachtRecargas);
                //NoMatchRecargas(DGVNoMachtRecargas);
                ////MatchRecargas(DGVMatchRecargas);
                //NoMatchR(DGVNoMachtRecargas);
                //DGVNoMachtRecargas.Rows.Add("", "", "", "", "");
                PBar.Value = cont;
                cont++;
            }
            MessageBox.Show("Proceso Finalizado.");
        }
        public void Recargas_MTC(DataGridView DGVRecargasMTC, DataGridView GVRecargasMTC, DataGridView DGVCajas, DataGridView DGVRecargas)
        {
            int NRecargasMTC = GVRecargasMTC.Rows.Count;
            //Filtra los pagos de servicios
            //0-No plus,1-Fecha y Hora, 2-vendedor, 3-concepto, 4-estado
            RecargasMTC = new string[NRecargasMTC, 5];
            int CDLim = 0;
            for (int i = 1; i < NRecargasMTC; i++)
            {
                //Datos de las recargas de MTCenter
                string tienda = GVRecargasMTC.Rows[i].Cells[3].Value.ToString();
                //DateTime fechaHora = DateTime.Parse(GVRecargasMTC.Rows[i].Cells[1].Value.ToString());
                string cajero = GVRecargasMTC.Rows[i].Cells[5].Value.ToString();
                string producto = GVRecargasMTC.Rows[i].Cells[7].Value.ToString();
                //cambia el concepto tienda a caja
                for (int filaCajas = 1; filaCajas < DGVCajas.Rows.Count; filaCajas++)
                {
                    string CajaCatalogo = DGVCajas.Rows[filaCajas].Cells[0].Value.ToString();

                    if (tienda == CajaCatalogo)
                    {
                        RecargasMTC[CDLim, 0] = DGVCajas.Rows[filaCajas].Cells[1].Value.ToString();
                        break;
                    }
                }
                //Fecha y Hora                   
                RecargasMTC[CDLim, 1] = GVRecargasMTC.Rows[i].Cells[1].Value.ToString();
                //Vendedores
                RecargasMTC[CDLim, 2] = GVRecargasMTC.Rows[i].Cells[5].Value.ToString();
                //cambiar concepto de MTCenter al concepto de Front y restar la comisión del servicios
                for (int CCSRC = 0; CCSRC < DGVRecargas.Rows.Count - 1/*NoSer*/; CCSRC++)
                {
                    string RecargaCatalogo = DGVRecargas.Rows[CCSRC].Cells[0].Value.ToString();
                    //procesa el concepto
                    if (producto == RecargaCatalogo)
                    {
                        RecargasMTC[CDLim, 3] = DGVRecargas.Rows[CCSRC].Cells[1].Value.ToString();
                    }
                }
                if (RecargasMTC[CDLim, 0] == null)
                {
                    RecargasMTC[CDLim, 0] = "";
                }
                if (RecargasMTC[CDLim, 1] == null)
                {
                    RecargasMTC[CDLim, 1] = "";
                }
                if (RecargasMTC[CDLim, 2] == null)
                {
                    RecargasMTC[CDLim, 2] = "";
                }
                if (RecargasMTC[CDLim, 3] == null)
                {
                    RecargasMTC[CDLim, 3] = "";
                }
                if (RecargasMTC[CDLim, 4] == null)
                {
                    RecargasMTC[CDLim, 4] = "";
                }
                //GVRecargasMTC.Rows.Add(RecargasMTC[CDLim, 0].ToString(), RecargasMTC[CDLim, 1].ToString(), RecargasMTC[CDLim, 2].ToString(), RecargasMTC[CDLim, 3].ToString());
                CDLim++;

            }
        }

        public void Diferncia_Tiempo(DataGridView DGVMatchRecargas, int Quitar, int Agregar)
        {
            for (int y = 0; y < RecargasMTC.GetLength(0)-1; y++)
            {
                for (int x = 1; x < NoRecargas; x++)
                {
                    string cajaMTC = RecargasMTC[y, 0].ToString();
                    string cajaMngr = RecargasMNGR[x, 0].ToString();
                    if (cajaMTC == cajaMngr)
                    {
                        DateTime FMNGR = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        DateTime FMTC = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        string ProductoMTC = RecargasMTC[y, 3].ToString();
                        string ProductoMNGR = RecargasMNGR[x, 3].ToString();
                        if (FMTC > FMNGR.AddMinutes(Quitar) && FMTC < FMNGR.AddMinutes(Agregar) && RecargasMTC[y, 2].ToString() == RecargasMNGR[x, 2].ToString() && ProductoMTC == ProductoMNGR && RecargasMTC[y, 4] != "MATCH" && RecargasMNGR[x, 4] != "MATCH")
                        {
                            DGVMatchRecargas.Rows.Add("Plataforma", RecargasMTC[y, 0].ToString(), FMTC.ToString(), RecargasMTC[y, 2].ToString(), RecargasMTC[y, 3].ToString());
                            RecargasMTC[y, 4] = "MATCH";
                            DGVMatchRecargas.Rows.Add("Manager", RecargasMNGR[x, 0].ToString(), FMNGR.ToString(), RecargasMNGR[x, 2].ToString(), RecargasMNGR[x, 3].ToString());
                            RecargasMNGR[x, 4] = "MATCH";
                            break;
                        }
                    }
                }
            }
        }
    }
}
