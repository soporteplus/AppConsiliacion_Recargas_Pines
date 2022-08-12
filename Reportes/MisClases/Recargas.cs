using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class Recargas
    {
        RenArchivos ReM = new RenArchivos();
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

        public void CargarCatalogos(DataGridView DGVCajas, DataGridView DGVRecargas, DataGridView DGVSupervisores)
        {
            //GUARDA LOS DATOS DE LAS CAJAS EN UN ARREGLO
            
            CatalogoCajas = new string[DGVCajas.Rows.Count - 1, 2];
            for (int filaCajas = 1; filaCajas < DGVCajas.Rows.Count; filaCajas++)
            {
                for (int col = 0; col < DGVCajas.Rows[filaCajas].Cells.Count; col++)
                {
                    CatalogoCajas[filaCajas - 1, col] = DGVCajas.Rows[filaCajas].Cells[col].Value.ToString();
                }
            }
            //GUARDA LOS DATOS DE LAS RECARGAS DE MTCENTER
            int conta = 0;
            CatalogoRecargasMTCe = new string[DGVRecargas.Rows.Count - 1, 2];
            for (int filaRecargasMTC = 1; filaRecargasMTC < DGVRecargas.Rows.Count; filaRecargasMTC++)
            {
                CatalogoRecargasMTCe[filaRecargasMTC - 1, 0] = DGVRecargas.Rows[filaRecargasMTC].Cells[0].Value.ToString();
                CatalogoRecargasMTCe[filaRecargasMTC - 1, 1] = DGVRecargas.Rows[filaRecargasMTC].Cells[1].Value.ToString();
                if (DGVRecargas.Rows[filaRecargasMTC].Cells[2].Value.ToString() != "")
                {
                    conta++;
                }
            }
            //GUARDA LOS DATOS DE LAS RECARGAS DE MANAGER
            CatalogoRecargasMNGR = new string[conta, 2];
            int proges = 0;
            for (int filaRecargaMNGR = 1; filaRecargaMNGR < DGVRecargas.Rows.Count; filaRecargaMNGR++)
            {
                if (DGVRecargas.Rows[filaRecargaMNGR].Cells[2].Value.ToString() != "")
                {
                    CatalogoRecargasMNGR[proges, 0] = DGVRecargas.Rows[filaRecargaMNGR].Cells[1].Value.ToString();
                    CatalogoRecargasMNGR[proges, 1] = DGVRecargas.Rows[filaRecargaMNGR].Cells[2].Value.ToString();
                    proges++;
                }
            }            
        }
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
        public void ConciliarRecargas(ProgressBar PBar, string ColumnasManejador, DataGridView DGVMTCenter, DataGridView DGVManager, DataGridView DGVMatchRecargas, DataGridView DGVNoMachtRecargas, DataGridView DGVTicketsFound, DataGridView DGVTickets, DataGridView DGVTicketsNotFound, DataGridView DGVReporte, DataGridView DGVEmpates, DataGridView DGVSupervisores, DataGridView DGVCajas, DataGridView DGVRecargas)
        {
            //Log.Clear();
            CargarCatalogos(DGVCajas, DGVRecargas, DGVSupervisores);
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
                DatosMNGRRecargas(DireccionCarpeta + "\\00.xls", ColumnasManejador, DGVManager, DGVTicketsFound, DGVTicketsNotFound, DGVEmpates);
                Recargas_MTCMNGR(PBar, DireccionCarpeta + "\\MTC\\", DGVMatchRecargas, DGVNoMachtRecargas, DGVSupervisores, DGVRecargas, DGVMTCenter);
            }
            PBar.Minimum = 0;
        }
        public void DatosMNGRRecargas(string ExcelName, string ColumnasManejador, DataGridView DGVManager, DataGridView DGVTicketsFound, DataGridView DGVTicketsNotFound, DataGridView DGVEmpates)
        {
            try
            {
                LLEnarGV.CargarGV(DGVManager, ExcelName, "00");
                string[] ColumnasM = ColumnasManejador.Split('/');
                //identificar la posicion de la columna
                int Caja = 0, Fecha = 0, Hora = 0, Vendedor = 0, Clasificacion = 0, Comentario = 0;
                for (int col = 0; col < DGVManager.Rows[0].Cells.Count; col++)
                {
                    string dato = DGVManager.Rows[0].Cells[col].Value.ToString();
                    if (dato == ColumnasM[0].ToString())
                    {
                        Caja = col;
                    }
                    if (dato == ColumnasM[1].ToString())
                    {
                        Fecha = col;
                    }
                    if (dato == ColumnasM[2].ToString())
                    {
                        Hora = col;
                    }
                    if (dato == ColumnasM[3].ToString())
                    {
                        Vendedor = col;
                    }
                    if (dato == ColumnasM[4].ToString())
                    {
                        Clasificacion = col;
                    }
                    if (dato == ColumnasM[5].ToString())
                    {
                        Comentario = col;
                    }

                }

                int NReg = 0;
                if (tICKETS == true)
                {
                    for (int fila = 0; fila < DGVManager.Rows.Count - 1; fila++)
                    {
                        for (int FilaT = 0; FilaT < NoTickets; FilaT++)
                        {
                            if (DGVManager.Rows[fila].Cells[6].Value.ToString() == Tickets[FilaT, 0].ToString() && DGVManager.Rows[fila].Cells[7].Value.ToString() == Tickets[FilaT, 1].ToString())
                            {
                                int unidades = Convert.ToInt16(DGVManager.Rows[fila].Cells[10].Value.ToString());
                                if (unidades > 1)
                                {
                                    for (int t = 0; t < unidades; t++)
                                    {
                                        DGVTicketsFound.Rows.Add(DGVManager.Rows[fila].Cells[6].Value.ToString(), DGVManager.Rows[fila].Cells[7].Value.ToString(), DGVManager.Rows[fila].Cells[3].Value.ToString(), DGVManager.Rows[fila].Cells[4].Value.ToString() + " " + DGVManager.Rows[fila].Cells[5].Value.ToString(), DGVManager.Rows[fila].Cells[0].Value.ToString());
                                    }
                                }
                                else
                                {
                                    DGVTicketsFound.Rows.Add(DGVManager.Rows[fila].Cells[6].Value.ToString(), DGVManager.Rows[fila].Cells[7].Value.ToString(), DGVManager.Rows[fila].Cells[3].Value.ToString(), DGVManager.Rows[fila].Cells[4].Value.ToString() + " " + DGVManager.Rows[fila].Cells[5].Value.ToString(), DGVManager.Rows[fila].Cells[0].Value.ToString());
                                }
                                DGVManager.Rows[fila].Cells[0].Value = "";
                                DGVManager.Rows[fila].Cells[1].Value = "";
                                DGVManager.Rows[fila].Cells[2].Value = 0;
                                DGVManager.Rows[fila].Cells[3].Value = "";
                                DGVManager.Rows[fila].Cells[4].Value = "";
                                DGVManager.Rows[fila].Cells[5].Value = "";
                                DGVManager.Rows[fila].Cells[6].Value = "";
                                DGVManager.Rows[fila].Cells[7].Value = 0;
                                DGVManager.Rows[fila].Cells[8].Value = 0;
                                DGVManager.Rows[fila].Cells[9].Value = "";
                                DGVManager.Rows[fila].Cells[10].Value = 0;
                                DGVManager.Rows[fila].Cells[11].Value = 0;
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
                    //El número 10 es por la columna donde se localizan las unidades                
                    int val = Convert.ToInt32(DGVManager.Rows[FUnidades].Cells[7].Value.ToString());
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
                    int unidades = Convert.ToInt32(DGVManager.Rows[fila].Cells[10].Value.ToString());
                    if (unidades != 0 && DGVManager.Rows[fila].Cells[10].Value.ToString() != "")
                    {
                        if (unidades < 0)
                        {
                            unidades = unidades * -1;
                        }
                        if (unidades > 1)
                        {
                            for (int unit = 0; unit < unidades; unit++)
                            {
                                RecargasMNGR[NDato, 0] = DGVManager.Rows[fila].Cells[6].Value.ToString().Replace("1T", "").Replace("2T", "");//Es la caja
                                RecargasMNGR[NDato, 1] = DGVManager.Rows[fila].Cells[4].Value.ToString() + " " + DGVManager.Rows[fila].Cells[5].Value.ToString();//Fecha Hora
                                RecargasMNGR[NDato, 2] = DGVManager.Rows[fila].Cells[3].Value.ToString();//Usuario
                                RecargasMNGR[NDato, 3] = Concepto(DGVManager.Rows[fila].Cells[0].Value.ToString());//Recarga
                                                                                                                   //RecargasMNGR[NDato, 3] = DGVManager.Rows[fila].Cells[0].Value.ToString();//Recarga
                                RecargasMNGR[NDato, 4] = "";
                                NDato++;
                            }
                        }
                        else
                        {
                            RecargasMNGR[NDato, 0] = DGVManager.Rows[fila].Cells[6].Value.ToString().Replace("1T", "").Replace("2T", "");//Es la caja
                            RecargasMNGR[NDato, 1] = DGVManager.Rows[fila].Cells[4].Value.ToString() + " " + DGVManager.Rows[fila].Cells[5].Value.ToString();//Fecha Hora
                            RecargasMNGR[NDato, 2] = DGVManager.Rows[fila].Cells[3].Value.ToString();//Usuario
                            RecargasMNGR[NDato, 3] = Concepto(DGVManager.Rows[fila].Cells[0].Value.ToString());//Recarga                                                                                                           
                            RecargasMNGR[NDato, 4] = "";
                            NDato++;
                        }
                    }

                }
                DGVTicketsFound.Sort(DGVTicketsFound.Columns[0], ListSortDirection.Ascending);
                //LimiarMNGR(DGVManager, DDGVManager, Log);
                MessageBox.Show("Numero de Registros: " + NReg.ToString() + "\r\n Numero de Recargas: " + NoRecargas.ToString());

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error, Verificar el archivo o el nombre de la hoja datos MANAGER Recargas", ex.Message);
            }
            
        }
        public string Concepto(string Cadena)
        {
            string ConceptoComun = "";            
            for (int CCSRC = 0; CCSRC < CatalogoRecargasMNGR.GetLength(0); CCSRC++)
            {
                if (Cadena == CatalogoRecargasMNGR[CCSRC, 1].ToString())
                {
                    ConceptoComun = CatalogoRecargasMNGR[CCSRC, 0].ToString();
                    break;
                }
            }
            return ConceptoComun;
        }
        public void Recargas_MTCMNGR(ProgressBar PBar, string ruta, DataGridView DGVMatchRecargas, DataGridView DGVNoMachtRecargas, DataGridView DGVSupervisores, DataGridView DGVRecargasMTC, DataGridView GVRecargasMTC)
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
                //Recargas_MTC(ruta, cont, DGVSupervisores, DGVRecargasMTC, GVRecargasMTC);

                LLEnarGV.CargarGV(GVRecargasMTC, ruta + "reporteVentas(" + cont + ").xls", "reporteVentas(" + cont + ")");
                Diferncia_Tiempo(DGVMatchRecargas, -05, 05);
                Diferncia_Tiempo(DGVMatchRecargas, -10, 10);
                Diferncia_Tiempo(DGVMatchRecargas, -20, 20);
                Diferncia_Tiempo(DGVMatchRecargas, -30, 30);
                MatchRecargas(DGVMatchRecargas);
                DGVMatchRecargas.Rows.Add("", "", "", "", "");
                MatchRecargasFechasVendedor(DGVNoMachtRecargas);
                MatchRecargasFechas(DGVNoMachtRecargas);
                NoMatchRecargas(DGVNoMachtRecargas);
                //MatchRecargas(DGVMatchRecargas);
                NoMatchR(DGVNoMachtRecargas);
                DGVNoMachtRecargas.Rows.Add("", "", "", "", "");
                PBar.Value = cont;
                cont++;
            }            
            MessageBox.Show("Proceso Finalizado.");
        }
        public void Recargas_MTC(string ruta, int Plus, DataGridView DGVSupervisores, DataGridView DGVRecargasMTC, DataGridView GVRecargasMTC)
        {
            string[] DatosR;
            int NLineas = 0;
            int NRecargasMTC = 0;
            int CArreglo = 0;
            List<string[]> parsedData = new List<string[]>();

            //Cuenta las lineas del archivo
            using (StreamReader readFile = new StreamReader(ruta + "reporteVentas (" + Plus + ").csv"))
            {
                string LCont;
                while ((LCont = readFile.ReadLine()) != null)
                {
                    NLineas++;
                }
            }
            DatosR = new string[NLineas];

            //limpia la cadena y la guarda en un arreglo
            using (StreamReader readFile = new StreamReader(ruta + "reporteVentas (" + Plus + ").csv"))
            {
                string line;
                string[] row;
                while ((line = readFile.ReadLine()) != null)
                {
                    row = line.Split(',');
                    parsedData.Add(row);
                    DatosR[CArreglo] = line.ToString().Replace("\t\t<td>", "").Replace("\t", "").Replace("<td>", "°").Replace("</td>", "°").Replace("</tr>", "").Replace("<tr>", "").Replace("°°", "°").Replace("&nbsp;", " ").Replace("</table>", "").TrimEnd('°');
                    if (DatosR[CArreglo] != "" && DatosR[CArreglo].Contains("TAE") == true && DatosR[CArreglo] != null)
                    {
                        NRecargasMTC++;
                    }
                    CArreglo++;
                }
            }
            //Filtra los pagos de servicios
            //0-No plus,1-Fecha y Hora, 2-vendedor, 3-concepto, 4-estado

            RecargasMTC = new string[NRecargasMTC, 5];
            int CDLim = 0;
            for (int i = 1; i < NLineas; i++)
            {
                if (DatosR[i] != "" && DatosR[i].Contains("TAE") == true && DatosR[i] != null)
                {
                    string Recarga = DatosR[i];
                    string[] Valores = Recarga.Split('°');
                    //cambia el concepto tienda a caja
                    for (int z = 0; z < CatalogoCajas.GetLength(0); z++)
                    {
                        string CajaMTC = Valores[3].ToString().Replace(" S", "S");
                        string CajaCatalogo = CatalogoCajas[z, 0].ToString().Replace(" S", "S");
                        if (CajaMTC == CajaCatalogo)
                        {
                            RecargasMTC[CDLim, 0] = CatalogoCajas[z, 1].ToString();
                            break;
                        }
                    }
                    tienda = RecargasMTC[CDLim, 0].ToString();
                    //Fecha y Hora                   
                    RecargasMTC[CDLim, 1] = Valores[1].ToString();
                    //cambiar concepto de MTCenter al concepto de Front y restar la comisión del servicios
                    
                    //cambiar concepto de MTCenter al concepto de Front y restar la comisión del servicios
                    for (int CCSRC = 0; CCSRC < DGVRecargasMTC.Rows.Count - 1/*NoSer*/; CCSRC++)
                    {
                        //procesa el concepto
                        if (Valores[7].ToString() == CatalogoRecargasMTCe[CCSRC, 0].ToString())
                        {
                            RecargasMTC[CDLim, 3] = CatalogoRecargasMTCe[CCSRC, 1].ToString();
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
                    GVRecargasMTC.Rows.Add(RecargasMTC[CDLim, 0].ToString(), RecargasMTC[CDLim, 1].ToString(), RecargasMTC[CDLim, 2].ToString(), RecargasMTC[CDLim, 3].ToString());
                    CDLim++;
                }
            }
        }
        public void Diferncia_Tiempo(DataGridView DGVMatchRecargas, int Quitar, int Agregar)
        {
            for (int y = 0; y < RecargasMTC.GetLength(0); y++)
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

        public void MatchRecargas(DataGridView DGVMatchRecargas)
        {
            for (int y = 0; y < RecargasMTC.GetLength(0); y++)
            {
                for (int x = 0; x < NoRecargas; x++)
                {
                    string cajaMTC = RecargasMTC[y, 0].ToString();
                    string cajaMngr = RecargasMNGR[x, 0].ToString();
                    if (cajaMTC == cajaMngr)
                    {
                        DateTime FMNGR = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        DateTime FMTC = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        string ProductoMTC = RecargasMTC[y, 3].ToString();
                        string ProductoMNGR = RecargasMNGR[x, 3].ToString();
                        if (FMTC > FMNGR.AddMinutes(-5) && FMTC < FMNGR.AddMinutes(5) && ProductoMTC == ProductoMNGR && RecargasMTC[y, 4] != "MATCH" && RecargasMNGR[x, 4] != "MATCH")
                        {
                            DGVMatchRecargas.Rows.Add("Plataforma", RecargasMTC[y, 0].ToString(), FMTC.ToString(), RecargasMTC[y, 2].ToString(), RecargasMTC[y, 3].ToString());
                            RecargasMTC[y, 4] = "MATCH";
                            DGVMatchRecargas.Rows.Add("Manager", RecargasMNGR[x, 0].ToString(), FMNGR.ToString(), RecargasMNGR[x, 2].ToString(), RecargasMNGR[x, 3].ToString());
                            RecargasMNGR[x, 4] = "MATCH";
                        }
                    }
                }
            }
            //InvertArryMTC();
            //InvertArryMNGR();
        }

        public void MatchRecargasFechasVendedor(DataGridView DGVNoMachtRecargas)
        {
            for (int y = 0; y < RecargasMTC.GetLength(0); y++)
            {
                for (int x = 0; x < NoRecargas; x++)
                {
                    string cajaMTC = RecargasMTC[y, 0].ToString();
                    string cajaMngr = RecargasMNGR[x, 0].ToString();
                    if (cajaMTC == cajaMngr)
                    {
                        DateTime FMNGR = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM")).Date;
                        DateTime FMTC = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM")).Date;

                        DateTime FMNGR2 = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        DateTime FMTC2 = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));

                        string ProductoMTC = RecargasMTC[y, 3].ToString();
                        string ProductoMNGR = RecargasMNGR[x, 3].ToString();

                        if (FMTC == FMNGR && ProductoMTC == ProductoMNGR && RecargasMTC[y, 2].ToString() == RecargasMNGR[x, 2].ToString() && RecargasMTC[y, 4] != "MATCH" && RecargasMNGR[x, 4] != "MATCH")
                        {
                            //Información MTCenter
                            DGVNoMachtRecargas.Rows.Add("Plataforma", RecargasMTC[y, 0].ToString(), FMTC2.ToString(), RecargasMTC[y, 2].ToString(), RecargasMTC[y, 3].ToString());
                            RecargasMTC[y, 4] = "MATCH";
                            //Información Manager
                            DGVNoMachtRecargas.Rows.Add("Manager", RecargasMNGR[x, 0].ToString(), FMNGR2.ToString(), RecargasMNGR[x, 2].ToString(), RecargasMNGR[x, 3].ToString());
                            RecargasMNGR[x, 4] = "MATCH";
                        }
                    }
                }
            }
            //InvertArryMTC();
            //InvertArryMNGR();
        }

        public void MatchRecargasFechas(DataGridView DGVNoMachtRecargas)
        {
            for (int y = 0; y < RecargasMTC.GetLength(0); y++)
            {
                for (int x = 0; x < NoRecargas; x++)
                {
                    string cajaMTC = RecargasMTC[y, 0].ToString();
                    string cajaMngr = RecargasMNGR[x, 0].ToString();
                    if (cajaMTC == cajaMngr)
                    {
                        DateTime FMNGR = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM")).Date;
                        DateTime FMTC = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM")).Date;

                        DateTime FMNGR2 = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        DateTime FMTC2 = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));

                        string ProductoMTC = RecargasMTC[y, 3].ToString();
                        string ProductoMNGR = RecargasMNGR[x, 3].ToString();

                        if (FMTC == FMNGR && ProductoMTC == ProductoMNGR && RecargasMTC[y, 4] != "MATCH" && RecargasMNGR[x, 4] != "MATCH")
                        {
                            //Información MTCenter
                            DGVNoMachtRecargas.Rows.Add("Plataforma", RecargasMTC[y, 0].ToString(), FMTC2.ToString(), RecargasMTC[y, 2].ToString(), RecargasMTC[y, 3].ToString());
                            RecargasMTC[y, 4] = "MATCH";
                            //Información Manager
                            DGVNoMachtRecargas.Rows.Add("Manager", RecargasMNGR[x, 0].ToString(), FMNGR2.ToString(), RecargasMNGR[x, 2].ToString(), RecargasMNGR[x, 3].ToString());
                            RecargasMNGR[x, 4] = "MATCH";
                        }
                    }
                }
            }
            //InvertArryMTC();
            //InvertArryMNGR();
        }

        public void NoMatchRecargas(DataGridView DGVNoMachtRecargas)
        {
            for (int y = 0; y < RecargasMTC.GetLength(0); y++)
            {
                for (int x = 0; x < NoRecargas; x++)
                {
                    string cajaMTC = RecargasMTC[y, 0].ToString();
                    string cajaMngr = RecargasMNGR[x, 0].ToString();
                    if (cajaMTC == cajaMngr)
                    {
                        DateTime FMNGR = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        DateTime FMTC = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                        if (FMTC > FMNGR.AddMinutes(-5) && FMTC < FMNGR.AddMinutes(5) && RecargasMTC[y, 4] != "MATCH" && RecargasMNGR[x, 4] != "MATCH")
                        {
                            //Información MTCenter      
                            DGVNoMachtRecargas.Rows.Add("Plataforma", RecargasMTC[y, 0].ToString(), FMTC.ToString(), RecargasMTC[y, 2].ToString(), RecargasMTC[y, 3].ToString());
                            RecargasMTC[y, 4] = "MATCH";
                            //Información Manager
                            DGVNoMachtRecargas.Rows.Add("Manager", RecargasMNGR[x, 0].ToString(), FMNGR.ToString(), RecargasMNGR[x, 2].ToString(), RecargasMNGR[x, 3].ToString());
                            RecargasMNGR[x, 4] = "MATCH";
                        }
                    }
                }
            }
            //InvertArryMTC();
            //InvertArryMNGR();
        }

        public void NoMatchR(DataGridView DGVNoMachtRecargas)
        {
            //se imprimen las recargas que no tuvieron macht en la plataforma
            for (int y = 0; y < RecargasMTC.GetLength(0); y++)
            {
                string cajaMTC = RecargasMTC[y, 0].ToString();
                DateTime FMTC = Convert.ToDateTime(RecargasMTC[y, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                if (RecargasMTC[y, 4] != "MATCH")
                {
                    //Información MTCenter
                    DGVNoMachtRecargas.Rows.Add("Plataforma", RecargasMTC[y, 0].ToString(), FMTC.ToString(), RecargasMTC[y, 2].ToString(), RecargasMTC[y, 3].ToString());
                    RecargasMTC[y, 4] = "MATCH";
                }
            }

            for (int x = 0; x < NoRecargas; x++)
            {
                string cajaMngr = RecargasMNGR[x, 0].ToString();
                if (tienda == cajaMngr && RecargasMNGR[x, 4] != "MATCH")
                {
                    DateTime FMNGR = Convert.ToDateTime(RecargasMNGR[x, 1].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM"));
                    //Información Manager
                    DGVNoMachtRecargas.Rows.Add("Manager", RecargasMNGR[x, 0].ToString(), FMNGR.ToString(), RecargasMNGR[x, 2].ToString(), RecargasMNGR[x, 3].ToString());
                    RecargasMNGR[x, 4] = "MATCH";
                }
            }
        }
    }
}
