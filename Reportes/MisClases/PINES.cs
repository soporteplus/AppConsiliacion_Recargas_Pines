using System;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class PINES
    {
        RenArchivos ReM = new RenArchivos();
        Configuracion Config = new Configuracion();

        string[,] DatosMNGR;
        string[,] DatosPlataforma;
        string[,] CatalogoCajas;
        string[,] CatalogoPINES;

        public void CargarCatalogos(DataGridView DGVPines, DataGridView DGVCajas)
        {
            //GUARDA LOS DATOS DE LOS SERVICIOS EN UN ARREGLO                                            
            CatalogoPINES = new string[DGVPines.Rows.Count - 1, 2];
            for (int fila = 1; fila < DGVPines.Rows.Count; fila++)
            {
                for (int col = 0; col < DGVPines.Rows[fila].Cells.Count; col++)
                {
                    CatalogoPINES[fila - 1, col] = DGVPines.Rows[fila].Cells[col].Value.ToString();
                }
            }
            //GUARDA LOS DATOS DE LAS CAJAS EN UN ARREGLO
            CatalogoCajas = new string[DGVCajas.Rows.Count - 1, 2];
            for (int fila = 1; fila < DGVCajas.Rows.Count; fila++)
            {
                for (int col = 0; col < DGVCajas.Rows[fila].Cells.Count; col++)
                {
                    CatalogoCajas[fila - 1, col] = DGVCajas.Rows[fila].Cells[col].Value.ToString();
                }
            }
        }
        public void ConciliarPines(string Conceptos, string ColumnasPlataforma, string MConceptos, string ColumnasManejador, ProgressBar PBar, DataGridView DGVManager, DataGridView DGVPines, DataGridView DGVMTCenter, DataGridView DGVCajas, DataGridView DGVPinesMatch, DataGridView DGVPinesNoMatch)
        {
            CargarCatalogos(DGVPines, DGVCajas);
            OpenFileDialog Carpeta = new OpenFileDialog();
            Carpeta.ValidateNames = false;
            Carpeta.CheckFileExists = false;
            Carpeta.CheckPathExists = true;
            Carpeta.FileName = "Seleccione...";
            if (Carpeta.ShowDialog() == DialogResult.OK)
            {
                string DireccionCarpeta = Path.GetDirectoryName(Carpeta.FileName);
                //llama la clase donde se encuentra el método encargado de renombrar los archivos
                ReM.Renombrar(DireccionCarpeta + "\\MTC\\");
                PINESMNGR(MConceptos, ColumnasManejador, DireccionCarpeta + "\\00.xls", DGVManager);
                Comparar_MTC_MNGR(Conceptos, ColumnasPlataforma, PBar, DireccionCarpeta + "\\MTC\\", DGVPines, DGVMTCenter, DGVPinesMatch, DGVPinesNoMatch);
            }
            PBar.Minimum = 0;
        }
        public void PINESMNGR(string MConceptos, string ColumnasManejador, string ExcelName, DataGridView DGVManager)
        {
            Config.CargarGV(DGVManager, ExcelName, "00");
            string[] ColumnasM = ColumnasManejador.Split('/');
            //identificar la posicion de la columna
            int Caja = 0, Fecha = 0, Hora = 0, Vendedor = 0, Clasificacion = 0, Concepto = 0, Comentario = 0, Importe = 0;
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
                if (dato == ColumnasM[6].ToString())
                {
                    Concepto = col;
                }
                if (dato == ColumnasM[7].ToString())
                {
                    Importe = col;
                }
            }
            //cuenta cuantos Registros se van conciliar *concepto*
            int Nregistros = 0;
            string[] Col = MConceptos.Split('*');
            for (int i = 0; i < Col.GetLength(0); i++)
            {
                for (int fila = 0; fila < DGVManager.Rows.Count; fila++)
                {
                    string dato = DGVManager.Rows[fila].Cells[Concepto].Value.ToString();
                    if (dato.Contains(Col[i].ToString()))
                    {
                        Nregistros++;
                    }
                }
            }
            DatosMNGR = new string[Nregistros, 8];
            int Paso = 0;
            for (int fila = 0; fila < DGVManager.Rows.Count; fila++)
            {
                for (int r = 0; r < Col.GetLength(0); r++)
                {
                    string Con = DGVManager.Rows[fila].Cells[Concepto].Value.ToString();
                    if (Con.Contains(Col[r].ToString()))
                    {
                        //Caja
                        string Caj = DGVManager.Rows[fila].Cells[Caja].Value.ToString();
                        if (Caj.EndsWith("2") == true || Caj.EndsWith("1") == true)
                        {
                            DatosMNGR[Paso, 0] = Caj.Substring(0, 2);
                        }
                        else
                        {
                            DatosMNGR[Paso, 0] = Caj;
                        }
                        DatosMNGR[Paso, 1] = DGVManager.Rows[fila].Cells[Fecha].Value.ToString();
                        DatosMNGR[Paso, 2] = DGVManager.Rows[fila].Cells[Hora].Value.ToString();
                        DatosMNGR[Paso, 3] = DGVManager.Rows[fila].Cells[Vendedor].Value.ToString();
                        DatosMNGR[Paso, 4] = DGVManager.Rows[fila].Cells[Concepto].Value.ToString();
                        //Referencia del pago
                        string RPUTexto = DGVManager.Rows[fila].Cells[Comentario].Value.ToString(), Referencia = string.Empty;
                        for (int i = 0; i < RPUTexto.Length; i++)
                        {
                            if (Char.IsDigit(RPUTexto[i]))
                                Referencia += RPUTexto[i];
                        }
                        DatosMNGR[Paso, 5] = Referencia;
                        DatosMNGR[Paso, 6] = DGVManager.Rows[fila].Cells[Importe].Value.ToString().Replace("-", "");
                        DatosMNGR[Paso, 7] = "";
                        //Fecha + Hora
                        Paso++;
                    }
                }
            }
        }
        public void Comparar_MTC_MNGR(string Conceptos, string ColumnasPlataforma, ProgressBar PBar, string ruta, DataGridView DGVPines, DataGridView DGVMTCenter, DataGridView DGVPinesMatch, DataGridView DGVPinesNoMatch)
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
                CargarDatosMTC(Conceptos, ColumnasPlataforma, ruta, cont, DGVPines, DGVMTCenter);//Carga la información de los archivos uno a la vez                              
                CincoMatch(DGVPinesMatch);
                CuatroMatch(DGVPinesMatch);
                TresNoMatch(DGVPinesNoMatch);
                NoMatch(DGVPinesNoMatch);
                PBar.Value = cont;
                cont++;
            }
            Revisar(DGVPinesNoMatch);
            MessageBox.Show("Se procesaron " + cont + "Plus");
        }
        public void CargarDatosMTC(string Conceptos, string ColumnasPlataforma, string ruta, int Plus, DataGridView DGVPines, DataGridView DGVMTCenter)
        {
            string[] ColP = ColumnasPlataforma.Split('/');
            string[] Concepto = Conceptos.Split('/');
            int tienda = 0, fecha = 0, Hora = 0, cajero = 0, clasificacion = 0, producto = 0, referencia1 = 0, entrada = 0, tamaño = 0;
            int NTArreglo = 0;
            int NTALimpio = 0;
            //Cuenta las lineas del archivo
            using (StreamReader readFile = new StreamReader(ruta + "reporteVentas (" + Plus + ").csv"))
            {
                string LCont;
                while ((LCont = readFile.ReadLine()) != null)
                {
                    NTArreglo++;
                    if (LCont.Contains(ColP[0].ToString()) && LCont.Contains(ColP[1].ToString()) && LCont.Contains(ColP[2].ToString()) && LCont.Contains(ColP[3].ToString()) && LCont.Contains(ColP[4].ToString()) && LCont.Contains(ColP[5].ToString()) && LCont.Contains(ColP[6].ToString()))
                    {
                        //TrimEnd Elimina todas las apariciones finales de un carácter de la cadena actual.
                        string NewLine = LCont.ToString().Replace("\t\t<td>", "").Replace("\t", "").Replace("<td>", "°").Replace("</td>", "°").Replace("</tr>", "").Replace("<tr>", "").Replace("°°", "°").Replace("&nbsp;", " ").Replace("</table>", "").TrimEnd('°');
                        string[] ColummnasNew = NewLine.Split('°');
                        if (tamaño == 0)
                        {
                            tamaño = ColummnasNew.GetLength(0);
                        }
                        for (int i = 0; i < ColummnasNew.GetLength(0); i++)
                        {
                            if (ColP[0].ToString() == ColummnasNew[i].ToString())
                            {
                                tienda = i;
                            }
                            if (ColP[1].ToString() == ColummnasNew[i].ToString())
                            {
                                fecha = i;
                            }
                            if (ColP[2].ToString() == ColummnasNew[i].ToString())
                            {
                                Hora = i;
                            }
                            if (ColP[3].ToString() == ColummnasNew[i].ToString())
                            {
                                cajero = i;
                            }
                            if (ColP[4].ToString() == ColummnasNew[i].ToString())
                            {
                                clasificacion = i;
                            }
                            if (ColP[5].ToString() == ColummnasNew[i].ToString())
                            {
                                producto = i;
                            }
                            if (ColP[6].ToString() == ColummnasNew[i].ToString())
                            {
                                referencia1 = i;
                            }
                            if (ColP[7].ToString() == ColummnasNew[i].ToString())
                            {
                                entrada = i;
                            }
                        }
                    }
                    if (LCont != "" && LCont.Contains(Concepto[2].ToString()) == true)
                    {
                        NTALimpio++;
                    }
                }
            }
            DatosPlataforma = new string[NTALimpio, 8];
            int CDLim = 0;
            //limpia la cadena y la guarda en un arreglo
            using (StreamReader readFile = new StreamReader(ruta + "reporteVentas (" + Plus + ").csv"))
            {
                string line, LineTemp = "";
                while ((line = readFile.ReadLine()) != null)
                {//CAMBIAR EL NUMERO DEL ARREGLO
                    string Unir = LineTemp + line;
                    string NLine = Unir.ToString().Replace("\t\t<td>", "").Replace("\t", "").Replace("<td>", "|").Replace("</td>", "|").Replace("</tr>", "").Replace("<tr>", "").Replace("||", "|").Replace("&nbsp;", " ").Replace("</table>", "").TrimEnd('|').Replace("@", "").Replace(".com", "");
                    string[] ArregloLine = NLine.Split('|');
                    if (ArregloLine.GetLength(0) == 9)
                    {
                        LineTemp = line.ToString();
                    }
                    else
                    {
                        if (Unir.Contains(Concepto[2].ToString()) == true)
                        {
                            //cambia el concepto tienda a caja
                            for (int z = 0; z < CatalogoCajas.GetLength(0); z++)
                            {
                                string CajaMTC = ArregloLine[tienda].ToString().Replace(" S", "S");
                                string CajaCatalogo = CatalogoCajas[z, 0].ToString().Replace(" S", "S");
                                if (CajaMTC == CajaCatalogo)
                                {
                                    DatosPlataforma[CDLim, 0] = CatalogoCajas[z, 1].ToString();
                                }
                            }
                            //Fecha y  Hora
                            if (ArregloLine[fecha].ToString() == ArregloLine[Hora].ToString())
                            {
                                string FechaHora = ArregloLine[fecha].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM");
                                int Caracter = FechaHora.IndexOf(" ");
                                DatosPlataforma[CDLim, 1] = FechaHora.Substring(0, Caracter);
                                int HrI = Caracter + 1;
                                int Dif = FechaHora.Length - HrI;
                                DateTime dt;
                                bool res = DateTime.TryParse(FechaHora.Substring(HrI, Dif), out dt);
                                DatosPlataforma[CDLim, 2] = dt.ToString("HH:mm:ss");
                            }
                            else
                            {
                                DatosPlataforma[CDLim, 1] = ArregloLine[fecha].ToString();
                                DatosPlataforma[CDLim, 2] = ArregloLine[Hora].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM");
                            }
                            //vendedor
                            DatosPlataforma[CDLim, 3] = ArregloLine[cajero].ToString();
                            //Concepto y Entrada
                            for (int CCSRC = 0; CCSRC < DGVPines.Rows.Count - 1/*NoSer*/; CCSRC++)
                            {
                                if (ArregloLine[producto].ToString() == CatalogoPINES[CCSRC, 0].ToString())
                                {
                                    DatosPlataforma[CDLim, 4] = CatalogoPINES[CCSRC, 1].ToString();
                                }
                            }
                            //Referencia del pago
                            string RPUTexto = ArregloLine[referencia1].ToString();
                            string[] Referencia = RPUTexto.Split('/');
                            string REF = Referencia[0].ToString();
                            string NReferencia = string.Empty;
                            for (int i = 0; i < REF.Length; i++)
                            {
                                if (Char.IsDigit(REF[i]))
                                    NReferencia += REF[i];
                            }
                            DatosPlataforma[CDLim, 5] = NReferencia;
                            //Entrada
                            DatosPlataforma[CDLim, 6] = ArregloLine[entrada].ToString();
                            //Estatus
                            DatosPlataforma[CDLim, 7] = "";
                            LineTemp = "";
                            CDLim++;
                        }
                    }
                }
            }
            //Validar que no tengan valores NULL
            for (int i = 0; i < NTALimpio; i++)
            {
                //cODIGO PARA QUE NO HAY VALORE NULLOS
                if (DatosPlataforma[i, 0] == null)
                {
                    DatosPlataforma[i, 0] = "";
                }
                if (DatosPlataforma[i, 1] == null)
                {
                    DatosPlataforma[i, 1] = "";
                }
                if (DatosPlataforma[i, 2] == null)
                {
                    DatosPlataforma[i, 2] = "";
                }
                if (DatosPlataforma[i, 3] == null)
                {
                    DatosPlataforma[i, 3] = "";
                }
                if (DatosPlataforma[i, 4] == null)
                {
                    DatosPlataforma[i, 4] = "";
                }
                if (DatosPlataforma[i, 5] == null)
                {
                    DatosPlataforma[i, 5] = "0";
                }
                if (DatosPlataforma[i, 6] == null)
                {
                    DatosPlataforma[i, 6] = "";
                }
                if (DatosPlataforma[i, 7] == null)
                {
                    DatosPlataforma[i, 7] = "";
                }
                DGVMTCenter.Rows.Add(DatosPlataforma[i, 0].ToString(), DatosPlataforma[i, 1].ToString(), DatosPlataforma[i, 2].ToString(), DatosPlataforma[i, 3].ToString(), DatosPlataforma[i, 4].ToString(), DatosPlataforma[i, 5].ToString(), DatosPlataforma[i, 6].ToString(), DatosPlataforma[i, 7].ToString());
            }
        }
        public void CincoMatch(DataGridView DGVPinesMatch)
        {
            for (int y = 0; y < DatosPlataforma.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosPlataforma[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                    {
                        string FechaMNGR = DatosMNGR[x, 1].ToString();
                        DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4) + " " + DatosMNGR[x, 2].ToString(), CultureInfo.InvariantCulture);
                        DateTime NEWFMGRLess = FMNGR.AddMinutes(-10);
                        DateTime NEWFMGRPlus = FMNGR.AddMinutes(10);
                        DateTime FMTC = Convert.ToDateTime(DatosPlataforma[y, 1].ToString() + " " + DatosPlataforma[y, 2].ToString());
                        float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                        float MontoMTC = float.Parse(DatosPlataforma[y, 6].ToString());
                        if (FMTC > NEWFMGRLess && FMTC < NEWFMGRPlus && DatosPlataforma[y, 3].ToString() == DatosMNGR[x, 3].ToString() && DatosPlataforma[y, 4].ToString() == DatosMNGR[x, 4].ToString() && DatosMNGR[x, 5].ToString().Contains(DatosPlataforma[y, 5].ToString()) && MontoMNGR == MontoMTC && DatosPlataforma[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVPinesMatch.Rows.Add("Plataforma", DatosPlataforma[y, 0].ToString(), DatosPlataforma[y, 1].ToString(), DatosPlataforma[y, 2].ToString(), DatosPlataforma[y, 3].ToString(), DatosPlataforma[y, 4].ToString(), DatosPlataforma[y, 5].ToString(), MontoMTC.ToString());
                            DatosPlataforma[y, 7] = "MATCH";
                            DGVPinesMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString());
                            DatosMNGR[x, 7] = "MATCH";
                            break;
                        }
                    }
                }
            }
        }
        public void CuatroMatch(DataGridView DGVPinesMatch)
        {
            for (int y = 0; y < DatosPlataforma.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosPlataforma[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                    {
                        string FechaMNGR = DatosMNGR[x, 1].ToString();
                        DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4) + " " + DatosMNGR[x, 2].ToString(), CultureInfo.InvariantCulture);
                        DateTime NEWFMGRLess = FMNGR.AddMinutes(-30);
                        DateTime NEWFMGRPlus = FMNGR.AddMinutes(30);
                        DateTime FMTC = Convert.ToDateTime(DatosPlataforma[y, 1].ToString() + " " + DatosPlataforma[y, 2].ToString());
                        float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                        float MontoMTC = float.Parse(DatosPlataforma[y, 6].ToString());
                        if (FMTC > NEWFMGRLess && FMTC < NEWFMGRPlus && DatosPlataforma[y, 4].ToString() == DatosMNGR[x, 4].ToString() && DatosMNGR[x, 5].ToString().Contains(DatosPlataforma[y, 5].ToString()) && MontoMNGR == MontoMTC && DatosPlataforma[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVPinesMatch.Rows.Add("Plataforma", DatosPlataforma[y, 0].ToString(), DatosPlataforma[y, 1].ToString(), DatosPlataforma[y, 2].ToString(), DatosPlataforma[y, 3].ToString(), DatosPlataforma[y, 4].ToString(), DatosPlataforma[y, 5].ToString(), MontoMTC.ToString());
                            DatosPlataforma[y, 7] = "MATCH";
                            DGVPinesMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString());
                            DatosMNGR[x, 7] = "MATCH";
                            break;
                        }
                    }
                }
            }
        }
        public void TresNoMatch(DataGridView DGVPinesNoMatch)
        {
            for (int y = 0; y < DatosPlataforma.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosPlataforma[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                    {
                        string FechaMNGR = DatosMNGR[x, 1].ToString();
                        DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                        DateTime FMTC = Convert.ToDateTime(DatosPlataforma[y, 1].ToString());
                        float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                        float MontoMTC = float.Parse(DatosPlataforma[y, 6].ToString());
                        if (FMTC == FMNGR && DatosPlataforma[y, 4].ToString() == DatosMNGR[x, 4].ToString() && DatosMNGR[x, 5].ToString().Contains(DatosPlataforma[y, 5].ToString()) && MontoMNGR == MontoMTC && DatosPlataforma[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVPinesNoMatch.Rows.Add("Plataforma", DatosPlataforma[y, 0].ToString(), DatosPlataforma[y, 1].ToString(), DatosPlataforma[y, 2].ToString(), DatosPlataforma[y, 3].ToString(), DatosPlataforma[y, 4].ToString(), DatosPlataforma[y, 5].ToString(), MontoMTC.ToString());
                            DatosPlataforma[y, 7] = "MATCH";
                            DGVPinesNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "ERROR EN LA REFERENCIA DE PAGO");
                            DatosMNGR[x, 7] = "MATCH";
                            break;
                        }
                    }
                }
            }
        }
        public void NoMatch(DataGridView DGVPinesNoMatch)
        {
            string caja = "";
            for (int y = 0; y < DatosPlataforma.GetLength(0); y++)
            {
                caja = DatosPlataforma[y, 0].ToString();
                //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                if (DatosPlataforma[y, 7] != "MATCH")
                {
                    DGVPinesNoMatch.Rows.Add("Plataforma", DatosPlataforma[y, 0].ToString(), DatosPlataforma[y, 1].ToString(), DatosPlataforma[y, 2].ToString(), DatosPlataforma[y, 3].ToString(), DatosPlataforma[y, 4].ToString(), DatosPlataforma[y, 5].ToString(), DatosPlataforma[y, 6].ToString(), "FALTA DAR DE BAJA");
                    DatosPlataforma[y, 7] = "MATCH";
                }
            }

            for (int x = 0; x < DatosMNGR.GetLength(0); x++)
            {
                if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                {
                    if (caja == DatosMNGR[x, 0].ToString())
                    {
                        if (DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVPinesNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), DatosMNGR[x, 6].ToString(), "ESTA DE MAS EN EL MANEJADOR");
                            DatosMNGR[x, 7] = "MATCH";
                        }
                    }
                }
            }
        }
        public void Revisar(DataGridView DGVPinesNoMatch)
        {
            DGVPinesNoMatch.Rows.Add("", "", " ", "", "", "", "");
            DGVPinesNoMatch.Rows.Add("", "", " ", "", "", "", "");
            DGVPinesNoMatch.Rows.Add("", "", " ", "", "", "", "");
            for (int x = 0; x < DatosMNGR.GetLength(0); x++)
            {
                if (DatosMNGR[x, 7] != "MATCH")
                {
                    DGVPinesNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), DatosMNGR[x, 6].ToString(), "ESTA DE MAS EN EL MANEJADOR");
                    DatosMNGR[x, 7] = "MATCH";
                }
            }
        }
    }
}
