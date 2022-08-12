using System;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class Servicios
    {
        RenArchivos ReM = new RenArchivos();
        Configuracion Config = new Configuracion();

        string[,] DatosMNGR;
        string[,] AServicios;
        string[,] CatalogoCajas;
        string[,] CatalogoServ;

        public void ConciliarServicios(string Conceptos, string ColumnasPlataforma, string MConceptos, string ColumnasManejador, ProgressBar PBar, DataGridView DGVManager, DataGridView DGVCatalogo, DataGridView DGVMTCenter, DataGridView DGVCajas, DataGridView DGVMatch, DataGridView DGVNoMatch)
        {
            CargarCatalogos(DGVCatalogo, DGVCajas);
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
                ServiciosMNGR(MConceptos, ColumnasManejador, DireccionCarpeta + "\\00.xls", DGVManager);
                Comparar_MTC_MNGR(Conceptos, ColumnasPlataforma, PBar, DireccionCarpeta + "\\MTC\\", DGVCatalogo, DGVMTCenter, DGVMatch, DGVNoMatch);
            }
            PBar.Minimum = 0;
        }

        //Identifica cuantos Reportes se descargaron de MTCenter, llama el methodo para cargar la información en el arreglo y posteriormente compara con los datos de Manager
        public void Comparar_MTC_MNGR(string Conceptos, string ColumnasPlataforma, ProgressBar PBar, string ruta, DataGridView DGVCatalogo, DataGridView DGVMTCenter, DataGridView DGVMatch, DataGridView DGVNoMatch)
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
                CargarDatosMTC(Conceptos, ColumnasPlataforma, ruta, cont, DGVCatalogo, DGVMTCenter);//Carga la información de los archivos uno a la vez
                CincoMatch(DGVMatch);
                CuatroMatch(DGVMatch);
                TresMatch(DGVMatch);
                TresNoMatch(DGVNoMatch);
                DosNoMatch(DGVNoMatch);
                UnoNoMatch(DGVNoMatch);
                TresMatchIusa(DGVNoMatch);
                DosMatchIusa(DGVNoMatch);
                UnoMatchIusa(DGVNoMatch);
                NoMatch(DGVNoMatch);
                PBar.Value = cont;
                cont++;
            }
            MessageBox.Show("Se procesaron " + cont + "Plus");
        }

        public void CargarCatalogos(DataGridView DGVCatalogo, DataGridView DGVCajas)
        {
            //GUARDA LOS DATOS DE LOS SERVICIOS EN UN ARREGLO                                            
            CatalogoServ = new string[DGVCatalogo.Rows.Count - 1, 3];
            for (int fila = 1; fila < DGVCatalogo.Rows.Count; fila++)
            {
                for (int col = 0; col < DGVCatalogo.Rows[fila].Cells.Count; col++)
                {
                    CatalogoServ[fila - 1, col] = DGVCatalogo.Rows[fila].Cells[col].Value.ToString();
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

        //Todo lo que se usa en la conciliación  de servicios
        //Lee el archivo 00, muestra el contenido en el Gridview y guarda la informacion en el arreglo DatosMNGR y cambia la caja con terminacion 2 a terminacion 1
        public void ServiciosMNGR(string MConceptos, string ColumnasManejador, string ExcelName, DataGridView DGVManager)
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
        //Lee uno por uno los reportes descargados de Manager los guarda en un arreglo y los muestra en el Gridview GVWMTCenter
        public void CargarDatosMTC(string Conceptos, string ColumnasPlataforma, string ruta, int Plus, DataGridView DGVCatalogo, DataGridView DGVMTCenter)
        {
            string[] ColP = ColumnasPlataforma.Split('/');
            string[] Concepto = Conceptos.Split('/');
            int tienda = 0, fecha = 0, Hora = 0, cajero = 0, clasificacion = 0, producto = 0, referencia1 = 0, entrada = 0;
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
                    if (LCont != "" && LCont.Contains(Concepto[1].ToString()) == true)
                    {
                        NTALimpio++;
                    }
                }
            }
            AServicios = new string[NTALimpio, 8];
            int CDLim = 0;
            //limpia la cadena y la guarda en un arreglo
            using (StreamReader readFile = new StreamReader(ruta + "reporteVentas (" + Plus + ").csv"))
            {
                string line;
                while ((line = readFile.ReadLine()) != null)
                {
                    if (line != "" && line.Contains(Concepto[1].ToString()) == true)
                    {
                        string NLine = line.ToString().Replace("\t\t<td>", "").Replace("\t", "").Replace("<td>", "°").Replace("</td>", "°").Replace("</tr>", "").Replace("<tr>", "").Replace("°°", "°").Replace("&nbsp;", " ").Replace("</table>", "").TrimEnd('°');
                        string[] ArregloLine = NLine.Split('°');
                        //cambia el concepto tienda a caja
                        for (int z = 0; z < CatalogoCajas.GetLength(0); z++)
                        {
                            string CajaMTC = ArregloLine[tienda].ToString().Replace(" S", "S");
                            string CajaCatalogo = CatalogoCajas[z, 0].ToString().Replace(" S", "S");
                            if (CajaMTC == CajaCatalogo)
                            {
                                AServicios[CDLim, 0] = CatalogoCajas[z, 1].ToString();
                            }
                        }
                        //Fecha y  Hora
                        if (ArregloLine[fecha].ToString() == ArregloLine[Hora].ToString())
                        {
                            string FechaHora = ArregloLine[fecha].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM");
                            int Caracter = FechaHora.IndexOf(" ");
                            AServicios[CDLim, 1] = FechaHora.Substring(0, Caracter);
                            int HrI = Caracter + 1;
                            int Dif = FechaHora.Length - HrI;
                            DateTime dt;
                            bool res = DateTime.TryParse(FechaHora.Substring(HrI, Dif), out dt);
                            AServicios[CDLim, 2] = dt.ToString("HH:mm:ss");
                        }
                        else
                        {
                            AServicios[CDLim, 1] = ArregloLine[fecha].ToString();
                            AServicios[CDLim, 2] = ArregloLine[Hora].ToString().Replace("p.m.", "PM").Replace("a.m.", "AM");
                        }
                        //vendedor
                        AServicios[CDLim, 3] = ArregloLine[cajero].ToString();
                        //Concepto y Entrada
                        for (int CCSRC = 0; CCSRC < DGVCatalogo.Rows.Count - 1/*NoSer*/; CCSRC++)
                        {
                            if (ArregloLine[producto].ToString() == CatalogoServ[CCSRC, 0].ToString())
                            {
                                AServicios[CDLim, 4] = CatalogoServ[CCSRC, 1].ToString();
                                string SComision = CatalogoServ[CCSRC, 2].ToString(), SImporte = ArregloLine[entrada].ToString();
                                float Comision = float.Parse(SComision.ToString()), Importe = float.Parse(SImporte.ToString());
                                float diferencia = Importe - Comision;
                                AServicios[CDLim, 6] = diferencia.ToString();
                            }
                        }
                        //Referencia
                        AServicios[CDLim, 5] = ArregloLine[referencia1].ToString();
                        //Estatus
                        AServicios[CDLim, 7] = "";
                        CDLim++;
                    }
                }
            }
            //MessageBox.Show(Plus.ToString());
            //Validar que no tengan valores NULL
            for (int i = 0; i < NTALimpio; i++)
            {
                //cODIGO PARA QUE NO HAY VALORE NULLOS
                if (AServicios[i, 1] == null)
                {
                    AServicios[i, 1] = "";
                }
                if (AServicios[i, 2] == null)
                {
                    AServicios[i, 2] = "";
                }
                if (AServicios[i, 3] == null)
                {
                    AServicios[i, 3] = "";
                }
                if (AServicios[i, 4] == null)
                {
                    AServicios[i, 4] = "";
                }
                if (AServicios[i, 5] == null)
                {
                    AServicios[i, 5] = "0";
                }
                if (AServicios[i, 6] == null)
                {
                    AServicios[i, 6] = "";
                }
                if (AServicios[i, 7] == null)
                {
                    AServicios[i, 7] = " ";
                }
                DGVMTCenter.Rows.Add(AServicios[i, 0].ToString(), AServicios[i, 1].ToString(), AServicios[i, 2].ToString(), AServicios[i, 3].ToString(), AServicios[i, 4].ToString(), AServicios[i, 5].ToString(), AServicios[i, 6].ToString(), AServicios[i, 7].ToString());
            }
        }

        //COMPARACION DE CAJA, FECHA Y HORA (60 MIN DIF - Y +), VENDEDOR, SERVICIO, MONTO y EVITA LOS QUE INCIAN CON 69 Y 70
        public void CincoMatch(DataGridView DGVMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                    {
                        //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                        if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                        {
                            //REORDENAR LA FECHA DEL MANAGER
                            string FechaMNGR = DatosMNGR[x, 1].ToString();
                            DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4) + " " + DatosMNGR[x, 2].ToString(), CultureInfo.InvariantCulture);
                            DateTime NEWFMGRLess = FMNGR.AddMinutes(-60);
                            DateTime NEWFMGRPlus = FMNGR.AddMinutes(60);
                            DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString() + " " + AServicios[y, 2].ToString());
                            float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                            float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                            //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                            if (FMTC > NEWFMGRLess && FMTC < NEWFMGRPlus && AServicios[y, 3].ToString() == DatosMNGR[x, 3].ToString() && AServicios[y, 4].ToString() == DatosMNGR[x, 4].ToString() && DatosMNGR[x, 5].ToString().Contains(AServicios[y, 5].ToString()) && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                DGVMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString());
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                        }
                    }
                }
            }
        }
        public void CuatroMatch(DataGridView DGVMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                    {
                        //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                        if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                        {
                            //REORDENAR LA FECHA DEL MANAGER
                            string FechaMNGR = DatosMNGR[x, 1].ToString();
                            DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4) + " " + DatosMNGR[x, 2].ToString(), CultureInfo.InvariantCulture);
                            DateTime NEWFMGRLess = FMNGR.AddMinutes(-60);
                            DateTime NEWFMGRPlus = FMNGR.AddMinutes(60);
                            DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString() + " " + AServicios[y, 2].ToString());
                            float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                            float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                            //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                            if (FMTC > NEWFMGRLess && FMTC < NEWFMGRPlus && AServicios[y, 4].ToString() == DatosMNGR[x, 4].ToString() && DatosMNGR[x, 5].ToString().Contains(AServicios[y, 5].ToString()) && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                DGVMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString());
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                        }
                    }
                }
            }

        }
        public void TresMatch(DataGridView DGVMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                    {
                        //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                        if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                        {
                            //REORDENAR LA FECHA DEL MANAGER
                            string FechaMNGR = DatosMNGR[x, 1].ToString();
                            DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                            DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString());
                            float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                            float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                            //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                            if (FMTC == FMNGR && DatosMNGR[x, 5].ToString().Contains(AServicios[y, 5].ToString()) && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                DGVMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString());
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                        }
                    }
                }
            }

        }
        public void TresNoMatch(DataGridView DGVNoMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                    {
                        //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                        if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                        {
                            //REORDENAR LA FECHA DEL MANAGER
                            string FechaMNGR = DatosMNGR[x, 1].ToString();
                            DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                            DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString());
                            float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                            float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                            //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                            if (FMTC == FMNGR && AServicios[y, 4].ToString() == DatosMNGR[x, 4].ToString() && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "ERROR EN LA REFERENCIA DE PAGO");
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                        }
                    }
                }
            }

        }
        public void UnoNoMatch(DataGridView DGVNoMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                    {
                        //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                        if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                        {
                            //REORDENAR LA FECHA DEL MANAGER
                            string FechaMNGR = DatosMNGR[x, 1].ToString();
                            DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                            DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString());
                            float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                            float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                            //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                            if (FMTC == FMNGR && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "Error en la referencia");
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                        }
                    }
                }
            }
        }
        public void DosNoMatch(DataGridView DGVNoMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                    {
                        //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                        if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                        {
                            //REORDENAR LA FECHA DEL MANAGER
                            string FechaMNGR = DatosMNGR[x, 1].ToString();
                            DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                            DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString());
                            float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                            float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                            //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                            if (MontoMNGR > MontoMTC && DatosMNGR[x, 5].ToString().Contains(AServicios[y, 5].ToString()) && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                float Diferencia = MontoMNGR - MontoMTC;
                                DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "SE DIO DE BAJA CON $" + Diferencia + " DE MAS");
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                            else if (MontoMNGR < MontoMTC && DatosMNGR[x, 5].ToString().Contains(AServicios[y, 5].ToString()) && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                            {
                                float Diferencia = MontoMNGR - MontoMTC;
                                DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                                AServicios[y, 7] = "MATCH";
                                DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "FALTA DAR DE BAJA $" + (Diferencia * -1));
                                DatosMNGR[x, 7] = "MATCH";
                                break;
                            }
                        }
                    }
                }
            }
        }
        public void TresMatchIusa(DataGridView DGVNoMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                    if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                    {
                        //REORDENAR LA FECHA DEL MANAGER
                        string FechaMNGR = DatosMNGR[x, 1].ToString();
                        DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                        DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString());
                        float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                        float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                        //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                        if (FMTC == FMNGR && DatosMNGR[x, 5].ToString().Contains(AServicios[y, 5].ToString()) && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                            AServicios[y, 7] = "MATCH";
                            DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "CAMBIO DE CONCEPTO A MTCENTER");
                            DatosMNGR[x, 7] = "MATCH";
                            break;
                        }
                    }
                }
            }

        }
        public void DosMatchIusa(DataGridView DGVNoMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                    if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                    {
                        //REORDENAR LA FECHA DEL MANAGER
                        string FechaMNGR = DatosMNGR[x, 1].ToString();
                        DateTime FMNGR = Convert.ToDateTime(FechaMNGR.Substring(3, 2) + "/" + FechaMNGR.Substring(0, 2) + "/" + FechaMNGR.Substring(6, 4), CultureInfo.InvariantCulture);
                        DateTime FMTC = Convert.ToDateTime(AServicios[y, 1].ToString());
                        float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                        float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                        //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                        if (FMTC == FMNGR && MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                            AServicios[y, 7] = "MATCH";
                            DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "CAMBIO DE CONCEPTO A MTCENTER");
                            DatosMNGR[x, 7] = "MATCH";
                            break;
                        }
                    }
                }
            }
        }
        public void UnoMatchIusa(DataGridView DGVNoMatch)
        {
            for (int y = 0; y < AServicios.GetLength(0); y++)
            {
                for (int x = 0; x < DatosMNGR.GetLength(0); x++)
                {
                    //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                    if (AServicios[y, 0].ToString() == DatosMNGR[x, 0].ToString())
                    {
                        float MontoMNGR = float.Parse(DatosMNGR[x, 6].ToString());
                        float MontoMTC = float.Parse(AServicios[y, 6].ToString());
                        //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                        if (MontoMNGR == MontoMTC && AServicios[y, 7] != "MATCH" && DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), MontoMTC.ToString());
                            AServicios[y, 7] = "MATCH";
                            DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), MontoMNGR.ToString(), "CAMBIO DE CONCEPTO A MTCENTER");
                            DatosMNGR[x, 7] = "MATCH";
                            break;
                        }
                    }
                }
            }
        }
        public void NoMatch(DataGridView DGVNoMatch)
        {
            string caja = "";
            int tam = AServicios.GetLength(0);
            for (int y = 0; y < tam; y++)
            {
                caja = AServicios[y, 0].ToString();
                //Match Perfecto Compara Fecha, Que el RPU sea el mismo, Que no esten empatados y que el monto sea el mismo
                if (AServicios[y, 7] != "MATCH")
                {
                    DGVNoMatch.Rows.Add("Plataforma", AServicios[y, 0].ToString(), AServicios[y, 1].ToString(), AServicios[y, 2].ToString(), AServicios[y, 3].ToString(), AServicios[y, 4].ToString(), AServicios[y, 5].ToString(), AServicios[y, 6].ToString(), "FALTA DAR DE BAJA");
                    AServicios[y, 7] = "MATCH";
                }
            }
            int tamano = DatosMNGR.GetLength(0);
            for (int x = 0; x < tamano; x++)
            {
                if (DatosMNGR[x, 5].ToString().StartsWith("69") == false && DatosMNGR[x, 5].ToString().StartsWith("70") == false)
                {
                    //valida que la caja sea la misma, que no tenga fechas vacias o que no tenga otro texto
                    if (caja == DatosMNGR[x, 0].ToString())
                    {
                        if (DatosMNGR[x, 7] != "MATCH")
                        {
                            DGVNoMatch.Rows.Add("Manager   ", DatosMNGR[x, 0].ToString(), DatosMNGR[x, 1].ToString(), DatosMNGR[x, 2].ToString(), DatosMNGR[x, 3].ToString(), DatosMNGR[x, 4].ToString(), DatosMNGR[x, 5].ToString(), DatosMNGR[x, 6].ToString(), "ESTA DE MAS EN EL MANEJADOR");
                            DatosMNGR[x, 7] = "MATCH";
                        }
                    }
                }
            }            
            //MessageBox.Show("");
        }
    }
}
