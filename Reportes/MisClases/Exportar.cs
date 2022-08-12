using System;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class Exportar
    {
        public void ExportarDataGridViewExcel(DataGridView grd, string Nombre)
        {
            try
            {
                SaveFileDialog archivo = new SaveFileDialog();
                archivo.Filter = "Excel (*.xls)|*.xls";
                archivo.FileName = "Reporte de Servicios " + Nombre + DateTime.Now.Date.ToShortDateString().Replace('/', '-');
                if (archivo.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libroDeTrabajo;
                    Microsoft.Office.Interop.Excel.Worksheet hojaDeTrabajo;
                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libroDeTrabajo = aplicacion.Workbooks.Add();
                    hojaDeTrabajo = (Microsoft.Office.Interop.Excel.Worksheet)libroDeTrabajo.Worksheets.get_Item(1);

                    hojaDeTrabajo.Cells[1, "A"] = grd.Columns[0].HeaderText;
                    hojaDeTrabajo.Cells[1, "B"] = grd.Columns[1].HeaderText;
                    hojaDeTrabajo.Cells[1, "C"] = grd.Columns[2].HeaderText;
                    hojaDeTrabajo.Cells[1, "D"] = grd.Columns[3].HeaderText;
                    hojaDeTrabajo.Cells[1, "E"] = grd.Columns[4].HeaderText;
                    hojaDeTrabajo.Cells[1, "F"] = grd.Columns[5].HeaderText;
                    hojaDeTrabajo.Cells[1, "G"] = grd.Columns[6].HeaderText;


                    hojaDeTrabajo.Columns[1].AutoFit();
                    hojaDeTrabajo.Columns[2].AutoFit();
                    hojaDeTrabajo.Columns[3].AutoFit();
                    hojaDeTrabajo.Columns[4].AutoFit();
                    hojaDeTrabajo.Columns[5].AutoFit();
                    hojaDeTrabajo.Columns[6].AutoFit();
                    hojaDeTrabajo.Columns[7].AutoFit();

                    hojaDeTrabajo.Name = Nombre;
                    //Recorremos el DataGridView rellenando la hoja de trabajo
                    for (int i = 0; i < grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            if (grd.Rows[i].Cells[j].Value != null)
                            {
                                hojaDeTrabajo.Cells[i + 2, j + 1] = grd.Rows[i].Cells[j].Value.ToString();
                            }
                        }
                    }
                    libroDeTrabajo.SaveAs(archivo.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libroDeTrabajo.Close(true);
                    aplicacion.Quit();
                    MessageBox.Show("Registro de " + Nombre + " Exportado a Excel", "Proceso finalizado");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar la informacion debido a: " + ex.ToString());
            }
        }
    }
}
