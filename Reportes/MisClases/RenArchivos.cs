using System;
using System.IO;
using System.Windows.Forms;

namespace Reportes.MisClases
{
    class RenArchivos
    {
        public void Renombrar(string ruta)
        {
            DirectoryInfo di = new DirectoryInfo(@ruta);
            foreach (var fi in di.GetFiles())
            {
                try
                {
                    string archivo = fi.Name.ToString();
                    string Nruta = ruta + archivo;
                    File.Move(Nruta, Path.ChangeExtension(Nruta, ".csv"));
                }
                catch (Exception e)
                {
                    MessageBox.Show("Error, Verificar el archivo", e.Message);
                }
            }
        }
    }
}
