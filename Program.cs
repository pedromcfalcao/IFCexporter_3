using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GetJarvisInfoTest;
using Tekla.Structures.Datatype;

namespace IFCexporter_3
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // Abrir o Form1 no início da aplicação
            Application.Run(new Form1());

           
        }

        
    }
}
