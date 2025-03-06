using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using GetJarvisInfoTest;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Tekla.Structures.Model;
using Tekla.Structures.Drawing;

namespace IFCexporter_3
{
    public partial class Form1 : Form
    {

        //Servidor de produção ou testes
        public static string servidorProd = "http://192.168.1.58";//Servidor de produção
        public static string servidorTest = "http://192.168.1.137"; //Servidor de testes

        public Form1()
        {
            InitializeComponent();

            // Escrever "teste" na textBox1
            textResult.Text += "Start Application" + Environment.NewLine;


            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                textResult.Text += "Tekla Structures model connected" + Environment.NewLine + Environment.NewLine;
                //CorrerPrograma();

                TeklaFuncoes.TeklaModelInfo();
                textResult.Text += TeklaFuncoes.TextForBox;

                TeklaFuncoes.ViewRepresentation();
            }
            else
            {
                textResult.Text += "Failed to connect to Tekla Structures model." + Environment.NewLine;
                //MessageBox.Show("Failed to connect to Tekla Structures model.");
                //Environment.Exit(1);
            }
        }

        private void CorrerPrograma()
        {
            textResult.Text += "Correr programa" + Environment.NewLine;

            Model model = new Model();
            if (model.GetConnectionStatus())
            {

                //Console.WriteLine("Starting program");
                textResult.Text += "Refresh pressed" + Environment.NewLine;

                TeklaFuncoes.Numbering();
                //TeklaFuncoes.BaseDadosFromModel();

                //Aqui é onde o usuário vai informar o número do pedido e a referência
                textResult.Text += Environment.NewLine + "Start Getting Jarvis Info";
                textResult.Text += Environment.NewLine + "xxxxxxxxxxxxxxxxxxxxxxxxx";
                DateTime timeIni = DateTime.Now;
                string textOrderNo = "";//Exemplo "SH173788";
                string textReference = "";

                // Vamos ler a lista de OrderNumbers e escrever num string
                List<string> orderNumbers = TeklaFuncoes.LerOrderNumber();

                foreach (string orderNumber in orderNumbers)
                {
                    textResult.Text += Environment.NewLine + "Order Number: " + orderNumber;

                    textOrderNo = orderNumber;

                    JarvisInfo jarvisInfo = GetJarvisInfo.GetItemsInfoByOrder(textOrderNo, textReference);
                    TeklaFuncoes.UpdateOutputText(jarvisInfo, timeIni);

                }
                textResult.Text += Environment.NewLine + "Verificação" + Environment.NewLine;
                textResult.Text += TeklaFuncoes.TextForBox1;
                textResult.Text += Environment.NewLine + "Verificação Concluida" + Environment.NewLine;
                textResult.Text += TeklaFuncoes.TextForBox2;
            }
            else
            {
                MessageBox.Show("Failed to connect to Tekla Structures model.");
                Environment.Exit(1);
            }
            Console.WriteLine("\nNumbering done\n");
        }

        private void Refresh_Click(object sender, EventArgs e)//Botão Refresh, Lê os objecto do modelo e mostra na caixa de texto
        {
            // Limpar o conteúdo da textResult
            textResult.Text += Environment.NewLine + "Refresh was pressed" + Environment.NewLine;

            // Executar o programa
            CorrerPrograma();

        }

        private void Update_Click(object sender, EventArgs e)//Botão Update, actualiza o modelo.
        {
            // Limpar o conteúdo da textResult
            textResult.Text += "";

            string textOrderNo = "";//Exemplo "SH173788";
            string textReference = "";
            DateTime timeIni = DateTime.Now;

            // Vamos ler a lista de OrderNumbers e escrever num string
            List<string> orderNumbers = TeklaFuncoes.LerOrderNumber();


            foreach (string orderNumber in orderNumbers)
            {
                textResult.Text += Environment.NewLine + "Order Number: " + orderNumber;

                textOrderNo = orderNumber;

                JarvisInfo jarvisInfo = GetJarvisInfo.GetItemsInfoByOrder(textOrderNo, textReference);
                TeklaFuncoes.UpdateOutputText(jarvisInfo, timeIni);

                string serv = servidorTest;
                //TeklaFuncoes.ActualizarModelo1(jarvisInfo, serv);
                TeklaFuncoes.ActualizarModelo(jarvisInfo, serv);
            }

            textResult.Text = TeklaFuncoes.TextForBox3;

            TeklaFuncoes.ViewRepresentation();
        }

        private void IFC_Export_Click(object sender, EventArgs e)
        {


            textResult.Text += "Start Export IFC";
            string folderPath = @"C:\TeklaStructuresModels\IFCfiles\";
            string fileName = "ExportX.ifc";

            //Exporta o arquivo IFC
            TeklaFuncoes.ExportIFCFile(folderPath, fileName);
        }

        private void Drawing_Schedule_Click_1(object sender, EventArgs e)
        {
            textResult.Text = TeklaDrawings.TextForBox1;
            //TeklaDrawings.DrawingsVerification();
            

            TeklaDrawings.SelectDrawing1();
            


        }

        private void NewDWG_Click(object sender, EventArgs e)
        {
            //Criar no novo desenho
            TeklaDrawings.CriarNovoDesenho();
        }
    }
}
