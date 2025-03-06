using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Tekla.Structures;
using Tekla.Structures.Model;
using Tekla.Structures.Model.Operations;
using TS = Tekla.Structures;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using System.Windows.Forms;
using Tekla.Structures.Model.UI;
using View = Tekla.Structures.Model.UI.View;
using GetJarvisInfoTest;
using static System.Net.WebRequestMethods;
using Tekla.Structures.RemotingHelper;
using ModelObjectSelector = Tekla.Structures.Model.UI.ModelObjectSelector;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Filtering;
using System.Windows.Forms.VisualStyles;
using System.Diagnostics;



namespace IFCexporter_3
{
    internal class TeklaFuncoes
    {
        public static string TextForBox = "";
        public static string TextForBox1 = "";
        public static string TextForBox2 = "";
        public static string TextForBox3 = "";//Texto do ActualizarModelo

        //Servidor de produção ou testes
        string servidorProd = "http://192.168.1.58";//Servidor de produção
        string servidorTest = "http://192.168.1.137"; //Servidor de testes

        enum MaterialTypes { Concrete, Steel }

        public static void TeklaConnection()
        {
            Model model = new Model();
            try
            {
                if (model.GetConnectionStatus())
                {
                    Console.WriteLine("Connected to Tekla Structures model.");
                    TextForBox += "Connected to Tekla Structures model." + Environment.NewLine;
                }
                else
                {
                    Console.WriteLine("Failed to connect to Tekla Structures model.");
                    TextForBox += "Failed to connect to Tekla Structures model." + Environment.NewLine;
                }
            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
                Console.WriteLine(e.Message);
                TextForBox += "Failed to connect to Tekla Structures model." + Environment.NewLine;
                TextForBox += e.Message;

            }

        }
        public static void TeklaModelInfo()
        {
            Model myModel = new Model();

            try
            {
                myModel.GetConnectionStatus();

                Console.WriteLine("Project Number: " + myModel.GetProjectInfo().ProjectNumber);
                Console.WriteLine("Project Name: " + myModel.GetProjectInfo().Name);
                Console.WriteLine("Model name: " + myModel.GetInfo().ModelName);
                Console.WriteLine("Model path: " + myModel.GetInfo().ModelPath);

                TextForBox += "Project Number: " + myModel.GetProjectInfo().ProjectNumber + Environment.NewLine;
                TextForBox += "Project Name: " + myModel.GetProjectInfo().Name + Environment.NewLine;
                TextForBox += "Model name: " + myModel.GetInfo().ModelName + Environment.NewLine;
                TextForBox += "Model path: " + myModel.GetInfo().ModelPath + Environment.NewLine;

            }
            catch (Exception e)
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
                Console.WriteLine(e.Message);
            }
        }
        public static void ExportIFCFile(string folderPath, string fileName)
        {
            string SelectedFormat_Name = string.Empty;
            int SelectedFormat_Number = 0;
            string SelectedFile = string.Empty;

            
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                ModelInfo modelInfo = model.GetInfo();

                folderPath = modelInfo.ModelPath + @"\IFC_Jarvis_Status\";

                // Verificar se a pasta existe, se não, criar a pasta
                 if (!System.IO.Directory.Exists(folderPath))
                 {
                 System.IO.Directory.CreateDirectory(folderPath);
                 }
                 // Obtém a data e hora atual do sistema
                DateTime dataAtual = DateTime.Now;

                // Converte a data para uma string formatada
                string dataString = dataAtual.ToString("yyyy-MM-dd");

                //MessageBox.Show(model.GetProjectInfo().ProjectNumber+ "");
                string PNumber = model.GetProjectInfo().ProjectNumber;

                fileName = PNumber +"_"+ dataString;
             
                SelectedFormat_Number = Tekla.Structures.Model.BaseComponent.PLUGIN_OBJECT_NUMBER; //Mandatory
                SelectedFormat_Name = "ExportIFC"; // Mandatory
                

                var componentInput = new Tekla.Structures.Model.ComponentInput();
                componentInput.AddOneInputPosition(new Tekla.Structures.Geometry3d.Point(0, 0, 0));
                var component = new Tekla.Structures.Model.Component(componentInput);
                component.Name = SelectedFormat_Name;
                component.Number = SelectedFormat_Number;
                component.LoadAttributesFromFile(SelectedFile);
                component.SetAttribute("CreateAll", 1);// 0 = selecionado, 1 = tudo
                component.SetAttribute("Format", 2); // 0 = ifc, 1 = ifc xml, 2 = ifc zip
                component.SetAttribute("OutputFile", folderPath + fileName);
             

                // Abre a pasta no Windows Explorer
                Process.Start("explorer.exe", folderPath);

                if (component.Insert())
                {
                    // MessageBox.Show("Export component inserted correctly!", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    TextForBox += "Export component inserted correctly!" + Environment.NewLine;
                }
                else
                {
                    MessageBox.Show("Error when inserting export component", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        public static void RedrawView()
        {


            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                model.CommitChanges();
                model.GetWorkPlaneHandler().SetCurrentTransformationPlane(new TransformationPlane());
                model.CommitChanges();
            }

        }

        public static void ViewRepresentation()
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Obter todas as vistas visíveis
                ViewHandler viewHandler = new ViewHandler();
                ModelViewEnumerator viewEnumerator = ViewHandler.GetVisibleViews();

                if (viewEnumerator.MoveNext())
                {
                    View currentView = viewEnumerator.Current as View;

                    if (currentView != null)
                    {
                        // Definir a propriedade de representação "status"
                        currentView.CurrentRepresentation = "Status_Jarvis";

                        // Atualizar a vista
                        currentView.Modify();
                        ViewHandler.RedrawView(currentView);
                    }
                    else
                    {
                        Console.WriteLine("Failed to get the current view.");
                    }
                }
                else
                {
                    Console.WriteLine("No visible views found.");
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
            }
        }

        public static void Numbering()
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                Tekla.Structures.ModelInternal.Operation.dotStartAction("FullNumbering", (string)null);
            }
            else
            {
                MessageBox.Show("Failed to connect to Tekla Structures model.");
                Environment.Exit(1);
            }
            Console.WriteLine("\nNumbering done\n");
        }

        public static void StatusFromJarvis(string ProjectNumber, string OrderNumber)
        {
            GetJarvisInfoTest.GetJarvisInfo GetJarvisInfo = new GetJarvisInfoTest.GetJarvisInfo();
        }

        public static void UpdateOutputText(JarvisInfo jarvisInfo, DateTime timeIni)
        {
            int num = 0;
            const int MAX_ITEMS_TO_SHOW = 200;

            string Info1 = "Web Service call result = " + jarvisInfo.StatusCode + Environment.NewLine +
                   "Returned Message = " + jarvisInfo.StatusMsg + Environment.NewLine +
                   "Number of returned items = " + jarvisInfo.ItemsCount + Environment.NewLine +
                   "Time waiting WS return + parsing JSON object =  " + (DateTime.Now - timeIni).ToString() + Environment.NewLine
                   ;

            Info1 += Environment.NewLine + "Items Info: " + Environment.NewLine;
            //TextForBox1 = "Teste ABC";

            timeIni = DateTime.Now;
            if (jarvisInfo.ItemsCount > MAX_ITEMS_TO_SHOW)
            {
                Info1 += Environment.NewLine + "   ----- Showing only first " + MAX_ITEMS_TO_SHOW + " items ------------" + Environment.NewLine + Environment.NewLine;
            }
            if (jarvisInfo.JobList != null)
            {
                foreach (JarvisJob jarvisJob in jarvisInfo.JobList)
                {
                    Info1 += " * Job: " + jarvisJob.JobNo + Environment.NewLine;
                    if (jarvisJob.OrderList != null)
                    {
                        foreach (JarvisOrder jarvisOrder in jarvisJob.OrderList)
                        {
                            Info1 += "  + Order: " + jarvisOrder.OrderNo + Environment.NewLine;
                            if (jarvisOrder.ItemList != null)
                            {
                                foreach (JarvisItem jarvisItem in jarvisOrder.ItemList)
                                {
                                    if (num > MAX_ITEMS_TO_SHOW) break;
                                    Info1 += "   - (Item " + num.ToString("0000") + ")   Product: '" + jarvisItem.Name + "' - Ref: '" + jarvisItem.Reference + "' - Q.O.: " + jarvisItem.Qty_ord + " - Q.M.: " + jarvisItem.Qty_man + " - Q.D.: " + jarvisItem.Qty_del + "" + Environment.NewLine;
                                    num++;
                                }
                            }
                        }
                    }
                }

            }
            if (num > 0)
            {
                Info1 += "Time printing results =  " + (DateTime.Now - timeIni).ToString();
            }
            Console.WriteLine(Info1);
            TextForBox1 += Info1;
        }

        /// <summary>
        /// Esta função atualiza os User Defined Attributes (UDAs) das peças no modelo Tekla com a informação do Jarvis.
        /// </summary>
        /// <param name="jarvisInfo"></param>
        /// <param name="ServidorJarvis"></param>
        public static void ActualizarModelo(JarvisInfo jarvisInfo, string ServidorJarvis)
        {
            string ServJarvis = ServidorJarvis;

            Model model = new Model();
            if (model.GetConnectionStatus())
            {

                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');
                TextForBox3 = "Project Number: " + projectNumber + Environment.NewLine;

                string FilterName = "Material - Concrete";

                // Selecionar todas as peças no modelo
                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                TSM.ModelObjectEnumerator selectByFilter = selector.GetObjectsByFilterName(FilterName);

                while (selectByFilter.MoveNext())
                {

                    Part part = selectByFilter.Current as Part;

                    if (part != null)
                    {
                        TextForBox3 += part.Profile.ProfileString + " | " + part.Name + Environment.NewLine;



                        string partOrderNo = string.Empty;
                        string partCastNumber = string.Empty;
                        part.GetUserProperty("SM_Order_01", ref partOrderNo);
                        part.GetReportProperty("CAST_UNIT_POS", ref partCastNumber);


                        if (jarvisInfo.JobList != null)
                        {

                            // Iterar sobre os Jobs do jarvisInfo
                            foreach (JarvisJob jarvisJob in jarvisInfo.JobList)
                            {
                                if (jarvisJob.JobNo.TrimStart('0') == projectNumber)
                                {
                                    // Iterar sobre os Orders do jarvisJob
                                    foreach (JarvisOrder jarvisOrder in jarvisJob.OrderList)
                                    {
                                        // Verificar se o número do pedido e o número da peça coincidem com os do Order
                                        if (partOrderNo.TrimStart('0') == jarvisOrder.OrderNo.TrimStart('0'))
                                        {
                                            foreach (JarvisItem jarvisItem in jarvisOrder.ItemList)
                                            {
                                                if (partCastNumber.TrimStart('0') == jarvisItem.Reference.TrimStart('0'))
                                                {
                                                    // Atualizar os UDAs da peça com a informação do Jarvis
                                                    part.SetUserProperty("SM_Jarvis_Ord", jarvisItem.Qty_ord.ToString());
                                                    part.SetUserProperty("SM_Jarvis_Man", jarvisItem.Qty_man.ToString());
                                                    part.SetUserProperty("SM_Jarvis_Ava", "");
                                                    part.SetUserProperty("SM_Jarvis_DO", "");
                                                    part.SetUserProperty("SM_Jarvis_LN", "");
                                                    part.SetUserProperty("SM_Jarvis_DN", jarvisItem.Qty_del.ToString());
                                                    part.SetUserProperty("SM_Jarvis_Rlin", jarvisItem.Rlin.ToString());
                                                    part.SetUserProperty("SM_Jarvis_Rlinpath", ServJarvis + "/intranet/w_pedidos/consulta-fab-alb.php?rlin=" + jarvisItem.Rlin + jarvisItem.Rlin.ToString());


                                                    part.Modify();

                                                    // Adicionar ao TextForBox3                            
                                                    TextForBox3 += $"Peça {jarvisItem.Reference} coincide com o número do pedido {jarvisOrder.OrderNo} | Qty_Ord {jarvisItem.Qty_ord} | Man {jarvisItem.Qty_man} | DN  - delivery note {jarvisItem.Qty_del}" + Environment.NewLine;



                                                }
                                                else
                                                {
                                                    TextForBox = "Nao escreveu nas peças";
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        }
                        else
                        {
                            // Handle the case when JobList is null
                            TextForBox3 = "JobList is null.";
                        }
                    }
                    else
                    {

                    }




                }
            }
        }

        /// <summary>
        /// Esta função lê todos os números de pedido (Order Numbers) das peças de concreto no modelo Tekla,
        /// garantindo que valores repetidos não sejam adicionados à lista.
        /// </summary>
        /// <returns>Uma lista de strings contendo os números de pedido únicos.</returns>
        public static List<string> LerOrderNumber()
        {
            List<string> orderNumbers = new List<string>();
            Model model = new Model();

            if (model.GetConnectionStatus())
            {
                // Selecionar todas as peças no modelo
                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                ModelObjectEnumerator selectConcrete = selector.GetObjectsByFilterName("Material - Concrete");

                // Esta a ler todas as peças de betão e a registar todos os OrderNumbers, mas não guarda repetidos.
                while (selectConcrete.MoveNext())
                {
                    Part part = selectConcrete.Current as Part;
                    if (part != null)
                    {
                        string orderNumber = string.Empty;
                        part.GetUserProperty("SM_Order_01", ref orderNumber);

                        if (!string.IsNullOrEmpty(orderNumber) && !orderNumbers.Contains(orderNumber))
                        {
                            orderNumbers.Add(orderNumber);
                        }
                    }
                }
            }
            else
            {
                //Console.WriteLine("Failed to connect to Tekla Structures model.");
                MessageBox.Show("Failed to connect to Tekla Structures model.");
            }

            return orderNumbers;
        }

        public static List<string> SelectByFilter(string FilterName)
        {
            List<string> partNames = new List<string>();
            Model model = new Model();

            TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
            TSM.ModelObjectEnumerator selectByFilter = selector.GetObjectsByFilterName(FilterName);

            while (selectByFilter.MoveNext())
            {
                Part part = selectByFilter.Current as Part;
                if (part != null)
                {
                    TextForBox += "";

                    string material = string.Empty;
                    string materialType = string.Empty;
                    part.GetReportProperty("MATERIAL", ref material);
                    part.GetReportProperty("MATERIAL_TYPE", ref materialType);

                    TextForBox += part.Name + " | " + materialType + " | " + material + Environment.NewLine;

                    // Verificar se o material é "concrete"
                    if (materialType.ToLower().Contains("CONCRETE"))
                    {
                        TextForBox += " NOME: " + part.Name + " | " + material + Environment.NewLine;
                        partNames.Add(part.Name);
                    }
                }
            }

            return partNames;
        }
        public static List<string> SelectMaterialType(string materialType)
        {
            List<string> partNames = new List<string>();
            Model model = new Model();

            if (model.GetConnectionStatus())
            {
                BinaryFilterExpression FilterExpression2 = new BinaryFilterExpression(new ObjectFilterExpressions.CustomString("MATERIAL_TYPE"), StringOperatorType.IS_EQUAL,
                                                                   new StringConstantFilterExpression("CONCRETE"));




                var moe2 = model.GetModelObjectSelector().GetObjectsByFilter(FilterExpression2);


                while (moe2.MoveNext())
                {
                    TextForBox += " SelectMaterialType entrou no while " + Environment.NewLine;

                    Part part = moe2.Current as Part;
                    if (part != null)
                    {
                        TextForBox += "";
                        string material = string.Empty;
                        materialType = string.Empty;
                        part.GetReportProperty("MATERIAL", ref material);
                        part.GetReportProperty("MATERIAL_TYPE", ref materialType);
                        TextForBox += part.Name + " | " + materialType + " | " + material + Environment.NewLine;
                        // Verificar se o material é "concrete"
                        if (materialType.ToLower().Contains("CONCRETE"))
                        {
                            TextForBox += " NOME: " + part.Name + " | " + material + Environment.NewLine;
                            partNames.Add(part.Name);
                        }
                        else
                        {
                            TextForBox += " ERROR " + Environment.NewLine;
                        }
                    }
                }
                TextForBox += " Terminou a função SelectMaterialType " + Environment.NewLine;

            }
            return partNames;
        }

        /// <summary>
        /// Executa uma macro gravada no Tekla Structures.
        /// </summary>
        public static void CorrerMacro(string Caminho)
        {
            //exemplo: "C:\Shaymurtagh_Programs\TEKLA\FIRM_MURTAGH\MACROS\modeling\CopyToNewSheet.cs"
            //string macroPath = @"C:\Caminho\Para\Sua\Macro.cs"; // Substitua pelo caminho real da sua macro
            string macroPath = Caminho;

            try
            {
                Operation.RunMacro(macroPath);
                MessageBox.Show("Macro executada com sucesso.", "Execução de Macro", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao executar a macro: {ex.Message}", "Erro de Execução", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public static void CorrerMacro1()
        {
            string NomeMacro = "CopyToNewSheet.cs";

            Model model = new Model();
            string macroPath = @"C:\Shaymurtagh_Programs\TEKLA\FIRM_MURTAGH\MACROS\modeling\" + NomeMacro; // Caminho da macro

            Operation.RunMacro(macroPath);
        }
    }
}


