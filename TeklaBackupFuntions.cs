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

namespace IFCexporter_3
{
    internal class TeklaBackupFuntions
    {
        public static string TextForBox = "";
        public static string TextForBox1 = "";
        public static string TextForBox2 = "";
        public static string TextForBox3 = "";//Texto do ActualizarModelo

        //Servidor de produção ou testes
        string servidorProd = "http://192.168.1.58";//Servidor de produção
        string servidorTest = "http://192.168.1.137"; //Servidor de testes

        enum MaterialTypes { Concrete, Steel }

        /// <summary>
        /// Esta função lê todos os números de pedido (Order Numbers) das peças de concreto no modelo Tekla,
        /// Esta a ler todas as peças de betão e a registar todos os OrderNumbers, mesmo que repetidos.
        /// </summary>
        /// <returns>Uma lista de strings contendo os números de pedido únicos.</returns>
        public static List<string> LerOrderNumber()//regista todos os OrderNumbers, mesmo que repetidos.
        {
            List<string> orderNumbers = new List<string>();
            Model model = new Model();

            if (model.GetConnectionStatus())
            {
                // Selecionar todas as peças no modelo
                //TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                //ModelObjectEnumerator allObjects = selector.GetAllObjects();

                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                ModelObjectEnumerator selectConcrete = selector.GetObjectsByFilterName("Material - Concrete");

                //Esta a ler todas as peças de betão e a registar todos os OrderNumbers, mesmo que repetidos.
                while (selectConcrete.MoveNext())
                {
                    Part part = selectConcrete.Current as Part;
                    if (part != null)
                    {
                        string orderNumber = string.Empty;
                        part.GetUserProperty("SM_Order_01", ref orderNumber);

                        if (!string.IsNullOrEmpty(orderNumber))
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


        public static List<string> LerOrderNumber1()//Selecionar apenas peças de betão
        {
            List<string> orderNumbers = new List<string>();
            Model model = new Model();

            if (model.GetConnectionStatus())
            {
                // Selecionar todas as peças no modelo
                //TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                //ModelObjectEnumerator allObjects = selector.GetAllObjects();



                //ModelObjectEnumerator selector = model.GetModelObjectSelector().GetObjectsByType(ModelObject.ModelObjectEnum.PART);


                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                ModelObjectEnumerator selectConcrete = selector.GetObjectsByFilterName("Material - Concrete");

                TextForBox = Environment.NewLine + "Please wait software is reading the parts" + Environment.NewLine;

                while (selectConcrete.MoveNext())
                {
                    Part part = selectConcrete.Current as Part;
                    if (part != null)
                    {
                        int X = 1;

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

                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
                MessageBox.Show("Failed to connect to Tekla Structures model.");
            }

            return orderNumbers;
        }


        private static void selecionarObjectosconcrete(string filterName)
        {
            Model myModel = new Model();
            if (myModel.GetConnectionStatus())
            {
                // Carregar o filtro existente para peças de betão
                filterName = "Material - Concrete"; // Nome do filtro existente para peças de betão
                TSM.ModelObjectSelector selector = myModel.GetModelObjectSelector();
                ModelObjectEnumerator myEnum = selector.GetObjectsByFilterName(filterName);

                while (myEnum.MoveNext())
                {
                    ModelObject myObject = myEnum.Current as ModelObject;
                    if (myObject != null)
                    {
                        Console.WriteLine(myObject.Identifier.ID);
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
            }
        }

        /// <summary>
        /// Esta função seleciona todas as peças do modelo Tekla que tiverem um determinado UDA.
        /// </summary>
        /// <param name="myUDA"></param>
        public static void SelecionarPorUDA(string myUDA)
        {
            Model myModel = new Model();
            var moe = myModel.GetModelObjectSelector().GetAllObjects();
            int nPlanes = 0;
            int nObjects = 0;

            while (moe.MoveNext())
            {
                moe.Current.SetUserProperty(myUDA, "");
                nObjects++;
                if (moe.Current is CutPlane)
                {
                    moe.Current.SetUserProperty(myUDA, "test");
                    nPlanes++;
                }
            }

            BinaryFilterExpression FilterExpression2 = new BinaryFilterExpression(new ObjectFilterExpressions.CustomString(myUDA), StringOperatorType.IS_NOT_EQUAL,
                                                                    new StringConstantFilterExpression(""));

            var moe2 = myModel.GetModelObjectSelector().GetObjectsByFilter(FilterExpression2);
            int nPlanes2 = 0;
            int nObjects2 = moe.GetSize();

            while (moe2.MoveNext())
            {
                if (moe2.Current is CutPlane) nPlanes2++;
            }

            bool result = nPlanes == nPlanes2;
            bool result2 = nObjects == nObjects2;
        }

        public static void ExportIFCFile2(string createFilePath)
        {
            Model Model = new Model();

            ComponentInput compInput = new ComponentInput();
            compInput.AddOneInputPosition(new TSG.Point(0, 0, 0));
            Component comp = new Component(compInput)
            {
                Name = "ExportIFC",
                Number = BaseComponent.PLUGIN_OBJECT_NUMBER
            };

            // This is code you can toggle to control whether pours are exported or not.
            // 0 = No
            // 1 = Yes
            bool PourManagement = false;
            Tekla.Structures.TeklaStructuresSettings.GetAdvancedOption("XS_ENABLE_POUR_MANAGEMENT", ref PourManagement);
            if (PourManagement)
                comp.SetAttribute("PourObjects", 1);

            comp.SetAttribute("OutputFile", createFilePath);
            comp.SetAttribute("Format", 0);
            comp.SetAttribute("ExportType", 1);
            comp.SetAttribute("CreateAll", 1);
            comp.SetAttribute("LocationBy", 0);
            comp.SetAttribute("Assemblies", 1);
            comp.SetAttribute("Bolts", 1);
            comp.SetAttribute("Welds", 0);
            comp.SetAttribute("GridExport", 1);
            comp.SetAttribute("ReinforcingBars", 1);
            comp.SetAttribute("SurfaceTreatments", 1);
            comp.SetAttribute("BaseQuantities", 1);
            comp.SetAttribute("LayersNameAsPart", 1);
            comp.SetAttribute("PLprofileToPlate", 1);
            comp.SetAttribute("ExcludeSnglPrtAsmb", 0);
            comp.SetAttribute("LocsFromOrganizer", 1);
            comp.SetAttribute("ViewColors", 0);

            if (!comp.Insert())
                Operation.DisplayPrompt("Exporting IFC File Failed.");
        }

        public static void ExportIFCFile1()
        {
            var model = new Tekla.Structures.Model.Model();
            if (model.GetConnectionStatus())
            {
                var componentInput = new ComponentInput();
                componentInput.AddOneInputPosition(new TSG.Point(0, 0, 0));

                var component = new Component(componentInput)
                {
                    Name = "JarvisIFC",
                    Number = BaseComponent.PLUGIN_OBJECT_NUMBER
                };

                component.SetAttribute("OutputFile", @"C:\TeklaStructuresModels\IFCfiles\Export.ifc");
                component.SetAttribute("Format", 0); // Escolher o formato desejado
                                                     // Configurações adicionais
                component.Insert();
            }

        }

        /// <summary>
        /// Esta função verifica se o Order Number coincide com o número do projecto.
        /// </summary>
        /// <param name="jarvisInfo"></param>
        /// <param name="timeIni"></param>
        public static void verificacao1(JarvisInfo jarvisInfo, DateTime timeIni)
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');

                // Iterar sobre os Jobs e Orders do jarvisInfo
                foreach (JarvisJob jarvisJob in jarvisInfo.JobList)
                {
                    if (jarvisJob.JobNo.TrimStart('0') == projectNumber)
                    {
                        Console.WriteLine($"Job {jarvisJob.JobNo} coincide com o número do projeto {projectNumber}");

                        foreach (JarvisOrder jarvisOrder in jarvisJob.OrderList)
                        {
                            // Verificar se o número do pedido coincide com o número do Order
                            if (jarvisOrder.OrderNo.TrimStart('0') == jarvisInfo.StatusMsg.TrimStart('0'))
                            {
                                Console.WriteLine($"Order {jarvisOrder.OrderNo} coincide com o número do pedido {jarvisInfo.StatusMsg}");
                            }
                            else
                            {
                                Console.WriteLine($"Order {jarvisOrder.OrderNo} não coincide com o número do pedido {jarvisInfo.StatusMsg}");
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Job {jarvisJob.JobNo} não coincide com o número do projeto {projectNumber}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
            }
        }

        public static void verificacao2(JarvisInfo jarvisInfo, DateTime timeIni)
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');

                // Iterar sobre os Jobs e Orders do jarvisInfo
                foreach (JarvisJob jarvisJob in jarvisInfo.JobList)
                {
                    if (jarvisJob.JobNo.TrimStart('0') == projectNumber)
                    {
                        Console.WriteLine($"Job {jarvisJob.JobNo} coincide com o número do projeto {projectNumber}");

                        foreach (JarvisOrder jarvisOrder in jarvisJob.OrderList)
                        {
                            // Selecionar todas as peças no modelo
                            TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                            ModelObjectEnumerator allObjects = selector.GetAllObjects();

                            while (allObjects.MoveNext())
                            {
                                Part part = allObjects.Current as Part;
                                if (part != null)
                                {
                                    string partOrderNo = string.Empty;
                                    part.GetUserProperty("SM_Order_01", ref partOrderNo);

                                    // Verificar se o número do pedido da peça coincide com o número do Order
                                    if (partOrderNo.TrimStart('0') == jarvisOrder.OrderNo.TrimStart('0'))
                                    {
                                        Console.WriteLine($"Peça {part.Identifier.ID} coincide com o número do pedido {jarvisOrder.OrderNo}");
                                        part.SetUserProperty("COMMENT", "match");
                                        part.Modify();
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Peça {part.Identifier.ID} não coincide com o número do pedido {jarvisOrder.OrderNo}");
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Job {jarvisJob.JobNo} não coincide com o número do projeto {projectNumber}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
            }
        }

        public static void verificacao(JarvisInfo jarvisInfo, DateTime timeIni)
        {
            TextForBox += "Verificação" + Environment.NewLine;

            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');

                // Iterar sobre os Jobs do jarvisInfo
                foreach (JarvisJob jarvisJob in jarvisInfo.JobList)
                {
                    if (jarvisJob.JobNo.TrimStart('0') == projectNumber)
                    {
                        Console.WriteLine($"Job {jarvisJob.JobNo} coincide com o número do projeto {projectNumber}");

                        // Iterar sobre os Orders do jarvisJob
                        foreach (JarvisOrder jarvisOrder in jarvisJob.OrderList)
                        {
                            // Selecionar todas as peças no modelo
                            TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                            ModelObjectEnumerator selectedObjects = selector.GetAllObjects();

                            while (selectedObjects.MoveNext())
                            {
                                Part part = selectedObjects.Current as Part;
                                if (part != null)
                                {
                                    string partOrderNo = string.Empty;
                                    string partCastNumber = string.Empty;
                                    part.GetUserProperty("SM_Order_01", ref partOrderNo);
                                    part.GetReportProperty("CAST_UNIT_POS", ref partCastNumber);

                                    // Verificar se o número do pedido e o número da peça coincidem com os do Order
                                    if (partOrderNo.TrimStart('0') == jarvisOrder.OrderNo.TrimStart('0'))
                                    {
                                        foreach (JarvisItem jarvisItem in jarvisOrder.ItemList)
                                        {
                                            if (partCastNumber.TrimStart('0') == jarvisItem.Reference.TrimStart('0'))
                                            {
                                                Console.WriteLine($"Peça {part.Identifier.ID} coincide com o número do pedido {jarvisOrder.OrderNo} e referência {jarvisItem.Reference}");
                                                part.SetUserProperty("comment", "MATCH");
                                                part.SetUserProperty("SM_Jarvis_Ava", "Avalible");
                                                part.Modify();
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Peça {part.Identifier.ID} não coincide com o número do pedido {jarvisOrder.OrderNo}");
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine($"Job {jarvisJob.JobNo} não coincide com o número do projeto {projectNumber}");
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
            }
        }

        /// <summary>
        /// Esta função atualiza os UDAs das peças no modelo Tekla com a informação do Jarvis.
        /// </summary>
        /// <param name="jarvisInfo"></param>
        public static void ActualizarModelo(JarvisInfo jarvisInfo)
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');

                // Selecionar todas as peças no modelo
                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                ModelObjectEnumerator allObjects = selector.GetAllObjects();

                while (allObjects.MoveNext())
                {
                    Part part = allObjects.Current as Part;
                    if (part != null)
                    {
                        string partMaterial = string.Empty;
                        part.GetReportProperty("MATERIAL_TYPE", ref partMaterial);

                        // Verificar se a peça é de betão
                        if (partMaterial.ToLower().Contains("CONCRETE"))
                        {
                            string partOrderNo = string.Empty;
                            string partCastNumber = string.Empty;
                            part.GetUserProperty("SM_Order_01", ref partOrderNo);
                            part.GetReportProperty("CAST_UNIT_POS", ref partCastNumber);

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
                                                    //part.SetUserProperty("SM_Jarvis_Ava", jarvisItem.Qty_ava.ToString());
                                                    //part.SetUserProperty("SM_Jarvis_DO", jarvisItem.Qty_do.ToString());
                                                    //part.SetUserProperty("SM_Jarvis_LN", jarvisItem.Qty_ln.ToString());
                                                    //part.SetUserProperty("SM_Jarvis_DN", jarvisItem.Qty_dn.ToString());
                                                    part.Modify();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
                TextForBox3 = "Failed to connect to Tekla Structures model.";
            }
        }

        public static void ActualizarModelo1_old(JarvisInfo jarvisInfo)
        {


            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');
                TextForBox3 = "Project Number: " + projectNumber + Environment.NewLine;

                TextForBox3 += Environment.NewLine + "Actualirzar" + Environment.NewLine;

                // Selecionar todas as peças no modelo
                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                ModelObjectEnumerator allObjects = selector.GetAllObjects();
                // foreach (ModelObject obj in allObjects) { TextForBox3 +=(obj.Identifier.ID)+Environment.NewLine; }

                while (allObjects.MoveNext())
                {
                    Part part = allObjects.Current as Part;
                    if (part != null)
                    {
                        string partMaterial = string.Empty;
                        part.GetReportProperty("MATERIAL", ref partMaterial);

                        //TextForBox3 += part.Identifier.ID +" | "+partMaterial+ Environment.NewLine;

                        // Verificar se a peça é de betão e se é a parte principal
                        if (partMaterial.ToLower().Contains("c"))// && part.IsMainPart())
                        {
                            string CastNumber = string.Empty;
                            string OrderNumber = string.Empty;
                            part.GetReportProperty("CAST_UNIT_POS", ref CastNumber);
                            part.GetReportProperty("SM_Order_01", ref OrderNumber);

                            TextForBox3 += CastNumber + " | " + OrderNumber + " | " + part.CastUnitType + " | " + part.Identifier.ID + " | " + partMaterial + Environment.NewLine;


                        }
                    }
                }
            }
            else
            {
                TextForBox3 = "Failed to connect to Tekla Structures model.";
            }

        }


        public static void ActualizarModelo1(JarvisInfo jarvisInfo, string ServidorJarvis)
        {
            string ServJarvis = ServidorJarvis;

            Model model = new Model();
            if (model.GetConnectionStatus())
            {

                // Obter o número do projeto do modelo Tekla
                string projectNumber = model.GetProjectInfo().ProjectNumber.TrimStart('0');
                TextForBox3 = "Project Number: " + projectNumber + Environment.NewLine;

                // Selecionar todas as peças no modelo
                TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
                ModelObjectEnumerator allObjects = selector.GetAllObjects();

                while (allObjects.MoveNext())
                {
                    TextForBox3 += "Please wait software is reading the parts" + Environment.NewLine;
                    Part part = allObjects.Current as Part;
                    if (part != null)
                    {
                        string partMaterial = string.Empty;
                        part.GetReportProperty("MATERIAL_TYPE", ref partMaterial);

                        // Verificar se a peça é de betão e se é a parte principal
                        if (partMaterial.ToLower().Contains("CONCRETE"))
                        {
                            string partOrderNo = string.Empty;
                            string partCastNumber = string.Empty;
                            part.GetUserProperty("SM_Order_01", ref partOrderNo);
                            part.GetReportProperty("CAST_UNIT_POS", ref partCastNumber);

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
                    }
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
                TextForBox3 = "Failed to connect to Tekla Structures model.";
            }
        }


        public static void ActualizarModelo2(JarvisInfo jarvisInfo, string ServidorJarvis)
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


        public static void BaseDadosFromModel()
        {
            TextForBox2 = Environment.NewLine + "Base de dados do modelo Tekla:";
            Console.WriteLine("Base de dados do modelo Tekla:");
            //NOTA: Design status: SENT, STOPPED, TO SEND
            List<Dictionary<string, object>> pecas = new List<Dictionary<string, object>> { };

            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Adicionar informações das peças
                pecas = new List<Dictionary<string, object>>
                {
                     new Dictionary<string, object>
                    {
                        { "Nome", "SW-7" },
                        { "Quantidade", 10 },
                        { "Manufacturado", 6 },
                        { "Ava", 1 },
                        { "Do", 1 },
                        { "LN", 1 },
                        { "DN", 0 }
                    },
                    new Dictionary<string, object>
                    {
                        { "Nome", "SW-8" },
                        { "Quantidade", 2 },
                        { "Manufacturado", 1 },
                        { "Ava", 1 },
                        { "Do", 1 },
                        { "LN", 1 },
                        { "DN", 0 }
                    },
                    new Dictionary<string, object>
                    {
                        { "Nome", "SW-10" },
                        { "Quantidade", 1 },
                        { "Manufacturado", 1 },
                        { "Ava", 1 },
                        { "Do", 1 },
                        { "LN", 1 },
                        { "DN", 1 }
                    }
                };

                foreach (var peca in pecas)
                {
                    Console.WriteLine($"Peça: {peca["Nome"]}, Quantidade: {peca["Quantidade"]}, Manufacturado: {peca["Manufacturado"]}, Ava: {peca["Ava"]}, Do: {peca["Do"]}, LN: {peca["LN"]}, DN: {peca["DN"]}");
                    TextForBox2 += ($"Peça: {peca["Nome"]}, Quantidade: {peca["Quantidade"]}, Manufacturado: {peca["Manufacturado"]}, Ava: {peca["Ava"]}, Do: {peca["Do"]}, LN: {peca["LN"]}, DN: {peca["DN"]}");
                }
            }
            else
            {
                Console.WriteLine("Failed to connect to Tekla Structures model.");
            }


            // Selecionar todas as peças no modelo
            TSM.ModelObjectSelector selector = model.GetModelObjectSelector();
            ModelObjectEnumerator allObjects = selector.GetAllObjects();

            while (allObjects.MoveNext())
            {
                Part part = allObjects.Current as Part;
                if (part != null)
                {
                    string partName = string.Empty;
                    string assNumb = string.Empty;
                    string CastNumber = string.Empty;
                    string OrderNumber = string.Empty;
                    part.GetReportProperty("NAME", ref partName);
                    part.GetReportProperty("AssemblyNumber", ref assNumb);
                    part.GetReportProperty("CAST_UNIT_POS", ref CastNumber);
                    part.GetReportProperty("SM_Order_01", ref OrderNumber);



                    Console.WriteLine(CastNumber + " " + partName + " " + assNumb + " " + OrderNumber);
                    TextForBox2 += (CastNumber + " " + partName + " " + assNumb + " " + OrderNumber) + Environment.NewLine;
                    // Verificar se a peça está na base de dados
                    var peca = pecas.FirstOrDefault(p => p["Nome"].ToString() == CastNumber);
                    if (peca != null)
                    {
                        // Atualizar o Design status para SENT
                        part.SetUserProperty("PLANS_STATUS", "SENT");
                        part.SetUserProperty("SM_Jarvis_Ord", "1");
                        part.SetUserProperty("SM_Jarvis_Ava", "Avalible");


                        // Atualizar o Fabrication status para Done se Manufacturado for diferente de 0
                        int quantidade = (int)peca["Quantidade"];
                        int manufacturado = (int)peca["Manufacturado"];
                        if (manufacturado > 0)
                        {
                            for (int i = 0; i < Math.Min(quantidade, manufacturado); i++)
                            {
                                part.SetUserProperty("FABRICATION_STATUS", "DONE");
                            }
                        }

                        part.Modify();
                    }
                }


            }
        }


    }

}

