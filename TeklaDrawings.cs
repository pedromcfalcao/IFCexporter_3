using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Tekla.Structures.Drawing;
using Tekla.Structures.Model;
using Tekla.Structures.Model.UI;
using TSM = Tekla.Structures.Model;
using TSG = Tekla.Structures.Geometry3d;
using TSD = Tekla.Structures.Drawing;
using Tekla.Structures.Geometry3d;
using Line = Tekla.Structures.Drawing.Line;
using Tekla.Structures.Model.Geometry;
using Tekla.Structures;
using Tekla;
using Tekla.Macros;
using Tekla.Structures.Filtering.Categories;
using Tekla.Structures.Filtering;


namespace IFCexporter_3
{
    internal class TeklaDrawings
    {
        public static string TextForBox1 = "Starting Drawings " + Environment.NewLine;

        /// <summary>
        /// Verifica se existe um desenho com o título '3D Schedule' e copia o desenho para uma nova folha.
        /// </summary>
        public static void DrawingsVerification()
        {
            TextForBox1 += "Starting Drawings Verification" + Environment.NewLine;

            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                DrawingHandler drawingHandler = new DrawingHandler();
                DrawingEnumerator drawingEnumerator = drawingHandler.GetDrawings();

                while (drawingEnumerator.MoveNext())
                {
                    Drawing drawing = drawingEnumerator.Current as Drawing;
                    if (drawing != null)
                    {
                        string drawingTitle = string.Empty;
                        drawingTitle = drawing.Name;

                        if (drawingTitle == "3D Schedule")
                        {
                            MessageBox.Show("Desenho com o título '3D Schedule' encontrado.", "Desenho Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                            //Copiar o desenho para uma nova folha
                            CopyDrawingToNewSheet(drawing);
                            
                            break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Não foi possível conectar ao Tekla Structures.", "Erro de Conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
        }

        /// <summary>
        /// Procura um desenho que tenha no Name ou Title1 a palavra "schedule".
        /// </summary>
        public static void SelectDrawing()
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                DrawingHandler drawingHandler = new DrawingHandler();
                DrawingEnumerator drawingEnumerator = drawingHandler.GetDrawings();

                while (drawingEnumerator.MoveNext())
                {
                    Drawing drawing = drawingEnumerator.Current as Drawing;
                    if (drawing != null)
                    {
                        string drawingName = drawing.Name.ToLower();
                        string drawingTitle1 = drawing.Title1.ToLower();

                        if (drawingName.Contains("schedule") || drawingTitle1.Contains("schedule"))
                        {
                            MessageBox.Show($"Desenho encontrado: {drawing.Name}", "Desenho Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            // Aqui você pode adicionar qualquer lógica adicional que deseja executar quando o desenho for encontrado
                            
                            
                            TeklaFuncoes.CorrerMacro(@"C:\Shaymurtagh_Programs\TEKLA\FIRM_MURTAGH\MACROS\modeling\CopyToNewSheet.cs");

                            break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Não foi possível conectar ao Tekla Structures.", "Erro de Conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// Procura e seleciona um desenho que tenha no Name ou Title1 a palavra "schedule".
        /// </summary>
        public static void SelectDrawing1()
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                AbrirDocumentManager();

                DrawingHandler drawingHandler = new DrawingHandler();
                DrawingEnumerator drawingEnumerator = drawingHandler.GetDrawings();

                while (drawingEnumerator.MoveNext())
                {
                    Drawing drawing = drawingEnumerator.Current as Drawing;
                    if (drawing != null)
                    {
                        string drawingName = drawing.Name.ToLower();
                        string drawingTitle1 = drawing.Title1.ToLower();

                        if (drawingName.Contains("schedule") || drawingTitle1.Contains("schedule"))
                        {
                            MessageBox.Show($"Desenho encontrado: {drawing.Name}", "Desenho Encontrado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                           
                            SelectDrawing2(drawing);
                            drawingHandler.GetActiveDrawing();

                            // Executar a macro
                            TeklaFuncoes.CorrerMacro(@"C:\Shaymurtagh_Programs\TEKLA\FIRM_MURTAGH\MACROS\modeling\CopyToNewSheet.cs");
                            

                            break;
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Não foi possível conectar ao Tekla Structures.", "Erro de Conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// Abre o Document Manager.
        /// </summary>
        public static void AbrirDocumentManager()
        {
            try
            {
                Tekla.Structures.Model.Operations.Operation.RunMacro("DocumentManager");
                MessageBox.Show("Document Manager aberto com sucesso.", "Document Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao abrir o Document Manager: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// Mantém o desenho selecionado sem abri-lo.
        /// </summary>
        /// <param name="Desenho">O enumerador de desenhos que contém o desenho previamente selecionado.</param>
        public static void SelectDrawing2(Drawing Desenho)
        {
            // Apenas armazena a referência ao desenho sem abri-lo
            // NÃO chama SetActiveDrawing(desenho);
            // Você pode usar o desenho para modificações sem abrir
            MessageBox.Show($"Desenho selecionado____: {Desenho.Name}", "Desenho Selecionado", MessageBoxButtons.OK, MessageBoxIcon.Information);

            //aqui vamos colocar um codigo para selecionar o desenho em questao

            DrawingHandler drawingHandler = new DrawingHandler();
            DrawingEnumerator drawings = drawingHandler.GetDrawings();

            while (drawings.MoveNext())
            {
                Drawing desenho = drawings.Current;

                if (desenho != null)
                {
                    MessageBox.Show($"Desenhoxxxx {desenho.Name} foi selecionado.", "Desenho Selecionado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    // Executar a macro
                    TeklaFuncoes.CorrerMacro(@"C:\Shaymurtagh_Programs\TEKLA\FIRM_MURTAGH\MACROS\modeling\CopyToNewSheet.cs");
                    break;
                }


                MessageBox.Show($"Desenho {Desenho.Name} foi selecionado.", "Desenho Selecionado", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }

      

        /// <summary>
        /// Copia o desenho para uma nova folha.
        /// </summary>
        /// <param name="originalDrawing"></param>
        private static void CopyDrawingToNewSheet(Drawing originalDrawing)
        {
            MessageBox.Show("Copiando o desenho para uma nova folha.", "Cópia de Desenho", MessageBoxButtons.OK, MessageBoxIcon.Information);

            DrawingHandler drawingHandler = new DrawingHandler();

            // Criar um novo desenho
            GADrawing newDrawing = new GADrawing
            {
                Name = originalDrawing.Name + " - Copy",
                Title1 = originalDrawing.Title1,
                Title2 = originalDrawing.Title2,
                Title3 = originalDrawing.Title3,

            };

            if (newDrawing.Insert())
            {

                MessageBox.Show("Novo desenho criado com sucesso.", "Cópia de Desenho", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // Copiar o conteúdo do desenho original para o novo desenho
                DrawingObjectEnumerator drawingObjects = originalDrawing.GetSheet().GetAllObjects();
                while (drawingObjects.MoveNext())
                {
                    DrawingObject drawingObject = drawingObjects.Current as DrawingObject;
                    if (drawingObject != null)
                    {
                        MessageBox.Show("Desenho copiado para uma nova folha com sucesso.", "Cópia de Desenho", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        //drawingObject.CopyTo(newDrawing.GetSheet(), new TSG.Point(0, 0));

                        CopiarObjetosDesenho(originalDrawing, newDrawing);
                        CopiarVistasDesenho(originalDrawing, newDrawing);
                        break;
                    }
                }

                //MessageBox.Show("Desenho copiado para uma nova folha com sucesso.", "Cópia de Desenho", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Falha ao criar o novo desenho.", "Erro de Cópia", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        public static void CopiarObjetosDesenho(Drawing originalDrawing, Drawing newDrawing)
        {
            // Obter todos os objetos da folha do desenho original
            DrawingObjectEnumerator drawingObjects = originalDrawing.GetSheet().GetAllObjects();

            while (drawingObjects.MoveNext())
            {
                DrawingObject drawingObject = drawingObjects.Current as DrawingObject;
                if (drawingObject != null)
                {
                    // Aqui você pode verificar o tipo de objeto e recriá-lo no novo desenho
                    // Exemplo para objetos de linha:
                    if (drawingObject is Line originalLine)
                    {
                        Line newLine = new Line(newDrawing.GetSheet(),
                                                new TSG.Point(originalLine.StartPoint.X, originalLine.StartPoint.Y),
                                                new TSG.Point(originalLine.EndPoint.X, originalLine.EndPoint.Y));
                        newLine.Insert();
                    }
                    // Adicione condições semelhantes para outros tipos de objetos conforme necessário
                 
                }
            }

            MessageBox.Show("Desenho copiado para uma nova folha com sucesso.", "Cópia de Desenho", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static void CopiarVistasDesenho(Drawing originalDrawing, Drawing newDrawing)
        {
            MessageBox.Show("Copiando vistas do desenho original para o novo desenho.", "Cópia de Vistas", MessageBoxButtons.OK, MessageBoxIcon.Information);
            // Obter todas as vistas do desenho original
            DrawingObjectEnumerator viewObjects = originalDrawing.GetSheet().GetViews();

            while (viewObjects.MoveNext())
            {
                Tekla.Structures.Drawing.View originalView = viewObjects.Current as Tekla.Structures.Drawing.View;
                if (originalView != null)
                {
                   
                    TSG.Point origin = originalView.Origin;
                    TSG.AABB aABB = originalView.GetAxisAlignedBoundingBox();
                    TSG.CoordinateSystem coordinateSystem = new CoordinateSystem();


                    // Criar uma nova vista no novo desenho com as mesmas propriedades
                    //Tekla.Structures.Drawing.View newView = new Tekla.Structures.Drawing.View(newDrawing.GetSheet(), originalView.Origin, originalView.ViewType)
                    TSD.View newView = new TSD.View(newDrawing.GetSheet(), coordinateSystem, coordinateSystem, aABB)
                    {

                        Name = originalView.Name,
                        Origin = originalView.Origin,
                        //Scale = originalView.Scale,
                        //Rotation = originalView.Rotation,
                        //ViewDepthUp = originalView.ViewDepthUp,
                        //ViewDepthDown = originalView.ViewDepthDown

                        ViewCoordinateSystem = originalView.ViewCoordinateSystem,
                    };
                    newView.Attributes.Scale = originalView.Attributes.Scale;
                    newView.Width = originalView.Width;
                    newView.Height = originalView.Height;
                    newView.Origin = originalView.Origin;
                    //newView.SetUserProperty("ViewType", originalView.GetUserProperty("ViewType"));
                    newView.FrameOrigin = originalView.FrameOrigin;
                    newView.RestrictionBox = originalView.RestrictionBox;
                    newView.Attributes.LoadAttributes("View_with_Status");



                    if (newView.Insert())
                    {
                        // Copiar objetos dentro da vista original para a nova vista
                        DrawingObjectEnumerator viewObjectsInOriginalView = originalView.GetObjects();
                        while (viewObjectsInOriginalView.MoveNext())
                        {
                            DrawingObject viewObject = viewObjectsInOriginalView.Current as DrawingObject;
                            if (viewObject != null)
                            {
                                // Aqui você pode verificar o tipo de objeto e recriá-lo na nova vista
                                // Exemplo para objetos de linha dentro da vista:
                                if (viewObject is Line originalLine)
                                {
                                    Line newLine = new Line(newView,
                                                            new TSG.Point(originalLine.StartPoint.X, originalLine.StartPoint.Y),
                                                            new TSG.Point(originalLine.EndPoint.X, originalLine.EndPoint.Y));
                                    newLine.Insert();
                                }
                                // Adicione condições semelhantes para outros tipos de objetos conforme necessário
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Cria um novo desenho GA (General Arrangement).
        /// </summary>
        public static void CriarNovoDesenho()
        {
            //Criar um novo desenho

            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                DrawingHandler drawingHandler = new DrawingHandler();

                // Definir as propriedades do novo desenho GA
                GADrawing newDrawing = new GADrawing
                {
                    Name = "Novo Desenho GA",
                    Title1 = "Título 1",
                    Title2 = "Título 2",
                    Title3 = "Título 3",
                };



                // Inserir o novo desenho no modelo
                if (newDrawing.Insert())
                {
                    MessageBox.Show("Novo desenho GA criado com sucesso.", "Criação de Desenho", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // Criar uma vista 3D e inserir no novo desenho
                    CriarVista3D(newDrawing);
                }
                else
                {
                    MessageBox.Show("Falha ao criar o novo desenho GA.", "Erro de Criação", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Não foi possível conectar ao Tekla Structures.", "Erro de Conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


        /// <summary>
        /// Cria uma vista 3D no modelo e insere no novo desenho GA.
        /// </summary>
        /// <param name="newDrawing">O novo desenho GA onde a vista será inserida.</param>
        private static void CriarVista3D(GADrawing newDrawing)
        {
            Model model = new Model();
            if (model.GetConnectionStatus())
            {
                // Definir o filtro para selecionar apenas peças de concreto
               /* BinaryFilterExpression filterExpression = new BinaryFilterExpression(
                    new ObjectFilterExpressions.CustomString("MATERIAL_TYPE"), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression("CONCRETE"));
               */

                // Criar a vista 3D
                TSD.View view = new TSD.View(newDrawing.GetSheet(), new TSG.CoordinateSystem(), new TSG.CoordinateSystem(), new TSG.AABB());

                // Definir a escala da vista
                view.Attributes.Scale = 250; // Defina a escala desejada aqui

                // Definir a área e a profundidade da vista
                view.Width = 500; // Defina a largura desejada aqui
                view.Height = 300; // Defina a altura desejada aqui
                //view.ViewDepthUp = 1000; // Defina a profundidade para cima desejada aqui
                //view.ViewDepthDown = 1000; // Defina a profundidade para baixo desejada aqui

                // Definir a origem da vista para posicioná-la no meio da folha
                view.Origin = new TSG.Point(150, 150, 0); // Defina a posição desejada aqui

                // Definir os limites da vista
                /*view.Attributes.XMin = 12345; // Defina o valor mínimo de X
                view.Attributes.XMax = 12345; // Defina o valor máximo de X
                view.Attributes.YMin = 12345; // Defina o valor mínimo de Y
                view.Attributes.YMax = 12345; // Defina o valor máximo de Y*/

                // Definir o sistema de coordenadas para uma vista 3D
                TSG.CoordinateSystem viewCoordinateSystem = new TSG.CoordinateSystem();
                viewCoordinateSystem.Origin = new TSG.Point(0, 0, 0);
                viewCoordinateSystem.AxisX = new TSG.Vector(1, 0, 0);
                viewCoordinateSystem.AxisY = new TSG.Vector(0, 1, 0);
                //viewCoordinateSystem.AxisZ = new TSG.Vector(0, 0, 1);
                view.ViewCoordinateSystem = viewCoordinateSystem;


                // Aplicar o filtro à vista
                //view.Filter = filterExpression;

                if (view.Insert())
                {
                    MessageBox.Show("Vista 3D criada e inserida no novo desenho GA com sucesso.", "Criação de Vista", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Falha ao criar a vista 3D.", "Erro de Criação", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Não foi possível conectar ao Tekla Structures.", "Erro de Conexão", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
