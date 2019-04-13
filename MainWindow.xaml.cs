using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using SolidWorks.Interop.cosworks;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using System.Collections;
using System.Threading;
using Excel = Microsoft.Office.Interop.Excel;
using System.Timers;

namespace W_Form_analyse_get_TrainingDaten
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private void Window_Closed(object sender, EventArgs e)
        {
            try
            {
                swApp.ExitApp();
            }
            catch (Exception)
            {

            }

        }
        #region Variable definition
        SldWorks swApp;
        int m;//how many cells
        int Count; // how many data
        double width, thickness, l1, l2, l3, F1, F2, F3;
        double l4, l1_, l2_, l3_, l4_, cell_length, cell_height, x_o, y_o, x_o_, y_o_;
        const double MeshEleSize = 2.0;
        const double MeshTol = 0.1;
        string strMaterialLib = null;
        object[] Disp = null;
        object[] Stress = null;
        ModelDoc2 swModel = null;
        SelectionMgr selectionMgr = null;
        CosmosWorks COSMOSWORKS = null;
        CwAddincallback COSMOSObject = default(CwAddincallback);
        CWModelDoc ActDoc = default(CWModelDoc);
        CWStudyManager StudyMngr = default(CWStudyManager);
        CWStudy Study = default(CWStudy);
        CWSolidManager SolidMgr = default(CWSolidManager);
        CWSolidBody SolidBody = default(CWSolidBody);
        CWSolidComponent SolidComp = default(CWSolidComponent);


        CWForce cwForce = default(CWForce);
        CWLoadsAndRestraintsManager LBCMgr = default(CWLoadsAndRestraintsManager);
        CWResults CWFeatobj = default(CWResults);
        bool isSelected;

        

        float maxDisp = 0.0f;
        float maxStress = 0.0f;
        int intStatus = 0;
        int errors = 0;
        int errCode = 0;
        int warnings = 0;


        /* change to double
        int ran_m, ran_width, ran_thickness, ran_l1, ran_l2, ran_l3, ran_F1, ran_F2, ran_F3;
        */
        int ran_m;
        double ran_width, ran_thickness, ran_l1, ran_l2, ran_l3, ran_F1, ran_F2, ran_F3;
        Excel.Application exlApp;
        Excel.Workbook exlBook;
        Excel.Worksheet exlSheet;



        #endregion
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_create_Geom_Click(object sender, RoutedEventArgs e)
        {
            Excel.Range r_m, r_width, r_thickness, r_l1, r_l2, r_l3, r_F1, r_F2, r_F3;

            Count = 849;

            for (int j = 777; j < Count; j++)
            {
                Console.WriteLine(string.Format("---------------------------------now No.{0}/{1}", j, Count));
                
                #region generate random Data

                r_m = (Excel.Range)exlSheet.Cells[j, 1];
                r_width = (Excel.Range)exlSheet.Cells[j, 2];
                r_thickness = (Excel.Range)exlSheet.Cells[j, 3];
                r_l1 = (Excel.Range)exlSheet.Cells[j, 4];
                r_l2 = (Excel.Range)exlSheet.Cells[j, 5];
                r_l3 = (Excel.Range)exlSheet.Cells[j, 6];
                r_F1 = (Excel.Range)exlSheet.Cells[j, 7];
                r_F2 = (Excel.Range)exlSheet.Cells[j, 8];
                r_F3 = (Excel.Range)exlSheet.Cells[j, 9];

                m = Convert.ToInt32(r_m.Value2);
                width = (double)(Convert.ToDouble(r_width.Value2) / 1000.0);
                thickness = (double)(Convert.ToDouble(r_thickness.Value2) / 1000.0);
                l1 = (double)(Convert.ToDouble(r_l1.Value2) / 1000.0);
                l2 = (double)(Convert.ToDouble(r_l2.Value2) / 1000.0);
                l3 = (double)(Convert.ToDouble(r_l3.Value2) / 1000.0);

                F1 = Convert.ToDouble(r_F1.Value2);
                F2 = Convert.ToDouble(r_F2.Value2);
                F3 = Convert.ToDouble(r_F3.Value2);

                l4 = l2;
                l1_ = l1 - 2 * width;
                l2_ = l2;
                l3_ = l3 + 2 * width;
                l4_ = l4;
                cell_length = l1 + l3;
                cell_height = l2 + width;


                #endregion
                


                Console.WriteLine(string.Format("{0},{1},{2},{3},{4},{5}", m, width, thickness, l1, l2, l3));

                #region geometrie
                Console.WriteLine("Start creating new Germetrie");
                try
                {
                    swModel = swApp.NewPart();
                }
                catch (Exception)
                {
                    
                    Console.WriteLine("starting SolidWorks");
                    try
                    {
                        swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    }
                    catch (Exception)
                    {
                        swApp = new SldWorks();
                        swApp.Visible = false;
                    }

                    //get MaterialLib
                    strMaterialLib = swApp.GetExecutablePath() + "\\lang\\english\\sldmaterials\\solidworks materials.sldmat";
                    //strMaterialLib = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\english\\sldmaterials\\solidworks materials.sldmat";
                    Console.WriteLine("SolidWorks successfully started");
                    j--;
                    continue;


                }
                swModel.Extension.SelectByID("前视基准面", "PLANE", 0, 0, 0, false, 1, null);
                swModel.InsertSketch2(true);
                #region sketch
                for (int i = 0; i < m; i++)
                {
                    if (i == 0)
                    {
                        swModel.SketchManager.CreateLine(0, 0, 0, l1 - width, 0, 0);
                        swModel.SketchManager.CreateLine(l1 - width, 0, 0, l1 - width, -l2, 0);
                        swModel.SketchManager.CreateLine(l1 - width, -l2, 0, l1 + l3 - width, -l2, 0);
                        swModel.SketchManager.CreateLine(l1 + l3 - width, -l2, 0, l1 + l3 - width, 0, 0);

                        swModel.SketchManager.CreateLine(0, 0, 0, 0, -width, 0);
                        swModel.SketchManager.CreateLine(0, -width, 0, l1_, -width, 0);
                        swModel.SketchManager.CreateLine(l1_, -width, 0, l1_, -(width + l2_), 0);
                        swModel.SketchManager.CreateLine(l1_, -(width + l2_), 0, l1_ + l3_, -(width + l2_), 0);
                        swModel.SketchManager.CreateLine(l1_ + l3_, -(width + l2_), 0, l1_ + l3_, -width, 0);
                    }
                    else
                    {
                        x_o = i * cell_length - width;
                        y_o = 0;
                        x_o_ = x_o + width;
                        y_o_ = y_o - width;

                        swModel.SketchManager.CreateLine(x_o, y_o, 0, x_o + l1, 0, 0);
                        swModel.SketchManager.CreateLine(x_o + l1, 0, 0, x_o + l1, -l2, 0);
                        swModel.SketchManager.CreateLine(x_o + l1, -l2, 0, x_o + l1 + l3, -l2, 0);
                        swModel.SketchManager.CreateLine(x_o + l1 + l3, -l2, 0, x_o + l1 + l3, 0, 0);

                        swModel.SketchManager.CreateLine(x_o_, y_o_, 0, x_o_ + l1_, y_o_, 0);
                        swModel.SketchManager.CreateLine(x_o_ + l1_, y_o_, 0, x_o_ + l1_, y_o_ - l2_, 0);
                        swModel.SketchManager.CreateLine(x_o_ + l1_, y_o_ - l2_, 0, x_o_ + l1_ + l3_, y_o_ - l2_, 0);
                        swModel.SketchManager.CreateLine(x_o_ + l1_ + l3_, y_o_ - l2_, 0, x_o_ + l1_ + l3_, y_o_, 0);
                    }
                }
                swModel.SketchManager.CreateLine(x_o_ + l1_ + l3_, y_o_, 0, x_o_ + l1_ + l3_, 0, 0);
                swModel.SketchManager.CreateLine(x_o_ + l1_ + l3_, 0, 0, x_o + l1 + l3, 0, 0);
               
                #endregion
                swModel.FeatureManager.FeatureExtrusion2(
                        true, false, false, 0, 0, thickness, 0, false, false, false, false, 0, 0, false, false, false, false, true, true, true, 0, 0, true
                        );

                swModel.SaveAsSilent("D:\\TUD\\7.Semeter\\SA\\SA_code\\c#\\W_Form_analyse_get_TrainingDaten\\Geometrie.sldprt", true);
                swApp.CloseAllDocuments(true);
                swApp.OpenDoc("D:\\TUD\\7.Semeter\\SA\\SA_code\\c#\\W_Form_analyse_get_TrainingDaten\\Geometrie.sldprt", (int)swOpenDocOptions_e.swOpenDocOptions_Silent);
                Console.WriteLine("Geometrie success");

  
                #endregion

                #region simulaiton
                string path_to_cosworks_dll = @"C:\Program Files\SOLIDWORKS Corp\SOLIDWORKS\Simulation\cosworks.dll";
                errors = swApp.LoadAddIn(path_to_cosworks_dll);
                COSMOSObject = (CwAddincallback)swApp.GetAddInObject("SldWorks.Simulation");
                try
                {
                    COSMOSWORKS = (CosmosWorks)COSMOSObject.CosmosWorks;
                }
                catch (Exception)
                {
                    Console.WriteLine("something wrong in Simulaiton Add In, start a new one");
                    continue;
                }


                COSMOSWORKS = COSMOSObject.CosmosWorks;
                //Get active document
                ActDoc = (CWModelDoc)COSMOSWORKS.ActiveDoc;

                //Create new static study
                try
                {
                    StudyMngr = (CWStudyManager)ActDoc.StudyManager;
                    Study = (CWStudy)StudyMngr.CreateNewStudy("static study", (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeStatic, 0, out errCode);
                }
                catch (Exception)
                {
                    errors = swApp.UnloadAddIn(path_to_cosworks_dll);
                    swApp.CloseAllDocuments(true);
                    continue;

                }

                //Add materials
                try
                {
                    SolidMgr = Study.SolidManager;
                }
                catch (Exception)
                {
                    Console.WriteLine("starting SolidWorks");
                    try
                    {
                        swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
                    }
                    catch (Exception)
                    {
                        swApp = new SldWorks();
                        swApp.Visible = false;
                    }

                    //get MaterialLib
                    strMaterialLib = swApp.GetExecutablePath() + "\\lang\\english\\sldmaterials\\solidworks materials.sldmat";
                    //strMaterialLib = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\english\\sldmaterials\\solidworks materials.sldmat";
                    Console.WriteLine("SolidWorks successfully started");
                    j--;
                    continue;
                }
                SolidComp = SolidMgr.GetComponentAt(0, out errCode);
                SolidBody = SolidComp.GetSolidBodyAt(0, out errCode);
                intStatus = SolidBody.SetLibraryMaterial(strMaterialLib, "AISI 1020");

                //fixed restraints
                LBCMgr = Study.LoadsAndRestraintsManager;

                swModel = (ModelDoc2)swApp.ActiveDoc;
                swModel.ShowNamedView2("", (int)swStandardViews_e.swIsometricView);

                selectionMgr = (SelectionMgr)swModel.SelectionManager;
                isSelected = swModel.Extension.SelectByID2("", "FACE", 0, -width / 2.0, thickness / 2.0, false, 0, null, 0);
                if (isSelected)
                {
                    object selectedFace = (object)selectionMgr.GetSelectedObject6(1, -1);
                    object[] fixedFaces = { selectedFace };
                    CWRestraint restraint = (CWRestraint)LBCMgr.AddRestraint((int)swsRestraintType_e.swsRestraintTypeFixed, fixedFaces, null, out errCode);
                }
                swModel.ClearSelection2(true);

                //add force
                selectionMgr = (SelectionMgr)swModel.SelectionManager;
                isSelected = swModel.Extension.SelectByID2("", "FACE", x_o + l1 + l3 + (width / 2.0), 0, thickness / 2.0, false, 0, null, 0);
                if (isSelected)
                {
                    object selectedFace = (object)selectionMgr.GetSelectedObject6(1, -1);
                    object[] forceAdd = { selectedFace };
                    selectionMgr = (SelectionMgr)swModel.SelectionManager;
                    swModel.Extension.SelectByID2("", "FACE", x_o + l1 + l3 + width, -(l4 + width) / 2.0, thickness / 2.0, false, 0, null, 0);
                    object selectedFaceToForceDir = (object)selectionMgr.GetSelectedObject6(1, -1);
                    double[] distValue = null;
                    double[] forceValue = null;
                    double[] Force = { F2, F3, F1 };
                    cwForce = (CWForce)LBCMgr.AddForce3((int)swsForceType_e.swsForceTypeForceOrMoment, (int)swsSelectionType_e.swsSelectionFaceEdgeVertexPoint,
                                                        2, 0, 0, 0,
                                                        (distValue),
                                                        (forceValue),
                                                        false, true,
                                                        (int)swsBeamNonUniformLoadDef_e.swsTotalLoad,
                                                        0, 7, 0.0,
                                                        Force,
                                                        false, false,
                                                        (forceAdd),
                                                        (selectedFaceToForceDir),
                                                        false, out errCode);//i have tried to figure out these arguments for one day, keep them and dont't change them.
                                                                            //the way to check cwForce : cwForce.GetForceComponentValues()
                                                                            //ForceComponet: [int b1, // 1 if x-direction hat Force-komponent, else 0
                                                                            //                int b2, // 1 if y-direction hat Force-komponent, else 0
                                                                            //                int b3, // 1 if z-direction hat Force-komponent, else 0
                                                                            //                double d1, // Force-komponent in x
                                                                            //                double d2, // Force-komponent in y
                                                                            //                double d3 // Force-komponent in z
                                                                            //                          ]
                                                                            //PS: the definition of xyz seems like not the same as the global XYZ system in SW.
                    swModel.ClearSelection2(true);

                    //meshing
                    CWMesh CWMeshObj = default(CWMesh);
                    CWMeshObj = Study.Mesh;
                    CWMeshObj.MesherType = (int)swsMesherType_e.swsMesherTypeStandard;
                    CWMeshObj.Quality = (int)swsMeshQuality_e.swsMeshQualityDraft;
                    errCode = Study.CreateMesh(0, MeshEleSize, MeshTol);
                    CWMeshObj = null;

                    //run analysis
                    errCode = Study.RunAnalysis();
                    if (errCode != 0)
                    {
                        Console.WriteLine(string.Format("RunAnalysis errCode = {0}", errCode));
                        Console.WriteLine(string.Format("RunAnalysis failed"));

                        

                        swApp.CloseAllDocuments(true);
                        errors = swApp.UnloadAddIn(path_to_cosworks_dll);
                        Console.WriteLine(string.Format("ready to start a new one"));
                        continue;
                    }
                    Console.WriteLine("RunAnalysis successed, ready to get results");

                    //get results
                    CWFeatobj = Study.Results;
                    //get max von Mieses stress
                    Stress = (object[])CWFeatobj.GetMinMaxStress((int)swsStressComponent_e.swsStressComponentVON,
                                                                 0, 0, null,
                                                                 (int)swsStrengthUnit_e.swsStrengthUnitNewtonPerSquareMillimeter,
                                                                 out errCode);
                    maxStress = (float)Stress[3]; //Stress: {node_with_minimum_stress, minimum_stress, node_with_maximum_stress, maximum_stress}
                    Console.WriteLine(maxStress);
                    /*
                    if (maxStress >= 351.6)
                    {
                        Console.WriteLine("out of yield stress, start a new example");
                        errors = swApp.UnloadAddIn(path_to_cosworks_dll);
                        swApp.CloseAllDocuments(true);
                        j--;
                        continue;
                    }
                    */
                    //get max URES displacement
                    Disp = (object[])CWFeatobj.GetMinMaxDisplacement((int)swsDisplacementComponent_e.swsDisplacementComponentURES,
                                                                     0, null,
                                                                     (int)swsLinearUnit_e.swsLinearUnitMillimeters,
                                                                     out errCode);
                    maxDisp = (float)Disp[3]; //Disp: {node_with_minimum_displacement, minimum_displacement, node_with_maximum_displacement, maximum_displacement}
                    CWFeatobj = null;
                    Console.WriteLine(string.Format("max Displacement: {0:f4} mm", maxDisp));

                   
                    //output to Excel
                    exlSheet.Cells[j, 1] = m;
                    exlSheet.Cells[j, 2] = width * 1000;
                    exlSheet.Cells[j, 3] = thickness * 1000;
                    exlSheet.Cells[j, 4] = l1 * 1000;
                    exlSheet.Cells[j, 5] = l2 * 1000;
                    exlSheet.Cells[j, 6] = l3 * 1000;
                    exlSheet.Cells[j, 7] = F1;
                    exlSheet.Cells[j, 8] = F2;
                    exlSheet.Cells[j, 9] = F3;
                    exlSheet.Cells[j, 10] = maxStress;
                    exlSheet.Cells[j, 11] = maxDisp;
                    if (j % 5 == 0)
                    {
                        exlBook.Save();
                    }

                    errors = swApp.UnloadAddIn(path_to_cosworks_dll);
                    swApp.CloseAllDocuments(true);

                }
                #endregion
            }
            exlBook.Save();
            exlApp.Quit();
        }
        private void button_start_SW_Click(object sender, RoutedEventArgs e)
        {
            Console.WriteLine("starting SolidWorks");
            try
            {
                swApp = (SldWorks)Marshal.GetActiveObject("SldWorks.Application");
            }
            catch (Exception)
            {
                swApp = new SldWorks();
                swApp.Visible = false;
            }

            //get MaterialLib
            strMaterialLib = swApp.GetExecutablePath() + "\\lang\\english\\sldmaterials\\solidworks materials.sldmat";
            //strMaterialLib = "C:\\Program Files\\SOLIDWORKS Corp\\SOLIDWORKS\\lang\\english\\sldmaterials\\solidworks materials.sldmat";
            Console.WriteLine("SolidWorks successfully started");
        }

        private void button_start_Excel_Click(object sender, RoutedEventArgs e)
        {
            exlApp = new Microsoft.Office.Interop.Excel.Application();
            exlApp.Visible = true;
            exlBook = exlApp.Workbooks.Open("D:\\TUD\\7.Semeter\\SA\\SA_code\\c#\\W_Form_analyse_get_TrainingDaten\\W_Form_simulationDaten_1553758068644_clean.xlsx");
            exlSheet = exlBook.ActiveSheet;

        }

        #region help function
        public static void ErrorMsg(SldWorks SwApp, string Message)
        {
            SwApp.SendMsgToUser2(Message, 0, 0);
            SwApp.RecordLine("'*** WARNING - General");
            SwApp.RecordLine("'*** " + Message);
            SwApp.RecordLine("");
        }

        /// <summary>
        /// 获得13位的时间戳
        /// </summary>
        /// <returns></returns>
        public static string GetTimeStamp()
        {
            System.DateTime time = System.DateTime.Now;
            long ts = ConvertDateTimeToInt(time);
            return ts.ToString();
        }
        /// <summary>  
        /// 将c# DateTime时间格式转换为Unix时间戳格式  
        /// </summary>  
        /// <param name="time">时间</param>  
        /// <returns>long</returns>  
        private static long ConvertDateTimeToInt(System.DateTime time)
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1, 0, 0, 0, 0));
            long t = (time.Ticks - startTime.Ticks) / 10000;   //除10000调整为13位      
            return t;
        }
        #endregion
    }
}
