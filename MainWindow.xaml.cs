﻿using System;
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
            #region a test example, not important
            m = 9;
            width = 8 / 1000.0;
            thickness = 1 / 1000.0;
            l1 = 30 / 1000.0;
            l2 = 16 / 1000.0;
            l3 = 21 / 1000.0;
            F1 = 1;
            F2 = 37;
            F3 = 49;

            l4 = l2;
            l1_ = l1 - 2 * width;
            l2_ = l2;
            l3_ = l3 + 2 * width;
            l4_ = l4;
            cell_length = l1 + l3;
            cell_height = l2 + width;
            #endregion

            Count = int.Parse(textBox_count.Text);
            Random randomSeed = new Random();
            Random random_m = new Random(randomSeed.Next());
            Random random_width = new Random(randomSeed.Next());
            Random random_thickness = new Random(randomSeed.Next());
            Random random_l1 = new Random(randomSeed.Next());
            Random random_l2 = new Random(randomSeed.Next());
            Random random_l3 = new Random(randomSeed.Next());
            Random random_F1 = new Random(randomSeed.Next());
            Random random_F2 = new Random(randomSeed.Next());
            Random random_F3 = new Random(randomSeed.Next());

            for (int j = 0; j < Count; j++)
            {
                Console.WriteLine(string.Format("---------------------------------now No.{0}", j));
                
                #region generate random Data
                try
                {
                    /* change to double
                    ran_m = random_m.Next(int.Parse(textBox_hm_cells_min.Text), int.Parse(textBox_hm_cellls_max.Text));
                    ran_width = random_width.Next(int.Parse(textBox_width_min.Text), int.Parse(textBox_width_max.Text));
                    ran_thickness = random_thickness.Next(int.Parse(textBox_thickness_min.Text), int.Parse(textBox_thickness_max.Text));
                    ran_l1 = random_l1.Next(int.Parse(textBox_l1_min.Text), int.Parse(textBox_l1_max.Text));
                    ran_l2 = random_l2.Next(int.Parse(textBox_l2_min.Text), int.Parse(textBox_l2_max.Text));
                    ran_l3 = random_l3.Next(int.Parse(textBox_l3_min.Text), int.Parse(textBox_l3_max.Text));
                    ran_F1 = random_F1.Next(int.Parse(textBox_F1_min.Text), int.Parse(textBox_F1_max.Text));
                    ran_F2 = random_F2.Next(int.Parse(textBox_F2_min.Text), int.Parse(textBox_F2_max.Text));
                    ran_F3 = random_F3.Next(int.Parse(textBox_F3_min.Text), int.Parse(textBox_F3_max.Text));
                    */
                    ran_m = random_m.Next(int.Parse(textBox_hm_cells_min.Text), int.Parse(textBox_hm_cellls_max.Text));
                    ran_width = random_width.NextDouble() * (int.Parse(textBox_width_max.Text) - int.Parse(textBox_width_min.Text)) + double.Parse(textBox_width_min.Text);
                    ran_thickness = random_thickness.NextDouble() * (int.Parse(textBox_thickness_max.Text) - int.Parse(textBox_thickness_min.Text)) + double.Parse(textBox_thickness_min.Text);
                    ran_l1 = random_l1.NextDouble() * (int.Parse(textBox_l1_max.Text) - int.Parse(textBox_l1_min.Text)) + double.Parse(textBox_l1_min.Text);
                    ran_l2 = random_l2.NextDouble() * (int.Parse(textBox_l2_max.Text) - int.Parse(textBox_l2_min.Text)) + double.Parse(textBox_l2_min.Text);
                    ran_l3 = random_l3.NextDouble() * (int.Parse(textBox_l3_max.Text) - int.Parse(textBox_l3_min.Text)) + double.Parse(textBox_l3_min.Text);
                    ran_F1 = random_F1.NextDouble() * (int.Parse(textBox_F1_max.Text) - int.Parse(textBox_F1_min.Text)) + double.Parse(textBox_F1_min.Text);
                    ran_F2 = random_F2.NextDouble() * (int.Parse(textBox_F2_max.Text) - int.Parse(textBox_F2_min.Text)) + double.Parse(textBox_F2_min.Text);
                    ran_F3 = random_F3.NextDouble() * (int.Parse(textBox_F3_max.Text) - int.Parse(textBox_F3_min.Text)) + double.Parse(textBox_F3_min.Text);

                }
                catch (Exception)
                {
                    Console.WriteLine("Input Format Error");
                    return;
                }


                m = ran_m;
                width = (double)(ran_width / 1000.0);
                thickness = (double)(ran_thickness / 1000.0);
                l1 = (double)(ran_l1 / 1000.0);
                l2 = (double)(ran_l2 / 1000.0);
                l3 = (double)(ran_l3 / 1000.0);

                F1 = (double)ran_F1;
                F2 = (double)ran_F2;
                F3 = (double)ran_F3;
                F3 = 0;

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
                StudyMngr = (CWStudyManager)ActDoc.StudyManager;
                Study = (CWStudy)StudyMngr.CreateNewStudy("static study", (int)swsAnalysisStudyType_e.swsAnalysisStudyTypeStatic, 0, out errCode);

                //Add materials
                SolidMgr = Study.SolidManager;
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
                    exlSheet.Cells[j + 2, 1] = m;
                    exlSheet.Cells[j + 2, 2] = width * 1000;
                    exlSheet.Cells[j + 2, 3] = thickness * 1000;
                    exlSheet.Cells[j + 2, 4] = l1 * 1000;
                    exlSheet.Cells[j + 2, 5] = l2 * 1000;
                    exlSheet.Cells[j + 2, 6] = l3 * 1000;
                    exlSheet.Cells[j + 2, 7] = F1;
                    exlSheet.Cells[j + 2, 8] = F2;
                    exlSheet.Cells[j + 2, 9] = F3;
                    exlSheet.Cells[j + 2, 10] = maxStress;
                    exlSheet.Cells[j + 2, 11] = maxDisp;
                    if (j % 5 == 0)
                    {
                        exlBook.Save();//C:\Users\Zhao\Documents\工作簿1.xlsx
                    }

                    errors = swApp.UnloadAddIn(path_to_cosworks_dll);
                    swApp.CloseAllDocuments(true);

                }
                #endregion
            }
            exlBook.SaveCopyAs("D:\\TUD\\7.Semeter\\SA\\SA_code\\c#\\W_Form_analyse_get_TrainingDaten\\W_Form_simulationDaten_" + GetTimeStamp() +"_F3=0" + ".xlsx");
            exlBook.Save();//C:\Users\zhaojie\Documents\工作簿1.xlsx
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

            Console.WriteLine("SolidWorks successfully started");
        }

        private void button_start_Excel_Click(object sender, RoutedEventArgs e)
        {
            exlApp = new Microsoft.Office.Interop.Excel.Application();
            exlApp.Visible = false;
            exlBook = exlApp.Workbooks.Add();
            exlSheet = exlBook.ActiveSheet;
            exlSheet.Cells[1, 1] = "cells number";
            exlSheet.Cells[1, 2] = "width(mm)";
            exlSheet.Cells[1, 3] = "thickness(mm)";
            exlSheet.Cells[1, 4] = "l1(mm)";
            exlSheet.Cells[1, 5] = "l2(mm)";
            exlSheet.Cells[1, 6] = "l3(mm)";
            exlSheet.Cells[1, 7] = "F1(N)";
            exlSheet.Cells[1, 8] = "F2(N)";
            exlSheet.Cells[1, 9] = "F3(N)";
            exlSheet.Cells[1, 10] = "maxStress(MPa)";
            exlSheet.Cells[1, 11] = "maxDisp(mm)";
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
