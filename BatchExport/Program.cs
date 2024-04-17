using System;
using System.IO;
using System.Collections.Generic;
using System.Windows.Forms;

using NXOpen;
using NXOpen.UF;
using NXOpen.Assemblies;
using NXOpen.Drawings;

namespace BatchExport
{
    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                #region Argument order
                /*
                 * "C:\Program Files\Siemens\NX 12.0\NXBIN\run_managed.exe" 
                 * "C:\Users\MOHAN DULAM\Documents\Visual Studio 2019\NX_Open\LN-Projects\Exe Files\StepFileCreator\StepFileCreator\bin\Debug\BatchExport.exe"
                 * "C:\Users\MOHAN DULAM\Documents\NX Models\Batch Process\Drawing\Drawing Labels Export.prt" 
                 * "PDF" 
                 * "C:\Users\MOHAN DULAM\Documents\NX Models\Batch Process\Drawing\Export File"
                 * 
                 * 3 Arumnents are 
                 * 1. Part file with file path(Directory)
                 * 2. Export file Extention are "STEP" "IGES" "XT" "DWG" "DXF" "PDF"
                 * 3. Location(Directory) for Exported file to save 
                 *                  
                 */
                #endregion

                // check for Arguments length
                if (args.Length != 3)
                {
                    //Console.WriteLine("No part(s) specified. Exit.");
                    MessageBox.Show("Input Arguments are not Found");
                    return;
                }

                // Input Arguments 
                string firstArgPartFile = args[0];
                string exportFileExtension = args[1];
                string exportFileDirectory = args[2];
                //Console.WriteLine("1 Arg: " + firstArgPartFile);
                //Console.WriteLine("2 Arg: " + exportFileExtension);
                //Console.WriteLine("3 Arg: " + exportFileDirectory);

                #region Check for File Extension
                if (exportFileExtension == "STEP")
                {
                    BatchExportStep batchExportStep = new BatchExportStep();
                    batchExportStep.CreateStepFile(firstArgPartFile, exportFileDirectory);
                }
                else if (exportFileExtension == "IGES")
                {
                    BatchExportIGES batchExportIGES = new BatchExportIGES();
                    batchExportIGES.CreateIGESFile(firstArgPartFile, exportFileDirectory);
                }
                else if (exportFileExtension == "XT")
                {
                    BatchExportParasolid batchExportParasolid = new BatchExportParasolid();
                    batchExportParasolid.CreateParasolidFile(firstArgPartFile, exportFileDirectory);
                }
                else if (exportFileExtension == "DWG")
                {
                    BatchExportDrawing batchExportDwg = new BatchExportDrawing();
                    batchExportDwg.CreateDwgDxfFile(firstArgPartFile, exportFileExtension, exportFileDirectory);
                }
                else if (exportFileExtension == "DXF")
                {
                    BatchExportDrawing batchExportDxf = new BatchExportDrawing();
                    batchExportDxf.CreateDwgDxfFile(firstArgPartFile, exportFileExtension, exportFileDirectory);
                }
                else if (exportFileExtension == "PDF")
                {
                    BatchExportPDF batchExportPDF = new BatchExportPDF();
                    batchExportPDF.CreatePdfFile(firstArgPartFile, exportFileExtension, exportFileDirectory);
                }

                #endregion

                #region Check for File Extension Switch Statement
                //switch (exportFileExtension)
                //{
                //    case "STEP":
                //        BatchExportStep batchExportStep = new BatchExportStep();
                //        batchExportStep.CreateStepFile(firstArgPartFile, exportFileDirectory);
                //        break;

                //    case "IGES":
                //        BatchExportIGES batchExportIGES = new BatchExportIGES();
                //        batchExportIGES.CreateIGESFile(firstArgPartFile, exportFileDirectory);
                //        break;

                //    case "XT":
                //        BatchExportParasolid batchExportParasolid = new BatchExportParasolid();
                //        batchExportParasolid.CreateParasolidFile(firstArgPartFile, exportFileDirectory);
                //        break;

                //    case "DWG":
                //        BatchExportDrawing batchExportDwg = new BatchExportDrawing();
                //        batchExportDwg.CreateDwgDxfFile(firstArgPartFile, exportFileExtension, exportFileDirectory);
                //        break;

                //    case "DXF":
                //        BatchExportDrawing batchExportDxf = new BatchExportDrawing();
                //        batchExportDxf.CreateDwgDxfFile(firstArgPartFile, exportFileExtension, exportFileDirectory);
                //        break;

                //    case "PDF":
                //        BatchExportPDF batchExportPDF = new BatchExportPDF();
                //        batchExportPDF.CreatePdfFile(firstArgPartFile, exportFileExtension, exportFileDirectory);
                //        break;

                //    default:
                //        break;
                //}
                #endregion

            }
            catch (Exception ex)
            {
                //throw;
            }
            
        }
        public static int GetUnloadOption(string dummy)
        {
            return (int)NXOpen.Session.LibraryUnloadOption.Immediately;
        }
    }

    /// <summary>
    /// Batch Export of STEP File from given Part file
    /// </summary>
    public class BatchExportStep
    {
        // class Memebers
        // Get the NX Session
        private static Session theSession = Session.GetSession();

        /// <summary>
        /// Batch Export of the STEP File
        /// </summary>
        /// <param name="partFilePath">First Argument as Part File Directory</param>
        /// <param name="exportDirectory">Directory of Exported File to Save</param>
        public void CreateStepFile(string partFilePath,string exportDirectory)
        {
            try
            {
                if (partFilePath == null && exportDirectory == null)
                {
                    //Console.WriteLine("STEP File Export input parameter are not Proper");
                    MessageBox.Show("STEP File Export input parameter are not Proper");
                    return;
                }

                // Check the Work Part load Status
                PartLoadStatus loadStatus;
                Part workPart = theSession.Parts.OpenDisplay(partFilePath, out loadStatus);
                if (workPart == null)
                {
                    //Console.WriteLine("Part file is null Return");
                    MessageBox.Show("Part file is not Found");
                    return;
                }

                // Calling StepFileExport function
                StepFileExport(workPart, exportDirectory);

                // Close the open Work Part
                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

            }
            catch (Exception ex)
            {
                //throw;
            }    
        }

        /// <summary>
        /// Export Step File form Part File
        /// </summary>
        /// <param name="part">Part File need to Export to STEP File</param>
        /// <param name="stepFileDir">Directory where STEP File to Save</param>
        private void StepFileExport(Part part, string stepFileDir)
        {
            try
            {
                // Name of Step file to save
                string displayPartName = part.Name;

                // location of the file
                string inputPartFileDir = part.FullPath;
                //string partFileDir = inputPartFileDir.Substring(0, inputPartFileDir.LastIndexOf("\\")) + "\\";
                //Console.WriteLine("Input File: "+inputPartFileDir);

                // File Path and Name of the Step file where to save it
                string outputStepFileDir = stepFileDir + $"\\" + displayPartName;
                //Console.WriteLine("Output File: "+outputStepFileDir);

                // Check for any old step file is exist in given directory 
                string delOldStepFile = outputStepFileDir + ".stp";
                if (File.Exists(delOldStepFile))
                {
                    File.Delete(delOldStepFile);
                }

                // Declaration of Step file Definition in NX
                string stepSettingFile;
                // Location of the Step file Definition in NX 
                string STEP214UG_Dir = theSession.GetEnvironmentVariableValue("STEP214UG_DIR");
                // Step file Definition File
                stepSettingFile = Path.Combine(STEP214UG_Dir, "ugstep214.def");
                //stepSettingFile = @"C:\Program Files\Siemens\NX 12.0\STEP214UG\ugstep214.def";

                #region STEP Creator Builder
                // Step File Creator Builder 
                NXOpen.StepCreator step214Creator1;
                step214Creator1 = theSession.DexManager.CreateStepCreator();
                step214Creator1.ExportAs = NXOpen.StepCreator.ExportAsOption.Ap214;
                step214Creator1.ObjectTypes.Solids = true;
                step214Creator1.ObjectTypes.Surfaces = true;
                step214Creator1.ObjectTypes.Curves = true;
                // File Path and Name of the Step file where to save it
                step214Creator1.OutputFile = outputStepFileDir;
                // Input File Path
                step214Creator1.InputFile = inputPartFileDir;

                // Path of the NX STEP file Definitation
                step214Creator1.SettingsFile = stepSettingFile;
                step214Creator1.ExportExtRef = false;
                step214Creator1.FileSaveFlag = false;
                step214Creator1.LayerMask = "1-256";
                step214Creator1.ProcessHoldFlag = true;

                // Commit the Builder
                NXObject nXObject;
                nXObject = step214Creator1.Commit();

                // Destroy the STEP File Creator Builder
                step214Creator1.Destroy();

                #endregion

                // Delete Log file
                string delLogFile = outputStepFileDir + ".log";
                if (File.Exists(delLogFile))
                {
                    File.Delete(delLogFile);
                }

            }
            catch (Exception ex)
            {
                string message = ex.ToString();
                // Message.Echo("Error While Exporting Part File" + ex.ToString());
            }
        }
    }
    
    /// <summary>
    /// Batch Export of IGES File from given Part file
    /// </summary>
    public class BatchExportIGES
    {
        // class Memebers
        // Get the NX Session
        private static Session theSession = Session.GetSession();

        /// <summary>
        /// Batch Export of the IGES File
        /// </summary>
        /// <param name="partFilePath">First Argument as Part File Directory</param>
        /// <param name="exportDirectory">Directory of Exported File to Save</param>
        public void CreateIGESFile(string partFilePath, string exportDirectory)
        {
            try
            {
                if (partFilePath == null && exportDirectory == null)
                {
                    //Console.WriteLine("IGES File Export input parameter are not Proper");
                    MessageBox.Show("IGES File Export input parameter are not Proper");
                    return;
                }

                // Check the Work Part load Status
                PartLoadStatus loadStatus;
                Part workPart = theSession.Parts.OpenDisplay(partFilePath, out loadStatus);
                if (workPart == null)
                {
                    //Console.WriteLine("Part file is null Return");
                    MessageBox.Show("Part file is not Found");
                    return;
                }

                // Calling IGESFileCreator function
                IGESFileCreator(workPart, exportDirectory);

                // Close the open Work Part
                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

            }
            catch (Exception ex)
            {
                //throw;
            }
        }

        /// <summary>
        /// Create Part File as IGES File
        /// </summary>
        /// <param name="part">Part File need to Export to IGES File</param>
        /// <param name = "igesFileDir" > Directory where IGES File to Save</param>
        private  void IGESFileCreator(Part part, string igesFileDir)
        {
            // Name of Step file to save
            string displayPartName = part.Name;

            // location of the file
            string inputPartFileDir = part.FullPath;
            //string partFileDir = inputPartFileDir.Substring(0, inputPartFileDir.LastIndexOf("\\")) + "\\";
            //Console.WriteLine("Input File: "+inputPartFileDir);

            // File Path and Name of the Iges file where to save it
            string outputIgesFileDir = igesFileDir + $"\\" + displayPartName;
            //Console.WriteLine("Output File: "+outputIgesFileDir);

            // Chack for any old iges file is exist in given directory 
            string delOldIgesFile = outputIgesFileDir + ".igs";
            if (File.Exists(delOldIgesFile))
            {
                File.Delete(delOldIgesFile);
            }

            string igesSettingFile; // Declaration of IGES file Definition in NX
            // Location of the IGES file Definition in NX
            string IGES_DIR = theSession.GetEnvironmentVariableValue("IGES_DIR");
            igesSettingFile = Path.Combine(IGES_DIR, "igesexport.def");

            #region IGES Creator Builder
            //  IGES Creator Builder
            NXOpen.IgesCreator igesCreator3;

            igesCreator3 = theSession.DexManager.CreateIgesCreator();
            // IGES file Definition File
            igesCreator3.SettingsFile = igesSettingFile;
            //igesCreator3.SettingsFile = "C:\\Program Files\\Siemens\\NX 12.0\\iges\\igesexport.def";

            // File Path and Name of the Step file where to save it
            igesCreator3.OutputFile = outputIgesFileDir + ".igs"; //partFilePath + displayName + ".igs";
            // Input File Path
            igesCreator3.InputFile = inputPartFileDir;

            igesCreator3.ExportModelData = true;
            igesCreator3.ExportDrawings = true;
            igesCreator3.MapTabCylToBSurf = true;
            igesCreator3.BcurveTol = 0.050799999999999998;
            igesCreator3.IdenticalPointResolution = 0.001;
            igesCreator3.MaxThreeDMdlSpace = 10000.0;
            igesCreator3.ObjectTypes.Curves = true;
            igesCreator3.ObjectTypes.Surfaces = true;
            igesCreator3.ObjectTypes.Annotations = true;
            igesCreator3.ObjectTypes.Structures = true;
            igesCreator3.ObjectTypes.Solids = true;
            igesCreator3.ExportDrawings = false;
            igesCreator3.FlattenAssembly = true;
            igesCreator3.MapRevolvedFacesTo = NXOpen.IgesCreator.MapRevolvedFacesOption.BSurfaces;
            igesCreator3.MapCrossHatchTo = NXOpen.IgesCreator.CrossHatchMapEnum.SectionArea;
            igesCreator3.BcurveTol = 0.050799999999999998;
            igesCreator3.FlattenAssembly = false;
            igesCreator3.FileSaveFlag = false;
            igesCreator3.LayerMask = "1-256";
            igesCreator3.ViewList = "Top,Front,Right,Back,Bottom,Left,Isometric,Trimetric,User Defined";
            igesCreator3.ProcessHoldFlag = true;

            // Commit the Builder
            NXOpen.NXObject nXObject3;
            nXObject3 = igesCreator3.Commit();

            // Destroy the IGES File Creator Builder
            igesCreator3.Destroy();
            #endregion

            // Delete Log file
            string delLogFile = outputIgesFileDir + ".log";
            if (File.Exists(delLogFile))
            {
                File.Delete(delLogFile);
            }

        }

    }

    /// <summary>
    /// Batch Export of Parasolid File from given Part file
    /// </summary>
    public class BatchExportParasolid
    {
        // class Memebers
        // Get the NX Session
        private static Session theSession = Session.GetSession();
        private static UFSession theUFSession = UFSession.GetUFSession();
        private static string parasolidFileDirectory;

        /// <summary>
        /// Batch Export of the Parasolid File
        /// </summary>
        /// <param name="partFilePath">First Argument as Part File Directory</param>
        /// <param name="exportDirectory">Directory of Exported File to Save</param>
        public void CreateParasolidFile(string partFilePath, string exportDirectory)
        {
            try
            {
                if (partFilePath == null && exportDirectory == null)
                {
                    //Console.WriteLine("Parasolid File Export input parameter are not Proper");
                    MessageBox.Show("Parasolid File Export input parameter are not Proper");
                    return;
                }

                // Check the Work Part load Status
                PartLoadStatus loadStatus;
                Part workPart = theSession.Parts.OpenDisplay(partFilePath, out loadStatus);
                if (workPart == null)
                {
                    //Console.WriteLine("Part file is null Return");
                    MessageBox.Show("Part file is not Found");
                    return;
                }

                // Directory to save parasolid file 
                parasolidFileDirectory = exportDirectory;

                // Calling ParaSolidFileExport function
                ParaSolidFileExport(workPart);

                // Close the open Work Part
                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

            }
            catch (Exception ex)
            {
                //throw;
            }
        }

        /// <summary>
        /// Create Parasolid File from Part File
        /// </summary>
        /// <param name="workPart">workPart</param>
        private void ParaSolidFileExport(Part workPart)
        {
            try
            {
                if (workPart == null)
                    return;

                // rootComponent for the Assembly Part file 
                Component rootComponent = workPart.ComponentAssembly.RootComponent;

                // Check for Part File or Assembly File
                if (rootComponent == null)
                    ParasolidExport(workPart);
                else
                    ParasolidExport(rootComponent);

            }
            catch (Exception ex)
            {
                //Message.Echo("Error While Exporting Parasolid File " + ex.ToString());
            }
        }

        /// <summary>
        /// Create Parasolid File
        /// </summary>
        /// <param name="BodyTags">Body Tag</param>
        /// <param name="fileName">File Name string</param>
        private void ParasolidFile(Part workPart, Tag[] BodyTags, string fileName)
        {
            try
            {
                // Name of the Parasolid file to save
                string displayBodyName = fileName;
                // location of the file
                string filePaths = workPart.FullPath;
                //filePaths = filePaths.Substring(0, filePaths.LastIndexOf("\\")) + "\\";

                // File Path and Name of the Parasolid file where to save it
                string parasolidFileName = parasolidFileDirectory + "\\" + displayBodyName + ".x_t";

                // Check for any old parasolid file is exist in given directory 
                string delOldParasolidFile = parasolidFileName;
                if (File.Exists(delOldParasolidFile))
                {
                    File.Delete(delOldParasolidFile);
                }

                int numUnexported;
                UFPs.Unexported[] unexported_tags = null;

                // UF function to Export Parasolid File
                theUFSession.Ps.ExportLinkedData(BodyTags, BodyTags.Length, parasolidFileName, 160,
                    null, out numUnexported, out unexported_tags);
            }
            catch (Exception ex)
            {
                // Message.Echo("Error While Parasolid File Funcion" + ex.ToString());
            }
        }

        /// <summary>
        /// Create Parasolid File from an Assembly File
        /// </summary>
        /// <param name="rootComponent">Component</param>
        private void ParasolidExport(Component rootComponent)
        {
            try
            {
                if (rootComponent == null) return;


                // Part from the rootComponent
                Part thePart = (Part)rootComponent.OwningPart;
                // Name of the Parasolid file to save
                string displayPartName = rootComponent.DisplayName;

                // Tag List to collect bodies of Child Components
                List<Tag> componentsBodies = new List<Tag>();

                //Loop all Child Components of the Assembly
                foreach (Component thecomponent in rootComponent.GetChildren())
                {
                    // Converting Component to Part                
                    Part thisPart = thecomponent.Prototype as Part;
                    // Loop all the Bodies in Part
                    foreach (Body body in thisPart.Bodies)
                    {
                        // Adding Bodies of Child Component to List
                        componentsBodies.Add(body.Tag);
                    }
                }
                // Calling Function to Create Parasolid File
                ParasolidFile(thePart, componentsBodies.ToArray(), displayPartName);

                componentsBodies.Clear(); // // Clear the List
            }
            catch (Exception ex)
            {
                //Message.Echo("Error in Assembly Parasolid Export" + ex.ToString());
            }
        }

        /// <summary>
        /// Create Parasolid File from Part File
        /// </summary>
        /// <param name="part">Part</param>
        private void ParasolidExport(Part part)
        {
            try
            {
                if (part == null) return;

                // Name of the Parasolid file to save
                string displayPartName = part.Name;
                // Tag Array to collect bodies of Part
                List<Tag> partBodies = new List<Tag>();

                // Loop all the Bodies in Part
                foreach (Body body in part.Bodies)
                {
                    partBodies.Add(body.Tag);  // Adding Bodies of Part to List
                }
                // Calling Function to Create Parasolid File
                ParasolidFile(part, partBodies.ToArray(), displayPartName);

                partBodies.Clear();  // Clear the List
            }
            catch (Exception ex)
            {
                // Message.Echo("Error in Part Parasolid Export" + ex.ToString());
            }
        }

    }

    /// <summary>
    /// Batch Export of DWG or DXF File from given Part file
    /// </summary>
    public class BatchExportDrawing
    {
        // class Memebers
        // Get the NX Session
        private static Session theSession = Session.GetSession();
        private static UFSession theUFSession = UFSession.GetUFSession();

        /// <summary>
        /// Batch Export of the DWG or DXF Files
        /// </summary>
        /// <param name="partFilePath">First Argument as Part File Directory</param>
        /// <param name="exportFileExtension">Export File Extension as "DWG" or "DXF"</param>
        /// <param name="exportDirectory">Directory of Exported File to Save</param>
        public void CreateDwgDxfFile(string partFilePath, string exportFileExtension, string exportDirectory)
        {
            try
            {
                if (partFilePath == null && exportFileExtension == null && exportDirectory == null)
                {
                    //Console.WriteLine("DWG/DXF File Export input parameter are not Proper");
                    MessageBox.Show("DWG/DXF File Export input parameter are not Proper");
                    return;
                }

                // Check the Work Part load Status
                PartLoadStatus loadStatus;
                Part workPart = theSession.Parts.OpenDisplay(partFilePath, out loadStatus);
                if (workPart == null)
                {
                    //Console.WriteLine("Part file is null Return");
                    MessageBox.Show("Part file is not Found");
                    return;
                }

                // Check for UF_APP_DRAFTING
                theUFSession.UF.AskApplicationModule(out int moduleID);
                if (moduleID != UFConstants.UF_APP_DRAFTING)
                    theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

                // Collect all Drawingsheet to Nxobject Array
                NXObject[] drwSheets = workPart.DrawingSheets.ToArray();

                // Check for drawing sheets in the workpart
                if (drwSheets.Length > 0)
                {
                    // Check for file extension
                    if (exportFileExtension == "DWG" || exportFileExtension == "DXF")
                    {
                        DrawingSheetCollection dSheets = workPart.DrawingSheets;
                        // loop through the drawing sheets
                        foreach (DrawingSheet theSheet in dSheets)
                        {
                            theSheet.Open();

                            //Console.WriteLine("Drawing Sheet Name is " + theSheet.Name);
                            //string exportpath = exportFileDirectory + "\\" + workPart.Name + "_" + theSheet.Name + ".dwg";
                            //Console.WriteLine("Export path " + exportpath);

                            ExportDrawingSheets(workPart, theSheet.Name, exportFileExtension, exportDirectory);
                        }
                    }
                }
                
                //Console.WriteLine("Part name : " +workPart.Name);

                // Close the open Work Part
                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

            }
            catch (Exception ex)
            {
                //throw;
            }
        }

        /// <summary>
        /// To Export the Drawing sheet to .dwg or .dxf Formates
        /// </summary>
        /// <param name="workPart">Work Part</param>
        /// <param name="sheetName">Sheet Name</param>
        /// <param name="fileExtension">Pass String "DWG" for .dwg file Export
        /// or "DXF" for .dxf file Export</param>
        /// <param name="exportFileDirectory">Directory where Export File to Save</param>
        private void ExportDrawingSheets(Part workPart, string sheetName, string fileExtension, string exportFileDirectory)
        {
            try
            {
                // Declaration of dxfdwg file Definition in NX
                string dwgDxfSettingFile;
                // Location of the dxfdwg file Definition in NX 
                string DXFDWG_DIR = theSession.GetEnvironmentVariableValue("DXFDWG_DIR");
                // dxfdwg file Definition File
                dwgDxfSettingFile = Path.Combine(DXFDWG_DIR, "dxfdwg.def");

                // File Path and Name of the dwg file where to save it
                string outfileName = exportFileDirectory + "\\" + workPart.Name + "_" + sheetName;

                #region DxfdwgCreator Builder
                // DWG and DXF File Creator Builder 
                NXOpen.DxfdwgCreator dxfdwgCreator;
                dxfdwgCreator = theSession.DexManager.CreateDxfdwgCreator();
                dxfdwgCreator.ExportData = NXOpen.DxfdwgCreator.ExportDataOption.Drawing;
                dxfdwgCreator.AutoCADRevision = NXOpen.DxfdwgCreator.AutoCADRevisionOptions.R2004;
                dxfdwgCreator.ViewEditMode = true;
                dxfdwgCreator.FlattenAssembly = true;
                dxfdwgCreator.ExportScaleValue = "1:1";
                dxfdwgCreator.SettingsFile = dwgDxfSettingFile;

                // Condition to check weather to Export Drawing to .dwg or .dxf formate
                if (fileExtension == "DWG")
                {
                    dxfdwgCreator.OutputFileType = NXOpen.DxfdwgCreator.OutputFileTypeOption.Dwg;

                    // File Path and Name of the dwg file where to save it
                    dxfdwgCreator.OutputFile = outfileName + ".dwg";
                }
                else if (fileExtension == "DXF")
                {
                    dxfdwgCreator.OutputFileType = NXOpen.DxfdwgCreator.OutputFileTypeOption.Dxf;
                    // File Path and Name of the dxf file where to save it
                    dxfdwgCreator.OutputFile = outfileName + ".dxf";
                }

                dxfdwgCreator.FlattenAssembly = false;
                dxfdwgCreator.InputFile = workPart.FullPath; // part file full Path
                dxfdwgCreator.WidthFactorMode = NXOpen.DxfdwgCreator.WidthfactorMethodOptions.AutomaticCalculation;
                dxfdwgCreator.LayerMask = "1-256";
                // Sheet name/number to Export
                dxfdwgCreator.DrawingList = sheetName;
                dxfdwgCreator.ProcessHoldFlag = true;

                // Commit the dxfdwgCreator Builder
                NXOpen.NXObject nXObject1;
                nXObject1 = dxfdwgCreator.Commit();
                // Destroy the dxfdwgCreator Builder
                dxfdwgCreator.Destroy();
                #endregion

                //Console.WriteLine("Output File name is  "+ outfileName);
                
                // Delete Log file
                string delLogFile = outfileName + ".log";
                //Console.WriteLine("Delete log file" + delLogFile);
                if (File.Exists(delLogFile))
                {
                    //Console.WriteLine(delLogFile + " is Found  ");
                    File.Delete(delLogFile);
                }

            }
            catch (Exception ex)
            {
                // throw;
            }
        }

    }

    /// <summary>
    /// Batch Export of PDF File from given Part file
    /// </summary>
    public class BatchExportPDF
    {
        // class Memebers
        // Get the NX Session
        private static Session theSession = Session.GetSession();
        private static UFSession theUFSession = UFSession.GetUFSession();

        /// <summary>
        /// Batch Export of the PDF File
        /// </summary>
        /// <param name="partFilePath">First Argument as Part File Directory</param>
        /// <param name="exportFileExtension">Export File Extension as "PDF"</param>
        /// <param name="exportDirectory">Directory of Exported File to Save</param>
        public void CreatePdfFile(string partFilePath, string exportFileExtension, string exportDirectory)
        {
            try
            {
                if (partFilePath == null && exportFileExtension == null && exportDirectory == null)
                {
                    //Console.WriteLine("PDF File Export input parameter are not Proper");
                    MessageBox.Show("PDF File Export input parameter are not Proper");
                    return;
                }

                // Check the Work Part load Status
                PartLoadStatus loadStatus;
                Part workPart = theSession.Parts.OpenDisplay(partFilePath, out loadStatus);
                if (workPart == null)
                {
                    //Console.WriteLine("Part file is null Return");
                    MessageBox.Show("Part file is not Found");
                    return;
                }

                // Check for UF_APP_DRAFTING
                theUFSession.UF.AskApplicationModule(out int moduleID);
                if (moduleID != UFConstants.UF_APP_DRAFTING)
                    theSession.ApplicationSwitchImmediate("UG_APP_DRAFTING");

                // Collect all Drawingsheet to Nxobject Array
                NXObject[] drwSheets = workPart.DrawingSheets.ToArray();

                // Check for drawing sheets in work part
                if (drwSheets.Length > 0)
                {
                    ExportDrawingToPDF(workPart, exportDirectory);
                }    

                // Close the open Work Part
                workPart.Close(BasePart.CloseWholeTree.True, BasePart.CloseModified.CloseModified, null);

            }
            catch (Exception ex)
            {
                //throw;
            }
        }

        /// <summary>
        /// To Export the Drawing to .pdf formate
        /// </summary>
        /// <param name="workPart">Work Part</param>
        /// <param name="exportFileDirectory">Directory where Export File to Save</param>
        private void ExportDrawingToPDF(Part workPart, string exportFileDirectory)
        {
            try
            {
                // Collect all Drawingsheet to Nxobject Array
                NXObject[] drwSheets = workPart.DrawingSheets.ToArray();

                // Check for drwSheets are not empty
                if (drwSheets.Length != 0)
                {
                    #region PrintPDFBuilder
                    // PDF File Creator Builder 
                    PrintPDFBuilder printPDFBuilder = workPart.PlotManager.CreatePrintPdfbuilder();
                    // Add collected NXobjects of drawing sheets to source builder
                    printPDFBuilder.SourceBuilder.SetSheets(drwSheets);
                    // File Path and Name of the pdf file where to save it
                    printPDFBuilder.Filename = exportFileDirectory + "\\" + workPart.Name + ".pdf"; //workPart.FullPath.Replace(".prt", ".pdf");
                    printPDFBuilder.Scale = 1.0;
                    printPDFBuilder.Size = PrintPDFBuilder.SizeOption.ScaleFactor;
                    printPDFBuilder.OutputText = PrintPDFBuilder.OutputTextOption.Polylines;
                    printPDFBuilder.Units = PrintPDFBuilder.UnitsOption.English;
                    printPDFBuilder.XDimension = 8.5;
                    printPDFBuilder.YDimension = 11.0;
                    printPDFBuilder.RasterImages = true;
                    printPDFBuilder.Append = false; // true
                    printPDFBuilder.Watermark = "";
                    printPDFBuilder.Colors = PrintPDFBuilder.Color.BlackOnWhite;
                    printPDFBuilder.Commit(); // Commit the builder
                    printPDFBuilder.Destroy(); // Destroy the builder

                    #endregion
                }
            }
            catch (Exception ex)
            {
                //throw;
            }
        }

    }

}
