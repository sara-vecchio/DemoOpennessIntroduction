using Siemens.Engineering.HW;
using Siemens.Engineering;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using Siemens.Engineering.Library;
using Siemens.Engineering.Library.MasterCopies;
using Siemens.Engineering.Library.Types;
using Siemens.Engineering.SW.Blocks;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Types;
using System.Reflection;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.HmiUnified.HmiTags;
using Siemens.Engineering.HmiUnified.UI.Controls;
using Siemens.Engineering.HmiUnified.UI.Parts;
using Siemens.Engineering.HmiUnified.UI.Screens;
using Siemens.Engineering.HmiUnified;

namespace DemoOpennessIntroduction
{
    public partial class Form1 : Form
    {
        string currentDirectory = "";
        string projectPath = "";
        String sProjectDataOra = "";
        string logResult = "";

        TiaPortalProcess process;
        TiaPortal portal = null;
        GlobalLibrary globalLib;
        ProjectLibrary projLib;
        Project project;
        DeviceComposition devices;
        Device[] device;
        DeviceItem[] deviceItem;
        string sPLCTypeName = "System:Device.ET200SP";//"System:Device.S71500";
        PlcSoftware programmaPLC;
        PlcBlockUserGroup groupFolder;

        string excelPath = "";
        Excel.Application xlApp = null;
        Excel.Workbook xlWBook = null;
        Excel.Worksheet xlWSheet = null;
        Excel.Range xlRange = null;

        string sTypeIdentifier = null;
        string sDeviceName = null;
        string sItemName = null;
        string sPlug = null;
        string sNetwork = null;

        int nElements = 1;
        int nRealNDevices = 0;

        OpenFileDialog selectFile;
        public Form1()
        {
            InitializeComponent();
            currentDirectory = System.IO.Directory.GetCurrentDirectory();

            SelectExcelFile.Enabled = true;
            CreateProject.Enabled = false;

            CopyFromLibrary.Enabled = false;
            CopyToLibrary.Enabled = false;
            GenerateHMIPage.Enabled = false;
            selectFile = new OpenFileDialog();
            selectFile.InitialDirectory = currentDirectory;  
        }

        private void SelectExcelFile_Click(object sender, EventArgs e)
        {

             
            selectFile.Filter = "Excel File (*.xlsx)|*.xlsx|General ExcelFile (*.xls)|*.xls";
            if (selectFile.ShowDialog() == DialogResult.OK)
            {
                excelPath = ExcelFileName.Text = selectFile.FileName;
                CreateProject.Enabled = true;
            }
        }

        private void CreateProject_Click(object sender, EventArgs e)
        {
            logResult = logResult + "Caricamento...\r";
            Verbose.Text = logResult;
            portal = new TiaPortal(TiaPortalMode.WithoutUserInterface);
            ProjectComposition projects = portal.Projects;
            sProjectDataOra = ProjectName.Text + DateTime.Now.ToString("_yyyMMdd_HHmmss");
            project = projects.Create(new DirectoryInfo(currentDirectory), sProjectDataOra);
            ProjectName.Text = sProjectDataOra;
            logResult = logResult + ".. Progetto creato. \r\n";
            Verbose.Text = logResult;


            xlApp = new Excel.Application();
            xlApp.Visible = true;
            logResult = logResult + "Caricamento Excel...\r";
            Verbose.Text = logResult;
            try
            {
                xlWBook = xlApp.Workbooks.Open(excelPath);
                xlWSheet = xlWBook.ActiveSheet;
                xlRange = xlWSheet.UsedRange;

                logResult = logResult + ".. Excel aperto.\r\n";
                nElements = 1;

                nElements = ReadExcel.ReadRange(xlRange);

                int nDeviceNameColumn = 1;
                int nDeviceItemNameColumn = 2;
                int nDeviceTypeColumn = 3;
                int nOrderNumberColumn = 4;
                int nFirmwareColumn = 5;
                int nPlugColumn = 6;
                int nControllerColumn = 7;
                int nFAddressColumn = 8;
                int nPluggedDevices = 1;
                int nPLC = 0;

                int nRowOffset = 1;
                // deviceItem = new DeviceItem[xlRange.Row];




                //NetworkInterface[] interfaces = new NetworkInterface[nElements - 1];

                device = new Device[nElements-nRowOffset];
                devices = project.Devices;
                int nRack = 0;

                Subnet PN1 = project.Subnets.Find("PN/IE AB");
                if (PN1 == null) project.Subnets.Create("System:Subnet.Ethernet", "PN/IE AB");
                //Subnet PN2 = project.Subnets.Create("System:Subnet.Ethernet", "PN/IE C");
                IoSystem ioSystem = null;

                for (int i = 1; i <= nElements - nRowOffset; i++)
                {
                    int iOff = i + nRowOffset;
                    nRack = 0;
                    string sFirmware = null, sType = null;

                    ReadExcel.ReadCell(xlRange, out sFirmware, nFirmwareColumn, iOff);
                    ReadExcel.ReadCell(xlRange, out sTypeIdentifier, nOrderNumberColumn, iOff);
                    sTypeIdentifier = "OrderNumber:" + sTypeIdentifier + "/" + sFirmware;
                    ReadExcel.ReadCell(xlRange, out sDeviceName, nDeviceNameColumn, iOff);
                    ReadExcel.ReadCell(xlRange, out sItemName, nDeviceItemNameColumn, iOff);
                    ReadExcel.ReadCell(xlRange, out sType, nDeviceTypeColumn, iOff);
                    ReadExcel.ReadCell(xlRange, out sPlug, nPlugColumn, iOff);
                    ReadExcel.ReadCell(xlRange, out sNetwork, nControllerColumn, iOff);


                    try
                    {
                        if (sDeviceName != null) logResult = logResult + "Creazione device \"" + sDeviceName + "\"\r";
                        else logResult = logResult + "Creazione device item \"" + sItemName + "\"\r";
                        if (sPlug != null && sPlug.Contains("x"))
                        {
                            if (device[nRealNDevices - 1].TypeIdentifier.ToString().Contains("ET200AL")) nRack = 1; //vanno identificati gli ET200AL

                            IList<Siemens.Engineering.HW.Extensions.PlugLocation> plugLocationsList = device[nRealNDevices - 1].DeviceItems[nRack].GetPlugLocations();
                            foreach (Siemens.Engineering.HW.Extensions.PlugLocation item in plugLocationsList)
                            {
                                if (item.PositionNumber > 0 && item.PositionNumber < 100)
                                {
                                    nPluggedDevices = item.PositionNumber;
                                    break;
                                }

                            }

                            device[nRealNDevices - 1].DeviceItems[nRack].PlugNew(sTypeIdentifier, sItemName, nPluggedDevices);



                        }
                        else
                        {
                            device[nRealNDevices] = devices.CreateWithItem(sTypeIdentifier, sItemName, sDeviceName);


                            // Interface is configured as io controller
                            if (sNetwork != null) if (sNetwork.Contains("x")) // && (device[nRealNDevices].DeviceItems[1].GetAttribute("Classification") == "CPU")
                                {
                                    int k = 0, j = 0;
                                    try
                                    {
                                        for (j = 0; j < device[nRealNDevices].DeviceItems.Count; j++)
                                        {
                                            if (device[nRealNDevices].DeviceItems[j].Name.Contains(sItemName)) break;
                                            //First(ee => ee.Name.Contains("PROFINET"))
                                        }
                                        for (k = 0; k < device[nRealNDevices].DeviceItems[j].DeviceItems.Count; k++)
                                        {
                                            if (device[nRealNDevices].DeviceItems[j].DeviceItems[k].Name.Contains("PROFINET")) break;
                                        }

                                        var ntwIntf = device[nRealNDevices].DeviceItems[j].DeviceItems[k].GetService<NetworkInterface>();
                                        ntwIntf.Nodes.First().ConnectToSubnet(PN1);

                                        IoController ioController = ntwIntf.IoControllers.First();

                                        if (ioController != null)
                                        {
                                            ioSystem = ioController.CreateIoSystem("IOsystem" + sItemName);
                                        }
                                    }
                                    catch
                                    {
                                        logResult += "No Item " + device[nRealNDevices].DeviceItems[j].DeviceItems[k].Name;
                                    }

                                }
                                else if (sNetwork.Contains("PLC")) // && (device[nRealNDevices].DeviceItems[1].GetAttribute("Classification") == "CPU")
                                {
                                    int k = 0, j = 0;
                                    try
                                    {
                                        for (j = 0; j < device[nRealNDevices].DeviceItems.Count; j++)
                                        {
                                            if (device[nRealNDevices].DeviceItems[j].Name.Contains(sItemName)) break;
                                        }
                                        for (k = 0; k < device[nRealNDevices].DeviceItems[j].DeviceItems.Count; k++)
                                        {
                                            if (device[nRealNDevices].DeviceItems[j].DeviceItems[k].Name.Contains("PROFINET")) break;
                                        }
                                        var ntwIntf = device[nRealNDevices].DeviceItems[j].DeviceItems[k].GetService<NetworkInterface>();
                                        ntwIntf.Nodes.First().ConnectToSubnet(PN1);
                                        IoConnector ioConnector = ntwIntf.IoConnectors.First();
                                        ioConnector.ConnectToIoSystem(ioSystem);
                                    }
                                    catch
                                    {
                                        logResult += "No Item " + device[nRealNDevices].DeviceItems[j].DeviceItems[k].Name;
                                    }


                                }

                            nRealNDevices++;

                            /*      if (device[nRealNDevices].DeviceItems[1].GetAttribute("Classification") == "CPU")
                                  {
                                     // interfaces[nPLC] = (IEngineeringServiceProvider)device[nRealNDevices].DeviceItems[1].DeviceItems[5].GetService<NetworkInterface>();
                                     // interfaces[nPLC].Nodes.First().CreateAndConnectToSubnet(sNetwork);
                                      nPLC++;
                                  }*/
                        }

                        Verbose.Text = logResult;
                    }
                    catch (EngineeringTargetInvocationException ex) //EngineeringException are those that can be recovered
                    {

                        logResult = logResult + "- eccezione: " + ex.GetType().ToString();
                        try
                        {

                            device[nRealNDevices] = devices.CreateWithItem(sTypeIdentifier, sDeviceName, null);
                            nRealNDevices++;


                        }
                        catch (Exception ex2)
                        {
                            logResult = logResult + " fallita.\r\n";
                            logResult = logResult + ex2.Message;
                            Verbose.Text = logResult;
                        }
                    }
                    finally
                    {
                        logResult = logResult + "\r\n";
                        Verbose.Text = logResult;
                    }
                }
                logResult = logResult + "Hardware Inserito.\r\n";
                Verbose.Text = logResult;
                
                project.Save();
                project.Close();
                portal.Dispose();
                xlWBook.Close();
                xlApp.Quit();
                CopyFromLibrary.Enabled = true;
                CopyToLibrary.Enabled = true;
                GenerateHMIPage.Enabled = true;
            }
            catch
            {
                logResult = logResult + ".. apertura Excel fallita. \r\n";
            }
            Verbose.Text = logResult;

        }

        private void CopyFromLibrary_Click(object sender, EventArgs e)
        {
            try
            {
                string message = "Procedo con la copia";
                groupFolder = programmaPLC.BlockGroup.Groups.Find("myOpennessGroup");
                if (groupFolder == null) groupFolder = programmaPLC.BlockGroup.Groups.Create("myOpennessGroup");
                MessageBox.Show(message);
                              
                foreach (MasterCopy copyOfPlcBlock in globalLib.MasterCopyFolder.MasterCopies)
                {
                    groupFolder.Blocks.CreateFrom(copyOfPlcBlock);
                    logResult = logResult + "Copia da System Folder - " + copyOfPlcBlock.Name + " \r\n";
                    groupFolder.Blocks.Last().Name = groupFolder.Blocks.Last().Name.Replace("in", "da");
                    groupFolder.Blocks.Last().Name = groupFolder.Blocks.Last().Name.Replace("per", "da");
                   
                }
                foreach (MasterCopyUserFolder subFolder in globalLib.MasterCopyFolder.Folders) foreach (MasterCopy copyOfPlcBlock in subFolder.MasterCopies)
                    {
                        groupFolder.Blocks.CreateFrom(copyOfPlcBlock);
                        logResult = logResult + "Copia da" + subFolder.Name + " - " + copyOfPlcBlock.Name + " \r\n";
                        groupFolder.Blocks.Last().Name = groupFolder.Blocks.Last().Name.Replace("in", "da");
                        groupFolder.Blocks.Last().Name = groupFolder.Blocks.Last().Name.Replace("_1", "");
                    }
                logResult = logResult + "Copia Eseguita. \r\n";
                Verbose.Text = logResult;
            }
            catch (Exception ex)
            {
                logResult = logResult + ex.Message;
                Verbose.Text = logResult;
            }


        }

        private void CopyToLibrary_Click(object sender, EventArgs e)
        {
            try
            {

                if(groupFolder==null) groupFolder = programmaPLC.BlockGroup.Groups.Find("myOpennessGroup");

                foreach (PlcBlock block in groupFolder.Blocks)
                {
                     if (block.Name.Contains("Libreria"))
                    {
                        logResult = logResult + "Copia da " + groupFolder.Name + " - " + block.Name + " \r\n";
                        IMasterCopySource cpBlock = block as IMasterCopySource;
                        MasterCopy masterCopy = projLib.MasterCopyFolder.MasterCopies.Create(cpBlock); ////
                        masterCopy.Name = masterCopy.Name.Replace("daLibreria", "daProgetto");
                    }
                }
                Verbose.Text = logResult;

                /* PlcTypeSystemGroup typeGroup = programmaPLC.TypeGroup;
                 Console.Write("TypeGroup: {0} \n", typeGroup.Name);
                 foreach (PlcStruct udtPLC in typeGroup.Types)
                 {
                     Console.Write("Type - Name: {0} - {1} \n", typeGroup.Name, udtPLC.Name);
                 }*/
            }
            catch(Exception ex)
            {
                logResult = logResult + ex.Message;
                Verbose.Text = logResult;
            }
        }

        private void SkipCreate_CheckedChanged(object sender, EventArgs e)
        {
            sProjectDataOra = ProjectName.Text;
            try
            {
                portal = new TiaPortal(TiaPortalMode.WithUserInterface);
                project = portal.Projects.Open(new FileInfo(currentDirectory + "\\" + sProjectDataOra + "\\" + sProjectDataOra + ".ap19"));
                globalLib = portal.GlobalLibraries.Open(new FileInfo(currentDirectory + "\\LibraryDemoOpenness\\LibraryDemoOpenness.al19"), OpenMode.ReadOnly);
                projLib = project.ProjectLibrary;
                selectFile.InitialDirectory = project.Path.Directory.ToString() + "\\UserFiles\\";

                string message = "Aggorno Libreria Progetto";
                MessageBox.Show(message);

                globalLib.UpdateLibrary(new[] { globalLib.TypeFolder }, projLib);

                DeviceComposition existingDeviceComposition = project.Devices;
                Device PLC_device = existingDeviceComposition.First(d => d.TypeIdentifier == sPLCTypeName); /**/
                DeviceItem PLC_item = PLC_device.DeviceItems.First(di => di.Classification == DeviceItemClassifications.CPU);
                //Console.Write("Item {0} \n", PLC_item.Name);
                programmaPLC = PLC_item.GetService<SoftwareContainer>().Software as PlcSoftware; ////
                //Console.Write("Software {0} \n", programmaPLC.Name);

                CopyFromLibrary.Enabled = true;
                CopyToLibrary.Enabled = true;
                GenerateHMIPage.Enabled = true;
            }
            catch (Exception ex)
            {
                if(portal!=null) portal.Dispose();
            }
        }

        private void GenerateHMIPage_Click(object sender, EventArgs e)
        {
            Random rnd = new Random();
            DateTime start = DateTime.Now;

            Device mioPannelloUnified = project.Devices.First(ucp => ucp.DeviceItems[0].TypeIdentifier.Contains("6AV2 128-3"));
            HmiSoftware mioSoftwareUnified = mioPannelloUnified.DeviceItems[3].GetService<SoftwareContainer>().Software as HmiSoftware;
            HmiScreenComposition miaComposizionePagine = mioSoftwareUnified.Screens;
            HmiTagTableComposition miaComposizioneTabelleTag = mioSoftwareUnified.TagTables;
            LibraryTypeComposition miaLibreriaTipi = projLib.TypeFolder.Types;

            /* --- AGGIUNGI PAGINA CON INFO DATA E ORA--- */
            String sDataOra = DateTime.Now.ToString("yyyMMdd_HHmmss");
            HmiScreen nuovaPagina = miaComposizionePagine.Create("PaginaOpenness_" + sDataOra);

            /* --- AGGIUNGI TABELLA CON INFO DATA E ORA--- */
            HmiTagTable nuovaTabella = miaComposizioneTabelleTag.Create("TabellaOpenness_" + sDataOra);

            /* --- AGGIUNGI TAG REAL CON INFO DATA E ORA--- */
            HmiTag nuovaTagSingola = nuovaTabella.Tags.Create("TagReal_" + sDataOra);
            nuovaTagSingola.DataType = "Real";

            /* --- TROVA PRIMO TIPO TAG UDT IN LIBRERIA --- */
            HmiUdtLibraryType miaUDTLibreria = miaLibreriaTipi.First(mudt => mudt.GetType().Name == "HmiUdtLibraryType") as HmiUdtLibraryType;

            /* --- AGGIUNGI TAG UDT VERSIONE DEFAULT CON INFO DATA E ORA--- */
            HmiTag nuovaTagUDT = nuovaTabella.Tags.Create("TagUDT_" + sDataOra); 
           // String sUDTDataType = @"\Project library\Types\" + miaUDTLibreria.Name + "\\V " + miaUDTLibreria.Versions.SingleOrDefault().VersionNumber;
            String sUDTDataType = miaUDTLibreria.Name + " V " + miaUDTLibreria.Versions.SingleOrDefault().VersionNumber;
            try
            {
                nuovaTagUDT.DataType = sUDTDataType;
            }
            catch { nuovaTagUDT.DataType = "Int"; };
            // @"\Project library\Types\Nome UDT\V x.x.x";

            /* --- TROVA TIPO FACEPLATE IN LIBRERIA DAL NOME--- */
            LibraryType mioFaceplateLibreria = miaLibreriaTipi.Find("Faceplate_Example");

            /* --- CREA ISTANZA FACEPLATE VERSIONE DEFAULT NELLA NUOVA PAGINA --- */
            String sContainedTypeValue = "V" + String.Join("\\", mioFaceplateLibreria.Versions.SingleOrDefault().VersionNumber, mioFaceplateLibreria.Name);
            // @"Vx.x.x\Nome Faceplate\";
            HmiFaceplateContainer mioFaceplateContainer = nuovaPagina.ScreenItems.Create<HmiFaceplateContainer>("IstanzaFaceplate_" + sDataOra, sContainedTypeValue);

            /* --- AGGANCIA INTERFACCIA FACEPLATE CON UDT E COLORE VERDE--- */
            int red = rnd.Next(0, 255);
            int green = rnd.Next(0, 255);
            int blu = rnd.Next(0, 255);
            foreach (HmiFaceplateInterface propInterface in mioFaceplateContainer.Interface)
            {
                if (propInterface.PropertyName == "Interface_Tag_1") propInterface.Value = nuovaTagUDT.Name;
                if (propInterface.PropertyName == "MotorColor") propInterface.Value = Color.FromArgb(red, green, blu);

            }

            DateTime end = DateTime.Now;

            logResult = logResult + "Tempo di esecuzione: " + (end - start).TotalMilliseconds.ToString("F0") + " ms\r\n";
            Verbose.Text = logResult;
        }
    }
}
