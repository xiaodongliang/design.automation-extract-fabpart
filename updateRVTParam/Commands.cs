using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using DesignAutomationFramework;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;

namespace Autodesk.Forge.Sample.DesignAutomation.Revit
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class Commands : IExternalDBApplication
    {
        //Path of the project(i.e)project where your Window family files are present
        string OUTPUT_FILE = "dump_mep_data.csv";

        IList<FabricationService> m_services { get; set; }

        IList<int> m_materialIds { get; set; }
        IList<string> m_materialGroups { get; set; }
        IList<string> m_materialNames { get; set; }

        IList<int> m_specIds { get; set; }
        IList<string> m_specGroups { get; set; }
        IList<string> m_specNames { get; set; }

        IList<int> m_insSpecIds { get; set; }
        IList<string> m_insSpecGroups { get; set; }
        IList<string> m_insSpecNames { get; set; }

        IList<int> m_connIds { get; set; }
        IList<string> m_connGroups { get; set; }
        IList<string> m_connNames { get; set; }

        public ExternalDBApplicationResult OnStartup(ControlledApplication application)
        {
            DesignAutomationBridge.DesignAutomationReadyEvent += HandleDesignAutomationReadyEvent;
            return ExternalDBApplicationResult.Succeeded;
        }

        private void HandleDesignAutomationReadyEvent(object sender, DesignAutomationReadyEventArgs e)
        {
            LogTrace("Design Automation Ready event triggered...");
            e.Succeeded = true;

            extract_mep_data(e.DesignAutomationData.RevitDoc);
        }

        private void extract_mep_data(Document rvtDoc)
        {
            try
            { 
                FabricationConfiguration config = FabricationConfiguration.GetFabricationConfiguration(rvtDoc);
                if (config == null)
                {
                    LogTrace("GetFabricationConfiguration failed");
                    return;
                }
                GetFabricationConnectors(config);

                FilteredElementCollector m_Collector = new FilteredElementCollector(rvtDoc);
                m_Collector.OfClass(typeof(FabricationPart));
                IList<Element> fps = m_Collector.ToElements();

                string outputFullPath = System.IO.Directory.GetCurrentDirectory() + "\\" + OUTPUT_FILE;

                FileStream fs = new FileStream(outputFullPath, System.IO.FileMode.Create, System.IO.FileAccess.Write);
                StreamWriter sw = new StreamWriter(fs, System.Text.Encoding.UTF8);

                foreach (Element el in fps)
                {

                    FabricationPart fp = el as FabricationPart;
                    LogTrace("Reading data of one FabricationPart: " + fp.Name); 
                    sw.WriteLine("Name: " + fp.Name);

                    foreach (Parameter para in el.Parameters){
                        sw.WriteLine("  " + GetParameterInformation(para, rvtDoc));
                    }

                    sw.WriteLine("  ProductName: " + fp.ProductName);
                    sw.WriteLine("  ServiceName: " + fp.ServiceName);
                    sw.WriteLine("  Service Type: " + config.GetServiceTypeName(fp.ServiceType)); 
                    sw.WriteLine("  SpoolName: " + fp.SpoolName);
                    sw.WriteLine("  Alias: " + fp.Alias);
                    sw.WriteLine("  CID: " + fp.ItemCustomId.ToString());
                    sw.WriteLine(" Domain Type: " + fp.DomainType.ToString()); 

                    if (fp.IsAHanger())
                    {
                        string rodKitName = "   None";
                        var rodKit = fp.HangerRodKit;
                        if (rodKit > 0)
                            rodKitName = config.GetAncillaryGroupName(fp.HangerRodKit) + ": " +
                                config.GetAncillaryName(fp.HangerRodKit);

                        sw.WriteLine("  Hanger Rod Kit: " + rodKitName);
                    }

                    var insSpec = config.GetInsulationSpecificationGroup(fp.InsulationSpecification)
                        + ": " + config.GetInsulationSpecificationName(fp.InsulationSpecification);
                    sw.WriteLine("  Insulation Specification: " + insSpec); 
                    sw.WriteLine("  Has No Connections: " + fp.HasNoConnections().ToString()); 
                    sw.WriteLine("  Item Number: " + fp.ItemNumber);

                    var material = config.GetMaterialGroup(fp.Material) + ": " + config.GetMaterialName(fp.Material);
                    sw.WriteLine("  Material: " + material);
                    sw.WriteLine("  Part Guid: " + fp.PartGuid.ToString());
                    sw.WriteLine("  Part Status: " + config.GetPartStatusDescription(fp.PartStatus)); 
                    sw.WriteLine("  Product Code: " + fp.ProductCode);
                     
                    var spec = config.GetSpecificationGroup(fp.Specification) + ": " +
                        config.GetSpecificationName(fp.Specification);
                    sw.WriteLine("  Specification: " + spec);

                    sw.WriteLine("  Dimensions from Fabrication Dimension Definition:");

                    foreach (FabricationDimensionDefinition dmdef in fp.GetDimensions()){
                        sw.WriteLine("      " + dmdef.Name + ":" + fp.GetDimensionValue(dmdef).ToString()); 
                    }

                    int index = 1;
                    sw.WriteLine("      Connectors:");

                    foreach (Connector con in fp.ConnectorManager.Connectors)
                    {
                        sw.WriteLine("          C" + index.ToString());

                        FabricationConnectorInfo fci = con.GetFabricationConnectorInfo();
                        sw.WriteLine("             Connector Info:" + m_connGroups[fci.BodyConnectorId] + 
                                            " " + m_connNames[fci.BodyConnectorId]);

                        sw.WriteLine("             IsBodyConnectorLocked:" + fci.IsBodyConnectorLocked.ToString());

                        sw.WriteLine("             Shape:" + con.Shape.ToString());
                        try
                        { sw.WriteLine("                Radius:" + con.Radius.ToString()); }
                        catch (Exception ex) { }

                        try
                        { sw.WriteLine("                PressureDrop:" + con.PressureDrop.ToString()); }
                        catch (Exception ex) { }

                        try
                        { sw.WriteLine("                Width:" + con.Width.ToString()); }
                        catch (Exception ex) { }

                        try
                        { sw.WriteLine("                Height:" + con.Height.ToString()); }
                        catch (Exception ex) { }

                        try
                        { sw.WriteLine("                PipeSystemType:" + con.PipeSystemType.ToString()); }
                        catch (Exception ex) { }

                        try
                        { sw.WriteLine("                VelocityPressure:" + con.VelocityPressure.ToString()); }
                        catch (Exception ex) { }

                        try
                        { sw.WriteLine("               DuctSystemType:" + con.DuctSystemType.ToString()); }
                        catch (Exception ex) { }
                        index++;
                    } 
                }

                LogTrace("Saving file...");

                sw.Close();
                fs.Close();
            }
            catch(Exception ex)
            {
                LogTrace("Unexpected Exception..." + ex.ToString());

            } 

         }

        public ExternalDBApplicationResult OnShutdown(ControlledApplication application)
        {
            return ExternalDBApplicationResult.Succeeded;
        }

        private void SetElementParameter(FamilySymbol FamSym, BuiltInParameter paraMeter, double parameterValue)
        {
            FamSym.get_Parameter(paraMeter).Set(parameterValue);
        }

        void GetFabricationConnectors(FabricationConfiguration config)
        {
            m_connIds = config.GetAllFabricationConnectorDefinitions(ConnectorDomainType.Undefined, ConnectorProfileType.Invalid);

            m_connGroups = new List<string>();
            m_connNames = new List<string>();

            for (int i = 0; i < m_connIds.Count; i++)
            {
                m_connGroups.Add(config.GetFabricationConnectorGroup(m_connIds[i]));
                m_connNames.Add(config.GetFabricationConnectorName(m_connIds[i]));
            }
        }

        static String GetParameterInformation(Parameter para, Document document)
        {
            string defName = para.Definition.Name + "\t : ";
            string defValue = string.Empty;
            // Use different method to get parameter data according to the storage type
            switch (para.StorageType)
            {
                case StorageType.Double:
                    //covert the number into Metric
                    defValue = para.AsValueString();
                    break;
                case StorageType.ElementId:
                    //find out the name of the element
                    Autodesk.Revit.DB.ElementId id = para.AsElementId();
                    if (id.IntegerValue >= 0)
                    {
                        defValue = document.GetElement(id).Name;
                    }
                    else
                    {
                        defValue = id.IntegerValue.ToString();
                    }
                    break;
                case StorageType.Integer:
                    if (ParameterType.YesNo == para.Definition.ParameterType)
                    {
                        if (para.AsInteger() == 0)
                        {
                            defValue = "False";
                        }
                        else
                        {
                            defValue = "True";
                        }
                    }
                    else
                    {
                        defValue = para.AsInteger().ToString();
                    }
                    break;
                case StorageType.String:
                    defValue = para.AsString();
                    break;
                default:
                    InternalDefinition def = para.Definition as InternalDefinition;
                    BuiltInParameter bp = def.BuiltInParameter;


                    defValue = "Unexposed parameter.";
                    break;
            } 

            return defName + defValue;
        }



        /// <summary>
        /// This will appear on the Design Automation output
        /// </summary>
        private static void LogTrace(string format, params object[] args) { System.Console.WriteLine(format, args); }
    }
}
