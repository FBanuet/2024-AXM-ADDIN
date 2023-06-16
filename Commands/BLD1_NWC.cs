using System;
using System.Threading;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using DATools.Util;

namespace DATools.Commands
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class BLD1_NWC : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            OpenOptions opt = new OpenOptions();
            NavisworksExportOptions nweOptions = new NavisworksExportOptions();

            Guid GlobalProjectGuid = new Guid("9352eecd-7567-4ebb-87a6-aacf7a964199");

            
            Guid BLD1 = new Guid("7eb873a6-3476-4148-9f64-d2bc47be03d0");


            Dictionary<string, Guid> PMGuids = new Dictionary<string, Guid>();

            //PMGuids.Add("BLD5-6", BLD56);
            //PMGuids.Add("BLD7-8", BLD78);
            //PMGuids.Add("BLD3-4", BLD34);
            PMGuids.Add("BLD1", BLD1);
            //PMGuids.Add("BLD2", BLD2);
            //PMGuids.Add("BLD9", BLD9);
            //PMGuids.Add("BLDLS", BLDLS);
            //PMGuids.Add("BLDRD", BLDRD);


            DefaultOpenFromCloudCallback defaultOpen = new DefaultOpenFromCloudCallback();

            string viewNames = "";

            foreach (KeyValuePair<string, Guid> PandMguid in PMGuids)
            {
                ModelPath mpath = ModelPathUtils.ConvertCloudGUIDsToCloudPath(GlobalProjectGuid, PandMguid.Value);

                Document doc360 = app.OpenDocumentFile(mpath, opt, defaultOpen);


                try
                {
                    //NEST A FOREACH LOOP THROUG EVERY SET OF VIEWS
                    Nwc3DViewExporter Exporter = new Nwc3DViewExporter();

                    List<View3D> vistasExport = Exporter.GetBIMviews(doc360);

                    foreach (View3D vistaMNSR in vistasExport)
                    {
                        string clave = vistaMNSR.Name;
                        viewNames += clave + Environment.NewLine;
                        nweOptions.ExportScope = NavisworksExportScope.View;
                        nweOptions.ViewId = vistaMNSR.Id;
                        nweOptions.ExportRoomGeometry = false;
                        nweOptions.ExportLinks = true;
                        doc360.Export(@"G:\Unidades compartidas\MANSUR\SURMAN PEDREGAL\01_RESIDENCIA MANSUR\01_AXM\01_MNSR_Project\01_BIM\09_NAVIS\COORDINATION\NWC\TESTING PLUGIN", clave + ".nwc", nweOptions);
                    }


                    doc360.Close(false);

                }
                catch (Exception e)
                {
                    TaskDialog.Show("OPERACIÓN INTERRUMPIDA CON ERRORES", "VERIFICAR NOMENCLATURA DE VISTAS 3D A EXPORTAR COMO NWC" + e.Message);
                }

            }
            TaskDialog.Show("NWC Exporter", "SE HAN EXPORTADO CON ÈXITO LOS SIGUIENTES NWC : " + " - " + viewNames);


            return Result.Succeeded;
        }


        public static string GetNameSpaceCmmnd()
        {
            return typeof(NavisWorksExporterCommand).Namespace + "." + nameof(NavisWorksExporterCommand);
        }
        public OpenConflictResult OnOpenConflict(OpenConflictScenario scenario)
        {
            throw new NotImplementedException();
        }

    }
}
