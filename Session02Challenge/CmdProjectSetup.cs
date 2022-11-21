#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;
using Forms = System.Windows.Forms;
using Application = Autodesk.Revit.ApplicationServices.Application;

#endregion

namespace Session02Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class CmdProjectSetup : IExternalCommand
    {
        
        public Result Execute(
          ExternalCommandData commandData,
          ref string message,
          ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Application app = uiapp.Application;
            Document doc = uidoc.Document;
            FilteredElementCollector tblockCollector = new FilteredElementCollector(doc);
            tblockCollector.OfCategory(BuiltInCategory.OST_TitleBlocks);
            ElementId tblockId = tblockCollector.FirstElementId();

            var fileName = OpenFile();

            ViewStruct struct1 = new ViewStruct();
            struct1.Name = "View Name";
            struct1.Discipline = "";
            struct1.Level = "";

            ViewStruct struct2 = new ViewStruct("View Name","", "");

            List<ViewStruct> myList = new List<ViewStruct>();
            myList.Add(struct1);
            myList.Add(struct2);

            foreach (var currentStruct in myList)
            {
                Debug.Print(currentStruct.Name);
            }

            FilteredElementCollector viewTemplateCollector = new FilteredElementCollector(doc);
            viewTemplateCollector.OfClass(typeof(ViewTemplateApplicationOption));
            

            FilteredElementCollector vftCollector = new FilteredElementCollector(doc);
            vftCollector.OfClass(typeof(ViewFamilyType));

            ViewFamilyType planVFT = null;
            ViewFamilyType rcpVFT = null;

            foreach (ViewFamilyType vft in vftCollector)
            {
                if(vft.ViewFamily == ViewFamily.FloorPlan) planVFT = vft;

                if(vft.ViewFamily == ViewFamily.CeilingPlan) rcpVFT = vft;
            }

            using (Transaction tx = new Transaction(doc))
            {
                
                tx.Start("Project Setup");
                foreach (var level in LevelsList())
                {
                    Level newLevel = Level.Create(doc, level.Elevation);
                    newLevel.Name = level.Name;
                    
                    ViewPlan newPlanVIew = ViewPlan.Create(doc, planVFT.Id, newLevel.Id);
                    ViewPlan newCeilingPlan = ViewPlan.Create(doc, rcpVFT.Id, newLevel.Id);

                    ViewSheet newSheet = ViewSheet.Create(doc, tblockId);
                    ViewSheet newCeilingSheet = ViewSheet.Create(doc, tblockId);

                    XYZ insertPoint = new XYZ(2, 1, 0);
                    XYZ secondInsertPoint = new XYZ(0, 1, 0);

                    Viewport newViewport = Viewport.Create(doc, newSheet.Id, newPlanVIew.Id, insertPoint);
                    Viewport newCeilingViewport = Viewport.Create(doc, newCeilingSheet.Id, newCeilingPlan.Id, secondInsertPoint);
                }

                foreach (var sheet in SheetList())
                {
                    ViewSheet newSheet = ViewSheet.Create(doc, tblockId);
                    newSheet.Name = sheet.Name;
                    newSheet.SheetNumber = sheet.Number;
                }

                tx.Commit();
                tx.Dispose();

            }

            return Result.Succeeded;
        }

        private static string OpenFile()
        {
            Forms.OpenFileDialog selectFile = new Forms.OpenFileDialog();
            selectFile.InitialDirectory = "C:\\";
            selectFile.Filter = "CSV Files|*.csv";
            selectFile.Multiselect = false;

            string fileName = "";
            if (selectFile.ShowDialog() == Forms.DialogResult.OK)
            {
                fileName = selectFile.FileName;
            }
            
            return fileName;
        }

        private static List<dSheet> SheetList()
        {
            string sheetsFilePath = OpenFile();

            List<dSheet> sheets = new List<dSheet>();
            string[] sheetsArray = File.ReadAllLines(sheetsFilePath);
            foreach (var sheetsRowString in sheetsArray)
            {
                string[] sheetsCellString = sheetsRowString.Split(',');
                var sheet = new dSheet
                {
                    Number = sheetsCellString[0],
                    Name = sheetsCellString[1]
                };

                sheets.Add(sheet);
            }
            return sheets;
        }

        private static List<dLevel> LevelsList()
        {
            string levelsFilePath = OpenFile();

            List<dLevel> levels = new List<dLevel>();
            string[] levelsArray = File.ReadAllLines(levelsFilePath);
            foreach (var levelsRowString in levelsArray)
            {
                string[] levelsCellString = levelsRowString.Split(',');
                var level = new dLevel
                {
                    Name = levelsCellString[0]
                };

                bool didItParse = double.TryParse(levelsCellString[1], out level.Elevation);

                levels.Add(level);
            }
            return levels;
        }
        public static List<View> GetAllViews(Document curDoc)
        {
            FilteredElementCollector allViews = new FilteredElementCollector(curDoc);
            allViews.OfCategory(BuiltInCategory.OST_Views);

            List<View> multiViews = new List<View>();
            foreach (View av in allViews.ToElements())
            {
                multiViews.Add(av);
            }

            return multiViews;
        }
        public static List<View> GetAllViewTemplates(Document curDoc)
        {
            List<View> returnList = new List<View>();
            List<View> viewList = GetAllViews(curDoc);
            foreach (View v in viewList)
            {
                if (v.IsTemplate == true)
                {
                    returnList.Add(v);
                }
            }

            return returnList;
        }
        struct ViewStruct
        {
            public string Name;
            public string Discipline;
            public string Level;

            public ViewStruct(string name, string discipline, string level)
            {
                Name =name;
                Discipline =discipline;
                Level =level;
            }
        }
    }
}
