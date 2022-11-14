#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;

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


            using (Transaction tx = new Transaction(doc))
            {
                
                tx.Start("Create Levels and Sheets");
                foreach (var level in LevelsList())
                {
                    Level newLevel = Level.Create(doc, level.elevation);
                    newLevel.Name = level.Name;
                }

                foreach (var sheet in SheetList())
                {
                    ViewSheet newSheet = ViewSheet.Create(doc, tblockId);
                    newSheet.Name = sheet.Name;
                    newSheet.SheetNumber = sheet.Number;
                }
                tx.Commit();

            }

            return Result.Succeeded;
        }

        private static List<dSheet> SheetList()
        {
            string sheetsFilePath = @"C:\Users\DGivner\Desktop\API Setup\RAB_Session_02_Challenge_Sheets.csv";

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
            string levelsFilePath = @"C:\Users\DGivner\Desktop\API Setup\RAB_Session_02_Challenge_Levels.csv";

            List<dLevel> levels = new List<dLevel>();
            string[] levelsArray = File.ReadAllLines(levelsFilePath);
            foreach (var levelsRowString in levelsArray)
            {
                string[] levelsCellString = levelsRowString.Split(',');
                var level = new dLevel
                {
                    Name = levelsCellString[0]
                };

                bool didItParse = double.TryParse(levelsCellString[1], out level.elevation);

                levels.Add(level);
            }

            return levels;
        }

    }
}
