using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;

[assembly: CommandClass(typeof(AutoCADPlugin.Exporter))]

namespace AutoCADPlugin
{
    public class Exporter
    {
        [CommandMethod("ExportarTubosExcel")]
        public void ExportarTubosExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                Application.ShowAlertDialog("No hay un documento activo de AutoCAD.");
                return;
            }

            Database db = doc.Database;
            Editor ed = doc.Editor;

            string pattern = "^\\\"?(\\\\d+(?:\\\\.\\\\d+)?)\\\"?-([A-ZÑ]+)-(\\d+)-([A-Z0-9]+)$";
            Regex regex = new Regex(pattern, RegexOptions.IgnoreCase);

            var lines = new List<string>();
            lines.Add("Diametro,Material,Linea,Norma,TextoOriginal");
            
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                foreach (ObjectId entId in ms)
                {
                    Entity ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    if (ent is DBText dbText)
                    {
                        ProcesarTexto(dbText.TextString, regex, lines);
                    }
                    else if (ent is MText mText)
                    {
                        ProcesarTexto(mText.Contents, regex, lines);
                    }
                }

                tr.Commit();
            }

            string filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "tubos_exportacion.csv");
            try
            {                File.WriteAllLines(filePath, lines);
                ed.WriteMessage($"\nArchivo exportado: {filePath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError al escribir el archivo: {ex.Message}");
            }
        }

        private void ProcesarTexto(string texto, Regex regex, List<string> lines)
        {
            if (string.IsNullOrWhiteSpace(texto))
                return;

            Match match = regex.Match(texto.Trim());
            if (!match.Success)
                return;

            string diametro = match.Groups[1].Value;
            string material = match.Groups[2].Value;
            string linea = match.Groups[3].Value;
            string norma = match.Groups[4].Value;

            lines.Add($"{diametro},{material},{linea},{norma},{texto}");
        }
    }
}
