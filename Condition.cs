using System.Linq;

namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Runtime;
    using System;
    using System.Diagnostics;
    using System.IO;
    using System.Text.RegularExpressions;

    /// <summary>
    /// Класс содержит методы для реализации различных проверок
    /// </summary>
    public sealed class Condition
    {
        private readonly Document _doc;
        private readonly string _pathTdms = @"C:\Temp";
        private readonly string _relativePath = @"..\";
        private readonly string _relativePathCut = @".\";

        public Condition()
        {
            _doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
        }

        public void Initialize()
        { }

        public void Terminate()
        { }

        /// <summary>
        /// Метод реализует проверку принадлежности чертежа базе TDMS. Если путь сохранения файла не совпадает с путём C:\TEMP\ то метод возвращает false, в противном случае true
        /// </summary>
        public bool CheckPath()
        {
            return (_doc.Name.ToLowerInvariant().Contains(_pathTdms.ToLowerInvariant()));
        }

        /// <summary>
        /// Перегрузка метода CheckPath(), реализует проверку принадлежности чертежа базе TDMS. Если путь сохранения файла не совпадает с путём C:\TEMP\ то метод возвращает false, в противном случае true.
        /// Метод в качестве параметра принимает путь к файлу типа string
        /// </summary>
        public bool CheckPath(string path)
        {
            try
            {
                if (path == String.Empty)
                {
                    _doc.Editor.WriteMessage("\n Path not found");
                    return false;
                }
                else
                {
                    return (String.Equals(path.Remove(7), _pathTdms, StringComparison.InvariantCultureIgnoreCase));
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                _doc.Editor.WriteMessage("\n Exception caught: \n" + ex);
                return false;
            }
        }

        /// <summary>
        /// Перегрузка метода CheckPath(), реализует проверку относительного пути. Если путь сохранения файла не относительный, то метод возвращает false, в противном случае true.
        /// Метод в качестве параметра принимает путь к файлу типа string
        /// </summary>
        public bool CheckRelativePath(string path)
        {
            try
            {
                if (path == String.Empty)
                {
                    _doc.Editor.WriteMessage("\n Path not found");
                    return false;
                }
                else
                {
                    return (String.Equals(path.Remove(3), _relativePath, StringComparison.InvariantCultureIgnoreCase));
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                _doc.Editor.WriteMessage("\n Exception caught: \n" + ex);
                return false;
            }
        }

        /// <summary>
        /// Перегрузка метода CheckPath(), реализует проверку относительного пути. Если путь сохранения файла не относительный, то метод возвращает false, в противном случае true.
        /// Метод в качестве параметра принимает путь к файлу типа string
        /// </summary>
        public bool CheckRelativePathCut(string path)
        {
            try
            {
                if (path == String.Empty)
                {
                    _doc.Editor.WriteMessage("\n Path not found");
                    return false;
                }
                else
                {
                    return (String.Equals(path.Remove(2), _relativePathCut, StringComparison.InvariantCultureIgnoreCase));
                }
            }
            catch (Autodesk.AutoCAD.Runtime.Exception ex)
            {
                _doc.Editor.WriteMessage("\n Exception caught: \n" + ex);
                return false;
            }
        }

        /// <summary>
        /// Метод проверяет существование процесса с именем TDMS, если такой процесс существует в единственном экземпляре, то возвращает true, в противном случае false
        /// </summary>
        public bool CheckTdmsProcess()
        {
            var process = Process.GetProcessesByName("TDMS");
            return (process.Length == 1);
        }

        /// <summary>
        /// Метод реализует поиск в переданной строке GUID объекта
        /// </summary>
        public string ParseGuid(string guidFromFile)
        {
            var regFf = new Regex("{.*?}", RegexOptions.IgnoreCase);
            var mcFf = regFf.Matches(guidFromFile);
            var parseGuidFromFile = mcFf.Cast<Match>().Aggregate<Match, string>(null, (current, mat) => current + mat.Value);
            parseGuidFromFile = parseGuidFromFile.Remove(0, 38);
            return parseGuidFromFile;
        }

        /// <summary>
        /// Метод очищает чертёж от пустых блоков
        /// </summary>
        [CommandMethod("ClearUnrefedBlocks")]
        public void ClearUnrefedBlocks()
        {
            try
            {
                using (var trans = _doc.Database.TransactionManager.StartTransaction())
                {
                    var bt = trans.GetObject(_doc.Database.BlockTableId, OpenMode.ForWrite) as BlockTable;

                    if (bt != null)
                        foreach (var btr in bt.Cast<ObjectId>().Select(oid =>
                        {
                            // ReSharper disable once AccessToDisposedClosure
                            Debug.Assert(trans != null, "trans != null");
                            // ReSharper disable once AccessToDisposedClosure
                            return trans.GetObject(oid, OpenMode.ForWrite) as BlockTableRecord;
                        }).Where(btr => btr.GetBlockReferenceIds(false, false).Count == 0 && !btr.IsLayout))
                        {
                            btr.Erase();
                        }
                    trans.Commit();
                }
            }
            catch (System.Exception ex)
            {
                _doc.Editor.WriteMessage("\n" + ex);
            }
        }

        [CommandMethod("UPURGE", CommandFlags.NoBlockEditor)]
        public static void Upurge()
        {
            var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
            var save = new SaveOptions();
            try
            {
                save.SaveActiveDrawing();

                //doc.SendStringToExecute("-PURGE" + "\n" + "All" + "\n" + "" + "\n" + "No" + "\n", true, false, false);
                //doc.SendStringToExecute("_AUDIT" + "\n" + "Yes" + "\n", true, false, false);
                //doc.SendStringToExecute("_EXPLODEALLPROXY" + "\n", true, false, false);
                //doc.SendStringToExecute("_REMOVEALLPROXY" + "\n", true, false, false);
                const string pathRemdgn = "V:\\\\remdgn";
                doc.SendStringToExecute("(load " + "\"" + pathRemdgn + "\"" + ")" + "\n", true, false, false);
                doc.SendStringToExecute("-PURGE" + "\n" + "All" + "\n" + "" + "\n" + "No" + "\n", true, false, false);
                doc.SendStringToExecute("_AUDIT" + "\n" + "Yes" + "\n", true, false, false);
                //doc.CloseAndDiscard();
            }
            catch (System.Exception ex) { doc.Editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        public class Commands
        {
            private const string DgnLsDefName = "DGNLSDEF";
            private const string DgnLsDictName = "ACAD_DGNLINESTYLECOMP";

            [CommandMethod("DGNPURGE")]
            public static void PurgeDgnLinetypes()
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                PurgeDgnLinetypesInDb(doc.Database, doc.Editor);
            }

            [CommandMethod("DGNPURGEEXT")]
            public static void PurgeDgnLinetypesExt()
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                var ed = doc.Editor;
                var pofo = new PromptOpenFileOptions("\nSelect file to purge");

                var fd = (short)Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("FILEDIA");
                var ca = (short)Autodesk.AutoCAD.ApplicationServices.Core.Application.GetSystemVariable("CMDACTIVE");

                pofo.PreferCommandLine = (fd == 0 || (ca & 36) > 0);
                pofo.Filter = "DWG (*.dwg)|*.dwg|All files (*.*)|*.*";

                var pfnr = ed.GetFileNameForOpen(pofo);

                if (pfnr.Status == PromptStatus.OK)
                {
                    if (!File.Exists(pfnr.StringResult))
                    {
                        ed.WriteMessage("\nCould not find file: \"{0}\".", pfnr.StringResult);
                        return;
                    }

                    try
                    {
                        var output = Path.GetDirectoryName(pfnr.StringResult) + "\\" + Path.GetFileNameWithoutExtension(pfnr.StringResult) + "-purged" + Path.GetExtension(pfnr.StringResult);

                        using (var db = new Database(false, true))
                        {
                            db.ReadDwgFile(pfnr.StringResult, FileOpenMode.OpenForReadAndReadShare, false, "");
                            db.RetainOriginalThumbnailBitmap = true;
                            var wdb = HostApplicationServices.WorkingDatabase;
                            HostApplicationServices.WorkingDatabase = db;
                            if (PurgeDgnLinetypesInDb(db, ed))
                            {
                                var ver = (db.LastSavedAsVersion == DwgVersion.MC0To0 ? DwgVersion.Current : db.LastSavedAsVersion);
                                db.SaveAs(output, ver);
                                ed.WriteMessage("\nSaved purged file to \"{0}\".", output);
                            }
                            HostApplicationServices.WorkingDatabase = wdb;
                        }
                    }
                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                    {
                        ed.WriteMessage("\nException: {0}", ex.Message);
                    }
                }
            }

            private static bool PurgeDgnLinetypesInDb(Database db, Editor ed)
            {
                using (var tr = db.TransactionManager.StartTransaction())
                {
                    var linetypes = CollectComplexLinetypeIds(db, tr);
                    var ltcnt = linetypes.Count;
                    var ltsToKeep = PurgeLinetypesReferencedNotByAnonBlocks(db, tr, linetypes);
                    var strokes = CollectStrokeIds(db, tr);
                    var strkcnt = strokes.Count;

                    PurgeStrokesReferencedByLinetypes(tr, ltsToKeep, strokes);

                    var erasedStrokes = 0;

                    foreach (ObjectId id in strokes)
                    {
                        try
                        {
                            var obj = tr.GetObject(id, OpenMode.ForWrite);

                            obj.Erase();

                            if (obj.GetRXClass().Name.Equals("AcDbLSSymbolComponent"))
                            {
                                EraseReferencedAnonBlocks(tr, obj);
                            }
                            erasedStrokes++;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nUnable to erase stroke ({0}): {1}", id.ObjectClass.Name, ex.Message);
                        }
                    }

                    var erasedLinetypes = 0;

                    foreach (ObjectId id in linetypes)
                    {
                        try
                        {
                            var obj = tr.GetObject(id, OpenMode.ForWrite);
                            obj.Erase();
                            erasedLinetypes++;
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage("\nUnable to erase linetype ({0}): {1}", id.ObjectClass.Name, ex.Message);
                        }
                    }

                    var erasedDict = false;

                    var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                    ed.WriteMessage("\nPurged {0} unreferenced complex linetype records" + " (of {1}).", erasedLinetypes, ltcnt);
                    ed.WriteMessage("\nPurged {0} unreferenced strokes (of {1}).", erasedStrokes, strkcnt);

                    if (nod.Contains(DgnLsDictName))
                    {
                        var dgnLsDict = (DBDictionary)tr.GetObject((ObjectId)nod[DgnLsDictName], OpenMode.ForRead);

                        if (dgnLsDict.Count == 0)
                        {
                            dgnLsDict.UpgradeOpen();
                            dgnLsDict.Erase();
                            ed.WriteMessage("\nRemoved the empty DGN linetype stroke dictionary.");
                            erasedDict = true;
                        }
                    }
                    tr.Commit();
                    return (erasedLinetypes > 0 || erasedStrokes > 0 || erasedDict);
                }
            }

            private static ObjectIdCollection CollectComplexLinetypeIds(Database db, Transaction tr)
            {
                var ids = new ObjectIdCollection();
                var lt = (LinetypeTable)tr.GetObject(db.LinetypeTableId, OpenMode.ForRead);

                foreach (var ltId in lt)
                {
                    var obj = tr.GetObject(ltId, OpenMode.ForRead);
                    if (obj.ExtensionDictionary != ObjectId.Null)
                    {
                        var exd = (DBDictionary)tr.GetObject(obj.ExtensionDictionary, OpenMode.ForRead);

                        if (exd.Contains(DgnLsDefName))
                        {
                            ids.Add(ltId);
                        }
                    }
                }
                return ids;
            }

            private static ObjectIdCollection CollectStrokeIds(Database db, Transaction tr)
            {
                var ids = new ObjectIdCollection();
                var nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                if (nod.Contains(DgnLsDictName))
                {
                    var dgnDict = (DBDictionary)tr.GetObject((ObjectId)nod[DgnLsDictName], OpenMode.ForRead);
                    foreach (var item in dgnDict)
                    {
                        ids.Add(item.Value);
                    }
                }
                return ids;
            }

            private static ObjectIdCollection

            PurgeLinetypesReferencedNotByAnonBlocks(Database db, Transaction tr, ObjectIdCollection ids)
            {
                var keepers = new ObjectIdCollection();
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);

                foreach (var btrId in bt)
                {
                    var btr = (BlockTableRecord)tr.GetObject(btrId, OpenMode.ForRead);

                    foreach (var id in btr)
                    {
                        var obj = tr.GetObject(id, OpenMode.ForRead, true);
                        var ent = obj as Entity;
                        if (ent != null && !ent.IsErased)
                        {
                            if (ids.Contains(ent.LinetypeId))
                            {
                                var owner = (BlockTableRecord)tr.GetObject(ent.OwnerId, OpenMode.ForRead);

                                if (!owner.Name.StartsWith("*") || owner.Name.ToUpper() == BlockTableRecord.ModelSpace || owner.Name.ToUpper().StartsWith(BlockTableRecord.PaperSpace))
                                {
                                    ids.Remove(ent.LinetypeId);
                                    keepers.Add(ent.LinetypeId);
                                }
                            }
                        }
                    }
                }
                return keepers;
            }

            private static void PurgeStrokesReferencedByLinetypes(Transaction tr, ObjectIdCollection tokeep, ObjectIdCollection nodtoremove)
            {
                foreach (ObjectId id in tokeep)
                {
                    PurgeStrokesReferencedByObject(tr, nodtoremove, id);
                }
            }

            private static void PurgeStrokesReferencedByObject(Transaction tr, ObjectIdCollection nodIds, ObjectId id)
            {
                var obj = tr.GetObject(id, OpenMode.ForRead);
                if (obj.ExtensionDictionary != ObjectId.Null)
                {
                    var exd = (DBDictionary)tr.GetObject(obj.ExtensionDictionary, OpenMode.ForRead);

                    if (exd.Contains(DgnLsDefName))
                    {
                        var lsdef = tr.GetObject(exd.GetAt(DgnLsDefName), OpenMode.ForRead);
                        var refFiler = new ReferenceFiler();
                        lsdef.DwgOut(refFiler);

                        foreach (ObjectId refid in refFiler.HardPointerIds)
                        {
                            if (nodIds.Contains(refid))
                            {
                                nodIds.Remove(refid);
                            }
                            PurgeStrokesReferencedByObject(tr, nodIds, refid);
                        }
                    }
                }
                else if (obj.GetRXClass().Name.Equals("AcDbLSCompoundComponent") || obj.GetRXClass().Name.Equals("AcDbLSPointComponent"))
                {
                    var refFiler = new ReferenceFiler();
                    obj.DwgOut(refFiler);

                    foreach (ObjectId refid in refFiler.HardPointerIds)
                    {
                        if (nodIds.Contains(refid))
                        {
                            nodIds.Remove(refid);
                        }
                        PurgeStrokesReferencedByObject(tr, nodIds, refid);
                    }
                }
            }

            private static void EraseReferencedAnonBlocks(Transaction tr, DBObject obj)
            {
                var refFiler = new ReferenceFiler();
                obj.DwgOut(refFiler);
                foreach (ObjectId refid in refFiler.HardPointerIds)
                {
                    var btr = tr.GetObject(refid, OpenMode.ForRead) as BlockTableRecord;
                    if (btr != null && btr.IsAnonymous)
                    {
                        btr.UpgradeOpen();
                        btr.Erase();
                    }
                }
            }
        }
    }
}