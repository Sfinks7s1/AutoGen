namespace Auto
{
    using Autodesk.AutoCAD.Colors;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Runtime;
    using System.IO;

    /// <summary>
    /// Класс содержит методы для генерации слоёв в чертеже AutoCAD
    /// </summary>
    public sealed class Layers
    {
        private const string pathREMDGN = @"V:\\lay-del";

        /// <summary>
        /// Метод создаёт слои, на которых затем генерируются штампы и рамки
        /// </summary>
        public static void CreateLayerStampAndFrame()
        {
            var editor = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                var acDoc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                var acCurDb = acDoc.Database;

                using (var acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    var acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead) as LayerTable;
                    var sLayerNames = new string[3];

                    sLayerNames[0] = "Z-TEXT";
                    sLayerNames[1] = "Z-STMP";
                    sLayerNames[2] = "Defpoints";

                    var acColors = new Color[3];
                    acColors[0] = Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[1] = Color.FromColorIndex(ColorMethod.ByAci, 7);
                    acColors[2] = Color.FromColorIndex(ColorMethod.ByAci, 7);

                    var acLineWeight = new LineWeight[3];

                    acLineWeight[0] = LineWeight.LineWeight030;
                    acLineWeight[1] = LineWeight.LineWeight050;
                    acLineWeight[2] = LineWeight.LineWeight000;

                    var nCnt = 0;

                    foreach (var sLayerName in sLayerNames)
                    {
                        LayerTableRecord acLyrTblRec;
                        if (acLyrTbl.Has(sLayerName) == false)
                        {
                            acLyrTblRec = new LayerTableRecord { Name = sLayerName };
                            if (acLyrTbl.IsWriteEnabled == false) acLyrTbl.UpgradeOpen();

                            acLyrTbl.Add(acLyrTblRec);
                            acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                        }
                        else { acLyrTblRec = acTrans.GetObject(acLyrTbl[sLayerName], OpenMode.ForWrite) as LayerTableRecord; }
                        acLyrTblRec.Color = acColors[nCnt];
                        acLyrTblRec.LineWeight = acLineWeight[nCnt];
                        nCnt += 1;
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Удаляет текущий слой вместе с объектами на нем
        /// </summary>
        [CommandMethod("LDLC")]
        public static void LDLC()
        {
            if (System.Windows.Forms.MessageBox.Show("Действительно удалить текущий слой?", "Работа со слоями", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

                try
                {
                    if (File.Exists(pathREMDGN))
                    {
                        doc.SendStringToExecute("(load " + "\"" + pathREMDGN + "\"" + ")" + "\n", true, false, false);
                        doc.SendStringToExecute("_DLC" + "\n", true, false, false);
                        TraceSource.TraceMessage(TraceType.Information, "Команда _DLC выполнена.");
                    }
                    else
                    {
                        TraceSource.TraceMessage(TraceType.Information, "Отсутствует файл lay - del по адресу: V:\\. Обратитесь к администратору.");
                        doc.Editor.WriteMessage("\n Отсутствует файл lay-del по адресу: V:\\. Обратитесь к администратору.");
                    }
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
                }
            }
        }

        /// <summary>
        /// Удаляет замороженные слои вместе с объектами на них
        /// </summary>
        [CommandMethod("LDLF")]
        public static void LDLF()
        {
            if (System.Windows.Forms.MessageBox.Show("Действительно удалить замороженные слои?", "Работа со слоями", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;

                try
                {
                    if (File.Exists(pathREMDGN))
                    {
                        doc.SendStringToExecute("(load " + "\"" + pathREMDGN + "\"" + ")" + "\n", true, false, false);
                        doc.SendStringToExecute("_DLF" + "\n", true, false, false);
                    }
                    else
                    {
                        doc.Editor.WriteMessage("\n Отсутствует файл lay-del по адресу: V:\\. Обратитесь к администратору.");
                    }
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
                }
            }
        }

        /// <summary>
        /// Удаляет непечатаемые слои вместе с объектами на них
        /// </summary>
        [CommandMethod("LDLNP")]
        public static void LDLNP()
        {
            if (System.Windows.Forms.MessageBox.Show("Действительно удалить непечатаемые слои?", "Работа со слоями", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                try
                {
                    if (File.Exists(pathREMDGN))
                    {
                        doc.SendStringToExecute("(load " + "\"" + pathREMDGN + "\"" + ")" + "\n", true, false, false);
                        doc.SendStringToExecute("_DLNP" + "\n", true, false, false);
                    }
                    else
                    {
                        doc.Editor.WriteMessage("\n Отсутствует файл lay-del по адресу: V:\\. Обратитесь к администратору.");
                    }
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
                }
            }
        }

        /// <summary>
        /// Удаляет выключенные слои вместе с объектами на них
        /// </summary>
        [CommandMethod("LDLO")]
        public static void LDLO()
        {
            if (System.Windows.Forms.MessageBox.Show("Действительно удалить выключенные слои?", "Работа со слоями", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.Yes)
            {
                var doc = Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument;
                try
                {
                    if (File.Exists(pathREMDGN))
                    {
                        doc.SendStringToExecute("(load " + "\"" + pathREMDGN + "\"" + ")" + "\n", true, false, false);
                        doc.SendStringToExecute("_DLO" + "\n", true, false, false);
                    }
                    else
                    {
                        doc.Editor.WriteMessage("\n Отсутствует файл lay-del по адресу: V:\\. Обратитесь к администратору.");
                    }
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
                }
            }
        }
    }
}