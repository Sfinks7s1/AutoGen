namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.GraphicsInterface;

    /// <summary>
    /// Класс содержит методы для генерации текста в чертежах AutoCAD
    /// </summary>
    public sealed class Creator
    {
        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        /// <summary>
        /// Метод для создания текстового стиля
        /// </summary>
        public static void CreateStyleText()
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl = tr.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    TextStyleTable st = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForWrite, false);
                    TextStyleTableRecord TStyle = new TextStyleTableRecord();
                    FontDescriptor font = TStyle.Font;

                    TStyle.Name = "STAMP";
                    TStyle.Font = new Autodesk.AutoCAD.GraphicsInterface.FontDescriptor("Tahoma", false, false, font.CharacterSet, font.PitchAndFamily);
                    st.Add(TStyle);

                    tr.AddNewlyCreatedDBObject(TStyle, true);
                    doc.Editor.Regen();
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage("\n" + ex);
            }
        }

        /// <summary>
        /// Метод возвращает название Layout и печатает в LineText
        /// </summary>
        public static void NameDrawingStamp(double x, double y, double z)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    LayoutManager acLayoutMgr;
                    acLayoutMgr = LayoutManager.Current;
                    Layout acLayout;
                    acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

                    using (MText acMText = new MText())
                    {
                        acMText.SetDatabaseDefaults();
                        acMText.Location = new Point3d(x, y, z);

                        acMText.Width = 68;
                        acMText.TextHeight = 2.25;

                        acMText.Contents = "\\pxqc;" + System.Convert.ToString(acLayout.LayoutName);
                        acMText.Layer = "Z-TEXT";
                        acBlkTblRec.AppendEntity(acMText);
                        acTrans.AddNewlyCreatedDBObject(acMText, true);
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Многострочный текст в штампе: OOO "Архитектурное бюро "Студия 44"
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="z"></param>
        public static void NameStudio(double x, double y, double z)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    using (MText acMText = new MText())
                    {
                        acMText.SetDatabaseDefaults();
                        acMText.Location = new Point3d(x, y, z);
                        acMText.Width = 36.838;
                        acMText.Height = 10.034;
                        acMText.TextHeight = 2.5;
                        acMText.LineSpaceDistance = 4.167;
                        acMText.LineSpacingFactor = 1.0;

                        acMText.Contents = "\\pxqc;" + "\"OOO Архитектурное бюро \"Студия 44\"";
                        acMText.Layer = "Z-TEXT";
                        acBlkTblRec.AppendEntity(acMText);
                        acTrans.AddNewlyCreatedDBObject(acMText, true);
                    }
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// метод реализует обновление атрибутов в блоках
        /// </summary>
        /// <param name="blockName"></param>
        /// <param name="text"></param>
        public static void RefreshAttribut(string blockName, string text)
        {
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    using (AttributeDefinition adAttr = new AttributeDefinition())
                    {
                        try
                        {
                            if (btTable.Has(blockName))
                                adAttr.TextString = text;
                        }
                        catch (System.Exception ex) { editor.WriteMessage("\n Invalid block name: " + ex.Message + "\n" + ex.StackTrace); }
                    }
                    trAdding.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Создаётся атрибут штампа, создавать можно как напрямую, так и через AddAttribute(там происходит запрос значений аттрибутов из ТДМС)
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="text"></param>
        /// <param name="height"></param>
        /// <param name="widthFactor"></param>
        /// <param name="rotate"></param>
        /// <param name="oblique"></param>
        /// <param name="attrName"></param>
        /// <param name="blockName"></param>
        /// <param name="textStyle"></param>
        public static void CreateStampAtribut(double x, double y, string text, double height, double widthFactor, double rotate, double oblique, string attrName, int blockName, string textStyle)
        {
            Database dbCurrent = Application.DocumentManager.MdiActiveDocument.Database;
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            ObjectId textStyleId = ObjectId.Null;
            try
            {
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    TextStyleTable TextST = trAdding.GetObject(dbCurrent.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    string strBlockName = "ATTRBLK" + System.Convert.ToString(blockName);
                    try
                    {
                        if (btTable.Has(strBlockName))
                            editor.WriteMessage("\n A block with this name already exist.");
                    }
                    catch (System.Exception ex) { editor.WriteMessage("\n Invalid block name: " + ex.Message + "\n" + ex.StackTrace); }

                    using (AttributeDefinition adAttr = new AttributeDefinition())
                    {
                        DBText acText = new DBText();
                        if (TextST.Has(textStyle))
                        {
                            textStyleId = TextST[textStyle];
                        }
                        else
                        {
                            Creator.CreateStyleText();
                            textStyleId = TextST[textStyle];
                        }

                        adAttr.TextStyleId = textStyleId;
                        adAttr.Position = new Point3d(x, y, 0);
                        adAttr.WidthFactor = widthFactor;
                        adAttr.Height = height;
                        adAttr.Rotation = rotate;
                        adAttr.Oblique = oblique;
                        adAttr.Tag = attrName;

                        using (BlockTableRecord btrRecord = new BlockTableRecord())
                        {
                            btrRecord.Name = strBlockName;
                            btTable.UpgradeOpen();

                            ObjectId btrID = btTable.Add(btrRecord);

                            trAdding.AddNewlyCreatedDBObject(btrRecord, true);

                            btrRecord.AppendEntity(adAttr);
                            trAdding.AddNewlyCreatedDBObject(adAttr, true);

                            BlockTableRecord btrPapperSpace = (BlockTableRecord)trAdding.GetObject(btTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);

                            using (BlockReference brRefBlock = new BlockReference(Point3d.Origin, btrID))
                            {
                                btrPapperSpace.AppendEntity(brRefBlock);
                                trAdding.AddNewlyCreatedDBObject(brRefBlock, true);

                                //задаём значение атрибута
                                using (AttributeReference arAttr = new AttributeReference())
                                {
                                    arAttr.SetAttributeFromBlock(adAttr, brRefBlock.BlockTransform);

                                    arAttr.TextString = text;
                                    arAttr.Layer = "Z-TEXT";

                                    brRefBlock.AttributeCollection.AppendAttribute(arAttr);

                                    trAdding.AddNewlyCreatedDBObject(arAttr, true);
                                }
                            }
                        }
                    }

                    trAdding.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод создаёт многострочный блок для названия чертежа по умолчания записывается наименование листа
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="widthFactor"></param>
        /// <param name="rotate"></param>
        /// <param name="oblique"></param>
        /// <param name="attrName"></param>
        /// <param name="blockName"></param>
        /// <param name="textStyle"></param>
        public static void CreateMultilineStampAtributNameDrawing(double x, double y, double height, double width, double widthFactor, double rotate, double oblique, string attrName, int blockName, string textStyle)
        {
            var docCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var dbCurrent = docCurrent.Database;
            var editor = docCurrent.Editor;
            string text;
            BlockTableRecord btrRecord = null;
            ObjectId btrID = ObjectId.Null;
            try
            {
                docCurrent.TransactionManager.EnableGraphicsFlush(true);
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    TextStyleTable TextST = trAdding.GetObject(dbCurrent.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    LayoutManager acLayoutMgr;
                    acLayoutMgr = LayoutManager.Current;
                    Layout acLayout;
                    acLayout = trAdding.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), OpenMode.ForRead) as Layout;

                    text = "\\pxqc;" + System.Convert.ToString(acLayout.LayoutName);

                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    string strBlockName = "ATTRBLK" + System.Convert.ToString(blockName);
                    using (btrRecord = new BlockTableRecord())
                    {
                        try
                        {
                            if (btTable.Has(strBlockName))
                                editor.WriteMessage("\n A block with this name already exist.");
                            else
                            {
                                using (AttributeDefinition adAttr = new AttributeDefinition())
                                {
                                    if (TextST.Has(textStyle))
                                    {
                                        btrID = TextST[textStyle];
                                    }
                                    else
                                    {
                                        Creator.CreateStyleText();
                                        btrID = TextST[textStyle];
                                    }

                                    adAttr.TextStyleId = btrID;

                                    adAttr.Position = new Point3d(x, y, 0);
                                    adAttr.WidthFactor = widthFactor;

                                    adAttr.Height = height;
                                    adAttr.Rotation = rotate;
                                    adAttr.Oblique = oblique;
                                    adAttr.Justify = AttachmentPoint.MiddleCenter;
                                    adAttr.Tag = attrName;
                                    btTable.UpgradeOpen();

                                    btrRecord.Name = strBlockName;
                                    btTable.Add(btrRecord);
                                    trAdding.AddNewlyCreatedDBObject(btrRecord, true);
                                    btTable.DowngradeOpen();
                                    adAttr.IsMTextAttributeDefinition = true;
                                    adAttr.MTextAttributeDefinition.Contents = text;
                                    adAttr.MTextAttributeDefinition.Width = width;
                                    btrRecord.AppendEntity(adAttr);
                                    trAdding.AddNewlyCreatedDBObject(adAttr, true);
                                }
                            }
                        }
                        catch { editor.WriteMessage("\n Invalid block name."); }

                        BlockTableRecord btrPapperSpace = (BlockTableRecord)trAdding.GetObject(btTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
                        BlockReference brRefBlock = new BlockReference(new Point3d(x, y, 0), btrRecord.ObjectId);
                        btrPapperSpace.AppendEntity(brRefBlock);
                        trAdding.AddNewlyCreatedDBObject(brRefBlock, true);
                        //задаём значение атрибута
                        foreach (ObjectId id in btrRecord)
                        {
                            DBObject obj = id.GetObject(OpenMode.ForRead);
                            AttributeDefinition adAttr = obj as AttributeDefinition;
                            if ((adAttr != null) && (!adAttr.Constant))
                            {
                                AttributeReference arAttr = new AttributeReference();
                                arAttr.SetAttributeFromBlock(adAttr, brRefBlock.BlockTransform);
                                if (arAttr.IsMTextAttribute)
                                {
                                    arAttr.Layer = "Z-TEXT";

                                    MText mText = arAttr.MTextAttribute;
                                    mText.Width = width;
                                    mText.Contents = text;
                                    mText.TextStyleId = btrID;
                                    arAttr.MTextAttribute = mText;
                                    arAttr.TextStyleId = btrID;
                                    arAttr.MTextAttribute.Width = 118;
                                    arAttr.Justify = AttachmentPoint.MiddleCenter;
                                    arAttr.UpdateMTextAttribute();
                                    brRefBlock.RecordGraphicsModified(true);
                                }
                                else
                                {
                                    // add usual attribute reference here
                                }
                                brRefBlock.AttributeCollection.AppendAttribute(arAttr);
                                trAdding.AddNewlyCreatedDBObject(arAttr, true);
                            }
                            trAdding.TransactionManager.QueueForGraphicsFlush();
                        }
                        docCurrent.TransactionManager.FlushGraphics();
                    }
                    trAdding.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
            finally { editor.WriteMessage("\n\t---\tSee results\t---\n"); }
        }

        /// <summary>
        /// Создание многострочного атрибута в штампе
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="text"></param>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="widthFactor"></param>
        /// <param name="rotate"></param>
        /// <param name="oblique"></param>
        /// <param name="attrName"></param>
        /// <param name="blockName"></param>
        /// <param name="textStyle"></param>
        public static void CreateMultilineStampAtribut(double x, double y, string text, double height, double width, double widthFactor, double rotate, double oblique, string attrName, int blockName, string textStyle)
        {
            var docCurrent = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            var dbCurrent = docCurrent.Database;
            var editor = docCurrent.Editor;
            BlockTableRecord btrRecord = null;
            ObjectId btrID = ObjectId.Null;
            try
            {
                docCurrent.TransactionManager.EnableGraphicsFlush(true);
                using (Transaction trAdding = dbCurrent.TransactionManager.StartTransaction())
                {
                    TextStyleTable TextST = trAdding.GetObject(dbCurrent.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;
                    BlockTable btTable = (BlockTable)trAdding.GetObject(dbCurrent.BlockTableId, OpenMode.ForRead);
                    string strBlockName = "ATTRBLK" + System.Convert.ToString(blockName);
                    try
                    {
                        if (btTable.Has(strBlockName))
                            editor.WriteMessage("\n A block with this name already exist.");
                        else
                        {
                            AttributeDefinition adAttr = new AttributeDefinition();
                            if (TextST.Has(textStyle))
                            {
                                btrID = TextST[textStyle];
                            }
                            else
                            {
                                Creator.CreateStyleText();
                                btrID = TextST[textStyle];
                            }
                            adAttr.TextStyleId = btrID;
                            adAttr.Position = new Point3d(x, y, 0);
                            adAttr.WidthFactor = widthFactor;

                            adAttr.Height = height;
                            adAttr.Rotation = rotate;
                            adAttr.Oblique = oblique;
                            adAttr.Justify = AttachmentPoint.MiddleCenter;
                            adAttr.Tag = attrName;
                            btTable.UpgradeOpen();
                            btrRecord = new BlockTableRecord();
                            btrRecord.Name = strBlockName;
                            btTable.Add(btrRecord);
                            trAdding.AddNewlyCreatedDBObject(btrRecord, true);
                            btTable.DowngradeOpen();
                            adAttr.IsMTextAttributeDefinition = true;
                            adAttr.MTextAttributeDefinition.Contents = text;
                            adAttr.MTextAttributeDefinition.Width = width;
                            btrRecord.AppendEntity(adAttr);
                            trAdding.AddNewlyCreatedDBObject(adAttr, true);
                        }
                    }
                    catch (System.Exception ex) { editor.WriteMessage("\n Invalid block name: " + ex.Message + "\n" + ex.StackTrace); }

                    BlockTableRecord btrPapperSpace = (BlockTableRecord)trAdding.GetObject(btTable[BlockTableRecord.PaperSpace], OpenMode.ForWrite);
                    BlockReference brRefBlock = new BlockReference(new Point3d(x, y, 0), btrRecord.ObjectId);
                    btrPapperSpace.AppendEntity(brRefBlock);
                    trAdding.AddNewlyCreatedDBObject(brRefBlock, true);
                    //задаём значение атрибута
                    foreach (ObjectId id in btrRecord)
                    {
                        DBObject obj = id.GetObject(OpenMode.ForRead);
                        AttributeDefinition adAttr = obj as AttributeDefinition;
                        if ((adAttr != null) && (!adAttr.Constant))
                        {
                            AttributeReference arAttr = new AttributeReference();
                            arAttr.SetAttributeFromBlock(adAttr, brRefBlock.BlockTransform);
                            if (arAttr.IsMTextAttribute)
                            {
                                arAttr.Layer = "Z-TEXT";

                                MText mText = arAttr.MTextAttribute;
                                mText.Width = width;
                                mText.Contents = text;
                                mText.TextStyleId = btrID;
                                arAttr.MTextAttribute = mText;
                                arAttr.TextStyleId = btrID;
                                //arAttr.TextString = mText;
                                arAttr.MTextAttribute.Width = width;
                                arAttr.Justify = AttachmentPoint.MiddleCenter;
                                arAttr.UpdateMTextAttribute();
                                brRefBlock.RecordGraphicsModified(true);
                            }
                            else
                            {
                                // add usual attribute reference here
                            }
                            brRefBlock.AttributeCollection.AppendAttribute(arAttr);
                            trAdding.AddNewlyCreatedDBObject(arAttr, true);
                        }
                        trAdding.TransactionManager.QueueForGraphicsFlush();
                    }
                    docCurrent.TransactionManager.FlushGraphics();
                    trAdding.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
            finally { Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\n\t---\tSee results\t---\n"); }
        }

        /// <summary>
        /// Создаём текст в штампе
        /// </summary>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="text"></param>
        /// <param name="height"></param>
        /// <param name="widthFactor"></param>
        /// <param name="rotate"></param>
        /// <param name="oblique"></param>
        /// <param name="textStyle"></param>
        public static void CreateTextStamp(double x, double y, string text, double height, double widthFactor, double rotate, double oblique, string textStyle)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                ObjectId textStyleId = ObjectId.Null;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    TextStyleTable TextST = acTrans.GetObject(acCurDb.TextStyleTableId, OpenMode.ForRead) as TextStyleTable;

                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;

                    DBText acText = new DBText();
                    if (TextST.Has(textStyle))
                    {
                        textStyleId = TextST[textStyle];
                    }
                    else
                    {
                        Creator.CreateStyleText();
                        textStyleId = TextST[textStyle];
                    }
                    acText.TextStyleId = textStyleId;
                    acText.SetDatabaseDefaults();
                    acText.Position = new Point3d(x, y, 0);
                    acText.Height = height;
                    acText.Rotation = rotate;
                    acText.WidthFactor = widthFactor;
                    acText.TextString = text;
                    acText.Oblique = oblique;
                    acText.Layer = "Z-TEXT";
                    acBlkTblRec.AppendEntity(acText);
                    acTrans.AddNewlyCreatedDBObject(acText, true);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод для генерации обычного текста
        /// </summary>
        /// <param name="nameObject"></param>
        /// <param name="x"></param>
        /// <param name="y"></param>
        /// <param name="str"></param>
        public static void CreateText(string nameObject, float x, float y, string str)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                // получаем актуальный документ (тот который открыт в данный момент времени) и базу данных
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                // Начало транзакций
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl; // Открывается блок таблиц для чтения
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec; // Открыть блок таблиц для чтения и записи в модели пространства
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;

                    DBText acText = new DBText();
                    acText.SetDatabaseDefaults();
                    acText.Position = new Point3d(x - 1, y + 4, 0);
                    acText.Height = 2.0;
                    acText.TextString = System.Convert.ToString(nameObject);
                    acBlkTblRec.AppendEntity(acText);
                    acTrans.AddNewlyCreatedDBObject(acText, true);

                    DBText acTextArea = new DBText();
                    acTextArea.SetDatabaseDefaults();
                    acTextArea.Position = new Point3d(x - 1, y, 0);
                    acTextArea.Height = 2.5;
                    acTextArea.TextString = System.Convert.ToString(str);

                    acBlkTblRec.AppendEntity(acTextArea);
                    acTrans.AddNewlyCreatedDBObject(acTextArea, true);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }
    }
}