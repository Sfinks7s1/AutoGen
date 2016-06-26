namespace Auto
{
    using Autodesk.AutoCAD.ApplicationServices;
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.EditorInput;
    using Autodesk.AutoCAD.Geometry;
    using Autodesk.AutoCAD.Runtime;
    using System;

    /// <summary>
    /// Класс содержит методы и конструкторы для генерации рамок и штампов
    /// </summary>
    public sealed class FrameStamp
    {
        /// <summary>
        /// метод возвращает случайное число заданного диапазона
        /// </summary>
        /// <param name="min"></param>
        /// <param name="max"></param>
        /// <returns></returns>
        private static int Rand(int min, int max)
        {
            Random r = new Random();
            return r.Next(min, max + 1);
        }

        private static int blk = Rand(0, 10000);

        private const string rab = "Разработал";
        private const string search = "Проверил";
        private const string GIP = "ГИП";
        private const string GAP = "ГAП";
        private const string normal = "Н.контроль";
        private const string IZM = "Изм.";
        private const string coluch = "Кол.уч.";
        private const string Paper = "Лист";
        private const string Ndoc = "N док.";
        private const string signature = "Подпись";
        private const string date = "Дата";
        private const string stage = "Стадия";
        private const string Papers = "Листов";
        private const string rukGroup = "Рук. группы";
        private const string glConstPr = "Гл.констр.пр.";
        private const string glConstAB = "Гл.констр.\"AБ\"";
        private const string textStyle = "STAMP";

        //Атрибуты штампа
        private const string nameAtrCode = "A_OBOZN_DOC"; //шифр объекта

        private const string nameAtrAddress = "GETOPADDRESS"; //адрес
        private const string nameAtrNameObj = "GETOPNAME"; //наименование объекта
        //private const string nameAtrProjectObj = "A_DESIGN_OBJ_REF"; //название

        private const string AttNameDrawing = "ATT_NAME_DRAWING"; //многострочный атрибут для ввода наименования чертежа, в дальнейшем используется для именования листов после разбивки.

        private const string nameAtrList = "A_PAGE_NUM";
        private const string nameAtrLists = "A_PAGE_COUNT";

        private const string moduleName = "CMD_SYSLIB";
        private const string _GDFOFA = "GETDATAFROMOBJFORACAD";
        private const string _GDFARTFA = "GETDATAFROMACTIVEROUTETABLEFORACAD";
        private const string _GDFSPFA = "GETDATAFROMSYSPROPSFORACAD";

        public void Initialize()
        {
        }

        public void Terminate()
        {
        }

        /// <summary>
        /// Метод добавляет заливку в архитектурные штампы
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="pointX"></param>
        /// <param name="pointY"></param>
        private void AddSolid(int height, int width, double pointX, double pointY)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    double fontSize = 0;
                    var acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline poly = new Polyline();
                    Polyline inContur = new Polyline();

                    poly.SetDatabaseDefaults();

                    //Нижний контур
                    poly.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + 5), 0, fontSize, fontSize);
                    poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                    poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 20), 0, fontSize, fontSize);
                    poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + 20), 0, fontSize, fontSize);
                    poly.Closed = true;
                    poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(poly);
                    acTrans.AddNewlyCreatedDBObject(poly, true);

                    //Верхний контур
                    roomNumPoint = 0;
                    fontSize = 0;
                    inContur.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + height - 25), 0, fontSize, fontSize);
                    inContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 25), 0, fontSize, fontSize);
                    inContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 5), 0, fontSize, fontSize);
                    inContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + height - 5), 0, fontSize, fontSize);
                    inContur.Closed = true;
                    inContur.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(inContur);
                    acTrans.AddNewlyCreatedDBObject(inContur, true);

                    ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                    acObjIdColl.Add(poly.ObjectId);
                    ObjectIdCollection acObjIdColl1 = new ObjectIdCollection();
                    acObjIdColl1.Add(inContur.ObjectId);

                    Hatch acHatch = new Hatch();
                    acBlkTblRec.AppendEntity(acHatch);

                    acTrans.AddNewlyCreatedDBObject(acHatch, true);

                    acHatch.SetDatabaseDefaults();
                    acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                    acHatch.Associative = true;

                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                    acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl1);

                    acHatch.EvaluateHatch(true);
                    acHatch.Layer = "Z-STMP";
                    acHatch.ColorIndex = 9;
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод создаёт форматную рамку с боковыми штампами и текстом в штампах
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="pointX"></param>
        /// <param name="pointY"></param>
        private void FrameSize(double height, double width, double pointX, double pointY)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    double fontSize = 0.25;
                    BlockTable acBlkTbl;
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    BlockTableRecord acBlkTblRec;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline Poly = new Polyline();
                    Polyline InContur = new Polyline();
                    Polyline Number = new Polyline();
                    Polyline sideStamp = new Polyline();

                    Poly.SetDatabaseDefaults();

                    //Внешний контур
                    Poly.AddVertexAt(roomNumPoint, new Point2d(pointX, pointY), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY + height), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(pointX, pointY + height), 0, fontSize, fontSize);
                    Poly.Closed = true;
                    Poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);

                    //Внутренний контур
                    roomNumPoint = 0;
                    fontSize = 0.5;
                    InContur.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 5), 0, fontSize, fontSize);
                    InContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + height - 5), 0, fontSize, fontSize);
                    InContur.Closed = true;
                    InContur.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(InContur);
                    acTrans.AddNewlyCreatedDBObject(InContur, true);

                    //Контур в правом верхнем углу, для сквозной нумерации
                    roomNumPoint = 0;
                    fontSize = 0.4;
                    Number.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + height - 12), 0, fontSize, fontSize);
                    Number.AddVertexAt(++roomNumPoint, new Point2d(pointX - 15, pointY + height - 12), 0, fontSize, fontSize);
                    Number.AddVertexAt(++roomNumPoint, new Point2d(pointX - 15, pointY + height - 5), 0, fontSize, fontSize);
                    Number.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(Number);
                    acTrans.AddNewlyCreatedDBObject(Number, true);

                    //Боковой штамп
                    roomNumPoint = 0;
                    fontSize = 0.5;
                    sideStamp.AddVertexAt(roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                    sideStamp.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 8, pointY + 5), 0, fontSize, fontSize);
                    sideStamp.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 8, pointY + 90), 0, fontSize, fontSize);
                    sideStamp.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 90), 0, fontSize, fontSize);
                    sideStamp.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(sideStamp);
                    acTrans.AddNewlyCreatedDBObject(sideStamp, true);

                    //Боковой штамп
                    CreateStamp(pointX - width + 20, pointY + 30, pointX - width + 8, pointY + 30, fontSize);
                    CreateStamp(pointX - width + 20, pointY + 65, pointX - width + 8, pointY + 65, fontSize);
                    CreateStamp(pointX - width + 20, pointY + 90, pointX - width + 8, pointY + 90, fontSize);
                    CreateStamp(pointX - width + 13, pointY + 5, pointX - width + 13, pointY + 90, fontSize);

                    //Штамп согласования
                    fontSize = 0.25;
                    CreateStamp(pointX - width + 20, pointY + 90, pointX - width, pointY + 90, fontSize);
                    CreateStamp(pointX - width + 20, pointY + 155, pointX - width, pointY + 155, fontSize);
                    CreateStamp(pointX - width + 20, pointY + 110, pointX - width + 5, pointY + 110, fontSize);
                    CreateStamp(pointX - width + 20, pointY + 130, pointX - width + 5, pointY + 130, fontSize);
                    CreateStamp(pointX - width + 20, pointY + 145, pointX - width + 5, pointY + 145, fontSize);

                    CreateStamp(pointX - width + 15, pointY + 90, pointX - width + 15, pointY + 155, fontSize);
                    CreateStamp(pointX - width + 10, pointY + 90, pointX - width + 10, pointY + 155, fontSize);
                    CreateStamp(pointX - width + 5, pointY + 90, pointX - width + 5, pointY + 155, fontSize);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод создаёт форматную рамку для архитекторов с боковыми штампами и текстом в штампах
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="pointX"></param>
        /// <param name="pointY"></param>
        private void FrameSizeArcitect(double height, double width, double pointX, double pointY)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    double fontSize = 0.25;
                    var acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline poly = new Polyline();
                    Polyline inContur = new Polyline();

                    poly.SetDatabaseDefaults();

                    //Внешний контур
                    poly.AddVertexAt(roomNumPoint, new Point2d(pointX, pointY), 0, fontSize, fontSize);
                    poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY), 0, fontSize, fontSize);
                    poly.AddVertexAt(++roomNumPoint, new Point2d(pointX - width, pointY + height), 0, fontSize, fontSize);
                    poly.AddVertexAt(++roomNumPoint, new Point2d(pointX, pointY + height), 0, fontSize, fontSize);
                    poly.Closed = true;
                    poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(poly);
                    acTrans.AddNewlyCreatedDBObject(poly, true);

                    //Внутренний контур
                    roomNumPoint = 0;
                    fontSize = 0.5;
                    inContur.AddVertexAt(roomNumPoint, new Point2d(pointX - 5, pointY + 5), 0, fontSize, fontSize);
                    inContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + 5), 0, fontSize, fontSize);
                    inContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - width + 20, pointY + height - 5), 0, fontSize, fontSize);
                    inContur.AddVertexAt(++roomNumPoint, new Point2d(pointX - 5, pointY + height - 5), 0, fontSize, fontSize);
                    inContur.Closed = true;
                    inContur.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(inContur);
                    acTrans.AddNewlyCreatedDBObject(inContur, true);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод служит для создания штампов
        /// </summary>
        /// <param name="X1"></param>
        /// <param name="Y1"></param>
        /// <param name="X2"></param>
        /// <param name="Y2"></param>
        /// <param name="fontSize"></param>
        private void CreateStamp(double X1, double Y1, double X2, double Y2, double fontSize)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    var acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
                    var acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace], OpenMode.ForWrite) as BlockTableRecord;
                    int roomNumPoint = 0;
                    Polyline Poly = new Polyline();
                    Poly.SetDatabaseDefaults();
                    Poly.AddVertexAt(roomNumPoint, new Point2d(X1, Y1), 0, fontSize, fontSize);
                    Poly.AddVertexAt(++roomNumPoint, new Point2d(X2, Y2), 0, fontSize, fontSize);
                    Poly.Layer = "Z-STMP";
                    acBlkTblRec.AppendEntity(Poly);
                    acTrans.AddNewlyCreatedDBObject(Poly, true);
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Конструктор стандартных рамок по формату
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        private void FrameBuilder(int height, int width)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    FrameStamp createStamp = new FrameStamp();
                    PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
                    PromptPointResult pointResult = editor.GetPoint(pointOptions);
                    Point3d pointFrame = pointResult.Value;
                    Layers.CreateLayerStampAndFrame();
                    createStamp.FrameSize(height, width, pointFrame.X, pointFrame.Y);
                    SideStampText(height, width, pointFrame.X, pointFrame.Y);
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Конструктор архитектурных рамок по формату
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        private void FrameBuilderArch(int height, int width)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                string nameDrawing = "Наименование чертежа";
                string nameObject = "Наименование объекта";

                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    FrameStamp createStamp = new FrameStamp();
                    Technical logo = new Technical();

                    PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
                    PromptPointResult pointResult = editor.GetPoint(pointOptions);
                    Point3d pointFrame = pointResult.Value;
                    Layers.CreateLayerStampAndFrame();

                    createStamp.AddSolid(height, width, pointFrame.X, pointFrame.Y);
                    createStamp.FrameSizeArcitect(height, width, pointFrame.X, pointFrame.Y);

                    logo.AddLogo(pointFrame.X - width + 74, pointFrame.Y + 5);

                    Creator.CreateTextStamp(pointFrame.X - 56, pointFrame.Y + 11, nameDrawing, 3, 1, 0, 0, textStyle);
                    Creator.CreateTextStamp(pointFrame.X - width + 25, pointFrame.Y + height - 16, nameObject, 3, 1, 0, 0, textStyle);

                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А4 горизонтальная
        /// </summary>
        [CommandMethod("A4H", CommandFlags.NoTileMode)]
        public static void A4H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(210, 297); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А4 вертикальная
        /// </summary>
        [CommandMethod("A4V", CommandFlags.NoTileMode)]
        public static void A4V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(297, 210); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка A4 x 3
        /// </summary>
        [CommandMethod("A4H3", CommandFlags.NoTileMode)]
        public static void A4H3()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(297, 630); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка A4 x 4
        /// </summary>
        [CommandMethod("A4H4", CommandFlags.NoTileMode)]
        public static void A4H4()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(297, 841); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А3 горизонтальная
        /// </summary>
        [CommandMethod("A3H", CommandFlags.NoTileMode)]
        public static void A3H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(297, 420); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А3 вертикальная
        /// </summary>
        [CommandMethod("A3V", CommandFlags.NoTileMode)]
        public static void A3V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(420, 297); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А3 x 3
        /// </summary>
        [CommandMethod("A3H3", CommandFlags.NoTileMode)]
        public static void A3H3()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(420, 891); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А2 горизонтальная
        /// </summary>
        [CommandMethod("A2H", CommandFlags.NoTileMode)]
        public static void A2H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(420, 594); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А2 вертикальная
        /// </summary>
        [CommandMethod("A2V", CommandFlags.NoTileMode)]
        public static void A2V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(594, 420); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А1 горизонтальная
        /// </summary>
        [CommandMethod("A1H", CommandFlags.NoTileMode)]
        public static void A1H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(594, 841); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А1 вертикальная
        /// </summary>
        [CommandMethod("A1V", CommandFlags.NoTileMode)]
        public static void A1V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(841, 594); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А1 горизонтальная удлинённая
        /// </summary>
        [CommandMethod("A1HL", CommandFlags.NoTileMode)]
        public static void A1HL()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(594, 1051); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А1 вертикальная удлинённая
        /// </summary>
        [CommandMethod("A1VL", CommandFlags.NoTileMode)]
        public static void A1VL()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(1051, 594); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А0 горизонтальная
        /// </summary>
        [CommandMethod("A0H", CommandFlags.NoTileMode)]
        public static void A0H()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp frame = new FrameStamp(); frame.FrameBuilder(841, 1189); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А0 вертикальная
        /// </summary>
        [CommandMethod("A0V", CommandFlags.NoTileMode)]
        public static void A0V()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilder(1189, 841); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А4 горизонтальная архитектурная
        /// </summary>
        [CommandMethod("A4HA", CommandFlags.NoTileMode)]
        public static void A4HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(210, 297); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А4 вертикальная архитектурная
        /// </summary>
        [CommandMethod("A4VA", CommandFlags.NoTileMode)]
        public static void A4VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(297, 210); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А3 горизонтальная архитектурная
        /// </summary>
        [CommandMethod("A3HA", CommandFlags.NoTileMode)]
        public static void A3HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(297, 420); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А3 вертикальная архитектурная
        /// </summary>
        [CommandMethod("A3VA", CommandFlags.NoTileMode)]
        public static void A3VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(420, 297); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А2 горизонтальная архитектурная
        /// </summary>
        [CommandMethod("A2HA", CommandFlags.NoTileMode)]
        public static void A2HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(420, 594); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А2 вертикальная архитектурная
        /// </summary>
        [CommandMethod("A2VA", CommandFlags.NoTileMode)]
        public static void A2VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(594, 420); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А1 горизонтальная архитектурная
        /// </summary>
        [CommandMethod("A1HA", CommandFlags.NoTileMode)]
        public static void A1HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(594, 841); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А1 вертикальная архитектурная
        /// </summary>
        [CommandMethod("A1VA", CommandFlags.NoTileMode)]
        public static void A1VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(841, 594); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А0 горизонтальная архитектурная
        /// </summary>
        [CommandMethod("A0HA", CommandFlags.NoTileMode)]
        public static void A0HA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(841, 1189); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Рамка А0 горизонтальная архитектурная
        /// </summary>
        [CommandMethod("A0VA", CommandFlags.NoTileMode)]
        public static void A0VA()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { FrameStamp Frame = new FrameStamp(); Frame.FrameBuilderArch(1189, 841); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Конструктор рамок по размерам пользователей
        /// </summary>
        [CommandMethod("FrameCustom", CommandFlags.NoTileMode)]
        public static void FrameCustom()
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            PromptIntegerResult inputHeight = editor.GetInteger("\n Input height: ");
            if (inputHeight.Status != PromptStatus.OK) { editor.WriteMessage("\n No integer was provided"); return; }

            int heightFrame = System.Convert.ToInt16(inputHeight.Value.ToString());

            if (heightFrame < 210 || heightFrame > 1189) { editor.WriteMessage("\n Incorrect frame size"); return; }

            PromptIntegerResult inputWidth = editor.GetInteger("\n Input width: ");

            if (inputWidth.Status != PromptStatus.OK) { editor.WriteMessage("\n No integer was provided"); return; }

            int widthFrame = System.Convert.ToInt16(inputWidth.Value.ToString());

            if (widthFrame < 210 || widthFrame > 1189) { editor.WriteMessage("\n Incorrect frame size"); return; }

            FrameStamp frame = new FrameStamp();
            frame.FrameBuilder(heightFrame, widthFrame);
            editor.WriteMessage("\n Create custom frame: height {0} width {1}", heightFrame, widthFrame);
        }

        /// <summary>
        /// Метод создаёт текст бокового штампа
        /// </summary>
        /// <param name="height"></param>
        /// <param name="width"></param>
        /// <param name="X"></param>
        /// <param name="Y"></param>
        private static void SideStampText(int height, int width, double X, double Y)
        {
            var editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                const double correctWidth = 11.75;
                const double heightText = 2.5;
                const double widthFactor = 0.8;
                const double angle = 1.57079632679; //90 градусов
                const double oblique = 0;
                Creator.CreateTextStamp(X - width + 3.4529, Y + 91.5604, "Согласовано:", heightText, widthFactor, angle, oblique, textStyle);
                Creator.CreateTextStamp(X - width + correctWidth, Y + 67.8355, "Взам. инв.N%%D", heightText, widthFactor, angle, oblique, textStyle);
                Creator.CreateTextStamp(X - width + correctWidth, Y + 36.5089, "Подп. и дата", heightText, widthFactor, angle, oblique, textStyle);
                Creator.CreateTextStamp(X - width + correctWidth, Y + 7.9314, "Инв.N%%D подп.", heightText, widthFactor, angle, oblique, textStyle);
                Technical.AddAttribute(X - 12.6667, Y + height - 10.25, 3.5, 1, 0, 0, "A_TOM_PAGE_NUM", blk++);
                Technical.AddAttribute(X - width + 17.7061, Y + 12.5476, 2.5, 1, angle, 0, "A_ARCH_SIGN", blk++);
                Technical.AddAttributeFunction(X - width + 17.7061, Y + 73.2334, 2.5, 1, angle, 0, "CMD_SYSLIB", "GETDATAFROMPAROBJFORACAD", blk++, "A_INSTEAD_OF_NUM");
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Метод для создания штампа в зависимости от мастерской
        /// </summary>
        /// <param name="stampType"></param>
        private static void StampBuilder(string stampType)
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try
            {
                Document acDoc = Application.DocumentManager.MdiActiveDocument;
                Database acCurDb = acDoc.Database;

                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    FrameStamp createStamp = new FrameStamp();

                    PromptPointOptions pointOptions = new PromptPointOptions("УКАЖИТЕ ТОЧКУ: ");
                    PromptPointResult pointResult = editor.GetPoint(pointOptions);
                    Point3d selectedPoint = pointResult.Value;
                    Layers.CreateLayerStampAndFrame();
                    //Горизонтальные линии штампа
                    createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y + 55, selectedPoint.X, selectedPoint.Y + 55, 0.5);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y + 15, selectedPoint.X, selectedPoint.Y + 15, 0.5);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y + 30, selectedPoint.X, selectedPoint.Y + 30, 0.5);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y + 45, selectedPoint.X, selectedPoint.Y + 45, 0.5);
                    createStamp.CreateStamp(selectedPoint.X - 50, selectedPoint.Y + 25, selectedPoint.X, selectedPoint.Y + 25, 0.5);
                    int step = 5;
                    double fontSize = 0.25;
                    for (int i = 0; i < 10; ++i)
                    {
                        if (i == 5 || i == 6)
                        {
                            fontSize = 0.5;
                            createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y + step, selectedPoint.X - 120, selectedPoint.Y + step, fontSize);
                            fontSize = 0.25;
                        }
                        createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y + step, selectedPoint.X - 120, selectedPoint.Y + step, fontSize);
                        step += 5;
                    }
                    //Вертикальные линии штампа
                    step = 10;
                    fontSize = 0.5;
                    createStamp.CreateStamp(selectedPoint.X - 185, selectedPoint.Y, selectedPoint.X - 185, selectedPoint.Y + 55, fontSize);
                    for (int i = 0; i < 4; ++i)
                    {
                        if (i == 1 || i == 3)
                        {
                            createStamp.CreateStamp(selectedPoint.X - 185 + step, selectedPoint.Y, selectedPoint.X - 185 + step, selectedPoint.Y + 55, fontSize);
                        }
                        createStamp.CreateStamp(selectedPoint.X - 185 + step, selectedPoint.Y + 30, selectedPoint.X - 185 + step, selectedPoint.Y + 55, fontSize);
                        step += 10;
                    }
                    createStamp.CreateStamp(selectedPoint.X - 130, selectedPoint.Y, selectedPoint.X - 130, selectedPoint.Y + 55, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 120, selectedPoint.Y, selectedPoint.X - 120, selectedPoint.Y + 55, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 50, selectedPoint.Y, selectedPoint.X - 50, selectedPoint.Y + 30, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 35, selectedPoint.Y + 15, selectedPoint.X - 35, selectedPoint.Y + 30, fontSize);
                    createStamp.CreateStamp(selectedPoint.X - 20, selectedPoint.Y + 15, selectedPoint.X - 20, selectedPoint.Y + 30, fontSize);

                    //Вставляем логотип
                    Technical logo = new Technical();
                    logo.AddLogo(selectedPoint.X, selectedPoint.Y);

                    //Создание текста в штампе
                    if (stampType != "stampar")
                    {
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 26.25, rab, 2.25, 1, 0, 0, textStyle);    //Разработал
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 21.25, search, 2.25, 1, 0, 0, textStyle);    //Проверил
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 16.25, rukGroup, 2.25, 1, 0, 0, textStyle);    //Руководитель группы
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 11.25, normal, 2.25, 1, 0, 0, textStyle);    //Нормоконтроль
                        Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 6.25, glConstPr, 2.25, 1, 0, 0, textStyle);    //Главный конструктор проекта

                        //ФАМИЛИИ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 26.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_User"); //Разработал
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 21.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_User");   //Проверил
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 16.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GR_HEAD", "A_User"); //Руководитель группы
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 11.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_User");  //Нормоконтроль
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 6.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GKP_", "A_User");     //Главный конструктор проекта

                        //ДАТЫ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 26.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 21.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 16.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GR_HEAD", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 11.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 6.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GKP_", "A_DATE");

                        if (stampType == "stampkab")
                        {
                            //Главный конструктор AБ
                            Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 1.25, 2.25, 1, 0, 0, moduleName, _GDFSPFA, blk++, "GKAB_");
                            Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 1.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GKAB_", "A_DATE");
                            Creator.CreateTextStamp(selectedPoint.X - 184.587, selectedPoint.Y + 1.25, glConstAB, 2.25, 0.92, 0, 0, textStyle);
                        }
                    }
                    else
                    {
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 26.25, rab, 2.25, 1, 0, 0, textStyle);     //Разработал
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 21.25, GIP, 2.25, 1, 0, 0, textStyle);     //ГИП
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 16.25, GAP, 2.25, 1, 0, 0, textStyle);     //ГАП
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 11.25, search, 2.25, 1, 0, 0, textStyle);     //Проверил
                        Creator.CreateTextStamp(selectedPoint.X - 183.8, selectedPoint.Y + 1.25, normal, 2.25, 1, 0, 0, textStyle);     //Нормоконтроль

                        //ФАМИЛИИ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 26.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_User"); //Разработал
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 21.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GIP_", "A_User");    //ГИП
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 16.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GAP_", "A_User");    //ГАП
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 11.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_User");   //Проверил
                        Technical.AddAttributeFunction(selectedPoint.X - 164.292, selectedPoint.Y + 1.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_User");   //Нормоконтроль

                        //ДАТЫ (Берутся из ТДМС)
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 26.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "DEVELOP", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 21.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GIP_", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 16.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "GAP_", "A_DATE");
                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 11.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "CHECK", "A_DATE");

                        Technical.AddAttributeFunction(selectedPoint.X - 128.793, selectedPoint.Y + 1.25, 2.25, 1, 0, 0, moduleName, _GDFARTFA, blk++, "NORMKL", "A_DATE");
                    }
                    Creator.CreateTextStamp(selectedPoint.X - 182.513, selectedPoint.Y + 31.55, IZM, 2, 0.9, 0, 0.1745329, textStyle);  //Изменения
                    Creator.CreateTextStamp(selectedPoint.X - 174.145, selectedPoint.Y + 31.55, coluch, 2, 0.9, 0, 0.1745329, textStyle);  //Количество участников
                    Creator.CreateTextStamp(selectedPoint.X - 162.9, selectedPoint.Y + 31.55, Paper, 2, 0.9, 0, 0.1745329, textStyle);  //Лист
                    Creator.CreateTextStamp(selectedPoint.X - 153.63, selectedPoint.Y + 31.55, Ndoc, 2, 0.9, 0, 0.1745329, textStyle);  //Номер документа
                    Creator.CreateTextStamp(selectedPoint.X - 142.81, selectedPoint.Y + 31.55, signature, 2, 0.9, 0, 0.1745329, textStyle);  //Подпись
                    Creator.CreateTextStamp(selectedPoint.X - 128.065, selectedPoint.Y + 31.55, date, 2, 0.9, 0, 0.1745329, textStyle);  //Дата

                    Creator.CreateTextStamp(selectedPoint.X - 46.610, selectedPoint.Y + 26.53, stage, 2, 0.9, 0, 0.1745329, textStyle);    //Стадия
                    Creator.CreateTextStamp(selectedPoint.X - 30.283, selectedPoint.Y + 26.53, Paper, 2, 0.9, 0, 0.1745329, textStyle);    //Лист
                    Creator.CreateTextStamp(selectedPoint.X - 13.757, selectedPoint.Y + 26.53, Papers, 2, 0.9, 0, 0.1745329, textStyle);    //Листов

                    //Cоздание атрибутов в штампе
                    //Середина штампа
                    Technical.AddAttributeMultiline(selectedPoint.X - 60, selectedPoint.Y + 50, 4, 118, 1, 0, 0, nameAtrCode, blk++);       //шифр объекта

                    //после внесения изменений
                    Technical.AddMultilineAttributeFunction(selectedPoint.X - 60, selectedPoint.Y + 37, 3, 118, 1, 0, 0, moduleName, nameAtrNameObj, blk++); //Наименование объекта
                    Technical.AddMultilineAttributeFunction(selectedPoint.X - 85, selectedPoint.Y + 22, 2.5, 68, 1, 0, 0, moduleName, nameAtrAddress, blk++);   //адрес объекта

                    //Номера листов и стадия (Берутся из ТДМС)
                    Technical.AddAttributeFunction(selectedPoint.X - 44.1065, selectedPoint.Y + 17.8617, 4, 1, 0, 0, moduleName, _GDFOFA, blk++, "A_STAGE_CLSF");
                    Technical.AddAttribute(selectedPoint.X - 29.4215, selectedPoint.Y + 17.8617, 4, 1, 0, 0, nameAtrList, blk++);
                    Technical.AddAttribute(selectedPoint.X - 12.0115, selectedPoint.Y + 17.8617, 4, 1, 0, 0, nameAtrLists, blk++);

                    Creator.CreateMultilineStampAtributNameDrawing(selectedPoint.X - 85, selectedPoint.Y + 7, 2.5, 68, 1, 0, 0, AttNameDrawing, blk++, textStyle); //Наименование чертежа из названия layout

                    editor.WriteMessage(selectedPoint.ToString());
                    acTrans.Commit();
                }
            }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Запуск создания форматного штампа с наименованием чертежа из названия layout
        /// Для конструкторов с корректно заполненным штампом
        /// </summary>
        [CommandMethod("StampKAB", CommandFlags.NoTileMode)]
        public void StampKAB()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { StampBuilder("stampkab"); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Запуск создания форматного штампа с наименованием чертежа из названия layout
        /// Для конструкторов с корректно заполненным штампом
        /// </summary>
        [CommandMethod("StampK", CommandFlags.NoTileMode)]
        public void StampK()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { StampBuilder("stampk"); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }

        /// <summary>
        /// Запуск создания форматного штампа с наименованием чертежа из названия layout
        /// Для архитекторов с корректно заполненным штампом
        /// </summary>
        [CommandMethod("StampAR", CommandFlags.NoTileMode)]
        public void StampAR()
        {
            Editor editor = Application.DocumentManager.MdiActiveDocument.Editor;
            try { StampBuilder("stampar"); }
            catch (System.Exception ex) { editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace); }
        }
    }
}