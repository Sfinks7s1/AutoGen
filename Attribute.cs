using System;

namespace Auto
{
    using Autodesk.AutoCAD.DatabaseServices;
    using Autodesk.AutoCAD.Runtime;

    /// <summary>
    /// Класс включает в себя методы для поиска атрибутов в блоках и значений этих атрибутов
    /// </summary>
    public sealed class Attribute : IExtensionApplication
    {
        public static string BlockName { get; } = "ATTRBLK";
        private const string NameAttrOboznach = "A_OBOZN_DOC";
        private const string NameAttrPageNum = "A_PAGE_NUM";
        private string _oboznach;
        private string _pageNum;

        /// <summary>
        /// Поиск атрибутов штампа, обозначения и номера листа.
        /// </summary>
        public void FindAttribute(Database db)
        {
            try
            {
                _oboznach = AttributeValueFind(NameAttrOboznach, db);
                _pageNum = AttributeValueFind(NameAttrPageNum, db);
            }
            catch (System.Exception)
            {
                // ignored
            }
        }

        public string GetOboznach()
        {
            return _oboznach;
        }

        public string GetPageNub()
        {
            return _pageNum;
        }

        public void Initialize()
        { }

        public void Terminate()
        { }

        /// <summary>
        /// Метод реализует поиск значений атрибутов.
        /// </summary>
        private static string AttributeValueFind(string attbName, Database db)
        {
            var allValueAttribute = string.Empty;

            using (var tr = db.TransactionManager.StartTransaction())
            {
                var bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                //msId = bt[BlockTableRecord.ModelSpace];
                var psId = bt[BlockTableRecord.PaperSpace];

                var btr = (BlockTableRecord)tr.GetObject(psId, OpenMode.ForRead);
                foreach (var entId in btr)
                {
                    var ent = tr.GetObject(entId, OpenMode.ForRead) as Entity;
                    var br = ent as BlockReference;
                    if (br == null) continue;
                    foreach (ObjectId arId in br.AttributeCollection)
                    {
                        var obj = tr.GetObject(arId, OpenMode.ForRead);
                        var ar = obj as AttributeReference;
                        if (ar == null) continue;
                        if (!String.Equals(ar.Tag, attbName, StringComparison.CurrentCultureIgnoreCase)) continue;
                        ar.UpgradeOpen();

                        allValueAttribute = ar.TextString;

                        ar.DowngradeOpen();
                    }
                }
                tr.Commit();
            }
            return allValueAttribute;
        }
    }
}