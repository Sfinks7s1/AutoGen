using Autodesk.AutoCAD.Ribbon;
using Autodesk.AutoCAD.Runtime;
using Autodesk.Windows;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Imaging;
using Application = Autodesk.AutoCAD.ApplicationServices.Core.Application;

[assembly: ExtensionApplication(typeof(Auto.Starter))]

namespace Auto
{
    /// <summary>
    /// Класс реализует меню ТДМС на ленте в AutoCAD
    /// </summary>
    public class Ribbon
    {
        public RibbonCombo Pan1Ribcombo1 = new RibbonCombo();
        public RibbonCombo Pan3Ribcombo = new RibbonCombo();

        [CommandMethod("MyRibbon")]
        public void MyRibbon()
        {
            var ribbonControl = ComponentManager.Ribbon;

            var tab = new RibbonTab
            {
                Title = "TDMS",
                Id = "RibbonSample_TAB_ID"
            };

            if (FindRibbon(ribbonControl))
            {
                ribbonControl.Tabs.Add(tab);

                //////////////////////////////
                // Создание панелей на ленте//
                //////////////////////////////

                var panelSaveFunction = PanelSourceSaveFunction(tab);
                var panelXref = PanelSourceXref(tab);
                var panelStamps = PanelSourceStamps(tab);
                var panelList = PanelSourceList(tab);
                var panelHelp = PanelSourceHelp(tab);
                var panelClose = PanelSourceClose(tab);

                /////////////////////////////////////
                //  ПАНЕЛЬ С ФУНКЦИЯМИ СОХРАНЕНИЯ  //
                // Кнопка сохранение чертежа в базе//
                /////////////////////////////////////

                // Кнопка Сохранения в базе
                var rbSave = new RibbonButton();
                var rttSave = new RibbonToolTip
                {
                    Title = "Сохранить",
                    Content =
                        "Сохранить изменения в чертеже в TDMS." + "\n" +
                        "Флаг редактирования чертежа в TDMS не снимается."
                };
                rbSave.CommandParameter = rttSave.Command = "_SaveActiveDrawing";
                rbSave.Text = "Сохранить";
                rbSave.ShowText = true;
                rbSave.ShowImage = true;
                rbSave.LargeImage = Images.GetBitmap(Properties.Resources.Save);
                rbSave.Size = RibbonItemSize.Large;
                rbSave.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbSave.CommandHandler = new RibbonCommandHandler();
                rbSave.ToolTip = rttSave;

                // Кнопка Сохранения и закрытия чертежа в базе
                var rbSaveAndClose = new RibbonButton();
                var rttSaveAndClose = new RibbonToolTip
                {
                    Title = "Сохранить и закрыть",
                    Content =
                        "Сохранить чертёж в TDMS и закрыть его в AutoCAD." + "\n" +
                        "Флаг редактирования чертежа в TDMS снимается."
                };
                rbSaveAndClose.CommandParameter = rttSaveAndClose.Command = "_SaveAndClose";
                rbSaveAndClose.Text = "Сохранить" + "\n" + "и закрыть";
                rbSaveAndClose.ShowText = true;
                rbSaveAndClose.ShowImage = true;
                rbSaveAndClose.LargeImage = Images.GetBitmap(Properties.Resources.SaveAndClose);
                rbSaveAndClose.Size = RibbonItemSize.Large;
                rbSaveAndClose.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbSaveAndClose.CommandHandler = new RibbonCommandHandler();
                rbSaveAndClose.ToolTip = rttSaveAndClose;

                // Кнопка Сохранения и закрытия всех чертежей в базе
                var rbSaveAndCloseAll = new RibbonButton();
                var rttSaveAndCloseAll = new RibbonToolTip
                {
                    Title = "Сохранить и закрыть все чертежи",
                    Content =
                        "Сохранить чертежи и закрыть в AutoCAD." + "\n" +
                        "Флаг редактирования чертежа в TDMS снимается."
                };
                rbSaveAndCloseAll.CommandParameter = rttSaveAndCloseAll.Command = "_SaveAndCloseAll";
                rbSaveAndCloseAll.Text = "Сохранить и закрыть" + "\n" + "все чертежи";
                rbSaveAndCloseAll.ShowText = true;
                rbSaveAndCloseAll.ShowImage = true;
                rbSaveAndCloseAll.LargeImage = Images.GetBitmap(Properties.Resources.SaveAndCloseAll);
                rbSaveAndCloseAll.Size = RibbonItemSize.Large;
                rbSaveAndCloseAll.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbSaveAndCloseAll.CommandHandler = new RibbonCommandHandler();
                rbSaveAndCloseAll.ToolTip = rttSaveAndCloseAll;

                //Кнопка для чистки чертежей
                var rbUberPurge = new RibbonButton();
                var rttUberPurge = new RibbonToolTip
                {
                    Title = "Очистка",
                    Content =
                        "После выполнения команды рекомендуется чертёж сохранить, закрыть и открыть заново. Выполняется следующая последовательность действий по очистке чертежа: "
                        + "\n" + "Выполняется чистка утилитой remdgn"
                        + "\n" + "_PURGE"
                        + "\n" + "_AUDIT"
                };
                rbUberPurge.CommandParameter = rttUberPurge.Command = "_UPURGE";
                rbUberPurge.Text = "Очистка";
                rbUberPurge.ShowText = true;
                rbUberPurge.ShowImage = true;
                rbUberPurge.LargeImage = Images.GetBitmap(Properties.Resources.Clear);
                rbUberPurge.Size = RibbonItemSize.Large;
                rbUberPurge.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbUberPurge.CommandHandler = new RibbonCommandHandler();
                rbUberPurge.ToolTip = rttUberPurge;

                // Кнопка закрытия чертежа в базе и в AutoCAD
                var rbClose = new RibbonButton();
                var rttClose = new RibbonToolTip
                {
                    Title = "Закрыть",
                    Content =
                        "Закрыть чертёж, открытый из TDMS без сохранения изменений." + "\n" +
                        "Флаг редактирования чертежа в TDMS снимается."
                };
                rbClose.CommandParameter = rttClose.Command = "_CloseAndDiscard";
                rbClose.Text = "Закрыть" + "\n" + "без сохранения";
                rbClose.ShowText = true;
                rbClose.ShowImage = true;
                rbClose.LargeImage = Images.GetBitmap(Properties.Resources.Close);
                rbClose.Size = RibbonItemSize.Large;
                rbClose.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbClose.CommandHandler = new RibbonCommandHandler();
                rbClose.ToolTip = rttClose;

                // Кнопка закрытия всех чертежей в базе и в AutoCAD
                var rbCloseAll = new RibbonButton();
                var rttCloseAll = new RibbonToolTip
                {
                    Title = "Закрыть все чертежи",
                    Content =
                        "Закрыть все открытые чертежи без сохранения изменений." + "\n" +
                        "Флаг редактирования чертежа в TDMS снимается."
                };
                rbCloseAll.CommandParameter = rttCloseAll.Command = "_CloseAndDiscardAll";
                rbCloseAll.Text = "Закрыть" + "\n" + "всё";
                rbCloseAll.ShowText = true;
                rbCloseAll.ShowImage = true;
                rbCloseAll.LargeImage = Images.GetBitmap(Properties.Resources.CloseAll);
                rbCloseAll.Size = RibbonItemSize.Large;
                rbCloseAll.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbCloseAll.CommandHandler = new RibbonCommandHandler();
                rbCloseAll.ToolTip = rttCloseAll;

                // Кнопка для сохранения чертежа с переопределением внешних ссылок в ТДМС (относительные ссылки)
                var rbSaveDwGwithXreFinTdms = new RibbonButton();
                var rttSaveDwGwithXreFinTdms = new RibbonToolTip
                {
                    Title = "Сохранить чертёж с внешними ссылками в TDMS",
                    Content =
                        "После выполнения команды следует в диалоговом окне TDMS выбрать каталог для импорта чертежа." +
                        "\n" +
                        "При этом основной чертёж сохранится в выбранном каталоге, а все входящие в него внешние ссылки " +
                        "\n" +
                        "будут сохранены в подкаталоге c подписью (XREF)." + "\n" +
                        "Пути для внешних ссылок в основном файле будут переопределены на относительные."
                };
                rbSaveDwGwithXreFinTdms.CommandParameter = rttSaveDwGwithXreFinTdms.Command = "_RELATIVEPATH";
                rbSaveDwGwithXreFinTdms.Text = "Сохранить в TDMS (XREF)";
                rbSaveDwGwithXreFinTdms.ShowText = true;
                rbSaveDwGwithXreFinTdms.ShowImage = true;
                rbSaveDwGwithXreFinTdms.Image = Images.GetBitmap(Properties.Resources.SaveDWGwithXREFinTDMS);
                rbSaveDwGwithXreFinTdms.CommandHandler = new RibbonCommandHandler();
                rbSaveDwGwithXreFinTdms.ToolTip = rttSaveDwGwithXreFinTdms;

                // Кнопка для сохранения чертежа с переопределением внешних ссылок на локальные HDD
                var rbSaveDwGwithXreFonHdd = new RibbonButton();
                var rttSaveDwGwithXreFonHdd = new RibbonToolTip
                {
                    Title = "Сохранить чертёж с внешними ссылками локально.",
                    Content =
                        "После выполнения команды следует в диалоговом окне задать путь для сохранения чертежа." + "\n" +
                        "При этом основной чертёж сохранится в выбранном каталоге, а все входящие в него внешние ссылки " +
                        "\n" +
                        "будут сохранены в подкаталоге c подписью (XREF)." + "\n" +
                        "Пути для внешних ссылок в основном файле будут переопределены на относительные."
                };
                rbSaveDwGwithXreFonHdd.CommandParameter = rttSaveDwGwithXreFonHdd.Command = "_EXPORTXREFFROMTDMS";
                rbSaveDwGwithXreFonHdd.Text = "Сохранить локально (XREF)";
                rbSaveDwGwithXreFonHdd.ShowText = true;
                rbSaveDwGwithXreFonHdd.ShowImage = true;
                rbSaveDwGwithXreFonHdd.Image = Images.GetBitmap(Properties.Resources.SaveDWGwithXREFonHDD);
                rbSaveDwGwithXreFonHdd.CommandHandler = new RibbonCommandHandler();
                rbSaveDwGwithXreFonHdd.ToolTip = rttSaveDwGwithXreFonHdd;

                // Кнопка для функции ETRANSMIT
                var rbEtransmit = new RibbonButton();
                var rttEtransmit = new RibbonToolTip
                {
                    Title = "Сформировать комплект",
                    Content =
                        "Команда предназначена для подготовки чертежа к отправке заказчикам или смежным отделам." + "\n" +
                        "При выполнении происходит следующее:" + "\n" +
                        "1) Выполняются команды purge и audit;" + "\n" +
                        "2) Все внешние ссылки внедряются в чертёж" + "\n" +
                        "3) Чертёж конвертируется в формат AutoCAD 2007" + "\n" +
                        "4) Сохранение результата в каталоге: D:\\ETRANSMIT"
                };
                rbEtransmit.CommandParameter = rttEtransmit.Command = "_PACKAGE";
                rbEtransmit.Text = "Сформировать комплект";
                rbEtransmit.ShowText = true;
                rbEtransmit.ShowImage = true;
                rbEtransmit.Image = Images.GetBitmap(Properties.Resources.etransmit);
                rbEtransmit.CommandHandler = new RibbonCommandHandler();
                rbEtransmit.ToolTip = rttEtransmit;

                var pan1Row1 = new RibbonRowPanel();

                pan1Row1.Items.Add(rbSaveDwGwithXreFinTdms);
                pan1Row1.Items.Add(new RibbonRowBreak());
                pan1Row1.Items.Add(rbSaveDwGwithXreFonHdd);
                pan1Row1.Items.Add(new RibbonRowBreak());
                pan1Row1.Items.Add(rbEtransmit);

                var pan2SplitButton = new RibbonSplitButton
                {
                    Text = "SplitButton",
                    CommandHandler = new RibbonCommandHandler(),
                    ShowText = true,
                    ShowImage = true,
                    IsSplit = true,
                    Size = RibbonItemSize.Large
                };

                var pan7SplitButton = new RibbonSplitButton
                {
                    Text = "SplitButton",
                    CommandHandler = new RibbonCommandHandler(),
                    ShowText = true,
                    ShowImage = true,
                    IsSplit = true,
                    Size = RibbonItemSize.Large
                };

                var pan8SplitButton = new RibbonSplitButton
                {
                    Text = "SplitButton",
                    CommandHandler = new RibbonCommandHandler(),
                    ShowText = true,
                    ShowImage = true,
                    IsSplit = true,
                    Size = RibbonItemSize.Large
                };

                //Кнопка для добавления штампа для архитекторов
                var rbStampAr = new RibbonButton();
                var rttStampAr = new RibbonToolTip
                {
                    Title = "Штамп для архитекторов",
                    Content =
                        "Команда выполняется в листах документа. В точке вставки добавляется штамп с атрибутами. Все атрибуты штампа ссылаются на данные из TDMS."
                };
                rbStampAr.CommandParameter = rttStampAr.Command = "_STAMPAR";
                rbStampAr.Text = "Архитекторы";
                rbStampAr.ShowText = true;
                rbStampAr.ShowImage = true;
                rbStampAr.LargeImage = Images.GetBitmap(Properties.Resources.stampar);
                rbStampAr.Size = RibbonItemSize.Large;
                rbStampAr.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbStampAr.CommandHandler = new RibbonCommandHandler();
                rbStampAr.ToolTip = rttStampAr;

                // Кнопка для добавления штампа для конструкторов
                var rbStampK = new RibbonButton();
                var rttStampK = new RibbonToolTip
                {
                    Title = "Штамп для конструкторов",
                    Content =
                        "Команда выполняется в листах документа. В точке вставки добавляется штамп с атрибутами. Все атрибуты штампа ссылаются на данные из TDMS."
                };
                rbStampK.CommandParameter = rttStampK.Command = "_STAMPK";
                rbStampK.Text = "Конструкторы";
                rbStampK.ShowText = true;
                rbStampK.ShowImage = true;
                rbStampK.LargeImage = Images.GetBitmap(Properties.Resources.stampk);
                rbStampK.Size = RibbonItemSize.Large;
                rbStampK.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbStampK.CommandHandler = new RibbonCommandHandler();
                rbStampK.ToolTip = rttStampK;

                // Кнопка для добавления штампа для конструкторов
                var rbStampKab = new RibbonButton();
                var rttStampKab = new RibbonToolTip
                {
                    Title = "Штамп для конструкторов (с главным конструктором архитектурного бюро)",
                    Content =
                        "Команда выполняется в листах документа. В точке вставки добавляется штамп с атрибутами. Все атрибуты штампа ссылаются на данные из TDMS."
                };
                rbStampKab.CommandParameter = rttStampKab.Command = "_STAMPKAB";
                rbStampKab.Text = "Конструкторы с АБ";
                rbStampKab.ShowText = true;
                rbStampKab.ShowImage = true;
                rbStampKab.LargeImage = Images.GetBitmap(Properties.Resources.stampk);
                rbStampKab.Size = RibbonItemSize.Large;
                rbStampKab.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbStampKab.CommandHandler = new RibbonCommandHandler();
                rbStampKab.ToolTip = rttStampKab;

                // Кнопка для удаления непечатаемых слоёв вместе с объектами на них
                var rbDlnp = new RibbonButton();
                var rttDlnp = new RibbonToolTip
                {
                    Title = "Удаление" + "\n" + "непечат.слоёв",
                    Content =
                        "Команда выполняет удаление непечатаемых слоёв вместе с расположенными на них объектами. Рекомендуется предварительно сохранить чертёж. Выполнение команды может занять некоторое время."
                };
                rbDlnp.CommandParameter = rttDlnp.Command = "_LDLNP";
                rbDlnp.Text = "Удаление" + "\n" + "непечат.слоёв";
                rbDlnp.ShowText = true;
                rbDlnp.ShowImage = true;
                rbDlnp.Image = Images.GetBitmap(Properties.Resources.DLNP);
                rbDlnp.LargeImage = Images.GetBitmap(Properties.Resources.DLNP);
                rbDlnp.Size = RibbonItemSize.Large;
                rbDlnp.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbDlnp.CommandHandler = new RibbonCommandHandler();
                rbDlnp.ToolTip = rttDlnp;

                // Кнопка для удаления замороженных слоёв вместе с объектами на них
                var rbLdlf = new RibbonButton();
                var rttLdlf = new RibbonToolTip
                {
                    Title = "Удаление" + "\n" + "заморож.слоёв",
                    Content =
                        "Команда выполняет удаление замороженных слоёв вместе с расположенными на них объектами. Рекомендуется предварительно сохранить чертёж. Выполнение команды может занять некоторое время."
                };
                rbLdlf.CommandParameter = rttLdlf.Command = "_LDLF";
                rbLdlf.Text = "Удаление" + "\n" + "заморож.слоёв";
                rbLdlf.ShowText = true;
                rbLdlf.ShowImage = true;
                rbLdlf.Image = Images.GetBitmap(Properties.Resources.DLF);
                rbLdlf.LargeImage = Images.GetBitmap(Properties.Resources.DLF);
                rbLdlf.Size = RibbonItemSize.Large;
                rbLdlf.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbLdlf.CommandHandler = new RibbonCommandHandler();
                rbLdlf.ToolTip = rttLdlf;

                // Кнопка для удаления выключенных слоёв вместе с объектами на них
                var rbLdlo = new RibbonButton();
                var rttLdlo = new RibbonToolTip
                {
                    Title = "Удаление" + "\n" + "выкл.слоёв",
                    Content =
                        "Команда выполняет удаление выключенных слоёв вместе с расположенными на них объектами. Рекомендуется предварительно сохранить чертёж. Выполнение команды может занять некоторое время."
                };
                rbLdlo.CommandParameter = rttLdlo.Command = "_LDLO";
                rbLdlo.Text = "Удаление" + "\n" + "выкл.слоёв";
                rbLdlo.ShowText = true;
                rbLdlo.ShowImage = true;
                rbLdlo.Image = Images.GetBitmap(Properties.Resources.DLO);
                rbLdlo.LargeImage = Images.GetBitmap(Properties.Resources.DLO);
                rbLdlo.Size = RibbonItemSize.Large;
                rbLdlo.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbLdlo.CommandHandler = new RibbonCommandHandler();
                rbLdlo.ToolTip = rttLdlo;

                // Кнопка для удаления текущего слоя вместе с объектами на нем
                var rbLdlc = new RibbonButton();
                var rttLdlc = new RibbonToolTip
                {
                    Title = "Удаление" + "\n" + "текущего_слоя",
                    Content =
                        "Команда выполняет удаление текущего слоя вместе с расположенными на нём объектами. Рекомендуется предварительно сохранить чертёж. Выполнение команды может занять некоторое время."
                };
                rbLdlc.CommandParameter = rttLdlc.Command = "_LDLC";
                rbLdlc.Text = "Удаление" + "\n" + "текущего_слоя";
                rbLdlc.ShowText = true;
                rbLdlc.ShowImage = true;
                rbLdlc.Image = Images.GetBitmap(Properties.Resources.DLC);
                rbLdlc.LargeImage = Images.GetBitmap(Properties.Resources.DLC);
                rbLdlc.Size = RibbonItemSize.Large;
                rbLdlc.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbLdlc.CommandHandler = new RibbonCommandHandler();
                rbLdlc.ToolTip = rttLdlc;

                pan2SplitButton.Items.Add(rbStampAr);
                pan2SplitButton.Items.Add(rbStampK);
                pan2SplitButton.Items.Add(rbStampKab);

                pan8SplitButton.Items.Add(rbDlnp);
                pan8SplitButton.Items.Add(rbLdlf);
                pan8SplitButton.Items.Add(rbLdlo);
                pan8SplitButton.Items.Add(rbLdlc);

                ////////////////////////////////////////////////
                //                                            //
                //         ПАНЕЛЬ ДЛЯ РАБОТЫ С ЛИСТАМИ        //
                //                                            //
                ////////////////////////////////////////////////

                //Кнопка для экспорта листов из чертежа
                var rbExportLayout = new RibbonButton();
                var rttExportLayout = new RibbonToolTip
                {
                    Title = "Экспорт листов из чертежа",
                    Content = "После выполнения команды происходит разбивка чертежа (Лист + Модель) и" + "\n" +
                              "сохранение результата в каталоге: d:\\_TDMSLAYOUT\\."
                };
                rbExportLayout.CommandParameter = rttExportLayout.Command = "_DivideLayout";
                rbExportLayout.Text = "Экспорт";
                rbExportLayout.ShowText = true;
                rbExportLayout.ShowImage = true;
                rbExportLayout.Image = Images.GetBitmap(Properties.Resources.ExportLayouts);
                rbExportLayout.LargeImage = Images.GetBitmap(Properties.Resources.ExportLayouts);
                rbExportLayout.Size = RibbonItemSize.Large;
                rbExportLayout.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbExportLayout.CommandHandler = new RibbonCommandHandler();
                rbExportLayout.ToolTip = rttExportLayout;

                ///////////////////////////////////////////////
                //            ПАНЕЛЬ ВНЕШНИХ ССЫЛОК          //
                ///////////////////////////////////////////////

                // Кнопка для добавления чертежа или изображения как внешней ссылки из ТДМС
                var rbDwgFromTdms = new RibbonButton();
                var rttDwgFromTdms = new RibbonToolTip
                {
                    Title = "Добавить внешнюю ссылку из TDMS",
                    Content = "При выполнении команды выберите в качестве объекта внешней ссылки" +
                              "любой чертёж из TDMS (формат *.dwg) или любое изображение из TDMS (*.jpeg, *.bmp, *.tiff, *.png)." +
                              "Укажите в пространстве модели точку вставки"
                };
                rbDwgFromTdms.CommandParameter = rttDwgFromTdms.Command = "_TDMSXREFADD";
                rbDwgFromTdms.Text = "Добавить" + "\n" + "внешнюю ссылку";
                rbDwgFromTdms.ShowText = true;
                rbDwgFromTdms.ShowImage = true;
                rbDwgFromTdms.LargeImage = Images.GetBitmap(Properties.Resources.DWGFromTDMS);
                rbDwgFromTdms.Size = RibbonItemSize.Large;
                rbDwgFromTdms.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbDwgFromTdms.CommandHandler = new RibbonCommandHandler();
                rbDwgFromTdms.ToolTip = rttDwgFromTdms;

                //Кнопка для добавления ссылки в таблицу спецификаций из внешнего чертежа.
                var rbAddTable = new RibbonButton();
                var rttAddTable = new RibbonToolTip
                {
                    Title = "Добавить внешнюю ссылку на таблицу из TDMS",
                    Content = "При выполнении команды:" +
                              "1) Выберите таблицу, которая будет ссылаться на спецификацию из другого файла; /n" +
                              "2) В диалоговом окне выбора файла выберите чертёж из папки проектно-сметной докуентации."
                };
                rbAddTable.CommandParameter = rttAddTable.Command = "_AddExternalSheduleTable";
                rbAddTable.Text = "Добавить ссылку на спецификацию";
                rbAddTable.ShowText = true;
                rbAddTable.ShowImage = true;
                rbAddTable.Image = Images.GetBitmap(Properties.Resources.tableadd);
                rbAddTable.CommandHandler = new RibbonCommandHandler();
                rbAddTable.ToolTip = rttAddTable;

                //Кнопка для обновления ссылки в таблице спецификаций.
                var rbUpdateTable = new RibbonButton();
                var rttUpdateTable = new RibbonToolTip
                {
                    Title = "Обновить внешнюю ссылку " + "\n" + " на таблицу из TDMS",
                    Content = "При выполнении команды происходит обновление таблицы спецификаций"
                };
                rbUpdateTable.CommandParameter = rttUpdateTable.Command = "_UpdateTable";
                rbUpdateTable.Text = "Обновить спецификации";
                rbUpdateTable.ShowText = true;
                rbUpdateTable.ShowImage = true;
                rbUpdateTable.Image = Images.GetBitmap(Properties.Resources.tableupdate);
                rbUpdateTable.CommandHandler = new RibbonCommandHandler();
                rbUpdateTable.ToolTip = rttUpdateTable;

                var pan4Row1 = new RibbonRowPanel();

                pan4Row1.Items.Add(rbAddTable);
                pan4Row1.Items.Add(new RibbonRowBreak());
                pan4Row1.Items.Add(rbUpdateTable);

                //Кнопка для редактирования внешней ссылки в пространстве чертежа.
                var rbEditXreFinPlace = new RibbonButton();
                var rttEditXreFinPlace = new RibbonToolTip
                {
                    Title = "Редактировать вхождение внешней ссылки из TDMS.",
                    Content = "После выполнения команды выберите блок внешней ссылки для редактирования."
                };
                rbEditXreFinPlace.CommandParameter = rttEditXreFinPlace.Command = "_XREFEDIT";
                rbEditXreFinPlace.Text = "Ред.Xref";
                rbEditXreFinPlace.ShowText = true;
                rbEditXreFinPlace.ShowImage = true;
                rbEditXreFinPlace.Image = Images.GetBitmap(Properties.Resources.EditXREFInPlace);
                rbEditXreFinPlace.CommandHandler = new RibbonCommandHandler();
                rbEditXreFinPlace.ToolTip = rttEditXreFinPlace;

                //Кнопка для сохранения изменений во внешней ссылке в пространстве чертежа.
                var rbSaveXreFinPlace = new RibbonButton();
                var rttSaveXreFinPlace = new RibbonToolTip
                {
                    Title = "Сохранение изменений во внешней ссылке из TDMS.",
                    Content = "После выполнения команды произойдёт сохранение и закрытие редактируемой внешней ссылки."
                };
                rbSaveXreFinPlace.CommandParameter = rttSaveXreFinPlace.Command = "_SaveCloseXREFInPlace";
                rbSaveXreFinPlace.Text = "Сохр.Xref";
                rbSaveXreFinPlace.ShowText = true;
                rbSaveXreFinPlace.ShowImage = true;
                rbSaveXreFinPlace.Image = Images.GetBitmap(Properties.Resources.SaveXREFinPlace);
                rbSaveXreFinPlace.CommandHandler = new RibbonCommandHandler();
                rbSaveXreFinPlace.ToolTip = rttSaveXreFinPlace;

                //Кнопка для отмены изменений во внешней ссылке в пространстве чертежа.
                var rbCloseXreFinPlace = new RibbonButton();
                var rttCloseXreFinPlace = new RibbonToolTip
                {
                    Title = "Отмена изменений во внешней ссылке из TDMS.",
                    Content =
                        "После выполнения команды произойдёт закрытие редактируемой внешней ссылки без сохранения изменений."
                };
                rbCloseXreFinPlace.CommandParameter = rttCloseXreFinPlace.Command = "_CloseXREFInPlace";
                rbCloseXreFinPlace.Text = "Закр.Xref";
                rbCloseXreFinPlace.ShowText = true;
                rbCloseXreFinPlace.ShowImage = true;
                rbCloseXreFinPlace.Image = Images.GetBitmap(Properties.Resources.CloseXREFinPlace);
                rbCloseXreFinPlace.CommandHandler = new RibbonCommandHandler();
                rbCloseXreFinPlace.ToolTip = rttCloseXreFinPlace;

                var pan4Row2 = new RibbonRowPanel();

                pan4Row2.Items.Add(rbEditXreFinPlace);
                pan4Row2.Items.Add(new RibbonRowBreak());
                pan4Row2.Items.Add(rbSaveXreFinPlace);
                pan4Row2.Items.Add(new RibbonRowBreak());
                pan4Row2.Items.Add(rbCloseXreFinPlace);

                /////////////////////////////////////////
                //          ПАНЕЛЬ ОБНОВЛЕНИЯ          //
                /////////////////////////////////////////

                // Кнопка для обновления внешних ссылок
                var rbXrefUpdate = new RibbonButton();
                var rttXrefUpdate = new RibbonToolTip
                {
                    Title = "Обновление внешних ссылок",
                    Content = "Команда выполняет обновление подгруженных в чертёж внешних ссылок. " + "\n" +
                              "Если объект ссылки расположен в TDMS, то он будет выгружен локально и обновлён." +
                              "Также выполняет обновление атрибутов, расположенных в штампах и рамках. " +
                              "Выполняется запрос данных из TDMS и происходит обновление в соответствующих полях. " +
                              "Обновление проиходит для следующих полей: Разработал, Проверил, Рук.Группы, Н.контроль, " +
                              "Гл. констр. проекта, ГИП, ГАП, Стадия, Лист, Листов, Шифр, Наименование объекта, Адрес объекта, Даты"
                };
                rbXrefUpdate.CommandParameter = rttXrefUpdate.Command = "_UPDATEXREFATTR";
                rbXrefUpdate.Text = "Обновить ссылки" + "\n" + "и атрибуты";
                rbXrefUpdate.ShowText = true;
                rbXrefUpdate.ShowImage = true;
                rbXrefUpdate.LargeImage = Images.GetBitmap(Properties.Resources.UpdateXREF);
                rbXrefUpdate.Size = RibbonItemSize.Large;
                rbXrefUpdate.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbXrefUpdate.CommandHandler = new RibbonCommandHandler();
                rbXrefUpdate.ToolTip = rttXrefUpdate;

                var pan4Button2 = new RibbonButton
                {
                    Text = "Button2",
                    ShowText = true,
                    ShowImage = true,
                    CommandHandler = new RibbonCommandHandler()
                };

                var pan4Button3 = new RibbonButton
                {
                    Text = "Button3",
                    ShowText = true,
                    ShowImage = true,
                    CommandHandler = new RibbonCommandHandler()
                };

                var pan4Ribcombobutton1 = new RibbonButton
                {
                    Text = "Button1",
                    ShowText = true,
                    ShowImage = true,
                    CommandHandler = new RibbonCommandHandler()
                };

                var pan4Ribcombobutton2 = new RibbonButton
                {
                    Text = "Button2",
                    ShowText = true,
                    ShowImage = true,
                    CommandHandler = new RibbonCommandHandler()
                };

                ///////////////////////////////
                //       ПАНЕЛЬ ПОМОЩЬ       //
                ///////////////////////////////

                // Кнопка для открытия документа с популярными вопросами и ответами
                var rbFaq = new RibbonButton();
                var rttFaq = new RibbonToolTip
                {
                    Title = "Часто задаваемые вопросы (FAQ - Frequently asked questions)",
                    Content =
                        "В предлагаемом документе отображены вопросы, на которые большинство ищет ответы. Обновление списка вопросов и ответов будет происходить через 1-2 недели, по мере накопления. Ответы в свободной форме, ссылки в интернет, рекомендации в отношении повышения компьютерной грамотности и прочее)"
                };
                rbFaq.CommandParameter = rttFaq.Command = "_Questions";
                rbFaq.Text = "Часто задаваемые" + "\n" + "вопросы";
                rbFaq.ShowText = true;
                rbFaq.ShowImage = true;
                rbFaq.Image = Images.GetBitmap(Properties.Resources.FAQ);
                rbFaq.LargeImage = Images.GetBitmap(Properties.Resources.FAQ);
                rbFaq.Size = RibbonItemSize.Large;
                rbFaq.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbFaq.CommandHandler = new RibbonCommandHandler();
                rbFaq.ToolTip = rttFaq;

                // Кнопка для открытия стандарта предприятия (СТП)
                var rbEnterpriseStandart = new RibbonButton();
                var rttEnterpriseStandart = new RibbonToolTip
                {
                    Title = "Стандарт предприятия (СТП)",
                    Content =
                        "Команда выполняет открытие действующего стандарта предприятия, последней редакции. Стандарт расположен по адресу (V:\\Help)."
                };
                rbEnterpriseStandart.CommandParameter = rttEnterpriseStandart.Command = "_OpenStandart";
                rbEnterpriseStandart.Text = "Стандарт" + "\n" + "предприятия";
                rbEnterpriseStandart.ShowText = true;
                rbEnterpriseStandart.ShowImage = true;
                rbEnterpriseStandart.Image = Images.GetBitmap(Properties.Resources.EnterpriseStandart);
                rbEnterpriseStandart.LargeImage = Images.GetBitmap(Properties.Resources.EnterpriseStandart);
                rbEnterpriseStandart.Size = RibbonItemSize.Large;
                rbEnterpriseStandart.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbEnterpriseStandart.CommandHandler = new RibbonCommandHandler();
                rbEnterpriseStandart.ToolTip = rttEnterpriseStandart;

                // Кнопка для открытия Хелпа с диска
                var rbHelp = new RibbonButton();
                var rttHelp = new RibbonToolTip
                {
                    Title = "Руководство пользователя",
                    Content =
                        "Команда выполняет открытие действующего руководства пользователя по TDMS и интерфейсам в AutoCAD и Word, последней редакции. Руководство пользователя расположено по адресу (D:\\_TDMSHELP\\TDMSHELP.chm)."
                };
                rbHelp.CommandParameter = rttHelp.Command = "_OpenManual";
                rbHelp.Text = "Руководство" + "\n" + "пользователя";
                rbHelp.ShowText = true;
                rbHelp.ShowImage = true;
                rbHelp.Image = Images.GetBitmap(Properties.Resources.Help);
                rbHelp.LargeImage = Images.GetBitmap(Properties.Resources.Help);
                rbHelp.Size = RibbonItemSize.Large;
                rbHelp.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbHelp.CommandHandler = new RibbonCommandHandler();
                rbHelp.ToolTip = rttHelp;

                // Кнопка для открытия папки с книгами и литературой по автокаду, ревиту...
                var rbLibrary = new RibbonButton();
                var rttLibrary = new RibbonToolTip
                {
                    Title = "Справка",
                    Content =
                        "Команда выполняет открытие каталога (V:\\Help) где хранятся учебные материалы и тренинги по работе с Revit и AutoCAD."
                };
                rbLibrary.CommandParameter = rttLibrary.Command = "_OPENHELP";
                rbLibrary.Text = "Библиотека";
                rbLibrary.ShowText = true;
                rbLibrary.ShowImage = true;
                rbLibrary.Image = Images.GetBitmap(Properties.Resources.Libary);
                rbLibrary.LargeImage = Images.GetBitmap(Properties.Resources.Libary);
                rbLibrary.Size = RibbonItemSize.Large;
                rbLibrary.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbLibrary.CommandHandler = new RibbonCommandHandler();
                rbLibrary.ToolTip = rttLibrary;

                pan7SplitButton.Items.Add(rbHelp);
                pan7SplitButton.Items.Add(rbLibrary);
                pan7SplitButton.Items.Add(rbEnterpriseStandart);
                pan7SplitButton.Items.Add(rbFaq);

                Pan3Ribcombo.Width = 150;
                Pan3Ribcombo.Items.Add(pan4Ribcombobutton1);
                Pan3Ribcombo.Items.Add(pan4Ribcombobutton2);
                Pan3Ribcombo.Current = pan4Ribcombobutton1;

                var vvorow1 = new RibbonRowPanel();

                vvorow1.Items.Add(pan4Button2);
                vvorow1.Items.Add(new RibbonRowBreak());
                vvorow1.Items.Add(pan4Button3);
                vvorow1.Items.Add(new RibbonRowBreak());
                vvorow1.Items.Add(Pan3Ribcombo);

                // Добавление кнопок на панели
                panelSaveFunction.Items.Add(new RibbonSeparator());
                panelSaveFunction.Items.Add(rbSave);
                panelSaveFunction.Items.Add(rbSaveAndClose);
                panelSaveFunction.Items.Add(new RibbonSeparator());
                panelSaveFunction.Items.Add(rbSaveAndCloseAll);

                panelSaveFunction.Items.Add(new RibbonSeparator());
                panelSaveFunction.Items.Add(pan1Row1);

                panelXref.Items.Add(rbDwgFromTdms);
                panelXref.Items.Add(rbXrefUpdate);
                panelXref.Items.Add(new RibbonSeparator());
                panelXref.Items.Add(pan4Row1);
                panelXref.Items.Add(new RibbonSeparator());
                panelXref.Items.Add(pan4Row2);

                panelStamps.Items.Add(pan2SplitButton);

                //панель для работы с листами и слоями
                panelList.Items.Add(rbExportLayout);
                panelList.Items.Add(new RibbonSeparator());
                panelList.Items.Add(pan8SplitButton);

                //панель со справкой
                panelHelp.Items.Add(rbUberPurge);
                panelHelp.Items.Add(new RibbonSeparator());
                panelHelp.Items.Add(pan7SplitButton);

                //панель с кнопками для закрытия
                panelClose.Items.Add(rbClose);
                panelClose.Items.Add(new RibbonSeparator());
                panelClose.Items.Add(rbCloseAll);
                
                tab.IsActive = true;
            }
        }

        private static RibbonPanelSource PanelSourceClose(RibbonTab tab)
        {
            var pan7Panel = new RibbonPanelSource { Title = "Функции завершения" };
            var panel7 = new RibbonPanel { Source = pan7Panel };
            tab.Panels.Add(panel7);
            return pan7Panel;
        }

        private static RibbonPanelSource PanelSourceHelp(RibbonTab tab)
        {
            var pan7Panel = new RibbonPanelSource {Title = "Помощь"};
            var panel7 = new RibbonPanel {Source = pan7Panel};
            tab.Panels.Add(panel7);
            return pan7Panel;
        }

        private static RibbonPanelSource PanelSourceList(RibbonTab tab)
        {
            var pan6Panel = new RibbonPanelSource {Title = "Листы"};
            var panel6 = new RibbonPanel {Source = pan6Panel};
            tab.Panels.Add(panel6);
            return pan6Panel;
        }

        private static RibbonPanelSource PanelSourceStamps(RibbonTab tab)
        {
            var pan4Panel = new RibbonPanelSource {Title = "Штампы"};
            var panel4 = new RibbonPanel {Source = pan4Panel};
            tab.Panels.Add(panel4);
            tab.Panels.Add(Frame());
            return pan4Panel;
        }

        private static RibbonPanelSource PanelSourceXref(RibbonTab tab)
        {
            var pan2Panel = new RibbonPanelSource {Title = "Внешние ссылки из TDMS"};
            var panel2 = new RibbonPanel {Source = pan2Panel};
            tab.Panels.Add(panel2);
            return pan2Panel;
        }

        private static RibbonPanelSource PanelSourceSaveFunction(RibbonTab tab)
        {
            var pan1Panel = new RibbonPanelSource {Title = "Функции сохранения"};
            var panel1 = new RibbonPanel {Source = pan1Panel};
            tab.Panels.Add(panel1);
            return pan1Panel;
        }

        private static bool FindRibbon(RibbonControl ribbonControl)
        {
            foreach (var fRibTab in ribbonControl.Tabs)
            {
                return (!String.Equals(fRibTab.Title, "TDMS", StringComparison.InvariantCultureIgnoreCase));
            }
            return true;
        }

        private static RibbonPanel Frame()
        {
            //var editor = Application.DocumentManager.MdiActiveDocument.Editor;

            var rttA4H = new RibbonToolTip();
            var rttA4V = new RibbonToolTip();

            var rttA4H3 = new RibbonToolTip();
            var rttA4H4 = new RibbonToolTip();

            var rttA3H = new RibbonToolTip();
            var rttA3V = new RibbonToolTip();

            var rttA3H3 = new RibbonToolTip();

            var rttA2H = new RibbonToolTip();
            var rttA2V = new RibbonToolTip();

            var rttA1H = new RibbonToolTip();
            var rttA1V = new RibbonToolTip();

            var rttA1Hl = new RibbonToolTip();
            var rttA1Vl = new RibbonToolTip();

            var rttA0H = new RibbonToolTip();
            var rttA0V = new RibbonToolTip();

            var rbA4H = new RibbonButton();
            var rbA4V = new RibbonButton();

            var rbA4H3 = new RibbonButton();
            var rbA4H4 = new RibbonButton();

            var rbA3H = new RibbonButton();
            var rbA3V = new RibbonButton();

            var rbA3H3 = new RibbonButton();

            var rbA2H = new RibbonButton();
            var rbA2V = new RibbonButton();

            var rbA1H = new RibbonButton();
            var rbA1V = new RibbonButton();

            var rbA1Hl = new RibbonButton();
            var rbA1Vl = new RibbonButton();

            var rbA0H = new RibbonButton();
            var rbA0V = new RibbonButton();

            var ribbPanelSource = new RibbonPanelSource { Title = "Рамки" };

            var ribbPanel = new RibbonPanel { Source = ribbPanelSource };

            var ribbSplitButton = new RibbonSplitButton
            {
                Text = "RibbonSplitButton",
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Size = RibbonItemSize.Large,
                ShowImage = true,
                ShowText = true,
                ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton,
                ResizeStyle = RibbonItemResizeStyles.NoResize,
                ListStyle = RibbonSplitButtonListStyle.List
            };
            // Стиль кнопки
            try
            {
                rttA4H.IsHelpEnabled = false;
                rttA4V.IsHelpEnabled = false;

                rttA4H3.IsHelpEnabled = false;
                rttA4H4.IsHelpEnabled = false;

                rttA3H.IsHelpEnabled = false;
                rttA3V.IsHelpEnabled = false;

                rttA3H3.IsHelpEnabled = false;

                rttA2H.IsHelpEnabled = false;
                rttA2V.IsHelpEnabled = false;

                rttA1H.IsHelpEnabled = false;
                rttA1V.IsHelpEnabled = false;

                rttA0H.IsHelpEnabled = false;
                rttA0V.IsHelpEnabled = false;

                rttA4H.Title = "Рамка по формату";
                rttA4H.Content = "Формат A4(210x297) горизонтальная";
                rbA4H.CommandParameter = rttA4H.Command = "_A4H";
                rbA4H.Name = "A4H(210x297)";
                rbA4H.CommandHandler = new RibbonCommandHandler();
                rbA4H.ShowText = true;
                rbA4H.ShowImage = true;
                rbA4H.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA4H.Size = RibbonItemSize.Large;
                rbA4H.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA4H.Text = "A4H" + "\n" + "(210x297)";
                rbA4H.ToolTip = rttA4H;

                rttA4V.Title = "Рамка по формату";
                rttA4V.Content = "Формат A4(297x210) вертикальная";
                rbA4V.CommandParameter = rttA4V.Command = "_A4V";
                rbA4V.Name = "A4V(297x210)";
                rbA4V.CommandHandler = new RibbonCommandHandler();
                rbA4V.ShowText = true;
                rbA4V.ShowImage = true;
                rbA4V.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA4V.Size = RibbonItemSize.Large;
                rbA4V.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA4V.Text = "A4V" + "\n" + "(297x210)";
                rbA4V.ToolTip = rttA4V;

                rttA4H3.Title = "Рамка по формату";
                rttA4H3.Content = "Формат A4(297x630) горизонтальная удлинённая";
                rbA4H3.CommandParameter = rttA4H3.Command = "_A4H3";
                rbA4H3.Name = "A4H3(297x630)";
                rbA4H3.CommandHandler = new RibbonCommandHandler();
                rbA4H3.ShowText = true;
                rbA4H3.ShowImage = true;
                rbA4H3.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA4H3.Size = RibbonItemSize.Large;
                rbA4H3.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA4H3.Text = "A4H3" + "\n" + "(297x630)";
                rbA4H3.ToolTip = rttA4H3;

                rttA4H4.Title = "Рамка по формату";
                rttA4H4.Content = "Формат A4(297x841) горизонтальная удлинённая";
                rbA4H4.CommandParameter = rttA4H4.Command = "_A4H4";
                rbA4H4.Name = "A4H4(297x841)";
                rbA4H4.CommandHandler = new RibbonCommandHandler();
                rbA4H4.ShowText = true;
                rbA4H4.ShowImage = true;
                rbA4H4.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA4H4.Size = RibbonItemSize.Large;
                rbA4H4.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA4H4.Text = "A4H4" + "\n" + "(297x841)";
                rbA4H4.ToolTip = rttA4H4;

                rttA3H.Title = "Рамка по формату";
                rttA3H.Content = "Формат A3(297x420) горизонтальная";
                rbA3H.CommandParameter = rttA3H.Command = "_A3H";
                rbA3H.Name = "A3H(297x420)";
                rbA3H.CommandHandler = new RibbonCommandHandler();
                rbA3H.ShowText = true;
                rbA3H.ShowImage = true;
                rbA3H.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA3H.Size = RibbonItemSize.Large;
                rbA3H.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA3H.Text = "A3H" + "\n" + "(297x420)";
                rbA3H.ToolTip = rttA3H;

                rttA3V.Title = "Рамка по формату";
                rttA3V.Content = "Формат A3(420x297) вертикальная";
                rbA3V.CommandParameter = rttA3V.Command = "_A3V";
                rbA3V.Name = "A3V(420x297)";
                rbA3V.CommandHandler = new RibbonCommandHandler();
                rbA3V.ShowText = true;
                rbA3V.ShowImage = true;
                rbA3V.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA3V.Size = RibbonItemSize.Large;
                rbA3V.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA3V.Text = "A3V" + "\n" + "(420x297)";
                rbA3V.ToolTip = rttA3V;

                rttA3H3.Title = "Рамка по формату";
                rttA3H3.Content = "Формат A3(420x891) горизонтальная удлинённая";
                rbA3H3.CommandParameter = rttA3H3.Command = "_A3H3";
                rbA3H3.Name = "A3H3(420x891)";
                rbA3H3.CommandHandler = new RibbonCommandHandler();
                rbA3H3.ShowText = true;
                rbA3H3.ShowImage = true;
                rbA3H3.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA3H3.Size = RibbonItemSize.Large;
                rbA3H3.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA3H3.Text = "A3H3" + "\n" + "(420x891)";
                rbA3H3.ToolTip = rttA3H3;

                rttA2H.Title = "Рамка по формату";
                rttA2H.Content = "Формат A2(420x594) горизонтальная";
                rbA2H.CommandParameter = rttA2H.Command = "_A2H";
                rbA2H.Name = "A2H(420x594)";
                rbA2H.CommandHandler = new RibbonCommandHandler();
                rbA2H.ShowText = true;
                rbA2H.ShowImage = true;
                rbA2H.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA2H.Size = RibbonItemSize.Large;
                rbA2H.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA2H.Text = "A2H" + "\n" + "(420x594)";
                rbA2H.ToolTip = rttA2H;

                rttA2V.Title = "Рамка по формату";
                rttA2V.Content = "Формат A2(594x420) вертикальная";
                rbA2V.CommandParameter = rttA2V.Command = "_A2V";
                rbA2V.Name = "A2V(594x420)";
                rbA2V.CommandHandler = new RibbonCommandHandler();
                rbA2V.ShowText = true;
                rbA2V.ShowImage = true;
                rbA2V.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA2V.Size = RibbonItemSize.Large;
                rbA2V.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA2V.Text = "A2V" + "\n" + "(594x420)";
                rbA2V.ToolTip = rttA2V;

                rttA1H.Title = "Рамка по формату";
                rttA1H.Content = "Формат A1(594x841) горизонтальная";
                rbA1H.CommandParameter = rttA1H.Command = "_A1H";
                rbA1H.Name = "A1H(594x841)";
                rbA1H.CommandHandler = new RibbonCommandHandler();
                rbA1H.ShowText = true;
                rbA1H.ShowImage = true;
                rbA1H.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA1H.Size = RibbonItemSize.Large;
                rbA1H.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA1H.Text = "A1H" + "\n" + "(594x841)";
                rbA1H.ToolTip = rttA1H;

                rttA1V.Title = "Рамка по формату";
                rttA1V.Content = "Формат A1(841x594) вертикальная";
                rbA1V.CommandParameter = rttA1V.Command = "_A1V";
                rbA1V.Name = "A1V(841x594)";
                rbA1V.CommandHandler = new RibbonCommandHandler();
                rbA1V.ShowText = true;
                rbA1V.ShowImage = true;
                rbA1V.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA1V.Size = RibbonItemSize.Large;
                rbA1V.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA1V.Text = "A1V" + "\n" + "(841x594)";
                rbA1V.ToolTip = rttA1V;

                rttA1Hl.Title = "Рамка по формату";
                rttA1Hl.Content = "Формат A1(594x1051) горизонтальная";
                rbA1Hl.CommandParameter = rttA1Hl.Command = "_A1HL";
                rbA1Hl.Name = "A1HL(594x1051)";
                rbA1Hl.CommandHandler = new RibbonCommandHandler();
                rbA1Hl.ShowText = true;
                rbA1Hl.ShowImage = true;
                rbA1Hl.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA1Hl.Size = RibbonItemSize.Large;
                rbA1Hl.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA1Hl.Text = "A1HL" + "\n" + "(594x1051)";
                rbA1Hl.ToolTip = rttA1Hl;

                rttA1Vl.Title = "Рамка по формату";
                rttA1Vl.Content = "Формат A1(1051x594) вертикальная";
                rbA1Vl.CommandParameter = rttA1V.Command = "_A1VL";
                rbA1Vl.Name = "A1VL(1051x594)";
                rbA1Vl.CommandHandler = new RibbonCommandHandler();
                rbA1Vl.ShowText = true;
                rbA1Vl.ShowImage = true;
                rbA1Vl.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA1Vl.Size = RibbonItemSize.Large;
                rbA1Vl.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA1Vl.Text = "A1VL" + "\n" + "(1051x594)";
                rbA1Vl.ToolTip = rttA1Vl;

                rttA0H.Title = "Рамка по формату";
                rttA0H.Content = "Формат A0(841x1189) горизонтальная";
                rbA0H.CommandParameter = rttA0H.Command = "_A0H";
                rbA0H.Name = "A0H(841x1189)";
                rbA0H.CommandHandler = new RibbonCommandHandler();
                rbA0H.ShowText = true;
                rbA0H.ShowImage = true;
                rbA0H.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA0H.Size = RibbonItemSize.Large;
                rbA0H.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA0H.Text = "A0H" + "\n" + "(841x1189)";
                rbA0H.ToolTip = rttA0H;

                rttA0V.Title = "Рамка по формату";
                rttA0V.Content = "Формат A0(1189x841) вертикальная";
                rbA0V.CommandParameter = rttA0V.Command = "_A0V";
                rbA0V.Name = "A0V(1189x841)";
                rbA0V.CommandHandler = new RibbonCommandHandler();
                rbA0V.ShowText = true;
                rbA0V.ShowImage = true;
                rbA0V.LargeImage = Images.GetBitmap(Properties.Resources.Frame);
                rbA0V.Size = RibbonItemSize.Large;
                rbA0V.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA0V.Text = "A0V" + "\n" + "(1189x841)";
                rbA0V.ToolTip = rttA0V;
            }
            catch (System.Exception)
            {
                // ignored
            }

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButton.Items.Add(rbA4H);
            ribbSplitButton.Items.Add(rbA4V);

            ribbSplitButton.Items.Add(rbA4H3);
            ribbSplitButton.Items.Add(rbA4H4);

            ribbSplitButton.Items.Add(rbA3H);
            ribbSplitButton.Items.Add(rbA3V);

            ribbSplitButton.Items.Add(rbA3H3);

            ribbSplitButton.Items.Add(rbA2H);
            ribbSplitButton.Items.Add(rbA2V);

            ribbSplitButton.Items.Add(rbA1H);
            ribbSplitButton.Items.Add(rbA1V);

            ribbSplitButton.Items.Add(rbA1Hl);
            ribbSplitButton.Items.Add(rbA1Vl);

            ribbSplitButton.Items.Add(rbA0H);
            ribbSplitButton.Items.Add(rbA0V);

            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButton);

            var rttA4Ha = new RibbonToolTip();
            var rttA4Va = new RibbonToolTip();

            var rttA3Ha = new RibbonToolTip();
            var rttA3Va = new RibbonToolTip();

            var rttA2Ha = new RibbonToolTip();
            var rttA2Va = new RibbonToolTip();

            var rttA1Ha = new RibbonToolTip();
            var rttA1Va = new RibbonToolTip();

            var rttA0Ha = new RibbonToolTip();
            var rttA0Va = new RibbonToolTip();

            var rbA4Ha = new RibbonButton();
            var rbA4Va = new RibbonButton();

            var rbA3Ha = new RibbonButton();
            var rbA3Va = new RibbonButton();

            var rbA2Ha = new RibbonButton();
            var rbA2Va = new RibbonButton();

            var rbA1Ha = new RibbonButton();
            var rbA1Va = new RibbonButton();

            var rbA0Ha = new RibbonButton();
            var rbA0Va = new RibbonButton();

            var ribbSplitButtonArch = new RibbonSplitButton
            {
                Text = "RibbonSplitButton",
                Orientation = System.Windows.Controls.Orientation.Vertical,
                Size = RibbonItemSize.Large,
                ShowImage = true,
                ShowText = true,
                ListButtonStyle = Autodesk.Private.Windows.RibbonListButtonStyle.SplitButton,
                ResizeStyle = RibbonItemResizeStyles.NoResize,
                ListStyle = RibbonSplitButtonListStyle.List
            };

            // Стиль кнопки

            try
            {
                rttA4Ha.IsHelpEnabled = false;
                rttA4Va.IsHelpEnabled = false;

                rttA3Ha.IsHelpEnabled = false;
                rttA3Va.IsHelpEnabled = false;

                rttA2Ha.IsHelpEnabled = false;
                rttA2Va.IsHelpEnabled = false;

                rttA1Ha.IsHelpEnabled = false;
                rttA1Va.IsHelpEnabled = false;

                rttA0Ha.IsHelpEnabled = false;
                rttA0Va.IsHelpEnabled = false;

                ///////////////////////////////////////////////////////

                rttA4Ha.Title = "Рамка для оформления";
                rttA4Ha.Content = "Рамка для оформления формата A4(210x297) горизонтальная";
                rbA4Ha.CommandParameter = rttA4Ha.Command = "_A4HA";
                rbA4Ha.Name = "A4HA(210x297)";
                rbA4Ha.CommandHandler = new RibbonCommandHandler();
                rbA4Ha.ShowText = true;
                rbA4Ha.ShowImage = true;
                rbA4Ha.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA4Ha.Size = RibbonItemSize.Large;
                rbA4Ha.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA4Ha.Text = "A4HA" + "\n" + "(210x297)";
                rbA4Ha.ToolTip = rttA4Ha;

                rttA4Ha.Title = "Рамка для оформления";
                rttA4Ha.Content = "Рамка для оформления формата A4(297x210) вертикальная";
                rbA4Va.CommandParameter = rttA4Va.Command = "_A4VA";
                rbA4Va.Name = "A4VA(297x210)";
                rbA4Va.CommandHandler = new RibbonCommandHandler();
                rbA4Va.ShowText = true;
                rbA4Va.ShowImage = true;
                rbA4Va.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA4Va.Size = RibbonItemSize.Large;
                rbA4Va.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA4Va.Text = "A4VA" + "\n" + "(297x210)";
                rbA4Va.ToolTip = rttA4Va;

                ///////////////////////////////////////////////////////

                rttA3Ha.Title = "Рамка для оформления";
                rttA3Ha.Content = "Рамка для оформления формата A3(297x420) горизонтальная";
                rbA3Ha.CommandParameter = rttA3Ha.Command = "_A3HA";
                rbA3Ha.Name = "A3HA(297x420)";
                rbA3Ha.CommandHandler = new RibbonCommandHandler();
                rbA3Ha.ShowText = true;
                rbA3Ha.ShowImage = true;
                rbA3Ha.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA3Ha.Size = RibbonItemSize.Large;
                rbA3Ha.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA3Ha.Text = "A3HA" + "\n" + "(297x420)";
                rbA3Ha.ToolTip = rttA3Ha;

                rttA3Va.Title = "Рамка для оформления";
                rttA3Va.Content = "Рамка для оформления формата A3(420x297) вертикальная";
                rbA3Va.CommandParameter = rttA3Va.Command = "_A3VA";
                rbA3Va.Name = "A3VA(420x297)";
                rbA3Va.CommandHandler = new RibbonCommandHandler();
                rbA3Va.ShowText = true;
                rbA3Va.ShowImage = true;
                rbA3Va.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA3Va.Size = RibbonItemSize.Large;
                rbA3Va.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA3Va.Text = "A3VA" + "\n" + "(420x297)";
                rbA3Va.ToolTip = rttA3Va;

                ///////////////////////////////////////////////////////

                rttA2Ha.Title = "Рамка для оформления";
                rttA2Ha.Content = "Рамка для оформления формата A2(420x594) горизонтальная";
                rbA2Ha.CommandParameter = rttA2Ha.Command = "_A2HA";
                rbA2Ha.Name = "A2HA(420x594)";
                rbA2Ha.CommandHandler = new RibbonCommandHandler();
                rbA2Ha.ShowText = true;
                rbA2Ha.ShowImage = true;
                rbA2Ha.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA2Ha.Size = RibbonItemSize.Large;
                rbA2Ha.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA2Ha.Text = "A2HA" + "\n" + "(420x594)";
                rbA2Ha.ToolTip = rttA2Ha;

                rttA2Va.Title = "Рамка для оформления";
                rttA2Va.Content = "Рамка для оформления формата A2(594x420) вертикальная";
                rbA2Va.CommandParameter = rttA2Va.Command = "_A2VA";
                rbA2Va.Name = "A2VA(594x420)";
                rbA2Va.CommandHandler = new RibbonCommandHandler();
                rbA2Va.ShowText = true;
                rbA2Va.ShowImage = true;
                rbA2Va.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA2Va.Size = RibbonItemSize.Large;
                rbA2Va.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA2Va.Text = "A2VA" + "\n" + "(594x420)";
                rbA2Va.ToolTip = rttA2Va;

                ///////////////////////////////////////////////////////

                rttA1Ha.Title = "Рамка для оформления";
                rttA1Ha.Content = "Рамка для оформления формата A1(594x841) горизонтальная";
                rbA1Ha.CommandParameter = rttA1Ha.Command = "_A1HA";
                rbA1Ha.Name = "A1HA(594x841)";
                rbA1Ha.CommandHandler = new RibbonCommandHandler();
                rbA1Ha.ShowText = true;
                rbA1Ha.ShowImage = true;
                rbA1Ha.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA1Ha.Size = RibbonItemSize.Large;
                rbA1Ha.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA1Ha.Text = "A1HA" + "\n" + "(594x841)";
                rbA1Ha.ToolTip = rttA1Ha;

                rttA1Va.Title = "Рамка для оформления";
                rttA1Va.Content = "Рамка для оформления формата A1(841x594) вертикальная";
                rbA1Va.CommandParameter = rttA1Va.Command = "_A1VA";
                rbA1Va.Name = "A1VA(841x594)";
                rbA1Va.CommandHandler = new RibbonCommandHandler();
                rbA1Va.ShowText = true;
                rbA1Va.ShowImage = true;
                rbA1Va.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA1Va.Size = RibbonItemSize.Large;
                rbA1Va.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA1Va.Text = "A1VA" + "\n" + "(841x594)";
                rbA1Va.ToolTip = rttA1Va;

                ///////////////////////////////////////////////////////

                rttA0Ha.Title = "Рамка для оформления";
                rttA0Ha.Content = "Рамка для оформления формата A0(841x1189) горизонтальная";
                rbA0Ha.CommandParameter = rttA0Ha.Command = "_A0HA";
                rbA0Ha.Name = "A0HA(841x1189)";
                rbA0Ha.CommandHandler = new RibbonCommandHandler();
                rbA0Ha.ShowText = true;
                rbA0Ha.ShowImage = true;
                rbA0Ha.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA0Ha.Size = RibbonItemSize.Large;
                rbA0Ha.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA0Ha.Text = "A0HA" + "\n" + "(841x1189)";
                rbA0Ha.ToolTip = rttA0Ha;

                rttA0Va.Title = "Рамка для оформления";
                rttA0Va.Content = "Рамка для оформления формата A0(1189x841) вертикальная";
                rbA0Va.CommandParameter = rttA0Va.Command = "_A0VA";
                rbA0Va.Name = "A0VA(1189x841)";
                rbA0Va.CommandHandler = new RibbonCommandHandler();
                rbA0Va.ShowText = true;
                rbA0Va.ShowImage = true;
                rbA0Va.LargeImage = Images.GetBitmap(Properties.Resources.FrameArch);
                rbA0Va.Size = RibbonItemSize.Large;
                rbA0Va.Orientation = System.Windows.Controls.Orientation.Vertical;
                rbA0Va.Text = "A0VA" + "\n" + "(1189x841)";
                rbA0Va.ToolTip = rttA0Va;
            }
            catch (System.Exception)
            {
                // ignored
            }
            ///////////////////////////////////////////////////////

            //Добавляем кнопку на раскрывающуюся панель
            ribbSplitButtonArch.Items.Add(rbA4Ha);
            ribbSplitButtonArch.Items.Add(rbA4Va);

            ribbSplitButtonArch.Items.Add(rbA3Ha);
            ribbSplitButtonArch.Items.Add(rbA3Va);

            ribbSplitButtonArch.Items.Add(rbA2Ha);
            ribbSplitButtonArch.Items.Add(rbA2Va);

            ribbSplitButtonArch.Items.Add(rbA1Ha);
            ribbSplitButtonArch.Items.Add(rbA1Va);

            ribbSplitButtonArch.Items.Add(rbA0Ha);
            ribbSplitButtonArch.Items.Add(rbA0Va);

            //Добавляем раскрывающуюся панель на панель вкладки
            ribbPanelSource.Items.Add(ribbSplitButtonArch);
            return ribbPanel;
        }

        public class RibbonCommandHandler : System.Windows.Input.ICommand
        {
            public bool CanExecute(object parameter)
            {
                return true;
            }

            public event EventHandler CanExecuteChanged;

            public void Execute(object parameter)
            {
                var editor = Application.DocumentManager.MdiActiveDocument.Editor;
                try
                {
                    var ribbonButton = parameter as RibbonButton;
                    if (ribbonButton == null) return;
                    var button = ribbonButton;
                    Application.DocumentManager.MdiActiveDocument.SendStringToExecute(button.CommandParameter + " ",
                        true, false, true);
                }
                catch (System.Exception ex)
                {
                    editor.WriteMessage("\n Exception caught: " + ex.Message + "\n" + ex.StackTrace);
                }
            }
        }

        public class Images
        {
            public static BitmapImage GetBitmap(Bitmap image)
            {
                var stream = new MemoryStream();
                image.Save(stream, ImageFormat.Png);
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.StreamSource = stream;
                bmp.EndInit();

                return bmp;
            }
        }
    }

    public class CreateRibbon : IExtensionApplication
    {
        public void Initialize()
        {
            try
            {
                RibbonHelper.OnRibbonFound(SetupRibbon);
            }
            catch (System.Exception ex)
            {
                var doc = Application.DocumentManager.MdiActiveDocument;
                doc?.Editor.WriteMessage(ex.ToString());
            }
        }

        public void Terminate()
        {
        }

        private void SetupRibbon(RibbonControl ribbon)
        {
            var ribbonCreate = new Ribbon();
            var addEve = new Xref();

            ribbonCreate.MyRibbon();

            addEve.AddEvent();
        }
    }

    public class RibbonHelper
    {
        private Action<RibbonControl> _action;
        private bool _idleHandled;
        private bool _created;

        private RibbonHelper(Action<RibbonControl> action)
        {
            if (action == null)
                throw new ArgumentNullException(nameof(action));
            _action = action;
            SetIdle(true);
        }

        public static void OnRibbonFound(Action<RibbonControl> action)
        {
            new RibbonHelper(action);
        }

        private void SetIdle(bool value)
        {
            if (!(value ^ _idleHandled)) return;
            if (value)
                Application.Idle += Idle;
            else
                Application.Idle -= Idle;
            _idleHandled = value;
        }

        private void Idle(object sender, EventArgs e)
        {
            SetIdle(false);
            if (_action == null) return;
            var ps = RibbonServices.RibbonPaletteSet;
            if (ps != null)
            {
                var ribbon = ps.RibbonControl;

                if (ribbon == null) return;
                _action(ribbon);
                _action = null;
            }
            else if (!_created)
            {
                _created = true;
                RibbonServices.RibbonPaletteSetCreated +=
                    RibbonPaletteSetCreated;
            }
        }

        private void RibbonPaletteSetCreated(object sender, EventArgs e)
        {
            RibbonServices.RibbonPaletteSetCreated
                -= RibbonPaletteSetCreated;
            SetIdle(true);
        }
    }
}