using System.Diagnostics;

namespace Auto
{
    using Microsoft.Win32;
    using System;
    using System.IO;
    using System.Windows.Forms;

    public partial class Property : Form
    {
        private int _updateXref = 1;
        private int _updateAttr;
        private int _updateScheduleTable;
        private int _searchChangeXref;

        private int _checkUpdateXref;
        private int _checkUpdateAttr;
        private int _checkUpdateScheduleTable;
        private int _checkSearchChangeXref;

        public int GetCheckUpdateXref()
        {
            return _checkUpdateXref;
        }

        public int GetCheckUpdateAttr()
        {
            return _checkUpdateAttr;
        }

        public int GetCheckScheduleTable()
        {
            return _checkUpdateScheduleTable;
        }

        public int GetCheckSearchChangeXref()
        {
            return _checkSearchChangeXref;
        }

        public Property()
        {
            InitializeComponent();
            GetPropertiesTdms();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists(Path.GetFullPath(tbTracePath.Text)))
                {
                    TraceSource.CreateTraceFile();
                    TraceSource.TraceMessage(TraceType.Information, "Задан новый путь к файлу логов.");
                }
            }
            catch (Exception ex)
            {
                if (ex is DirectoryNotFoundException)
                {
                    TraceSource.TraceMessage(TraceType.Information, "Директория не найдена.");
                }
            }

            RegisterProperty();

            SetPropertiesTdms();

            ActiveForm?.Close();
        }

        private void RegisterProperty()
        {
            try
            {
                _updateXref = chboxUpdateXref.Checked ? 1 : 0;

                _updateAttr = chboxUpdateAttr.Checked ? 1 : 0;

                _searchChangeXref = chboxSearchChangeXref.Checked ? 1 : 0;

                _updateScheduleTable = chboxUpdateScheduleTable.Checked ? 1 : 0;
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void SetPropertiesTdms()
        {
            try
            {
                // Из реестра получаем ключ AutoCAD
                var sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
                const string sAppName = "TDMSProperties";

                using (var regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey))
                {
                    using (var regAcadAppKey = regAcadProdKey?.OpenSubKey("Applications", true))
                    {
                        // Регистрируем изменения
                        using (var regAppAddInKey = regAcadAppKey?.CreateSubKey(sAppName))
                        {
                            regAppAddInKey?.SetValue("DESCRIPTION", sAppName, RegistryValueKind.String);
                            regAppAddInKey?.SetValue("UpdateXref", _updateXref, RegistryValueKind.DWord);
                            regAppAddInKey?.SetValue("UpdateAttr", _updateAttr, RegistryValueKind.DWord);
                            regAppAddInKey?.SetValue("UpdateScheduleTable", _updateScheduleTable, RegistryValueKind.DWord);
                            regAppAddInKey?.SetValue("SearchChangeXref", _searchChangeXref, RegistryValueKind.DWord);
                            regAcadAppKey?.Close();
                        }
                    }
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void GetPropertiesTdms()
        {
            try
            {
                // Из реестра получаем ключ AutoCAD
                var sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;
                const string sAppName = "TDMSProperties";

                using (var regAcadProdKey = Registry.CurrentUser.OpenSubKey(sProdKey))
                {
                    Debug.Assert(regAcadProdKey != null, "regAcadProdKey != null");
                    using (var regAcadAppKey = regAcadProdKey.OpenSubKey("Applications", true))
                    {
                        Debug.Assert(regAcadAppKey != null, "regAcadAppKey != null");
                        using (var regAppAddInKey = regAcadAppKey.OpenSubKey(sAppName))
                        {
                            Debug.Assert(regAppAddInKey != null, "regAppAddInKey != null");
                            _checkUpdateXref = (int)regAppAddInKey.GetValue("UpdateXref");
                            _checkUpdateAttr = (int)regAppAddInKey.GetValue("UpdateAttr");
                            _checkUpdateScheduleTable = (int)regAppAddInKey.GetValue("UpdateScheduleTable");
                            _checkSearchChangeXref = (int)regAppAddInKey.GetValue("SearchChangeXref");
                            regAcadAppKey.Close();
                        }
                    }
                }

                if (_checkUpdateXref == 1)
                {
                    chboxUpdateXref.Checked = true;
                }
                if (_checkUpdateAttr == 1)
                {
                    chboxUpdateAttr.Checked = true;
                }
                if (_checkUpdateScheduleTable == 1)
                {
                    chboxUpdateScheduleTable.Checked = true;
                }
                if (_checkSearchChangeXref == 1)
                {
                    chboxSearchChangeXref.Checked = true;
                }
            }
            catch (Exception)
            {
                // ignored
            }
        }

        private void btnSetPath_Click(object sender, EventArgs e)
        {
            var sfd = new Autodesk.AutoCAD.Windows.SaveFileDialog("Сохранить файл логов на локальном диске",
                                                                                                       Path.GetFileNameWithoutExtension(TraceSource.LogFileFullPath),
                                                                                                       "log",
                                                                                                       "saving file",
                                                                                                       Autodesk.AutoCAD.Windows.SaveFileDialog.SaveFileDialogFlags.DefaultIsFolder);
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                tbTracePath.Text = TraceSource.LogFileFullPath = sfd.Filename;
            }
            else
            {
                tbTracePath.Text = TraceSource.LogFileFullPath;
            }
        }
    }
}