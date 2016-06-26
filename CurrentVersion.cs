namespace Auto
{
    /// <summary>
    /// Методы данного класса обращаются к ветке реестра Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey
    /// и возвращают соответствующие значения локализации\языка в AutoCAD и используемый стиль для команды Package.
    /// </summary>
    public sealed class CurrentVersion
    {
        private string sProdKey = Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey;

        private void Initialize()
        {
        }

        private void Terminate()
        {
        }

        /// <summary>
        /// Метод возвращает Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey + \\LanguagePack
        /// </summary>
        /// <returns></returns>
        public string pathLanguage() { return (this.sProdKey + "\\LanguagePack"); }

        /// <summary>
        /// Метод возвращает Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.UserRegistryProductRootKey + \\ETransmit\\setups\\Enterprise
        /// </summary>
        /// <returns></returns>
        public string pathEtransmit() { return (this.sProdKey + "\\ETransmit\\setups\\Enterprise"); }
    }
}