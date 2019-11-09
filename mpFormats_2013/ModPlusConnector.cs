#pragma warning disable SA1600 // Elements should be documented
namespace mpFormats
{
    using System;
    using System.Collections.Generic;
    using ModPlusAPI.Interfaces;

    public class ModPlusConnector : IModPlusFunctionInterface
    {
        private static ModPlusConnector _instance;

        public static ModPlusConnector Instance => _instance ?? (_instance = new ModPlusConnector());

        public SupportedProduct SupportedProduct => SupportedProduct.AutoCAD;
        
        public string Name => "mpFormats";

#if A2013
        public string AvailProductExternalVersion => "2013";
#elif A2014
        public string AvailProductExternalVersion => "2014";
#elif A2015
        public string AvailProductExternalVersion => "2015";
#elif A2016
        public string AvailProductExternalVersion => "2016";
#elif A2017
        public string AvailProductExternalVersion => "2017";
#elif A2018
        public string AvailProductExternalVersion => "2018";
#elif A2019
        public string AvailProductExternalVersion => "2019";
#elif A2020
        public string AvailProductExternalVersion => "2020";
#endif

        public string FullClassName => string.Empty;
        
        public string AppFullClassName => string.Empty;
        
        public Guid AddInId => Guid.Empty;
        
        public string LName => "Форматки";
        
        public string Description => "Плагин вставки в чертеж форматок по ГОСТ 2.301-68/ISO 216";
        
        public string Author => "Пекшев Александр aka Modis";
        
        public string Price => "0";
        
        public bool CanAddToRibbon => true;
        
        public string FullDescription => "Форматки вставляются в виде блока со специальной меткой в расширенных данных. Метка в расширенных данных требуется для определения форматок другими функциями (например, \"Нумератор листов\"). Поэтому не следует вставлять форматки стандартными средствами вставки блока в AutoCAD. Плагин позволяет добавлять к форматке штампы – для этого должна быть установлена и хотя бы раз запущена функция \"Штампы\"";
        
        public string ToolTipHelpImage => string.Empty;
        
        public List<string> SubFunctionsNames => new List<string>();
        
        public List<string> SubFunctionsLames => new List<string>();
        
        public List<string> SubDescriptions => new List<string>();
        
        public List<string> SubFullDescriptions => new List<string>();
        
        public List<string> SubHelpImages => new List<string>();
        
        public List<string> SubClassNames => new List<string>();
    }
}
#pragma warning restore SA1600 // Elements should be documented