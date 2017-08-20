using mpPInterface;

namespace mpFormats
{
    public class Interface : IPluginInterface
    {
        public string Name => "mpFormats";
        public string AvailCad => "2017";
        public string LName => "Форматки";
        public string Description => "Функция вставки в чертеж форматок по ГОСТ 2.301-68";
        public string Author => "Пекшев Александр aka Modis";
        public string Price => "0";
    }
    public class VersionData
    {
        public const string FuncVersion = "2017";
    }
}
