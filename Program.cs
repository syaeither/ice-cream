using System;
using System.Windows.Forms;

namespace ice_cream
{
    internal static class Program
    {
        public static string ConnectionString { get; } = "Host=localhost;Port=5432;Username=postgres;Password=998877fff;Database=ice_cream";

        //public static string ConnectionString { get; } = "Host=shinkansen.proxy.rlwy.net;Port=14484;Username=postgres;Password=tzlSIuwBvtOGXksCKjimNfuayjsNDuwu;Database=railway;SSL Mode=Require;Trust Server Certificate=true;";

        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Authorization());
        }
    }
}
