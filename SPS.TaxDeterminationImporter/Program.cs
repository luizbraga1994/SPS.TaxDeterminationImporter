using SBO.Hub;
using SBO.Hub.Services;
using SPS.TaxDeterminationImporter.Core.BLL;
using System.Threading;
using System.Windows.Forms;

namespace SPS.TaxDeterminationImporter
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Application.Exit();
                return;
            }

            SBOApp sboApp = new SBOApp(args[0], $"{Application.StartupPath}\\SPS.TaxDeterminationImporter.Core.dll");

            sboApp.InitializeApplication();

            InitializeBLL.Initialize();

            var oListener = new Listener();
            var oThread = new Thread(oListener.startListener) { IsBackground = true };
            oThread.Start();

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run();
        }
    }
}
