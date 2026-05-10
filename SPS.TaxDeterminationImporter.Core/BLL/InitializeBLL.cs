using SPS.TaxDeterminationImporter.Core.DAO;

namespace SPS.TaxDeterminationImporter.Core.BLL
{
    public class InitializeBLL
    {
        public static void Initialize()
        {
            Scripts.SetResourceManager();
            EventFilterBLL.SetDefaultEvents();
        }
    }
}
