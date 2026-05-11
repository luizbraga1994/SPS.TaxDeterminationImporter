using AAC.TaxDeterminationImporter.Core.DAO;

namespace AAC.TaxDeterminationImporter.Core.BLL
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
