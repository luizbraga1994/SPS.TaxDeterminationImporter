using SBO.Hub;
using System.Resources;

namespace SPS.TaxDeterminationImporter.Core.DAO
{
    public class Scripts
    {
        public static ResourceManager Resource;

        public static void SetResourceManager()
        {
            if (SBOApp.Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                Resource = new ResourceManager("SPS.TaxDeterminationImporter.Core.DAO.Hana", typeof(Hana).Assembly);
            }
            else
            {
                Resource = new ResourceManager("SPS.TaxDeterminationImporter.Core.DAO.SQL", typeof(SQL).Assembly);
            }
        }
    }
}
