using SBO.Hub;
using System.Resources;

namespace AAC.TaxDeterminationImporter.Core.DAO
{
    public class Scripts
    {
        public static ResourceManager Resource;

        public static void SetResourceManager()
        {
            if (SBOApp.Company.DbServerType == SAPbobsCOM.BoDataServerTypes.dst_HANADB)
            {
                Resource = new ResourceManager("AAC.TaxDeterminationImporter.Core.DAO.Hana", typeof(Hana).Assembly);
            }
            else
            {
                Resource = new ResourceManager("AAC.TaxDeterminationImporter.Core.DAO.SQL", typeof(SQL).Assembly);
            }
        }
    }
}
