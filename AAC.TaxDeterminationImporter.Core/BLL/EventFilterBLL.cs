using SAPbouiCOM;
using SBO.Hub.Helpers;

namespace AAC.TaxDeterminationImporter.Core.BLL
{
    class EventFilterBLL
    {
        public static void SetDefaultEvents()
        {
            EventFilterHelper.SetFormEvent("80401", BoEventTypes.et_FORM_LOAD);
            EventFilterHelper.SetFormEvent("80401", BoEventTypes.et_CLICK);

            EventFilterHelper.SetFormEvent("FrmRemoveTax", BoEventTypes.et_CLICK);
            EventFilterHelper.SetFormEvent("FrmRemoveTax", BoEventTypes.et_DOUBLE_CLICK);

            //EventFilterHelper.SetFormEvent("80402", BoEventTypes.et_FORM_LOAD);
            //EventFilterHelper.SetFormEvent("80402", BoEventTypes.et_CLICK);
            EventFilterHelper.EnableEvents();
        }
    }
}
