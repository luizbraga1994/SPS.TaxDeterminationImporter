using SAPbouiCOM;
using SBO.Hub.Forms;
using AAC.TaxDeterminationImporter.Core.Model;
using System.Collections.Generic;

namespace AAC.TaxDeterminationImporter.Core.Forms
{
    public class FrmImportLog : BaseForm
    {
        Form Form;

        public void Show(List<TaxDeterminationModel> list)
        {
            Form = (Form)base.Show();

            Form.Freeze(true);
            DataTable dt_Log = Form.DataSources.DataTables.Item("dt_Log");

            dt_Log.Rows.Add(list.Count);
            int i = 0;
            foreach (var item in list)
            {
                dt_Log.SetValue("Linha Planilha", i, item.Line);
                dt_Log.SetValue("Detalhes do Erro", i, item.Error);
                i++;
            }

            Grid gr_Log = Form.Items.Item("gr_Log").Specific as Grid;
            gr_Log.AutoResizeColumns();

            Form.Freeze(false);
        }
    }
}
