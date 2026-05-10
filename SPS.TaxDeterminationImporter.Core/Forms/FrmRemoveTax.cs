using SAPbouiCOM;
using SBO.Hub;
using SBO.Hub.Forms;
using SPS.TaxDeterminationImporter.Core.BLL;
using SPS.TaxDeterminationImporter.Core.DAO;
using SPS.TaxDeterminationImporter.Core.Model;
using System;
using System.Collections.Generic;
namespace SPS.TaxDeterminationImporter.Core.Forms
{
    public class FrmRemoveTax : BaseForm
    {
        public FrmRemoveTax()
        {

        }

        public FrmRemoveTax(ItemEvent itemEvent)
        {
            ItemEventInfo = itemEvent;
        }

        Form Form;
        private static int LineKey;

        public void Show(int lineKey, string field1, string field2, string field3, string field4, string field5, string description)
        {
            LineKey = lineKey;
            Form = (Form)base.Show();

            ((EditText)Form.Items.Item("et_Desc").Specific).Value = description;

            if (String.IsNullOrEmpty(field2))
            {
                field2 = "Campo-chave 2";
            }
            if (String.IsNullOrEmpty(field3))
            {
                field3 = "Campo-chave 3";
            }
            if (String.IsNullOrEmpty(field4))
            {
                field4 = "Campo-chave 4";
            }
            if (String.IsNullOrEmpty(field5))
            {
                field5 = "Campo-chave 5";
            }

            Form.Freeze(true);
            DataTable dt_Remove = Form.DataSources.DataTables.Item("dt_Remove");
            string sql = String.Format(Scripts.Resource.GetString("TaxDetermination_GetFields"), lineKey, field1, field2, field3, field4, field5);
            dt_Remove.ExecuteQuery(sql);

            Grid gr_Remove = Form.Items.Item("gr_Remove").Specific as Grid;
            gr_Remove.Columns.Item("#").Type = BoGridColumnType.gct_CheckBox;
            gr_Remove.AutoResizeColumns();

            for (int i = 1; i < gr_Remove.Columns.Count; i++)
            {
                gr_Remove.Columns.Item(i).Editable = false;
            }

            Form.Freeze(false);
        }

        public override bool ItemEvent()
        {
            if (!ItemEventInfo.BeforeAction)
            {
                Form = SBOApp.Application.Forms.GetForm(ItemEventInfo.FormTypeEx, ItemEventInfo.FormTypeCount);
                if (ItemEventInfo.EventType == BoEventTypes.et_CLICK)
                {
                    if (ItemEventInfo.ItemUID == "bt_Remove")
                    {
                        this.Remove();
                    }
                }
                if (ItemEventInfo.EventType == BoEventTypes.et_DOUBLE_CLICK)
                {
                    if (ItemEventInfo.Row == -1 && ItemEventInfo.ColUID == "#")
                    {
                        this.SelectAllNone();
                    }
                }
            }

            return true;
        }

        private void SelectAllNone()
        {
            Form.Freeze(true);
            DataTable dt_Remove = Form.DataSources.DataTables.Item("dt_Remove");
            if (dt_Remove.Rows.Count > 0)
            {
                string selected = dt_Remove.GetValue("#", 0).ToString() == "Y" ? "N" : "Y";

                for (int i = 0; i < dt_Remove.Rows.Count; i++)
                {
                    dt_Remove.SetValue("#", i, selected);
                }
            }
            Form.Freeze(false);
        }
        
        private void Remove()
        {
            List<int> selectedLines = new List<int>();
            DataTable dt_Remove = Form.DataSources.DataTables.Item("dt_Remove");
            for (int i = 0; i < dt_Remove.Rows.Count; i++)
            {
                if (dt_Remove.GetValue("#", i).ToString() == "Y")
                {
                    selectedLines.Add(i);
                }
            }

            if (selectedLines.Count == 0)
            {
                SBOApp.Application.SetStatusBarMessage("Nenhuma linha selecionada");
                return;
            }

            TaxDeterminationBLL taxDeterminationBLL = new TaxDeterminationBLL();
            string error = taxDeterminationBLL.RemoveLines(LineKey, selectedLines);
            if (String.IsNullOrEmpty(error))
            {
                SBOApp.Application.MessageBox("Linhas removidas com sucesso");
                Form.Close();

                SBOApp.Application.Forms.ActiveForm.Close();
                SBOApp.Application.ActivateMenuItem("8497");
            }
            else
            {
                SBOApp.Application.MessageBox(error);
            }
        }
    }
}
