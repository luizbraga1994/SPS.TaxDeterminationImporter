using SAPbouiCOM;
using SBO.Hub;
using SBO.Hub.Forms;
using SBO.Hub.Util;
using SPS.TaxDeterminationImporter.Core.BLL;
using SPS.TaxDeterminationImporter.Core.Model;
using System;
using System.Collections.Generic;

namespace SPS.TaxDeterminationImporter.Core.Forms
{
    /// <summary>
    /// Determinação de Código de Imposto - Valores campos chave
    /// </summary>
    public class f80402 : SystemForm
    {
        public f80402(ItemEvent itemEvent)
        {
            ItemEventInfo = itemEvent;
        }

        public override bool ItemEvent()
        {
            if (!ItemEventInfo.BeforeAction)
            {
                if (ItemEventInfo.EventType == BoEventTypes.et_FORM_LOAD)
                {
                    var itemBase = Form.Items.Item("2002");
                    var it_Import = Form.Items.Add("bt_Import", BoFormItemTypes.it_BUTTON);
                    it_Import.LinkTo = "2002";
                    it_Import.Top = itemBase.Top;
                    it_Import.Width = 150;
                    it_Import.Left = itemBase.Left - 160;
                    it_Import.Height = itemBase.Height;

                    Button bt_Import = (Button)it_Import.Specific;
                    bt_Import.Caption = "Atualizar Cód. Imposto";
                }

                if (ItemEventInfo.EventType == BoEventTypes.et_CLICK)
                {
                    if (ItemEventInfo.ItemUID == "bt_Import")
                    {
                        if (Form.Mode == BoFormMode.fm_OK_MODE)
                        {
                            Matrix mt_Tax = (Matrix)Form.Items.Item("2003").Specific;
                            int selectedRow = mt_Tax.GetNextSelectedRow();
                            if (selectedRow != -1)
                            {
                                if (ItemEventInfo.ItemUID == "bt_Import")
                                {
                                    this.Import(selectedRow);
                                }
                            }
                            else
                            {
                                SBOApp.Application.SetStatusBarMessage("Por favor, selecione a linha que a atualização deverá ser iniciada");
                            }
                        }
                        else
                        {
                            SBOApp.Application.SetStatusBarMessage("Por favor, salve os dados antes de continuar");
                        }
                    }
                }
            }

            return true;
        }

        private void Import(int selectedRow)
        {
            DialogUtil dialogUtil = new DialogUtil();
            string filePath = dialogUtil.OpenFileDialog("Arquivo Excel|*.xlsx");
            if (!String.IsNullOrEmpty(filePath))
            {
                TaxDeterminationBLL taxDeterminationBLL = new TaxDeterminationBLL();
                List<TaxDeterminationModel> list = taxDeterminationBLL.UpdateTaxes(filePath, f80401.SelectedKey, selectedRow);
                if (list != null)
                {
                    FrmImportLog frmImportLog = new FrmImportLog();
                    frmImportLog.Show(list);
                }
                else
                {
                    SBOApp.Application.MessageBox("Importação finalizada com sucesso");
                    Form.Close();
                    SBOApp.Application.ActivateMenuItem("8497");
                }
            }
        }
    }
}
