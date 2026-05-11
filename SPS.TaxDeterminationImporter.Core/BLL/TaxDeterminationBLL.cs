using ClosedXML.Excel;
using SAPbobsCOM;
using SAPbouiCOM;
using SBO.Hub;
using SBO.Hub.DAO;
using SPS.TaxDeterminationImporter.Core.DAO;
using SPS.TaxDeterminationImporter.Core.Enum;
using SPS.TaxDeterminationImporter.Core.Model;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;

namespace SPS.TaxDeterminationImporter.Core.BLL
{
    public class TaxDeterminationBLL
    {
        #region Properties

        private static List<string> BusinesPartnerList;
        private static List<string> ItemList;
        private static List<string> MaterialList;
        private static List<ValidatorModel> NcmList;
        private static List<string> StateList;
        private static List<ValidatorModel> ItemGroupList;
        private static List<ValidatorModel> CustomerGroupList;
        private static List<ValidatorModel> SupplierGroupList;
        private static List<string> BranchList;
        private static List<string> TaxList;

        #endregion Properties

        #region ImportData

        public List<TaxDeterminationModel> ImportData(string filePath, int lineKey)
        {
            List<TaxDeterminationModel> list = GetListFromExcelFile(filePath);
            if (list.Any(m => !String.IsNullOrEmpty(m.Error)))
            {
                return list.Where(m => !String.IsNullOrEmpty(m.Error)).ToList();
            }

            SBOApp.Application.SetStatusBarMessage("Iniciando importação, por favor aguarde", BoMessageTime.bmt_Long, false);
            TaxCodeDeterminationsTCDService oTcdService = SBOApp.Company.GetCompanyService().GetBusinessService(ServiceTypes.TaxCodeDeterminationsTCDService) as TaxCodeDeterminationsTCDService;
            TaxCodeDeterminationTCDParams oTcdParams = oTcdService.GetDataInterface(TaxCodeDeterminationsTCDServiceDataInterfaces.tcdsTaxDeterminationTCDParams) as TaxCodeDeterminationTCDParams;
            oTcdParams.AbsId = 1;
            TaxCodeDeterminationTCD oTcd = oTcdService.GetTaxCodeDeterminationTCD(oTcdParams) as TaxCodeDeterminationTCD;

            TaxCodeDeterminationTCDKeyField keyField = oTcd.KeyFields.Item(lineKey - 1);
            List<TaxDeterminationModel> errorList = ValidateList(list, keyField);
            if (errorList.Count > 0)
            {
                return errorList;
            }
            list = this.ConvertDescriptionToId(list, keyField);

            int displayOrder = 0;
            if (keyField.Values.Count > 0)
            {
                displayOrder = Convert.ToInt32(CrudDAO.ExecuteScalar(String.Format(Scripts.Resource.GetString("TaxDetermination_GetMaxDisplayOrder"), keyField.AbsId)));
            }

            TaxCodeDeterminationTCDValue values = null;
            TaxCodeDeterminationTCDPeriod period = null;
            TaxCodeDeterminationTCDByUsage usage = null;
            DateTime startdate = DateTime.Now;

            ProgressBar pgb = SBOApp.Application.StatusBar.CreateProgressBar("Inserindo dados", list.Count, false);
            List<TaxDeterminationModel> existingList = null;
            if (list.Count > 500)
            {
                existingList = new CrudDAO().FillModelListFromSql<TaxDeterminationModel>(String.Format(Scripts.Resource.GetString("TaxDetermination_GetKeyValues"), keyField.AbsId));
            }

            try
            {
                Dictionary<int, int> displayOrderValues = null;

                foreach (var model in list)
                {
                    pgb.Value++;
                    int existingOrder = 0;

                    if (existingList != null)
                    {
                        TaxDeterminationModel existingModel = existingList.FirstOrDefault(m => m.Value1 == model.Value1 && m.Value2 == model.Value2 && m.Value3 == model.Value3 && m.Value4 == model.Value4 && m.Value5 == model.Value5);
                        if (existingModel != null)
                        {
                            existingOrder = existingModel.DisplayOrder;
                        }
                    }
                    else
                    {
                        existingOrder = Convert.ToInt32(CrudDAO.ExecuteScalar(String.Format(Scripts.Resource.GetString("TaxDetermination_GetDisplayOrder"), keyField.AbsId, model.Value1, model.Value2, model.Value3, model.Value4, model.Value5)));
                    }
                    if (existingOrder == 0)
                    {
                        displayOrder++;
                        values = keyField.Values.Add();

                        values.Value1 = model.Value1;
                        values.Value2 = model.Value2;
                        values.Value3 = model.Value3;
                        values.Value4 = model.Value4;
                        values.Value5 = model.Value5;
                        values.DispOrder = displayOrder;
                    }
                    else
                    {
                        values = keyField.Values.Item(existingOrder - 1);

                        // Se está fora da sequência, copia para o dictonary para ser consultado
                        if (values.DispOrder != existingOrder)
                        {
                            if (displayOrderValues == null)
                            {
                                displayOrderValues = new Dictionary<int, int>();

                                for (int i = 0; i < keyField.Values.Count; i++)
                                {
                                    values = keyField.Values.Item(i);
                                    displayOrderValues.Add(values.DispOrder, i);
                                }
                            }

                            values = keyField.Values.Item(displayOrderValues[existingOrder]);
                        }
                    }

                    if (values.Periods.Count > 0)
                    {
                        period = values.Periods.Item(0);
                    }
                    else
                    {
                        period = values.Periods.Add();
                    }

                    period.EffectFrom = model.DateFrom;
                    if (model.DateTo.HasValue)
                    {
                        period.EffectTo = model.DateTo.Value;
                    }
                    foreach (var taxUsage in model.TaxUsageList)
                    {
                        bool usageFound = false;
                        for (int i = 0; i < period.ByUsages.Count; i++)
                        {
                            usage = period.ByUsages.Item(i);
                            if (usage.UsageCode == taxUsage.UsageId)
                            {
                                usageFound = true;
                                break;
                            }
                        }
                        if (!usageFound)
                        {
                            usage = period.ByUsages.Add();
                        }
                        usage.UsageCode = taxUsage.UsageId;
                        usage.TaxCode = taxUsage.TaxCode;
                        usage.FreightTaxCode = taxUsage.TaxCodeExpense;
                        usage.PurchaseTaxCode = taxUsage.TaxCodePurchase;
                    }
                }

                SBOApp.Application.SetStatusBarMessage("Atualizando informações no banco de dados, por favor aguarde", BoMessageTime.bmt_Long, false);
                oTcdService.UpdateTaxCodeDeterminationTCD(oTcd);
            }
            catch
            {
                throw;
            }
            finally
            {
                pgb.Stop();
                Marshal.ReleaseComObject(pgb);
                pgb = null;

                Marshal.ReleaseComObject(oTcd);
                Marshal.ReleaseComObject(oTcdParams);
                Marshal.ReleaseComObject(oTcdService);
                Marshal.ReleaseComObject(keyField);
                if (values != null)
                {
                    Marshal.ReleaseComObject(values);
                }
                if (period != null)
                {
                    Marshal.ReleaseComObject(period);
                }
                if (usage != null)
                {
                    Marshal.ReleaseComObject(usage);
                }
                oTcd = null;
                oTcdParams = null;
                oTcdService = null;
                keyField = null;
                values = null;
                period = null;
                usage = null;
                double minutes = DateTime.Now.Subtract(startdate).TotalMinutes;
                SBOApp.Application.SetStatusBarMessage($"Importação finalizada - Tempo de execução: {minutes.ToString("f2")} minutos", BoMessageTime.bmt_Medium, false);
            }

            return null;
        }

        #endregion ImportData

        public string ExportData(string path, int lineKey, string description)
        {
            ProgressBar pgb = null;
            try
            {
                SBOApp.Application.SetStatusBarMessage("Iniciando exportação, por favor aguarde", BoMessageTime.bmt_Long, false);
                using (var wb = new XLWorkbook())
                {
                    var ws = wb.AddWorksheet("Determinação Imposto");

                    List<TaxDeterminationModel> list = new CrudDAO().FillModelListFromSql<TaxDeterminationModel>(String.Format(Scripts.Resource.GetString("TaxDeterminationValues_GetToExport"), lineKey));

                    pgb = SBOApp.Application.StatusBar.CreateProgressBar("Exportando dados", list.Count, false);

                    List<TaxDeterminationUsageModel> taxCodeList = new CrudDAO().FillModelListFromSql<TaxDeterminationUsageModel>(String.Format(Scripts.Resource.GetString("TaxDeterminationUsage_GetToExport"), lineKey));

                    System.Data.DataTable dataTable = new System.Data.DataTable();
                    dataTable.Columns.Add("Valor1", typeof(string));
                    dataTable.Columns.Add("Valor2", typeof(string));
                    dataTable.Columns.Add("Valor3", typeof(string));
                    dataTable.Columns.Add("Valor4", typeof(string));
                    dataTable.Columns.Add("Valor5", typeof(string));
                    dataTable.Columns.Add("Efetivo De", typeof(DateTime));
                    dataTable.Columns.Add("Efetivo Ate", typeof(DateTime));

                    List<string> usageList = taxCodeList.Select(s => s.Usage).Distinct().ToList();
                    foreach (var usage in usageList)
                    {
                        dataTable.Columns.Add(usage, typeof(string));
                        dataTable.Columns.Add("IVA - " + usage, typeof(string));
                        dataTable.Columns.Add("DA - " + usage, typeof(string));
                    }

                    foreach (var determination in list)
                    {
                        pgb.Value++;

                        DataRow dataRow = dataTable.NewRow();
                        dataRow["Valor1"] = determination.Value1;
                        dataRow["Valor2"] = determination.Value2;
                        dataRow["Valor3"] = determination.Value3;
                        dataRow["Valor4"] = determination.Value4;
                        dataRow["Valor5"] = determination.Value5;

                        dataRow["Efetivo De"] = determination.DateFrom;
                        if (determination.DateTo.HasValue)
                        {
                            dataRow["Efetivo Ate"] = determination.DateTo;
                        }
                        List<TaxDeterminationUsageModel> taxCodeListLine = taxCodeList.Where(m => m.Tcd2Id == determination.Tcd2Id).ToList();

                        foreach (var taxCode in taxCodeListLine)
                        {
                            dataRow[taxCode.Usage] = taxCode.TaxCodePurchase;
                            dataRow["IVA - " + taxCode.Usage] = taxCode.TaxCode;
                            dataRow["DA - " + taxCode.Usage] = taxCode.TaxCodeExpense;
                        }
                        dataTable.Rows.Add(dataRow);
                    }

                    ws.Cell(1, "A").InsertTable(dataTable);
                    ws.Columns().AdjustToContents();

                    wb.SaveAs(Path.Combine(path, lineKey + " - " + description + ".xlsx"));
                } // using wb
                return String.Empty;
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                if (pgb != null)
                {
                    pgb.Stop();
                    Marshal.ReleaseComObject(pgb);
                    pgb = null;
                }
            }
        }

        #region UpdateTaxes
        /// <summary>
        /// Obsoleto - atualizando de acordo com as chaves
        /// </summary>
        /// <param name="filePath"></param>
        /// <param name="lineKey"></param>
        /// <param name="lineValue"></param>
        /// <returns></returns>
        public List<TaxDeterminationModel> UpdateTaxes(string filePath, int lineKey, int lineValue)
        {
            List<TaxDeterminationModel> list = GetListFromExcelFile(filePath);
            List<TaxDeterminationModel> errorList = this.ValidateTax(list);
            if (errorList.Any(m => !String.IsNullOrEmpty(m.Error)))
            {
                return list.Where(m => !String.IsNullOrEmpty(m.Error)).ToList();
            }

            TaxCodeDeterminationsTCDService oTcdService = SBOApp.Company.GetCompanyService().GetBusinessService(ServiceTypes.TaxCodeDeterminationsTCDService) as TaxCodeDeterminationsTCDService;
            TaxCodeDeterminationTCDParams oTcdParams = oTcdService.GetDataInterface(TaxCodeDeterminationsTCDServiceDataInterfaces.tcdsTaxDeterminationTCDParams) as TaxCodeDeterminationTCDParams;
            oTcdParams.AbsId = 1;
            TaxCodeDeterminationTCD oTcd = oTcdService.GetTaxCodeDeterminationTCD(oTcdParams) as TaxCodeDeterminationTCD;

            TaxCodeDeterminationTCDKeyField keyField = oTcd.KeyFields.Item(lineKey - 1);
            errorList = ValidateList(list, keyField);
            if (errorList.Count > 0)
            {
                return errorList;
            }
            list = this.ConvertDescriptionToId(list, keyField);

            TaxCodeDeterminationTCDValue values = null;
            TaxCodeDeterminationTCDPeriod period = null;
            TaxCodeDeterminationTCDByUsage usage = null;
            DateTime startdate = DateTime.Now;

            ProgressBar pgb = SBOApp.Application.StatusBar.CreateProgressBar("Inserindo dados", list.Count, false);
            try
            {
                int currentLine = lineValue - 1;

                if (list.Count > keyField.Values.Count - currentLine)
                {
                    throw new Exception("Quantidade de linhas do arquivo é maior que o total de linhas a serem atualizadas, por favor verifique");
                }

                foreach (var model in list)
                {
                    pgb.Value++;

                    values = keyField.Values.Item(currentLine);
                    if (values.Periods.Count > 0)
                    {
                        period = values.Periods.Item(0);
                    }
                    else
                    {
                        period = values.Periods.Add();
                    }
                    period.EffectFrom = model.DateFrom;
                    if (model.DateTo.HasValue)
                    {
                        period.EffectTo = model.DateTo.Value;
                    }
                    foreach (var taxUsage in model.TaxUsageList)
                    {
                        bool exists = false;
                        for (int i = 0; i < period.ByUsages.Count; i++)
                        {
                            usage = period.ByUsages.Item(i);
                            if (usage.UsageCode == taxUsage.UsageId)
                            {
                                usage.TaxCode = taxUsage.TaxCode;
                                usage.FreightTaxCode = taxUsage.TaxCodeExpense;
                                usage.PurchaseTaxCode = taxUsage.TaxCodePurchase;
                                exists = true;
                                break;
                            }
                        }

                        if (!exists)
                        {
                            usage = period.ByUsages.Add();
                            usage.UsageCode = taxUsage.UsageId;
                            usage.TaxCode = taxUsage.TaxCode;
                            usage.FreightTaxCode = taxUsage.TaxCodeExpense;
                            usage.PurchaseTaxCode = taxUsage.TaxCodePurchase;
                        }
                    }
                    currentLine++;
                }

                SBOApp.Application.SetStatusBarMessage("Atualizando informações no banco de dados, por favor aguarde", BoMessageTime.bmt_Long, false);
                oTcdService.UpdateTaxCodeDeterminationTCD(oTcd);
            }
            catch
            {
                throw;
            }
            finally
            {
                pgb.Stop();
                Marshal.ReleaseComObject(pgb);
                pgb = null;

                Marshal.ReleaseComObject(oTcd);
                Marshal.ReleaseComObject(oTcdParams);
                Marshal.ReleaseComObject(oTcdService);
                Marshal.ReleaseComObject(keyField);
                if (values != null)
                {
                    Marshal.ReleaseComObject(values);
                }
                if (period != null)
                {
                    Marshal.ReleaseComObject(period);
                }
                if (usage != null)
                {
                    Marshal.ReleaseComObject(usage);
                }
                oTcd = null;
                oTcdParams = null;
                oTcdService = null;
                keyField = null;
                values = null;
                period = null;
                usage = null;
                double minutes = DateTime.Now.Subtract(startdate).TotalMinutes;
                SBOApp.Application.SetStatusBarMessage($"Importação finalizada - Tempo de execução: {minutes.ToString("f2")} minutos", BoMessageTime.bmt_Medium, false);
            }
            return null;
        }

        #endregion UpdateTaxes

        #region ClearValues

        public string RemoveLines(int lineKey, List<int> linesToRemove)
        {
            TaxCodeDeterminationsTCDService oTcdService = SBOApp.Company.GetCompanyService().GetBusinessService(ServiceTypes.TaxCodeDeterminationsTCDService) as TaxCodeDeterminationsTCDService;
            TaxCodeDeterminationTCDParams oTcdParams = oTcdService.GetDataInterface(TaxCodeDeterminationsTCDServiceDataInterfaces.tcdsTaxDeterminationTCDParams) as TaxCodeDeterminationTCDParams;
            oTcdParams.AbsId = 1;
            TaxCodeDeterminationTCD oTcd = oTcdService.GetTaxCodeDeterminationTCD(oTcdParams) as TaxCodeDeterminationTCD;
            TaxCodeDeterminationTCDKeyField keyField = oTcd.KeyFields.Item(lineKey - 1);

            try
            {
                for (int i = linesToRemove.Count - 1; i >= 0; i--)
                {
                    keyField.Values.Remove(linesToRemove[i]);
                }

                oTcdService.UpdateTaxCodeDeterminationTCD(oTcd);
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
            finally
            {
                Marshal.ReleaseComObject(oTcd);
                Marshal.ReleaseComObject(oTcdParams);
                Marshal.ReleaseComObject(oTcdService);
                Marshal.ReleaseComObject(keyField);

                oTcd = null;
                oTcdParams = null;
                oTcdService = null;
                keyField = null;
            }

            return String.Empty;
        }

        #endregion ClearValues

        #region GetListFromExcelFile

        public List<TaxDeterminationModel> GetListFromExcelFile(string filePath)
        {
            List<TaxDeterminationModel> list = new List<TaxDeterminationModel>();

            using (var wb = new XLWorkbook(filePath))
            {
                var ws = wb.Worksheet(1);

                int lastRow = ws.LastRowUsed().RowNumber() + 1;
                int lastColumn = ws.LastColumnUsed().ColumnNumber() + 1;

                ProgressBar pgb = SBOApp.Application.StatusBar.CreateProgressBar("Carregando dados do arquivo", lastRow, false);

                try
                {
                    List<TaxDeterminationUsageModel> usageList = new List<TaxDeterminationUsageModel>();

                    IXLRow headerRow = ws.Row(1);
                    for (int i = 8; i < lastColumn; i += 3)
                    {
                        TaxDeterminationUsageModel usageModel = new TaxDeterminationUsageModel();
                        usageModel.Usage = headerRow.Cell(i).Value.ToString();

                        usageModel.UsageId = Convert.ToInt32(CrudDAO.ExecuteScalar(String.Format(Scripts.Resource.GetString("Usage_GetIdByDescription"), headerRow.Cell(i).Value.ToString().ToLower())));
                        if (usageModel.UsageId == 0)
                        {
                            throw new Exception($"Utilização '{headerRow.Cell(i).Value.ToString()}' não encontrada");
                        }
                        usageList.Add(usageModel);
                    }

                    for (int i = 2; i < lastRow; i++)
                    {
                        pgb.Value++;
                        IXLRow row = ws.Row(i);

                        TaxDeterminationModel model = new TaxDeterminationModel();
                        model.Line = i;
                        model.Value1 = row.Cell("A").Value.ToString();
                        model.Value2 = row.Cell("B").Value.ToString();
                        model.Value3 = row.Cell("C").Value.ToString();
                        model.Value4 = row.Cell("D").Value.ToString();
                        model.Value5 = row.Cell("E").Value.ToString();
                        DateTime date;
                        if (DateTime.TryParse(row.Cell("F").Value.ToString(), out date))
                        {
                            model.DateFrom = date;
                        }
                        else
                        {
                            model.Error = "Coluna 'Efetivo de': Formato da data inválido";
                        }
                        if (!String.IsNullOrEmpty(row.Cell("G").Value.ToString().Trim()))
                        {
                            if (DateTime.TryParse(row.Cell("G").Value.ToString(), out date))
                            {
                                model.DateTo = date;
                            }
                            else
                            {
                                model.Error = "Coluna 'Efetivo até': Formato da data inválido";
                            }
                        }
                        model.TaxUsageList = new List<TaxDeterminationUsageModel>();
                        for (int j = 8; j < lastColumn; j += 3)
                        {
                            TaxDeterminationUsageModel usageModel = new TaxDeterminationUsageModel();
                            usageModel.Usage = headerRow.Cell(j).Value.ToString();
                            usageModel.UsageId = usageList.FirstOrDefault(m => m.Usage == headerRow.Cell(j).Value.ToString()).UsageId;
                            usageModel.TaxCodePurchase = row.Cell(j).Value.ToString();
                            usageModel.TaxCode = row.Cell(j + 1).Value.ToString();
                            usageModel.TaxCodeExpense = row.Cell(j + 2).Value.ToString();
                            model.TaxUsageList.Add(usageModel);
                        }

                        TaxDeterminationModel duplicated = list.FirstOrDefault(m => m.Value1 == model.Value1 && m.Value2 == model.Value2 && m.Value3 == model.Value3 && m.Value4 == model.Value4 && m.Value5 == model.Value5);

                        if (duplicated != null)
                        {
                            model.Error = "Valores já existentes na linha " + duplicated.Line;
                        }
                        list.Add(model);
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception("Erro ao ler arquivo: " + ex.Message);
                }
                finally
                {
                    pgb.Stop();
                    Marshal.ReleaseComObject(pgb);
                    pgb = null;
                }
            } // using wb
            return list;
        }

        #endregion GetListFromExcelFile

        #region ConvertDescriptionToId

        public List<TaxDeterminationModel> ConvertDescriptionToId(List<TaxDeterminationModel> list, TaxCodeDeterminationTCDKeyField keyField)
        {
            TaxKeyFieldTypeEnum fieldType1 = (TaxKeyFieldTypeEnum)keyField.KeyField1;
            TaxKeyFieldTypeEnum fieldType2 = (TaxKeyFieldTypeEnum)keyField.KeyField2;
            TaxKeyFieldTypeEnum fieldType3 = (TaxKeyFieldTypeEnum)keyField.KeyField3;
            TaxKeyFieldTypeEnum fieldType4 = (TaxKeyFieldTypeEnum)keyField.KeyField4;
            TaxKeyFieldTypeEnum fieldType5 = (TaxKeyFieldTypeEnum)keyField.KeyField5;

            switch (fieldType1)
            {
                case TaxKeyFieldTypeEnum.NcmCode:
                    foreach (var item in list)
                    {
                        var ncm = NcmList?.FirstOrDefault(m => m.Code == item.Value1);
                        if (ncm != null) item.Value1 = ncm.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.ItemGroup:
                    foreach (var item in list)
                    {
                        var ig = ItemGroupList?.FirstOrDefault(m => m.Code == item.Value1);
                        if (ig != null) item.Value1 = ig.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.CustomerGroup:
                    foreach (var item in list)
                    {
                        var cg = CustomerGroupList?.FirstOrDefault(m => m.Code == item.Value1);
                        if (cg != null) item.Value1 = cg.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.SupplierGroup:
                    foreach (var item in list)
                    {
                        var sg = SupplierGroupList?.FirstOrDefault(m => m.Code == item.Value1);
                        if (sg != null) item.Value1 = sg.Id.ToString();
                    }
                    break;
            }

            switch (fieldType2)
            {
                case TaxKeyFieldTypeEnum.NcmCode:
                    foreach (var item in list)
                    {
                        var ncm = NcmList?.FirstOrDefault(m => m.Code == item.Value2);
                        if (ncm != null) item.Value2 = ncm.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.ItemGroup:
                    foreach (var item in list)
                    {
                        var ig = ItemGroupList?.FirstOrDefault(m => m.Code == item.Value2);
                        if (ig != null) item.Value2 = ig.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.CustomerGroup:
                    foreach (var item in list)
                    {
                        var cg = CustomerGroupList?.FirstOrDefault(m => m.Code == item.Value2);
                        if (cg != null) item.Value2 = cg.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.SupplierGroup:
                    foreach (var item in list)
                    {
                        var sg = SupplierGroupList?.FirstOrDefault(m => m.Code == item.Value2);
                        if (sg != null) item.Value2 = sg.Id.ToString();
                    }
                    break;
            }

            switch (fieldType3)
            {
                case TaxKeyFieldTypeEnum.NcmCode:
                    foreach (var item in list)
                    {
                        var ncm = NcmList?.FirstOrDefault(m => m.Code == item.Value3);
                        if (ncm != null) item.Value3 = ncm.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.ItemGroup:
                    foreach (var item in list)
                    {
                        var ig = ItemGroupList?.FirstOrDefault(m => m.Code == item.Value3);
                        if (ig != null) item.Value3 = ig.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.CustomerGroup:
                    foreach (var item in list)
                    {
                        var cg = CustomerGroupList?.FirstOrDefault(m => m.Code == item.Value3);
                        if (cg != null) item.Value3 = cg.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.SupplierGroup:
                    foreach (var item in list)
                    {
                        var sg = SupplierGroupList?.FirstOrDefault(m => m.Code == item.Value3);
                        if (sg != null) item.Value3 = sg.Id.ToString();
                    }
                    break;
            }

            switch (fieldType4)
            {
                case TaxKeyFieldTypeEnum.NcmCode:
                    foreach (var item in list)
                    {
                        var ncm = NcmList?.FirstOrDefault(m => m.Code == item.Value4);
                        if (ncm != null) item.Value4 = ncm.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.ItemGroup:
                    foreach (var item in list)
                    {
                        var ig = ItemGroupList?.FirstOrDefault(m => m.Code == item.Value4);
                        if (ig != null) item.Value4 = ig.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.CustomerGroup:
                    foreach (var item in list)
                    {
                        var cg = CustomerGroupList?.FirstOrDefault(m => m.Code == item.Value4);
                        if (cg != null) item.Value4 = cg.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.SupplierGroup:
                    foreach (var item in list)
                    {
                        var sg = SupplierGroupList?.FirstOrDefault(m => m.Code == item.Value4);
                        if (sg != null) item.Value4 = sg.Id.ToString();
                    }
                    break;
            }

            switch (fieldType5)
            {
                case TaxKeyFieldTypeEnum.NcmCode:
                    foreach (var item in list)
                    {
                        var ncm = NcmList?.FirstOrDefault(m => m.Code == item.Value5);
                        if (ncm != null) item.Value5 = ncm.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.ItemGroup:
                    foreach (var item in list)
                    {
                        var ig = ItemGroupList?.FirstOrDefault(m => m.Code == item.Value5);
                        if (ig != null) item.Value5 = ig.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.CustomerGroup:
                    foreach (var item in list)
                    {
                        var cg = CustomerGroupList?.FirstOrDefault(m => m.Code == item.Value5);
                        if (cg != null) item.Value5 = cg.Id.ToString();
                    }
                    break;

                case TaxKeyFieldTypeEnum.SupplierGroup:
                    foreach (var item in list)
                    {
                        var sg = SupplierGroupList?.FirstOrDefault(m => m.Code == item.Value5);
                        if (sg != null) item.Value5 = sg.Id.ToString();
                    }
                    break;
            }

            return list;
        }

        #endregion ConvertDescriptionToId

        #region ValidateList

        public List<TaxDeterminationModel> ValidateList(List<TaxDeterminationModel> list, TaxCodeDeterminationTCDKeyField keyField)
        {
            List<TaxDeterminationModel> errorList = ValidateTax(list);

            TaxKeyFieldTypeEnum fieldType1 = (TaxKeyFieldTypeEnum)keyField.KeyField1;
            TaxKeyFieldTypeEnum fieldType2 = (TaxKeyFieldTypeEnum)keyField.KeyField2;
            TaxKeyFieldTypeEnum fieldType3 = (TaxKeyFieldTypeEnum)keyField.KeyField3;
            TaxKeyFieldTypeEnum fieldType4 = (TaxKeyFieldTypeEnum)keyField.KeyField4;
            TaxKeyFieldTypeEnum fieldType5 = (TaxKeyFieldTypeEnum)keyField.KeyField5;

            CrudDAO dao = new CrudDAO();

            if ((fieldType1 == TaxKeyFieldTypeEnum.BusinessPartner || fieldType2 == TaxKeyFieldTypeEnum.BusinessPartner || fieldType3 == TaxKeyFieldTypeEnum.BusinessPartner || fieldType4 == TaxKeyFieldTypeEnum.BusinessPartner || fieldType5 == TaxKeyFieldTypeEnum.BusinessPartner))
            {
                List<string> bpValues = new List<string>();
                if (fieldType1 == TaxKeyFieldTypeEnum.BusinessPartner)
                {
                    bpValues = list.Select(m => m.Value1).Distinct().ToList();
                }
                if (fieldType2 == TaxKeyFieldTypeEnum.BusinessPartner)
                {
                    bpValues = list.Select(m => m.Value2).Distinct().ToList();
                }
                if (fieldType3 == TaxKeyFieldTypeEnum.BusinessPartner)
                {
                    bpValues = list.Select(m => m.Value3).Distinct().ToList();
                }
                if (fieldType4 == TaxKeyFieldTypeEnum.BusinessPartner)
                {
                    bpValues = list.Select(m => m.Value4).Distinct().ToList();
                }
                if (fieldType5 == TaxKeyFieldTypeEnum.BusinessPartner)
                {
                    bpValues = list.Select(m => m.Value5).Distinct().ToList();
                }

                BusinesPartnerList = new List<string>();
                foreach (var bp in bpValues)
                {
                    BusinesPartnerList.Add(CrudDAO.ExecuteScalar(String.Format(Scripts.Resource.GetString("BusinessPartner_Get"), bp)).ToString());
                }
            }

            if (ItemList == null && (fieldType1 == TaxKeyFieldTypeEnum.Item || fieldType2 == TaxKeyFieldTypeEnum.Item || fieldType3 == TaxKeyFieldTypeEnum.Item || fieldType4 == TaxKeyFieldTypeEnum.Item || fieldType5 == TaxKeyFieldTypeEnum.Item))
            {
                List<string> itemValues = new List<string>();
                if (fieldType1 == TaxKeyFieldTypeEnum.Item)
                {
                    itemValues = list.Select(m => m.Value1).Distinct().ToList();
                }
                if (fieldType2 == TaxKeyFieldTypeEnum.Item)
                {
                    itemValues = list.Select(m => m.Value2).Distinct().ToList();
                }
                if (fieldType3 == TaxKeyFieldTypeEnum.Item)
                {
                    itemValues = list.Select(m => m.Value3).Distinct().ToList();
                }
                if (fieldType4 == TaxKeyFieldTypeEnum.Item)
                {
                    itemValues = list.Select(m => m.Value4).Distinct().ToList();
                }
                if (fieldType5 == TaxKeyFieldTypeEnum.Item)
                {
                    itemValues = list.Select(m => m.Value5).Distinct().ToList();
                }

                ItemList = new List<string>();
                foreach (var item in itemValues)
                {
                    ItemList.Add(CrudDAO.ExecuteScalar(String.Format(Scripts.Resource.GetString("Item_Get"), item)).ToString());
                }
            }
            if (MaterialList == null && (fieldType1 == TaxKeyFieldTypeEnum.MaterialGroup || fieldType2 == TaxKeyFieldTypeEnum.MaterialGroup || fieldType3 == TaxKeyFieldTypeEnum.MaterialGroup || fieldType4 == TaxKeyFieldTypeEnum.MaterialGroup || fieldType5 == TaxKeyFieldTypeEnum.MaterialGroup))
            {
                MaterialList = dao.FillStringList(Scripts.Resource.GetString("MaterialGroup_Get"));
            }
            if (fieldType1 == TaxKeyFieldTypeEnum.NcmCode || fieldType2 == TaxKeyFieldTypeEnum.NcmCode || fieldType3 == TaxKeyFieldTypeEnum.NcmCode || fieldType4 == TaxKeyFieldTypeEnum.NcmCode || fieldType5 == TaxKeyFieldTypeEnum.NcmCode)
            {
                List<string> ncmValues = new List<string>();
                if (fieldType1 == TaxKeyFieldTypeEnum.NcmCode)
                {
                    ncmValues = list.Select(m => m.Value1).Distinct().ToList();
                }
                if (fieldType2 == TaxKeyFieldTypeEnum.NcmCode)
                {
                    ncmValues = list.Select(m => m.Value2).Distinct().ToList();
                }
                if (fieldType3 == TaxKeyFieldTypeEnum.NcmCode)
                {
                    ncmValues = list.Select(m => m.Value3).Distinct().ToList();
                }
                if (fieldType4 == TaxKeyFieldTypeEnum.NcmCode)
                {
                    ncmValues = list.Select(m => m.Value4).Distinct().ToList();
                }
                if (fieldType5 == TaxKeyFieldTypeEnum.NcmCode)
                {
                    ncmValues = list.Select(m => m.Value5).Distinct().ToList();
                }

                NcmList = new List<ValidatorModel>();
                foreach (var ncm in ncmValues)
                {
                    NcmList.Add(dao.FillModelFromSql<ValidatorModel>(String.Format(Scripts.Resource.GetString("Ncm_GetByCode"), ncm)));
                }
            }
            if (StateList == null && (fieldType1 == TaxKeyFieldTypeEnum.State || fieldType2 == TaxKeyFieldTypeEnum.State || fieldType3 == TaxKeyFieldTypeEnum.State || fieldType4 == TaxKeyFieldTypeEnum.State || fieldType5 == TaxKeyFieldTypeEnum.State))
            {
                StateList = dao.FillStringList(Scripts.Resource.GetString("State_Get"));
            }
            if (ItemGroupList == null && (fieldType1 == TaxKeyFieldTypeEnum.ItemGroup || fieldType2 == TaxKeyFieldTypeEnum.ItemGroup || fieldType3 == TaxKeyFieldTypeEnum.ItemGroup || fieldType4 == TaxKeyFieldTypeEnum.ItemGroup || fieldType5 == TaxKeyFieldTypeEnum.ItemGroup))
            {
                ItemGroupList = dao.FillModelListFromSql<ValidatorModel>(Scripts.Resource.GetString("ItemGroup_Get"));
            }
            if (CustomerGroupList == null && (fieldType1 == TaxKeyFieldTypeEnum.CustomerGroup || fieldType2 == TaxKeyFieldTypeEnum.CustomerGroup || fieldType3 == TaxKeyFieldTypeEnum.CustomerGroup || fieldType4 == TaxKeyFieldTypeEnum.CustomerGroup || fieldType5 == TaxKeyFieldTypeEnum.CustomerGroup))
            {
                CustomerGroupList = dao.FillModelListFromSql<ValidatorModel>(Scripts.Resource.GetString("CustomerGroup_Get"));
            }
            if (SupplierGroupList == null && (fieldType1 == TaxKeyFieldTypeEnum.SupplierGroup || fieldType2 == TaxKeyFieldTypeEnum.SupplierGroup || fieldType3 == TaxKeyFieldTypeEnum.SupplierGroup || fieldType4 == TaxKeyFieldTypeEnum.SupplierGroup || fieldType5 == TaxKeyFieldTypeEnum.SupplierGroup))
            {
                SupplierGroupList = dao.FillModelListFromSql<ValidatorModel>(Scripts.Resource.GetString("SupplierGroup_Get"));
            }
            if (BranchList == null && (fieldType1 == TaxKeyFieldTypeEnum.Branch || fieldType2 == TaxKeyFieldTypeEnum.Branch || fieldType3 == TaxKeyFieldTypeEnum.Branch || fieldType4 == TaxKeyFieldTypeEnum.Branch || fieldType5 == TaxKeyFieldTypeEnum.Branch))
            {
                BranchList = dao.FillStringList(Scripts.Resource.GetString("Branch_Get"));
            }

            List<string> notFound = Validate(list.Select(m => m.Value1).ToList(), fieldType1, keyField.UDFTable1, keyField.UDFAlias1);
            if (notFound.Count > 0)
            {
                List<TaxDeterminationModel> value1List = list.Where(m => notFound.Contains(m.Value1)).ToList();
                value1List.ForEach(m => m.Error += $" {fieldType1.ToString()}: Valor1 ({m.Value1}) não encontrado!");
                errorList.AddRange(value1List);
            }
            notFound = Validate(list.Select(m => m.Value2).Where(m => !String.IsNullOrEmpty(m)).ToList(), fieldType2, keyField.UDFTable2, keyField.UDFAlias2);
            if (notFound.Count > 0)
            {
                List<TaxDeterminationModel> value2List = list.Where(m => notFound.Contains(m.Value2)).ToList();
                value2List.ForEach(m => m.Error += $" {fieldType2.ToString()}: Valor2 ({m.Value2}) não encontrado!");
                errorList.AddRange(value2List);
            }
            notFound = Validate(list.Select(m => m.Value3).Where(m => !String.IsNullOrEmpty(m)).ToList(), fieldType3, keyField.UDFTable3, keyField.UDFAlias3);
            if (notFound.Count > 0)
            {
                List<TaxDeterminationModel> value3List = list.Where(m => notFound.Contains(m.Value3)).ToList();
                value3List.ForEach(m => m.Error += $" {fieldType3.ToString()}: Valor3 ({m.Value3}) não encontrado!");
                errorList.AddRange(value3List);
            }

            notFound = Validate(list.Select(m => m.Value4).Where(m => !String.IsNullOrEmpty(m)).ToList(), fieldType4, keyField.UDFTable4, keyField.UDFAlias4);
            if (notFound.Count > 0)
            {
                List<TaxDeterminationModel> value4List = list.Where(m => notFound.Contains(m.Value4)).ToList();
                value4List.ForEach(m => m.Error += $" {fieldType4.ToString()}: Valor4 ({m.Value4}) não encontrado!");
                errorList.AddRange(value4List);
            }

            notFound = Validate(list.Select(m => m.Value5).Where(m => !String.IsNullOrEmpty(m)).ToList(), fieldType5, keyField.UDFTable5, keyField.UDFAlias5);
            if (notFound.Count > 0)
            {
                List<TaxDeterminationModel> value5List = list.Where(m => notFound.Contains(m.Value5)).ToList();
                value5List.ForEach(m => m.Error += $" {fieldType5.ToString()}: Valor5 ({m.Value5}) não encontrado!");
                errorList.AddRange(value5List);
            }

            return errorList;
        }

        #endregion ValidateList

        #region ValidateTax

        private List<TaxDeterminationModel> ValidateTax(List<TaxDeterminationModel> list)
        {
            if (TaxList == null)
            {
                TaxList = new CrudDAO().FillStringList(Scripts.Resource.GetString("Tax_Get"));
            }

            List<TaxDeterminationModel> errorList = new List<TaxDeterminationModel>();

            foreach (var model in list)
            {
                foreach (var tax in model.TaxUsageList)
                {
                    if (!String.IsNullOrEmpty(tax.TaxCode) && !TaxList.Contains(tax.TaxCode))
                    {
                        TaxDeterminationModel errorModel = new TaxDeterminationModel();
                        errorModel.Line = model.Line;
                        errorModel.Error += $" {tax.TaxCode} - Código de imposto IVA não encontrado!";
                        errorList.Add(errorModel);
                    }
                    if (!String.IsNullOrEmpty(tax.TaxCodeExpense) && !TaxList.Contains(tax.TaxCodeExpense))
                    {
                        TaxDeterminationModel errorModel = new TaxDeterminationModel();
                        errorModel.Line = model.Line;
                        errorModel.Error += $" {tax.TaxCodeExpense} - Código de imposto de Despesa Adicional não encontrado!";
                        errorList.Add(errorModel);
                    }
                    if (!String.IsNullOrEmpty(tax.TaxCodePurchase) && !TaxList.Contains(tax.TaxCodePurchase))
                    {
                        TaxDeterminationModel errorModel = new TaxDeterminationModel();
                        errorModel.Line = model.Line;
                        errorModel.Error += $" {tax.TaxCodePurchase} - Código de imposto sobre compra não encontrado!";
                        errorList.Add(errorModel);
                    }
                }
            }
            return errorList;
        }

        #endregion ValidateTax

        #region Validate

        private List<string> Validate(List<string> valueList, TaxKeyFieldTypeEnum fieldType, string udfTable, string udfAlias)
        {
            List<string> notFound = new List<string>();

            switch (fieldType)
            {
                case TaxKeyFieldTypeEnum.BusinessPartner:
                    notFound = valueList.Where(x => !BusinesPartnerList.Contains(x)).ToList();
                    break;

                case TaxKeyFieldTypeEnum.Item:
                    notFound = valueList.Where(x => !ItemList.Contains(x)).ToList();
                    break;

                case TaxKeyFieldTypeEnum.MaterialGroup:
                    break;

                case TaxKeyFieldTypeEnum.NcmCode:
                    List<string> ncmValues = NcmList.Select(m => m.Code).ToList();
                    notFound = valueList.Where(x => !ncmValues.Contains(x)).ToList();
                    break;

                case TaxKeyFieldTypeEnum.State:
                    notFound = valueList.Where(x => !StateList.Contains(x)).ToList();
                    break;

                case TaxKeyFieldTypeEnum.ItemGroup:
                    List<string> itemGroupValues = ItemGroupList.Select(m => m.Code).ToList();
                    notFound = valueList.Where(x => !itemGroupValues.Contains(x)).ToList();
                    break;

                case TaxKeyFieldTypeEnum.CustomerGroup:
                    break;

                case TaxKeyFieldTypeEnum.SupplierGroup:
                    break;

                case TaxKeyFieldTypeEnum.Branch:
                    notFound = valueList.Where(x => !BranchList.Contains(x)).ToList();
                    break;

                case TaxKeyFieldTypeEnum.UDF:
                    List<string> validValues = new CrudDAO().FillStringList(String.Format(Scripts.Resource.GetString("UDF_GetValidValues"), udfTable, udfAlias));
                    notFound = valueList.Where(x => !validValues.Contains(x)).ToList();
                    break;

                default:
                    break;
            }

            return notFound;
        }

        #endregion Validate
    }
}