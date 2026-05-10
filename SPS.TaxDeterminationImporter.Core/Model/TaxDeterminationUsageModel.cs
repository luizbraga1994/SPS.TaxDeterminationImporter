namespace SPS.TaxDeterminationImporter.Core.Model
{
    public class TaxDeterminationUsageModel
    {
        public int Tcd2Id { get; set; }
        public int UsageId { get; set; }
        public string Usage { get; set; }
        public string TaxCode { get; set; }
        public string TaxCodeExpense { get; set; }
        public string TaxCodePurchase { get; set; }
    }
}
