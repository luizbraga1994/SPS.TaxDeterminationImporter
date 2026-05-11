using SBO.Hub.Attributes;
using System;
using System.Collections.Generic;

namespace SPS.TaxDeterminationImporter.Core.Model
{
    public class TaxDeterminationModel
    {
        public int DisplayOrder { get; set; }
        public int Line { get; set; }
        public int Tcd2Id { get; set; }

        public string Value1 { get; set; }
        public string Value2 { get; set; }
        public string Value3 { get; set; }
        public string Value4 { get; set; }
        public string Value5 { get; set; }

        public DateTime DateFrom { get; set; }
        public DateTime? DateTo { get; set; }

        public List<TaxDeterminationUsageModel> TaxUsageList { get; set; }

        public string Error { get; set; } = String.Empty;
    }
}
