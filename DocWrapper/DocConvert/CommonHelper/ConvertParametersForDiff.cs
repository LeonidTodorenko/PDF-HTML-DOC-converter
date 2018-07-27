using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonHelper
{
    public class ConvertParametersForDiff
    {
        public ConvertParametersForDiff()
        {
        }

        public ConvertParametersForDiff(String originalDiffName, String modifiedDiffName, String diffResult, String partnerDiffName, String diffResultEnd)
        {
            OriginalDiffName = originalDiffName;
            ModifiedDiffName = modifiedDiffName;
            DiffResult = diffResult;
            PartnerDiffName = partnerDiffName;
            DiffResultEnd = diffResultEnd;
        }

        public String OriginalDiffName { get; set; }
        public String ModifiedDiffName { get; set; }
        public String DiffResult { get; set; }
        public String PartnerDiffName { get; set; }
        public String DiffResultEnd { get; set; }
    }
}
