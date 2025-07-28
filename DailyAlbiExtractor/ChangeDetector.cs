using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System;
using System.Collections.Generic;

namespace DailyAlbiExtractor
{
    public class ChangeDetector
    {
        public List<ApiItem> DetectChanges(List<ApiItem> previous, List<ApiItem> current)
        {
            var changes = new List<ApiItem>();

            // Use ID as unique key
            var prevDict = previous.ToDictionary(p => p.Id);
            foreach (var item in current)
            {
                if (!prevDict.TryGetValue(item.Id, out var prevItem))
                {
                    // New addition
                    changes.Add(item);
                }
                else if (!ItemsAreEqual(prevItem, item))
                {
                    // Modified
                    changes.Add(item);
                }
            }

            return changes;
        }

        private bool ItemsAreEqual(ApiItem a, ApiItem b)
        {
            // Compare key fields; expand as needed
            return a.CodiceFiscale == b.CodiceFiscale &&
                   a.PartitaIVA == b.PartitaIVA &&
                   a.RagioneSociale == b.RagioneSociale &&
                   a.StatoIscrizione == b.StatoIscrizione;
            // Add more comparisons if necessary
        }
    }
}
