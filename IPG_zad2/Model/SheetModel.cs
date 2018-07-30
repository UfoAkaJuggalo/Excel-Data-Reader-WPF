using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IPG_zad2.Model
{
    public class SheetModel
    {
        public List<int> Id { get; set; }
        public List<string> Name { get; set; }
        public List<int> Price { get; set; }
        public List<int> Position { get; set; }
        public List<string> Level { get; set; }
        public List<string> Description { get; set; }
        public List<string> Order { get; set; }
        public List<DateRangeColumn> EmissionDatesList { get; set; }
        public List<LevelAveragePrice> AveragePricePerLevel { get; set; }

        public SheetModel()
        {
            EmissionDatesList = new List<DateRangeColumn>();
        }

        public void CalcAveragePricePerLevel()
        {
            AveragePricePerLevel = new List<LevelAveragePrice>();
            List<string> levelList = Level.Select(s=>s.Trim().ToUpper()).Distinct().ToList();
            levelList.Sort();
            foreach (string level in levelList)
            {
                int daysCounter = 0;
                int priceCounter = 0;
                List<int> levelIndex = Level.Select((value, idx) => new { value, idx })
                    .Where(w => string.Compare(w.value.Trim().ToUpper(), level) == 0)
                    .Select(s => s.idx)
                    .ToList();
                foreach (int index in levelIndex)
                {
                    priceCounter += Price.ElementAt(index);
                    foreach (DateRangeColumn emissionRange in EmissionDatesList)
                    {
                        if (emissionRange.EmissionsList.ElementAt(index))
                            daysCounter += (int)(emissionRange.DtTo - emissionRange.DtFrom).TotalDays;
                    }
                }
                AveragePricePerLevel.Add(new LevelAveragePrice
                {
                    Level = level,
                    AveragePrice = priceCounter / daysCounter
                });
            }
        }
    }
}
