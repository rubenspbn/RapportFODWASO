using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Attributes;

namespace RapportTewerkstellingCorona.Models
{
    public class SubTotalTW
    {
        [EpplusIgnore]
        private static int counter = 0;
        [EpplusIgnore]
        public int ID { get; set; }
        public string WerknemerTypeCaptionNL { get; set; }
        public int Jaar { get; set; }
        public int Week { get; set; }
        public decimal AndereAfwezigheid { get; set; }
        public decimal Gepresteerd { get; set; }
        [EpplusIgnore]
        public decimal GewoneEconomischeWerkloosheid { get; set; }
        [EpplusIgnore]
        public decimal WerkloosheidCorona { get; set; }
        [EpplusIgnore]
        public decimal ZiekteGewaarborgdLoon { get; set; }
        [EpplusIgnore]
        public decimal ZiekteNa1Jaar { get; set; }
        [EpplusIgnore]
        public decimal ZiekteNaGewaarborgdLoon { get; set; }
        public decimal TotaalWerkloosheid
        {
            get { return GewoneEconomischeWerkloosheid + WerkloosheidCorona; }
        }
        public decimal TotaalZiekte
        {
            get { return ZiekteGewaarborgdLoon + ZiekteNa1Jaar + ZiekteNaGewaarborgdLoon; }
        }
        public decimal GrandTotal { get; set; }
        public SubTotalTW()
        {
            ID = System.Threading.Interlocked.Increment(ref counter);
        }
        public override string ToString()
        {
            return $"{WerknemerTypeCaptionNL,-50}|{Jaar,-10}|{Week,-10}|{AndereAfwezigheid,-25:00.##}|{Gepresteerd,-25:00.##}|{TotaalWerkloosheid,-25:00.##}|{TotaalZiekte,-15:00.##}|{GrandTotal,-15}";
        }
    }
}
