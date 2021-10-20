using OfficeOpenXml.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RapportTewerkstellingCorona.Models
{
    public class Tewerkstellingslijn:IComparable<Tewerkstellingslijn>
    {
        #region PROPERTIES
        [EpplusIgnore]
        private static int counter = 0;
        [EpplusIgnore]
        public int ID { get; set; }
        [EpplusIgnore]
        public string OfficieelPC { get; set; }
        [EpplusIgnore]
        public string AcertaPC { get; set; }
        [EpplusIgnore]
        public string OfficieelSubPC { get; set; }
        public string PC
        {
            get 
            {
                if (OfficieelSubPC == "NULL")
                    return OfficieelPC;
                else
                    return OfficieelSubPC; 
            }
        }
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
        #endregion
        #region CONSTRUCTORS
        public Tewerkstellingslijn()
        {
            ID = System.Threading.Interlocked.Increment(ref counter);
        }
        #endregion
        #region METHODS
        public override string ToString()
        {
            return $"{PC,-8}|{WerknemerTypeCaptionNL,-50}|{Jaar,-10}|{Week,-10}|{AndereAfwezigheid.ToString("00.##"),-25}|{Gepresteerd.ToString("00.##"),-25}|{TotaalWerkloosheid.ToString("00.##"),-25}|{TotaalZiekte.ToString("00.##"),-15}|{GrandTotal,-15}";
        }

        public int CompareTo(Tewerkstellingslijn other)
        {
            int result = Jaar.CompareTo(other.Jaar);
            if (result != 0) //CHECK IF PROPERTY IS EQUAL
                return result; //PROPERTY WAS NOT THE SAME
            result = Week.CompareTo(other.Week);
            if (result != 0)
                return result;
            result = WerknemerTypeCaptionNL.CompareTo(other.WerknemerTypeCaptionNL);
            if (result != 0)
                return result;
            result = PC.CompareTo(other.PC);
            if (result != 0)
                return result;
            return 0; //ALL PROPERTIES WERE EQUAL
        }

        public override bool Equals(object obj)
        {
            if (obj == null)
                return false;
            if (!(obj is Tewerkstellingslijn))
                return false;
            Tewerkstellingslijn other = (Tewerkstellingslijn)obj;
            if (CompareTo(other) == 0)
                return true;
            else
                return false;
        }
        #endregion
    }
}
