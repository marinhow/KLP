using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace KLPOutlook2007AddIn
{
    public class Helper
    {

        //Fødselsnummer Validator Method
        public Tuple<bool, string> CheckFødselsnummer(string number)
        {
            string result = string.Empty;

            //sjekker lengden 
            if (number.Length != 11)
            {
                result = "Fødselsnummer er ikke korrekt!";
                return new Tuple<bool, string>(false, result);
            }

            //D-nummer transformation
            if (Convert.ToInt32(number.Substring(0, 1)) > 3)
            {
                number = (Convert.ToInt32(number.Substring(0, 1)) - 4).ToString() + number.Substring(1, number.Length - 1);
            }

            long foo;
            //Sjekker om det faktisk er et tall det som er skrevet inn
            //Checks if the input is just number
            if (!Int64.TryParse(number, out foo))
            {
                result = "Fødselsnummer er ikke korrekt!";
                return new Tuple<bool, string>(false, result);
            }

            //bare fordi jeg ikke gidder å skrive inputssn.Text hele tiden.
            string num = number;

            //Deler opp i litt mer håndterlige deler
            string day = num.Substring(0, 2);
            string month = num.Substring(2, 2);
            string year = num.Substring(4, 2);
            string individual = num.Substring(6, 3);
            string k1 = num.Substring(9, 1);
            string k2 = num.Substring(10, 1);



            //Her kan du validere litt dato. Jeg gidder ikke skrive det nå, men du kan ta med sjekk på antall dager i den 
            // aktuelle mnd, osv... Husk at i et D-nummer legges det til 40 på dagen.

            //dersom individnummeret er mellom 500 og 750 er vedkommende født mellom 1855 og 1899
            if (Convert.ToInt32(individual) > 500 && Convert.ToInt32(individual) < 750)
                result = "Er du sikker på at denne personen er født FØR 1900? - ";

            //individnummerets tredje siffer bestemmer kjønn. partall: kvinne, oddetall: mann
            if (Convert.ToInt32(individual.Substring(2, 1)) % 2 == 0)
                result = "Dette er en kvinne - ";
            else
                result = "Dette er en mann - ";

            //Deler opp alle sifferne i hver sin int (bare for å gjøre utregningen lettere)
            int d1 = Convert.ToInt32(day.Substring(0, 1));
            int d2 = Convert.ToInt32(day.Substring(1, 1));
            int m1 = Convert.ToInt32(month.Substring(0, 1));
            int m2 = Convert.ToInt32(month.Substring(1, 1));
            int y1 = Convert.ToInt32(year.Substring(0, 1));
            int y2 = Convert.ToInt32(year.Substring(1, 1));
            int i1 = Convert.ToInt32(individual.Substring(0, 1));
            int i2 = Convert.ToInt32(individual.Substring(1, 1));
            int i3 = Convert.ToInt32(individual.Substring(2, 1));

            //Regner ut k1 (første kontrollsiffer)
            int k1Calculated = 11 -
                               (((3 * d1) + (7 * d2) + (6 * m1) + (1 * m2) + (8 * y1) + (9 * y2) + (4 * i1) + (5 * i2) + (2 * i3)) % 11);
            k1Calculated = (k1Calculated == 11 ? 0 : k1Calculated);

            //fødselsnummer som ville gitt k1 = 10 tildeles ikke
            if (k1Calculated == 10)
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k1 kan aldri bli 10";
                return new Tuple<bool, string>(false, result);
            }

            //Sjekker om den utregnede k1 er den samme som den som er tastet inn
            if (k1Calculated != Convert.ToInt32(k1))
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k1 feil!";
                return new Tuple<bool, string>(false, result);
            }

            //regner ut k2 (andre kontrolliffer)
            int k2Calculated = 11 -
                               (((5 * d1) + (4 * d2) + (3 * m1) + (2 * m2) + (7 * y1) + (6 * y2) + (5 * i1) + (4 * i2) + (3 * i3) +
                                 (2 * k1Calculated)) % 11);
            k2Calculated = (k2Calculated == 11 ? 0 : k2Calculated);

            //fødselsnummer som ville gitt k2 = 10 tildeles ikke
            if (k2Calculated == 10)
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k2 kan aldri bli 10";
                return new Tuple<bool, string>(false, result);
            }

            //sjekker om den utregnede k2 er den samme som den som er tatet inn
            if (k2Calculated != Convert.ToInt32(k2))
            {
                result = "Fødselsnummer er ikke korrekt!";
                //result += "k2 feil";
                return new Tuple<bool, string>(false, result);
            }

            //siden alle feil returnerer test-funksjonen, så har den aldrå nå passert :)
            result += "Passerte alle tester";
            return new Tuple<bool, string>(true, result);
        }

        //Read Product Codes File
        public List<Tuple<string, string>> Read(string resourceName)
        {
            XDocument xDoc = XDocument.Load(resourceName);
            //XDocument xDoc = XDocument.Parse(resourceName);

            var result = from b in xDoc.Descendants("DocumentCodes")
                         select new
                         {
                             code = b.Element("BREVKODE").Value,
                             description = b.Element("Beskrivelse").Value.Replace("/", "|").Replace(@"\", "|")
                         };

            List<Tuple<string, string>> productCodes = new List<Tuple<string, string>>();

            foreach (var row in result)
                productCodes.Add(new Tuple<string, string>(row.code, row.description));


            return productCodes;
        }

    }
}
