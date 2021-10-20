using OfficeOpenXml;
using RapportTewerkstellingCorona.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml.Table;
using OfficeOpenXml.Attributes;

namespace RapportTewerkstellingCorona
{
    class Program
    {
        static async Task Main(string[] args)
        {
            #region CONNECTIONS
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var savefile = new FileInfo(@"D:\Ruben Spillebeen\Documents\thuis\Papa\FODWASOEXCELdata.xlsx");
            var editfile = new FileInfo(@"D:\Ruben Spillebeen\Documents\thuis\Papa\FODWASOEXCEL.xlsx");
            DeleteIfExists(editfile);
            #endregion
            List<Tewerkstellingslijn> lijnen = await LoadExcelFile(savefile);
            Console.WriteLine("done loading");
            List<SubTotalTW> subTotals = new List<SubTotalTW>();
            List<SubTotalTW> prcSubTotals = new List<SubTotalTW>();
            RemoveFalse(lijnen);
            await SaveExcelFile(lijnen, editfile, "RemoveFalseLines");
            //Console.WriteLine($"{"PC",-8}|{"WerknemerTypeCaptionNL",-50}|{"Jaar",-10}|{"Week",-10}|{"TotaalWerkloosheid",-25}|{"TotaalZiekte",-15}|{"GrandTotal",-15}");
            RemoveDuplicates(lijnen);
            await SaveExcelFile(lijnen, editfile, "RemoveDubplicates");
            //AddMissing(lijnen);
            lijnen.Sort();
            await SaveExcelFile(lijnen, editfile, "SortedList");
            await SaveExcelFile(lijnen, editfile, "Aantal Dagen");
            GetSubTotals(lijnen, subTotals);
            await SaveExcelFile(subTotals, editfile, "totaal dagen");
            //ShowSubTotals(subTotals);
            List<Tewerkstellingslijn> prclijnen = GetPercentages(lijnen);
            await SaveExcelFile(prclijnen, editfile, "% Dagen");
            prcSubTotals = GetPercentages(subTotals);
            await SaveExcelFile(prcSubTotals, editfile, "totaal % Dagen");
            //ShowSubTotals(prcSubTotals);
            //ShowLijnen(lijnen);
            //ShowLijnen(prclijnen);
            //Console.WriteLine(lijnen.Count);
            ReorderSheets(editfile);
            Console.WriteLine("Edits done");
            Console.ReadLine();
        }
        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists)
            {
                file.Delete();
            }
        }

        private static void ReorderSheets(FileInfo file)
        {
            using (var package = new ExcelPackage(file))
            {
                package.Workbook.Worksheets.MoveToStart("aantal dagen");
                package.Workbook.Worksheets.MoveAfter("% dagen", "aantal dagen");
                package.Workbook.Worksheets.MoveAfter("totaal dagen", "% dagen");
                package.Workbook.Worksheets.MoveAfter("totaal % dagen", "totaal dagen");
            };
        }

        private static async Task SaveExcelFile(List<Tewerkstellingslijn> tewerkstellingslijns, FileInfo file, string wsName)
        {
            using (var package = new ExcelPackage(file))
            {
                MemberInfo[] membersToInclude = typeof(Tewerkstellingslijn)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => !Attribute.IsDefined(p, typeof(EpplusIgnore)))
                .ToArray();
                if (package.Workbook.Worksheets.Any(sheet => sheet.Name == wsName))
                    package.Workbook.Worksheets.Delete(wsName);
                var ws = package.Workbook.Worksheets.Add(wsName);
                var range = ws.Cells["A1"].LoadFromCollection(tewerkstellingslijns, true, TableStyles.None, BindingFlags.Public, membersToInclude);
                range.AutoFitColumns();
                await package.SaveAsync();
            }
        }
        private static async Task SaveExcelFile(List<SubTotalTW> subTotalTWs, FileInfo file, string wsName)
        {
            using (var package = new ExcelPackage(file))
            {
                MemberInfo[] membersToInclude = typeof(SubTotalTW)
                .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                .Where(p => !Attribute.IsDefined(p, typeof(EpplusIgnore)))
                .ToArray();
                if (package.Workbook.Worksheets.Any(sheet => sheet.Name == wsName))
                    package.Workbook.Worksheets.Delete(wsName);
                var ws = package.Workbook.Worksheets.Add(wsName);
                var range = ws.Cells["A1"].LoadFromCollection(subTotalTWs, true, TableStyles.None, BindingFlags.Public, membersToInclude);
                range.AutoFitColumns();
                await package.SaveAsync();
            }
        }
        private static void ShowSubTotals(List<SubTotalTW> subTotals)
        {
            foreach (var item in subTotals)
            {
                Console.WriteLine(item);
            }
        }

        private static void GetSubTotals(List<Tewerkstellingslijn> lijnen, List<SubTotalTW> subTotalTWs)
        {
            var query = lijnen.GroupBy(x => new { x.Jaar, x.Week, x.WerknemerTypeCaptionNL })
                .Select(g => new
                {
                    returnlist = g.ToList()
                });
            foreach (var group in query)
            {
                subTotalTWs.Add(new SubTotalTW()
                {
                    WerknemerTypeCaptionNL = group.returnlist.First().WerknemerTypeCaptionNL,
                    Jaar = group.returnlist.First().Jaar,
                    Week = group.returnlist.First().Week,
                    AndereAfwezigheid = group.returnlist.Sum(x => x.AndereAfwezigheid),
                    Gepresteerd = group.returnlist.Sum(x => x.Gepresteerd),
                    GewoneEconomischeWerkloosheid = group.returnlist.Sum(x => x.GewoneEconomischeWerkloosheid),
                    WerkloosheidCorona = group.returnlist.Sum(x => x.WerkloosheidCorona),
                    ZiekteGewaarborgdLoon = group.returnlist.Sum(x => x.ZiekteGewaarborgdLoon),
                    ZiekteNa1Jaar = group.returnlist.Sum(x => x.ZiekteNa1Jaar),
                    ZiekteNaGewaarborgdLoon = group.returnlist.Sum(x => x.ZiekteNaGewaarborgdLoon),
                    GrandTotal = group.returnlist.Sum(x => x.GrandTotal)
                });
            }
        }

        private static void RemoveFalse(List<Tewerkstellingslijn> lijnen)
        {
            lijnen.RemoveAll(x => x.WerknemerTypeCaptionNL.ToLower() != "arbeider" && x.WerknemerTypeCaptionNL.ToLower() != "bediende");
            lijnen.RemoveAll(x => x.PC.StartsWith("1") && x.WerknemerTypeCaptionNL.ToLower() == "bediende");
            lijnen.RemoveAll(x => x.PC.StartsWith("2") && x.WerknemerTypeCaptionNL.ToLower() == "arbeider");
        }
        //TODO: reference types xdxp prclijnen = lijnen
        private static List<Tewerkstellingslijn> GetPercentages(List<Tewerkstellingslijn> lijnen)
        {
            List<Tewerkstellingslijn> returnlist = new List<Tewerkstellingslijn>();
            Tewerkstellingslijn tl = new Tewerkstellingslijn();
            foreach (var lijn in lijnen)
            {
                tl = lijn;
                tl.AndereAfwezigheid = lijn.AndereAfwezigheid / lijn.GrandTotal;
                tl.Gepresteerd = lijn.Gepresteerd / lijn.GrandTotal;
                tl.GewoneEconomischeWerkloosheid = lijn.GewoneEconomischeWerkloosheid / lijn.GrandTotal;
                tl.WerkloosheidCorona = lijn.WerkloosheidCorona / lijn.GrandTotal;
                tl.ZiekteGewaarborgdLoon = lijn.ZiekteGewaarborgdLoon / lijn.GrandTotal;
                tl.ZiekteNa1Jaar = lijn.ZiekteNa1Jaar / lijn.GrandTotal;
                tl.ZiekteNaGewaarborgdLoon = lijn.ZiekteNaGewaarborgdLoon / lijn.GrandTotal;
                tl.GrandTotal = lijn.GrandTotal / lijn.GrandTotal;
                returnlist.Add(tl);
            }
            return returnlist;
        }
        private static List<SubTotalTW> GetPercentages(List<SubTotalTW> lijnen)
        {
            List<SubTotalTW> returnlist = new List<SubTotalTW>();
            SubTotalTW tl = new SubTotalTW();
            foreach (var lijn in lijnen)
            {
                tl = lijn;
                tl.AndereAfwezigheid = lijn.AndereAfwezigheid / lijn.GrandTotal;
                tl.Gepresteerd = lijn.Gepresteerd / lijn.GrandTotal;
                tl.GewoneEconomischeWerkloosheid = lijn.GewoneEconomischeWerkloosheid / lijn.GrandTotal;
                tl.WerkloosheidCorona = lijn.WerkloosheidCorona / lijn.GrandTotal;
                tl.ZiekteGewaarborgdLoon = lijn.ZiekteGewaarborgdLoon / lijn.GrandTotal;
                tl.ZiekteNa1Jaar = lijn.ZiekteNa1Jaar / lijn.GrandTotal;
                tl.ZiekteNaGewaarborgdLoon = lijn.ZiekteNaGewaarborgdLoon / lijn.GrandTotal;
                tl.GrandTotal = lijn.GrandTotal / lijn.GrandTotal;
                returnlist.Add(tl);
            }
            return returnlist;
        }
        private static void ShowLijnen(List<Tewerkstellingslijn> lijnen)
        {
            foreach (var item in lijnen)
            {
                Console.WriteLine(item);
            }
        }

        private static void RemoveDuplicates(List<Tewerkstellingslijn> lijnen)
        {
            for (int i = 0; i < lijnen.Count; i++)
            {
                for (int j = 0; j < lijnen.Count; j++)
                {
                    if (lijnen[i].Equals(lijnen[j]) && j > i)
                    {
                        lijnen[i].AndereAfwezigheid += lijnen[j].AndereAfwezigheid;
                        lijnen[i].Gepresteerd += lijnen[j].Gepresteerd;
                        lijnen[i].GewoneEconomischeWerkloosheid += lijnen[j].GewoneEconomischeWerkloosheid;
                        lijnen[i].WerkloosheidCorona += lijnen[j].WerkloosheidCorona;
                        lijnen[i].ZiekteGewaarborgdLoon += lijnen[j].ZiekteGewaarborgdLoon;
                        lijnen[i].ZiekteNa1Jaar += lijnen[j].ZiekteNa1Jaar;
                        lijnen[i].ZiekteNaGewaarborgdLoon += lijnen[j].ZiekteNaGewaarborgdLoon;
                        lijnen[i].GrandTotal += lijnen[j].GrandTotal;
                        lijnen.RemoveAt(j);
                        j--;
                    }
                }
            }
        }

        private static async Task<List<Tewerkstellingslijn>> LoadExcelFile(FileInfo file)
        {
            List<Tewerkstellingslijn> output = new List<Tewerkstellingslijn>();
            using (var package = new ExcelPackage(file))
            {
                await package.LoadAsync(file);
                var ws = package.Workbook.Worksheets["aantal dagen"];
                int row = 2;
                int col = 1;
                while (string.IsNullOrWhiteSpace(ws.Cells[row, col].Value?.ToString()) == false)
                {
                    Tewerkstellingslijn lijn = new Tewerkstellingslijn()
                    {
                        OfficieelPC = ws.Cells[row, col].Value.ToString(),
                        OfficieelSubPC = ws.Cells[row, col + 1].Value.ToString(),
                        AcertaPC = ws.Cells[row, col + 2].Value.ToString(),
                        WerknemerTypeCaptionNL = ws.Cells[row, col + 3].Value.ToString(),
                        Jaar = ws.Cells[row, col + 4].Value.ToString().ToNullableInt(),
                        Week = ws.Cells[row, col + 5].Value.ToString().ToNullableInt(),
                        AndereAfwezigheid = ws.Cells[row, col + 6].Value.ToString().ToNullableDecimal(),
                        Gepresteerd = ws.Cells[row, col + 7].Value.ToString().ToNullableDecimal(),
                        GewoneEconomischeWerkloosheid = ws.Cells[row, col + 8].Value.ToString().ToNullableDecimal(),
                        WerkloosheidCorona = ws.Cells[row, col + 10].Value.ToString().ToNullableDecimal(),
                        ZiekteGewaarborgdLoon = ws.Cells[row, col + 11].Value.ToString().ToNullableDecimal(),
                        ZiekteNa1Jaar = ws.Cells[row, col + 12].Value.ToString().ToNullableDecimal(),
                        ZiekteNaGewaarborgdLoon = ws.Cells[row, col + 13].Value.ToString().ToNullableDecimal(),
                        GrandTotal = ws.Cells[row, col + 14].Value.ToString().ToNullableDecimal(),
                    };
                    output.Add(lijn);
                    row++;
                }
                return output;
            }
        }
    }
    public static class StringExtensions
    {
        public static decimal ToNullableDecimal(this string s)
        {
            decimal i;
            if (decimal.TryParse(s, out i)) return i;
            return 0;
        }
        public static int ToNullableInt(this string s)
        {
            int i;
            if (int.TryParse(s, out i)) return i;
            return 0;
        }
    }
}
