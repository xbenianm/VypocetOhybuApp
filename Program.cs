using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using ClosedXML.Excel;

namespace VypocetOhybuApp
{
    class Program
    {
        static string excelPath = "VypocetOhybu.xlsx";
        static string vdiePath = "optimal_vdie.xlsx";

        static List<string> ZoznamStrojov = new List<string> { "Trumpf", "Amada", "Ursviken" };
        static List<string> ZoznamMaterialov = new List<string> { "11373", "Pozink", "AISI", "Meď", "Hliník", "RAL" };

        static double VypocitajToleranciu(double dlzka)
        {
            if (dlzka <= 3) return 0.1;
            else if (dlzka <= 6) return 0.2;
            else if (dlzka <= 30) return 0.5;
            else if (dlzka <= 120) return 1.0;
            else if (dlzka <= 400) return 1.5;
            else if (dlzka <= 1000) return 2.5;
            else return 4.0;
        }

        static string VyberZoZoznamu(List<string> zoznam, string prompt)
        {
            Console.WriteLine(prompt);
            for (int i = 0; i < zoznam.Count; i++)
                Console.WriteLine($"{i + 1} - {zoznam[i]}");
            int volba = int.Parse(Console.ReadLine());
            return zoznam[volba - 1];
        }

        static double ZiskajOptimalneVdie(double hrubka)
        {
            var workbook = new XLWorkbook(vdiePath);
            var ws = workbook.Worksheet(1);
            double najblizsiaRozdiel = double.MaxValue;
            double odporucaneVdie = 0;

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                double h = row.Cell(1).GetDouble();
                double v = row.Cell(2).GetDouble();
                double rozdiel = Math.Abs(h - hrubka);
                if (rozdiel < najblizsiaRozdiel)
                {
                    najblizsiaRozdiel = rozdiel;
                    odporucaneVdie = v;
                }
            }

            return odporucaneVdie;
        }

        static void NovyVypocet()
        {
            string stroj = VyberZoZoznamu(ZoznamStrojov, "Zvoľ stroj:");
            string material = VyberZoZoznamu(ZoznamMaterialov, "Zvoľ materiál:");

            Console.Write("Hrúbka plechu (T): ");
            double T = Convert.ToDouble(Console.ReadLine());

            Console.Write("Uhol ohybu (°): ");
            double uhol = Convert.ToDouble(Console.ReadLine());

            Console.Write("Rameno A: ");
            double ramenoA = Convert.ToDouble(Console.ReadLine());

            Console.Write("Rameno B: ");
            double ramenoB = Convert.ToDouble(Console.ReadLine());

            Console.Write("Typ OWR (napr. 0.5, 1.0, 2.0): ");
            double owr = Convert.ToDouble(Console.ReadLine());

            Console.Write("Zvoliť V-die automaticky podľa hrúbky? (y/n): ");
            string autoVdie = Console.ReadLine().Trim().ToLower();

            double vdie;
            if (autoVdie == "y")
            {
                vdie = ZiskajOptimalneVdie(T);
                Console.WriteLine($"➡️ Automaticky zvolené V-die: {vdie}");
            }
            else
            {
                Console.Write("Zadaj V-die (napr. 12 alebo W12): ");
                string vdieInput = Console.ReadLine().Trim();
                if (vdieInput.StartsWith("W")) vdieInput = vdieInput.Substring(1);
                vdie = Convert.ToDouble(vdieInput, CultureInfo.InvariantCulture);
            }

            Console.Write("Spôsob výpočtu R (1 = Podmienka podľa V-die, 2 = 0.16 × V-die, 3 = zadaj vlastné R): ");
            int rmetoda = Convert.ToInt32(Console.ReadLine());

            double R;
            double korekcia = (material == "AISI") ? 0.4 : (material == "Meď") ? 0.0 : (material == "Hliník") ? 0.1 : 0.2;

            if (rmetoda == 1)
            {
                double hranica = (owr + korekcia) / 0.16;
                if (vdie < hranica)
                {
                    Console.WriteLine($"➡️ Vdie ({vdie}) < (OWR + korekcia)/0.16 = {hranica:F2} → Použije sa R = OWR + korekcia");
                    R = owr + korekcia;
                }
                else
                {
                    Console.WriteLine($"➡️ Vdie ({vdie}) ≥ (OWR + korekcia)/0.16 = {hranica:F2} → Použije sa R = 0.16 × Vdie");
                    R = 0.16 * vdie;
                }
            }
            else if (rmetoda == 2)
            {
                R = 0.16 * vdie;
            }
            else
            {
                Console.Write("Zadaj vlastné R: ");
                R = Convert.ToDouble(Console.ReadLine());
            }

            Console.Write("BD pri 90° (BD90): ");
            double BD90 = Convert.ToDouble(Console.ReadLine());

            Console.Write("Použiť špeciálny OSSB výpočet aj pre väčšie a rovne ako 90°? (y/n): ");
            bool special = Console.ReadLine().Trim().ToLower() == "y";

            double BA90 = 2 * (R + T) * Math.Tan(Math.PI / 4) - BD90;
            double kfactor = (BA90 / (Math.PI / 2 * T)) - (R / T);

            double radian = (180 - uhol) * Math.PI / 180.0;
            double BA = (R + kfactor * T) * radian;

            double ossb = (uhol < 90 || special)
                ? Math.Tan((180 - uhol) / 2 * Math.PI / 180.0) * (R + T)
                : R + kfactor * T;

            double BD = 2 * ossb - BA;

            double x_min = vdie / 2 + 1.5 * T;
            double minRameno = x_min;
            double minRamenoS = x_min + 2 * T;

            double tolA = VypocitajToleranciu(ramenoA);
            double tolB = VypocitajToleranciu(ramenoB);

            Console.WriteLine($"\n📐 R = {R:F2} mm");
            Console.WriteLine($"📏 BD = {BD:F2} mm");
            Console.WriteLine($"📏 BA = {BA:F2} mm");
            Console.WriteLine($"📌 OSSB = {ossb:F2} mm");
            Console.WriteLine($"📌 K-faktor = {kfactor:F4}");
            Console.WriteLine($"📌 Tolerancia A = ±{tolA} mm");
            Console.WriteLine($"📌 Tolerancia B = ±{tolB} mm");
            Console.WriteLine($"📌 Min. rameno = {minRameno:F2} mm");
            Console.WriteLine($"📌 Min. rameno + s = {minRamenoS:F2} mm");

            var workbook = new XLWorkbook(excelPath);
            var ws = workbook.Worksheets.Contains("Výpočty") ? workbook.Worksheet("Výpočty") : workbook.AddWorksheet("Výpočty");

            if (ws.Cell(1, 1).IsEmpty())
            {
                string[] hlavicky = {
                    "Stroj", "Materiál", "Hrúbka", "Uhol", "OWR", "V-die", "R", "BD α°",
                    "Rameno A", "Tol. A", "Rameno B", "Tol. B", "Min x", "BA", "K-faktor", "Dátum"
                };
                for (int i = 0; i < hlavicky.Length; i++)
                    ws.Cell(1, i + 1).Value = hlavicky[i];
            }

            int row = ws.LastRowUsed().RowNumber() + 1;
            ws.Cell(row, 1).Value = stroj;
            ws.Cell(row, 2).Value = material;
            ws.Cell(row, 3).Value = T;
            ws.Cell(row, 4).Value = uhol;
            ws.Cell(row, 5).Value = owr;
            ws.Cell(row, 6).Value = vdie;
            ws.Cell(row, 7).Value = R;
            ws.Cell(row, 8).Value = BD;
            ws.Cell(row, 9).Value = ramenoA;
            ws.Cell(row, 10).Value = tolA;
            ws.Cell(row, 11).Value = ramenoB;
            ws.Cell(row, 12).Value = tolB;
            ws.Cell(row, 13).Value = x_min;
            ws.Cell(row, 14).Value = BA;
            ws.Cell(row, 15).Value = kfactor;
            ws.Cell(row, 16).Value = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");

            workbook.SaveAs(excelPath);
            Console.WriteLine("✅ Výsledok uložený do Excelu.");
        }

        static void ZobrazZaznamy()
        {
            if (!File.Exists(excelPath))
            {
                Console.WriteLine("❌ Súbor neexistuje.");
                return;
            }

            var workbook = new XLWorkbook(excelPath);
            var ws = workbook.Worksheet("Výpočty");

            Console.WriteLine("\n📄 ZOZNAM ZÁZNAMOV:\n");

            var hlavicka = ws.Row(1).Cells(1, 16).Select(c => c.GetString()).ToList();
            Console.WriteLine(string.Join(" | ", hlavicka));
            Console.WriteLine(new string('-', 160));

            foreach (var row in ws.RowsUsed().Skip(1))
            {
                List<string> hodnoty = new List<string>();
                for (int i = 1; i <= 16; i++)
                {
                    hodnoty.Add(row.Cell(i).GetFormattedString());
                }
                Console.WriteLine(string.Join(" | ", hodnoty));
            }
        }

        static void VymazZaznam()
        {
            if (!File.Exists(excelPath))
            {
                Console.WriteLine("❌ Súbor neexistuje.");
                return;
            }

            var workbook = new XLWorkbook(excelPath);
            var ws = workbook.Worksheet("Výpočty");
            int pocetRiadkov = ws.LastRowUsed().RowNumber();

            ZobrazZaznamy();
            Console.Write("\nZadaj číslo záznamu na vymazanie: ");
            if (!int.TryParse(Console.ReadLine(), out int index) || index < 1 || index >= pocetRiadkov)
            {
                Console.WriteLine("❌ Neplatný index.");
                return;
            }

            ws.Row(index + 1).Delete();
            workbook.SaveAs(excelPath);
            Console.WriteLine("🗑️ Záznam bol vymazaný.");
        }

        static void ExportDoCsv()
        {
            if (!File.Exists(excelPath))
            {
                Console.WriteLine("❌ Súbor neexistuje.");
                return;
            }

            var workbook = new XLWorkbook(excelPath);
            var ws = workbook.Worksheet("Výpočty");

            string csvPath = Path.ChangeExtension(excelPath, ".csv");
            using (var writer = new StreamWriter(csvPath))
            {
                foreach (var row in ws.RowsUsed())
                {
                    var hodnoty = row.Cells().Select(c => $"\"{c.GetValue<string>().Replace("\"", "\"\"")}\"");
                    writer.WriteLine(string.Join(",", hodnoty));
                }
            }

            Console.WriteLine($"✅ Exportované do {csvPath}");
        }

        static void Main()
        {
            while (true)
            {
                Console.WriteLine("\n=== HLAVNÉ MENU ===");
                Console.WriteLine("1 - Nový výpočet");
                Console.WriteLine("2 - Zobraziť všetky záznamy");
                Console.WriteLine("3 - Vymazať záznam");
                Console.WriteLine("4 - Export do CSV");
                Console.WriteLine("5 - Koniec");
                Console.Write("Zvoľ možnosť: ");

                string volba = Console.ReadLine().Trim();
                if (volba == "1") NovyVypocet();
                else if (volba == "2") ZobrazZaznamy();
                else if (volba == "3") VymazZaznam();
                else if (volba == "4") ExportDoCsv();
                else if (volba == "5") break;
                else Console.WriteLine("❌ Neplatná voľba.");
            }
        }
    }
}







