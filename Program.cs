using System.Collections;
using System.ComponentModel.DataAnnotations;
using System.Diagnostics;
using System.Diagnostics.Metrics;
using System.Linq;
using Google.Api.Gax.ResourceNames;
using Google.Cloud.Translate.V3;
using OfficeOpenXml;
using OfficeOpenXml.Sparkline;
using OfficeOpenXml.Style;
using static Google.Api.ResourceDescriptor.Types;

internal class Program
{
    private static async Task Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        Console.WriteLine("Wskaz plik raportu maintenance z dostępnych:");
        var file = new FileInfo(DisplayFilesInFolder());

        Console.Write("Wprowadz date w formacie DD.MM.RRR: ");
        string reportDate = Console.ReadLine();

        Console.WriteLine("Prosze czekac...");
        await ModifyAndSaveExcelFile(file, reportDate);
        Console.WriteLine("Gotowe!");


        //Awaits for user input to finish program.
        //Console.ReadKey();
    }

    private static async Task ModifyAndSaveExcelFile(FileInfo file, string date)
    {
        using var package = new ExcelPackage(file);

        var wsRaportDzienny = package.Workbook.Worksheets.First();

        int flag = 0;
        foreach (var item in package.Workbook.Worksheets)
        {
            if (item.Name == "Points") flag = 1;
        }

        if (flag == 1)
        {
            package.Workbook.Worksheets.Delete("Points");
            flag = 0;
        }

        package.Workbook.Worksheets.Add("Points");
        var wsPoints = package.Workbook.Worksheets.Last();

        int dziennyLineColumn;
        AssigningLines(wsRaportDzienny, 3, out dziennyLineColumn);
        //Coping column date - GOOD
        wsRaportDzienny.Cells[1, dziennyLineColumn, 1000, dziennyLineColumn].Copy(wsPoints.Cells[1, 5, 1000, 5]);

        //Coping column date - GOOD
        wsRaportDzienny.Cells[1, 1, 1000, 1].Copy(wsPoints.Cells[1, 7, 1000, 7]);

        //Coping column event time - GOOD
        wsRaportDzienny.Cells[1, 2, 1000, 2].Copy(wsPoints.Cells[1, 8, 1000, 8]);

        //Creating and coping column problem + remarks - GOOD
        for (int i = 1; i <= 1000; i++)
        {
            wsRaportDzienny.Cells[i, 50].Formula = "=" + wsRaportDzienny.Cells[i, 3].Address + "&\" - \"&" + wsRaportDzienny.Cells[i, 4];
        }
        wsRaportDzienny.Calculate();
        wsRaportDzienny.ClearFormulas();
        wsRaportDzienny.Cells[1, 50, 1000, 50].Copy(wsPoints.Cells[1, 2, 1000, 2]);

        //Coping column stop duration - GOOD
        wsRaportDzienny.Cells[1, 5, 1000, 5].Copy(wsPoints.Cells[1, 10, 1000, 10]);

        //Creating and coping column end time - GOOD
        for (int i = 1; i <= 1000; i++)
        {
            wsRaportDzienny.Cells[i, 51].Formula = "=" + wsRaportDzienny.Cells[i, 2].Address + "+" + wsRaportDzienny.Cells[i, 5];
        }
        wsRaportDzienny.Calculate();
        wsRaportDzienny.ClearFormulas();
        wsRaportDzienny.Cells[1, 51, 1000, 51].Copy(wsPoints.Cells[1, 9, 1000, 9]);

        //Coping column conveyors - GOOD
        wsRaportDzienny.Cells[1, 8, 1000, 8].Copy(wsPoints.Cells[1, 12, 1000, 12]);

        //Coping column comments - GOOD
        wsRaportDzienny.Cells[1, 9, 1000, 9].Copy(wsPoints.Cells[1, 4, 1000, 4]);

        DeletingOnesfullRows(wsPoints, date);
        //MoveShortFullDowntimes(wsPoints, 2, 10);
        ClearAllFormatting(wsPoints);
        AddCorrectFormating(wsPoints);
        FormatDateAndTime(wsPoints);
        //TranslateText(wsPoints, 2);
        
        //Saving
        await package.SaveAsync();
    }

    private static void FormatDateAndTime(ExcelWorksheet worksheet)
    {
        int dateC = 7;
        int timeStartC = 8;
        int timeEndC = 9;
        int timeDurationC = 10;

        var dateColumn = worksheet.Cells[1, dateC, 1000, dateC];
        var timeStartColumn = worksheet.Cells[1, timeStartC, 1000, timeStartC];
        var timeEndColumn = worksheet.Cells[1, timeEndC, 1000, timeEndC];
        var timeDurationColumn = worksheet.Cells[1, timeDurationC, 1000, timeDurationC];

        foreach (var item in dateColumn)
        {
            dateColumn.Style.Numberformat.Format = "yyyy-MM-dd";
            timeStartColumn.Style.Numberformat.Format = "HH:mm";
            timeEndColumn.Style.Numberformat.Format = "HH:mm";
            timeDurationColumn.Style.Numberformat.Format = "H:mm:ss";
        }

        worksheet.Cells[1, 1, 1000, 1000].Style.WrapText = true;
    }

    private static void AddCorrectFormating(ExcelWorksheet worksheet)
    {
        int lastColumn = worksheet.Dimension.Columns + 1;
        int lastRow = worksheet.Dimension.Rows;

        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        worksheet.Cells[1, 1, lastRow, lastColumn].Style.Border.Right.Style = ExcelBorderStyle.Thin;

        var problemColumn = worksheet.Cells[1, 5, 1000, 5];

        foreach (var item in problemColumn)
        {
            string itemValue = Convert.ToString(item.Value);
            string address = item.Address;

            if (itemValue.Contains("HC1-2") || itemValue.Contains("MV1-2") || itemValue.Contains("MV3-4"))
            {
                worksheet.Cells[address].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                worksheet.Cells[address].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
            }
        }
    }

    private static void TranslateText(ExcelWorksheet worksheet, int column)
    {
        TranslationServiceClient client = TranslationServiceClient.Create();
        TranslateTextRequest request = new TranslateTextRequest
        {
            Contents = { "It is raining." },
            TargetLanguageCode = "fr-FR",
            //Parent = new ProjectName(Program).ToString()
        };
        TranslateTextResponse response = client.TranslateText(request);
        // response.Translations will have one entry, because request.Contents has one entry.
        Translation translation = response.Translations[0];
        Console.WriteLine($"Detected language: {translation.DetectedLanguageCode}");
        Console.WriteLine($"Translated text: {translation.TranslatedText}");

    }

    private static void AssigningLines(ExcelWorksheet worksheet, int problemC, out int column)
    {
        //Column used to temporary store what line in maint report
        column = 52;

        //Checking what lines are available in report
        //string[] maintLines = { "PBS", "HC1-2", "MV1-2", "MV3-4", "LOADING LINE", "KOKPIT", "ED", "AGV HC", "AGV MV", "ZALEWANIE UKŁADÓW" };

        var maintLinesDictionary = new Dictionary<string, string>()
        {
            {"PBS", "PBS"},
            {"HC1-2", "HC1-2"},
            {"MV1-2", "MV1-2"},
            {"MV3-4", "MV3-4"},
            {"LOADING LINE", "OM2"},
            {"KOKPIT", "PB1"},
            {"ED", "EDU"},
            {"AGV HC", "AGV"},
            {"AGV MV", "AGV"},
            {"ZALEWANIE UKŁADÓW", "MV3"},
            {"DECKING", "DECKING"},
            {"UNDECKING", "UNDECKING"}
        };

        //string[] convLines = { "PBS", "HC1", "HC2", "MV1", "MV2", "DECKING", "UNDECKING", "MV3", "MV4", "OM2", "PB1", "AGV" };

        List<string> lines = new();
        var problemColumn = worksheet.Cells[1, problemC, 1000, problemC];
        string actualLine = "";
        int i = 1;
        int sihiFlag = 0;

        //Searching for used lines
        foreach (var item in problemColumn)
        {
            string itemValue = Convert.ToString(item.Value).ToUpper();

            //Assigning correct lines
            if      ((actualLine == "HC1-2" || actualLine == "HC1" || actualLine == "HC2") && itemValue.Contains("HC1")) actualLine = "HC1";
            else if ((actualLine == "HC1-2" || actualLine == "HC1" || actualLine == "HC2") && itemValue.Contains("HC2")) actualLine = "HC2";
            if      ((actualLine == "MV1-2" || actualLine == "MV1" || actualLine == "MV2" || actualLine == "DECKING" || actualLine == "UNDECKING") && itemValue.Contains("MV1")) actualLine = "MV1";
            else if ((actualLine == "MV1-2" || actualLine == "MV1" || actualLine == "MV2" || actualLine == "DECKING" || actualLine == "UNDECKING") && itemValue.Contains("MV2")) actualLine = "MV2";
            else if ((actualLine == "MV1-2" || actualLine == "MV1" || actualLine == "MV2" || actualLine == "DECKING" || actualLine == "UNDECKING") && itemValue.Contains("DECKING")) actualLine = "DECKING";
            else if ((actualLine == "MV1-2" || actualLine == "MV1" || actualLine == "MV2" || actualLine == "DECKING" || actualLine == "UNDECKING") && itemValue.Contains("UNDECKING")) actualLine = "UNDECKING";
            if      ((actualLine == "MV3-4" || actualLine == "MV3" || actualLine == "MV4") && itemValue.Contains("MV3")) actualLine = "MV3";
            else if ((actualLine == "MV3-4" || actualLine == "MV3" || actualLine == "MV4") && itemValue.Contains("MV4")) actualLine = "MV4";
            if      ((actualLine == "OM2" || actualLine == "LOADING LINE") && (itemValue.Conatins("SLT2600") || itemValue.Conatins("SLT 2600")) actualLine = "UNDECKING";

            //Assigning basic lines
            if (maintLinesDictionary.ContainsKey(itemValue.ToUpper()))
            {
                actualLine = maintLinesDictionary[itemValue.ToUpper()];
                lines.Add(actualLine);
                if (itemValue.Contains("ZALEWANIE UKŁADÓW"))
                {
                    sihiFlag = 1;
                }
            }

            worksheet.Cells[i, column].Value = actualLine;

            if (actualLine == "HC1" || actualLine == "HC2") actualLine = "HC1-2";
            if (actualLine == "MV1" || actualLine == "MV2" || actualLine == "DECKING" || actualLine == "UNDECKING") actualLine = "MV1-2";
            if ((actualLine == "MV3" || actualLine == "MV4") && sihiFlag == 1) actualLine = "MV3";
            else if ((actualLine == "MV3" || actualLine == "MV4") && sihiFlag == 0) actualLine = "MV3-4";

            i++;
        }

        sihiFlag = 0;
        //Assigning correct lines

    }

    private static void ClearAllFormatting(ExcelWorksheet worksheet)
    {
        var clearCell = worksheet.Cells["zz10000"];

        foreach (var item in worksheet.Cells.Where(item => item.Value != null || item.Style.Border != null))
        {
            clearCell.CopyStyles(item);
        }
    }

    private static void MoveShortFullDowntimes(ExcelWorksheet worksheet, int text, int downtime)
    {
        string[] downtimeColumn = new string[1000];
        string[] textColumn = new string[1000];
        var columnD = worksheet.Cells[1, downtime, 1000, downtime];
        var columnT = worksheet.Cells[1, text, 1000, text];
        int i = 0;

        foreach (var item in columnT)
        {
            textColumn[i] = item.GetValue<string>().ToLower();

            if (textColumn[i].Contains("short") || textColumn[i].Contains("full"))
            {
                worksheet.Cells[i + 1, downtime].Copy(worksheet.Cells[i + 1, downtime + 1]);
                worksheet.Cells[i + 1, downtime].Clear();
            }

            i++;
        }
    }

    private static void DeletingOnesfullRows(ExcelWorksheet worksheet, string reportDate)
    {
        string date, conv;
        int length;

        for (int i = 1; i <= 1000; i++)
        {
            date = Convert.ToString(worksheet.Cells[i, 7].Value);
            conv = Convert.ToString(worksheet.Cells[i, 12].Value);   

            if (date != "" && conv == "")
            {
                worksheet.Cells[i, 12].Value = "n";
            }

            length = date.IndexOf(" ");
            if (length > 0)
            {
                date = date.Substring(0, length);
            }

            
            if (!string.Equals(date, reportDate) && worksheet.Cells[i, 2].Value != null)
            {
                worksheet.DeleteRow(i);
                i--;
            }
        }
    }

    private static string DisplayFilesInFolder()
    {
        Process currentProcess = Process.GetCurrentProcess();

        string actualFolder = Path.GetDirectoryName(currentProcess.MainModule.FileName);
        string[] files = Directory.GetFiles(actualFolder);

        string path;

        for (int i = 1; i <= files.Length; i++)
        {
            path = files[i-1];
            Console.WriteLine(i + ": " + System.IO.Path.GetFileName(path));
        }

        Console.Write("Twoj wybor: ");
        int fileNumber = int.Parse(Console.ReadLine());

        return files[fileNumber - 1];
    }
}
