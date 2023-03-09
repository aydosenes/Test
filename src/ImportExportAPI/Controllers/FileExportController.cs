using ClosedXML.Excel;
using ImportExportAPI.Model.DataModel;
using ImportExportAPI.Utils;
using ImportExportAPI.Utils.ExcelUtils;
using Microsoft.AspNetCore.Mvc;
using Spire.Xls;

namespace ImportExportAPI.Controllers;

[ApiController]
[Route("[controller]")]
public class FileExportController : ControllerBase
{
    private readonly ILogger<FileExportController> _logger;

    public FileExportController(ILogger<FileExportController> logger)
    {
        _logger = logger;
    }

    [HttpPost]
    [Route("ExcelExport")]
    public IActionResult FileCreate(IFormFile personelCardsFile,
        IFormFile projectCardsFile,
        Boolean convertPdf = true,
        String filename = "Çıktı",
        Boolean doneWorks = true)
    {
        try
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            PersonelCard personelCard = ExcelRead.readPersonelCardData(personelCardsFile);
            _logger.LogInformation("Personel Başarıyla Kartları Okundu");
            ProjectCard projectCard = ExcelRead.readProjectCardData(projectCardsFile);
            _logger.LogInformation("Proje Başarıyla Kartları Okundu");

            if (!doneWorks)
            {
                projectCard.ProjectList.Clear();
                projectCard.ProjectList.AddRange(projectCard.CurrentProjectList);
            }
            else
            {
                projectCard.ProjectList.Clear();
                projectCard.ProjectList.AddRange(projectCard.CurrentProjectList);
                projectCard.ProjectList.AddRange(projectCard.DoneProjectList);
            }


            using (var workbook = new XLWorkbook())
            {
                var doneWorkColor = XLColor.FromArgb(211, 211, 211);
                var yapimIsi_Color = XLColor.FromArgb(245, 191, 157);
                var kdIsi_Color = XLColor.FromArgb(186, 219, 165);
                var gpIsi_Color = XLColor.FromArgb(176, 206, 234);
                var generalWorkerColor = XLColor.FromArgb(255, 240, 193);

                var worksheet = workbook.Worksheets.Add(filename);

                var startRowIndex = 1;
                var startColumnIndex = 1;
                var endRowIndex = personelCard.PersonelList.Count + 6;
                var endColIndex = 6 + projectCard.ProjectList.Count;

                var currentRow = 3;

                {
                    var lastIndexOfCurrentWorks = projectCard.ProjectList.Select(x => x.Status != "Tamamlandı").Count();

                    worksheet.Cell(currentRow - 2, 7).Value = projectCard.ProjectList.First().Directorate.ToUpper() +
                                                              " (DEVAM EDEN İŞLER)";
                    worksheet.Range(currentRow - 2, 7, currentRow, lastIndexOfCurrentWorks).Merge();
                    worksheet.Cell(currentRow - 2, 7).Style.Font.FontSize = 25;
                    if (doneWorks)
                    {
                        worksheet.Cell(currentRow - 2, lastIndexOfCurrentWorks + 1).Value =
                            projectCard.ProjectList.First().Directorate.ToUpper() + " (TAMAMLANAN İŞLER)";
                        worksheet.Cell(currentRow - 2, lastIndexOfCurrentWorks + 1).Style.Font.FontSize = 25;
                        worksheet.Range(currentRow - 2, lastIndexOfCurrentWorks + 1, currentRow,
                            7 + projectCard.ProjectList.Count() - 1).Merge();
                    }
                }

                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = "";
                    worksheet.Cell(currentRow, 2).Value = "";
                    worksheet.Cell(currentRow, 3).Value = "";
                    worksheet.Cell(currentRow, 4).Value = "";
                    worksheet.Cell(currentRow, 5).Value = "";
                    worksheet.Cell(currentRow, 6).Value = "";
                    for (int i = 0; i < projectCard.ProjectList.Count; i++)
                    {
                        worksheet.Cell(currentRow, 6 + i + 1).Value =
                            projectCard.ProjectList[i].Id + Environment.NewLine;
                    }
                }


                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = "Sıra No";
                    worksheet.Cell(currentRow, 2).Value = "Sicil No";
                    worksheet.Cell(currentRow, 3).Value = "Adı Soyadı";
                    worksheet.Cell(currentRow, 4).Value = "Mesleği";
                    worksheet.Cell(currentRow, 5).Value = "Görevi";
                    worksheet.Cell(currentRow, 6).Value = "Proje Sayısı";

                    for (int i = 0; i < projectCard.ProjectList.Count; i++)
                    {
                        worksheet.Cell(currentRow, 6 + i + 1).Value =
                            StringUtils.sliceString(projectCard.ProjectList[i].Name);
                        if (projectCard.ProjectList[i].Status == "Tamamlandı")
                            worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                                .SetBackgroundColor(doneWorkColor);
                        else if (projectCard.ProjectList[i].Type == "Yapım İşi")
                            worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                                .SetBackgroundColor(yapimIsi_Color);
                        else if (projectCard.ProjectList[i].Type == "Kentsel Dönüşüm İşi")
                            worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                                .SetBackgroundColor(kdIsi_Color);
                        else if (projectCard.ProjectList[i].Type == "Gelir Paylaşımı İşi")
                            worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                                .SetBackgroundColor(gpIsi_Color);
                    }
                }


                currentRow++;
                worksheet.Cell(currentRow, 1).Value = "";
                worksheet.Cell(currentRow, 2).Value = "";
                worksheet.Cell(currentRow, 3).Value = "";
                worksheet.Cell(currentRow, 4).Value = "";
                worksheet.Cell(currentRow, 5).Value = "";
                worksheet.Cell(currentRow, 6).Value = "";
                for (int i = 0; i < projectCard.ProjectList.Count; i++)
                {
                    worksheet.Cell(currentRow, 6 + i + 1).Value = StringUtils.sliceString(
                        projectCard.ProjectList[i].ShortName + "(" + projectCard.ProjectList[i].ContractorName + ")");
                    worksheet.Cell(currentRow, 6 + i + 1).Style.Font.FontColor = XLColor.Red;
                    if (projectCard.ProjectList[i].Status == "Tamamlandı")
                        worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                            .SetBackgroundColor(doneWorkColor);
                    else if (projectCard.ProjectList[i].Type == "Yapım İşi")
                        worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                            .SetBackgroundColor(yapimIsi_Color);
                    else if (projectCard.ProjectList[i].Type == "Kentsel Dönüşüm İşi")
                        worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                            .SetBackgroundColor(kdIsi_Color);
                    else if (projectCard.ProjectList[i].Type == "Gelir Paylaşımı İşi")
                        worksheet.Cell(currentRow, 6 + i + 1).AddConditionalFormat().WhenGreaterThan(0).Fill
                            .SetBackgroundColor(gpIsi_Color);
                }


                foreach (var personel in personelCard.PersonelList)
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = StringUtils.sliceStringInEverySpace(personel.OrderNo);
                    worksheet.Cell(currentRow, 2).Value = StringUtils.sliceStringInEverySpace(personel.RegistrationNo);
                    worksheet.Cell(currentRow, 3).Value = StringUtils.sliceStringInEverySpace(personel.FullName);
                    worksheet.Cell(currentRow, 4).Value = StringUtils.sliceStringInEverySpace(personel.Job);
                    worksheet.Cell(currentRow, 5).Value = StringUtils.sliceStringInEverySpace(personel.Role);
                    worksheet.Cell(currentRow, 6).Value =
                        StringUtils.sliceStringInEverySpace(personel.ProjectCount.ToString());
                    if (personel.IsGeneralWorker)
                    {
                        for (int i = 0; i < endColIndex; i++)
                        {
                            worksheet.Cell(currentRow, i + 1).Style.Fill.SetBackgroundColor(generalWorkerColor);
                        }
                    }


                    for (int i = 0; i < personel.Projects.Count; i++)
                    {
                        if (personel.Projects[i].Trim() == "")
                            continue;
                        int projectColIndex = -1;
                        ProjectRow project = new ProjectRow();
                        for (int index = 0; index < projectCard.ProjectList.Count; index++)
                        {
                            var prj = projectCard.ProjectList[index];
                            if (prj.Id == personel.Projects[i])
                            {
                                project = prj;
                                projectColIndex = index + 1;
                                break;
                            }
                        }


                        if (project.Status == "Tamamlandı")
                        {
                            worksheet.Cell(currentRow, 6 + projectColIndex).Value = "";
                            worksheet.Cell(currentRow, 6 + projectColIndex).AddConditionalFormat().WhenEquals("").Fill
                                .SetBackgroundColor(doneWorkColor);
                        }
                        else if (project.Type == "Yapım İşi")
                        {
                            worksheet.Cell(currentRow, 6 + projectColIndex).Value = "";
                            worksheet.Cell(currentRow, 6 + projectColIndex).AddConditionalFormat().WhenEquals("").Fill
                                .SetBackgroundColor(yapimIsi_Color);
                        }

                        else if (project.Type == "Kentsel Dönüşüm İşi")
                        {
                            worksheet.Cell(currentRow, 6 + projectColIndex).Value = "";
                            worksheet.Cell(currentRow, 6 + projectColIndex).AddConditionalFormat().WhenEquals("").Fill
                                .SetBackgroundColor(kdIsi_Color);
                        }

                        else if (project.Type == "Gelir Paylaşımı İşi")
                        {
                            worksheet.Cell(currentRow, 6 + projectColIndex).Value = "";
                            worksheet.Cell(currentRow, 6 + projectColIndex).AddConditionalFormat().WhenEquals("").Fill
                                .SetBackgroundColor(gpIsi_Color);
                        }
                    }
                }


                worksheet.CellsUsed().Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                worksheet.CellsUsed().Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);
                worksheet.CellsUsed().Style.Font.Bold = true;

                worksheet.PageSetup.PrintAreas.Add(1, 1, endRowIndex, endColIndex);


                var pageRange = worksheet.Range(4, 1, endRowIndex, endColIndex);

                pageRange.Style.Border.TopBorder = XLBorderStyleValues.Thin;
                pageRange.Style.Border.TopBorderColor = XLColor.Black;
                pageRange.Style.Border.BottomBorder = XLBorderStyleValues.Thin;
                pageRange.Style.Border.BottomBorderColor = XLColor.Black;
                pageRange.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                pageRange.Style.Border.RightBorderColor = XLColor.Black;
                pageRange.Style.Border.LeftBorder = XLBorderStyleValues.Thin;
                pageRange.Style.Border.LeftBorderColor = XLColor.Black;

                /*worksheet.Range(1,1,1, endColIndex).Style.Border.TopBorder = XLBorderStyleValues.Thick;
                worksheet.Range(1, 1, 1, endColIndex).Style.Border.TopBorderColor = XLColor.Blue;
                worksheet.Range(1,1, endRowIndex, 1).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                worksheet.Range(1, 1, endRowIndex, 1).Style.Border.LeftBorderColor = XLColor.Blue;
                worksheet.Range(endRowIndex, 1, endRowIndex, endColIndex).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                worksheet.Range(endRowIndex, 1, endRowIndex, endColIndex).Style.Border.BottomBorderColor = XLColor.Blue;
                worksheet.Range(1, endColIndex , endRowIndex, endColIndex).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                worksheet.Range(1, endColIndex, endRowIndex, endColIndex).Style.Border.RightBorderColor = XLColor.Blue;*/

                worksheet.ColumnsUsed().AdjustToContents();
                worksheet.RowsUsed().AdjustToContents();

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);

                    var FileDic = "files";
                    Directory.CreateDirectory(FileDic);
                    var tempGuid = Guid.NewGuid();

                    var tempExcelFilePath = $"{FileDic}/{tempGuid}.xlsx";
                    FileStream excelFile =
                        new FileStream(tempExcelFilePath, FileMode.Create, System.IO.FileAccess.Write);
                    stream.WriteTo(excelFile);

                    var content = stream.ToArray();

                    var exlsFile = File(
                        content,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        filename + ".xlsx");


                    if (!convertPdf)
                        return exlsFile;

                    Workbook spireWorkbbok = new Workbook();
                    spireWorkbbok.LoadFromFile(tempExcelFilePath);

                    for (int i = 0; i < workbook.Worksheets.Count; i++)
                    {
                        //spireWorkbbok.Worksheets[i].PageSetup.IsFitToPage = true;
                        spireWorkbbok.Worksheets[i].PageSetup.FitToPagesWide = 1;
                        spireWorkbbok.Worksheets[i].PageSetup.Orientation = PageOrientationType.Landscape;
                    }

                    var tempPdfFileName = $"{tempGuid}.pdf";
                    spireWorkbbok.SaveToFile($@"{FileDic}/{tempPdfFileName}", Spire.Xls.FileFormat.PDF);

                    byte[] fileBytes = System.IO.File.ReadAllBytes($"{FileDic}/{tempPdfFileName}");

                    return File(fileBytes, "application/force-download", filename + ".pdf");
                }
            }


            return BadRequest();
        }
        catch (Exception ex)
        {
            _logger.LogError(ex.StackTrace);
            return BadRequest(ex.StackTrace);
        }
    }


    [HttpPost]
    [Route("KuzuDailyReport")]
    public IActionResult KuzuDailyReport(
        IFormFile imalatPlanlanan,
        IFormFile imalatYapilan,
        IFormFile makinaParkı,
        IFormFile personelPuantaj,
        String filename = "Çıktı")
    {
        Workbook spireWorkbbok = new Workbook();
        spireWorkbbok.LoadFromFile("templateFiles/tempKuzuDailyOutput.xlsx");


        Worksheet sheet = spireWorkbbok.Worksheets[0];

        DateTime currentDateTime = DateTime.Now;

        DateTime projectStartDate = new DateTime(2022, 09, 19);
        DateTime projectFinishDate = new DateTime(2024, 09, 19);

        int workedDayCount = (int)(currentDateTime - projectStartDate).TotalDays;
        int leftDayCount = (int)(projectFinishDate - currentDateTime).TotalDays;

        sheet.SetCellValue(12, 9, leftDayCount.ToString());
        sheet.SetCellValue(8, 34, workedDayCount.ToString());


        /*List<ImalatPlanlanan> imalatPlanlanans = ExcelRead.readImalatPlanlan(imalatPlanlanan);
        List<ImalatYapilan> imalatYapilans = ExcelRead.readImalatYapilan(imalatYapilan);*/

        {
            List<PersonelPuantaj> personelPuantajs = ExcelRead.readPersonelPuantaj(personelPuantaj);
            List<PersonelPuantaj> endirekPersonel = new List<PersonelPuantaj>();
            List<PersonelPuantaj> direktPersonel = new List<PersonelPuantaj>();
            foreach (PersonelPuantaj item in personelPuantajs)
            {
                if (item.Durumu.ToLower() == "endirekt")
                {
                    endirekPersonel.Add(item);
                }
                else
                {
                    direktPersonel.Add(item);
                }
            }

            {
                sheet.SetCellValue(6, 36, endirekPersonel.Select(x => int.Parse(x.KisiSayısı) * 8).Sum().ToString());
                sheet.SetCellValue(6, 37, direktPersonel.Select(x => int.Parse(x.KisiSayısı) * 8).Sum().ToString());
                sheet.SetCellValue(6, 38, personelPuantajs.Select(x => int.Parse(x.KisiSayısı) * 8).Sum().ToString());

                sheet.SetCellValue(8, 36,
                    personelPuantajs.Where(x => x.Firma == "KUZU").Select(x => int.Parse(x.KisiSayısı) * 8).Sum()
                        .ToString());
                sheet.SetCellValue(8, 37,
                    personelPuantajs.Where(x => x.Firma != "KUZU").Select(x => int.Parse(x.KisiSayısı) * 8).Sum()
                        .ToString());
                sheet.SetCellValue(8, 38, personelPuantajs.Select(x => int.Parse(x.KisiSayısı) * 8).Sum().ToString());
            }

            {
                List<String> endirektTasks = new List<String>();
                List<String> altFirmas = new List<String>();
                foreach (PersonelPuantaj item in endirekPersonel)
                {
                    if (!endirektTasks.Contains(item.Görevi))
                    {
                        endirektTasks.Add(item.Görevi);
                    }

                    if (!altFirmas.Contains(item.Firma) && item.Firma.ToLower() != "kuzu")
                    {
                        altFirmas.Add(item.Firma);
                    }
                }

                for (int i = 0; i < altFirmas.Count; i++)
                {
                    sheet.SetCellValue(15, i + 5, altFirmas[i].ToString());
                }

                for (int i = 0; i < endirektTasks.Count; i++)
                {
                    String endirektTask = endirektTasks[i].ToString();
                    int totalSumForTask = 0;
                    sheet.SetCellValue(i + 16, 3, endirektTask);
                    int AnaSum = endirekPersonel
                        .Where(x => x.Görevi.ToLower() == endirektTask.ToLower() && x.Firma == "KUZU")
                        .Select(x => int.Parse(x.KisiSayısı)).Sum();
                    if (AnaSum > 0)
                        sheet.SetCellValue(i + 16, 4, AnaSum.ToString());
                    totalSumForTask += AnaSum;

                    for (int j = 0; j < altFirmas.Count; j++)
                    {
                        int sum = endirekPersonel.Where(x =>
                                x.Görevi.ToLower() == endirektTask.ToLower() && x.Firma == altFirmas[j].ToString())
                            .Select(x => int.Parse(x.KisiSayısı)).Sum();
                        if (sum > 0)
                            sheet.SetCellValue(i + 16, j + 5, sum.ToString());
                        totalSumForTask += sum;
                    }

                    sheet.SetCellValue(i + 16, 11, (totalSumForTask * 8).ToString());
                }

                int toplamEndirektCalisanSayisi = 0;
                int anaCalisanSayisi = endirekPersonel.Where(x => x.Firma == "KUZU")
                    .Select(x => int.Parse(x.KisiSayısı)).Sum();
                sheet.SetCellValue(58, 4, anaCalisanSayisi.ToString());
                sheet.SetCellValue(61, 4, anaCalisanSayisi.ToString());
                toplamEndirektCalisanSayisi += anaCalisanSayisi;

                for (int j = 0; j < altFirmas.Count; j++)
                {
                    int sum = endirekPersonel.Where(x => x.Firma == altFirmas[j].ToString())
                        .Select(x => int.Parse(x.KisiSayısı)).Sum();
                    if (sum > 0)
                    {
                        sheet.SetCellValue(58, j + 5, sum.ToString());
                        sheet.SetCellValue(61, j + 5, sum.ToString());
                    }


                    toplamEndirektCalisanSayisi += sum;
                }

                sheet.SetCellValue(58, 11, toplamEndirektCalisanSayisi.ToString());
                sheet.SetCellValue(61, 11, toplamEndirektCalisanSayisi.ToString());
            }

            {
                List<String> direktTasks = new List<String>();
                List<String> direktAltFirmas = new List<String>();
                foreach (PersonelPuantaj item in direktPersonel)
                {
                    if (!direktTasks.Contains(item.Görevi))
                    {
                        direktTasks.Add(item.Görevi);
                    }

                    if (!direktAltFirmas.Contains(item.Firma) && item.Firma.ToLower() != "kuzu")
                    {
                        direktAltFirmas.Add(item.Firma);
                    }
                }

                for (int i = 0; i < direktAltFirmas.Count; i++)
                {
                    sheet.SetCellValue(15, i + 14, direktAltFirmas[i].ToString());
                }

                for (int i = 0; i < direktTasks.Count; i++)
                {
                    String direktTask = direktTasks[i].ToString();
                    int totalSumForTask = 0;

                    sheet.SetCellValue(i + 16, 12, direktTask);
                    int AnaSum = direktPersonel
                        .Where(x => x.Görevi.ToLower() == direktTask.ToLower() && x.Firma == "KUZU")
                        .Select(x => int.Parse(x.KisiSayısı)).Sum();
                    if (AnaSum > 0)
                        sheet.SetCellValue(i + 16, 13, AnaSum.ToString());
                    totalSumForTask += AnaSum;

                    for (int j = 0; j < direktAltFirmas.Count; j++)
                    {
                        int sum = direktPersonel.Where(x =>
                                x.Görevi.ToLower() == direktTask.ToLower() && x.Firma == direktAltFirmas[j].ToString())
                            .Select(x => int.Parse(x.KisiSayısı)).Sum();
                        if (sum > 0)
                            sheet.SetCellValue(i + 16, j + 14, sum.ToString());

                        totalSumForTask += sum;

                        sheet.SetCellValue(i + 16, 33, (totalSumForTask * 8).ToString());
                    }
                }

                int toplamDirektCalisanSayisi = 0;
                int anaCalisanSayisi = direktPersonel.Where(x => x.Firma == "KUZU")
                    .Select(x => int.Parse(x.KisiSayısı)).Sum();
                sheet.SetCellValue(58, 13, anaCalisanSayisi.ToString());
                sheet.SetCellValue(61, 13, anaCalisanSayisi.ToString());
                toplamDirektCalisanSayisi += anaCalisanSayisi;

                for (int j = 0; j < direktAltFirmas.Count; j++)
                {
                    int sum = direktPersonel.Where(x => x.Firma == direktAltFirmas[j].ToString())
                        .Select(x => int.Parse(x.KisiSayısı)).Sum();
                    if (sum > 0)
                    {
                        sheet.SetCellValue(58, j + 14, sum.ToString());
                        sheet.SetCellValue(61, j + 14, sum.ToString());
                    }


                    toplamDirektCalisanSayisi += sum;
                }

                sheet.SetCellValue(58, 33, toplamDirektCalisanSayisi.ToString());
                sheet.SetCellValue(61, 33, toplamDirektCalisanSayisi.ToString());
            }
        }

        {
            List<MakinaParkı> makinaParkis = ExcelRead.readMakinaParki(makinaParkı);
            List<String> aracList = new List<String>();

            foreach (MakinaParkı item in makinaParkis)
            {
                if (!aracList.Contains(item.Araccinsi))
                {
                    aracList.Add(item.Araccinsi);
                }
            }

            int totalCount = 0;
            int totalCalisanCount = 0;
            int totalSaat = 0;


            for (int i = 0; i < aracList.Count; i++)
            {
                sheet.SetCellValue(34 + i, 34, aracList[i].ToString());
                int aracCount = 0;
                int aracCalisanCount = 0;
                int aracTotalSaat = 0;
                foreach (MakinaParkı item in makinaParkis)
                {
                    if (item.Araccinsi == aracList[i])
                    {
                        aracCount += 1;
                        if (item.Calısmadurumu == "Çalıştı")
                        {
                            aracCalisanCount += 1;
                            aracTotalSaat += int.Parse(item.Saat);
                        }
                    }
                }

                totalCount += aracCount;
                totalCalisanCount += aracCalisanCount;
                totalSaat += aracTotalSaat;

                sheet.SetCellValue(34 + i, 36, aracCount.ToString());
                sheet.SetCellValue(34 + i, 37, aracCalisanCount.ToString());
                sheet.SetCellValue(34 + i, 38, aracTotalSaat.ToString());
            }

            sheet.SetCellValue(55, 36, totalCount.ToString());
            sheet.SetCellValue(55, 37, totalCalisanCount.ToString());
            sheet.SetCellValue(55, 38, totalSaat.ToString());
        }


        for (int i = 0; i < spireWorkbbok.Worksheets.Count; i++)
        {
            //spireWorkbbok.Worksheets[i].PageSetup.IsFitToPage = true;
            spireWorkbbok.Worksheets[i].PageSetup.FitToPagesWide = 1;
            spireWorkbbok.Worksheets[i].PageSetup.Orientation = PageOrientationType.Landscape;
        }

        var FileDic = "files";
        Directory.CreateDirectory(FileDic);
        var tempGuid = Guid.NewGuid();

        var tempPdfFileName = $"{tempGuid}.xlsx";
        spireWorkbbok.SaveToFile($@"{FileDic}/{tempPdfFileName}", Spire.Xls.FileFormat.Version2016);

        byte[] fileBytes = System.IO.File.ReadAllBytes($"{FileDic}/{tempPdfFileName}");

        return File(fileBytes, "application/force-download", filename + ".xlsx");
    }
}