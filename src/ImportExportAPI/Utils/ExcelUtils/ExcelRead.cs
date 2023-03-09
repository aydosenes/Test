using ExcelDataReader;
using ImportExportAPI.Model.DataModel;

namespace ImportExportAPI.Utils.ExcelUtils;

    public static class ExcelRead
    {

        public static PersonelCard readPersonelCardData(IFormFile personelCardsFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            PersonelCard personelCard = new PersonelCard();
            using (var stream = new MemoryStream())
            {
                personelCardsFile.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;
                    while (reader.Read()) //Each row of the file
                    {
                        if (rowIndex == 0)
                        {
                            rowIndex = rowIndex + 1;
                            continue;
                        }
                        PersonelRow personelRow = new PersonelRow();

                        personelRow.OrderNo = reader.GetValue(0).ToString();
                        personelRow.RegistrationNo = reader.GetValue(1).ToString();
                        personelRow.FullName = reader.GetValue(2).ToString();
                        personelRow.Job = reader.GetValue(3).ToString();
                        personelRow.Role = reader.GetValue(4).ToString();
                        personelRow.IsGeneralWorker = reader.GetValue(5).ToString().ToLower() == "evet" ? true : false;
                        personelRow.ProjectCount = int.Parse(reader.GetValue(6).ToString());
                        String projectsStr = reader.GetValue(7) == null ? "" : reader.GetValue(7).ToString().Trim();
                        personelRow.Projects = projectsStr.Split(",").ToList();
                        personelCard.PersonelList.Add(personelRow);

                        rowIndex = rowIndex + 1;


                    }
                }
            }

            return personelCard;
        }

        public static ProjectCard readProjectCardData(IFormFile projectCardsFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            ProjectCard projectCard = new ProjectCard();
            using (var stream = new MemoryStream())
            {
                projectCardsFile.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;
                    while (reader.Read()) //Each row of the file
                    {
                        if (rowIndex == 0)
                        {
                            rowIndex = rowIndex + 1;
                            continue;
                        }
                        ProjectRow projectRow = new ProjectRow();

                        projectRow.Id = reader.GetValue(0) == null ? "" : reader.GetValue(0).ToString();
                        projectRow.Code = reader.GetValue(1) == null ? "" : reader.GetValue(1).ToString();
                        projectRow.SapCode = reader.GetValue(2) == null ? "" : reader.GetValue(2).ToString();
                        projectRow.Name = reader.GetValue(3) == null ? "" : reader.GetValue(3).ToString();
                        projectRow.ShortName = reader.GetValue(4) == null ? "" : reader.GetValue(4).ToString();
                        projectRow.ContractorName = reader.GetValue(5) == null ? "" : reader.GetValue(5).ToString();
                        projectRow.Type = reader.GetValue(6) == null ? "" : reader.GetValue(6).ToString();
                        projectRow.Directorate = reader.GetValue(7) == null ? "" : reader.GetValue(7).ToString();
                        projectRow.HeadShip = reader.GetValue(8) == null ? "" : reader.GetValue(8).ToString();
                        projectRow.Status = reader.GetValue(9) == null ? "" : reader.GetValue(9).ToString();

                        projectCard.ProjectList.Add(projectRow);
                        if (projectRow.Status == "Tamamlandı")
                        {
                            projectCard.DoneProjectList.Add(projectRow);
                        }
                        else
                        {
                            projectCard.CurrentProjectList.Add(projectRow);
                        }

                        rowIndex = rowIndex + 1;


                    }
                }
            }
            return projectCard;
        }

        public static List<MakinaParkı> readMakinaParki(IFormFile makinaParkiInputFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            List<MakinaParkı> makinaParkiList = new List<MakinaParkı>();
            using (var stream = new MemoryStream())
            {
                makinaParkiInputFile.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;

                    while (reader.Read()) //Each row of the file
                    {
                        if (rowIndex == 0)
                        {
                            rowIndex = rowIndex + 1;
                            continue;
                        }
                        MakinaParkı item = new MakinaParkı();
                        item.Firma = reader.GetValue(0) == null ? "" : reader.GetValue(0).ToString();
                        item.Araccinsi = reader.GetValue(1) == null ? "" : reader.GetValue(1).ToString();
                        item.Calısmadurumu = reader.GetValue(2) == null ? "" : reader.GetValue(2).ToString();
                        item.Saat = reader.GetValue(3) == null ? "" : reader.GetValue(3).ToString();


                        makinaParkiList.Add(item);
                        rowIndex = rowIndex + 1;


                    }
                }
            }
            return makinaParkiList;
        }

        public static List<PersonelPuantaj> readPersonelPuantaj(IFormFile personelPuantajInputFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            List<PersonelPuantaj> personelPuantajs = new List<PersonelPuantaj>();
            using (var stream = new MemoryStream())
            {
                personelPuantajInputFile.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;

                    while (reader.Read()) //Each row of the file
                    {
                        if (rowIndex == 0)
                        {
                            rowIndex = rowIndex + 1;
                            continue;
                        }
                        PersonelPuantaj item  = new PersonelPuantaj();
                        item.Firma = reader.GetValue(0) == null ? "" : reader.GetValue(0).ToString();
                        item.Tarih = reader.GetValue(1) == null ? "" : reader.GetValue(1).ToString();
                        item.AdSoyad = reader.GetValue(2) == null ? "" : reader.GetValue(2).ToString();
                        item.Görevi = reader.GetValue(3) == null ? "" : reader.GetValue(3).ToString();
                        item.CalısmaDurumu = reader.GetValue(4) == null ? "" : reader.GetValue(4).ToString();
                        item.KisiSayısı = reader.GetValue(5) == null ? "" : reader.GetValue(5).ToString();
                        item.Durumu = reader.GetValue(6) == null ? "" : reader.GetValue(6).ToString();
                        item.Vardiya = reader.GetValue(7) == null ? "" : reader.GetValue(7).ToString();




                        personelPuantajs.Add(item);
                       

                        rowIndex = rowIndex + 1;


                    }
                }
            }
            return personelPuantajs;
        }

        public static List<ImalatYapilan> readImalatYapilan(IFormFile imalatYapilanInputFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            List<ImalatYapilan>imalatYapilans = new List<ImalatYapilan>();
            using (var stream = new MemoryStream())
            {
                imalatYapilanInputFile.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;
                    
                    while (reader.Read()) //Each row of the file
                    {
                        if (rowIndex == 0)
                        {
                            rowIndex = rowIndex + 1;
                            continue;
                        }
                        ImalatYapilan item = new ImalatYapilan();
                        item.Firma = reader.GetValue(0) == null ? "" : reader.GetValue(0).ToString();
                        item.Tarih = reader.GetValue(1) == null ? "" : reader.GetValue(1).ToString();
                        item.Disiplin = reader.GetValue(2) == null ? "" : reader.GetValue(2).ToString();
                        item.Altgrupkodu = reader.GetValue(3) == null ? "" : reader.GetValue(3).ToString();
                        item.Aktivitekodu = reader.GetValue(4) == null ? "" : reader.GetValue(4).ToString();
                        item.Blokadı = reader.GetValue(5) == null ? "" : reader.GetValue(5).ToString();
                        item.Parsel = reader.GetValue(6) == null ? "" : reader.GetValue(6).ToString();
                        item.Katkod = reader.GetValue(7) == null ? "" : reader.GetValue(7).ToString();
                        item.AktiviteId = reader.GetValue(8) == null ? "" : reader.GetValue(8).ToString();
                        item.Aktiviteadı = reader.GetValue(9) == null ? "" : reader.GetValue(9).ToString();
                        item.Acıklama = reader.GetValue(10) == null ? "" : reader.GetValue(10).ToString();
                        item.Birim = reader.GetValue(11) == null ? "" : reader.GetValue(11).ToString();
                        item.Miktar = reader.GetValue(12) == null ? "" : reader.GetValue(12).ToString();
                        item.Direkpersonel = reader.GetValue(13) == null ? "" : reader.GetValue(13).ToString();
                        item.Vardiya = reader.GetValue(14) == null ? "" : reader.GetValue(14).ToString();

                        imalatYapilans.Add(item);


                        rowIndex = rowIndex + 1;


                    }
                }
            }
            return imalatYapilans;
        }

        public static List<ImalatPlanlanan> readImalatPlanlan(IFormFile imalatPlanlananInputFile)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            List<ImalatPlanlanan> imalatPlanlanans = new List<ImalatPlanlanan>();
            using (var stream = new MemoryStream())
            {
                imalatPlanlananInputFile.CopyTo(stream);
                stream.Position = 0;
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    int rowIndex = 0;

                    while (reader.Read()) //Each row of the file
                    {
                        if (rowIndex == 0)
                        {
                            rowIndex = rowIndex + 1;
                            continue;
                        }
                        ImalatPlanlanan item = new ImalatPlanlanan();
                        item.Firma = reader.GetValue(0) == null ? "" : reader.GetValue(0).ToString();
                        item.Tarih = reader.GetValue(1) == null ? "" : reader.GetValue(1).ToString();
                        item.Disiplin = reader.GetValue(2) == null ? "" : reader.GetValue(2).ToString();
                        item.Altgrupkodu = reader.GetValue(3) == null ? "" : reader.GetValue(3).ToString();
                        item.Aktivitekodu = reader.GetValue(4) == null ? "" : reader.GetValue(4).ToString();
                        item.Blokadı = reader.GetValue(5) == null ? "" : reader.GetValue(5).ToString();
                        item.Parsel = reader.GetValue(6) == null ? "" : reader.GetValue(6).ToString();
                        item.Katkod = reader.GetValue(7) == null ? "" : reader.GetValue(7).ToString();
                        item.AktiviteId = reader.GetValue(8) == null ? "" : reader.GetValue(8).ToString();
                        item.Aktiviteadı = reader.GetValue(9) == null ? "" : reader.GetValue(9).ToString();
                        item.Acıklama = reader.GetValue(10) == null ? "" : reader.GetValue(10).ToString();

                        imalatPlanlanans.Add(item);


                        rowIndex = rowIndex + 1;


                    }
                }
            }
            return imalatPlanlanans;
        }
    }
