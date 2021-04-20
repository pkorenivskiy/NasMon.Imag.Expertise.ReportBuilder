using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Extensions.Logging;
using NasMon.Imag.Expertise.Models;
using NasMon.Imag.Expertise.Reports.DataReaders;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NasMon.Imag.Expertise.Reports
{
    public class ExpertiseReport : IReport
    {
        private readonly IExpertiseDataReader _expertiseDataReader;
        private readonly ILogger<ExpertiseReport> _logger;
        public ExpertiseReport(ILogger<ExpertiseReport> logger, IExpertiseDataReader expertiseDataReader)
        {
            _logger = logger;
            _expertiseDataReader = expertiseDataReader;
        }

        public void Generate()
        {
            _logger.LogInformation("Start report generate");

            _logger.LogInformation("Get Report Data");
            var data = _expertiseDataReader.GetExpertiseData();

            _logger.LogInformation("Calculate average");
            data.ForEach(x => x.AvgPoints = x.ExpertPoints.Average(p => p.Points));

            _logger.LogInformation("Get Projects infos");
            var infos = _expertiseDataReader.GetProjectInfos();
            foreach(var project in data)
            {
                var title = infos.FirstOrDefault(i => i.Code == project.ProjectCode)?.Title;
                if (title != null)
                    project.ProjectTitle = title;
                else
                    _logger.LogError($"title for ptoject [{project.ProjectCode}] not found");
            }
            data.ForEach(x => x.ProjectTitle = infos.FirstOrDefault(i => i.Code == x.ProjectCode)?.Title);

            if (System.IO.File.Exists(("result.xlsx")))
            {
                _logger.LogInformation("Delete old result file");
                System.IO.File.Delete("result.xlsx");
            }

            _logger.LogInformation("Generate result");
            CreateResult(data);
        }

        private void CreateResult(List<ExpertiseData> data)
        {
            using (var xls = new ExcelPackage(new System.IO.FileInfo("result.xlsx")))
            {
                var xlsSheet = xls.Workbook.Worksheets.Add("Result");
                var row = 1;
                var n = 1;
                foreach (var expertise in data.OrderByDescending(x => x.AvgPoints))
                {
                    xlsSheet.Cells[row, 1].Value = n++;
                    xlsSheet.Cells[row, 2].Value = expertise.ProjectLead;
                    xlsSheet.Cells[row, 3].Value = expertise.ProjectCode;
                    xlsSheet.Cells[row, 4].Value = expertise.ProjectTitle;
                    xlsSheet.Cells[row, 7].Value = expertise.AvgPoints;
                    xlsSheet.Cells[row, 7].Style.Numberformat.Format = "0.00";

                    foreach (var expert in expertise.ExpertPoints)
                    {
                        xlsSheet.Cells[row, 5].Value = expert.ExpertName;
                        xlsSheet.Cells[row++, 6].Value = expert.Points;
                    }

                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 1, row - 1, 1].Merge = true;
                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 1, row - 1, 1].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 2, row - 1, 2].Merge = true;
                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 2, row - 1, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 3, row - 1, 3].Merge = true;
                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 3, row - 1, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 4, row - 1, 4].Merge = true;
                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 4, row - 1, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 7, row - 1, 7].Merge = true;
                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 7, row - 1, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    xlsSheet.Cells[row - expertise.ExpertPoints.Count, 1, row - 1, 7].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                }

                xlsSheet.Column(3).Width = 12;

                xlsSheet.Column(2).Width = 35;
                xlsSheet.Cells[row, 2].Style.WrapText = true;

                xlsSheet.Column(4).Width = 35;
                xlsSheet.Cells[row, 4].Style.WrapText = true;                
                
                xlsSheet.Column(5).AutoFit();
                xlsSheet.Column(6).AutoFit();

                xls.Save();
            }
        }
    }
}
