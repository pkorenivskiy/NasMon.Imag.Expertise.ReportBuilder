using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;
using NasMon.Imag.Expertise.Models;
using OfficeOpenXml;

namespace NasMon.Imag.Expertise.Reports.DataReaders
{
    public class ExpertiseDataReader : IExpertiseDataReader
    {
        private readonly string _excelDataFileName;
        private readonly string _excelProjectInfoFileName;
        private readonly ILogger<ExpertiseDataReader> _logger;

        public ExpertiseDataReader(IConfiguration configuration, ILogger<ExpertiseDataReader> logger)
        {
            _excelDataFileName = configuration["InputArgs:dataFileName"];
            _excelProjectInfoFileName = configuration["InputArgs:projectInfoFileName"];
            _logger = logger;
        }

        public List<ExpertiseData> GetExpertiseData()
        {
            Dictionary<string, ExpertiseData> data = new Dictionary<string, ExpertiseData>();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var xls = new ExcelPackage(new FileInfo(_excelDataFileName)))
            {
                var xlsSheet = xls.Workbook.Worksheets["Projects"];

                var row = 2;

                while (string.IsNullOrEmpty(xlsSheet.Cells[row, 1].Text) == false)
                {
                    ExpertiseData expertiseData;
                    var projectCode = xlsSheet.Cells[row, 3].Text.Trim();

                    if (data.ContainsKey(projectCode))
                    {
                        expertiseData = data[projectCode];
                    }
                    else
                    {
                        expertiseData = new ExpertiseData();
                        expertiseData.ProjectLead = xlsSheet.Cells[row, 2].Text.Trim();
                        expertiseData.ProjectCode = xlsSheet.Cells[row, 3].Text.Trim();
                        data.Add(projectCode, expertiseData);
                    }

                    expertiseData.ExpertPoints.Add(new ExpertPoint
                    {
                        ExpertName = xlsSheet.Cells[row, 5].Text.Trim(),
                        Points = Convert.ToDecimal(xlsSheet.Cells[row, 4].Value)
                    });

                    row++;
                }
            }

            return data.Values
                .ToList();
        }

        public List<ProjectInfo> GetProjectInfos()
        {
            var result = new List<ProjectInfo>();

            using (var xls = new ExcelPackage(new FileInfo(_excelProjectInfoFileName)))
            {
                var xlsSheet = xls.Workbook.Worksheets["Projects"];

                var row = 2;

                while (string.IsNullOrEmpty(xlsSheet.Cells[row, 1].Text) == false)
                {
                    var strValue = xlsSheet.Cells[row, 3].Text.Trim();

                    if (strValue.Length > 1)
                    {
                        result.Add(new ProjectInfo
                        {
                            Code = strValue.Substring(0, strValue.IndexOf(',')).Trim(),
                            Title = strValue.Substring(strValue.IndexOf(',') + 1, strValue.LastIndexOf(',') - strValue.IndexOf(',')).Trim()
                        });
                    }
                    else
                    {
                        _logger.LogError($"Unknown code value: [\"{xlsSheet.Cells[row, 3].Text.Trim()}\"] in Row {row}");
                    }

                    row++;
                }
            }

            return result;
        }
    }
}
