using System.Collections.Generic;
using NasMon.Imag.Expertise.Models;

namespace NasMon.Imag.Expertise.Reports.DataReaders
{
    public interface IExpertiseDataReader
    {
        List<ExpertiseData> GetExpertiseData();
        List<ProjectInfo> GetProjectInfos();
    }
}