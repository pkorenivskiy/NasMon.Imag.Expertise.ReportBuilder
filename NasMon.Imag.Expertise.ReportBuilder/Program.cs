using System;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using NasMon.Imag.Expertise.Reports;
using NasMon.Imag.Expertise.Reports.DataReaders;

namespace NasMon.Imag.Expertise.ReportBuilder
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            using (IHost host = CreateHostBuilder().Build())
            {
                var report = host.Services.GetRequiredService<IReport>();
                report.Generate();
            }
        }

        static IHostBuilder CreateHostBuilder()
        {
            return Host.CreateDefaultBuilder()
                .ConfigureServices((hostContext, services) =>
                {
                    services                        
                        .AddScoped<IExpertiseDataReader, ExpertiseDataReader>()
                        .AddScoped<IReport, ExpertiseReport>()
                        .AddLogging();                    
                });
        }
    }
}
