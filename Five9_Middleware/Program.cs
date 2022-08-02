using System;
using System.ServiceModel;
using System.Threading.Tasks;
using Five9_Middleware.Helpers;
using Five9ConfigService;
using System.IO;
using System.Text;
using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Linq;

namespace Five9_Middleware
{
    class Program
    {
        static void Main(string[] args)
        {

            string _username = "jesuscueva100@gmail.com";
            string _password = "JCmanitoCueva1109!";

            var binding = new BasicHttpBinding
            {

                MaxBufferSize = int.MaxValue,
                MaxBufferPoolSize = int.MaxValue,
                MaxReceivedMessageSize = int.MaxValue,
                ReaderQuotas =
                {

                    MaxArrayLength = int.MaxValue,
                    MaxBytesPerRead = int.MaxValue,
                    MaxNameTableCharCount = int.MaxValue,
                    MaxStringContentLength = int.MaxValue

                },
                Security =
                {

                    Mode = BasicHttpSecurityMode.Transport,
                    Transport = {ClientCredentialType = HttpClientCredentialType.None}

                }
            };

            var adminUrl = $"https://api.five9.com/wsadmin/v12/AdminWebService";
            var address = new EndpointAddress(adminUrl);
            var adminClient = new WsAdminClient(binding, address);
            var inserter = new AuthHeaderInserter
            {

                Username = _username,
                Password = _password

            };

            adminClient.Endpoint.EndpointBehaviors.Add(new AuthHeaderBehavior(inserter));

            Console.WriteLine("Iniciando Reporte Call Logs...");
            string Call_Logs = generateCallLogs_Report(adminClient, "Call Log Reports", "Call Log");
            Console.WriteLine("Reporte Call Logs ha finalizado...");

            Console.WriteLine("Iniciando Reporte Agent State Details...");
            string Agent_State = generateAgentStateDetails(adminClient, "Agent Reports", "Agent State Details");
            Console.WriteLine("Reporte Agent State Details ha finalizado...");

            Console.WriteLine("Iniciando Reporte Agent State Summary by State...");
            string AgentSummary_State = generateAgentStateDetails(adminClient, "Agent Reports", "Agent State Summary by State");
            Console.WriteLine("Reporte Agent State Summary by State ha finalizado...");

            //Save report to file
            StringBuilder csvcontent = new StringBuilder();
            csvcontent.AppendLine(Call_Logs);
            csvcontent.AppendLine(Agent_State);
            csvcontent.AppendLine(AgentSummary_State);
            string csvpath = "C:/Users/jesus/Documents/ReportData.csv";
            Console.WriteLine("Se ha generado el archivo CSV exitosamente!");
            File.AppendAllText(csvpath, csvcontent.ToString());

            // Convert CSV to XLSX
            CSV2XLS(csvpath);
            Console.WriteLine("Se ha generado el archivo XLSX exitosamente!");

            Console.WriteLine("Finalizando ejecución...");
            Console.ReadLine();
        }

        static string generateCallLogs_Report(WsAdminClient adminClient, string folderName,string reportName)
        {

            customReportCriteria reqRunReport = new customReportCriteria();
            reqRunReport.time = new reportTimeCriteria();
            reqRunReport.time.start = DateTime.Parse("2022-07-01 07:01:00+00:00");
            reqRunReport.time.startSpecified = true;
            reqRunReport.time.end = DateTime.Parse("2022-07-01 18:59:00+00:00");
            reqRunReport.time.endSpecified = true;

            Task<runReportResponse> response = adminClient.runReportAsync(folderName, reportName, reqRunReport);
            string reportId = response.Result.@return;

            var isRunning = true;

            while (isRunning)
            {


                Task<isReportRunningResponse> respIsReportRunning = adminClient.isReportRunningAsync(reportId, 5);
                isRunning = respIsReportRunning.Result.@return;

            }

            _ = new getReportResultCsv
            {
                identifier = reportId
            };

            Task<getReportResultCsvResponse> getReportResultCsv = adminClient.getReportResultCsvAsync(reportId);
            string reportData = getReportResultCsv.Result.@return;

            return reportData;

        }

        static string generateAgentStateDetails(WsAdminClient adminClient, string folderName, string reportName)
        {

            customReportCriteria runReport_2 = new customReportCriteria
            {
                time = new reportTimeCriteria()
            };
            runReport_2.time.start = DateTime.Parse("2022-07-01 07:01:00+00:00");
            runReport_2.time.startSpecified = true;
            runReport_2.time.end = DateTime.Parse("2022-07-01 18:59:00+00:00");
            runReport_2.time.endSpecified = true;

            Task<runReportResponse> response_2 = adminClient.runReportAsync(folderName, reportName, runReport_2);
            string reportId_2 = response_2.Result.@return;

            var isRunning_2 = true;

            while (isRunning_2)
            {


                Task<isReportRunningResponse> respIsReportRunning_2 = adminClient.isReportRunningAsync(reportId_2, 5);
                isRunning_2 = respIsReportRunning_2.Result.@return;

            }

            _ = new getReportResultCsv
            {
                identifier = reportId_2
            };

            Task<getReportResultCsvResponse> getReportResultCsv_2 = adminClient.getReportResultCsvAsync(reportId_2);
            string reportData_2 = getReportResultCsv_2.Result.@return;

            return reportData_2;

        }

        static string generateAgentSummaryByState(WsAdminClient adminClient, string folderName, string reportName)
        {

            customReportCriteria runReport_2 = new customReportCriteria
            {
                time = new reportTimeCriteria()
            };
            runReport_2.time.start = DateTime.Parse("2022-07-01 07:01:00+00:00");
            runReport_2.time.startSpecified = true;
            runReport_2.time.end = DateTime.Parse("2022-07-01 18:59:00+00:00");
            runReport_2.time.endSpecified = true;

            Task<runReportResponse> response_2 = adminClient.runReportAsync(folderName, reportName, runReport_2);
            string reportId_2 = response_2.Result.@return;

            var isRunning_2 = true;

            while (isRunning_2)
            {


                Task<isReportRunningResponse> respIsReportRunning_2 = adminClient.isReportRunningAsync(reportId_2, 5);
                isRunning_2 = respIsReportRunning_2.Result.@return;

            }

            _ = new getReportResultCsv
            {
                identifier = reportId_2
            };

            Task<getReportResultCsvResponse> getReportResultCsv_2 = adminClient.getReportResultCsvAsync(reportId_2);
            string reportData_2 = getReportResultCsv_2.Result.@return;

            return reportData_2;

        }

        static void CSV2XLS(string csv)
        {

            string xls = "Output_CSV.xlsx";

            Excel.Application xl = new Excel.Application();
            
            // Open Excel Workbook for conversion
            Excel.Workbook wb = xl.Workbooks.Open(csv);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets.get_Item(1);

            // Select the UsedRange
            Excel.Range used = ws.UsedRange;

            // Autofit the columns
            used.EntireColumn.AutoFit();

            // Save
            wb.SaveAs(xls, 51);

            // Close the workbook
            wb.Close();

            // Quit Excel
            xl.Quit();

        }

    }
}
