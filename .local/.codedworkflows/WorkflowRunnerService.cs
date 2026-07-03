using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using UiPath.CodedWorkflows;
using UiPath.CodedWorkflows.Interfaces;
using UiPath.Activities.Contracts;
using AutoFact_IDP_Core;

[assembly: WorkflowRunnerServiceAttribute(typeof(AutoFact_IDP_Core.WorkflowRunnerService))]
namespace AutoFact_IDP_Core
{
    public class WorkflowRunnerService
    {
        private readonly ICodedWorkflowServices _services;
        public WorkflowRunnerService(ICodedWorkflowServices services)
        {
            _services = services;
        }

        /// <summary>
        /// Invokes the Framework/Process.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public (string out_SupplierName, string out_InvoiceDate, double out_AmountRON, System.Collections.Generic.Dictionary<string, System.Tuple<double, System.DateTime>> out_RatesCache) Framework_Process(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, string in_TransactionItem, System.Collections.Generic.Dictionary<string, string> in_Config, System.Collections.Generic.Dictionary<string, System.Tuple<double, System.DateTime>> in_RatesCache, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"Framework\Process.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_TransactionItem", in_TransactionItem }, { "in_Config", in_Config }, { "in_RatesCache", in_RatesCache } }, default, isolated, default, GetAssemblyName());
            return ((string)result["out_SupplierName"], (string)result["out_InvoiceDate"], (double)result["out_AmountRON"], (System.Collections.Generic.Dictionary<string, System.Tuple<double, System.DateTime>>)result["out_RatesCache"]);
        }

        /// <summary>
        /// Invokes the Framework/InitAllSettings.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public System.Collections.Generic.Dictionary<string, object> InitAllSettings(string in_ConfigFile, string[] in_ConfigSheets, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"Framework\InitAllSettings.xaml", new Dictionary<string, object> { { "in_ConfigFile", in_ConfigFile }, { "in_ConfigSheets", in_ConfigSheets } }, default, isolated, default, GetAssemblyName());
            return (System.Collections.Generic.Dictionary<string, object>)result["out_Config"];
        }

        /// <summary>
        /// Invokes the M1_ExtractInvoiceData.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public (string out_SupplierName, string out_InvoiceDate, double out_TotalAmount, string out_Currency, double out_Confidence, string out_Status) M1_ExtractInvoiceData(string in_FilePath, string in_DUApiKey, string in_DUEndpoint, string in_ConfidenceThreshold, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"M1_ExtractInvoiceData.xaml", new Dictionary<string, object> { { "in_FilePath", in_FilePath }, { "in_DUApiKey", in_DUApiKey }, { "in_DUEndpoint", in_DUEndpoint }, { "in_ConfidenceThreshold", in_ConfidenceThreshold } }, default, isolated, default, GetAssemblyName());
            return ((string)result["out_SupplierName"], (string)result["out_InvoiceDate"], (double)result["out_TotalAmount"], (string)result["out_Currency"], (double)result["out_Confidence"], (string)result["out_Status"]);
        }

        /// <summary>
        /// Invokes the M2_WriteToExcel.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public void M2_WriteToExcel(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, System.Data.DataTable in_DataTable, string in_OutputPath, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"M2_WriteToExcel.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_DataTable", in_DataTable }, { "in_OutputPath", in_OutputPath } }, default, isolated, default, GetAssemblyName());
        }

        /// <summary>
        /// Invokes the M3_CurrencyConversion.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public (double out_AmountRON, double out_ExchangeRate) M3_CurrencyConversion(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, double in_Amount, string in_FromCurrency, System.Collections.Generic.Dictionary<string, string> in_Config, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"M3_CurrencyConversion.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_Amount", in_Amount }, { "in_FromCurrency", in_FromCurrency }, { "in_Config", in_Config } }, default, isolated, default, GetAssemblyName());
            return ((double)result["out_AmountRON"], (double)result["out_ExchangeRate"]);
        }

        /// <summary>
        /// Invokes the M4_GenerateInvoice.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public (string out_PdfPath, string out_InvoiceNum, bool out_Success, string out_Status) M4_GenerateInvoice(System.Collections.Generic.Dictionary<string, string> in_Config, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"M4_GenerateInvoice.xaml", new Dictionary<string, object> { { "in_Config", in_Config } }, default, isolated, default, GetAssemblyName());
            return ((string)result["out_PdfPath"], (string)result["out_InvoiceNum"], (bool)result["out_Success"], (string)result["out_Status"]);
        }

        /// <summary>
        /// Invokes the M5_SendConfirmationEmail.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public bool M5_SendConfirmationEmail(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, System.Data.DataTable in_DataTable, System.Collections.Generic.List<string> in_ErrorList, string in_ExcelPath, System.Collections.Generic.Dictionary<string, string> in_Config, string in_GeneratedPdf, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"M5_SendConfirmationEmail.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_DataTable", in_DataTable }, { "in_ErrorList", in_ErrorList }, { "in_ExcelPath", in_ExcelPath }, { "in_Config", in_Config }, { "in_GeneratedPdf", in_GeneratedPdf } }, default, isolated, default, GetAssemblyName());
            return (bool)result["out_EmailSent"];
        }

        /// <summary>
        /// Invokes the Main.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public void Main(string in_OrchestratorQueueName, string in_OrchestratorQueueFolder, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"Main.xaml", new Dictionary<string, object> { { "in_OrchestratorQueueName", in_OrchestratorQueueName }, { "in_OrchestratorQueueFolder", in_OrchestratorQueueFolder } }, default, isolated, default, GetAssemblyName());
        }

        /// <summary>
        /// Invokes the ParseAmount.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public double ParseAmount(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, string in_RawAmount, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"ParseAmount.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_RawAmount", in_RawAmount } }, default, isolated, default, GetAssemblyName());
            return (double)result["out_Amount"];
        }

        /// <summary>
        /// Invokes the ParseAndFormatDate.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public string ParseAndFormatDate(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, string in_RawDate, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"ParseAndFormatDate.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_RawDate", in_RawDate } }, default, isolated, default, GetAssemblyName());
            return (string)result["out_FormattedDate"];
        }

        /// <summary>
        /// Invokes the Process.xaml
        /// </summary>
        /// <param name="isolated">Indicates whether to isolate executions (run them within a different process)</param>
        public (string out_SupplierName, string out_InvoiceDate, double out_AmountRON, System.Collections.Generic.Dictionary<string, System.Tuple<double, System.DateTime>> out_RatesCache) Process(string UiPathEventConnector, string UiPathEvent, string UiPathEventObjectType, string UiPathEventObjectId, string UiPathAdditionalEventData, string in_TransactionItem, System.Collections.Generic.Dictionary<string, string> in_Config, System.Collections.Generic.Dictionary<string, System.Tuple<double, System.DateTime>> in_RatesCache, System.Boolean isolated = false)
        {
            var result = _services.WorkflowInvocationService.RunWorkflow(@"Process.xaml", new Dictionary<string, object> { { "UiPathEventConnector", UiPathEventConnector }, { "UiPathEvent", UiPathEvent }, { "UiPathEventObjectType", UiPathEventObjectType }, { "UiPathEventObjectId", UiPathEventObjectId }, { "UiPathAdditionalEventData", UiPathAdditionalEventData }, { "in_TransactionItem", in_TransactionItem }, { "in_Config", in_Config }, { "in_RatesCache", in_RatesCache } }, default, isolated, default, GetAssemblyName());
            return ((string)result["out_SupplierName"], (string)result["out_InvoiceDate"], (double)result["out_AmountRON"], (System.Collections.Generic.Dictionary<string, System.Tuple<double, System.DateTime>>)result["out_RatesCache"]);
        }

        private string GetAssemblyName()
        {
            var assemblyProvider = _services.Container.Resolve<ILibraryAssemblyProvider>();
            return assemblyProvider.GetLibraryAssemblyName(GetType().Assembly);
        }
    }
}