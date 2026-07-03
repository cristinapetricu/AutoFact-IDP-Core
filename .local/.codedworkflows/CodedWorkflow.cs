using System;
using System.Collections.Generic;
using System.Data;
using System.Net;
using UiPath.Activities.System.Jobs.Coded;
using UiPath.CodedWorkflows;
using UiPath.Core;
using UiPath.Core.Activities.Storage;
using UiPath.Excel;
using UiPath.Excel.Activities;
using UiPath.Excel.Activities.API;
using UiPath.Excel.Activities.API.Models;
using UiPath.GSuite.Activities.Api;
using UiPath.Mail.Activities.Api;
using UiPath.MicrosoftOffice365.Activities.Api;
using UiPath.OCR.Activities.Api;
using UiPath.OCR.Contracts.Activities;
using UiPath.OCR.Contracts.DataContracts;
using UiPath.Orchestrator.Client.Models;
using UiPath.UIAutomationNext.API.Contracts;
using UiPath.UIAutomationNext.API.Models;
using UiPath.UIAutomationNext.Enums;
using UiPath.Web.Activities.API;
using UiPath.Web.Activities.API.Models;
using UiPath.Web.Activities.Http.Models;
using UiPath.Word;
using UiPath.Word.Activities;
using UiPath.Word.Activities.API;
using UiPath.Word.Activities.API.Models;

namespace AutoFact_IDP_Core
{
    public partial class CodedWorkflow : CodedWorkflowBase
    {
        private Lazy<global::AutoFact_IDP_Core.WorkflowRunnerService> _workflowRunnerServiceLazy;
        private Lazy<ConnectionsManager> _connectionsManagerLazy;
        public CodedWorkflow()
        {
            _ = new System.Type[]
            {
                typeof(UiPath.Core.Activities.API.ISystemService),
                typeof(UiPath.Excel.Activities.API.IExcelService),
                typeof(UiPath.GSuite.Activities.Api.IGoogleConnectionsService),
                typeof(UiPath.Mail.Activities.Api.IMailService),
                typeof(UiPath.MicrosoftOffice365.Activities.Api.IOffice365ConnectionsService),
                typeof(UiPath.OCR.Activities.Api.IOcrService),
                typeof(UiPath.UIAutomationNext.API.Contracts.IUiAutomationAppService),
                typeof(UiPath.Web.Activities.API.ICurlImportService),
                typeof(UiPath.Web.Activities.API.IHttpService),
                typeof(UiPath.Web.Activities.API.ISoapService),
                typeof(UiPath.Word.Activities.API.IWordService)
            };
            _workflowRunnerServiceLazy = new Lazy<global::AutoFact_IDP_Core.WorkflowRunnerService>(() => new global::AutoFact_IDP_Core.WorkflowRunnerService(this.services));
#pragma warning disable
            _connectionsManagerLazy = new Lazy<ConnectionsManager>(() => new ConnectionsManager(serviceContainer));
#pragma warning restore
        }

        protected global::AutoFact_IDP_Core.WorkflowRunnerService workflows => _workflowRunnerServiceLazy.Value;
        protected ConnectionsManager connections => _connectionsManagerLazy.Value;
#pragma warning disable
        protected UiPath.Web.Activities.API.ICurlImportService curl { get => serviceContainer.Resolve<UiPath.Web.Activities.API.ICurlImportService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.Excel.Activities.API.IExcelService excel { get => serviceContainer.Resolve<UiPath.Excel.Activities.API.IExcelService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.GSuite.Activities.Api.IGoogleConnectionsService google { get => serviceContainer.Resolve<UiPath.GSuite.Activities.Api.IGoogleConnectionsService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.Web.Activities.API.IHttpService http { get => serviceContainer.Resolve<UiPath.Web.Activities.API.IHttpService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.Mail.Activities.Api.IMailService mail { get => serviceContainer.Resolve<UiPath.Mail.Activities.Api.IMailService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.OCR.Activities.Api.IOcrService ocr { get => serviceContainer.Resolve<UiPath.OCR.Activities.Api.IOcrService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.MicrosoftOffice365.Activities.Api.IOffice365ConnectionsService office365 { get => serviceContainer.Resolve<UiPath.MicrosoftOffice365.Activities.Api.IOffice365ConnectionsService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.Web.Activities.API.ISoapService soap { get => serviceContainer.Resolve<UiPath.Web.Activities.API.ISoapService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.Core.Activities.API.ISystemService system { get => serviceContainer.Resolve<UiPath.Core.Activities.API.ISystemService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.UIAutomationNext.API.Contracts.IUiAutomationAppService uiAutomation { get => serviceContainer.Resolve<UiPath.UIAutomationNext.API.Contracts.IUiAutomationAppService>() ; }
#pragma warning restore
#pragma warning disable
        protected UiPath.Word.Activities.API.IWordService word { get => serviceContainer.Resolve<UiPath.Word.Activities.API.IWordService>() ; }
#pragma warning restore
    }
}