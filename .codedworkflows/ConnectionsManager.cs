using UiPath.CodedWorkflows;
using System;

namespace AutoFact_IDP_Core
{
    public class ConnectionsManager
    {
        public GoogleDocsFactory GoogleDocs { get; set; }
        public DriveFactory Drive { get; set; }
        public GoogleFormsFactory GoogleForms { get; set; }
        public GmailFactory Gmail { get; set; }
        public GoogleSheetsFactory GoogleSheets { get; set; }
        public GoogleTasksFactory GoogleTasks { get; set; }
        public GoogleWorkspaceFactory GoogleWorkspace { get; set; }
        public ExcelFactory Excel { get; set; }
        public O365MailFactory O365Mail { get; set; }
        public Office365Factory Office365 { get; set; }
        public OneDriveFactory OneDrive { get; set; }

        public ConnectionsManager(ICodedWorkflowsServiceContainer resolver)
        {
            GoogleDocs = new GoogleDocsFactory(resolver);
            Drive = new DriveFactory(resolver);
            GoogleForms = new GoogleFormsFactory(resolver);
            Gmail = new GmailFactory(resolver);
            GoogleSheets = new GoogleSheetsFactory(resolver);
            GoogleTasks = new GoogleTasksFactory(resolver);
            GoogleWorkspace = new GoogleWorkspaceFactory(resolver);
            Excel = new ExcelFactory(resolver);
            O365Mail = new O365MailFactory(resolver);
            Office365 = new Office365Factory(resolver);
            OneDrive = new OneDriveFactory(resolver);
        }
    }
}