using UiPath.CodedWorkflows;
using System;

namespace AutoFact_IDP_Core
{
    public class GoogleDocsFactory
    {
        public GoogleDocsFactory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class DriveFactory
    {
        public DriveFactory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class GoogleFormsFactory
    {
        public GoogleFormsFactory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class GmailFactory
    {
        public UiPath.GSuite.Activities.Api.GmailConnection My_Workspace_cristinapetricu_gmail_com { get; set; }

        public GmailFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_cristinapetricu_gmail_com = new UiPath.GSuite.Activities.Api.GmailConnection("cfca04c6-595f-4dd1-93b4-c4dd37ebf671", resolver);
        }
    }

    public class GoogleSheetsFactory
    {
        public GoogleSheetsFactory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class GoogleTasksFactory
    {
        public GoogleTasksFactory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class GoogleWorkspaceFactory
    {
        public GoogleWorkspaceFactory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class ExcelFactory
    {
        public UiPath.MicrosoftOffice365.Activities.Api.ExcelConnection My_Workspace_Microsoft_OneDrive___SharePoint { get; set; }

        public ExcelFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_Microsoft_OneDrive___SharePoint = new UiPath.MicrosoftOffice365.Activities.Api.ExcelConnection("399cf810-e7b9-4f63-9f99-6967c775cd13", resolver);
        }
    }

    public class O365MailFactory
    {
        public UiPath.MicrosoftOffice365.Activities.Api.MailConnection My_Workspace_Microsoft_Outlook_365 { get; set; }
        public UiPath.MicrosoftOffice365.Activities.Api.MailConnection My_Workspace_cristinapetricu_gmail_com { get; set; }
        public UiPath.MicrosoftOffice365.Activities.Api.MailConnection My_Workspace_Microsoft_Outlook_365__3 { get; set; }
        public UiPath.MicrosoftOffice365.Activities.Api.MailConnection My_Workspace_Microsoft_Outlook_365__4 { get; set; }
        public UiPath.MicrosoftOffice365.Activities.Api.MailConnection My_Workspace_Microsoft_Outlook_365__5 { get; set; }

        public O365MailFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_Microsoft_Outlook_365 = new UiPath.MicrosoftOffice365.Activities.Api.MailConnection("cc23fcdb-d862-4f34-a4dc-dd589cd8240d", resolver);
            My_Workspace_cristinapetricu_gmail_com = new UiPath.MicrosoftOffice365.Activities.Api.MailConnection("056b81dd-b785-4ba0-a067-e64456bed77b", resolver);
            My_Workspace_Microsoft_Outlook_365__3 = new UiPath.MicrosoftOffice365.Activities.Api.MailConnection("703b7993-a819-454f-8354-a85880f343b3", resolver);
            My_Workspace_Microsoft_Outlook_365__4 = new UiPath.MicrosoftOffice365.Activities.Api.MailConnection("a80f4636-4cc0-4cdd-9604-1e66bd3f1595", resolver);
            My_Workspace_Microsoft_Outlook_365__5 = new UiPath.MicrosoftOffice365.Activities.Api.MailConnection("bfeeb3b0-554c-4ce2-8396-675981236389", resolver);
        }
    }

    public class Office365Factory
    {
        public Office365Factory(ICodedWorkflowsServiceContainer resolver)
        {
        }
    }

    public class OneDriveFactory
    {
        public UiPath.MicrosoftOffice365.Activities.Api.OneDriveConnection My_Workspace_Microsoft_OneDrive___SharePoint { get; set; }

        public OneDriveFactory(ICodedWorkflowsServiceContainer resolver)
        {
            My_Workspace_Microsoft_OneDrive___SharePoint = new UiPath.MicrosoftOffice365.Activities.Api.OneDriveConnection("399cf810-e7b9-4f63-9f99-6967c775cd13", resolver);
        }
    }
}