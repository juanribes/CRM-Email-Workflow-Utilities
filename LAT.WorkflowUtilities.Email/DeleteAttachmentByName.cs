using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Text;

namespace LAT.WorkflowUtilities.Email
{
    public class DeleteAttachmentByName : CodeActivity
    {
        [RequiredArgument]
        [Input("Email With Attachments To Remove")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailWithAttachments { get; set; }

        [RequiredArgument]
        [Input("File Name With Extension")]
        public InArgument<string> FileName { get; set; }

        [RequiredArgument]
        [Input("Add Delete Notice As Note?")]
        [Default("false")]
        public InArgument<bool> AppendNotice { get; set; }

        [OutputAttribute("Number Of Attachments Deleted")]
        public OutArgument<int> NumberOfAttachmentsDeleted { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference emailWithAttachments = EmailWithAttachments.Get(executionContext);
                string fileName = FileName.Get(executionContext);
                bool appendNotice = AppendNotice.Get(executionContext);

                EntityCollection attachments = GetAttachments(service, emailWithAttachments.Id);
                if (attachments.Entities.Count == 0) return;

                StringBuilder notice = new StringBuilder();
                int numberOfAttachmentsDeleted = 0;

                foreach (Entity activityMineAttachment in attachments.Entities)
                {
                    if (!String.Equals(activityMineAttachment.GetAttributeValue<string>("filename"), fileName, StringComparison.CurrentCultureIgnoreCase))
                        continue;

                    DeleteEmailAttachment(service, activityMineAttachment.Id);
                    numberOfAttachmentsDeleted++;

                    if (appendNotice)
                        notice.AppendLine("Deleted Attachment: " + activityMineAttachment.GetAttributeValue<string>("filename") + " " +
                                          DateTime.Now.ToShortDateString() + "\r\n");
                }

                if (appendNotice && notice.Length > 0)
                    UpdateEmail(service, emailWithAttachments.Id, notice.ToString());

                NumberOfAttachmentsDeleted.Set(executionContext, numberOfAttachmentsDeleted);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }

        private EntityCollection GetAttachments(IOrganizationService service, Guid emailId)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "activitymimeattachment",
                ColumnSet = new ColumnSet("filesize", "filename"),
                Criteria = new FilterExpression
                {
                    Conditions = 
                    {
                        new ConditionExpression
                        {
                            AttributeName = "objectid",
                            Operator = ConditionOperator.Equal,
                            Values = { emailId }
                        }
                    }
                }
            };

            return service.RetrieveMultiple(query);
        }

        private void DeleteEmailAttachment(IOrganizationService service, Guid activitymimeattachmentId)
        {
            service.Delete("activitymimeattachment", activitymimeattachmentId);
        }

        private void UpdateEmail(IOrganizationService service, Guid emailId, string notice)
        {
            Entity note = new Entity("annotation");
            note["notetext"] = notice;
            note["objectid"] = new EntityReference("email", emailId);

            service.Create(note);
        }
    }
}
