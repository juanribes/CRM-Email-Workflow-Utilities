using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Text;

namespace LAT.WorkflowUtilities.Email
{
    public sealed class DeleteAttachment : CodeActivity
    {
        [RequiredArgument]
        [Input("Email With Attachments To Remove")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailWithAttachments { get; set; }

        [Input("Delete <= Than X Bytes (Empty = 0)")]
        public InArgument<int> DeleteSizeMax { get; set; }

        [Input("Delete >= Than X Bytes (Empty = 2,147,483,647)")]
        public InArgument<int> DeleteSizeMin { get; set; }

        [Input("Limit To Extensions (Comma Delimited, Empty = Ignore)")]
        public InArgument<string> Extensions { get; set; }

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
                int deleteSizeMax = DeleteSizeMax.Get(executionContext);
                int deleteSizeMin = DeleteSizeMin.Get(executionContext);
                string extensions = Extensions.Get(executionContext);
                bool appendNotice = AppendNotice.Get(executionContext);

                if (deleteSizeMax == 0) deleteSizeMax = int.MaxValue;
                if (deleteSizeMin > deleteSizeMax)
                {
                    tracer.Trace("Exception: {0}", "Min:" + deleteSizeMin + " Max:" + deleteSizeMax);
                    throw new InvalidPluginExecutionException("Minimum Size Cannot Be Greater Than Maximum Size");
                }

                EntityCollection attachments = GetAttachments(service, emailWithAttachments.Id);
                if (attachments.Entities.Count == 0) return;

                string[] filetypes = new string[0];
                if (!string.IsNullOrEmpty(extensions))
                    filetypes = extensions.Replace(".", string.Empty).Split(',');

                StringBuilder notice = new StringBuilder();
                int numberOfAttachmentsDeleted = 0;

                bool delete = false;
                foreach (Entity activityMineAttachment in attachments.Entities)
                {
                    delete = false;

                    if (activityMineAttachment.GetAttributeValue<int>("filesize") >= deleteSizeMax)
                        delete = true;

                    if (activityMineAttachment.GetAttributeValue<int>("filesize") <= deleteSizeMin)
                        delete = true;

                    if (filetypes.Length > 0 && delete)
                        delete = ExtensionMatch(filetypes, activityMineAttachment.GetAttributeValue<string>("filename"));

                    if (!delete) continue;

                    DeleteEmailAttachment(service, activityMineAttachment.Id);
                    numberOfAttachmentsDeleted++;

                    if (appendNotice)
                        notice.AppendLine("Deleted Attachment: " + activityMineAttachment.GetAttributeValue<string>("filename") + " " +
                                          DateTime.Now.ToShortDateString() + "\r\n");
                }

                if (delete && appendNotice && notice.Length > 0)
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

        private static bool ExtensionMatch(IEnumerable<string> extenstons, string filename)
        {
            foreach (string ex in extenstons)
            {
                if (filename.EndsWith("." + ex))
                    return true;
            }
            return false;
        }
    }
}