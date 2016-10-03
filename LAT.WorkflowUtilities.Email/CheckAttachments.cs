using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Email
{
    public class CheckAttachments : CodeActivity
    {
        [RequiredArgument]
        [Input("Email To Check")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToCheck { get; set; }

        [OutputAttribute("Has Attachments")]
        public OutArgument<bool> HasAttachments { get; set; }

        [OutputAttribute("Attachment Count")]
        public OutArgument<int> AttachmentCount { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference emailToCheck = EmailToCheck.Get(executionContext);

                int count = GetAttachmentCount(service, emailToCheck.Id);

                AttachmentCount.Set(executionContext, count);
                HasAttachments.Set(executionContext, count > 0);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }

        private int GetAttachmentCount(IOrganizationService service, Guid emailId)
        {
            FetchExpression query = new FetchExpression(@"<fetch aggregate='true' >
                                                            <entity name='email' >
                                                            <attribute name='activityid' alias='count' aggregate='count' />
                                                            <filter>
                                                                <condition entityname='am' attribute='activityid' operator='eq' value='" + emailId + @"' />
                                                            </filter>
                                                            <link-entity name='activitymimeattachment' from='activityid' to='activityid' link-type='inner' alias='am' />
                                                            </entity>
                                                        </fetch>");

            EntityCollection results = service.RetrieveMultiple(query);

            return (int)results.Entities[0].GetAttributeValue<AliasedValue>("count").Value;
        }
    }
}
