using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;

namespace LAT.WorkflowUtilities.Email
{
    public sealed class SendEmail : CodeActivity
    {
        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [OutputAttribute("Email Sent")]
        public OutArgument<bool> EmailSent { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference emailToSend = EmailToSend.Get(executionContext);

                SendEmailRequest request = new SendEmailRequest
                {
                    EmailId = emailToSend.Id,
                    TrackingToken = string.Empty,
                    IssueSend = true
                };

                SendEmailResponse response = (SendEmailResponse)service.Execute(request);

                bool emailSent = (response != null);

                EmailSent.Set(executionContext, emailSent);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }
    }
}