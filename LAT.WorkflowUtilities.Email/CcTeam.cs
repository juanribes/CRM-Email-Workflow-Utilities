using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System;
using System.Activities;
using System.Collections.Generic;
using System.Linq;

namespace LAT.WorkflowUtilities.Email
{
    public sealed class CcTeam : CodeActivity
    {
        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("CC Team")]
        [ReferenceTarget("team")]
        public InArgument<EntityReference> RecipientTeam { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Send Email?")]
        public InArgument<bool> SendEmail { get; set; }

        [OutputAttribute("Users Added")]
        public OutArgument<int> UsersAdded { get; set; }

        protected override void Execute(CodeActivityContext executionContext)
        {
            ITracingService tracer = executionContext.GetExtension<ITracingService>();
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                EntityReference emailToSend = EmailToSend.Get(executionContext);
                EntityReference recipientTeam = RecipientTeam.Get(executionContext);
                bool sendEmail = SendEmail.Get(executionContext);

                List<Entity> ccList = new List<Entity>();

                Entity email = RetrieveEmail(service, emailToSend.Id);

                if (email == null)
                {
                    UsersAdded.Set(executionContext, 0);
                    return;
                }

                //Add any pre-defined recipients specified to the array               
                foreach (Entity activityParty in email.GetAttributeValue<EntityCollection>("cc").Entities)
                {
                    ccList.Add(activityParty);
                }

                EntityCollection teamMembers = GetTeamMembers(service, recipientTeam.Id);

                ccList = ProcessUsers(service, teamMembers, ccList);

                //Update the email
                email["cc"] = ccList.ToArray();
                service.Update(email);

                //Send
                if (sendEmail)
                {
                    SendEmailRequest request = new SendEmailRequest
                    {
                        EmailId = emailToSend.Id,
                        TrackingToken = string.Empty,
                        IssueSend = true
                    };

                    service.Execute(request);
                }

                UsersAdded.Set(executionContext, ccList.Count);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }

        private Entity RetrieveEmail(IOrganizationService service, Guid emailId)
        {
            return service.Retrieve("email", emailId, new ColumnSet("cc"));
        }

        private List<Entity> ProcessUsers(IOrganizationService service, EntityCollection teamMembers, List<Entity> ccList)
        {
            foreach (Entity e in teamMembers.Entities)
            {
                Entity user = service.Retrieve("systemuser", e.GetAttributeValue<Guid>("systemuserid"),
                    new ColumnSet("internalemailaddress"));

                if (string.IsNullOrEmpty(user.GetAttributeValue<string>("internalemailaddress"))) continue;

                Entity activityParty = new Entity("activityparty");
                activityParty["partyid"] = new EntityReference("systemuser", e.GetAttributeValue<Guid>("systemuserid"));

                if (ccList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.GetAttributeValue<Guid>("systemuserid"))) continue;

                ccList.Add(activityParty);
            }

            return ccList;
        }

        private EntityCollection GetTeamMembers(IOrganizationService service, Guid teamId)
        {
            QueryExpression query = new QueryExpression
            {
                EntityName = "teammembership",
                ColumnSet = new ColumnSet(true),
                LinkEntities = 
                    {
                        new LinkEntity
                        {
                            LinkFromEntityName = "teammembership",
                            LinkFromAttributeName = "teamid",
                            LinkToEntityName = "team",
                            LinkToAttributeName = "teamid",
                            LinkCriteria = new FilterExpression
                            {
                                FilterOperator = LogicalOperator.And,
                                Conditions = 
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "teamid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { teamId }
                                    }
                                }
                            }
                        }
                    }
            };

            return service.RetrieveMultiple(query);
        }
    }
}