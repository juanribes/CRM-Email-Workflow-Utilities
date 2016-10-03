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
    public sealed class EmailTeam : CodeActivity
    {
        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("Recipient Team")]
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

                List<Entity> toList = new List<Entity>();

                Entity email = RetrieveEmail(service, emailToSend.Id);

                if (email == null)
                {
                    UsersAdded.Set(executionContext, 0);
                    return;
                }

                //Add any pre-defined recipients specified to the array               
                foreach (Entity activityParty in email.GetAttributeValue<EntityCollection>("to").Entities)
                {
                    toList.Add(activityParty);
                }

                EntityCollection teamMembers = GetTeamMembers(service, recipientTeam.Id);

                toList = ProcessUsers(service, teamMembers, toList);

                //Update the email
                email["to"] = toList.ToArray();
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

                UsersAdded.Set(executionContext, toList.Count);
            }
            catch (Exception ex)
            {
                tracer.Trace("Exception: {0}", ex.ToString());
            }
        }

        private Entity RetrieveEmail(IOrganizationService service, Guid emailId)
        {
            return service.Retrieve("email", emailId, new ColumnSet("to"));
        }

        private List<Entity> ProcessUsers(IOrganizationService service, EntityCollection teamMembers, List<Entity> toList)
        {
            foreach (Entity e in teamMembers.Entities)
            {
                Entity user = service.Retrieve("systemuser", e.GetAttributeValue<Guid>("systemuserid"),
                    new ColumnSet("internalemailaddress"));

                if (string.IsNullOrEmpty(user.GetAttributeValue<string>("internalemailaddress"))) continue;

                Entity activityParty = new Entity("activityparty");
                activityParty["partyid"] = new EntityReference("systemuser", e.GetAttributeValue<Guid>("systemuserid"));

                if (toList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.GetAttributeValue<Guid>("systemuserid"))) continue;

                toList.Add(activityParty);
            }

            return toList;
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