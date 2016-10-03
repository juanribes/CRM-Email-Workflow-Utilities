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
    public class EmailSecurityRole : CodeActivity
    {
        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("Security Role GUID")]
        [ReferenceTarget("role")]
        public InArgument<string> RecipientRole { get; set; }

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
                Guid roleId = IsGuid(RecipientRole.Get(executionContext));
                bool sendEmail = SendEmail.Get(executionContext);

                if (roleId == Guid.Empty)
                {
                    tracer.Trace("Invalid Role GUID");
                    throw new InvalidWorkflowException("Invalid Role GUID");
                }

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

                EntityCollection users = GetRoleUsers(service, roleId);

                toList = ProcessUsers(users, toList);

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

        private List<Entity> ProcessUsers(EntityCollection users, List<Entity> toList)
        {
            foreach (Entity e in users.Entities)
            {
                Entity activityParty = new Entity("activityparty");
                activityParty["partyid"] = new EntityReference("systemuser", e.Id);

                if (toList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

                toList.Add(activityParty);
            }

            return toList;
        }

        private EntityCollection GetRoleUsers(IOrganizationService service, Guid id)
        {
            //Query for the users with security role
            QueryExpression query = new QueryExpression
            {
                EntityName = "systemuser",
                ColumnSet = new ColumnSet("systemuserid"),
                LinkEntities = 
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "systemuser",
                        LinkFromAttributeName = "systemuserid",
                        LinkToEntityName = "systemuserroles",
                        LinkToAttributeName = "systemuserid",
                        LinkCriteria = new FilterExpression
                        {
                            FilterOperator = LogicalOperator.And,
                            Conditions = 
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "roleid",
                                    Operator = ConditionOperator.Equal,
                                    Values = { id }
                                }
                            }
                        }
                    }
                },
                Criteria = new FilterExpression
                {
                    Conditions =
                    {
                        new ConditionExpression
                        {
                            AttributeName = "internalemailaddress",
                            Operator = ConditionOperator.NotNull
                        }
                    }
                }
            };

            return service.RetrieveMultiple(query);
        }

        private Entity RetrieveEmail(IOrganizationService service, Guid emailId)
        {
            return service.Retrieve("email", emailId, new ColumnSet("to"));
        }

        private static Guid IsGuid(string value)
        {
            Guid parsed;
            return Guid.TryParse(value, out parsed) ? parsed : Guid.Empty;
        }
    }
}
