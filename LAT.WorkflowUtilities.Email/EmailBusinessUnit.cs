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
    public sealed class EmailBusinessUnit : CodeActivity
    {
        [RequiredArgument]
        [Input("Email To Send")]
        [ReferenceTarget("email")]
        public InArgument<EntityReference> EmailToSend { get; set; }

        [RequiredArgument]
        [Input("Recipient Business Unit")]
        [ReferenceTarget("businessunit")]
        public InArgument<EntityReference> RecipientBusinessUnit { get; set; }

        [RequiredArgument]
        [Default("false")]
        [Input("Include All Child Business Units?")]
        public InArgument<bool> IncludeChildren { get; set; }

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
                EntityReference recipientBusinessUnit = RecipientBusinessUnit.Get(executionContext);
                bool includeChildren = IncludeChildren.Get(executionContext);
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

                EntityCollection buUsers = GetBuUsers(service, recipientBusinessUnit.Id);

                toList = ProcessUsers(buUsers, toList);

                if (includeChildren)
                    toList = DrillDownBu(service, recipientBusinessUnit.Id, toList);

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

        private List<Entity> ProcessUsers(EntityCollection teamMembers, List<Entity> toList)
        {
            foreach (Entity e in teamMembers.Entities)
            {
                Entity activityParty = new Entity("activityparty");
                activityParty["partyid"] = new EntityReference("systemuser", e.Id);

                if (toList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

                toList.Add(activityParty);
            }

            return toList;
        }

        private List<Entity> DrillDownBu(IOrganizationService service, Guid businessUnitId, List<Entity> toList)
        {
            //Find and process child business units
            EntityCollection childBu = GetChildBu(service, businessUnitId);

            foreach (Entity businessUnit in childBu.Entities)
            {
                EntityCollection childUsers = GetBuUsers(service, businessUnit.Id);
                toList = ProcessUsers(childUsers, toList);
                toList = DrillDownBu(service, businessUnit.Id, toList);
            }

            return toList;
        }

        private EntityCollection GetChildBu(IOrganizationService service, Guid businessUnitId)
        {
            //Query for the child business units
            QueryExpression query = new QueryExpression
            {
                EntityName = "businessunit",
                ColumnSet = new ColumnSet("businessunitid"),
                LinkEntities = 
                    {
                        new LinkEntity
                        {
                            LinkFromEntityName = "businessunit",
                            LinkFromAttributeName = "businessunitid",
                            LinkToEntityName = "businessunit",
                            LinkToAttributeName = "businessunitid",
                            LinkCriteria = new FilterExpression
                            {
                                FilterOperator = LogicalOperator.And,
                                Conditions = 
                                {
                                    new ConditionExpression
                                    {
                                        AttributeName = "parentbusinessunitid",
                                        Operator = ConditionOperator.Equal,
                                        Values = { businessUnitId }
                                    }
                                }
                            }
                        }
                    }
            };

            return service.RetrieveMultiple(query);
        }

        private EntityCollection GetBuUsers(IOrganizationService service, Guid businessUnitId)
        {
            //Query for the business unit members
            QueryExpression query = new QueryExpression
            {
                EntityName = "systemuser",
                ColumnSet = new ColumnSet("systemuserid"),
                LinkEntities = 
                {
                    new LinkEntity
                    {
                        LinkFromEntityName = "systemuser",
                        LinkFromAttributeName = "businessunitid",
                        LinkToEntityName = "businessunit",
                        LinkToAttributeName = "businessunitid",
                        LinkCriteria = new FilterExpression
                        {
                            FilterOperator = LogicalOperator.And,
                            Conditions = 
                            {
                                new ConditionExpression
                                {
                                    AttributeName = "businessunitid",
                                    Operator = ConditionOperator.Equal,
                                    Values = { businessUnitId }
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
                        },
                        new ConditionExpression
                        {
                            AttributeName = "accessmode",
                            Operator = ConditionOperator.NotIn,
                            Values = { 3, 4 }
                        }
                    }
                }
            };

            return service.RetrieveMultiple(query);
        }
    }
}