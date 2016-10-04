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
	public class CcQueueMembers : CodeActivity
	{
		[RequiredArgument]
		[Input("Email To Send")]
		[ReferenceTarget("email")]
		public InArgument<EntityReference> EmailToSend { get; set; }

		[RequiredArgument]
		[Input("Recipient Queue")]
		[ReferenceTarget("queue")]
		public InArgument<EntityReference> RecipientQueue { get; set; }

		[RequiredArgument]
		[Default("false")]
		[Input("Include Owner?")]
		public InArgument<bool> IncludeOwner { get; set; }

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
				EntityReference recipientQueue = RecipientQueue.Get(executionContext);
				bool sendEmail = SendEmail.Get(executionContext);
				bool includeOwner = IncludeOwner.Get(executionContext);

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

				bool is2011 = Is2011(service);
				EntityCollection users = is2011 ? GetQueueOwner(service, recipientQueue.Id) : GetQueueMembers(service, recipientQueue.Id);

				if (!is2011)
				{
					if (includeOwner)
						users.Entities.AddRange(GetQueueOwner(service, recipientQueue.Id).Entities);
				}

				ccList = ProcessUsers(users, ccList);

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
			catch (Exception e)
			{
				throw new InvalidPluginExecutionException(e.Message);
			}
		}

		private Entity RetrieveEmail(IOrganizationService service, Guid emailId)
		{
			return service.Retrieve("email", emailId, new ColumnSet("cc"));
		}

		private List<Entity> ProcessUsers(EntityCollection queueMembers, List<Entity> ccList)
		{
			foreach (Entity e in queueMembers.Entities)
			{
				Entity activityParty = new Entity("activityparty");
				activityParty["partyid"] = new EntityReference("systemuser", e.Id);

				if (ccList.Any(t => t.GetAttributeValue<EntityReference>("partyid").Id == e.Id)) continue;

				ccList.Add(activityParty);
			}

			return ccList;
		}

		private bool Is2011(IOrganizationService service)
		{
			//Check if 2011
			RetrieveVersionRequest request = new RetrieveVersionRequest();
			OrganizationResponse response = service.Execute(request);

			return response.Results["Version"].ToString().StartsWith("5");
		}

		private EntityCollection GetQueueOwner(IOrganizationService service, Guid queueId)
		{
			//Retrieve the queue owner
			Entity queue = service.Retrieve("queue", queueId, new ColumnSet("ownerid"));

			if (queue == null) return new EntityCollection();

			Entity owner = new Entity("systemuser") { Id = queue.GetAttributeValue<EntityReference>("ownerid").Id };

			EntityCollection ownerCollection = new EntityCollection();
			ownerCollection.Entities.Add(owner);

			return ownerCollection;
		}

		private EntityCollection GetQueueMembers(IOrganizationService service, Guid queueId)
		{
			//Query for the business unit members
			QueryExpression query = new QueryExpression
			{
				EntityName = "systemuser",
				ColumnSet = new ColumnSet(false),
				LinkEntities =
				{
					new LinkEntity
					{
						LinkFromEntityName = "systemuser",
						LinkFromAttributeName = "systemuserid",
						LinkToEntityName = "queuemembership",
						LinkToAttributeName = "systemuserid",
						Columns = new ColumnSet("systemuserid"),
						LinkCriteria = new FilterExpression
						{
							Conditions =
							{
								new ConditionExpression
								{
									AttributeName = "queueid",
									Operator = ConditionOperator.Equal,
									Values = { queueId }
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
