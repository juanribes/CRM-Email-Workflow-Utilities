using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Moq;
using System;
using System.Activities;
using System.Collections.Generic;
using Microsoft.Crm.Sdk.Messages;

namespace LAT.WorkflowUtilities.Email.Tests
{
	[TestClass]
	public class EmailQueueMembersTests
	{
		#region Class Constructor
		private readonly string _namespaceClassAssembly;
		public EmailQueueMembersTests()
		{
			//[Namespace.class name, assembly name] for the class/assembly being tested
			//Namespace and class name can be found on the class file being tested
			//Assembly name can be found under the project properties on the Application tab
			_namespaceClassAssembly = "LAT.WorkflowUtilities.Email.EmailQueueMembers" + ", " + "LAT.WorkflowUtilities.Email";
		}
		#endregion
		#region Test Initialization and Cleanup
		// Use ClassInitialize to run code before running the first test in the class
		[ClassInitialize()]
		public static void ClassInitialize(TestContext testContext) { }

		// Use ClassCleanup to run code after all tests in a class have run
		[ClassCleanup()]
		public static void ClassCleanup() { }

		// Use TestInitialize to run code before running each test 
		[TestInitialize()]
		public void TestMethodInitialize() { }

		// Use TestCleanup to run code after each test has run
		[TestCleanup()]
		public void TestMethodCleanup() { }
		#endregion

		[TestMethod]
		public void NoMembersWithoutOwnerOneExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 1;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, NoMembersWithoutOwnerOneExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> NoMembersWithoutOwnerOneExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			Entity activityParty = new Entity("activityparty") { Id = Guid.NewGuid() };
			activityParty["partyid"] = new EntityReference("contact", Guid.NewGuid());
			EntityCollection existingRecipients = new EntityCollection();
			existingRecipients.Entities.Add(activityParty);

			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email);

			EntityCollection queueMembers = new EntityCollection();

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void NoMembersWithoutOwnerNoExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 0;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, NoMembersWithoutOwnerNoExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> NoMembersWithoutOwnerNoExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email);

			EntityCollection queueMembers = new EntityCollection();

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithoutOwnerNoExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 1;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithoutOwnerNoExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithoutOwnerNoExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithoutOwnerTwoExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 3;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithoutOwnerTwoExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithoutOwnerTwoExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity activityParty1 = new Entity("activityparty") { Id = Guid.NewGuid() };
			activityParty1["partyid"] = new EntityReference("contact", Guid.NewGuid());	
			existingRecipients.Entities.Add(activityParty1);
			Entity activityParty2 = new Entity("activityparty") { Id = Guid.NewGuid() };
			activityParty2["partyid"] = new EntityReference("contact", Guid.NewGuid());
			existingRecipients.Entities.Add(activityParty2);

			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithoutOwnerOneExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 2;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithoutOwnerOneExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithoutOwnerOneExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity activityParty = new Entity("activityparty") { Id = Guid.NewGuid() };
			activityParty["partyid"] = new EntityReference("contact", Guid.NewGuid());
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;
			existingRecipients.Entities.Add(activityParty);

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void TwoMembersWithoutOwnerNoExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 2;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoMembersWithoutOwnerNoExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> TwoMembersWithoutOwnerNoExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);
			Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user2);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithOwnerDifferentNoExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", true},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 2;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithOwnerDifferentNoExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithOwnerDifferentNoExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			Entity queue = new Entity("queue");
			queue.Attributes["ownerid"] = new EntityReference("systemuser", Guid.NewGuid());

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email, queue);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithOwnerSameNoExisting()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", true},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 1;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithOwnerSameNoExistingSetup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithOwnerSameNoExistingSetup(Mock<IOrganizationService> serviceMock)
		{
			Guid sameId = Guid.NewGuid();

			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			Entity queue = new Entity("queue");
			queue.Attributes["ownerid"] = new EntityReference("systemuser", sameId);

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email, queue);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = sameId };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "8.1.0.512" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithoutOwnerNoExisting2011()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", false},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 1;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithoutOwnerNoExisting2011Setup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithoutOwnerNoExisting2011Setup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			Entity queue = new Entity("queue");
			queue.Attributes["ownerid"] = new EntityReference("systemuser", Guid.NewGuid());

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email, queue);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "5.0.9690.4376" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}

		[TestMethod]
		public void OneMemberWithOwnerNoExisting2011()
		{
			//Target
			Entity targetEntity = null;

			//Input parameters
			var inputs = new Dictionary<string, object>
			{
				{ "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
				{ "RecipientQueue", new EntityReference("queue", Guid.NewGuid()) },
				{ "IncludeOwner", true},
				{ "SendEmail", false }
			};

			//Expected value(s)
			const int expected = 1;

			//Invoke the workflow
			var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneMemberWithOwnerNoExisting2011Setup);

			//Test(s)
			Assert.AreEqual(expected, output["UsersAdded"]);
		}

		/// <summary>
		/// Modify to mock CRM Organization Service actions
		/// </summary>
		/// <param name="serviceMock">The Organization Service to mock</param>
		/// <returns>Configured Organization Service</returns>
		private static Mock<IOrganizationService> OneMemberWithOwnerNoExisting2011Setup(Mock<IOrganizationService> serviceMock)
		{
			EntityCollection existingRecipients = new EntityCollection();
			Entity email = new Entity("email") { Id = Guid.NewGuid() };
			email["to"] = existingRecipients;

			Entity queue = new Entity("queue");
			queue.Attributes["ownerid"] = new EntityReference("systemuser", Guid.NewGuid());

			serviceMock.Setup(t =>
				t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
				.ReturnsInOrder(email, queue);

			EntityCollection queueMembers = new EntityCollection();
			Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
			queueMembers.Entities.Add(user1);

			serviceMock.Setup(t =>
				t.RetrieveMultiple(It.IsAny<QueryExpression>()))
				.ReturnsInOrder(queueMembers);

			OrganizationResponse version = new OrganizationResponse();
			ParameterCollection results = new ParameterCollection { { "Version", "5.0.9690.4376" } };
			version.Results = results;

			serviceMock.Setup(t =>
				t.Execute(It.IsAny<OrganizationRequest>()))
				.ReturnsInOrder(version);

			return serviceMock;
		}	

		/// <summary>
		/// Invokes the workflow.
		/// </summary>
		/// <param name="name">Namespace.Class, Assembly</param>
		/// <param name="target">The target entity</param>
		/// <param name="inputs">The workflow input parameters</param>
		/// <param name="configuredServiceMock">The function to configure the Organization Service</param>
		/// <returns>The workflow output parameters</returns>
		private static IDictionary<string, object> InvokeWorkflow(string name, ref Entity target, Dictionary<string, object> inputs,
			Func<Mock<IOrganizationService>, Mock<IOrganizationService>> configuredServiceMock)
		{
			var testClass = Activator.CreateInstance(Type.GetType(name)) as CodeActivity;

			var serviceMock = new Mock<IOrganizationService>();
			var factoryMock = new Mock<IOrganizationServiceFactory>();
			var tracingServiceMock = new Mock<ITracingService>();
			var workflowContextMock = new Mock<IWorkflowContext>();

			//Apply configured Organization Service Mock
			if (configuredServiceMock != null)
				serviceMock = configuredServiceMock(serviceMock);

			IOrganizationService service = serviceMock.Object;

			//Mock workflow Context
			var workflowUserId = Guid.NewGuid();
			var workflowCorrelationId = Guid.NewGuid();
			var workflowInitiatingUserId = Guid.NewGuid();

			//Workflow Context Mock
			workflowContextMock.Setup(t => t.InitiatingUserId).Returns(workflowInitiatingUserId);
			workflowContextMock.Setup(t => t.CorrelationId).Returns(workflowCorrelationId);
			workflowContextMock.Setup(t => t.UserId).Returns(workflowUserId);
			var workflowContext = workflowContextMock.Object;

			//Organization Service Factory Mock
			factoryMock.Setup(t => t.CreateOrganizationService(It.IsAny<Guid>())).Returns(service);
			var factory = factoryMock.Object;

			//Tracing Service - Content written appears in output
			tracingServiceMock.Setup(t => t.Trace(It.IsAny<string>(), It.IsAny<object[]>())).Callback<string, object[]>(MoqExtensions.WriteTrace);
			var tracingService = tracingServiceMock.Object;

			//Parameter Collection
			ParameterCollection inputParameters = new ParameterCollection { { "Target", target } };
			workflowContextMock.Setup(t => t.InputParameters).Returns(inputParameters);

			//Workflow Invoker
			var invoker = new WorkflowInvoker(testClass);
			invoker.Extensions.Add(() => tracingService);
			invoker.Extensions.Add(() => workflowContext);
			invoker.Extensions.Add(() => factory);

			return invoker.Invoke(inputs);
		}
	}
}
