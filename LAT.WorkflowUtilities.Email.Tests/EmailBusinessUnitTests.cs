using Microsoft.VisualStudio.TestTools.UnitTesting;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using Moq;
using System;
using System.Activities;
using System.Collections.Generic;

namespace LAT.WorkflowUtilities.Email.Tests
{
    [TestClass]
    public class EmailBusinessUnitTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public EmailBusinessUnitTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.Email.EmailBusinessUnit" + ", " + "LAT.WorkflowUtilities.Email";
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
        public void NoUserBuNoExisting()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, NoUserBuNoExistingSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> NoUserBuNoExistingSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection existingRecipients = new EntityCollection();

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email);

            EntityCollection buMembers = new EntityCollection();

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers);

            return serviceMock;
        }

        [TestMethod]
        public void OneUserBuNoExisting()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneUserBuNoExistingSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> OneUserBuNoExistingSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection existingRecipients = new EntityCollection();

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email);

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test@test.com";

            EntityCollection buMembers = new EntityCollection();
            buMembers.Entities.Add(user1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers);

            return serviceMock;
        }

        [TestMethod]
        public void OneUserBuNoExistingNoEmail()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneUserBuNoExistingNoEmailSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> OneUserBuNoExistingNoEmailSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection existingRecipients = new EntityCollection();

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email);

            EntityCollection buMembers = new EntityCollection();

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers);

            return serviceMock;
        }

        [TestMethod]
        public void TwoUserBuNoExisting()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 2;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoUserBuNoExistingSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> TwoUserBuNoExistingSetup(Mock<IOrganizationService> serviceMock)
        {
            EntityCollection existingRecipients = new EntityCollection();

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email);

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test1@test.com";
            Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user2["internalemailaddress"] = "test2@test.com";

            EntityCollection buMembers = new EntityCollection();
            buMembers.Entities.Add(user1);
            buMembers.Entities.Add(user2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers);

            return serviceMock;
        }

        [TestMethod]
        public void TwoUserBuOneExisting()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", false},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 3;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoUserBuOneExistingSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> TwoUserBuOneExistingSetup(Mock<IOrganizationService> serviceMock)
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

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test1@test.com";
            Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user2["internalemailaddress"] = "test2@test.com";

            EntityCollection buMembers = new EntityCollection();
            buMembers.Entities.Add(user1);
            buMembers.Entities.Add(user2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers);

            return serviceMock;
        }

        [TestMethod]
        public void TwoUserBuOneExistingOneChildBu()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", true},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 4;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoUserBuOneExistingOneChildBuSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> TwoUserBuOneExistingOneChildBuSetup(Mock<IOrganizationService> serviceMock)
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

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test1@test.com";
            Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user2["internalemailaddress"] = "test2@test.com";

            EntityCollection buMembers = new EntityCollection();
            buMembers.Entities.Add(user1);
            buMembers.Entities.Add(user2);

            Entity childBu = new Entity("businessunit") { Id = Guid.NewGuid() };
            EntityCollection childBus = new EntityCollection();
            childBus.Entities.Add(childBu);

            Entity user3 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user3["internalemailaddress"] = "test3@test.com";

            EntityCollection childBuMembers = new EntityCollection();
            childBuMembers.Entities.Add(user3);

            EntityCollection childChildBus = new EntityCollection();

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers, childBus, childBuMembers, childChildBus);

            return serviceMock;
        }

        [TestMethod]
        public void TwoUserBuOneExistingTwoChildBu()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", true},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 5;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoUserBuOneExistingTwoChildBuSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> TwoUserBuOneExistingTwoChildBuSetup(Mock<IOrganizationService> serviceMock)
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

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test1@test.com";
            Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user2["internalemailaddress"] = "test2@test.com";

            EntityCollection buMembers = new EntityCollection();
            buMembers.Entities.Add(user1);
            buMembers.Entities.Add(user2);

            Entity childBu1 = new Entity("businessunit") { Id = Guid.NewGuid() };
            Entity childBu2 = new Entity("businessunit") { Id = Guid.NewGuid() };
            EntityCollection childBus = new EntityCollection();
            childBus.Entities.Add(childBu1);
            childBus.Entities.Add(childBu2);

            Entity user3 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user3["internalemailaddress"] = "test3@test.com";

            EntityCollection childBu1Members = new EntityCollection();
            childBu1Members.Entities.Add(user3);

            Entity user4 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user4["internalemailaddress"] = "test4@test.com";

            EntityCollection childBu2Members = new EntityCollection();
            childBu2Members.Entities.Add(user4);

            EntityCollection childChildBus = new EntityCollection();

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers, childBus, childBu1Members, childChildBus, childBu2Members, childChildBus);

            return serviceMock;
        }

        [TestMethod]
        public void TwoUserBuOneExistingChildChildBu()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientBusinessUnit", new EntityReference("team", Guid.NewGuid()) },
                { "IncludeChildren", true},
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 5;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoUserBuOneExistingChildChildBuSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> TwoUserBuOneExistingChildChildBuSetup(Mock<IOrganizationService> serviceMock)
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

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test1@test.com";
            Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user2["internalemailaddress"] = "test2@test.com";

            EntityCollection buMembers = new EntityCollection();
            buMembers.Entities.Add(user1);
            buMembers.Entities.Add(user2);

            Entity childBu = new Entity("businessunit") { Id = Guid.NewGuid() };
            EntityCollection childBus = new EntityCollection();
            childBus.Entities.Add(childBu);

            Entity user3 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user3["internalemailaddress"] = "test3@test.com";

            EntityCollection childBu1Members = new EntityCollection();
            childBu1Members.Entities.Add(user3);

            Entity childChildBu = new Entity("businessunit") { Id = Guid.NewGuid() };
            EntityCollection childChildBus = new EntityCollection();
            childChildBus.Entities.Add(childChildBu);

            Entity user4 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user4["internalemailaddress"] = "test4@test.com";

            EntityCollection childChildMembers = new EntityCollection();
            childChildMembers.Entities.Add(user4);

            EntityCollection emptyChildBus = new EntityCollection();

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(buMembers, childBus, childBu1Members, childChildBus, childChildMembers, emptyChildBus);

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
            var testClass = Activator.CreateInstance(Type.GetType(name)) as CodeActivity; ;

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
