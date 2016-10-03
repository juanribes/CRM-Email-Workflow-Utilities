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
    public class EmailTeamTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public EmailTeamTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.Email.EmailTeam" + ", " + "LAT.WorkflowUtilities.Email";
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
        public void OneUserTeamWithOneExisting()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientTeam", new EntityReference("team", Guid.NewGuid()) },
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 2;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneUserTeamSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> OneUserTeamSetup(Mock<IOrganizationService> serviceMock)
        {
            Entity activityParty = new Entity("activityparty") { Id = Guid.NewGuid() };
            activityParty["partyid"] = new EntityReference("contact", Guid.NewGuid());
            EntityCollection existingRecipients = new EntityCollection();
            existingRecipients.Entities.Add(activityParty);

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test@test.com";

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email, user1);

            Entity teamMember1 = new Entity("teammembership") { Id = Guid.NewGuid() };
            teamMember1["systemuserid"] = Guid.NewGuid();

            EntityCollection teamMembers = new EntityCollection();
            teamMembers.Entities.Add(teamMember1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(teamMembers);

            return serviceMock;
        }

        [TestMethod]
        public void OneUserTeamWithOneExistingNoEmail()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientTeam", new EntityReference("team", Guid.NewGuid()) },
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, OneUserTeamWithOneExistingNoEmailSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> OneUserTeamWithOneExistingNoEmailSetup(Mock<IOrganizationService> serviceMock)
        {
            Entity activityParty = new Entity("activityparty") { Id = Guid.NewGuid() };
            activityParty["partyid"] = new EntityReference("contact", Guid.NewGuid());
            EntityCollection existingRecipients = new EntityCollection();
            existingRecipients.Entities.Add(activityParty);

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = null;

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email, user1);

            Entity teamMember1 = new Entity("teammembership") { Id = Guid.NewGuid() };
            teamMember1["systemuserid"] = Guid.NewGuid();

            EntityCollection teamMembers = new EntityCollection();
            teamMembers.Entities.Add(teamMember1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(teamMembers);

            return serviceMock;
        }

        [TestMethod]
        public void TwoUserTeamWithOneExisting()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailToSend", new EntityReference("email", Guid.NewGuid()) },
                { "RecipientTeam", new EntityReference("team", Guid.NewGuid()) },
                { "SendEmail", false }
            };

            //Expected value
            const int expected = 3;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, TwoUserTeamSetup);

            //Test
            Assert.AreEqual(expected, output["UsersAdded"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> TwoUserTeamSetup(Mock<IOrganizationService> serviceMock)
        {
            Entity activityParty = new Entity("activityparty") { Id = Guid.NewGuid() };
            activityParty["partyid"] = new EntityReference("contact", Guid.NewGuid());
            EntityCollection existingRecipients = new EntityCollection();
            existingRecipients.Entities.Add(activityParty);

            Entity email = new Entity("email") { Id = Guid.NewGuid() };
            email["to"] = existingRecipients;

            Entity user1 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user1["internalemailaddress"] = "test1@test.com";

            Entity user2 = new Entity("systemuser") { Id = Guid.NewGuid() };
            user2["internalemailaddress"] = "test2@test.com";

            serviceMock.Setup(t =>
                t.Retrieve(It.IsAny<string>(), It.IsAny<Guid>(), It.IsAny<ColumnSet>()))
                .ReturnsInOrder(email, user1, user2);

            Entity teamMember1 = new Entity("teammembership") { Id = Guid.NewGuid() };
            teamMember1["systemuserid"] = Guid.NewGuid();

            Entity teamMember2 = new Entity("teammembership") { Id = Guid.NewGuid() };
            teamMember2["systemuserid"] = Guid.NewGuid();

            EntityCollection teamMembers = new EntityCollection();
            teamMembers.Entities.Add(teamMember1);
            teamMembers.Entities.Add(teamMember2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(teamMembers);

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
