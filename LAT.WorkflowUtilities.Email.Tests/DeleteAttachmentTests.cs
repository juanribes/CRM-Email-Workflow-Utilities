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
    public class DeleteAttachmentTests
    {
        #region Class Constructor
        private readonly string _namespaceClassAssembly;
        public DeleteAttachmentTests()
        {
            //[Namespace.class name, assembly name] for the class/assembly being tested
            //Namespace and class name can be found on the class file being tested
            //Assembly name can be found under the project properties on the Application tab
            _namespaceClassAssembly = "LAT.WorkflowUtilities.Email.DeleteAttachment" + ", " + "LAT.WorkflowUtilities.Email";
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
        public void DeleteZeroGreater()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteZeroGreaterSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteZeroGreaterSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneGreater()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneGreaterSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneGreaterSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteTwoGreater()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 2;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteTwoGreaterSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteTwoGreaterSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text1.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 500000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneGreaterOneOfTwo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 10000},
                { "DeleteSizeMin", 0 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneGreaterOneOfTwoSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneGreaterOneOfTwoSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 500000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteZeroLess()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 10000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteZeroLessSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteZeroLessSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 50000;
            attachment1["filename"] = "text.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneLess()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 10000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneLessSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneLessSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteTwoLess()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 10000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 2;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteTwoLessSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteTwoLessSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 3000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneLessOneOfTwo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 5000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneLessOneOfTwoSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneLessOneOfTwoSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 500000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteZeroMixed()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 75000},
                { "DeleteSizeMin", 3000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteZeroMixedSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteZeroMixedSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 50000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneMixed()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 75000},
                { "DeleteSizeMin", 3000 },
                { "Extensions" , null },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneMixedSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneMixedSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 500000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteZeroExtension()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteZeroExtensionSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteZeroExtensionSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text.docx";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 5000;
            attachment2["filename"] = "text.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneExtension()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneExtensionSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneExtensionSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text.pdf";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteTwoExtension()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 2;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteTwoExtensionSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteTwoExtensionSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text1.pdf";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 100000;
            attachment2["filename"] = "text2.pdf";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneExtensionOneOfTwo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneExtensionOneOfTwoSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneExtensionOneOfTwoSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text1.pdf";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 100000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneGreaterExtensionOneOfTwo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneGreaterExtensionOneOfTwoSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneGreaterExtensionOneOfTwoSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.pdf";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 5000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteOneLessExtensionOneOfTwo()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 0},
                { "DeleteSizeMin", 5000 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 1;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteOneLessExtensionOneOfTwoSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteOneLessExtensionOneOfTwoSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.pdf";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 5000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteZeroMixedExtension()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 75000},
                { "DeleteSizeMin", 3000 },
                { "Extensions" , "pdf" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 0;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteZeroMixedExtensionSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteZeroMixedExtensionSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 5000;
            attachment1["filename"] = "text1.pdf";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 50000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

            return serviceMock;
        }

        [TestMethod]
        public void DeleteTwoMultipleExtension()
        {
            //Target
            Entity targetEntity = null;

            //Input parameters
            var inputs = new Dictionary<string, object> 
            {
                { "EmailWithAttachments", new EntityReference("email", Guid.NewGuid()) },
                { "DeleteSizeMax", 1},
                { "DeleteSizeMin", 0 },
                { "Extensions" , "pdf,docx" },
                { "AppendNotice", false }
            };

            //Expected value
            const int expected = 2;

            //Invoke the workflow
            var output = InvokeWorkflow(_namespaceClassAssembly, ref targetEntity, inputs, DeleteTwoMultipleExtensionSetup);

            //Test
            Assert.AreEqual(expected, output["NumberOfAttachmentsDeleted"]);
        }

        /// <summary>
        /// Modify to mock CRM Organization Service actions
        /// </summary>
        /// <param name="serviceMock">The Organization Service to mock</param>
        /// <returns>Configured Organization Service</returns>
        private static Mock<IOrganizationService> DeleteTwoMultipleExtensionSetup(Mock<IOrganizationService> serviceMock)
        {
            //Attachment List
            Entity attachment1 = new Entity("activitymimeattachment");
            attachment1["filesize"] = 100000;
            attachment1["filename"] = "text1.pdf";

            Entity attachment2 = new Entity("activitymimeattachment");
            attachment2["filesize"] = 100000;
            attachment2["filename"] = "text2.docx";

            EntityCollection attachments = new EntityCollection();
            attachments.Entities.Add(attachment1);
            attachments.Entities.Add(attachment2);

            serviceMock.Setup(t =>
                t.RetrieveMultiple(It.IsAny<QueryExpression>()))
                .ReturnsInOrder(attachments);

            //New Note
            Guid newNoteId = new Guid();

            serviceMock.Setup(t =>
                t.Create(It.IsAny<Entity>()))
                .ReturnsInOrder(newNoteId);

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
