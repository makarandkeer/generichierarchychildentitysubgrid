// <copyright file="CreateConfigXml.cs" company="">
// Copyright (c) 2016 All Rights Reserved
// </copyright>
// <author></author>
// <date>5/24/2016 8:44:33 PM</date>
// <summary>Implements the CreateConfigXml Workflow Activity.</summary>
namespace Generic.Hierarchy.Rollup.Actions
{
    using System;
    using System.Collections.Generic;
    using System.Text;
    using System.IO;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Activities;
    using System.Xml;
    using System.Xml.Linq;
    using System.Xml.Xsl;
    using System.Xml.XPath;
    using System.ServiceModel;
    using Microsoft.Xrm.Sdk;
    using Microsoft.Xrm.Sdk.Query;
    using Microsoft.Xrm.Sdk.Client;
    using Microsoft.Xrm.Sdk.Linq;
    using Microsoft.Xrm.Sdk.Messages;
    using Microsoft.Crm.Sdk.Messages;
    using Microsoft.Xrm.Sdk.Workflow;

    public sealed class CreateConfigXml : CodeActivity
    {
        #region Input Properties
        [Input("EnabledItems")]
        [Default("")]
        public InArgument<string> EnabledItems { get; set; }

        [Input("RefreshConfig")]
        [Default("false")]
        public InArgument<bool> RefreshConfig { get; set; }

        [Input("Namespace")]
        [Output("NamespaceOut")]
        public InOutArgument<string> Namespace { get; set; }

        [Input("EnableTrace")]
        [Output("EnableTraceOut")]
        public InOutArgument<bool> EnableTrace { get; set; }

        [Input("CreateConfigWR")]
        [Default("false")]
        public InArgument<bool> CreateConfigWR { get; set; }

        #endregion

        #region Output Properties
        [Output("ConfigHtml")]
        public OutArgument<string> ConfigHtml { get; set; }

        [Output("ConfigWRExists")]
        public OutArgument<bool> ConfigWRExists { get; set; }
        #endregion

        public const string XML_WEBRESOURCE = "mkisv_/xml/GenericHierarchyRollupXml.xml";
        public const string XSL_WEBRESOURCE = "mkisv_/xsl/GenericHierarchyRollupXsl.xsl";

        /// <summary>
        /// Executes the workflow activity.
        /// </summary>
        /// <param name="executionContext">The execution context.</param>
        protected override void Execute(CodeActivityContext executionContext)
        {
            // Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();

            if (tracingService == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve tracing service.");
            }

            tracingService.Trace("Entered CreateConfigXml.Execute(), Activity Instance Id: {0}, Workflow Instance Id: {1}",
                executionContext.ActivityInstanceId,
                executionContext.WorkflowInstanceId);

            // Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();

            if (context == null)
            {
                throw new InvalidPluginExecutionException("Failed to retrieve workflow context.");
            }

            tracingService.Trace("CreateConfigXml.Execute(), Correlation Id: {0}, Initiating User: {1}",
                context.CorrelationId,
                context.InitiatingUserId);

            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            try
            {
                ConfigWRExists.Set(executionContext, true);
                string ns = Namespace.Get(executionContext);
                bool traceEnable = EnableTrace.Get(executionContext);
                XDocument xDoc = null;
                Guid webresourceId = Guid.Empty;
                QueryByAttribute query = new QueryByAttribute("webresource");
                query.AddAttributeValue("name", XML_WEBRESOURCE);
                query.AddAttributeValue("webresourcetype", 4);
                query.ColumnSet = new ColumnSet("name", "content");
                EntityCollection ec = service.RetrieveMultiple(query);
                
                #region Create or Update existing XML WebResource

                if (ec == null || ec.Entities.Count == 0)
                {
                    ConfigWRExists.Set(executionContext, false);
                    if (CreateConfigWR.Get(executionContext))
                    {
                        #region 
                        xDoc = RetrieveEntityMetadata(service, ns: ns, traceEnable: traceEnable);
                        Entity xmlWebResource = new Entity("webresource");
                        xmlWebResource["name"] = XML_WEBRESOURCE;
                        xmlWebResource["displayname"] = "GenericHierarchyRollupXml.xml";
                        xmlWebResource["content"] = Convert.ToBase64String(UnicodeEncoding.UTF8.GetBytes(xDoc.ToString()));
                        xmlWebResource["webresourcetype"] = new OptionSetValue(4);
                        webresourceId = service.Create(xmlWebResource);
                        ec.Entities.Add(xmlWebResource);
                        PublishWebResource(webresourceId, service);
                        #endregion
                    }
                }
                else if (RefreshConfig.Get<bool>(executionContext))
                {
                    #region
                    xDoc = RetrieveEntityMetadata(service, ns: ns, traceEnable: traceEnable);
                    bool webResourceChanged = false;
                    byte[] binary = Convert.FromBase64String(ec.Entities[0].Attributes["content"].ToString());
                    XDocument xDocExisting = XDocument.Parse(UnicodeEncoding.UTF8.GetString(binary));

                    foreach (XElement xe in xDoc.Root.Elements("entity"))
                    {
                        if (xDocExisting.Root.Elements("entity").Where(a => a.Attribute("ReferencingEntity").Value == xe.Attribute("ReferencingEntity").Value && a.Attribute("ReferencingAttribute").Value == xe.Attribute("ReferencingAttribute").Value && a.Attribute("ReferencedEntity").Value == xe.Attribute("ReferencedEntity").Value).Count() == 0)
                        {
                            xDocExisting.Root.Add(xe);
                            webResourceChanged = true;
                        }
                    }

                    if (webResourceChanged)
                    {
                        ec.Entities[0].Attributes["content"] = Convert.ToBase64String(UnicodeEncoding.UTF8.GetBytes(xDocExisting.ToString()));
                        service.Update(ec.Entities[0]);
                        PublishWebResource(ec.Entities[0].Id, service);
                    }
                    #endregion
                }
                #endregion

                if (!string.IsNullOrEmpty(EnabledItems.Get<string>(executionContext)))
                {
                    #region Set IsEnable
                    string[] enabledItemList = EnabledItems.Get<string>(executionContext).Split(new char[] { ',' });
                    byte[] binary = Convert.FromBase64String(ec.Entities[0].Attributes["content"].ToString());
                    XDocument xDocExisting = XDocument.Parse(UnicodeEncoding.UTF8.GetString(binary));
                    UpdateConfigXML(enabledItemList, xDocExisting, ns: ns, traceEnable: traceEnable);

                    ec.Entities[0].Attributes["content"] = Convert.ToBase64String(UnicodeEncoding.UTF8.GetBytes(xDocExisting.ToString()));
                    service.Update(ec.Entities[0]);

                    PublishWebResource(ec.Entities[0].Id, service);
                    #endregion
                }

                if (ec != null && ec.Entities.Count > 0)
                {
                    XDocument xDocData = XDocument.Parse(UnicodeEncoding.UTF8.GetString(Convert.FromBase64String(ec.Entities[0].Attributes["content"].ToString())));

                    Namespace.Set(executionContext, xDocData.Root.Attribute("HierarchyColumnNamespace").Value);
                    EnableTrace.Set(executionContext, string.Compare(xDocData.Root.Attribute("EnableTrace").Value, "true", true) == 0 ? true : false);

                    ConfigHtml.Set(executionContext, XmlToHtmlTransformation(service, xDocData));
                }
                
            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                tracingService.Trace("Exception: {0}", e.ToString());

                // Handle the exception.
                throw;
            }

            tracingService.Trace("Exiting CreateConfigXml.Execute(), Correlation Id: {0}", context.CorrelationId);
        }

        public XDocument RetrieveEntityMetadata(IOrganizationService service, string ns = "new", bool traceEnable = false)
        {
            if(string.IsNullOrEmpty(ns))
            {
                ns = "new";
            }
            
            RetrieveAllEntitiesRequest req = new RetrieveAllEntitiesRequest();
            req.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Relationships;

            RetrieveAllEntitiesResponse ec = (RetrieveAllEntitiesResponse)service.Execute(req);
            XDocument xDoc = XDocument.Parse(string.Format("<root HierarchyColumnNamespace='{0}' EnableTrace='{1}'></root>", ns, traceEnable.ToString().ToLower()));

            foreach (var entity in ec.EntityMetadata.Where(a => a.IsChildEntity == false && a.IsIntersect == false && a.IsValidForAdvancedFind == true && a.CanCreateForms.Value == true && (a.OwnershipType == Microsoft.Xrm.Sdk.Metadata.OwnershipTypes.UserOwned || a.OwnershipType == Microsoft.Xrm.Sdk.Metadata.OwnershipTypes.OrganizationOwned)).OrderBy(a => a.LogicalName))
            {
                foreach (var one2Many in entity.ManyToOneRelationships.Where(a => a.ReferencedEntity != "systemuser" && a.ReferencedEntity != a.ReferencingEntity))
                {
                    var referencedEntityCollection = ec.EntityMetadata.Where(a => a.LogicalName == one2Many.ReferencedEntity).First().ManyToOneRelationships.Where(p => p.ReferencingEntity == p.ReferencedEntity && p.IsHierarchical == true);

                    foreach (var referencedEntity in referencedEntityCollection)
                    {
                        xDoc.Root.Add(new XElement("entity", new XAttribute("ReferencingEntity", one2Many.ReferencingEntity),
                            new XAttribute("ReferencingAttribute", one2Many.ReferencingAttribute),
                            new XAttribute("ReferencedEntity", one2Many.ReferencedEntity),
                            new XAttribute("ReferencedAttribute", one2Many.ReferencedAttribute),
                            new XAttribute("ReferencedParentAttribute", referencedEntity.ReferencingAttribute),
                            new XAttribute("IsEnabled", false)));
                    }
                }
            }
            return xDoc;
        }

        public string XmlToHtmlTransformation(IOrganizationService service, XDocument xDocData = null)
        {
            string xDocStyle = null;
            string HTMLString = string.Empty;

            QueryExpression query = new QueryExpression
            {
                EntityName = "webresource",
                ColumnSet = new ColumnSet("name", "content"),
                Criteria = new FilterExpression(LogicalOperator.Or)
                {
                    //mkisv_/xml/GenericHierarchyRollupXml.xml
                    Conditions = { 
                        new ConditionExpression("name", ConditionOperator.Equal, XSL_WEBRESOURCE), 
                        new ConditionExpression("name", ConditionOperator.Equal, XML_WEBRESOURCE) 
                    }
                }
            };

            EntityCollection ec = service.RetrieveMultiple(query);

            foreach (Entity e in ec.Entities)
            {
                if (e.GetAttributeValue<string>("name") == XSL_WEBRESOURCE)
                {
                    byte[] binary = Convert.FromBase64String(e.Attributes["content"].ToString());
                    xDocStyle = UnicodeEncoding.UTF8.GetString(binary);
                }
                else if (e.GetAttributeValue<string>("name") == XML_WEBRESOURCE && xDocData == null)
                {
                    byte[] binary = Convert.FromBase64String(e.Attributes["content"].ToString());
                    xDocData = XDocument.Parse(UnicodeEncoding.UTF8.GetString(binary));
                }
            }
            XPathDocument myXPathDoc = new XPathDocument(xDocData.CreateReader());
            XslCompiledTransform myXslTrans = new XslCompiledTransform();
            myXslTrans.Load(XmlReader.Create(new StringReader(xDocStyle)));

            using (var sw = new StringWriter())
            {
                using (var xw = XmlWriter.Create(sw))
                {
                    myXslTrans.Transform(myXPathDoc, null, xw);
                }
                HTMLString = sw.ToString();
            }
            return HTMLString;
        }

        public void UpdateConfigXML(string[] IdList, XDocument xDocData, string ns = "new", bool traceEnable = false)
        {
            string ReferencingEntity = string.Empty;
            string ReferencingAttribute = string.Empty;
            string ReferencedEntity = string.Empty;

            foreach (XElement xe in xDocData.Root.Elements("entity"))
            {
                xe.Attribute("IsEnabled").Value = false.ToString().ToLower();
            }

            xDocData.Root.Attribute("HierarchyColumnNamespace").Value = ns;
            xDocData.Root.Attribute("EnableTrace").Value = traceEnable.ToString().ToLower();

            foreach (string id in IdList.Where(a=>a.Trim().Length > 0))
            {
                string[] columnArray = id.Split(new char[] { '|' });
                ReferencingEntity = columnArray[0];
                ReferencingAttribute = columnArray[1];
                ReferencedEntity = columnArray[2];
                var xEntity = xDocData.Root.Elements("entity").Where(xe => xe.Attribute("ReferencingEntity").Value == ReferencingEntity && xe.Attribute("ReferencingAttribute").Value == ReferencingAttribute && xe.Attribute("ReferencedEntity").Value == ReferencedEntity).FirstOrDefault<XElement>();
                if (xEntity != null)
               {
                   xEntity.Attribute("IsEnabled").Value = true.ToString().ToLower();
               }
            }
        }

        private void PublishWebResource(Guid webresourceId, IOrganizationService service)
        {
            OrganizationRequest request = new OrganizationRequest { RequestName = "PublishXml" };

            request.Parameters = new ParameterCollection();
            request.Parameters.Add(new KeyValuePair<string, object>("ParameterXml",
            string.Format("<importexportxml><webresources>{0}</webresources></importexportxml>",
            string.Format("<webresource>{0}</webresource>", webresourceId)
            )));
            service.Execute(request);
        }
    }
}