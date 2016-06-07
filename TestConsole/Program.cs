using System;
using System.Linq;
using System.Linq.Expressions;
using System.Text;
using System.Net;
using System.ServiceModel;
using System.ServiceModel.Description;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using System.Xml.XPath;
using System.Configuration;
using System.Collections.ObjectModel;
using System.IO;
using System.Runtime.Serialization;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Linq;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;

using Microsoft.Xrm.Client;
using Microsoft.Xrm.Client.Services;

namespace TestConsole
{
    class Program
    {
        private static IOrganizationService _orgSvc;

        static void Main(string[] args)
        {
            InitializeOrgService(Guid.Empty);
            CallAction();
            //ExecuteCreateConfigXml();
        }

        private static void XmlToHtmlTransform()
        {
            XDocument xDocData = null;
            string xDocStyle = null;

            QueryExpression query = new QueryExpression
            {
                EntityName = "webresource",
                ColumnSet = new ColumnSet("name", "content"),
                Criteria = new FilterExpression(LogicalOperator.Or)
                {
                    //mkisv_/xml/GenericHierarchyRollupXml.xml
                    Conditions = { 
                        new ConditionExpression("name", ConditionOperator.Equal, "mkisv_/xsl/GenericHierarchyRollupXsl.xsl"), 
                        new ConditionExpression("name", ConditionOperator.Equal, "mkisv_/xml/GenericHierarchyRollupXml.xml") 
                    }
                }
            };

            EntityCollection ec = _orgSvc.RetrieveMultiple(query);

            foreach(Entity e in ec.Entities)
            {
                if(e.GetAttributeValue<string>("name") == "mkisv_/xsl/GenericHierarchyRollupXml.xsl")
                {
                    byte[] binary = Convert.FromBase64String(e.Attributes["content"].ToString());
                    xDocStyle = UnicodeEncoding.UTF8.GetString(binary);
                }
                else if (e.GetAttributeValue<string>("name") == "mkisv_/xml/GenericHierarchyRollupXml.xml")
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
            }
        }

        private static void InitializeOrgService(Guid CallerId)
        {
            String connectionString = GetServiceConfiguration();

            // Establish a connection to the organization web _orgSvc using CrmConnection.
            Microsoft.Xrm.Client.CrmConnection connection = CrmConnection.Parse(connectionString);

            if (CallerId != null && CallerId != Guid.Empty)
            {
                connection.CallerId = CallerId;
            }
            _orgSvc = new OrganizationService(connection);
        }

        /// <summary>
        /// Gets web _orgSvc connection information from the app.config file.
        /// If there is more than one available, the user is prompted to select
        /// the desired connection configuration by name.
        /// </summary>
        /// <returns>A string containing web _orgSvc connection configuration information.</returns>
        public static String GetServiceConfiguration()
        {
            // Get available connection strings from app.config.
            int count = ConfigurationManager.ConnectionStrings.Count;

            // Create a filter list of connection strings so that we have a list of valid
            // connection strings for Microsoft Dynamics CRM only.
            List<KeyValuePair<String, String>> filteredConnectionStrings = new List<KeyValuePair<String, String>>();

            for (int a = 0; a < count; a++)
            {
                if (isValidConnectionString(ConfigurationManager.ConnectionStrings[a].ConnectionString))
                    filteredConnectionStrings.Add(new KeyValuePair<string, string>
                                                    (ConfigurationManager.ConnectionStrings[a].Name,
                                                    ConfigurationManager.ConnectionStrings[a].ConnectionString));
            }

            // No valid connections strings found. Write out and error message.
            if (filteredConnectionStrings.Count == 0)
            {
                throw new Exception("No valid CRM Connection string found");
            }

            // If one valid connection string is found, use that.
            if (filteredConnectionStrings.Count == 1)
            {
                return filteredConnectionStrings[0].Value;
            }

            // If more than one valid connection string is found, let the user decide which to use.
            if (filteredConnectionStrings.Count > 1)
            {
                throw new Exception("More than 1 CRM Connections strings found. Only one CRM Connection is expected.");
            }
            return null;
        }

        /// <summary>
        /// Verifies if a connection string is valid for Microsoft Dynamics CRM.
        /// </summary>
        /// <returns>True for a valid string, otherwise False.</returns>
        public static Boolean isValidConnectionString(String connectionString)
        {
            // At a minimum, a connection string must contain one of these arguments.
            if (connectionString.Contains("Url=") ||
                connectionString.Contains("Server=") ||
                connectionString.Contains("ServiceUri="))
                return true;
            return false;
        }

        public static void CallAction()
        {
            //mkisv_GenericHierarchyAction
            OrganizationRequest req = new OrganizationRequest("mkisv_GenericHierarchyAction");
            //req["EnabledEntityIDs"] = "connection|record1id|knowledgearticle";
            req["CustomFieldPrefix"] = "abcd";
            req["RefreshConfigXML"] = false;
            req["EnabledEntityIDs"] = string.Empty;
            req["CreateConfigWebResource"] = false;
            req["EnableTrace"] = true;

            OrganizationResponse response = _orgSvc.Execute(req);
        }


        public static void ExecuteCreateConfigXml()
        {
            string EnabledItems =  "connection|record1id|knowledgearticle";
            EnabledItems = string.Empty;
            try
            {
                Guid webresourceId = Guid.Empty;

                XDocument xDoc = RetrieveEntityMetadata(_orgSvc);

                QueryByAttribute query = new QueryByAttribute("webresource");
                query.AddAttributeValue("name", "mkisv_/xml/GenericHierarchyRollupXml.xml");
                query.AddAttributeValue("webresourcetype", 4);
                query.ColumnSet = new ColumnSet("name", "content");
                EntityCollection ec = _orgSvc.RetrieveMultiple(query);

                if (xDoc != null)
                {
                    #region Create or Update existing XML WebResource

                    if (ec == null || ec.Entities.Count == 0)
                    {
                        Entity xmlWebResource = new Entity("webresource");
                        xmlWebResource["name"] = "mkisv_/xml/GenericHierarchyRollupXml.xml";
                        xmlWebResource["displayname"] = "GenericHierarchyRollupXml.xml";
                        xmlWebResource["content"] = Convert.ToBase64String(UnicodeEncoding.UTF8.GetBytes(xDoc.ToString()));
                        xmlWebResource["webresourcetype"] = new OptionSetValue(4);
                        webresourceId = _orgSvc.Create(xmlWebResource);
                        PublishWebResource(webresourceId, _orgSvc);
                    }
                    else
                    {
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
                            _orgSvc.Update(ec.Entities[0]);
                            PublishWebResource(ec.Entities[0].Id, _orgSvc);
                        }
                    }

                    #endregion
                }

                if (!string.IsNullOrEmpty(EnabledItems))
                {
                    string[] enabledItemList = EnabledItems.Split(new char[] { ',' });
                    byte[] binary = Convert.FromBase64String(ec.Entities[0].Attributes["content"].ToString());
                    XDocument xDocExisting = XDocument.Parse(UnicodeEncoding.UTF8.GetString(binary));
                    UpdateConfigXML(enabledItemList, _orgSvc, xDocExisting);

                    ec.Entities[0].Attributes["content"] = Convert.ToBase64String(UnicodeEncoding.UTF8.GetBytes(xDocExisting.ToString()));
                    _orgSvc.Update(ec.Entities[0]);

                    PublishWebResource(ec.Entities[0].Id, _orgSvc);
                }

                XDocument xDocData = XDocument.Parse(UnicodeEncoding.UTF8.GetString(Convert.FromBase64String(ec.Entities[0].Attributes["content"].ToString())));
                string ConfigHtml = XmlToHtmlTransformation(_orgSvc, xDocData);

            }
            catch (FaultException<OrganizationServiceFault> e)
            {
                // Handle the exception.
                throw;
            }
        }

        public static XDocument RetrieveEntityMetadata(IOrganizationService _orgSvc)
        {
            RetrieveAllEntitiesRequest req = new RetrieveAllEntitiesRequest();
            req.EntityFilters = Microsoft.Xrm.Sdk.Metadata.EntityFilters.Entity | Microsoft.Xrm.Sdk.Metadata.EntityFilters.Relationships;

            RetrieveAllEntitiesResponse ec = (RetrieveAllEntitiesResponse)_orgSvc.Execute(req);
            XDocument xDoc = XDocument.Parse("<root HierarchyColumnNamespace='new' EnableTrace='false'></root>");

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

        public static string XmlToHtmlTransformation(IOrganizationService _orgSvc, XDocument xDocData = null)
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
                        new ConditionExpression("name", ConditionOperator.Equal, "mkisv_/xsl/GenericHierarchyRollupXsl.xsl"), 
                        new ConditionExpression("name", ConditionOperator.Equal, "mkisv_/xml/GenericHierarchyRollupXml.xml") 
                    }
                }
            };

            EntityCollection ec = _orgSvc.RetrieveMultiple(query);

            foreach (Entity e in ec.Entities)
            {
                if (e.GetAttributeValue<string>("name") == "mkisv_/xsl/GenericHierarchyRollupXsl.xsl")
                {
                    byte[] binary = Convert.FromBase64String(e.Attributes["content"].ToString());
                    xDocStyle = UnicodeEncoding.UTF8.GetString(binary);
                }
                else if (e.GetAttributeValue<string>("name") == "mkisv_/xml/GenericHierarchyRollupXml.xml" && xDocData == null)
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

        public static void UpdateConfigXML(string[] IdList, IOrganizationService _orgSvc, XDocument xDocData)
        {
            string ReferencingEntity = string.Empty;
            string ReferencingAttribute = string.Empty;
            string ReferencedEntity = string.Empty;

            foreach (string id in IdList)
            {
                string[] columnArray = id.Split(new char[] { '|' });
                ReferencingEntity = columnArray[0];
                ReferencingAttribute = columnArray[1];
                ReferencedEntity = columnArray[2];
            }

            foreach (XElement xe in xDocData.Root.Elements("entity"))
            {
                xe.Attribute("IsEnabled").Value = false.ToString().ToLower();
                if (xe.Attribute("ReferencingEntity").Value == ReferencingEntity && xe.Attribute("ReferencingAttribute").Value == ReferencingAttribute && xe.Attribute("ReferencedEntity").Value == ReferencedEntity)
                {
                    xe.Attribute("IsEnabled").Value = true.ToString().ToLower();
                }
            }
        }

        private static void PublishWebResource(Guid webresourceId, IOrganizationService _orgSvc)
        {
            OrganizationRequest request = new OrganizationRequest { RequestName = "PublishXml" };

            request.Parameters = new ParameterCollection();
            request.Parameters.Add(new KeyValuePair<string, object>("ParameterXml",
            string.Format("<importexportxml><webresources>{0}</webresources></importexportxml>",
            string.Format("<webresource>{0}</webresource>", webresourceId)
            )));

            OrganizationResponse res = _orgSvc.Execute(request);
        }
    }
}
