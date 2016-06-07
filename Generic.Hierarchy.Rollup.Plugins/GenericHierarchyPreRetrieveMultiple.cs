using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;
using System.Xml.XPath;

using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Linq;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Crm.Sdk.Messages;

namespace Generic.Hierarchy.Rollup.Plugins
{
    public class GenericHierarchyPreRetrieveMultiple : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {
            IPluginExecutionContext executionContext = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));
            IOrganizationServiceFactory serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(executionContext.InitiatingUserId);
            OrganizationServiceContext sContext = new OrganizationServiceContext(service);
            ITracingService Trace = (ITracingService)serviceProvider.GetService(typeof(ITracingService));
            try
            {
                string ReferencingEntity = executionContext.PrimaryEntityName;
                string ReferencingAttribute = string.Empty;
                string ReferencedEntity = string.Empty;
                string ReferencedAttribute = string.Empty;
                string ReferencedParentAttribute = string.Empty;
                string HierarchyColumnNamespace = string.Empty;

                bool TraceEnabled = false;

  //<entity ReferencingEntity="incident" ReferencingAttribute="subjectid" ReferencedEntity="subject" ReferencedAttribute="subjectid" ReferencedParentAttribute="parentsubject" IsEnabled="false" />

                if (executionContext.MessageName == "RetrieveMultiple")
                {
                    foreach (var v in executionContext.InputParameters)
                    {
                        if (v.Key == "Query" && v.Value.GetType() == typeof(Microsoft.Xrm.Sdk.Query.QueryExpression))
                        {
                            XDocument xDoc = ReadGenericHierarchyRollupXml(service);
                            if (xDoc == null)
                            {
                                CreateTraceNote("GenericHierarchyPreRetrieveMultiple - xDoc", string.Format("No xDoc found for ReferencingEntity {0}", ReferencingEntity), service);
                                return;
                            }
                            HierarchyColumnNamespace = xDoc.Root.Attribute("HierarchyColumnNamespace").Value;
                            TraceEnabled = string.Compare(xDoc.Root.Attribute("EnableTrace").Value, "true", true) == 0 ? true : false;
                            var xmlConfig = xDoc.Root.Elements("entity").Where(a => a.Attribute("ReferencingEntity").Value == ReferencingEntity && string.Compare(a.Attribute("IsEnabled").Value ,"true", true) == 0);
                            if (xmlConfig == null)
                            {
                                if (TraceEnabled)
                                {
                                    CreateTraceNote("GenericHierarchyPreRetrieveMultiple - xmlConfig", string.Format("No xmlConfig found for ReferencingEntity {0}", ReferencingEntity), service);
                                }
                                return;
                            }

                            QueryExpression query = (QueryExpression)v.Value;

                            if (TraceEnabled)
                            {
                                QueryExpressionToFetchXmlRequest req001 = new QueryExpressionToFetchXmlRequest();
                                req001.Query = query;
                                QueryExpressionToFetchXmlResponse res001 = (QueryExpressionToFetchXmlResponse)service.Execute(req001);
                                CreateTraceNote("GenericHierarchyPreRetrieveMultiple-Original", res001.FetchXml, service);
                            }
                            if (query.ColumnSet.AllColumns == false && query.ColumnSet.Columns.Contains(HierarchyColumnNamespace + "_isretrievehierarchy"))
                            {
                                foreach (var xe in xmlConfig)
                                {
                                    ReferencingAttribute = xe.Attribute("ReferencingAttribute").Value;
                                    ReferencedEntity = xe.Attribute("ReferencedEntity").Value;
                                    ReferencedAttribute = xe.Attribute("ReferencedAttribute").Value;
                                    ReferencedParentAttribute = xe.Attribute("ReferencedParentAttribute").Value;

                                    if (query.Criteria.Conditions.Where(a => a.AttributeName == ReferencingAttribute).Count() > 0)
                                    {
                                        #region if condition
                                        var referenceEntityIdCondition = query.Criteria.Conditions.Where(a => a.AttributeName == ReferencingAttribute).FirstOrDefault();
                                        FilterExpression hierarchyFilter = new FilterExpression(LogicalOperator.Or);

                                        hierarchyFilter.AddCondition(new ConditionExpression(ReferencedAttribute, ConditionOperator.UnderOrEqual, (Guid)referenceEntityIdCondition.Values[0]));

                                        LinkEntity accountLink = new LinkEntity()
                                        {
                                            LinkFromEntityName = executionContext.PrimaryEntityName,
                                            LinkToEntityName = ReferencedEntity,
                                            LinkFromAttributeName = ReferencingAttribute,
                                            LinkToAttributeName = ReferencedAttribute,
                                            LinkCriteria = hierarchyFilter
                                        };
                                        query.LinkEntities.Add(accountLink);

                                        query.Criteria.Conditions.Remove(referenceEntityIdCondition);

                                        if (TraceEnabled)
                                        {
                                            QueryExpressionToFetchXmlRequest req001 = new QueryExpressionToFetchXmlRequest();
                                            req001.Query = query;
                                            QueryExpressionToFetchXmlResponse res001 = (QueryExpressionToFetchXmlResponse)service.Execute(req001);
                                            CreateTraceNote("GenericHierarchyPreRetrieveMultiple-Modified", res001.FetchXml, service);
                                        }
                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Trace.Trace(ex.Message, ex);
                throw new InvalidPluginExecutionException("Exception on Generic.Hierarchy.Rollup.Plugins.GenericHierarchyPreRetrieveMultiple Plugin: " + ex.Message);
            }
        }

        private XDocument ReadGenericHierarchyRollupXml(IOrganizationService service)
        {
            QueryByAttribute query = new QueryByAttribute("webresource");
            query.AddAttributeValue("name", "mkisv_/xml/GenericHierarchyRollupXml.xml");
            query.ColumnSet = new ColumnSet(true);

            EntityCollection ec = service.RetrieveMultiple(query);
            if (ec != null && ec.Entities.Count > 0)
            {
                byte[] binary = Convert.FromBase64String(ec.Entities[0].Attributes["content"].ToString());
                return XDocument.Parse(UnicodeEncoding.UTF8.GetString(binary));
            }
            return null;
        }

        private void CreateTraceNote(string subject, string noteText, IOrganizationService service)
        {
            Entity e = new Entity("annotation");
            e["subject"] = subject;
            e["notetext"] = noteText;
            service.Create(e);
        }
    }
}
