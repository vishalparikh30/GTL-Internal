using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using System.ServiceModel;

namespace GTL.D365Processes
{
    /// <summary>
    /// This WF copies all the activities from Contact to Lead
    /// </summary>
    public class wfPropagateActivities : CodeActivity
    {
        [Input("Contact")]
        [ReferenceTarget("contact")]
        public InArgument<EntityReference> argLead { get; set; }
        protected override void Execute(CodeActivityContext executionContext)
        {
            #region CRM Init
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            //Create the context
            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.InitiatingUserId);
            #endregion
            tracingService.Trace("0");

            EntityCollection lstActivities = RetrieveDataQueryExpression("activitypointer", service,
                new string[] { "activityid", "regardingobjectid", "activitytypecode", "subject", "statecode" }, "regardingobjectid", ConditionOperator.Equal, argLead.Get<EntityReference>(executionContext).Id);
            tracingService.Trace("1");
            foreach (var item in lstActivities.Entities)
            {
                if ((item["statecode"] as OptionSetValue).Value != 0)
                    continue;

                Entity e = new Entity(item["activitytypecode"].ToString(), item.Id);
                e["regardingobjectid"] = new EntityReference(context.PrimaryEntityName, context.PrimaryEntityId);
                tracingService.Trace("2");

                service.Update(e);
            }

        }
        public static EntityCollection RetrieveDataQueryExpression(string entityName, IOrganizationService client, string[] fields, string conditionAttributeName, ConditionOperator conditionOperator, object conditonValue)
        {
            EntityCollection entityCollection = null;
            try
            {
                QueryExpression configQuery = new QueryExpression()
                {
                    EntityName = entityName,
                    ColumnSet = new ColumnSet(fields),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                      {
                            new ConditionExpression(conditionAttributeName, conditionOperator,conditonValue)
                      }
                    }
                };
                entityCollection = new EntityCollection();
                entityCollection = client.RetrieveMultiple(configQuery);
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                // HealthMonitor.CreateLog("FaultException. Error in GetEntityMetadata.",
                // ex, System.Diagnostics.EventLogEntryType.Error);

                throw ex;

            }
            catch (System.TimeoutException ex)
            {
                //HealthMonitor.CreateLog("TimeoutException. Error in GetEntityMetadata.",
                // ex, System.Diagnostics.EventLogEntryType.Error);

                throw ex;
            }
            catch (Exception ex)
            {
                //HealthMonitor.CreateLog("Exception. Error in GetEntityMetadata.",
                //ex, System.Diagnostics.EventLogEntryType.Error);
                throw ex;
            }

            return entityCollection;
        }
    }
}
