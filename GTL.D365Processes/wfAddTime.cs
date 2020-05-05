using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Activities;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Workflow;
using Microsoft.Xrm.Sdk.Query;
using System.ServiceModel;

namespace GTL.D365Processes
{
    public class wfAddTime : CodeActivity
    {
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
            tracingService.Trace("start");
            var currentUserSettings = service.RetrieveMultiple(
                new QueryExpression("usersettings")
                {
                    ColumnSet = new ColumnSet("localeid", "timezonecode"),
                    Criteria = new FilterExpression
                    {
                        Conditions =
                        {
                            new ConditionExpression("systemuserid", ConditionOperator.EqualUserId)
                        }
                    }
                }).Entities[0];

            //TimeZoneInfo timeOffSet = (ubl.RetrieveTimeZoneDefinitionForCurrentUser(new string[] { "standardname" }, svc, Convert.ToInt16(userSettings["timezonecode"])));

            EntityCollection entityCollection = null;
            try
            {
                string[] fields = new string[] { "standardname" };
                entityCollection = RetrieveDataQueryExpression("timezonedefinition", service, fields, "timezonecode", ConditionOperator.Equal, currentUserSettings["timezonecode"]);

            }
            catch (Exception ex)
            {
                throw ex;
            }
            tracingService.Trace("before");
            TimeZoneInfo offset = (TimeZoneInfo.FindSystemTimeZoneById(Convert.ToString(entityCollection.Entities.First()["standardname"])));
            tracingService.Trace("after");
            Entity task = new Entity("task");
            task["subject"] = "My Test";
            tracingService.Trace("before--1");
            tracingService.Trace(DateTime.UtcNow.ToString());
            task["scheduledend"] = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow.AddHours(1), offset);
            tracingService.Trace(TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow.AddHours(1), offset).ToString());
            tracingService.Trace("after--1");
            service.Create(task);
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
