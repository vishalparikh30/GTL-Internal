using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Client;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Crm.Sdk.Messages;

namespace GTL.D365Processes
{
    public class wfShareAndAssignRecord : CodeActivity
    {
        [RequiredArgument]
        [Input("Root Business unit")]
        [ReferenceTarget("businessunit")]
        public InArgument<EntityReference> argRootBu { get; set; }
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

            //get current entity (lead,account etc)
            Entity currentEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));

            //get current user entity 
            Entity currentUser = service.Retrieve("systemuser", context.InitiatingUserId, new ColumnSet(true));

            EntityReference businessunitref = argRootBu.Get<EntityReference>(executionContext);

            if (currentUser.Contains("businessunitid") && currentUser["businessunitid"] != null && businessunitref.Id != (currentUser["businessunitid"] as EntityReference).Id)
            {
                //get the current user business unit entity

                EntityReference entityReferenceCurrentUserBU = currentUser["businessunitid"] as EntityReference;
                Entity entityCurrentBusuinessUnit = service.Retrieve(entityReferenceCurrentUserBU.LogicalName, entityReferenceCurrentUserBU.Id, new ColumnSet(true));

                //get current business unit's parent business unit
                Entity entParentBusinesssUnit = entityCurrentBusuinessUnit;


                if (entityCurrentBusuinessUnit.Contains("new_isdeliverybu") && Convert.ToBoolean(entityCurrentBusuinessUnit["new_isdeliverybu"]) == false)// false means sub business unit ,so find parent business unit
                {

                    if (entityCurrentBusuinessUnit.Contains("parentbusinessunitid") && entityCurrentBusuinessUnit["parentbusinessunitid"] != null)
                    {

                        EntityReference refParentBusinesssUnit = entityCurrentBusuinessUnit["parentbusinessunitid"] as EntityReference;
                        entParentBusinesssUnit = service.Retrieve(refParentBusinesssUnit.LogicalName, refParentBusinesssUnit.Id, new ColumnSet(true));

                    }
                }

                tracingService.Trace("Good");
                //get delivery business unit entity
                if (currentEntity.Contains("new_deliverysbuid") && currentEntity["new_deliverysbuid"] != null && currentEntity.Contains("new_indirectsalesbuid") && currentEntity["new_indirectsalesbuid"] != null)
                {
                    tracingService.Trace("Good  offf");
                    Entity entityTeam = null;
                    Entity entityManagementTeam = null;
                    EntityReference refDeliveryBU = currentEntity["new_deliverysbuid"] as EntityReference;
                    Entity entityDeliveryBU = service.Retrieve(refDeliveryBU.LogicalName, refDeliveryBU.Id, new ColumnSet(true));

                    EntityReference refSalesBU = currentEntity["new_indirectsalesbuid"] as EntityReference;
                    Entity entSalesBU = service.Retrieve(refSalesBU.LogicalName, refSalesBU.Id, new ColumnSet(true));


                    if (entityDeliveryBU.Id != entParentBusinesssUnit.Id) //if not equal assign record to delivery BU and share record with current user
                    {
                        using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(service))
                        {
                            //get the sales team 
                            var qryTeams = from team in orgSvcContext.CreateQuery("team")
                                           where (team["businessunitid"] as EntityReference).Id == entSalesBU.Id
                                           && team["name"] == entityDeliveryBU["name"]
                                           select team;
                            if (qryTeams != null)
                            {
                                List<Entity> lstTeams = qryTeams.ToList<Entity>();
                                if (lstTeams != null && lstTeams.Count > 0)
                                {
                                    entityTeam = new Entity();
                                    entityTeam = lstTeams.FirstOrDefault();
                                }
                            }
                            // tracingService.Trace(lstTeams.Count.ToString());
                            tracingService.Trace("sales team sucess---");
                        }

                        using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(service))
                        {
                            //get the users' management team  
                            var qryTeams = from team in orgSvcContext.CreateQuery("team")
                                           where team["name"] == Convert.ToString(entParentBusinesssUnit["name"]) + "-Management"
                                           select team;
                            if (qryTeams != null)
                            {
                                List<Entity> lstTeams = qryTeams.ToList<Entity>();
                                if (lstTeams != null && lstTeams.Count > 0)
                                {
                                    entityManagementTeam = new Entity();
                                    entityManagementTeam = lstTeams.FirstOrDefault();
                                }
                            }
                            tracingService.Trace(entParentBusinesssUnit["name"].ToString() + "-Management");
                            //tracingService.Trace(lstTeams.Count.ToString());
                        }
                        if (entityTeam != null && entityManagementTeam != null)
                        {
                            // Grant current  user read , write,assign and share access to the created record by him/herself.
                            var grantUserAccessRequest = new GrantAccessRequest
                            {
                                PrincipalAccess = new PrincipalAccess
                                {
                                    AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AssignAccess | AccessRights.ShareAccess,
                                    Principal = new EntityReference(currentUser.LogicalName, context.InitiatingUserId)
                                },
                                Target = new EntityReference(currentEntity.LogicalName, currentEntity.Id)
                            };
                            service.Execute(grantUserAccessRequest);


                            //grand users' creator users' team manager team only read access as per documents 
                            var grantTeamManagementAccessRequest = new GrantAccessRequest
                            {
                                PrincipalAccess = new PrincipalAccess
                                {
                                    AccessMask = AccessRights.ReadAccess,
                                    Principal = new EntityReference(currentUser.LogicalName, context.InitiatingUserId)
                                },
                                Target = new EntityReference(currentEntity.LogicalName, currentEntity.Id)
                            };
                            service.Execute(grantTeamManagementAccessRequest);


                            currentEntity["ownerid"] = new EntityReference(entityTeam.LogicalName, entityTeam.Id);// or  use below code to assign the request
                            //AssignRequest req = new AssignRequest()
                            //{
                            //    Target = new EntityReference(currentEntity.LogicalName, currentEntity.Id),
                            //    Assignee = new EntityReference(entityTeam.LogicalName, entityTeam.Id)
                            //};
                            //service.Execute(req);
                            service.Update(currentEntity);

                        }

                    }
                }

            }
        }

    }

}
