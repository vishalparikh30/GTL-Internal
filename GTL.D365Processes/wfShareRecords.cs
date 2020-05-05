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
    public class wfShareRecords : CodeActivity
    {
        /// <summary>
        /// This wf is created on 17th-Mar-2020
        /// based on practice BU and Geo BU
        /// currently works only for oppor. and on create action only.
        /// </summary>
        /// <param name="executionContext"></param>
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

            Entity currentEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));
            Entity AccountEntity = null;
            if (currentEntity.Contains("customerid") && currentEntity["customerid"] != null)
            {
                EntityReference AccountEntityRef = currentEntity["customerid"] as EntityReference;
                AccountEntity = service.Retrieve(AccountEntityRef.LogicalName, AccountEntityRef.Id, new ColumnSet(true));
            }
           // if (currentEntity.Contains("owningbusinessunit") && currentEntity["owningbusinessunit"] != null)
            //{
               // tracingService.Trace("Get Owning Bu");
               // EntityReference owningBURef = currentEntity["owningbusinessunit"] as EntityReference; //originating BU

               // Entity owningBU = service.Retrieve(owningBURef.LogicalName, owningBURef.Id, new ColumnSet(true));
               // if (owningBU.Contains("new_butype") && owningBU["new_butype"] != null)
               // {
                    //tracingService.Trace("Getbutype");
                   // if ((owningBU["new_butype"] as OptionSetValue).Value == 100000000) //geo Bu
                   // {
                       // tracingService.Trace("BU is Geo BU");
                        if (currentEntity.Contains("new_practicebu") && currentEntity["new_practicebu"] != null)
                        {
                            tracingService.Trace("Share record With Practice BU Management");
                            string sharedBUName = (currentEntity["new_practicebu"] as EntityReference).Name;
                            //find management team
                            tracingService.Trace("call default team function");
                            Entity sharedBUDefaultTeam = GetManagementTeam(tracingService, service, sharedBUName);
                            if (sharedBUDefaultTeam != null)
                            {
                                //share record with  practice BU default team 
                                tracingService.Trace("call share record function");
                                ShareRecords(tracingService, service, context, currentEntity, sharedBUDefaultTeam);
                                if (AccountEntity != null)
                                {
                                    tracingService.Trace("Share Account");
                                    ShareRecords(tracingService, service, context, AccountEntity, sharedBUDefaultTeam);
                                }

                            }

                            //flag = Geo
                        }
                        if(currentEntity.Contains("new_geosbu") && currentEntity["new_geosbu"] != null)
                        {
                            tracingService.Trace("Share record With Geo BU Management");
                            string sharedBUName = (currentEntity["new_geosbu"] as EntityReference).Name;
                            //find management team
                            tracingService.Trace("call default team function");
                            Entity sharedBUDefaultTeam = GetManagementTeam(tracingService, service, sharedBUName);
                            if (sharedBUDefaultTeam != null)
                            {
                                //share record with  practice BU default team 
                                tracingService.Trace("call share record function");
                                ShareRecords(tracingService, service, context, currentEntity, sharedBUDefaultTeam);
                                if (AccountEntity != null)
                                {
                                    tracingService.Trace("Share Account");
                                    ShareRecords(tracingService, service, context, AccountEntity, sharedBUDefaultTeam);
                                }

                            }
                        }
                        if (currentEntity.Contains("new_deliverybu") && currentEntity["new_deliverybu"] != null)
                        {
                            tracingService.Trace("Share record With delivery Management");
                            string sharedBUName = (currentEntity["new_deliverybu"] as EntityReference).Name;
                            //find management team
                            tracingService.Trace("call default team function");
                            Entity sharedBUDefaultTeam = GetManagementTeam(tracingService, service, sharedBUName);
                            if (sharedBUDefaultTeam != null)
                            {
                                //share record with  practice BU default team 
                                tracingService.Trace("call share record function");
                                ShareRecords(tracingService, service, context, currentEntity, sharedBUDefaultTeam);
                                if (AccountEntity != null)
                                {
                                    tracingService.Trace("Share Account");
                                    ShareRecords(tracingService, service, context, AccountEntity, sharedBUDefaultTeam);
                                }

                            }
                       }
                    //else if ((owningBU["new_butype"] as OptionSetValue).Value == 100000001) //practice BU
                    //{
                    //    tracingService.Trace("BU is Practice BU");

                    //    if (currentEntity.Contains("new_geosbu") && currentEntity["new_geosbu"] != null)
                    //    {

                    //        tracingService.Trace("Share record With geo BU management");
                    //        string sharedBUName = (currentEntity["new_geosbu"] as EntityReference).Name;
                    //        //find management team
                    //        tracingService.Trace("call default team function");
                    //        Entity sharedBUDefaultTeam = GetManagementTeam(tracingService, service, sharedBUName);
                    //        if (sharedBUDefaultTeam != null)
                    //        {
                    //            //share record with  Geo BU default team 
                    //            tracingService.Trace("call share record function");
                    //            ShareRecords(tracingService, service, context, currentEntity, sharedBUDefaultTeam);
                    //            if (AccountEntity != null)
                    //            {
                    //                tracingService.Trace("Share Account");
                    //                ShareRecords(tracingService, service, context, AccountEntity, sharedBUDefaultTeam);
                    //            }
                    //        }

                    //        //flag = P-Geo
                    //    }
                    //    if (currentEntity.Contains("new_deliverybu") && currentEntity["new_deliverybu"] != null && (currentEntity["new_deliverybu"] as EntityReference).Id != (currentEntity["owningbusinessunit"] as EntityReference).Id)
                    //    {
                    //        tracingService.Trace("get Delivery BU");
                    //        string sharedBUName = (currentEntity["new_deliverybu"] as EntityReference).Name;
                    //        //find management team
                    //        tracingService.Trace("call default team function");
                    //        Entity sharedBUDefaultTeam = GetManagementTeam(tracingService, service, sharedBUName);
                    //        if (sharedBUDefaultTeam != null)
                    //        {
                    //            //share record with  Geo BU default team 
                    //            tracingService.Trace("call share record function");
                    //            ShareRecords(tracingService, service, context, currentEntity, sharedBUDefaultTeam);
                    //            if (AccountEntity != null)
                    //            {
                    //                tracingService.Trace("Share Account");
                    //                ShareRecords(tracingService, service, context, AccountEntity, sharedBUDefaultTeam);
                    //            }
                    //        }
                    //    }

                    //}
                //}

            //}

        }

        private void ShareRecords(ITracingService tracingService, IOrganizationService service, IWorkflowContext context, Entity currentEntity, Entity defaultTeam)
        {
            tracingService.Trace("ShareRecords called");
            var grantUserAccessRequest = new GrantAccessRequest
            {
                PrincipalAccess = new PrincipalAccess
                {
                    AccessMask = AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AssignAccess | AccessRights.ShareAccess,
                    Principal = new EntityReference(defaultTeam.LogicalName, defaultTeam.Id)
                },
                Target = new EntityReference(currentEntity.LogicalName, currentEntity.Id)
            };
            service.Execute(grantUserAccessRequest);
        }

        private Entity GetManagementTeam(ITracingService tracingService, IOrganizationService service, string buName)
        {
            tracingService.Trace("GetDefaultTeam called");

            OrganizationServiceContext serviceContext = new OrganizationServiceContext(service);

            var qry = from c in serviceContext.CreateQuery("team")
                      where c["name"] == buName + "-Management"
                      select c;
            List<Entity> team = qry.ToList();
            if (team.Any())
            {
                return team.FirstOrDefault();
            }
            return null;
        }

    }
}
