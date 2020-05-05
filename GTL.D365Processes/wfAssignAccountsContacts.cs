using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Query;
using Microsoft.Xrm.Sdk.Workflow;
using System.Activities;
using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Messages;

namespace GTL.D365Processes
{
    /// <summary>
    /// This workflow will trigger, when BA assign the campaign to respective BM
    ///result:
    ///campaign ,its related marketingg list, contacts or account will be assigned to same BM as of campaign  and original entity(record) will be shared with creator(creartedby)
    /// </summary>
    public class wfAssignAccountsContacts : CodeActivity
    {

        protected override void Execute(CodeActivityContext executionContext)
        {

            #region CRM Init
            //Create the tracing service
            ITracingService tracingService = executionContext.GetExtension<ITracingService>();
            //Create the context

            IWorkflowContext context = executionContext.GetExtension<IWorkflowContext>();
            IOrganizationServiceFactory serviceFactory = executionContext.GetExtension<IOrganizationServiceFactory>();
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);
            #endregion
            tracingService.Trace(context.UserId.ToString());
            try
            {
                tracingService.Trace("start New");
                tracingService.Trace(context.PrimaryEntityName);
                tracingService.Trace(context.PrimaryEntityId.ToString());
                Entity currentEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(new string[] { "ownerid", "createdby" })); //campaign in our case
                tracingService.Trace(currentEntity.LogicalName);
                EntityReference refOwnerId = currentEntity["ownerid"] as EntityReference;
                //share the campaign with createdby
                EntityReference refSystemUser = currentEntity["createdby"] as EntityReference;
                tracingService.Trace("entity set");
                ShareRecord(refSystemUser, currentEntity, service, AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess, tracingService);

                //retrieve marketinglists associated with  campaign
                List<Guid> lstmarketingListids = RetrieveMarketingList(context.PrimaryEntityId, service, tracingService);
                tracingService.Trace("markwtting list retrieved");
                if (lstmarketingListids != null && lstmarketingListids.Count > 0)
                {
                    //retrieve detail of each marketlist
                    foreach (var marketingListid in lstmarketingListids)
                    {
                        Entity marketinglist = service.Retrieve("list", marketingListid, new ColumnSet(true));
                        if (marketinglist != null)
                        {
                            string targetEntity = string.Empty;
                            string[] columnnames = new string[1];
                            string fromAttribute = string.Empty;
                            ColumnSet cs = null;


                            //retrieve marketting list type(static/dynamic) and target entity(account/contact)
                            EntityCollection collection = null;

                            if (marketinglist.Contains("createdfromcode") && marketinglist["createdfromcode"] != null)
                            {

                                if ((marketinglist["createdfromcode"] as OptionSetValue).Value == 1) //account
                                {
                                    targetEntity = "account";
                                    columnnames[0] = "accountid";
                                    fromAttribute = "accountid";
                                    cs = new ColumnSet(new string[] { "ownerid", "createdby" });
                                }
                                else if ((marketinglist["createdfromcode"] as OptionSetValue).Value == 2) //contact
                                {
                                    targetEntity = "contact";
                                    columnnames[0] = "contactid";
                                    fromAttribute = "contactid";
                                    cs = new ColumnSet(new string[] { "ownerid", "createdby", "parentcustomerid" });
                                }

                                if (marketinglist.Contains("type") && marketinglist["type"] != null)
                                {

                                    if (Convert.ToBoolean(marketinglist["type"])) // dynamic list
                                    {
                                        collection = new EntityCollection();
                                        collection = GetDynamicTargetList(marketingListid, service, tracingService);
                                    }
                                    else //static list
                                    {
                                        //fetch static list
                                        collection = new EntityCollection();
                                        collection = GetStaticTargetList(targetEntity, columnnames, fromAttribute, marketingListid, service, tracingService);
                                    }
                                }
                            }



                            //share and assign the marketinglist with createdby
                            EntityReference refCreatdBy = marketinglist["createdby"] as EntityReference;
                            // Entity CreatdBy = service.Retrieve(refCreatdBy.LogicalName, refCreatdBy.Id, new ColumnSet(true));
                            ShareRecord(refCreatdBy, marketinglist, service, AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess, tracingService);
                            AssignRecord(marketinglist, refOwnerId, service, tracingService);
                            if (collection != null && collection.Entities != null && collection.Entities.Count > 0)
                            {
                                foreach (Entity e in collection.Entities)
                                {

                                    Entity ent = service.Retrieve(targetEntity, e.Id, cs);  //may be account or contact
                                    EntityReference refEntCreatdBy = ent["createdby"] as EntityReference;
                                    // Entity entCreatdBy = service.Retrieve(refEntCreatdBy.LogicalName, refEntCreatdBy.Id, new ColumnSet(new string[] { "ownerid", "createdby" }));
                                    ShareRecord(refEntCreatdBy, ent, service, AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess, tracingService);
                                    AssignRecord(ent, refOwnerId, service, tracingService);
                                    if (targetEntity == "contact")
                                    {
                                        if (ent.Contains("parentcustomerid") && ent["parentcustomerid"] != null)
                                        {
                                            Entity entParent = service.Retrieve((ent["parentcustomerid"] as EntityReference).LogicalName, (ent["parentcustomerid"] as EntityReference).Id, new ColumnSet(new string[] { "ownerid", "createdby" }));
                                            EntityReference refEntParentCreatdBy = entParent["createdby"] as EntityReference;
                                            Entity entParentCreatdBy = service.Retrieve(refEntParentCreatdBy.LogicalName, refEntParentCreatdBy.Id, new ColumnSet(true));
                                            ShareRecord(refEntParentCreatdBy, entParent, service, AccessRights.ReadAccess | AccessRights.WriteAccess | AccessRights.AppendAccess | AccessRights.AppendToAccess, tracingService);
                                            AssignRecord(entParent, refOwnerId, service, tracingService);
                                        }
                                    }

                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Exceptiom Message--Excute Method {0} Stack Trace {1}", ex.Message, ex.StackTrace));
            }

        }

        static void ShareRecord(EntityReference principal, Entity target, IOrganizationService service, AccessRights accessmask, ITracingService tracingService)
        {
            try
            {
                var grantUserAccessRequest = new GrantAccessRequest
                {
                    PrincipalAccess = new PrincipalAccess
                    {
                        AccessMask = accessmask,
                        Principal = principal
                    },
                    Target = new EntityReference(target.LogicalName, target.Id)
                };
                service.Execute(grantUserAccessRequest);
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Exceptiom Message--ShareRecord {0} Stack Trace {1},shared entity name {5}, shared with {2} , entity shared {3}, shared entity id {4}", ex.Message, ex.StackTrace, principal.Name, target.LogicalName, target.Id, principal.LogicalName));
            }

        }

        static List<Guid> RetrieveMarketingList(Guid campaignId, IOrganizationService service, ITracingService tracingService)
        {

            List<Guid> lstmarketingListids = null;
            try
            {
                QueryExpression query = new QueryExpression();
                query.EntityName = "list";
                query.ColumnSet = new ColumnSet("listid");
                // Get the relationship between contact and note
                Relationship relationship = new Relationship("campaignlist_association");
                relationship.PrimaryEntityRole = EntityRole.Referenced;

                // Create list of related entities
                RelationshipQueryCollection relatedEntity = new RelationshipQueryCollection();
                relatedEntity.Add(relationship, query);

                // Request to get a specified list record with the related records
                RetrieveRequest request = new RetrieveRequest();
                request.RelatedEntitiesQuery = relatedEntity;
                request.ColumnSet = new ColumnSet();
                request.Target = new EntityReference("campaign", campaignId);

                RetrieveResponse response = (RetrieveResponse)service.Execute(request);

                if (response.Entity.RelatedEntities[relationship].Entities != null
                    && response.Entity.RelatedEntities[relationship].Entities.Count > 0)
                {
                    lstmarketingListids = new List<Guid>();
                    foreach (var item in response.Entity.RelatedEntities[relationship].Entities)
                    {
                        lstmarketingListids.Add(new Guid(item["listid"].ToString()));
                    }
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Exceptiom Message--RetrieveMarketingList {0} Stack Trace {1}", ex.Message, ex.StackTrace));
            }

            return lstmarketingListids;
        }

        static EntityCollection GetStaticTargetList(string entityName, string[] columnlist, string fromAttribute, Guid listID, IOrganizationService service, ITracingService tracingService)
        {
            QueryExpression qe = new QueryExpression();
            qe.EntityName = entityName;
            //Initialize columnset
            ColumnSet col = new ColumnSet();
            EntityCollection bec = null;
            try
            {
                //add columns to columnset for the acc to retrieve each acc from the acc list
                col.AddColumns(columnlist);

                qe.ColumnSet = col;

                // link from account to listmember
                LinkEntity le = new LinkEntity();
                le.LinkFromEntityName = entityName;
                le.LinkFromAttributeName = fromAttribute;
                le.LinkToEntityName = "listmember";
                le.LinkToAttributeName = "entityid";

                //link from listmember to list
                LinkEntity le2 = new LinkEntity();
                le2.LinkFromEntityName = "listmember";
                le2.LinkFromAttributeName = "listid";
                le2.LinkToEntityName = "list";
                le2.LinkToAttributeName = "listid";

                // add condition for listid
                ConditionExpression ce = new ConditionExpression();
                ce.AttributeName = "listid";
                ce.Operator = ConditionOperator.Equal;
                ce.Values.Add(listID);

                //add condition to linkentity
                le2.LinkCriteria = new FilterExpression();
                le2.LinkCriteria.Conditions.Add(ce);

                //add linkentity2 to linkentity
                le.LinkEntities.Add(le2);
                qe.LinkEntities.Add(le);
                bec = new EntityCollection();
                bec = service.RetrieveMultiple(qe);
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Exceptiom Message--GetStaticTargetList {0} Stack Trace {1}, list ID {3}", ex.Message, ex.StackTrace, listID));
            }
            return bec;
        }

        static EntityCollection GetDynamicTargetList(Guid ListID, IOrganizationService service, ITracingService tracingService)
        {
            EntityCollection entities = null;
            try
            {
                Entity markettingList = service.Retrieve("list", ListID, new ColumnSet(true));

                if (markettingList.Contains("query"))
                {
                    entities = new EntityCollection();
                    FetchExpression query = new FetchExpression()
                    {
                        Query = markettingList["query"].ToString()
                    };

                    entities = service.RetrieveMultiple(query);
                }
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Exceptiom Message--GetDynamicTargetList {0} Stack Trace {1}, list ID {3}", ex.Message, ex.StackTrace, ListID));
            }
            return entities;

        }

        static void AssignRecord(Entity target, EntityReference assignee, IOrganizationService service, ITracingService tracingService)
        {
            try
            {
                AssignRequest req = new AssignRequest()
                {
                    Target = new EntityReference(target.LogicalName, target.Id),
                    Assignee = assignee
                };
                service.Execute(req);
            }
            catch (Exception ex)
            {
                tracingService.Trace(string.Format("Exceptiom Message--AssignRecord {0} Stack Trace {1},shared entity name {5}, shared with {2} , entity shared {3}, shared entity id {4}", ex.Message, ex.StackTrace, assignee.Name, target.LogicalName, target.Id, assignee.LogicalName));
            }
        }
    }
}
