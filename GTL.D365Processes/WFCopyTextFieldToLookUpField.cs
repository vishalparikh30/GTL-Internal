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


namespace GTL.D365Processes
{
    /// <summary>
    /// Use for coping single line type text field of D365's any entity to look up field 
    /// e.g. coping camapign name text field of contact entity to campaign look up field of contact entity
    /// </summary>
    public class WFCopyTextFieldToLookUpField : CodeActivity
    {
        [RequiredArgument]
        [Input("Source String Field")]
        public InArgument<string> argSourceStringField { get; set; }

        [RequiredArgument]
        [Input("LookUp Entity")]
        public InArgument<string> argLookUpEntityName { get; set; }

        [RequiredArgument]
        [Input("LookUp Entity Field")]
        public InArgument<string> argLookUpEntityField { get; set; }

        [RequiredArgument]
        [Input("Destination lookUp Field Field")]
        public InArgument<string> argDestinationlookUpField { get; set; }
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

            Entity currentEntity = service.Retrieve(context.PrimaryEntityName, context.PrimaryEntityId, new ColumnSet(true));

            #region get the input fields
            string SourceStringField = argSourceStringField.Get<string>(executionContext);
            string LookUpEntityName = argLookUpEntityName.Get<string>(executionContext);
            string LookUpEntityField = argLookUpEntityField.Get<string>(executionContext);
            string DestinationlookUpField = argDestinationlookUpField.Get<string>(executionContext);
            #endregion
         

            if (currentEntity.Contains(SourceStringField) && (!string.IsNullOrEmpty(Convert.ToString(currentEntity[SourceStringField]))))
            {
                
                string SourceStringFieldValue = Convert.ToString(currentEntity[SourceStringField]);
                using (OrganizationServiceContext orgSvcContext = new OrganizationServiceContext(service))
                {
                    var qry = from Lookup in orgSvcContext.CreateQuery(LookUpEntityName)
                              where Lookup[LookUpEntityField] == SourceStringFieldValue
                              select Lookup;


                    List<Entity> lstLookupEntities = qry.ToList<Entity>();
                    if (lstLookupEntities != null && lstLookupEntities.Count > 0)
                    {
                        Entity LookupEntity = lstLookupEntities.FirstOrDefault();
                        if (LookupEntity != null)
                        {

                            Entity e = currentEntity;
                            e[DestinationlookUpField] = new EntityReference(LookupEntity.LogicalName, LookupEntity.Id);
                            service.Update(e);
                        };
                    }
                }
            }

       
        }
    }
}
