using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;
using System.ServiceModel;

namespace DataversePowerAutomateHelpers
{
    public class AddUserToRecordTeam : IPlugin
    {

        /*
        INPUT:
        Entity Target Required Required  the ID of system user (user) to add to the auto created access team. Required.
        String RecordEntityLogicalName Required  Gets or sets the record for which the access team is auto created. Required.
        Guid RecordId Required  Gets or sets the record for which the access team is auto created. Required.
        String TeamTemplateName Required  The name of team template which is used to create the access team. Required.
         -- Team template is bound to a specific entity by ETC

        OUTPUT:
        Guid AccessTeamId 
        */


        public void Execute(IServiceProvider serviceProvider)
        {
            EntityReference target = null;
            string recordEntityLogicalName = null;
            Guid recordId = Guid.Empty;
            string teamTemplateName = null;
            Guid teamTemplateId = Guid.Empty;


            // Obtain the tracing service
            ITracingService tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            IPluginExecutionContext context = (IPluginExecutionContext)
                serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for  
            // web service calls.  
            IOrganizationServiceFactory serviceFactory =
                (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            //Assign variables from context
            if (context.InputParameters.Contains("Target"))
            {
                target = (EntityReference)context.InputParameters["Target"];
            }

            if (context.InputParameters.Contains("RecordEntityLogicalName"))
            {
                recordEntityLogicalName = (string)context.InputParameters["RecordEntityLogicalName"];
            }

            if (context.InputParameters.Contains("RecordId"))
            {
                recordId = (Guid)context.InputParameters["RecordId"];
            }

            if (context.InputParameters.Contains("TeamTemplateName"))
            {
                teamTemplateName = (string)context.InputParameters["TeamTemplateName"];
            }

            try
            {
                //Get the team template Id & verify it is of the right type

                //Get the metadata for the record
                var retrieveEntityRequest = new RetrieveEntityRequest
                {
                    EntityFilters = EntityFilters.Entity,
                    LogicalName = recordEntityLogicalName
                };

                EntityMetadata entityMetadata = ((RetrieveEntityResponse)service.Execute(retrieveEntityRequest)).EntityMetadata;

                //Get the Id teamtemplate that matches the type of record
                var query = new QueryExpression("teamtemplate")
                {
                    ColumnSet = new ColumnSet("teamtemplateid")
                };
                query.Criteria.AddCondition(new ConditionExpression("teamtemplatename", ConditionOperator.Equal, teamTemplateName));
                query.Criteria.AddCondition(new ConditionExpression("objecttypecode", ConditionOperator.Equal, entityMetadata.ObjectTypeCode));
                query.TopCount = 1;
                var results = service.RetrieveMultiple(query).Entities;
                if (results.Count < 1)
                {
                    throw new InvalidPluginExecutionException($"No Team Template record found with the name '{teamTemplateName}' for the {recordEntityLogicalName} entity.");
                }
                teamTemplateId = (Guid)results.FirstOrDefault()["teamtemplateid"];

                var request = new AddUserToRecordTeamRequest
                {
                    Record = new EntityReference(recordEntityLogicalName, recordId),
                    SystemUserId = target.Id,
                    TeamTemplateId = teamTemplateId
                };

                context.OutputParameters["AccessTeamId"] = ((AddUserToRecordTeamResponse)service.Execute(request)).AccessTeamId;

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in AddUserToRecordTeam: {ex.Message}", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("AddUserToRecordTeam: {0}", ex.ToString());
                throw;
            }

        }
    }
}
