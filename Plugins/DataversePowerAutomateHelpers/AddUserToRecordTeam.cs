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
    /// <summary>
    /// Provides logic implementation for the sample_AddUserToRecordTeam Custom API
    /// </summary>
    public class AddUserToRecordTeam : IPlugin
    {

        /*
        Input Parameters:
        | EntityReference | Target                  | Required | The ID of system user (user) to add to the auto created access team. |
        | String          | RecordEntityLogicalName | Required | Sets the record for which the access team is auto created.           |
        | Guid            | RecordId                | Required | Sets the record for which the access team is auto created.           |
        | String          | TeamTemplateName        | Required | The name of team template which is used to create the access team.   |

                    Team template is bound to a specific entity by integer entity type code value

        Output Parameters:
        | Guid | AccessTeamId | The ID of the auto created access team. |

        */


        public void Execute(IServiceProvider serviceProvider)
        {
            // Obtain the tracing service
            var tracingService = (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for web service calls.  
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            IOrganizationService service = serviceFactory.CreateOrganizationService(context.UserId);

            //Assign variables from context
            var target = (EntityReference)context.InputParameters["Target"];
            var recordEntityLogicalName = (string)context.InputParameters["RecordEntityLogicalName"];
            var recordId = (Guid)context.InputParameters["RecordId"];
            var teamTemplateName = (string)context.InputParameters["TeamTemplateName"];

            try
            {

                //Get the Object Type Code for the record
                var entityMetadata = Utility.GetEntityProperties(service, recordEntityLogicalName, "ObjectTypeCode");
                tracingService.Trace($"Object Type Code: {entityMetadata.ObjectTypeCode}");

                //Get the Id teamtemplate that matches the type of record
                var query = new QueryExpression("teamtemplate")
                {
                    ColumnSet = new ColumnSet("teamtemplateid"),
                    Criteria = new FilterExpression(LogicalOperator.And)
                    {
                        Conditions =
                        {
                            { new ConditionExpression("teamtemplatename", ConditionOperator.Equal, teamTemplateName)},
                            { new ConditionExpression("objecttypecode", ConditionOperator.Equal, entityMetadata.ObjectTypeCode)}
                        }
                    },
                    TopCount = 1
                };

                var results = service.RetrieveMultiple(query).Entities;
                tracingService.Trace($"Number of team templates records returned: {results.Count}");
                if (results.Count < 1)
                {
                    throw new InvalidPluginExecutionException($"No Team Template record found with the name " +
                        $"'{teamTemplateName}' for the {recordEntityLogicalName} entity.");
                }

                //Instantiate the request
                var request = new AddUserToRecordTeamRequest
                {
                    Record = new EntityReference(recordEntityLogicalName, recordId),
                    SystemUserId = target.Id,
                    TeamTemplateId = (Guid)results.FirstOrDefault()["teamtemplateid"]
                };

                //Send the request
                var response = ((AddUserToRecordTeamResponse)service.Execute(request));
                tracingService.Trace("AddUserToRecordTeamRequest sent.");

                //Set the output parameter
                context.OutputParameters["AccessTeamId"] = response.AccessTeamId;
                tracingService.Trace($"AccessTeamId value returned: {response.AccessTeamId}");

            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in sample_AddUserToRecordTeam: {ex.Message}", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("sample_AddUserToRecordTeam: {0}", ex.ToString());
                throw;
            }

        }
    }
}
