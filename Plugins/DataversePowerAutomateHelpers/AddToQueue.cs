using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.ServiceModel;

namespace DataversePowerAutomateHelpers
{
    /// <summary>
    /// Provides logic implementation for the sample_AddToQueue Custom API
    /// </summary>
    public class AddToQueue : IPlugin
    {
        /*
        Input Parameters:
        | String | SourceQueueName         | Optional | The name of the queue that the item should be moved from.                    |
        | String | TargetEntityLogicalName | Required | The logical name of the entity that represents the item to add to the queue. |
        | Guid   | TargetId                | Required | The Id of the item to add to the queue.                                      |
        | String | DestinationQueueName    | Required | The name of the queue to add the item to.                                    |

        Output Parameters:
        | Guid | QueueItemId |
        */

        public void Execute(IServiceProvider serviceProvider)
        {

            // Obtain the tracing service
            var tracingService =
            (ITracingService)serviceProvider.GetService(typeof(ITracingService));

            // Obtain the execution context from the service provider.  
            var context = (IPluginExecutionContext)serviceProvider.GetService(typeof(IPluginExecutionContext));

            // Obtain the organization service reference which you will need for  web service calls.  
            var serviceFactory = (IOrganizationServiceFactory)serviceProvider.GetService(typeof(IOrganizationServiceFactory));
            var service = serviceFactory.CreateOrganizationService(context.UserId);

            //Assign variables from context
            string sourceQueueName = (string)context.InputParameters["SourceQueueName"];
            string destinationQueueName = (string)context.InputParameters["DestinationQueueName"];
            string targetEntityLogicalName = (string)context.InputParameters["TargetEntityLogicalName"];
            Guid targetId = (Guid)context.InputParameters["TargetId"];


            try
            {

                //Verify required parameters have values
                if (!string.IsNullOrEmpty(destinationQueueName) &&
                    !string.IsNullOrEmpty(targetEntityLogicalName) &&
                    targetId != Guid.Empty)
                {

                    //Validate target entity logical name exists
                    try
                    {
                        Utility.EntityExists(service, targetEntityLogicalName);
                        tracingService.Trace($"TargetEntityLogicalName is: '{targetEntityLogicalName}'");
                    }
                    catch (Exception)
                    {
                        tracingService.Trace($"'{targetEntityLogicalName}' is not a valid entity logical name.");
                        throw;
                    }


                    //Get the ID of the destination queue by name
                    Guid destinationQueueId;

                    try
                    {
                        destinationQueueId = Utility.GetEntityIdByName(service, "queue", "name", "queueid", destinationQueueName);
                        tracingService.Trace($"DestinationQueueId is: '{destinationQueueId}'");
                    }
                    catch (Exception)
                    {
                        tracingService.Trace($"There is no destination queue named '{destinationQueueName}'");
                        throw;
                    }

                    //Instantiate the request
                    var request = new AddToQueueRequest
                    {
                        DestinationQueueId = destinationQueueId,
                        Target = new EntityReference(targetEntityLogicalName, targetId)
                    };

                    //If optional SourceQueueName parameter value included, retrieve it and add it to the request
                    if (!string.IsNullOrEmpty(sourceQueueName))
                    {
                        tracingService.Trace($"SourceQueueName is: '{sourceQueueName}'.");

                        try
                        {
                            request.SourceQueueId = Utility.GetEntityIdByName(service, "queue", "name", "queueid", sourceQueueName);
                            tracingService.Trace($"SourceQueueId is: '{request.SourceQueueId}'");
                        }
                        catch (Exception)
                        {
                            tracingService.Trace($"There is no source queue named '{sourceQueueName}'");
                            throw;
                        }                        
                    }


                    //Send the request
                    var response = (AddToQueueResponse)service.Execute(request);
                    tracingService.Trace("AddToQueueRequest sent and recieved.");

                    //Set the output parameter
                    context.OutputParameters["QueueItemId"] = response.QueueItemId;
                    tracingService.Trace($"QueueItemId value returned: {response.QueueItemId}");

                }
                else
                {
                    //Trace the expected input parameters
                    tracingService.Trace($"Expected AddToQueue parameter values: " +
                        $"destinationQueueName = {destinationQueueName}, " +
                        $"targetEntityLogicalName = {targetEntityLogicalName}, " +
                        $"targetId = {targetId}");

                    //Throw error when require parameters not set
                    throw new InvalidPluginExecutionException($"Required parameter values for sample_AddToQueue not present.");
                }

            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in sample_AddToQueue: {ex.Message}", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("sample_AddToQueue: {0}", ex.ToString());
                throw;
            }
        }
    }
}
