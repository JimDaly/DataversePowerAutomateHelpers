using Microsoft.Crm.Sdk.Messages;
using Microsoft.Xrm.Sdk;
using System;
using System.ServiceModel;

namespace DataversePowerAutomateHelpers
{
    public class AddToQueue : IPlugin
    {
        public void Execute(IServiceProvider serviceProvider)
        {

            string sourceQueueName = null;
            Guid? sourceQueueId = null;
            string targetEntityLogicalName = null;
            Guid targetId = Guid.Empty;
            string destinationQueueName = null;
            Guid destinationQueueId;

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
            if (context.InputParameters.Contains("SourceQueueName"))
            {
                sourceQueueName = (string)context.InputParameters["SourceQueueName"];
            }

            if (context.InputParameters.Contains("DestinationQueueName"))
            {
                destinationQueueName = (string)context.InputParameters["DestinationQueueName"];
            }

            if (context.InputParameters.Contains("TargetEntityLogicalName"))
            {
                targetEntityLogicalName = (string)context.InputParameters["TargetEntityLogicalName"];
            }

            if (context.InputParameters.Contains("TargetId"))
            {
                targetId = (Guid)context.InputParameters["TargetId"];
            }

            try
            {
                /*
                INPUT:
                String SourceQueueName Optional
                String TargetEntityLogicalName Required
                Guid TargetId Required
                String DestinationQueueName Required

                OUTPUT:
                Guid QueueItemId 
                */

                if (!string.IsNullOrEmpty(destinationQueueName) && !string.IsNullOrEmpty(targetEntityLogicalName) && targetId != Guid.Empty)
                {
                    destinationQueueId = Utility.GetEntityIdByName(service, "queue", "name", "queueid", destinationQueueName);

                    var request = new AddToQueueRequest
                    {
                        DestinationQueueId = destinationQueueId,
                        Target = new EntityReference(targetEntityLogicalName, targetId)
                    };

                    if (!string.IsNullOrEmpty(sourceQueueName))
                    {
                        sourceQueueId = Utility.GetEntityIdByName(service, "queue", "name", "queueid", sourceQueueName);
                        if (sourceQueueId.HasValue)
                        {
                            request.SourceQueueId = sourceQueueId.Value;
                        }
                    }

                    var response = (AddToQueueResponse)service.Execute(request);
                    context.OutputParameters["QueueItemId"] = response.QueueItemId;
                }
                else
                {
                    tracingService.Trace("Expected AddToQueue parameter values: destinationQueueName = {destinationQueueName}, targetEntityLogicalName = {targetEntityLogicalName}, targetId = {targetId}");
                    throw new InvalidPluginExecutionException($"Required parameter values for AddToQueue not present.");                    
                }

            }

            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException("An error occurred in AddToQueue.", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("AddToQueue: {0}", ex.ToString());
                throw;
            }


        }

    }
}
