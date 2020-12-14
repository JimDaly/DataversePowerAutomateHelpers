using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using System;
using System.Linq;
using System.ServiceModel;

namespace DataversePowerAutomateHelpers
{
    public class RetrieveOptions : IPlugin
    {
        /*
        Input Parameters:
        | String | EntityLogicalName    | Required | The LogicalName of the entity (or table) that contains the attribute (or column). |
        | String | LogicalName | Required | The LogicalName of the attribute (or column) that contains the options.           |


        Output Parameters:
        | EntityCollection | Options | A List of 'expando' entities where properties represent option properties|

        'Expando entities' are entities with no LogicalName set. Expando entities do not need to map to any entity metadata.
        
        When returned using Web API they have this @odata.type value:

            {
                "@odata.type": "#Microsoft.Dynamics.CRM.expando",
                "value": 1,
                "label": "Preferred Customer"
            }

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

            try
            {

                //Assign variables from context
                string entityLogicalName = (string)context.InputParameters["EntityLogicalName"];
                string logicalName = (string)context.InputParameters["LogicalName"];


                //Compose the request
                var req = new RetrieveAttributeRequest
                {
                    EntityLogicalName = entityLogicalName,
                    LogicalName = logicalName
                };

                tracingService.Trace($"Sending RetrieveAttributeRequest using values EntityLogicalName:{entityLogicalName}, LogicalName:{ logicalName}");
                //Send the request
                var response = (RetrieveAttributeResponse)service.Execute(req);
                tracingService.Trace("RetrieveAttributeRequest call succeeded.");

                //Define the EntityCollection to return
                var options = new EntityCollection();

                //Add expando entities to the collection based on the type of attribute.
                switch (response.AttributeMetadata)
                {
                    case BooleanAttributeMetadata b:

                        //True Option
                        options.Entities.Add(new Entity()
                        {
                            Attributes = {
                                { "value", b.OptionSet.TrueOption.Value},
                                { "label", b.OptionSet.TrueOption.Label.UserLocalizedLabel.Label}
                            }
                        });

                        //False Option
                        options.Entities.Add(new Entity()
                        {
                            Attributes = {
                                { "value", b.OptionSet.FalseOption.Value},
                                { "label", b.OptionSet.FalseOption.Label.UserLocalizedLabel.Label}
                            }
                        });

                        break;
                    case MultiSelectPicklistAttributeMetadata m:
                        m.OptionSet.Options.ToList().ForEach(o =>
                        {
                            var option = new Entity();
                            option["value"] = o.Value;
                            option["label"] = o.Label.UserLocalizedLabel.Label;
                            options.Entities.Add(option);
                        });
                        break;
                    case PicklistAttributeMetadata p:
                        p.OptionSet.Options.ToList().ForEach(o =>
                        {
                            var option = new Entity();
                            option["value"] = o.Value;
                            option["label"] = o.Label.UserLocalizedLabel.Label;
                            options.Entities.Add(option);
                        });
                        break;
                    case StateAttributeMetadata se:
                        se.OptionSet.Options.ToList().ForEach(o =>
                        {
                            var option = new Entity();
                            option["value"] = o.Value;
                            option["label"] = o.Label.UserLocalizedLabel.Label;
                            //Special properties that only this type of attribute has
                            option["defaultstatus"] = ((StateOptionMetadata)o).DefaultStatus;
                            option["invariantname"] = ((StateOptionMetadata)o).InvariantName;
                            options.Entities.Add(option);
                        });
                        break;
                    case StatusAttributeMetadata su:
                        su.OptionSet.Options.ToList().ForEach(o =>
                        {
                            var option = new Entity();
                            option["value"] = o.Value;
                            option["label"] = o.Label.UserLocalizedLabel.Label;
                            //Special property that only this type of attribute has
                            option["state"] = ((StatusOptionMetadata)o).State;
                            options.Entities.Add(option);
                        });
                        break;
                    default:
                        throw new InvalidPluginExecutionException($"The {logicalName} attribute doesn't have options.");

                }

                context.OutputParameters["Options"] = options;

                tracingService.Trace("sample_RetrieveOptions completed.");
            }
            catch (FaultException<OrganizationServiceFault> ex)
            {
                throw new InvalidPluginExecutionException($"An error occurred in sample_RetrieveOptions: {ex.Message}", ex);
            }

            catch (Exception ex)
            {
                tracingService.Trace("sample_RetrieveOptions: {0}", ex.ToString());
                throw;
            }
        }
    }
}
