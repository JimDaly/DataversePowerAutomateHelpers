using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using Microsoft.Xrm.Sdk.Metadata.Query;
using Microsoft.Xrm.Sdk.Query;
using System;
using System.Linq;

namespace DataversePowerAutomateHelpers
{
    public static class Utility
    {
        /// <summary>
        /// Retrieves the id of the first instance of a record with the specified primary name value
        /// </summary>
        /// <param name="service">IOrganization service to us to make the request</param>
        /// <param name="entityLogicalName">The logical name of the entity</param>
        /// <param name="primaryNameAttributeLogicalName">The logical name of the primary name attribute</param>
        /// <param name="primaryKeyName">The name of the primary key</param>
        /// <param name="primaryNameValue">The primary name value to use as a condition.</param>
        /// <returns>The id of the matching record</returns>
        public static Guid GetEntityIdByName(IOrganizationService service, string entityLogicalName, string primaryNameAttributeLogicalName, string primaryKeyName, string primaryNameValue)
        {

            var query = new QueryExpression(entityLogicalName)
            {
                ColumnSet = new ColumnSet(primaryKeyName)
            };
            query.Criteria.AddCondition(new ConditionExpression(primaryNameAttributeLogicalName, ConditionOperator.Equal, primaryNameValue));
            query.TopCount = 1;
            var results = service.RetrieveMultiple(query).Entities;
            if (results.Count < 1)
            {
                throw new InvalidPluginExecutionException($"No {entityLogicalName} record found with the {primaryNameAttributeLogicalName} value '{primaryNameValue}'.");
            }
            return (Guid)results.FirstOrDefault()[primaryKeyName];
        }

        /// <summary>
        /// Whether an entity with the provided logical name exists.
        /// </summary>
        /// <param name="service">IOrganization service to us to make the request</param>
        /// <param name="entityLogicalName">The LogicalName value to test</param>
        public static void EntityExists(IOrganizationService service, string entityLogicalName) {

            //RetrieveMetadataChanges is more performant than RetrieveEntity
            var req = new RetrieveMetadataChangesRequest()
            {
                Query = new EntityQueryExpression()
                {
                    Criteria = new MetadataFilterExpression(LogicalOperator.And)
                    {
                        Conditions = {
                                {new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, entityLogicalName) }
                            }
                    },
                    Properties = new MetadataPropertiesExpression()
                    {
                        AllProperties = false,
                        PropertyNames = { "LogicalName" }
                    }
                }
            };

            //Send the request
            var response = (RetrieveMetadataChangesResponse)service.Execute(req);

            //Only one item should be returned if the name is valid
            if (response.EntityMetadata.Count.Equals(1))
            {
                return;
            }
            else
            {
                throw new InvalidPluginExecutionException($"'{entityLogicalName}' is not a valid entity LogicalName.");
            }

        }


        /// <summary>
        /// Returns specific Entity Properties
        /// </summary>
        /// <param name="service">IOrganization service to us to make the request.</param>
        /// <param name="entityLogicalName">The logical name of the entity.</param>
        /// <param name="entityPropertyNames">The properties of the entity.</param>
        /// <returns>EntityMetadata</returns>
        public static EntityMetadata GetEntityProperties(IOrganizationService service, string entityLogicalName, params string[] entityPropertyNames) {

            //RetrieveMetadataChanges is more performant than RetrieveEntity
            var req = new RetrieveMetadataChangesRequest()
            {
                Query = new EntityQueryExpression()
                {
                    Criteria = new MetadataFilterExpression(LogicalOperator.And)
                    {
                        Conditions = {
                                {new MetadataConditionExpression("LogicalName", MetadataConditionOperator.Equals, entityLogicalName) }
                            }
                    },
                    Properties = new MetadataPropertiesExpression(entityPropertyNames)
                }
            };

            try
            {
                //Send the request
                var response = (RetrieveMetadataChangesResponse)service.Execute(req);
                if (response.EntityMetadata.Count.Equals(1)) {
                    return response.EntityMetadata.FirstOrDefault();
                }
                else
                {
                    throw new InvalidPluginExecutionException($"'{entityLogicalName}' is not a valid entity logical name.");
                }
                
            }
            catch (Exception ex)
            {
                throw new InvalidPluginExecutionException(ex.Message, ex);
            }
            

        }
    }
}
