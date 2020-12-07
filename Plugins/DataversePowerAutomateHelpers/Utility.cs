using Microsoft.Xrm.Sdk;
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
                throw new InvalidPluginExecutionException($"No {entityLogicalName} record found with the {primaryNameAttributeLogicalName} value {primaryNameValue}.");
            }
            return (Guid)results.FirstOrDefault()[primaryKeyName];
        }
    }
}
