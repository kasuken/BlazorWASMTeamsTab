using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace BlazorTeamsTabExample.Models
{
    public class Group
    {
        [JsonProperty("mail")]
        public string Mail { get; set; }

        [JsonProperty("onPremisesSyncEnabled")]
        public bool? OnPremisesSyncEnabled { get; set; }

        [JsonProperty("description")]
        public string Description { get; set; }

        [JsonProperty("createdDateTime")]
        public string CreatedDateTime { get; set; }

        [JsonProperty("classification")]
        public object Classification { get; set; }

        [JsonProperty("deletedDateTime")]
        public object DeletedDateTime { get; set; }

        [JsonProperty("groupTypes")]
        public string[] GroupTypes { get; set; }

        [JsonProperty("displayName")]
        public string DisplayName { get; set; }

        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("membershipRuleProcessingState")]
        public string MembershipRuleProcessingState { get; set; }

        [JsonProperty("mailNickname")]
        public string MailNickname { get; set; }

        [JsonProperty("mailEnabled")]
        public bool MailEnabled { get; set; }

        [JsonProperty("membershipRule")]
        public object MembershipRule { get; set; }

        [JsonProperty("onPremisesProvisioningErrors")]
        public object[] OnPremisesProvisioningErrors { get; set; }

        [JsonProperty("onPremisesLastSyncDateTime")]
        public string OnPremisesLastSyncDateTime { get; set; }

        [JsonProperty("onPremisesSecurityIdentifier")]
        public string OnPremisesSecurityIdentifier { get; set; }

        [JsonProperty("resourceBehaviorOptions")]
        public string[] ResourceBehaviorOptions { get; set; }

        [JsonProperty("proxyAddresses")]
        public string[] ProxyAddresses { get; set; }

        [JsonProperty("preferredLanguage")]
        public object PreferredLanguage { get; set; }

        [JsonProperty("renewedDateTime")]
        public string RenewedDateTime { get; set; }

        [JsonProperty("securityEnabled")]
        public bool SecurityEnabled { get; set; }

        [JsonProperty("resourceProvisioningOptions")]
        public object[] ResourceProvisioningOptions { get; set; }

        [JsonProperty("theme")]
        public object Theme { get; set; }

        [JsonProperty("visibility")]
        public string Visibility { get; set; }
    }
}
