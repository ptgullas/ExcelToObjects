using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;
using Serilog;

namespace ExcelToObjects {
    public class MemberProcessor {
        private ZipCodeRetrieverService _zipRetrieverService;

        public static Dictionary<string, string> stateToAbbrev = new Dictionary<string, string>() { { "alabama", "AL" }, { "alaska", "AK" }, { "arizona", "AZ" }, { "arkansas", "AR" }, { "california", "CA" }, { "colorado", "CO" }, { "connecticut", "CT" }, { "delaware", "DE" }, { "district of columbia", "DC" }, { "florida", "FL" }, { "georgia", "GA" }, { "hawaii", "HI" }, { "idaho", "ID" }, { "illinois", "IL" }, { "indiana", "IN" }, { "iowa", "IA" }, { "kansas", "KS" }, { "kentucky", "KY" }, { "louisiana", "LA" }, { "maine", "ME" }, { "maryland", "MD" }, { "massachusetts", "MA" }, { "michigan", "MI" }, { "minnesota", "MN" }, { "mississippi", "MS" }, { "missouri", "MO" }, { "montana", "MT" }, { "nebraska", "NE" }, { "nevada", "NV" }, { "new hampshire", "NH" }, { "new jersey", "NJ" }, { "new mexico", "NM" }, { "new york", "NY" }, { "north carolina", "NC" }, { "north dakota", "ND" }, { "ohio", "OH" }, { "oklahoma", "OK" }, { "oregon", "OR" }, { "pennsylvania", "PA" }, { "rhode island", "RI" }, { "south carolina", "SC" }, { "south dakota", "SD" }, { "tennessee", "TN" }, { "texas", "TX" }, { "utah", "UT" }, { "vermont", "VT" }, { "virginia", "VA" }, { "washington", "WA" }, { "west virginia", "WV" }, { "wisconsin", "WI" }, { "wyoming", "WY" } };
        public static Dictionary<string, string> abbrevToState = new Dictionary<string, string>() { { "AK", "alaska" }, { "AL", "alabama" }, { "AR", "arkansas" }, { "AZ", "arizona" }, { "CA", "california" }, { "CO", "colorado" }, { "CT", "connecticut" }, { "DC", "district of columbia" }, { "DE", "delaware" }, { "FL", "florida" }, { "GA", "georgia" }, { "HI", "hawaii" }, { "IA", "iowa" }, { "ID", "idaho" }, { "IL", "illinois" }, { "IN", "indiana" }, { "KS", "kansas" }, { "KY", "kentucky" }, { "LA", "louisiana" }, { "MA", "massachusetts" }, { "MD", "maryland" }, { "ME", "maine" }, { "MI", "michigan" }, { "MN", "minnesota" }, { "MO", "missouri" }, { "MS", "mississippi" }, { "MT", "montana" }, { "NC", "north carolina" }, { "ND", "north dakota" }, { "NE", "nebraska" }, { "NH", "new hampshire" }, { "NJ", "new jersey" }, { "NM", "new mexico" }, { "NV", "nevada" }, { "NY", "new york" }, { "OH", "ohio" }, { "OK", "oklahoma" }, { "OR", "oregon" }, { "PA", "pennsylvania" }, { "RI", "rhode island" }, { "SC", "south carolina" }, { "SD", "south dakota" }, { "TN", "tennessee" }, { "TX", "texas" }, { "UT", "utah" }, { "VA", "virginia" }, { "VT", "vermont" }, { "WA", "washington" }, { "WI", "wisconsin" }, { "WV", "west virginia" }, { "WY", "wyoming" } };

        public MemberProcessor(ZipCodeRetrieverService zipCodeRetrieverService) {
            _zipRetrieverService = zipCodeRetrieverService;
        }

        public async Task<string> GetZipFromMemberAddress(Member member) {
            string zip = member.ZipCode;
            if (member.ContainsFullAddressAndNoZip()) {
                string fullMemberAddress = member.GetFullAddress();
                zip = await _zipRetrieverService.GetZip(fullMemberAddress);
            }
            return zip;
        }

        public string GetStateAbbreviation(Member member, int index) {
            if (string.IsNullOrEmpty(member.State)) {
                return null;
            }
            else {
                return ConvertStateToAbbreviation(member.State, index);
            }
        }

        private string ConvertStateToAbbreviation(string stateName, int index) {
            string abbreviation = null;
            if (string.IsNullOrEmpty(stateName) || (stateName.Length == 2)) {
                abbreviation = stateName;
            }
            else if (stateToAbbrev.ContainsKey(stateName.ToLower())) {
                abbreviation = stateToAbbrev[stateName.ToLower()];
            }
            else {
                Log.Warning("Row {row} contains invalid US State: {stateName}", index, stateName);
            }
            return abbreviation;
        }

    }
}
