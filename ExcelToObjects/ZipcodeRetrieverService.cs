using System;
using System.Collections.Generic;
using System.Text;
using System.Net.Http;
using System.Net.Http.Headers;
using Google.Apis.CivicInfo.v2;
using Google.Apis.Services;
using System.Threading.Tasks;
using Serilog;

namespace ExcelToObjects {
    public class ZipCodeRetrieverService {
        private readonly string _apiKey;
        public ZipCodeRetrieverService(string GoogleApiKey) {
            _apiKey = GoogleApiKey;
        }

        public async Task<string> GetZip(string streetAddress, long electionId = 2000) {
            string zip = null;
            var voterQueryRequest = SetUpVoterQueryRequest(streetAddress, electionId);
            Log.Information("Looking up zip code info for {streetAddress} from Google", streetAddress);
            try {
                var result = await voterQueryRequest.ExecuteAsync();
                if (result != null) {
                    zip = result.NormalizedInput.Zip;
                    Log.Information("Success! Found zip code: {zip}", zip);
                }
                else {
                    Log.Information("Failure. Google Civic API returned null");
                }
            }
            catch (Exception e) {
                Log.Error(e, "Couldn't retrieve VoterInfo from Google Civic API");
            }
            return zip;
        }

        private ElectionsResource.VoterInfoQueryRequest SetUpVoterQueryRequest(string streetAddress, long electionId) {
            CivicInfoService service = new CivicInfoService(new BaseClientService.Initializer {
                ApplicationName = "ExcelToObjects",
                ApiKey = _apiKey
            });
            var voterQueryRequest = new ElectionsResource.VoterInfoQueryRequest(service, streetAddress);
            voterQueryRequest.ElectionId = electionId;
            return voterQueryRequest;
        }
    }
}
