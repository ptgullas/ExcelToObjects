using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToObjects {
    public class MemberProcessor {
        private ZipCodeRetrieverService _zipRetrieverService;
        public MemberProcessor(ZipCodeRetrieverService zipCodeRetrieverService) {
            _zipRetrieverService = zipCodeRetrieverService;
        }

        public async Task<string> GetZipFromMemberAddress(Member member) {
            string zip = null;
            if (member.ContainsFullAddressAndNoZip()) {
                string fullMemberAddress = member.GetFullAddress();
                zip = await _zipRetrieverService.GetZip(fullMemberAddress);
            }
            return zip;
        }

    }
}
