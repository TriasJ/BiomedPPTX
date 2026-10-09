using System;
using System.Collections.Generic;
using System.Net;
using System.Threading;

using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PowerPointLabs.SmartBrowserLab.Models
{
    public class PubChemClient
    {
        private const string BaseUrl = "https://pubchem.ncbi.nlm.nih.gov/rest/pug/compound";
        private const int RateLimitMs = 200;

        private DateTime _lastRequest;

        public PubChemClient()
        {
            _lastRequest = DateTime.MinValue;

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                | SecurityProtocolType.Tls11
                | SecurityProtocolType.Tls;
        }

        public List<int> SearchByName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return new List<int>();
            }

            EnforceRateLimit();

            try
            {
                string encoded = Uri.EscapeDataString(name);
                string url = string.Format("{0}/name/{1}/cids/JSON", BaseUrl, encoded);

                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                    string json = client.DownloadString(url);
                    _lastRequest = DateTime.Now;

                    JObject obj = JObject.Parse(json);
                    JToken listToken = obj.SelectToken("IdentifierList.CID");
                    if (listToken == null)
                    {
                        return new List<int>();
                    }

                    return listToken.ToObject<List<int>>();
                }
            }
            catch (WebException)
            {
                return new List<int>();
            }
        }

        public MoleculeSearchResult GetProperties(int cid)
        {
            EnforceRateLimit();

            string url = string.Format(
                "{0}/cid/{1}/property/MolecularFormula,MolecularWeight,IUPACName/JSON",
                BaseUrl,
                cid);

            using (var client = new WebClient())
            {
                client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                string json = client.DownloadString(url);
                _lastRequest = DateTime.Now;

                JObject obj = JObject.Parse(json);
                JToken propsToken = obj.SelectToken("PropertyTable.Properties[0]");
                if (propsToken == null)
                {
                    return null;
                }

                var result = new MoleculeSearchResult();
                result.Cid = propsToken.Value<int>("CID");
                result.MolecularFormula = propsToken.Value<string>("MolecularFormula") ?? "";
                result.MolecularWeight = propsToken.Value<double>("MolecularWeight");
                result.IupacName = propsToken.Value<string>("IUPACName") ?? "";
                return result;
            }
        }

        public byte[] Download2DImage(int cid)
        {
            return Download2DImage(cid, 0);
        }

        public byte[] Download2DImage(int cid, int size)
        {
            EnforceRateLimit();

            try
            {
                string url;
                if (size > 0)
                {
                    url = string.Format("{0}/cid/{1}/PNG?image_size={2}x{2}", BaseUrl, cid, size);
                }
                else
                {
                    url = string.Format("{0}/cid/{1}/PNG", BaseUrl, cid);
                }

                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                    byte[] data = client.DownloadData(url);
                    _lastRequest = DateTime.Now;
                    return data;
                }
            }
            catch (WebException)
            {
                return null;
            }
        }

        public PubChemCompoundData Get3DConformer(int cid)
        {
            EnforceRateLimit();

            try
            {
                string url = string.Format("{0}/cid/{1}/record/JSON/?record_type=3d", BaseUrl, cid);

                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                    string json = client.DownloadString(url);
                    _lastRequest = DateTime.Now;

                    PubChemResponse response = JsonConvert.DeserializeObject<PubChemResponse>(json);
                    if (response == null || response.Compounds == null || response.Compounds.Length == 0)
                    {
                        return null;
                    }

                    return response.Compounds[0];
                }
            }
            catch (WebException)
            {
                return null;
            }
        }

        public PubChemCompoundData Get2DStructure(int cid)
        {
            EnforceRateLimit();

            try
            {
                string url = string.Format("{0}/cid/{1}/record/JSON/?record_type=2d", BaseUrl, cid);

                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                    string json = client.DownloadString(url);
                    _lastRequest = DateTime.Now;

                    PubChemResponse response = JsonConvert.DeserializeObject<PubChemResponse>(json);
                    if (response == null || response.Compounds == null || response.Compounds.Length == 0)
                    {
                        return null;
                    }

                    return response.Compounds[0];
                }
            }
            catch (WebException)
            {
                return null;
            }
        }

        public string GetRawJson(int cid, bool is3D)
        {
            EnforceRateLimit();

            try
            {
                string recordType = is3D ? "3d" : "2d";
                string url = string.Format("{0}/cid/{1}/record/JSON/?record_type={2}", BaseUrl, cid, recordType);

                using (var client = new WebClient())
                {
                    client.Headers.Add("User-Agent", "BiomedPPTX/1.0");
                    string json = client.DownloadString(url);
                    _lastRequest = DateTime.Now;
                    return json;
                }
            }
            catch (WebException)
            {
                return null;
            }
        }

        private void EnforceRateLimit()
        {
            TimeSpan elapsed = DateTime.Now - _lastRequest;
            if (elapsed.TotalMilliseconds < RateLimitMs)
            {
                int waitMs = RateLimitMs - (int)elapsed.TotalMilliseconds;
                Thread.Sleep(waitMs);
            }
        }
    }
}
