using Newtonsoft.Json.Linq;
using System.Net;

namespace innovation_tracker_backend.Helper
{
    public class EncodeData
    {
        public static string[] HtmlEncodeObject(JObject data)
        {
            var encodedData = new List<string>();

            foreach (var property in data.Properties())
            {
                encodedData.Add(WebUtility.HtmlEncode(property.Value.ToString()));
            }

            return encodedData.ToArray();
        }
    }
}
