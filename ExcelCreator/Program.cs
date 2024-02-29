// Create a http call using this curl command
/*
 * curl --location --request POST 'https://uk.api.diligentcloudservices.com/entities/reports/appointments' \
--header 'Content-Type: application/json' \
--header 'authorization: Bearer <Token>' \
--header 'X-Api-Key: Bknbtl5IRt6YMsUyJsPhb3LtsYRaKk2D87TzeZo7' \
--header 'Cookie: cookie_bhs=rd40o00000000000000000000ffff0a37a2bbo80' \
--data-raw '{
    "Columns": [
    "UDFCompanies0.BPUI","INDIVIDUALS.NAME","OFFICERS.OFFPOS","LMOFFICERS.OFFPOSNAME.OFFICERS",
  "LMOFFICERCATS.OFFPOSCATNAME.OFFICERS","OFFICERS.EVENTDATE","OFFICERS.EVENTTYPE","LMEVENTS.EVENTNAME.OFFICERS",
  "OFFICERS.PERSONQR","INDIVIDUALS.TITLE","INDIVIDUALS.SURNAME","INDIVIDUALS.FORENAMES"
  ],
    "PageNumber": 1,
    "PageSize": 200,
    "FilterCombinator": 1,
    "Filters": [
    ]
}'
*/
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace ExcelCreator
{
    class Program
    {
        private static string[] columns = [
            "INDIVIDUALS.NAME","OFFICERS.OFFPOS","LMOFFICERS.OFFPOSNAME.OFFICERS",
            "LMOFFICERCATS.OFFPOSCATNAME.OFFICERS","OFFICERS.EVENTDATE","OFFICERS.EVENTTYPE","LMEVENTS.EVENTNAME.OFFICERS",
            "OFFICERS.PERSONQR","INDIVIDUALS.TITLE","INDIVIDUALS.SURNAME","INDIVIDUALS.FORENAMES"
        ];

        static async Task Main(string[] args)
        {
            await CreateReport();
        }

        static async Task CreateReport(string XxApiKey, string bearerToken)
        {
            var client = new HttpClient();
            client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", "eyJhbGciOiJSUzI1NiIsImtpZCI6IjgzNEUwMDUyOTI3MzYwRTNDNDRDNEI2M0RBNzIwQTlGNjEzRTE1ODkiLCJ4NXQiOiJnMDRBVXBKellPUEVURXRqMm5JS24yRS1GWWsiLCJ0eXAiOiJhdCtqd3QifQ.eyJpc3MiOiJodHRwczovL2lkcy1xYTAxLmJsdWVwcmludHNlcnZlci5jb20vZGVxYS0wMzEiLCJuYmYiOjE3MDkxOTkxNzQsImlhdCI6MTcwOTE5OTE3NCwiZXhwIjoxNzExNzkxMTc0LCJhdWQiOiJuZXh0Z2VuLmFwaSIsInNjb3BlIjpbInB1YmxpYy5hcGkucmVhZCIsIm5leHRnZW4uYXBpIl0sImFtciI6WyJwd2QiXSwiY2xpZW50X2lkIjoibmV4dGdlbi5hcGkuc3dhZ2dlciIsInN1YiI6IkdCT1ItRiIsImF1dGhfdGltZSI6MTcwOTE5ODYzMCwiaWRwIjoibG9jYWwiLCJ1c2VyaWQiOiJHQk9SLUYiLCJzaWQiOiI5NEQwRDFERkMyMjVDMjA3MThBRTkxNTBFQkNGQUEzNSJ9.EW6pz3hseC-TfxX16yFK6pq-SgiB3Z6OXFLkLQ3aSOqLRGFuAy8ck4sSXCgrwinQlFuDpYCR6jzEO1ucAwQqFa2UAoZfsIO4BKpuQRJ1OxGeByXJJ864MLY1sjrZVsSqghxOjRQBT_UnDivFUGb1U-JOa6-uj2hvavLg1QQwBLNvSWBZ0w9Q2A2POKm9wiDxEGbkn9TN4kEWqp0gl8XPV4hx1MPojN8_dVPHoJ8-fLqRuAhBj-_R6PxLS3rhtAuqgDQuQx-Zwr8vmvt5huhsmJprQUzsB3A2eUUNM6Sw63HZ9qFl9b__3u2ETwUjglzSkqfxYYey2rbphgUjW5eqUg");


            var request = new HttpRequestMessage(HttpMethod.Post, "https://deqa-031.blueprintserver.com/nextgenapi/v1/reports/appointments");
            request.Content = new StringContent(
                               JsonSerializer.Serialize(new
                               {
                                   Columns = columns,
                                   PageNumber = 1,
                                   PageSize = 200,
                                   FilterCombinator = 1,
                                   Filters = new List<object>()
                               }),
                                              Encoding.UTF8,
                                                             "application/json"
                                                                        );

            var response = await client.SendAsync(request);
            response.EnsureSuccessStatusCode();

            var responseStream = await response.Content.ReadAsStreamAsync();
            var report = await JsonSerializer.DeserializeAsync<Root>(responseStream);
            Console.WriteLine(report);

            //create csv from report.Rows.Fields
            var csv = new StringBuilder();
            csv.AppendLine(string.Join(";", report.ColumnMetadata.GetType().GetProperties().Select(p => ((dynamic?)p.GetValue(report.ColumnMetadata))?.Key)));
            csv.AppendLine(string.Join(";", report.ColumnMetadata.GetType().GetProperties().Select(p => ((dynamic?)p.GetValue(report.ColumnMetadata))?.Name)));
            foreach (var row in report.Rows)
            {
                csv.AppendLine(string.Join(";", row.Fields.Select(f => f?.ToString())));
            }

            var pageCount = report.Count / 200;
            for (int i = 2; i <= pageCount + 1; i++)
            {
                request = new HttpRequestMessage(HttpMethod.Post, "https://deqa-031.blueprintserver.com/nextgenapi/v1/reports/appointments");
                request.Content = new StringContent(
                                                  JsonSerializer.Serialize(new
                                                  {
                                   Columns = columns,
                                   PageNumber = i,
                                   PageSize = 200,
                                   FilterCombinator = 1,
                                   Filters = new List<object>()
                               }),
                                                                                               Encoding.UTF8,
                                                                                                                                                           "application/json"
                                                                                                                                                                                                                                  );

                response = await client.SendAsync(request);
                response.EnsureSuccessStatusCode();

                responseStream = await response.Content.ReadAsStreamAsync();
                report = await JsonSerializer.DeserializeAsync<Root>(responseStream);
                Console.WriteLine(report);

                foreach (var row in report.Rows)
                {
                    csv.AppendLine(string.Join(";", row.Fields.Select(f => f?.ToString())));
                }
            }

            // write to a file
            File.WriteAllText("report.csv", csv.ToString());
        }
    }

    // Root myDeserializedClass = JsonSerializer.Deserialize<Root>(myJsonResponse);
    public class ColumnMetadata
    {
        [JsonPropertyName("NAME")]
        public NAME NAME { get; set; }

        [JsonPropertyName("OFFPOS")]
        public OFFPOS OFFPOS { get; set; }

        [JsonPropertyName("OFFPOSNAME")]
        public OFFPOSNAME OFFPOSNAME { get; set; }

        [JsonPropertyName("OFFPOSCATNAME")]
        public OFFPOSCATNAME OFFPOSCATNAME { get; set; }

        [JsonPropertyName("EVENTDATE")]
        public EVENTDATE EVENTDATE { get; set; }

        [JsonPropertyName("EVENTTYPE")]
        public EVENTTYPE EVENTTYPE { get; set; }

        [JsonPropertyName("EVENTNAME")]
        public EVENTNAME EVENTNAME { get; set; }

        [JsonPropertyName("PERSONQR")]
        public PERSONQR PERSONQR { get; set; }

        [JsonPropertyName("TITLE")]
        public TITLE TITLE { get; set; }

        [JsonPropertyName("SURNAME")]
        public SURNAME SURNAME { get; set; }

        [JsonPropertyName("FORENAMES")]
        public FORENAMES FORENAMES { get; set; }

        [JsonPropertyName("EVENTID")]
        public EVENTID EVENTID { get; set; }

        [JsonPropertyName("COMPQR")]
        public COMPQR COMPQR { get; set; }

        [JsonPropertyName("NAME1")]
        public NAME1 NAME1 { get; set; }

        [JsonPropertyName("NAME2")]
        public NAME2 NAME2 { get; set; }
    }

    public class COMPQR
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class EVENTDATE
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class EVENTID
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class EVENTNAME
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class EVENTTYPE
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class FORENAMES
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class NAME
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class NAME1
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class NAME2
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class OFFPOS
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class OFFPOSCATNAME
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class OFFPOSNAME
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class PERSONQR
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class Root
    {
        [JsonPropertyName("rows")]
        public List<Row> Rows { get; set; }

        [JsonPropertyName("count")]
        public int Count { get; set; }

        [JsonPropertyName("reportHeaders")]
        public List<string> ReportHeaders { get; set; }

        [JsonPropertyName("columnMetadata")]
        public ColumnMetadata ColumnMetadata { get; set; }
    }

    public class Row
    {
        [JsonPropertyName("fields")]
        public List<object> Fields { get; set; }
    }

    public class SURNAME
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }

    public class TITLE
    {
        [JsonPropertyName("key")]
        public string Key { get; set; }

        [JsonPropertyName("section")]
        public string Section { get; set; }

        [JsonPropertyName("name")]
        public string Name { get; set; }

        [JsonPropertyName("description")]
        public string Description { get; set; }

        [JsonPropertyName("type")]
        public string Type { get; set; }

        [JsonPropertyName("requiredAccessRight")]
        public List<string> RequiredAccessRight { get; set; }
    }
}
