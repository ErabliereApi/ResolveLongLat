using System.Net.Http.Headers;
using System.Text.Json;
using System.Web;
using ClosedXML.Excel;
using ErabliereAPI.Proxy;
using ResolveLongLat;

var path = @"{excel-data-path}";

var client = new HttpClient();

var httpClientEapi = new HttpClient();

httpClientEapi.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", 
"{your-bearer-token}");

var eapi = new ErabliereAPIProxy("https:/localhost:5001", httpClientEapi);

using (var workbook = new XLWorkbook(path))
{
    var ws = workbook.Worksheet(2);
    var rows = ws.RowsUsed().Skip(1);
    try
    {
        foreach (var row in rows.Where(r => r.Cell(4).IsEmpty() || r.Cell(5).IsEmpty()))
        {
            var nom = row.Cell(1).Value.GetText();
            var adresse = row.Cell(2).GetText();
            Console.Write("{0} ", nom);

            var urlBuilder = new UriBuilder();
            urlBuilder.Scheme = "https";
            urlBuilder.Host = "servicescarto.mern.gouv.qc.ca";
            urlBuilder.Path = "pes/rest/services/Territoire/AdressesQuebec_Geocodage/GeocodeServer/findAddressCandidates";
            var addressUrl = HttpUtility.UrlEncode(adresse);
            urlBuilder.Query = $"singleLine={addressUrl}&f=json&outSR=4326";

            // Call the AdressesQuebec API to get the coordinates
            var finalUrl = urlBuilder.ToString();

            HttpResponseMessage response;

            try
            {
                response = await client.GetAsync(finalUrl);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error: {0} Url: {1}", e.Message, finalUrl);
                throw;
            }

            var content = await response.Content.ReadAsStringAsync();

            if (response.IsSuccessStatusCode)
            {
                AdressesQuebec? json = null;

                try
                {
                    json = JsonSerializer.Deserialize<AdressesQuebec>(content);
                }
                catch (Exception e)
                {
                    Console.WriteLine("Error: {0} Content: {1} Url: {2}", e.Message, content, finalUrl);
                    throw;
                }

                var location = json?.candidates?.FirstOrDefault()?.location;
                if (location != null)
                {
                    Console.WriteLine("(long,lat): ({0},{1})", location.x, location.y);

                    row.Cell(4).SetValue(location.x);
                    row.Cell(5).SetValue(location.y);
                }
                else
                {
                    Console.WriteLine("No location found");

                    row.Cell(4).SetValue("N/A");
                    row.Cell(5).SetValue("N/A");
                }

                await Task.Delay(1000);
            }
            else
            {
                Console.WriteLine("Error: {0}", content);
            }
        }
    }
    catch (Exception e)
    {
        await Console.Error.WriteLineAsync(e.Message);
    }

    workbook.Save();

    // Pour chaque ligne avec une longitude et latitude
    // Trouvé l'érablière associée et ajouter les informations
    // Et sauvegarder l'ID de l'érablière dans la colonne 6
    var rowLonLat = ws.RowsUsed().Skip(1).Where(r => !r.Cell(4).IsEmpty() && !r.Cell(5).IsEmpty() && r.Cell(6).IsEmpty());

    int i = -1;

    try
    {
        foreach (var row in rowLonLat)
        {
            i++;

            var name = row.Cell(1).Value.GetText();
            var lonTxt = row.Cell(4).Value.ToString();
            var latTxt = row.Cell(5).Value.ToString();

            if (lonTxt == "N/A" || latTxt == "N/A")
            {
                Console.WriteLine("{0}: No longitude or latitude for {1}", i, name);
                continue;
            }

            var lon = double.Parse(lonTxt);
            var lat = double.Parse(latTxt);

            var filter = "nom eq '" + name.Replace("'", "''") + "'";

            var erabliere = await eapi.ErablieresAllAsync(
                select: null,
                filterQuery: filter,
                expand: null,
                orderbyQuery: null,
                topQuery: null,
                skip: null);

            if (erabliere != null && erabliere.Count == 1)
            {
                var erab = erabliere.First();
                Console.WriteLine("{0}: Erabliere {1} found with ID {2}", i, name, erab.Id);
                erab.Longitude = lon;
                erab.Latitude = lat;

                if (!erab.Id.HasValue)
                {
                    Console.WriteLine("{0}: Error: Erabliere {1} has no ID", i, name);
                    continue;
                }

                await eapi.ErablieresPUTAsync(erab.Id.Value, new PutErabliere
                {
                    Id = erab.Id.Value,
                    Longitude = lon,
                    Latitude = lat
                });

                Console.WriteLine("{0}: Erabliere {1} updated with longitude {2} and latitude {3}", i, name, lon, lat);

                row.Cell(6).SetValue(erab.Id.Value.ToString());
            }
            else
            {
                Console.WriteLine("{0}: Error: Erabliere {1} not found. Filter: {2}", i, name, filter);
            }
        }
    }
    catch (Exception e)
    {
        Console.Error.WriteLine("{0}: Error while posting data to ErabliereAPI: {1} {2}", i, e.Message, e.StackTrace);
    }

    workbook.Save();
}