using System.IO.Compression;
using System.Net;
using System.Xml;
using Newtonsoft.Json.Linq;
using static System.Net.WebRequestMethods;

public class Script : ScriptBase
{
    const string __NS_URI = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// Resolve potential issue with base64 encoding of the OperationId. Test and decode if it's base64 encoded
    /// </summary>
    private string OperationId 
    {
        get
        {
            var value = this.Context.OperationId;

            try
            {
                byte[] data = Convert.FromBase64String(value);
                value = System.Text.Encoding.UTF8.GetString(data);
            }
            catch (FormatException ex) { }

            return value;
        }
    }

    public override async Task<HttpResponseMessage> ExecuteAsync()
    {
        // Check if the operation ID matches what is specified in the OpenAPI definition of the connector

        if (this.OperationId == "TrackChanges")
        {
            return await this.HandleTrackChangesOperation().ConfigureAwait(false);
        }


        // Handle an invalid operation ID

        HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.BadRequest);
        response.Content = CreateJsonContent($"Unknown operation ID '{this.Context.OperationId}'");
        return response;
    }

    private async Task<HttpResponseMessage> HandleTrackChangesOperation()
    {
        HttpResponseMessage response;

        if (this.Context.Request.Content == null)
        {
            return new HttpResponseMessage(HttpStatusCode.BadRequest)
            {
                Content = CreateJsonContent("Value cannot be null.")
            };
        }

        // We assume the body of the incoming request looks like this:
        // {
        //   "Content": "<base64>"
        // }

        var contentAsString = await this.Context.Request.Content.ReadAsStringAsync().ConfigureAwait(false);


        // Parse as JSON object
        
        var contentAsJson = JObject.Parse(contentAsString);


        // Get the file content
        
        var content = Convert.FromBase64String((string)contentAsJson["Content"]);


        // Memory stream

        using (var stream = new System.IO.MemoryStream())
        {
            // Copy file to memory stream and reset position

            stream.Write(content, 0, content.Length);
            stream.Position = 0;
            
            
            // Read the file as a zip

            using (var zip = new ZipArchive(stream, ZipArchiveMode.Update))
            {
                // Open word/settings.xml

                var entry = zip.GetEntry("word/settings.xml");
                if (entry != null)
                {
                    // Load word/settings.xml to an XmlDocument

                    var doc = new XmlDocument();
                    using (var settingsStream = entry.Open())
                    {
                        doc.Load(settingsStream);
                    }
                    
                
                    // Delete word/settings.xml from the zip - Will recreate it later.

                    entry.Delete();


                    // Create a namespace manager

                    var nsMgr = new XmlNamespaceManager(doc.NameTable);
                    nsMgr.AddNamespace("w", __NS_URI);


                    // Append /w:settings/w:trackRevisions if it doesn't already exist

                    if (doc.SelectSingleNode("/w:settings/w:trackRevisions", nsMgr) == null)
                    {
                        doc.DocumentElement!.AppendChild(doc.CreateElement("w:trackRevisions", __NS_URI));
                    }


                    // Recreate word/settings.xml

                    entry = zip.CreateEntry("word/settings.xml");
                    using (var settingsStream = entry.Open())
                    {
                        doc.Save(settingsStream);
                    }
                }
            }

            // Write the changes to the zip in memory.

            stream.Flush();


            // Return a JSON object containing the file content as a base64 string

            JObject output = new JObject
            {
                ["Content"] = Convert.ToBase64String(stream.ToArray())
            };

            response = new HttpResponseMessage(HttpStatusCode.OK);
            response.Content = CreateJsonContent(output.ToString());
            return response;
        }
    }
}