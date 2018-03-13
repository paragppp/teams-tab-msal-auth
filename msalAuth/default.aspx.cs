using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Net.Http;
using Newtonsoft.Json.Linq;

namespace msalAuth
{
    public partial class _default : System.Web.UI.Page
    {
        string token = "";

        protected void Page_Load(object sender, EventArgs e)
        {
            if (Request.QueryString.Count > 0)
            {
                if (Request.QueryString["tk"] != "")
                {
                    token = Request.QueryString["tk"];

                    using (var client = new HttpClient())
                    {
                        var url = "https://graph.microsoft.com/v1.0/me/";
                        client.DefaultRequestHeaders.Add("Authorization", "Bearer " + token);
                        var response = client.GetAsync(url).Result;

                        var responseContent = response.Content;

                        string responseString = responseContent.ReadAsStringAsync().Result;

                        if (response.IsSuccessStatusCode)
                        {
                            JObject jObject = JObject.Parse(responseString);

                            Label1.Text = "Welcome " + jObject["displayName"].ToString();
                        }
                        else
                            Label1.Text = responseString + "[" + token + "]";
                    }

                }
            }
        }
    }
}