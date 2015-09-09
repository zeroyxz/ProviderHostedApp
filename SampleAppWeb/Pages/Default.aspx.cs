using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using Microsoft.SharePoint.Client;
using Microsoft.IdentityModel.S2S.Tokens;
using System.Net;
using System.IO;
using System.Xml;


namespace SampleAppWeb
{
    public partial class Default : System.Web.UI.Page
    {
        SharePointContextToken contextToken;
        string accessToken;
        Uri sharePointUrl;
        string siteName;
        string currentUser;
        List<string> listOfUsers = new List<string>();
        List<string> listOfLists = new List<string>();

        protected void Page_PreInit(object sender, EventArgs e)
        {
            Uri redirectUrl;
            switch (SharePointContextProvider.CheckRedirectionStatus(Context, out redirectUrl))
            {
                case RedirectionStatus.Ok:
                    return;
                case RedirectionStatus.ShouldRedirect:
                    Response.Redirect(redirectUrl.AbsoluteUri, endResponse: true);
                    break;
                case RedirectionStatus.CanNotRedirect:
                    Response.Write("An error occurred while processing your request.");
                    Response.End();
                    break;
            }
        }

        protected void Page_Load(object sender, EventArgs e)
        {
            string contextTokenString = TokenHelper.GetContextTokenFromRequest(Request);

            if (contextTokenString != null)
            {
                contextToken = TokenHelper.ReadAndValidateContextToken(contextTokenString, Request.Url.Authority);
                sharePointUrl = new Uri(Request.QueryString["SPHostUrl"]);
                accessToken = TokenHelper.GetAccessToken(contextToken, sharePointUrl.Authority).AccessToken;
            }
            else if (!IsPostBack)
            {
                Response.Write("Couldn't find a context token");
                return;
            }


        }

        private void RetrieveWithCSOM(string accessToken)
        {
            if (IsPostBack)
                sharePointUrl = new Uri(Request.QueryString["SPHostUrl"]);

            ClientContext context = TokenHelper.GetClientContextWithAccessToken(sharePointUrl.ToString(), accessToken);

            Web web = context.Web;
            context.Load(web);
            context.ExecuteQuery();

            siteName = web.Title;

            context.Load(web.CurrentUser);
            context.ExecuteQuery();
            currentUser = context.Web.CurrentUser.LoginName;

            ListCollection lists = web.Lists;
            context.Load<ListCollection>(lists);
            context.ExecuteQuery();

            UserCollection users = web.SiteUsers;
            context.Load<UserCollection>(users);
            context.ExecuteQuery();

            foreach(User siteuser in users)
            {
                listOfUsers.Add(siteuser.LoginName);
            }

            foreach(List list in lists)
            {
                listOfLists.Add(list.Title);
            }


        }

        protected void CSOM_Click(object sender, EventArgs e)
        {
            string commandAccessToken = ((LinkButton)sender).CommandArgument;
            RetrieveWithCSOM(commandAccessToken);
            WebTitleLabel.Text = siteName;
            CurrentUserLabel.Text = currentUser;
            UserList.DataSource = listOfUsers;
            UserList.DataBind();
            ListList.DataSource = listOfLists;
            ListList.DataBind();

        }
    }
}