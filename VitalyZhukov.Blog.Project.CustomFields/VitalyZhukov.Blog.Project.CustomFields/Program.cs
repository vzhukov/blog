using System;
using System.Configuration;
using System.Security;
using Microsoft.ProjectServer.Client;
using Microsoft.SharePoint.Client;

namespace VitalyZhukov.Blog.Project.CustomFields
{
    class Program
    {
        static void Main(string[] args)
        {
            // Environment variables
            var siteUrl = ConfigurationManager.AppSettings["pwa:SiteUrl"];
            var login = ConfigurationManager.AppSettings["pwa:Login"];
            var password = ConfigurationManager.AppSettings["pwa:Password"];
            var fieldId = new Guid(ConfigurationManager.AppSettings["pwa:FieldId"]);
            var resourceId = new Guid(ConfigurationManager.AppSettings["pwa:ResourceId"]);

            // Store password in secure string
            var securePassword = new SecureString();
            foreach (var c in password) securePassword.AppendChar(c);

            // Project instance credentials
            var creds = new SharePointOnlineCredentials(login, securePassword);

            // Initiation of the client context
            using (var ctx = new ProjectContext(siteUrl))
            {
                ctx.Credentials = creds;

                // Retrieve Enterprise Custom Field
                var field = ctx.CustomFields.GetByGuid(fieldId);

                // Load InernalName property, we will use it to get the value
                ctx.Load(field,
                    x=>x.InternalName);

                // Execture prepared query on server side
                ctx.ExecuteQuery();

                var fieldInternalName = field.InternalName;

                // Retrieve recource by its Id
                var resource = ctx.EnterpriseResources.GetByGuid(resourceId);

                // !
                // Load custom field value
                ctx.Load(resource,
                    x => x[fieldInternalName]);
                ctx.ExecuteQuery();


                // Update ECF value
                resource[fieldInternalName] = "Vitaly Zhukov";
                ctx.EnterpriseResources.Update();
                ctx.ExecuteQuery();
                
                // Get ECF value from server
                ctx.Load(resource,
                    x=>x[fieldInternalName]);
                ctx.ExecuteQuery();

                Console.WriteLine("ECF value: " + resource[fieldInternalName]);

                Console.ReadLine();
            }
        }
    }
}
