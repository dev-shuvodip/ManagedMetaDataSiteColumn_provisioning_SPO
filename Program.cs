using System;
using System.Configuration;
using System.Security;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace SP_CSOM_DEMO2
{
    class Program
    {
        private static readonly ClientContext context = new ClientContext(ConfigurationManager.AppSettings["SPOSite"]);

        static void Main(string[] args)
        {
            InitiateAuthentication(context);

            CreateSPManagedMetaDataField("MetaDataLanguage", "MetaData Language", "RFP Columns", "KPMGLanguages", "Indian");
            Console.ReadLine();
        }
        private static void CreateSPManagedMetaDataField(string columnName, string displayName, string columnGroup, string TermGroup, string TermSet)
        {
            InitiateAuthentication(context);

            Web web = context.Web;
            FieldCollection fieldCollection = web.AvailableFields;
            context.Load(web, w => w.AvailableFields);
            context.ExecuteQuery();

            bool siteColumnExists = false;
            foreach (Field field in fieldCollection)
            {
                if (columnName.ToLower() == field.InternalName.ToLower())
                {
                    siteColumnExists = true;
                    break;
                }
            }

            if (!siteColumnExists)
            {
                Guid FieldID = Guid.NewGuid();
                string columnTaxonomySchema = "<Field ID='{" + FieldID + "}'" + $" Type='TaxonomyFieldType' Name='{columnName}' DisplayName='{displayName}' Description='Managed Metadata Field' Required='False' EnforceUniqueValues='False' Group='{columnGroup}'   />";

                TaxonomySession session = TaxonomySession.GetTaxonomySession(context);
                TermStore store = session.GetDefaultSiteCollectionTermStore();
                TermGroup group = store.Groups.GetByName(TermGroup);
                TermSet set = group.TermSets.GetByName(TermSet);
                context.Load(store, s => s.Id);
                context.Load(set, s => s.Id);
                context.ExecuteQuery();

                Field siteTaxonomyColumn = web.Fields.AddFieldAsXml(columnTaxonomySchema, false, AddFieldOptions.DefaultValue);
                context.Load(siteTaxonomyColumn);
                context.ExecuteQuery();
                Console.WriteLine($"Site Column {siteTaxonomyColumn.Title} created.");

                TaxonomyField siteTaxonomyColumnBind = context.CastTo<TaxonomyField>(siteTaxonomyColumn);
                siteTaxonomyColumnBind.SspId = store.Id;
                siteTaxonomyColumnBind.TermSetId = set.Id;
                siteTaxonomyColumnBind.Update();
                context.ExecuteQuery();
                Console.WriteLine($"Term Set unique ID: {set.Id} mapped to Site Column {siteTaxonomyColumn.Title}");

                List targetlist = web.Lists.GetByTitle("RFP");
                context.Load(targetlist);
                context.ExecuteQuery();

                Field targetSiteColumn = web.AvailableFields.GetByInternalNameOrTitle(columnName);
                context.Load(targetSiteColumn);
                context.ExecuteQuery();

                targetlist.Fields.Add(targetSiteColumn);
                targetlist.Update();
                context.ExecuteQuery();
                Console.WriteLine($"Site Column {targetSiteColumn.Title} added to list {targetlist.Title}");
            }
            else
            {
                Console.WriteLine($"Site Column {displayName} already exists.");
            }
        }

        private static string GetSPOUserName()
        {
            try
            {
                return ConfigurationManager.AppSettings["SPOUsername"];
            }
            catch
            {
                throw;
            }
        }

        private static SecureString GetSPOSecureStringPassword()
        {
            try
            {
                SecureString secureString = new SecureString();
                foreach (char c in ConfigurationManager.AppSettings["SPOPassword"])
                {
                    secureString.AppendChar(c);
                }
                return secureString;
            }
            catch
            {
                throw;
            }
        }

        private static void InitiateAuthentication(ClientContext ctx)
        {
            ClientContext context = ctx;
            context.AuthenticationMode = ClientAuthenticationMode.Default;
            context.Credentials = new SharePointOnlineCredentials(GetSPOUserName(), GetSPOSecureStringPassword());
        }
    }
}
