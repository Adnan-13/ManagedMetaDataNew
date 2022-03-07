using System;
using OfficeDevPnP.Core;
using Microsoft.SharePoint.Client;
using System.Security;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client.Taxonomy;

namespace ManagedMetaDataNew
{
    class Api
    {
        string siteURL = "";
        string listTitle = "";
        string username = "shahidul@genweb2bd.onmicrosoft.com";
        string pasword = "!!Arizona11";

        public Api(string siteURL, string listTitle)
        {
            this.siteURL = siteURL;
            this.listTitle = listTitle;
        }

        ClientContext getClientContext()
        {
            SecureString securePassword = new SecureString();

            foreach (char c in pasword)
            {
                securePassword.AppendChar(c);
            }

            var credentials = new SharePointOnlineCredentials(username, securePassword);
            ClientContext clientContext = new ClientContext(siteURL);
            clientContext.Credentials = credentials;

            return clientContext;
        }
        public void addMultiTerm()
        {
            ClientContext clientContext = getClientContext();

            List list = clientContext.Web.GetListByTitle(listTitle);
            Field field = list.Fields.GetByTitle("Product");

            TaxonomyField txField = clientContext.CastTo<TaxonomyField>(field);

            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            ListItem newItem = list.AddItem(listItemCreationInformation);

            newItem["Title"] = "Some Item 2";
            newItem["ThirdParty"] = "Some Item 2s third party";
            newItem["VersionNumber"] = 101;

            var txCollection = new TaxonomyFieldValueCollection(clientContext, null, txField);

            txCollection.PopulateFromLabelGuidPairs(@"GWP 5.1|609120c7-33ed-47b2-9ab6-a812ffea7675");
            txCollection.PopulateFromLabelGuidPairs(@"GWP 5.2|dea1e840-377a-47c9-bc5f-9e4d5eebd925"); //needs to be a different value.


            //USE SetFieldValueByValueCollection Not SetFieldValueByCollection as we have a TaxonomyFieldValueCollection
            txField.SetFieldValueByValueCollection(newItem, txCollection);

            newItem.Update();

            clientContext.Load(field);
            clientContext.ExecuteQuery();

            Console.WriteLine("Done");
        }
        public void addSingleTerm()
        {
            ClientContext clientContext = getClientContext();

            List list = clientContext.Web.GetListByTitle(listTitle);
            Field field = list.Fields.GetByTitle("Product");

            TaxonomyField txField = clientContext.CastTo<TaxonomyField>(field);

            ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
            ListItem newItem = list.AddItem(listItemCreationInformation);

            newItem["Title"] = "Some Item 2";
            newItem["ThirdParty"] = "Some Item 2s third party";
            newItem["VersionNumber"] = 101;


            //Useless for multivalue
            TaxonomyFieldValue termValue1 = new TaxonomyFieldValue
            {
                Label = "GWP 5.1",
                TermGuid = "609120c7-33ed-47b2-9ab6-a812ffea7675",
                WssId = -1
            };


            txField.SetFieldValueByValue(newItem, termValue1);

            newItem.Update();

            clientContext.Load(field);
            clientContext.ExecuteQuery();

            Console.WriteLine("Done");
        }

        void getTerms()
        {
            using (ClientContext clientContext = getClientContext())
            {
                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

                TermStore termStore = taxonomySession.TermStores.GetById(new Guid("85692aed-c030-4329-a547-1cbcd8c84444"));

                TermGroup termGroup = termStore.Groups.GetById(new Guid("7f5fdac7-f384-4e8c-b420-eecd7463c1fb"));

                TermSet termSet = termGroup.TermSets.GetByName("CMP");

                TermCollection terms = termSet.Terms;

                clientContext.Load(terms);

                clientContext.ExecuteQuery();


                foreach (Term term in terms)
                {
                    Console.WriteLine(term.Name);
                    /*Console.WriteLine(term.TermsCount);*/
                    TermCollection subTerms = term.Terms;

                    clientContext.Load(subTerms);
                    clientContext.ExecuteQuery();

                    foreach (Term subTerm in subTerms)
                    {
                        Console.WriteLine("     " + subTerm.Name);
                    }

                    /*termToAdd = subTerms[0];
                    break;*/
                }


            }
        }

        void addByTerm()
        {
            using (ClientContext clientContext = getClientContext())
            {
                List list = clientContext.Web.GetListByTitle(listTitle);
                Field field = list.Fields.GetByTitle("Product");
                TaxonomyField txField = clientContext.CastTo<TaxonomyField>(field);

                ListItemCreationInformation listItemCreationInformation = new ListItemCreationInformation();
                ListItem newItem = list.AddItem(listItemCreationInformation);




                TaxonomySession taxonomySession = TaxonomySession.GetTaxonomySession(clientContext);

                TermStore termStore = taxonomySession.TermStores.GetById(new Guid("85692aed-c030-4329-a547-1cbcd8c84444"));

                TermGroup termGroup = termStore.Groups.GetById(new Guid("7f5fdac7-f384-4e8c-b420-eecd7463c1fb"));

                TermSet termSet = termGroup.TermSets.GetByName("CMP");

                TermCollection terms = termSet.Terms;

                clientContext.Load(terms);
                clientContext.ExecuteQuery();

                Term term1 = terms.GetByName("Apex").Terms.GetByName("GWP 5.1");
                Term term2 = terms.GetByName("Apex").Terms.GetByName("GWP 5.2");

                newItem["Title"] = "Some Item Added using term";
                newItem["ThirdParty"] = "Some Item's third party";
                newItem["VersionNumber"] = 101;
                txField.SetFieldValueByTerm(newItem, term1, 102);

                newItem.Update();
                clientContext.Load(field);
                clientContext.ExecuteQuery();

                Console.WriteLine("Done");
            }


        }

        public void Test()
        {
            Console.WriteLine("Testing");

            /*getTerms();*/
            addByTerm();

        }
    }
}
