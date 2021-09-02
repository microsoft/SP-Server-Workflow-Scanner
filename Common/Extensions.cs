using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Common
{
    public static partial class WebExtensions
    {

        /// <summary>
        /// Gets a list of SharePoint lists to scan for modern compatibility
        /// </summary>
        /// <param name="web">Web to check</param>
        /// <returns>List of SharePoint lists to scan</returns>
        public static List<List> GetListsToScan(this Microsoft.SharePoint.Client.Web web, bool showHidden = false)
        {
            List<List> lists = new List<List>(10);

            // Ensure timeout is set on current context as this can be an operation that times out
            web.Context.RequestTimeout = Timeout.Infinite;

            ListCollection listCollection = web.Lists;
            listCollection.EnsureProperties(coll => coll.Include(li => li.Id, li => li.ForceCheckout, li => li.Title, li => li.Hidden, li => li.DefaultViewUrl,
                                                                 li => li.BaseTemplate, li => li.RootFolder, li => li.ListExperienceOptions, li => li.ItemCount,
                                                                 li => li.UserCustomActions, li => li.LastItemUserModifiedDate));
            // Let's process the visible lists
            IQueryable<List> listsToReturn = null;

            if (showHidden)
            {
                listsToReturn = listCollection;
            }
            else
            {
                listsToReturn = listCollection.Where(p => p.Hidden == false);
            }

            foreach (List list in listsToReturn)
            {
                if (list.DefaultViewUrl.Contains("_catalogs"))
                {
                    // skip catalogs
                    continue;
                }

                if (list.BaseTemplate == 544)
                {
                    // skip MicroFeed (544)
                    continue;
                }

                lists.Add(list);
            }

            return lists;
        }
    }
    public static class SiteExtensions
    {
        /// <summary>
        /// Gets all sub sites for a given site
        /// </summary>
        /// <param name="site">Site to find all sub site for</param>
        /// <returns>IEnumerable of strings holding the sub site urls</returns>
        public static IEnumerable<string> GetAllSubSites(this Site site)
        {
            var siteContext = site.Context;
            siteContext.Load(site, s => s.Url);

            try
            {
                siteContext.ExecuteQueryRetry();
            }
            catch (System.Net.WebException clientException)
            {
                Console.WriteLine(clientException.Message.ToString());
                yield break;
            }
            catch (Microsoft.SharePoint.Client.ServerUnauthorizedAccessException unauthorizedException)
            {
                Console.WriteLine(unauthorizedException.Message.ToString());
                yield break;
            }

            var queue = new Queue<string>();
            queue.Enqueue(site.Url);
            while (queue.Count > 0)
            {
                var currentUrl = queue.Dequeue();
                try
                {
                    using (var webContext = siteContext.Clone(currentUrl))
                    {
                        webContext.Load(webContext.Web, web => web.Webs.Include(w => w.Url, w => w.WebTemplate));
                        webContext.ExecuteQueryRetry();
                        foreach (var subWeb in webContext.Web.Webs)
                        {
                            // Ensure these props are loaded...sometimes the initial load did not handle this
                            subWeb.EnsureProperties(s => s.Url, s => s.WebTemplate);

                            // Don't dive into App webs and Access Services web apps
                            if (!subWeb.WebTemplate.Equals("App", StringComparison.InvariantCultureIgnoreCase) &&
                                !subWeb.WebTemplate.Equals("ACCSVC", StringComparison.InvariantCultureIgnoreCase))
                            {
                                queue.Enqueue(subWeb.Url);
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // Eat exceptions when certain subsites aren't accessible, better this then breaking the complete fMedium
                }

                yield return currentUrl;
            }
        }
    }

}
