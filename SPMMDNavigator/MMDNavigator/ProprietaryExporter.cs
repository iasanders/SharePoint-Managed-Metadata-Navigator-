using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Windows.Forms;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Taxonomy;

namespace MMDNavigator
{
    public class ProprietaryExporter
    {


        /// <summary>
        /// </summary>
        public static bool ExportToPropFormat(SaveFileDialog saveFileDialog, string siteUrl, bool splitSyns, TreeNode tNode, out string msg)
        {
            msg = "OK";

            try
            {
                if (tNode == null)
                {
                    msg = "Cannot export, please select a tree node.";
                    return false;
                }

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    if (tNode.Level == 0)
                    {
                        // export entire termstore
                        Guid termStoreId = new Guid(tNode.Name);

                        ProprietaryExporter.ExportToCsv(saveFileDialog.FileName, siteUrl, termStoreId, null, null, null, splitSyns);

                    }
                    else if (tNode.Level == 1)
                    {
                        // export group
                        Guid termStoreId = new Guid(tNode.Parent.Name);
                        Guid groupId = new Guid(tNode.Name);

                        ProprietaryExporter.ExportToCsv(saveFileDialog.FileName, siteUrl, termStoreId, groupId, null, null, splitSyns);

                    }
                    else if (tNode.Level == 2)
                    {
                        // export termset
                        Guid termStoreId = new Guid(tNode.Parent.Parent.Name);
                        Guid groupId = new Guid(tNode.Parent.Name);
                        Guid termSetId = new Guid(tNode.Name);

                        ProprietaryExporter.ExportToCsv(saveFileDialog.FileName, siteUrl, termStoreId, groupId, termSetId, null, splitSyns);

                    }
                    else if (tNode.Level == 3)
                    {
                        // export term and subterms
                        // export termset
                        Guid termStoreId = new Guid(tNode.Parent.Parent.Parent.Name);
                        Guid groupId = new Guid(tNode.Parent.Parent.Name);
                        Guid termSetId = new Guid(tNode.Parent.Name);

                        SortedList<int, Guid> slTerms = new SortedList<int, Guid>();
                        slTerms.Add(3, new Guid(tNode.Name));

                        ProprietaryExporter.ExportToCsv(saveFileDialog.FileName, siteUrl, termStoreId, groupId, termSetId, slTerms, splitSyns);

                    }
                    else
                    {
                        // export term and subterms

                        SortedList<int, Guid> slTerms = new SortedList<int, Guid>();
                        TreeNode nodeTermSet = tNode;

                        while (nodeTermSet.Level != 2)
                        {
                            if (nodeTermSet.Level > 2)
                                slTerms.Add(nodeTermSet.Level, new Guid(nodeTermSet.Name));
                            nodeTermSet = nodeTermSet.Parent;
                        }

                        Guid termStoreId = new Guid(nodeTermSet.Parent.Parent.Name);
                        Guid groupId = new Guid(nodeTermSet.Parent.Name);
                        Guid termSetId = new Guid(nodeTermSet.Name);


                        ExportToCsv(saveFileDialog.FileName, siteUrl, termStoreId, groupId, termSetId, slTerms, splitSyns);

                    }

                }

            }
            catch (Exception exc)
            {
                msg = exc.Message;
            }

            return (msg == "OK");
        }


        /// <summary>
        /// </summary>
        private static StringBuilder sbOutput = null;


        /// <summary>
        /// Currently ignores any selected terms or subterms to base export on.
        /// </summary>
        private static void ExportToCsv(
                                string fileName,
                                string siteUrl,
                                Guid? termStoreId,
                                Guid? groupId,
                                Guid? termSetId,
                                SortedList<int, Guid> slTerms,
                                bool splitLevels = false
                                )
        {
            sbOutput = new StringBuilder("");

            sbOutput.Append(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}{9}",
                GenUtil.CSVer("termStoreName"),
                GenUtil.CSVer("groupName"),
                GenUtil.CSVer("termSetName"),
                GenUtil.CSVer("level"),
                GenUtil.CSVer("parentId"),
                GenUtil.CSVer("parentName"),
                GenUtil.CSVer("termId"),
                GenUtil.CSVer("termName"),
                GenUtil.CSVer("termLabels"),
                Environment.NewLine));

            using (SPSite site = new SPSite(siteUrl))
            {
                TaxonomySession txsn = new TaxonomySession(site);

                foreach (TermStore _curTermStore in txsn.TermStores)
                {
                    if (termStoreId != null && _curTermStore.Id != termStoreId)
                        continue;

                    foreach (Group _curGroup in _curTermStore.Groups)
                    {
                        if (groupId != null && _curGroup.Id != groupId)
                            continue;

                        foreach (TermSet _curTermSet in _curGroup.TermSets)
                        {
                            if (termSetId != null && _curTermSet.Id != termSetId)
                                continue;

                            foreach (Term _curTerm in _curTermSet.Terms)
                            {
                                WriteTerm(_curTermStore.Name, _curGroup.Name, _curTermSet.Name, _curTerm, 1, splitLevels);
                            }
                        }
                    }
                }
            }


            FileStream fs = new FileStream(fileName, FileMode.Append);
            StreamWriter writer = new StreamWriter(fs);
            writer.Write(sbOutput.ToString());
            writer.Close();
            fs.Close();

        }


        /// <summary>
        /// Recursive writer of terms
        /// </summary>
        private static void WriteTerm(string termStoreName, string groupName, string termSetName, Term term, int level, bool splitLevels)
        {
            string id = GenUtil.CSVer(term.Id.ToString());
            string name = GenUtil.CSVer(term.Name);

            // get labels
            List<string> lstLabels = term.Labels.Where(x => x.Value != term.Name).Select(x => GenUtil.CSVer(x.Value)).ToList<string>();

            //lstLabels.Remove(name);

            string labels = string.Join(",", lstLabels.ToArray<string>());

            if (term.Labels.Count <= 0)
                labels = "";

            // get parent
            Term tParent = term.Parent;

            string parentId = "null";
            string parentName = "null";

            if (tParent != null)
            {
                parentId = GenUtil.CSVer(tParent.Id.ToString());
                parentName = GenUtil.CSVer(tParent.Name);
            }


            if (splitLevels && lstLabels.Count > 0)
            {
                foreach (string lbl in lstLabels)
                {
                    sbOutput.Append(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}{9}",
                        termStoreName,
                        groupName,
                        termSetName,
                        level.ToString(),
                        parentId,
                        parentName,
                        id,
                        name,
                        lbl,
                        Environment.NewLine));
                }
            }
            else
            {
                sbOutput.Append(string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}{9}",
                    termStoreName,
                    groupName,
                    termSetName,
                    level.ToString(),
                    parentId,
                    parentName,
                    id,
                    name,
                    labels,
                    Environment.NewLine));
            }


            if (term.Terms.Count > 0)
            {
                level++;

                foreach (Term childTerm in term.Terms)
                    WriteTerm(termStoreName, groupName, termSetName, childTerm, level, splitLevels);
            }

        }


    }
}
