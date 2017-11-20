<Query Kind="Statements">
  <NuGetReference>Microsoft.Office.Interop.Outlook</NuGetReference>
  <Namespace>Microsoft.Office</Namespace>
  <Namespace>Microsoft.Office.Interop</Namespace>
  <Namespace>Microsoft.Office.Interop.Outlook</Namespace>
  <Namespace>System</Namespace>
  <Namespace>System.Collections.Generic</Namespace>
  <Namespace>System.Configuration</Namespace>
  <Namespace>System.Linq</Namespace>
  <Namespace>System.Runtime.InteropServices</Namespace>
  <Namespace>System.Text</Namespace>
  <Namespace>System.Threading.Tasks</Namespace>
</Query>

var app = new Microsoft.Office.Interop.Outlook.Application();
var mapi = app.GetNamespace("mapi");
var sentItems = mapi.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderSentMail).Items;

var peopleIveSentMailToOrderedByFrequency = sentItems.OfType<MailItem>()
.SelectMany(c=>c.Recipients.Cast<Recipient>().Select(r=>r.Name))
.GroupBy(r=>r)
.Select(g=>new {g.Key, Count=g.Count()})
.OrderByDescending(g=>g.Count)
.Distinct()
.Select(r=>r.Key)
.ToArray();

File.WriteAllLines("c:\\temp\\interactees.txt",peopleIveSentMailToOrderedByFrequency);
 	
