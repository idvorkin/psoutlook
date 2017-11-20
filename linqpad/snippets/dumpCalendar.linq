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

// C:\Users\igord\Documents\LINQPad Queries>c:\dropbox\bin_drop\lprun.exe -format=csv dumpCalendar.linq > c:\temp\CalenarItem.csv

var app = new Microsoft.Office.Interop.Outlook.Application();
var mapi = app.GetNamespace("mapi");
var calItems = mapi.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar).Items;
calItems.Sort("[Start]");
calItems.IncludeRecurrences = true;
var start = new DateTime(2014,1,1);
// Dump forward the next week.
var end = DateTime.Now;
var searchString = String.Format ("[Start] > '{0}' AND [Start] < '{1}'", start.ToShortDateString(),end.ToShortDateString());
var appointments = calItems.Restrict(searchString);

appointments.Cast<AppointmentItem>()
.Select(ai=>new {ai.Categories, ai.Start, ai.Duration,
	IsSelfAppointment= ai.Recipients.Count == 1 , IsOneOnOne = ai.Recipients.Count == 2, ai.Organizer
// , ai.Subject
// , recipients = ai.Recipients.Cast<Recipient>().Select(r=>r.Name)}
}) .Dump();