using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

using PdfSharp.Pdf;
using PdfSharp.Pdf.Content;
using PdfSharp.Pdf.Content.Objects;
using System.IO;
using System.Reflection;
using System.Net;

namespace DebtCollectionDownloader
{
    class Program
    {
        static void Main(string[] args)
        {
            /*
             * The structure of this program is as follows: first, it downloads the PDF containing the docket
             * Then, it splits the PDF into individual pages
             * Then, each page is split into each individual entry
             * These entries are then filtered for 3 things
             *      1) whether they are debt collection actions
             *      2) the date and time (here, Wednesday at 1PM)
             *      3) whether the second party (generally the debtor) is represented by counsel
             * All actions which meet all three criteria are then saved into a CSV file
             * Both the CSV file and original PDF are saved to the directory where the program runs from
             * 
             */

            //Downloads the file and opens as a PDF
            WebClient downloader = new WebClient();
            string directory = "http://www.utcourts.gov/cal/data/";
            string filename = "SLC_Calendar.pdf";
            downloader.DownloadFile(directory + filename, filename);
            var data = downloader.OpenRead(directory + filename);
            PdfDocument document = PdfSharp.Pdf.IO.PdfReader.Open(filename);

            //Converts the PDF into individual pages
            //Note the file format used here is a nested list
            //The inner list contains all of the strings on a page, the outer list contains those inner lists
            List<IEnumerable<string>> ExtractedText = ReadPdf(document);

            //Takes each page, splits into individual entries and filters for debt collection
            //A better version of this code would probably make each of these actions separate functions
            //However, there are two reasons why this was not done
            //First, the memory requirements would be unnecessarily large if all entries were saved, then filtered
            //Second, filtering for debt collection on the processed entries is surprisingly difficult
            List<Entry> ExtractedEntries = ExtractData(ExtractedText);

            //Finally, the filtered entries are saved to a CSV file
            //This is where the filtering by time and date takes place, as well as filtering for pro se
            //CSV is essentially a simplified spreadsheet format
            WriteToCSV(ExtractedEntries);
        }

        static List<IEnumerable<string>> ReadPdf(PdfDocument doc)
        {
            doc.Pages.RemoveAt(0);
            List<IEnumerable<string>> texts = new List<IEnumerable<string>>();
            foreach (PdfPage page in doc.Pages)
            {
                //Converts each PDF page into a collection of strings
                //The actual code to do so was copied from StackOverflow and is not important
                IEnumerable<string> q = PdfSharpExtensions.ExtractText(page);
                texts.Add(q);
            }
            return texts;
        }

        static List<Entry> ExtractData(List<IEnumerable<string>> extractedPDF)
        {
            List<Entry> ToReturn = new List<Entry>();
            foreach (IEnumerable<string> page in extractedPDF)
            {
                ToReturn.AddRange(ExtractEntries(page));
            }
            return ToReturn;
        }

        static List<Entry> ExtractEntries(IEnumerable<string> strings)
        {
            /*
             * All entries have a few things in common:
             * First, they all have a timestamp before the first real entry, and the ##:## pattern is repeated nowhere else
             * Second, all entries are separated by a line of hyphens
             * Third, the date is done on a per-page basis, unlike the per-entry basis for everything else
             */
            List<Entry> ToReturn = new List<Entry>();
            bool PastTimeStamp = false;
            List<List<string>> Entries = new List<List<string>>
            {
                new List<string>()
            };

            //This is a Regular expression used to do advanced comparisons of strings
            //Here, the goal is to find the time of day in which the action takes place
            Regex TimeChecker = new Regex(@"[0-2][0-9]:[0-5][0-9] (AM|PM)");

            //Same idea, but for Date
            //All date formats are in the form "Month dd, yyyy"
            Regex DateCheck = new Regex(@"(Jan(uary)?|Feb(ruary)?|Mar(ch)?|Apr(il)?|May|Jun(e)?|Jul(y)?|Aug(ust)?|Sep(tember)?|Oct(ober)?|Nov(ember)?|Dec(ember)?)\s+\d{1,2},\s+\d{4}");

            //Finally, the court number
            //The only relevant court for us is S34, which makes the regex trivial
            //To be honest, this could be done with String.Contains(), but this is more future-proof (ie, checking for several at once)
            Regex CourtCheck = new Regex(@"S34");

            //The variables above do not record the time and date itself, only that the string matches the format given
            //Therefore, we need two more variables to store the actual time and date
            string Datestamp = "empty";
            string Timestamp = "empty";
            string Courtstamp = "empty";

            foreach (string s in strings)
            {
                //This portion of the code looks a bit weird, but it is important
                //Because the date is per-page and time is per-entry, the date must be recorded before any time  is recorded
                //Many of the entries list a date inside (seemingly to reference prior court orders)
                //We do not want to capture these dates by mistake
                //Therefore, once a time is detected, no more dates matter
                var datecheck = DateCheck.Match(s);
                if (datecheck.Success & !PastTimeStamp)
                {
                    Datestamp = datecheck.Value;
                }
                var courtcheck = CourtCheck.Match(s);
                if (courtcheck.Success & !PastTimeStamp)
                {
                    Courtstamp = courtcheck.Value;
                }

                var matches = TimeChecker.Match(s);
                if (matches.Success & PastTimeStamp)
                {
                    Timestamp = matches.Value;
                }
                if (matches.Success)
                {
                    Timestamp = matches.Value;
                    PastTimeStamp = true;
                }


                if (PastTimeStamp)
                {
                    if (s.Contains("-----")) //entries are generally separated by a line of hyphens (except the end of the page)
                    {
                        //Attaching the time and date stamps to the entry
                        //This prevents having to make the date a separate variable, and re-parsing for time
                        Entries.Last().Add(Datestamp);
                        Entries.Last().Add(Timestamp);
                        Entries.Last().Add(Courtstamp);

                        Entries.Add(new List<string>()); // where the separator appears, we start a new list
                    }
                    else if (s == strings.Last())
                    {
                        //Sometimes, there is no dividing hyphen line for the last entry on a page
                        //This requires a different conditional to catch
                        //Note that the last line is not added, as it is generally "Page 123 of" and therefore not helpful
                        //Note also that no next string is created, as this is the last entry on the page
                        Entries.Last().Add(Datestamp);
                        Entries.Last().Add(Timestamp);
                        Entries.Last().Add(Courtstamp);
                        break;
                    }
                    else
                    {
                        Entries.Last().Add(s); // where there isn't a separator, we add the string to the list
                    }
                }
            }

            foreach (var stringlist in Entries)
            {
                if (IsDebtCollection(stringlist)) //Any other type of action is not saved
                {
                    ToReturn.Add(ConvertStringsToEntries(stringlist));

                }
            }

            return ToReturn;
        }

        static Entry ConvertStringsToEntries(List<string> strings)
        {
            Entry toReturn = new Entry();
            //This is the Regex match for the case number
            Regex TitleCapture = new Regex(@"\s[0-9]+\s+");
            //The attorney is generally on the same line as the party, separated by this phrase
            //Note that, because there can be multiple parties and multiple attorneys, it is hard to separate the two after the first line
            //This is fine, however, for two reasons:
            //First, only the first creditor and first creditor's attorney are relevant on the creditor side
            //Second, we only keep debtors who do not have counsel, so any name on the debtor side is of a party, not counsel
            Regex PartyCapture = new Regex(@"ATTY");

            //The time and date stamps were added in as the last line before
            //They are saved separately and removed now to avoid complicating the text extraction
            string courtstamp = strings.Last();
            strings.RemoveAt(strings.Count() - 1);
            string timestamp = strings.Last();
            strings.RemoveAt(strings.Count() - 1);
            string datestamp = strings.Last();
            strings.RemoveAt(strings.Count() - 1);

            //Because they are saved separately, they need to be parsed separately
            //The easiest way to do this is to create DateTime variables for each, then combine the relevant info into a new DateTime
            //These intermediate variables are not strictly necessary, but do improve readability
            //It could alternately be written as new DateTime(DateTime.Parse(datestamp).Year,DateTime.Parse(datestamp).Month ...
            DateTime intermediatedate = DateTime.Parse(datestamp);
            DateTime intermediatetime = DateTime.Parse(timestamp);

            toReturn.TimeAndDate = new DateTime(intermediatedate.Year, intermediatedate.Month, intermediatedate.Day,
                intermediatetime.Hour, intermediatetime.Minute, intermediatetime.Second);
            toReturn.CourtNumber = courtstamp;

            /*Current strategy:
             * 
             * Right now, only want defendants who are unrepresented, but want all such defendants
             * 
             * All cases follow essentially the same pattern:
             * Handful of irrelevant space lines
             * Among relevant lines:
             * First one has the title
             * Second one has the case number and type of claim (always debt collection)
             * Third has party and attorney
             * All subsequent lines before the V are additional parties or attorneys
             * The V
             * Second Party and Attorney
             * More parties and attorneys
             * 
             * makes sense to first clear all irrelevant lines, then work through each one to identify issues
             */

            //Because of the nature of the PDF->text conversion, much of the formatting turns into empty lines
            //These are removed because they interfere with the order described above
            for (int i = strings.Count - 1; i >= 0; i--)
            {
                if (strings[i].Length < 2)
                    strings.RemoveAt(i);
                else
                    strings[i] = strings[i].Replace(',', '|'); //Removing commas because they interfere with the CSV format
            }


            // the first string has number on the list (which is unimportant), and the title of the action (motion in limine etc)
            var titlematch = TitleCapture.Match(strings[0]);
            toReturn.Title = strings[0].Remove(0, titlematch.Index + titlematch.Length);

            // the second string contains the type (always debt collection because of previous filtering) and case number
            var casenumbermatch = Regex.Match(strings[1], @"[0-9]+");
            toReturn.CaseNumber = Int32.Parse(casenumbermatch.Captures[0].ToString());

            // the third string contains the first party and their attorney, the only relevant ones here
            var firstpartyattorney = IsolatePartyAttorney(strings[2]);
            toReturn.Party1 = firstpartyattorney[0];
            toReturn.Attorney1 = firstpartyattorney[1];

            int index = 3;
            while (!strings[index].Contains("VS.")) index++; // finding the string with the VS. to iterate past the remaining coplaintiffs
            index++; //moving one past it into the useful range (otherwise the index would be on the line with "VS.")

            //As written here, this only reads additional parties if the debtors are pro se
            //This is fine for now, because only pro se defendants are considered here
            //However, if subsequent work includes parties who are represented by counsel, this will likely need to change
            var secondpartyattorney = IsolatePartyAttorney(strings[index]);

            if (secondpartyattorney[1] == "None") //filtering for pro se defendants
            {
                toReturn.Party2.Add(secondpartyattorney[0]);
                index++;
                while (index < strings.Count)
                {
                    toReturn.Party2.Add(strings[index].Trim());
                    index++;
                }
            }
            toReturn.Attorney2 = secondpartyattorney[1];

            return toReturn;
        }

        static List<string> IsolatePartyAttorney(string raw)
        {
            //The first line of the party section will always be in the format "Party Name (spaces) ATTY: Attorney Name"
            //Therefore, this function splits it in two
            //The first entry is the party, the second entry is the attorney (or None if pro se)
            List<string> ToReturn = new List<string>();
            var partystartpoint = Regex.Match(raw, @"\S\B.+ATTY").Index;//everything before the ATTY: label
            var partyendpoint = Regex.Match(raw, @"\s+ATTY").Index; //all of the whitespace between the party and ATTY:
            var partyname = raw.Substring(partystartpoint, partyendpoint - partystartpoint);

            ToReturn.Add(partyname);

            var attorneymatcher = Regex.Match(raw, @"ATTY:(.*)?");
            string attorneyname = attorneymatcher.Captures[0].ToString().Substring(6); //start beyond ATTY: label
            if (attorneyname == "") attorneyname = "None";

            ToReturn.Add(attorneyname);

            return ToReturn;
        }

        static bool IsDebtCollection(List<string> strings)
        {
            foreach (string s in strings)
            {
                if (s.Contains("Debt") & s.Contains("Collection"))
                    return true;
            }

            return false;


        }

        static void WriteToCSV(List<Entry> entries)
        {
            //The CSV file format is an extremely simple way of creating spreadsheets
            //On each individual row, entries are separated by a comma (hence the name "Comma Separated Values")
            //      This is why commas in the party names are separated by pipelines instead of commas
            //      Otherwise and attorney listed as "Doe, John" may produce two separate entries, "Doe" and "John"
            //      This would interfere with the formatting of the list, so it instead entered as "Doe | John"
            //Individual rows are separated by a newline operator (generally "\n", but here using C# functions instead)
            var csv = new StringBuilder();
            var header = "Case Number,Title,Day,Time,First Party,First Party Attorney,Second Party,Date Ran, Court Number,";
            var blanktitles = "Address,City,State,Zip,Phone Number,Randomize1,Randomize2,Attended Hearing,";
            csv.AppendLine(header + blanktitles);
            foreach (Entry e in entries)
            {
                if (e.TimeAndDate.DayOfWeek == DayOfWeek.Wednesday && e.TimeAndDate.TimeOfDay.TotalHours == 13)
                {
                    if (e.Attorney2 == "None")
                    {
                        if (e.CourtNumber == "S34")
                        {
                            foreach (string secondpartyname in e.Party2)
                            {
                                var newline = string.Format("{0},{1},{2},{3},{4},{5},{6},{7},{8}",
                                    e.CaseNumber, e.Title, e.TimeAndDate.ToShortDateString(), e.TimeAndDate.ToShortTimeString(),
                                    e.Party1, e.Attorney1, secondpartyname, DateTime.Now.ToShortDateString(), e.CourtNumber);
                                csv.AppendLine(newline); //This adds the newline automatically
                            }
                        }

                    }
                }

            }

            File.WriteAllText("ProSeDebtCollection.csv", csv.ToString());
        }

    }

    //This is the class which contains the information relevant to final saving
    class Entry
    {
        public string Title;
        public string Party1 = null;
        public string Attorney1 = "none";
        public List<string> Party2 = new List<string>();
        public string Attorney2 = "none";
        public DateTime TimeAndDate;
        public string CourtNumber;

        public int CaseNumber;

    }
}

//Ignore all of this, it is code copied from StackOverflow to split the PDF into pages
public static class PdfSharpExtensions
{
    public static IEnumerable<string> ExtractText(this PdfPage page)
    {
        var content = ContentReader.ReadContent(page);
        var text = content.ExtractText();
        return text;
    }

    public static IEnumerable<string> ExtractText(this CObject cObject)
    {
        if (cObject is COperator)
        {
            var cOperator = cObject as COperator;
            if (cOperator.OpCode.Name == OpCodeName.Tj.ToString() ||
                cOperator.OpCode.Name == OpCodeName.TJ.ToString())
            {
                foreach (var cOperand in cOperator.Operands)
                    foreach (var txt in ExtractText(cOperand))
                        yield return txt;
            }
        }
        else if (cObject is CSequence)
        {
            var cSequence = cObject as CSequence;
            foreach (var element in cSequence)
                foreach (var txt in ExtractText(element))
                    yield return txt;
        }
        else if (cObject is CString)
        {
            var cString = cObject as CString;
            yield return cString.Value;
        }
    }
}