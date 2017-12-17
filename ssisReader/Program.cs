using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.IO;
using System.Text.RegularExpressions;
using System.Data;
using MarkdownLog;

namespace ssisReader
{
    public enum SqlCompatibilityType { SQL2008, SQL2005 };

    public class Program
    {
        public static SqlCompatibilityType gSqlMode = SqlCompatibilityType.SQL2008;

        public static void Main()
        {
            string package = "";

            //package = "C:\\SSIS\\packages\\Load_CMOR_DELINQUENCY_GROUP.dtsx";
            //package = "C:\\SSIS\\packages\\Datafeed_Quandis_Main.dtsx";
            //package = "C:\\SSIS\\packages\\Load_CMOR_COMMUNICATION_REF_GROUP.dtsx"; 
            package = "C:\\SSIS\\packages\\Source_HHF.dtsx";

            SsisObject ssisobj = new SsisObject();
            
            ssisobj = ParseSsisPackage(package);

            string filename = "C:\\SSIS\\program.md";
            string htmlOutputFilename = "C:\\SSIS\\program.html";
            ProduceSsisDocumentation(ssisobj, filename, htmlOutputFilename);

        }

    /// <summary>
    /// Attempt to read an SSIS package and produce a meaningful C# program
    /// </summary>
    /// <param name="ssis_filename"></param>
    /// <param name="output_folder"></param>
    public static SsisObject ParseSsisPackage(string ssis_filename, SqlCompatibilityType SqlMode = SqlCompatibilityType.SQL2008, bool UseSqlSMO = true)
        {
            XmlReaderSettings set = new XmlReaderSettings();
            set.IgnoreWhitespace = true;
            SsisObject o = new SsisObject();
            gSqlMode = SqlMode;

            // Set the appropriate flag for SMO usage
            ProjectWriter.UseSqlServerManagementObjects = UseSqlSMO;
            
            // Read in the file, one element at a time
            XmlDocument xd = new XmlDocument();
            xd.Load(ssis_filename);
            XmlNodeReader xl = new XmlNodeReader(xd);
            DataSet d = new DataSet();
            d.ReadXml(xl, XmlReadMode.IgnoreSchema);

            ReadObject(xd.DocumentElement, o);

            return o;
        }

        #region Create a document based on the SSIS package 
        /// <summary>
        /// Produce Documentation of an SSIS package
        /// </summary>
        /// <param name="o"></param>
        private static void ProduceSsisDocumentation(SsisObject o, string filename, string htmlOutputFilename)
        {
            // First find all the connection strings 
            var connstrings = from SsisObject c in o.Children where c.DtsObjectType == "DTS:ConnectionManager" select c;
            //ConnectionWriter.WriteConnections(connstrings, filename);

            // Find all the Variables 
            var variables = from SsisObject c in o.Children where c.DtsObjectType == "DTS:Variable" select c;

            // Next, get all the executable functions 
            var functions = from SsisObject c in o.Children where c.DtsObjectType == "DTS:Executable" select c;

            if (!functions.Any())
            {
                var executables = from SsisObject c in o.Children where c.DtsObjectType == "DTS:Executables" select c;
                List<SsisObject> flist = new List<SsisObject>();
                foreach (var exec in executables)
                {
                    flist.AddRange(from e in exec.Children where e.DtsObjectType == "DTS:Executable" select e);
                }
                if (flist.Count == 0)
                {
                    Console.WriteLine("No functions ('DTS:Executable') objects found in the specified file.");
                    return;
                }
                functions = flist;
            }
            
            WriteDocumentation(o, variables, connstrings, functions, filename);
            ConvertToHtml(filename, htmlOutputFilename);
        }

        private static void ConvertToHtml(string inputfilename, string outputfilename)
        {
            string mdText = File.ReadAllText(inputfilename);
            var md = new MarkdownDeep.Markdown();
            md.ExtraMode = true;
            string output = md.Transform(mdText);
            File.WriteAllText(outputfilename, output);            
        }
        #endregion

        /// <summary>
        /// Write a program file that has all the major executable instructions as functions
        /// </summary>
        /// <param name="variables"></param>
        /// <param name="functions"></param>
        /// <param name="p"></param>
        private static void WriteDocumentation(SsisObject o, IEnumerable<SsisObject> variables, IEnumerable<SsisObject> connstrings, IEnumerable<SsisObject> functions, string filename)
        {
            //Get Package Creator Name
            var creatorName = "CreatorName - " + o.Properties["CreatorName"] ;

            //Get Version GUID
            string versionGuid = "VersionGUID - " + o.Properties["VersionGUID"];
                                   
            //Get Package Name
            var packageName = o.DtsObjectName + ".dtsx";
            //var packageName = "PackageName - " + o.DtsObjectName;

            using (SourceWriter.SourceFileStream = new StreamWriter(filename, false, Encoding.UTF8))
            {
                SourceWriter.WriteLine(packageName.ToMarkdownHeader());
                SourceWriter.WriteLine(creatorName.ToMarkdownSubHeader());
                SourceWriter.WriteLine(versionGuid.ToMarkdownSubHeader());

                WriteVariables(variables);
                WriteConnectionManagers(connstrings);

                SourceWriter.WriteLine(@"");
                SourceWriter.WriteLine(@"");
                //Write each executable out as a function
                SourceWriter.WriteLine(@"Executables".ToMarkdownSubHeader());

                SourceWriter.WriteLine(@"* #### Root Level Executables");
                foreach (SsisObject v in functions)
                {
                    SourceWriter.WriteLine(@"   * {0}", v.DtsObjectName);
                    SourceWriter.WriteLine(@"       * {0}", v.Description);
                }

                SourceWriter.WriteLine(@"* #### Executable Flows");
                foreach (SsisObject v in functions)
                {
                    v.EmitFunctionsAsSequence("\t", new List<ProgramVariable>());
                }

                SourceWriter.WriteLine(@"* #### Executables");
                foreach (SsisObject v in functions)
                {
                    v.EmitFunction("\t", new List<ProgramVariable>());
                }
            }
        }

        /// <summary>
        /// Write all variables to file
        /// </summary>
        /// <param name="variables"></param>
        private static void WriteVariables(IEnumerable<SsisObject> variables)
        {
            SourceWriter.WriteLine(@"***");
            // Write each variable out as if it's a global
            SourceWriter.WriteLine(@"Variables".ToMarkdownSubHeader());
                foreach (SsisObject v in variables)
                {
                    v.EmitVariable("        ", true);
                }                
        }
        
        /// <summary>
        /// Write all ConnectionManagers to file
        /// </summary>
        /// <param name="conns"></param>
        private static void WriteConnectionManagers(IEnumerable<SsisObject> conns)
        {
            SourceWriter.WriteLine(@"***");
            SourceWriter.WriteLine(@"ConnectionsManagers".ToMarkdownSubHeader());
            
            // Write each Conection String
            foreach (SsisObject connstr in conns)
            {
                string connString = "Not Found";
                string creationName = "Not Found";
                string DtsId = connstr.DtsId.ToString();
                string description = "Not Found";
                //string connObject = "Not Found";

                SourceWriter.WriteLine(String.Format(@"* **{0}**", connstr.DtsObjectName));
                connstr.Properties.TryGetValue("CreationName", out creationName);
                connstr.Properties.TryGetValue("Description", out description);
                
                var v = connstr.GetChildByType("DTS:ObjectData");
                if (v != null)
                {
                    // Look for a SQL Connection string
                    var v2 = v.GetChildByType("DTS:ConnectionManager");
                    if (v2 != null)
                    {
                        v2.Properties.TryGetValue("ConnectionString", out connString);                        
                        // If not, look for an SMTP connection string
                    }
                    else
                    {
                        v2 = v.GetChildByType("SmtpConnectionManager");
                        if (v2 != null)
                        {
                            v2.Attributes.TryGetValue("ConnectionString", out connString);
                        }
                        else
                        {
                            Console.WriteLine("Help");
                        }
                    }
                }

                SourceWriter.WriteLine();

                var data = new[]
                {
                    new {Name = "Connection String", Value =  connString},
                    new {Name = "Creation Name", Value =  creationName},
                    new {Name = "DtsId", Value =  DtsId},
                    new {Name = "Description", Value =  description}
                };

                SourceWriter.WriteLine(data.ToMarkdownTable());
              
            }
        }
    
        #region Read in an SSIS DTSX file
        /// <summary>
        /// Recursive read function
        /// </summary>
        /// <param name="xr"></param>
        /// <param name="o"></param>
        private static void ReadObject(XmlElement el, SsisObject o)
                {
                    // Read in the object name
                    o.DtsObjectType = el.Name;

                    // Read in attributes
                    foreach (XmlAttribute xa in el.Attributes)
                    {
                        o.Attributes.Add(xa.Name, xa.Value);
                    }

                    // Iterate through all children of this element
                    foreach (XmlNode child in el.ChildNodes)
                    {

                        // For child elements
                        if (child is XmlElement)
                        {
                            XmlElement child_el = child as XmlElement;

                            // Read in a DTS Property
                            if (child.Name == "DTS:Property" || child.Name == "DTS:PropertyExpression")
                            {
                                ReadDtsProperty(child_el, o);

                                // Everything else is a sub-object
                            }
                            else
                            {
                                SsisObject child_obj = new SsisObject();
                                ReadObject(child_el, child_obj);
                                child_obj.Parent = o;
                                o.Children.Add(child_obj);
                            }
                        }
                        else if (child is XmlText)
                        {
                            o.ContentValue = child.InnerText;
                        }
                        else if (child is XmlCDataSection)
                        {
                            o.ContentValue = child.InnerText;
                        }
                        else
                        {
                            Console.WriteLine("Help");
                        }
                    }
                }

                /// <summary>
                /// Read in a DTS property from the XML stream
                /// </summary>
                /// <param name="xr"></param>
                /// <param name="o"></param>
                private static void ReadDtsProperty(XmlElement el, SsisObject o)
                {
                    string prop_name = null;

                    // Read all the attributes
                    foreach (XmlAttribute xa in el.Attributes)
                    {
                        if (String.Equals(xa.Name, "DTS:Name", StringComparison.CurrentCultureIgnoreCase))
                        {
                            prop_name = xa.Value;
                            break;
                        }
                    }

                    // Set the property
                    o.SetProperty(prop_name, el.InnerText);
                }
                #endregion
            }
        }
