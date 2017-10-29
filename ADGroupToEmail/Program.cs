using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.DirectoryServices.AccountManagement;
using System.Windows.Forms;

namespace ADGroupToEmail
{
    class Program
    {
        static string DomainOption = "--domain";
        static string RecursiveOption = "--recursive";

        public static void Main(string[] args)
        {
            string Group;
            string Domain;
            bool Recursive = false;
            if (args.Length < 2)
            {
                PrintHelpMessage();
            } else
            {
                Group = args[0].Trim();
                Domain = args[1].Trim();
                if (args.Contains(RecursiveOption))
                {
                    Recursive = true;
                }

                QueryActiveDirectory(Group, Domain, Recursive);
            }
        }

        /// <summary>
        /// Query ActiveDirectory group
        /// </summary>
        /// <param name="Group"></param>
        /// <param name="Domain"></param>
        /// <param name="Recursive"></param>
        /// <returns></returns>
        private static void QueryActiveDirectory(string Group, string Domain, bool Recursive) 
        {
            try
            {
                PrincipalContext Context = new PrincipalContext(ContextType.Domain, Domain);
                GroupPrincipal GroupPrincipal = GroupPrincipal.FindByIdentity(Context, Domain);
                if (GroupPrincipal == null)
                {
                    WriteErrorMessage($"Error: Unable to find group {Domain}\\{Group}.");
                }
                else
                {
                    var Users = GroupPrincipal.GetMembers(Recursive);
                    string UserString = string.Empty;
                    foreach (UserPrincipal User in Users)
                    {
                        UserString = string.Concat(UserString, User.EmailAddress, ",");
                    }
                    UserString = UserString.TrimEnd(',');

                    if (string.IsNullOrEmpty(UserString)) WriteErrorMessage("Error: No users found in the group specified.");
                    else
                    {
                        Clipboard.SetText(UserString);
                        WriteInfoMessage("Copied user emails to clipboard. Press any key to exit.");
                        Console.ReadKey();
                    }

                }

            } catch (PrincipalServerDownException exc)
            {
                WriteErrorMessage("Error: Unable to contact LDAP server specified. Check the domain and try again.");
            } catch (MultipleMatchesException exc)
            {
                WriteErrorMessage($"Error: Found multiple groups with name {Domain}\\{Group}. \r\nCheck your setttings and try again.");
            }
        }

        private static void WriteInfoMessage(string Message)
        {
            Console.WriteLine(Message);
        }

        private static void WriteErrorMessage(string Message)
        {
            Console.Error.WriteLine(Message);
        }

        /// <summary>
        /// Print the usage method to the Console.
        /// </summary>
        private static void PrintHelpMessage()
        {
            string Message = @"Usage: ADGroupToEmail.exe [group] [domain] [--recursive]
Parameters
    group: the name of the Group you want to query.
    domain: the domain of the group that you want to query.
Optional Parameters:
    recursive: whether or not to copy the emails of users in any subgroups found in the query.
            ";
            Console.WriteLine(Message);
        }
    }
}
