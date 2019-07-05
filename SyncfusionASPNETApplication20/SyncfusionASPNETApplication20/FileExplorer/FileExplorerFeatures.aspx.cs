using Syncfusion.JavaScript;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
namespace SyncfusionASPNETApplication3
{
    public partial class FileExplorerFeatures : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
        }
        [System.Web.Services.WebMethod]
        public static object FileActionDefault(string ActionType, string Path, string ExtensionsAllow, string LocationFrom, string LocationTo, string Name, string[] Names, string NewName, string Action, bool CaseSensitive,string SearchString, IEnumerable<CommonFileDetails> CommonFiles)
        {
            FileExplorerOperations opeartion = new FileExplorerOperations();
            switch (ActionType)
            {
                case "Read":
                    return (opeartion.Read(Path, ExtensionsAllow));
                case "CreateFolder":
                    return (opeartion.CreateFolder(Path, Name));
                case "Paste":
                    opeartion.Paste(LocationFrom, LocationTo, Names, Action, CommonFiles);
                    break;
                case "Remove":
                    opeartion.Remove(Names, Path);
                    break;
                case "Rename":
                    opeartion.Rename(Path, Name, NewName, CommonFiles);
                    break;
                case "GetDetails":
                    return (opeartion.GetDetails(Path, Names));
                case "Search":
                    return (opeartion.Search(Path, ExtensionsAllow, SearchString, CaseSensitive));
            }
            return "";
        }  
    }
}
