<%@ Page Language="C#" MasterPageFile="~/Site.Master" Title="FileExplorer" AutoEventWireup="true" CodeBehind="FileExplorerFeatures.aspx.cs" Inherits="SyncfusionASPNETApplication3.FileExplorerFeatures" %>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
<h2>FileExplorer Features:</h2>
<br />
<li> RTL</li>
<li> Localization - en-US</li>
<li> API</li>
<li> Keyboard Interaction</li>
<li> Custom Tool</li>
<li> Theme - Flat-Azure</li>
<br/>
     <script src='<%= Page.ResolveClientUrl("~/Scripts/ej/i18n/ej.culture.en-US.min.js")%>' type="text/javascript"></script>
			<script src='<%= Page.ResolveClientUrl("~/Scripts/ej/l10n/ej.localetexts.en-US.min.js")%>' type="text/javascript"></script>
<div id = "ControlRegion">
<ej:FileExplorer ID="fileexplorer" runat="server" IsResponsive="true" Width="100%" AjaxAction="FileExplorerFeatures.aspx/FileActionDefault" Path="~/content/images/FileExplorer/" EnableRTL="true" Locale="en-US" Layout="Tile">
        <AjaxSettings>
            <Download Url="downloadFile.ashx{0}" />
            <Upload Url="uploadFiles.ashx{0}" />
        </AjaxSettings>
 <Tools>
            <CustomTool>
                <ej:FileExplorerCustomTool Name="Help" Tooltip="Help" Action="dialogOpen" Css="e-fileExplorer-toolbar-icon Help" />
            </CustomTool>
        </Tools>
    </ej:FileExplorer>
       <ej:Dialog ID="helpDialog" Title="FileExplorer Help" ShowOnInit="false" EnableModal="true" Width="350" EnableResize="false" runat="server">
        <DialogContent>
            <div class="text-content">
                <div class="header-content">Need assistance?</div>
                Our help document assists you to know more about FileExplorer control.<br /><br />
                Please refer -> <a href="http://help.syncfusion.com/web" target="_blank">Help Document</a>
            </div>
        </DialogContent>
    </ej:Dialog>
    <script type="text/javascript">
        function dialogOpen() {
            $('#<%=helpDialog.ClientID%>').ejDialog('open')
        }
    </script>
    <style type="text/css">
        .e-fileExplorer-toolbar-icon {
            height: 22px;
            width: 22px;
            font-family: 'ej-webfont';
            font-size: 18px;
            margin-top: 2px;
            text-align: center;
        }
        .e-fileExplorer-toolbar-icon.Help:before {
            content: "\e72b";
        }
        .e-dialog .header-content {
           font-size:16px;
           margin-top: .5em;
           margin-bottom: 1em;
        }
        .e-dialog>.e-titlebar {
            padding: .25em .25em .25em 1em;
        }
        .e-dialog.e-dialog-wrap {
            border: none;
        }
        .e-dialog .e-dialog-icon {
            right: 0;
        }
    </style>
 <script type="text/javascript" >
    $(function () {
           $(document).on("keydown", function (e) {
                if (e.altKey && e.keyCode === 74) { // j- key code.
                    $("#<%=fileexplorer.ClientID%>").find(".e-toolbar").focus();
                }
            });
    });
       </script>
            <h2>API</h2>
            <div id="sampleProperties">
        <div class="prop-grid jumbotron">
            <div class="row">
                <div class="col-md-3">
                    Toolbar
                </div>
                <div class="col-md-3">
                    <ej:ToggleButton ID="check1" runat="server" Width="105px" Size="Normal" ContentType="TextOnly" DefaultText="Hide" ActiveText="Show" ClientSideOnClick="Toolbar"></ej:ToggleButton>                  
                </div>
            </div>
            <div class="row">
                <div class="col-md-3">
                    Status Bar
                </div>
                <div class="col-md-3">
                    <ej:ToggleButton ID="check2" runat="server" Width="105px" Size="Normal" ContentType="TextOnly" DefaultText="Hide" ActiveText="Show" ClientSideOnClick="Statusbar"></ej:ToggleButton>                                     
                </div>
            </div>
            <div class="row">
                <div class="col-md-3">
                    Treeview
                </div>
                <div class="col-md-3">
                    <ej:ToggleButton ID="check3" runat="server" Width="105px" Size="Normal" ContentType="TextOnly" DefaultText="Hide" ActiveText="Show" ClientSideOnClick="Treeview"></ej:ToggleButton>                                                         
                </div>
            </div>
            <div class="row">
                <div class="col-md-3">
                    Destroy/Restore
                </div>
                <div class="col-md-3">
                    <ej:ToggleButton ID="check6" runat="server" Width="105px" Size="Normal" ContentType="TextOnly" DefaultText="Destroy" ActiveText="Restore" ClientSideOnClick="onDestoryRestore"></ej:ToggleButton>                                                                                                                     
                </div>
            </div>
            <div class="row">
                <div class="col-md-3">
                    Diable/Enable AddFolder
                </div>
                <div class="col-md-3">
                    <ej:ToggleButton ID="check7" runat="server" Width="105px" Size="Normal" ContentType="TextOnly" DefaultText="Disable" ActiveText="Enable" ClientSideOnClick="onDisableEnable"></ej:ToggleButton>                                                                                                                     
                </div>
            </div>
            <div class="row">
                <div class="col-md-3">
                    Get Current Path
                </div>
                <div class="col-md-3">
                    <ej:Button ID="getPath" runat="server" Type="Button" Width="105px" Text="Get Path" ClientSideOnClick="getCurrentPath"></ej:Button>                    
                </div>
            </div>
        </div>
    </div>
     <script type="text/javascript">
         var rte;
         $(function () {
             feObj = $("#<%=fileexplorer.ClientID%>").data("ejFileExplorer");
        });
        function Toolbar(args) {
    if (feObj)
        feObj.option("showToolbar", !args.isChecked);
}
function Statusbar(args) {
    if (feObj)
        feObj.option("showFooter", !args.isChecked);
}
function Treeview(args) {
    if (feObj)
        feObj.option("showNavigationPane", !args.isChecked);
}
function ContextMenu(args) {
    if (feObj)
        feObj.option("showContextMenu", !args.isChecked);
}
function onDestoryRestore(args) {
    if (args.isChecked) {
        feObj.destroy();
        feObj = null;
        if (!$("#MainContent_check6").hasClass("not-disable"))
            $("#MainContent_check6").addClass("not-disable");
        $("#sampleProperties .e-togglebutton.e-js").not(".not-disable").ejToggleButton("disable");
        $("#MainContent_getPath").ejButton("disable");       
    }
    else {
        var localServ = "FileExplorerFeatures.aspx/FileActionDefault";
        $("#MainContent_fileexplorer").ejFileExplorer({
            isResponsive: true,
            width: "100%",
            path: "~/content/images/FileExplorer/",
            ajaxAction: localServ,
            ajaxSettings: {
                upload: {
                    url: "uploadFiles.ashx{0}"
                },
                download: {
                    url: "downloadFile.ashx{0}"
                }
            }
        });
        feObj = $("#MainContent_fileexplorer").data("ejFileExplorer");
        $("#sampleProperties .e-togglebutton.e-js").not(".not-disable").ejToggleButton("enable");
        $("#MainContent_getPath").ejButton("enable");
    }
}
function getCurrentPath() {
    if (feObj)
        alert(feObj.option("selectedFolder"));
}
function onDisableEnable(args) {
    if (args.isChecked) {
        if (feObj){
            feObj.disableToolbarItem("NewFolder");
			feObj.disableMenuItem("NewFolder");
			}
    }
    else
        if (feObj){
            feObj.enableToolbarItem("NewFolder");
			feObj.enableMenuItem("NewFolder");
			}
}
    </script>
</div>
//FeatureScript###
</asp:Content>
