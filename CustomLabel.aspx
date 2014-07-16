<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="CustomLabel.aspx.vb" Inherits="_1_ClickPublish.WebForm3" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

    <asp:DropDownList id="dropdownlist" runat="server"/>
    <asp:Button id="printbutton" runat="server" Text="Print" /><br />
    <asp:TextBox id="tbWarning" runat="server" Text="" ReadOnly="true" BorderColor="Transparent" Width="400px" style="color: Red"/><br />

    <table>
        <tr>
            <td>
                Label Type
            </td>
            <td style="padding-left: 20px;">
                Copies
            </td>
        </tr>
        <tr>
            <td>
                <asp:DropDownList id="dropdownlistlabel" runat="server"/>
            </td>
            <td style="padding-left: 20px;">
                <asp:TextBox id="copies" runat="server" Text="1" Width="30px"/>
            </td>
        </tr>
    </table><br />
    
    <textarea id="textarea" rows="5" cols="40" maxlength="125" runat="server" draggable="false" style="text-align: center;"></textarea>
    <br />
    <br />
    <i>Important Notes:</i> <br />
    <ul>
        <li><i>A separate line will <b>only</b> be created when you press <b>Enter</b>.</i></li>
        <li><i><b>Do not</b> copy and paste into the <b> Text Area</b>.</i></li>
        <li><i>If a line overflows onto the next line, it will appear as <b>one line</b> on the label.</i></li>
    </ul>
</asp:Content>

