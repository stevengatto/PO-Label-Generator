<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Default.aspx.vb" Inherits="_1_ClickPublish.WebForm1" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>

<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">

    <h2> Purchase Order Lookup </h2>
   
    <asp:textbox id="POsearch" runat="server"/>
    <asp:Button id="PObutton" runat="server" Text="Search"/>
    <asp:TextBox id="tbWarning" runat="server" Text="" ReadOnly="true" BorderColor="Transparent" Width="400px" style="color: Red"/>
    <asp:Button id="printbutton" runat="server" Text="Print" Visible="false" style="float: right" />
    <asp:DropDownList id="dropdownlist" runat="server"  Visible="false" style="float: right"/>
    <br />

    <table runat="server" style="width:100%; border:2px; padding-top:5px">
        <tr>
            <td id="tableCol1" runat="server" style="color: Black; text-align:center"></td>
        </tr>
        <tr>
            <td id="tableCol2" runat="server" style="color: Black; text-align:center"></td>
        </tr>
    </table>

    <br />
    <asp:GridView ID="GridView1" Width="100%" runat="server"   
        AutoGenerateColumns="False"  
        GridLines="None"  
        CssClass="mGrid"  
        PagerStyle-CssClass="pgr"  
        AlternatingRowStyle-CssClass="alt"
        EmptyDataText="No matches found"
        >
        <Columns>
            <asp:BoundField DataField="No_" HeaderText="Part No." 
                SortExpression="No_" ItemStyle-Width="15%"/>
            <asp:BoundField DataField="Description" HeaderText="Description" 
                SortExpression="Description" ItemStyle-Width="36%"/>
            <asp:TemplateField HeaderText="Quantity" SortExpression="Quantity" 
                ItemStyle-HorizontalAlign="Center" ItemStyle-Width="8%">
                <ItemTemplate>
                    <asp:TextBox id="quantityTextBox" runat="server" Width="100%"/>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Units" SortExpression="Units" ItemStyle-Width="7%">
                <ItemTemplate>
                    <asp:TextBox id="unitsTextBox" runat="server" Width="100%"/>
                </ItemTemplate>
            </asp:TemplateField>
            <asp:BoundField DataField="Location Code" HeaderText="Location Code" 
                SortExpression="Location Code" ItemStyle-Width="8%"/>
            <asp:BoundField DataField="Purchaser Code" HeaderText="Purchaser" 
                SortExpression="Purchaser Code" ItemStyle-Width="14%"/>
            <asp:TemplateField HeaderText="Small Labels" SortExpression="Small Labels" 
                ItemStyle-HorizontalAlign="Right" ItemStyle-Width="6%">
                <ItemTemplate>
                    <asp:TextBox ID="smallLabelsTextBox" runat="server" Text="0" Width="100%" style="text-align: right" />
                </ItemTemplate>
            </asp:TemplateField>
            <asp:TemplateField HeaderText="Large Labels" SortExpression="Large Labels" 
                ItemStyle-HorizontalAlign="Right" ItemStyle-Width="6%">
                <ItemTemplate>
                    <asp:TextBox ID="largeLabelsTextBox" runat="server" Text="0" Width="100%" style="text-align: right" />
                </ItemTemplate>
            </asp:TemplateField>
        </Columns>
    </asp:GridView>
</asp:Content>
