<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="HowToUse.aspx.vb" Inherits="_1_ClickPublish.WebForm2" %>
<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <h2 style="padding-left: 10px; color: Black;">
    Searching for a PO and Printing Labels
    </h2>
    <p>
    1) Type a valid <b>PO Number</b> into the Text Box in the upper left hand corner of the <b>PO Lookup</b> page  and click search.<br />
    2) Make sure the <b>Sender Name</b> displayed in the center above the results table is correct.<br />
    3) Edit the <b>Quantity</b> and <b>Units</b> columns if necessary.<br />
    4) Edit the <b>Small Labels</b> and <b>Large Labels</b> columns with the desired number of labels.<br />
    5) Select a <b>Printer</b> from the drop down list in the upper right hand corner.<br />
    6) Click the <b>Print</b> button.<br />
    </p>

    <h2 style="padding-left: 10px; color: Black;">
    Requirements
    </h2>
    <p>
    1) Use with DYMO LabelWriter 450 Twin Turbo <b>ONLY</b>.<br />
    2) Use <b>ULine S-8503</b> labels on the left roller and <b>ULine S-8505</b> labels on the right roller.<br />
    3) Report any bugs to the <a href="mailto:helpdesk.us@ipgphotonics.com&subject=Receiving Dept. Label Generator Bug Report"><b>IT Helpdesk</b></a> x2550<br />
    </p>

    <h2 style="padding-left: 10px; color: Black;">
    Label Examples
    </h2><br />
    <table>
        <tr>
            <td style="text-align: center;">S-8503 (Left Roll)</td>
            <td style="text-align: center;">S-8505 (Right Roll)</td>
        </tr>
        <tr>
            <td style="padding: 0px 30px 30px 30px; vertical-align: top;"><img src="SmallRecvLabel_small.jpg" alt="Small Receiving Label"/></td>
            <td style="padding: 0px 30px 30px 30px; vertical-align: top;"><img src="LargeRecvLabel_small.jpg" alt="Small Receiving Label"/></td>
        </tr>
    </table>
</asp:Content>
