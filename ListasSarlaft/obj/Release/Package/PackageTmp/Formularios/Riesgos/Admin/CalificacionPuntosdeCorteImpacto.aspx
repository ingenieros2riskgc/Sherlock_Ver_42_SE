<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="CalificacionPuntosdeCorteImpacto.aspx.cs" MasterPageFile="~/MastersPages/Admin.Master" Inherits="ListasSarlaft.Formularios.Riesgos.Admin.CalificacionPuntosdeCorteImpacto" %>

<%@ OutputCache Location="None" %>


<%@ Register TagPrefix="T"  TagName="CalificacionPuntosdeCorteImpactos" Src="~/UserControls/Riesgos/CalificacionExpPuntosDeCorteImpacto.ascx"%>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server"></asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="ContentPlaceHolder2" runat="server">
    
</asp:Content>
<asp:Content ID="Content4" ContentPlaceHolderID="ContentPlaceHolder3" runat="server">
</asp:Content>
<asp:Content ID="Content5" ContentPlaceHolderID="ContentPlaceHolder4" runat="server">
</asp:Content>
<asp:Content ID="Content6" ContentPlaceHolderID="ContentPlaceHolder5" runat="server">
    <T:CalificacionPuntosdeCorteImpactos ID="CalificacionPuntosdeCorteImpactos" runat="server" />
</asp:Content>
<asp:Content ID="Content7" ContentPlaceHolderID="ContentPlaceHolder6" runat="server">
</asp:Content>
<asp:Content ID="Content8" ContentPlaceHolderID="ContentPlaceHolder7" runat="server">
</asp:Content>