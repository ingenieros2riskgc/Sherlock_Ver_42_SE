<%@ Page Title="Sherlock" Language="C#" MasterPageFile="~/MastersPages/Admin.Master" AutoEventWireup="true" CodeBehind="ImpactoCualitativo.aspx.cs" Inherits="ListasSarlaft.Formularios.Eventos.ImpactoCualitativo" %>
<%@ OutputCache Location="None" %>
<%@ Register TagPrefix="ICT" TagName="ImpactoCualitativoo" Src="~/UserControls/Eventos/ImpactoCualitativo.ascx" %>

<asp:Content ID="Content1" ContentPlaceHolderID="head" runat="server">
</asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="ContentPlaceHolder1" runat="server">
    <ICT:ImpactoCualitativoo ID="ImpactoCualitativoo" runat="server"/>
</asp:Content>

