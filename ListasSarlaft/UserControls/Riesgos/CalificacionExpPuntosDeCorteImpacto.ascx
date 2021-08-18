<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="CalificacionExpPuntosDeCorteImpacto.ascx.cs" Inherits="ListasSarlaft.UserControls.Riesgos.CalificacionExpPuntosDeCorteImpacto" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>
<%@ Register Src="~/UserControls/Sitio/OkMessageBox.ascx" TagPrefix="uc" TagName="OkMessageBox" %>

<script type="text/javascript">
    function fnClick(sender, e) {
        __doPostBack(sender, e);
    }
</script>
<script src="/JS/jquery.min.js"></script>
<script src="/JS/jsFunciones.js"></script>

<style type="text/css">
    .gridViewHeader a:link {
        text-decoration: none;
    }

    .FondoAplicacion {
        background-color: Gray;
        filter: alpha(opacity=80);
        opacity: 0.8;
    }

    .scrollingControlContainer {
        overflow-x: hidden;
        overflow-y: scroll;
    }

    .scrollingCheckBoxList {
        border: 1px #808080 solid;
        margin: 10px 10px 10px 10px;
        height: 200px;
    }

    .popup {
        border: Silver 1px solid;
        color: #060F40;
        background: #ffffff;
    }

    .style1 {
        height: 74px;
    }

    .autocomplete_completionListElement {
        margin: 0px !important;
        background-color: inherit;
        color: windowtext;
        border: buttonshadow;
        border-width: 1px;
        cursor: 'default';
        overflow: auto;
        height: 200px;
        text-align: left;
        list-style-type: none;
    }
    /* AutoComplete highlighted item */
    .autocomplete_highlightedListItem {
        background-color: #ffff99;
        color: black;
        padding: 1px;
    }
    /* AutoComplete item */
    .autocomplete_listItem {
        background-color: window;
        color: windowtext;
        padding: 1px;
    }

    .FiltroPlanes {
        text-align: center;
        width: 29%;
        margin-left: 36%;
        border: solid 1px #999999;
        margin-top: 22px;
        margin-bottom: 10px;
    }

    .EtiquetaCarga, .ArchivosCargados, .Justificacion, .jCambios, .pRiesgosAsociados, .sAsociados, .pkJustificacion {
        margin: 18px 0 7px 0;
        font-family: Calibri;
        font-size: Small;
        color: #4d4d4d;
    }

    .EtiquetaCarga, .RAsociados, .EAsociados {
        border-radius: 0px 55px 57px 0px;
        -moz-border-radius: 0px 55px 57px 0px;
        -webkit-border-radius: 0px 55px 57px 0px;
        border: 0px solid #000000;
        background: rgba(227,227,227,1);
        background: -moz-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -webkit-gradient(left top, right top, color-stop(0%, rgba(227,227,227,1)), color-stop(0%, rgba(166,166,166,1)), color-stop(43%, rgba(178,178,178,0.58)), color-stop(73%, rgba(186,186,186,0.58)), color-stop(100%, rgba(201,201,201,0.58)));
        background: -webkit-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -o-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -ms-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: linear-gradient(to right, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(255, 255, 255, 0.45) 100%);
        filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e3e3e3', endColorstr='#c9c9c9', GradientType=1 );
    }

    .pJustificacion {
        border-radius: 0px 55px 57px 0px;
        -moz-border-radius: 0px 55px 57px 0px;
        -webkit-border-radius: 0px 55px 57px 0px;
        border: 0px solid #000000;
        background: rgba(227,227,227,1);
        background: -moz-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -webkit-gradient(left top, right top, color-stop(0%, rgba(227,227,227,1)), color-stop(0%, rgba(166,166,166,1)), color-stop(43%, rgba(178,178,178,0.58)), color-stop(73%, rgba(186,186,186,0.58)), color-stop(100%, rgba(201,201,201,0.58)));
        background: -webkit-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -o-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -ms-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: linear-gradient(to right, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.31) 100%);
        filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e3e3e3', endColorstr='#c9c9c9', GradientType=1 );
    }

    .pRiesgosAsociados, .sAsociados {
        border-radius: 0px 55px 57px 0px;
        -moz-border-radius: 0px 55px 57px 0px;
        -webkit-border-radius: 0px 55px 57px 0px;
        border: 0px solid #000000;
        background: rgba(227,227,227,1);
        background: -moz-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -webkit-gradient(left top, right top, color-stop(0%, rgba(227,227,227,1)), color-stop(0%, rgba(166,166,166,1)), color-stop(43%, rgba(178,178,178,0.58)), color-stop(73%, rgba(186,186,186,0.58)), color-stop(100%, rgba(201,201,201,0.58)));
        background: -webkit-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -o-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: -ms-linear-gradient(left, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0.58) 43%, rgba(186,186,186,0.58) 73%, rgba(201,201,201,0.58) 100%);
        background: linear-gradient(to right, rgba(227,227,227,1) 0%, rgba(166,166,166,1) 0%, rgba(178,178,178,0) 43%, rgba(186,186,186,0.21) 73%, rgba(201,201,201,0.31) 100%);
        filter: progid:DXImageTransform.Microsoft.gradient( startColorstr='#e3e3e3', endColorstr='#c9c9c9', GradientType=1 );
    }
</style>

<%--Modal Editar Puntos de Corte--%>
<asp:ModalPopupExtender ID="ModalEditarPuntosCorte" runat="server" TargetControlID="btndummy4"
    PopupControlID="PCEditarPuntosCorte" OkControlID="CancelarC" BackgroundCssClass="FondoAplicacion"
    Enabled="True" DropShadow="true">
</asp:ModalPopupExtender>
<asp:Button ID="btndummy4" runat="server" Text="Button1" Style="display: none" />
<asp:Panel ID="PCEditarPuntosCorte" runat="server" Width="400px" Style="display: none;" BorderColor="#575757"
    BackColor="#FFFFFF" BorderStyle="Solid">
    <table width="100%">
        <tr class="topHandle" style="background-color: #5D7B9D">
            <td colspan="2" align="center" runat="server" id="td3">&nbsp;
                        <asp:Label ID="Label9" runat="server" Text="Atención" Font-Names="Calibri" Font-Size="Small"></asp:Label><br />
            </td>
        </tr>
        <tr>
            <td style="width: 60px" valign="middle" align="center">
                <asp:Image ID="Image2" runat="server" ImageUrl="~/Imagenes/Icons/icontexto-webdev-about.png" />
            </td>
            <td valign="middle" align="left">
                <div style="text-align: center; margin-top: 2%; margin-right: 25%;">
                    <asp:Label ID="Label20" runat="server" Text="Nombre:" Font-Names="Calibri" Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="ModalNombreFrecuencia" runat="server" MaxLength="100" Width="77%" autofocus="true" placeholder="Nombre Frecuencia" Enabled="false"></asp:TextBox>
                </div>
                <div style="text-align: center; margin-top: 2%; margin-right: 40%;">
                    <asp:Label ID="Label18" runat="server" Text="Min (%)" Font-Names="Calibri" Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="ModalMin" runat="server" Width="57%" MaxLength="3" autocomplete="off" placeholder="Min"
                        onkeyup="SoloNumerosMaxCien(this);" onchange="SoloNumerosMaxCien(this);"></asp:TextBox>
                </div>
                <div style="text-align: center; margin-top: 2%; margin-right: 40%;">
                    <asp:Label ID="Label1" runat="server" Text="Max (%)" Font-Names="Calibri" Font-Size="Small"></asp:Label>
                    <asp:TextBox ID="ModalMax" runat="server" Width="57%" MaxLength="3" autocomplete="off" placeholder="Max"
                        onkeyup="SoloNumerosMaxCien(this);" onchange="SoloNumerosMaxCien(this);"></asp:TextBox>
                </div>
            </td>
        </tr>
        <tr align="right">
            <td align="right" colspan="2">
                <asp:Button ID="Actualizar" runat="server" Font-Names="Calibri" Font-Size="Small" Text="Actualizar" OnClick="Actualizar_Click"  />
                <asp:Button ID="CancelarC" runat="server" Font-Names="Calibri" Font-Size="Small" Text="Cancelar" />
            </td>
        </tr>
    </table>
</asp:Panel>

<asp:UpdatePanel ID="Tbody" runat="server">
    <ContentTemplate>
        <uc:OkMessageBox ID="omb" runat="server" />
        <div style="text-align: -webkit-center;">
            <table style="width: 30%;">
                <tr align="center" bgcolor="#333399">
                    <td>
                        <asp:Label ID="Label140" runat="server" Font-Bold="False" Font-Names="Calibri" Font-Size="Large"
                            ForeColor="White" Text="Puntos de Corte - Calificación Cuantitativa - Impacto"></asp:Label>
                    </td>
                </tr>
            </table>
             <table>
                <tr>
                    <td>
                        <asp:GridView ID="GvPuntosCorte" runat="server" AutoGenerateColumns="False" CellPadding="4"
                            ForeColor="#333333" GridLines="Vertical" ShowHeaderWhenEmpty="True" HeaderStyle-CssClass="gridViewHeader"
                            BorderStyle="Solid" HorizontalAlign="Justify" Font-Names="Calibri" Font-Size="Small" AllowPaging="True"
                            DataKeyNames="IdPuntoCorte, IdFrecuenciaEventos, NombreFrecuencia, Min, Max" Width="414px"  OnRowCommand="GvPuntosCorte_RowCommand"
                            OnPageIndexChanging="GvPuntosCorte_PageIndexChanging" >
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <Columns>
                                <asp:BoundField HeaderText="Id PuntoCorte" DataField="IdPuntoCorte" Visible="false" />
                                <asp:BoundField HeaderText="Valor" DataField="IdFrecuenciaEventos"  />
                                <asp:BoundField HeaderText="Nivel" DataField="NombreFrecuencia" />
                                <asp:BoundField HeaderText="Min (%)" DataField="Min" />
                                <asp:BoundField HeaderText="Max (%)" DataField="Max" />
                                <asp:ButtonField HeaderText="Editar" ButtonType="Image" ImageUrl="~/Imagenes/Icons/edit.png" Text="Editar"
                                    CommandName="Editar" />
                            </Columns>
                            <EditRowStyle BackColor="#999999" />
                            <FooterStyle BackColor="White" Font-Bold="True" ForeColor="White" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" HorizontalAlign="Center" />
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <SortedAscendingCellStyle BackColor="#E9E7E2" />
                            <SortedAscendingHeaderStyle BackColor="#506C8C" />
                            <SortedDescendingCellStyle BackColor="#FFFDF8" />
                            <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                        </asp:GridView>
                    </td>
                </tr>
            </table>
        </div>
    </ContentTemplate>
</asp:UpdatePanel>
