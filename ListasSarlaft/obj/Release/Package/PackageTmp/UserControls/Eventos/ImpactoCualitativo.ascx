<%@ Control Language="C#" AutoEventWireup="true" CodeBehind="ImpactoCualitativo.ascx.cs" Inherits="ListasSarlaft.UserControls.Eventos.ImpactoCualitativo" %>
<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="asp" %>

<style type="text/css">
    .gridViewHeader a:link {
        text-decoration: none;
    }

    .style1 {
        width: 100%;
    }
</style>


<asp:UpdatePanel ID="UpdatePanel1" runat="server">
    <ContentTemplate>
        <asp:Panel ID="pnlMsgBox" runat="server" Width="400px" CssClass="modalPopup" Style="display: none;">
            <table width="100%">
                <tr class="topHandle">
                    <td colspan="2" align="center" runat="server" id="tdCaption">
                        <asp:Label ID="lblCaption" runat="server" Text="Atención"
                            Font-Bold="False" Font-Names="Tahoma" Font-Size="11px"></asp:Label><br />
                    </td>
                </tr>
                <tr>
                    <td style="width: 60px" valign="middle" align="center">
                        <asp:Image ID="imgInfo" runat="server"
                            ImageUrl="~/Imagenes/Icons/icontexto-webdev-about.png" />
                    </td>
                    <td valign="middle" align="left">
                        <asp:Label ID="lblMsgBox" runat="server" Font-Bold="False" Font-Names="Tahoma"
                            Font-Size="11px"></asp:Label>
                    </td>
                </tr>
                <tr>
                    <td align="right" colspan="2">
                         <asp:Button ID="btnImgokEliminar" runat="server" Visible="false" Text="Eliminar" OnClick="btnImgokEliminar_Click" />
                        <asp:Button ID="btnCancelar" runat="server" Text="Aceptar" OnClientClick="$find('mypopup').hide(); return false;" />
                    </td>
                </tr>
            </table>
        </asp:Panel>



         <asp:ModalPopupExtender ID="mpeMsgBox" runat="server" PopupControlID="pnlMsgBox" BehaviorID="mypopup"
            Enabled="True" TargetControlID="btndummy" BackgroundCssClass="modalBackground" DropShadow="true" >
        </asp:ModalPopupExtender>
        <asp:Button ID="btndummy" runat="server" Text="Button" Style="display: none" />






        <asp:Label ID="lblOk" runat="server" Text="Label" Visible="false"></asp:Label>
        <asp:Label ID="lblCancelar" runat="server" Text="Label" Visible="false"></asp:Label>
        <table align="center">
            <tr align="center" bgcolor="#333399">
                <td>
                    <asp:Label ID="Label13" runat="server" ForeColor="White" Text="Impacto cualitativo" Font-Bold="False"
                        Font-Names="Calibri" Font-Size="Large"></asp:Label>
                </td>
            </tr>
            <tr align="center" bgcolor="#EEEEEE" id="filaGrid" runat="server" visible="true">
                <td bgcolor="#EEEEEE">
                    <br />
                    <table>
                        <tr>
                            <td>
                                <asp:GridView ID="gvImpactoCualitativo" runat="server" CellPadding="4"
                                    ForeColor="#333333" AllowPaging="True" AllowSorting="True" AutoGenerateColumns="False"
                                    ShowHeaderWhenEmpty="True" OnRowCommand="gvImpactoCualitativo_RowCommand" OnPageIndexChanging="gvImpactoCualitativo_PageIndexChanging"
                                    OnSelectedIndexChanged="gvImpactoCualitativo_SelectedIndexChanged"
                                    HeaderStyle-CssClass="gridViewHeader" BorderStyle="Solid" GridLines="Vertical"
                                    DataKeyNames="idImpactoCualitativo"
                                    CssClass="Apariencia" Font-Bold="False">
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <Columns>
                                        <asp:BoundField DataField="idImpactoCualitativo" HeaderText="Id" InsertVisible="False" Visible="false" ReadOnly="True">
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Nombre" HeaderText="Nombre" ItemStyle-HorizontalAlign="Center"></asp:BoundField>
                                        <asp:BoundField DataField="IdUsuario" HeaderText="IdUsuario" ItemStyle-HorizontalAlign="Center" Visible="false"></asp:BoundField>
                                        <asp:BoundField DataField="FechaRegistro" HeaderText="FechaRegistro" ItemStyle-HorizontalAlign="Center" Visible="false"></asp:BoundField>
                                        <asp:TemplateField ShowHeader="False" HeaderText="Acción">
                                            <ItemTemplate>
                                                <asp:ImageButton ID="ImageButton1" runat="server" CausesValidation="False" CommandName="Select"
                                                    ImageUrl="~/Imagenes/Icons/edit.png" Text="Seleccionar" ToolTip="Editar"/>
                                                <asp:ImageButton ID="btnImgEliminar" runat="server" CausesValidation="False" OnClick="btnImgEliminar_Click" CommandArgument="<%# Container.DataItemIndex %>"
                                                    ImageUrl="~/Imagenes/Icons/delete.png" Text="Eliminar" CommandName="Eliminar" ToolTip="Eliminar"/>
                                            </ItemTemplate>
                                            <ItemStyle HorizontalAlign="Center" />
                                        </asp:TemplateField>
                                        <%--<asp:ButtonField ButtonType="Image" ImageUrl="~/Imagenes/Icons/edit.png" Text="Editar"
                                            CommandName="Modificar" />
                                        <asp:ButtonField ButtonType="Image" ImageUrl="~/Imagenes/Icons/delete.png" Text="Eliminar"
                                            CommandName="Eliminar" />--%>
                                    </Columns>
                                    <EditRowStyle BackColor="#999999" />
                                    <FooterStyle BackColor="White" Font-Bold="True" ForeColor="White" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                    <SortedAscendingCellStyle BackColor="#E9E7E2" />
                                    <SortedAscendingHeaderStyle BackColor="#506C8C" />
                                    <SortedDescendingCellStyle BackColor="#FFFDF8" />
                                    <SortedDescendingHeaderStyle BackColor="#6F8DAE" />
                                </asp:GridView>
                            </td>
                        </tr>
                        <tr>
                            <td align="center">
                                <asp:ImageButton ID="imgBtnInsertar" runat="server" CausesValidation="False" CommandName="Insert"
                                    ImageUrl="~/Imagenes/Icons/Add.png" Text="Insert"
                                    OnClick="imgBtnInsertar_Click" ToolTip="Insertar" />
                            </td>
                        </tr>
                    </table>
                    <br />
                </td>
            </tr>
            <tr align="left" id="filaDetalle" runat="server" visible="false">
                <td bgcolor="#EEEEEE">

                    <table width="50%" align="center">

                        <tr>
                            <td align="left" bgcolor="#BBBBBB">
                                <asp:Label ID="lblNombreImpacto" runat="server" Text="Nombre:" CssClass="Apariencia"></asp:Label>
                            </td>
                            <td>
                                <asp:TextBox ID="txtNombreImpacto" runat="server" Width="300px" CssClass="Apariencia" MaxLength="50"></asp:TextBox>

                            </td>
                        </tr>
                        <tr align="center">
                            <td colspan="2">
                                <table class="tablaSinBordes">
                                    <tr>
                                        <td>
                                            <asp:ImageButton ID="btnImgInsertar" runat="server" ImageUrl="~/Imagenes/Icons/guardar.png"
                                                Visible="False" OnClick="btnImgInsertar_Click" ToolTip="Guardar" ValidationGroup="iServ" />
                                        </td>
                                        <td>
                                            <asp:ImageButton ID="btnImgActualizar" runat="server" ImageUrl="~/Imagenes/Icons/guardar.png"
                                                Style="text-align: right" OnClick="btnImgActualizar_Click" ToolTip="Guardar" ValidationGroup="iServ" />
                                        </td>
                                        <td>
                                            <asp:ImageButton ID="btnImgCancelar" runat="server" ImageUrl="~/Imagenes/Icons/cancel.png"
                                                OnClick="btnImgCancelar_Click" ToolTip="Cancelar" />
                                        </td>
                                    </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>


        <%--<asp:ModalPopupExtender ID="mpeMsgBox" runat="server" TargetControlID="btndummy"
            PopupControlID="pnlMsgBox" OkControlID="btnAceptar" BackgroundCssClass="FondoAplicacion"
            Enabled="True" DropShadow="true">
        </asp:ModalPopupExtender>
        <asp:Button ID="btndummy" runat="server" Text="Button" Style="display: none" />--%>



    </ContentTemplate>

    <Triggers>
        <asp:AsyncPostBackTrigger ControlID="gvImpactoCualitativo" EventName="SelectedIndexChanged" />
    </Triggers>
</asp:UpdatePanel>
<%--The fnClick javascript function will force the btnOk button to post back. 
(to actually achieve this we need to add a little piece of code in 
the Page_Load method of the user control.--%>
<script type="text/javascript">
    function fnClick(sender, e) {
        __doPostBack(sender, e);
    }
</script>
