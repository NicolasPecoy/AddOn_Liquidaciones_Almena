using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAPbouiCOM;
using SAPbobsCOM;
using System.Data;
using System.Xml.Serialization;
using System.Xml;
using System.IO;
using System.Web;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.Management;
using Microsoft.Win32;
using FactElectronicaSICFE;
using System.Runtime.InteropServices;
using System.Threading;
using System.Globalization;

namespace AddOn_Liquidaciones_Almena
{
    public class SboClass
    {
        #region Global Definitions
        private SAPbouiCOM.Application SBO_Application;
        private SAPbobsCOM.Company oCompany;
        private String sPath;
        char c = Convert.ToChar(92);
        private String empresa = "ALMEN";
        private String tipoArchivo = "";
        private bool guardaLog = true;
        private bool actualizaCentroCostos = false; // Indica si actualiza los Centros de Costos de los articulos
        private String usuarioLogueado = ""; // Usuario logueado
        private int idUsuarioLogueado = 0; // Id del Usuario logueado
        private int idSucursalUsuario = 0; // Id de la Sucursal correspondiente al Usuario logueado
        private bool esSuperUsuario = false; // Indica si el usuario logueado es superusuario
        private bool visualizaSoloLiquidaciones = false; // Indica si el usuario logueado ve el addOn Completo o solo la parte de Ingreso de Liquidaciones
        SAPbouiCOM.Form oFormDatosPedido;
        SAPbouiCOM.Form oFormDatosPedidoVisor;
        SAPbouiCOM.Form oFormDatosIngresoCobros;
        SAPbouiCOM.Form oFormDatosCheques;
        SAPbouiCOM.Form oFormDatosIngresoLiquidaciones;

        #endregion

        #region Estructura SBO
        private void init()
        {
            try
            {
                try
                {
                    SetApplication();
                    getUsuarioLogueado();
                }
                catch (Exception ex)
                {
                    //System.Windows.Forms.MessageBox.Show("SBO no se encontro cargado!.");
                }
                try
                {
                    SBO_Application.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(this.SBO_Application_MenuEvent);
                    SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                    SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                }
                catch
                {
                    //System.Windows.Forms.MessageBox.Show("Error al cargar eventos.");
                }
                try
                {
                    if (SetConnectionContext() != 0)
                    {
                        //SBO_Application.MessageBox("Error al Conectar con DI API");
                    }
                }
                catch (Exception e)
                {
                    //System.Windows.Forms.MessageBox.Show("Error al conectar con DI API: " + e.Message);
                }
                try
                {
                    oCompany = SBO_Application.Company.GetDICompany();
                }
                catch
                {
                    //SBO_Application.MessageBox("Error al intentar conectar con Base de Datos.");
                }
                try
                {
                    obtenerDatosConfiguracion();
                    getIdUsuarioIdSucursal(); // Busco el id de Sucursal y el id de Usuario
                    AddMenuItems();
                }
                catch
                {
                    //SBO_Application.MessageBox("Error al agregar menu items.");
                }
            }
            catch (Exception ex)
            {
                //System.Windows.Forms.MessageBox.Show("SBO no se encontro cargado!.");
            }
        }

        public SboClass()
        {
            //matarProcesos();
            init();
        }

        private void SetApplication()
        {
            SAPbouiCOM.SboGuiApi SboGuiApi;
            String sConnectionString;
            SboGuiApi = new SAPbouiCOM.SboGuiApi();
            sConnectionString = Environment.GetCommandLineArgs().GetValue(1).ToString(); // 1 Estaba antes el uno pero no lo traia como param
            SboGuiApi.Connect(sConnectionString);
            SBO_Application = SboGuiApi.GetApplication();
        }

        private int SetConnectionContext()
        {

            String sCookie;
            String sConnectionContext;
            oCompany = new SAPbobsCOM.Company();
            sCookie = oCompany.GetContextCookie();
            sConnectionContext = SBO_Application.Company.GetConnectionContext(sCookie);
            if (oCompany.Connected == true)
            {
                oCompany.Disconnect();
            }
            return oCompany.SetSboLoginContext(sConnectionContext);
        }

        private String LoadFromXML(String FileName)
        {
            System.Xml.XmlDocument oXmlDoc;
            String sPath;

            oXmlDoc = new System.Xml.XmlDocument();
            sPath = System.Windows.Forms.Application.StartupPath + c;

            oXmlDoc.Load(sPath + FileName);
            return (oXmlDoc.InnerXml);
        }
        #endregion

        #region Eventos

        //Menu Events
        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (pVal.BeforeAction == false)
                {
                    if (pVal.MenuUID.Equals("Liquidaciones"))
                    {
                        CargarFormulario();
                    }
                    if (pVal.MenuUID.Equals("Pedidos"))
                    {
                        CargarFormularioPedidosVisor();
                    }
                    if (pVal.MenuUID.Equals("Cobros"))
                    {
                        CargarFormularioIngresoCobros();
                    }
                    if (pVal.MenuUID.Equals("Cheques"))
                    {
                        CargarFormularioCheques();
                    }
                    if (pVal.MenuUID.Equals("Nuevas Liquidaciones"))
                    {
                        CargarFormularioIngresoLiquidaciones();
                    }
                }
            }
            catch
            {
                BubbleEvent = false;
            }
            BubbleEvent = true;
        }

        //Items Events (Every item click is detected here)
        private void SBO_Application_ItemEvent(String FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            try
            {
                if (!pVal.BeforeAction)
                {
                    #region "InicializaAddOn"
                    try
                    {
                        // Entra aca cuando se da enter en la pantalla de Bloqueo
                        if (pVal.FormTypeEx.Equals("821") && pVal.ItemUID.Equals("1") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
                        {
                            init();
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "Inicializa AddOn", ""); // Guarda log del Proceso
                        }
                    }
                    catch (Exception ex)
                    { }
                    #endregion

                    #region "Ingresa Liquidaciones"
                    // Agrega una nueva Liquidacion F"
                    if (pVal.ItemUID.Equals("500") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("IngresoLiquidaciones", 0);

                            SAPbouiCOM.ComboBox oStaticCombo;
                            SAPbouiCOM.EditText oStaticTexto;

                            oStaticCombo = oFormDatosIngresoLiquidaciones.Items.Item("5").Specific;
                            string nombreRepartidor = ""; string codeRepartidor = "";
                            DateTime fecha = DateTime.Now;
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                            {
                                codeRepartidor = oStaticCombo.Selected.Value.ToString();
                                nombreRepartidor = oStaticCombo.Selected.Description.ToString();
                            }

                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea agregar una nueva Liquidacion?", 2, "Aceptar", "Cancelar");

                            if (respuestaMge == 1)
                            {
                                if (!String.IsNullOrEmpty(codeRepartidor.ToString()))
                                {
                                    oStaticTexto = oFormDatosIngresoLiquidaciones.Items.Item("6").Specific; // Fecha
                                    if (!String.IsNullOrEmpty(oStaticTexto.String))
                                        fecha = Convert.ToDateTime(oStaticTexto.String);

                                    int code = obtenerCodeProxLiquidacion();
                                    int docEntry = obtenerDocEntryLiquidaciones();
                                    string name = fecha.ToString("yyyyMMdd") + " - " + nombreRepartidor;
                                    //DateTime fechaCierre = Convert.ToDateTime("01-01-1980");

                                    SAPbobsCOM.Recordset oRSMyTable = null;
                                    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    String query = "INSERT into [@LIQUIDACIONES] (Code,Name,DocEntry,CreateDate,U_FECHA, U_ESTADO, U_REPARTIDOR) values ('" + // U_FECHA_CIERRE
                                    code + "','" + name + "','" + docEntry + "','" + DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss") + "','" + fecha.ToString("yyyy-MM-dd") + "','1','" + codeRepartidor + "')"; // ,'" + fechaCierre.ToString("yyyy-MM-dd") + "'

                                    oRSMyTable.DoQuery(query);

                                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                                    oRSMyTable = null;

                                    oStaticTexto = oFormDatosIngresoLiquidaciones.Items.Item("3").Specific; // Code proxima liquidacion
                                    oStaticTexto.Value = obtenerCodeProxLiquidacion().ToString();

                                    SBO_Application.MessageBox("Usted ha creado la Liquidación número " + code + " para la fecha " + fecha.ToShortDateString() + " correspondiente al repartidor " + nombreRepartidor.ToString());
                                    //CargarMatrixIngresoLiquidaciones();
                                }
                                else
                                    SBO_Application.MessageBox("Debe seleccionar un repartidor");
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al asignar Liquidacion al Registro", ex.Message.ToString()); // Guarda log del Proceso
                        }
                    }

                    if (pVal.ItemUID.Equals("16") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoLiquidaciones"))
                    {
                        try
                        {
                            CargarMatrixIngresoLiquidaciones();
                        }
                        catch (Exception ex)
                        { }
                    }

                    // Cierra Liquidacion seleccionada
                    if (pVal.ItemUID.Equals("17") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("IngresoLiquidaciones", 0);
                            oStatic = oForm.Items.Item("100").Specific;
                            int row = oStatic.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                            if (row == -1)
                                SBO_Application.MessageBox("Debe seleccionar una fila");
                            else
                            {
                                bool cambioAlgunRegistro = false; // Variable para saber si hubo algún cambio 
                                while (row != -1)
                                {
                                    SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_4", row); // Número Liquidacion
                                    string numeroLiquidacion = ed.Value;
                                    ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_0", row); // Estado
                                    string estadoLiquidacion = ed.Value;

                                    int respuestaMge = SBO_Application.MessageBox("Está seguro que desea cerrar la Liquidación " + numeroLiquidacion + "?", 2, "Aceptar", "Cancelar");

                                    if (respuestaMge == 1)
                                    {
                                        if (estadoLiquidacion.ToString().Equals("Abierta")) // Si está abierta
                                        {
                                            DateTime fechaCierre = DateTime.Now; // Por defecto es la fecha del día
                                            SAPbouiCOM.EditText oStaticTexto;
                                            oStaticTexto = oForm.Items.Item("19").Specific; // Fecha de Cierre

                                            if (!String.IsNullOrEmpty(oStaticTexto.String) || esSuperUsuario == false)
                                            {
                                                if (esSuperUsuario == false) // Si no es un Super Usuario
                                                    fechaCierre = DateTime.Now;
                                                else
                                                    fechaCierre = Convert.ToDateTime(oStaticTexto.String);

                                                clsLiquidacion liquidacionDatos = obtenerDatosLiquidacion(numeroLiquidacion); // Obtengo fecha de cierre si la misma ya había sido cerrada alguna vez

                                                bool accionValida = true;

                                                //if (liquidacionDatos.FechaCierre != fechaUno.Date && liquidacionDatos.FechaCierre != fechaDos.Date) // Entonces quiere decir que la Liq ya tiene una fecha de Cierre
                                                if (liquidacionDatos.FechaCierre.Year > 1980) // Entonces quiere decir que la Liq ya tiene una fecha de Cierre
                                                {
                                                    //fechaCierre = liquidacionDatos.FechaCierre; // Por lo tanto le cargo la fecha anterior
                                                    int respuestaMge2 = SBO_Application.MessageBox("La Liquidación ya tiene una fecha de cierre correspondiente al día " + liquidacionDatos.FechaCierre.ToShortDateString() + ", está seguro que desea continuar y cerrar con la fecha " + fechaCierre.ToShortDateString() + " ?", 2, "Aceptar", "Cancelar");
                                                    if (respuestaMge2 != 1)
                                                        accionValida = false;
                                                }

                                                if (accionValida == true)
                                                {

                                                    bool res = cambiarEstadoLiquidacion(numeroLiquidacion, "2", fechaCierre);

                                                    if (res == false)
                                                        SBO_Application.MessageBox("No es posible cambiar el estado de ésta Liquidación");
                                                    else
                                                    {
                                                        if (cambioAlgunRegistro == false)  // Si cambio el estado correctamente entonces recargo la grilla
                                                            cambioAlgunRegistro = true;
                                                    }
                                                }
                                            }
                                            else
                                                SBO_Application.MessageBox("Debe ingresar la fecha de cierre para la Liquidación");
                                        }
                                        else
                                            SBO_Application.MessageBox("La Liquidación seleccionada ya se encuentra cerrada");

                                        row = oStatic.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                    else
                                        break;
                                }

                                if (cambioAlgunRegistro == true) // Si cambio el estado correctamente entonces recargo la grilla
                                    CargarMatrixIngresoLiquidaciones();
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al Cerrar Liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                        }
                    }

                    // Re abre una Liquidacion seleccionada
                    if (pVal.ItemUID.Equals("18") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("IngresoLiquidaciones", 0);
                            oStatic = oForm.Items.Item("100").Specific;
                            int row = oStatic.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                            if (row == -1)
                                SBO_Application.MessageBox("Debe seleccionar una fila");
                            else
                            {
                                while (row != -1)
                                {
                                    SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_4", row); // Número Liquidacion
                                    string numeroLiquidacion = ed.Value;
                                    ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_0", row); // Estado
                                    string estadoLiquidacion = ed.Value;

                                    int respuestaMge = SBO_Application.MessageBox("Está seguro que desea re-abrir la Liquidación " + numeroLiquidacion + "?", 2, "Aceptar", "Cancelar");

                                    if (respuestaMge == 1)
                                    {
                                        if (estadoLiquidacion.ToString().Equals("Cerrada")) // Si está cerrada
                                        {
                                            if (esSuperUsuario == true) // Si no es un Super Usuario
                                            {
                                                bool res = cambiarEstadoLiquidacion(numeroLiquidacion, "1", Convert.ToDateTime("01-01-1980"));

                                                if (res == false)
                                                    SBO_Application.MessageBox("No es posible cambiar el estado de ésta Liquidación");
                                            }
                                            else
                                                SBO_Application.MessageBox("Su usuario no tiene permitido re-abrir una Liquidación");
                                        }
                                        else
                                            SBO_Application.MessageBox("La Liquidación seleccionada ya se encuentra abierta");

                                        row = oStatic.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }
                                    else
                                        break;
                                }

                                CargarMatrixIngresoLiquidaciones();
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al Cerrar Liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                        }
                    }
                    #endregion

                    #region "AsignarLiquidaciones"
                    // Asigna nuevo numero de Liquidacion al documento seleccionado
                    if (pVal.ItemUID.Equals("7") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarLiquidaciones", 0);
                            oStatic = oForm.Items.Item("2").Specific;

                            SAPbouiCOM.ComboBox oStaticCombo;
                            oStaticCombo = oFormDatosPedido.Items.Item("1000001").Specific; // Nro de liquidación nuevo
                            string nuevoNroLiquidacion = "";
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                nuevoNroLiquidacion = oStaticCombo.Selected.Value.ToString();

                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea asignar el número " + nuevoNroLiquidacion + " de liquidación a los documentos seleccionados?", 2, "Aceptar", "Cancelar");

                            if (respuestaMge == 1)
                            {
                                DateTime fechaLiquidacion = obtenerFechaLiquidacion(nuevoNroLiquidacion);
                                if (String.IsNullOrEmpty(nuevoNroLiquidacion))
                                    fechaLiquidacion = Convert.ToDateTime("01-01-1980");

                                int row = oStatic.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                if (row == -1)
                                    SBO_Application.MessageBox("Debe seleccionar una fila");
                                else
                                {
                                    while (row != -1)
                                    {
                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row); // DocNum
                                        clsDocumento docSeleccionado = new clsDocumento();
                                        docSeleccionado.DocNum = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row); // DocEntry
                                        docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row); // Tipo
                                        docSeleccionado.Tipo = ed.Value;

                                        bool res = asignarLiquidacion(docSeleccionado, nuevoNroLiquidacion, fechaLiquidacion);

                                        if (res == false)
                                            SBO_Application.MessageBox("Ha ocurrido un error al asignar el número de liquidación al documento " + docSeleccionado.Tipo + " " + docSeleccionado.DocNum);

                                        row = oStatic.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                    }

                                    // Valor por defecto al combo
                                    SAPbouiCOM.ComboBox combo = null;
                                    combo = oForm.Items.Item("1000001").Specific;
                                    combo.Select("No filtrar", BoSearchKey.psk_ByDescription);
                                    CargarGrilla(); // Recargo la grilla
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al asignar Liquidacion al Registro", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }

                    // // Asigna nuevo numero de Liquidacion a todos los documentos de la grilla cuando presiona clic en "Procesar Todos"
                    if (pVal.ItemUID.Equals("15") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Button oStaticText;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarLiquidaciones", 0);
                            oStatic = oForm.Items.Item("2").Specific;
                            SAPbouiCOM.ComboBox oStaticCombo;
                            oStaticCombo = oFormDatosPedido.Items.Item("1000001").Specific; // Nro de liquidación nuevo
                            string nuevoNroLiquidacion = "";
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                nuevoNroLiquidacion = oStaticCombo.Selected.Value.ToString();
                            oStaticText = oForm.Items.Item("15").Specific;

                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea asignar el número " + nuevoNroLiquidacion + " de liquidación a todos los documentos?", 2, "Aceptar", "Cancelar");

                            if (respuestaMge == 1)
                            {
                                DateTime fechaLiquidacion = obtenerFechaLiquidacion(nuevoNroLiquidacion);
                                if (String.IsNullOrEmpty(nuevoNroLiquidacion))
                                    fechaLiquidacion = Convert.ToDateTime("01-01-1980");

                                int cantRows = oStatic.RowCount; // Saco la cantidad de registros que tiene la Grilla
                                int row = 0;

                                //if (!String.IsNullOrEmpty(nuevoNroLiquidacion.ToString()))
                                //{
                                for (int j = 0; j < cantRows; j++) // Mientras tenga registros 
                                {
                                    row = j + 1;

                                    oStaticText.Caption = "Asignando " + row + "/" + cantRows;

                                    SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row); // DocNum
                                    clsDocumento docSeleccionado = new clsDocumento();
                                    docSeleccionado.DocNum = Int32.Parse(ed.Value);
                                    ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row); // DocEntry
                                    docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                    ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row); // Tipo
                                    docSeleccionado.Tipo = ed.Value;

                                    bool res = asignarLiquidacion(docSeleccionado, nuevoNroLiquidacion, fechaLiquidacion);
                                }
                                //}
                                //else
                                //    SBO_Application.MessageBox("Debe ingresar el nuevo número de liquidación para el documento");

                                oStaticText.Caption = "Asignar Todos"; // Le pongo el nombro original al Botón

                                // Valor por defecto al combo
                                SAPbouiCOM.ComboBox combo = null;
                                combo = oForm.Items.Item("1000001").Specific;
                                combo.Select("No filtrar", BoSearchKey.psk_ByDescription);
                                CargarGrilla(); // Recargo la grilla
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al querer Generar el XML", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }
                    #endregion

                    #region "ActualizarComboLiquidaciones"
                    // Clic en el botón Actualizar, para refrescar el combo de Liquidaciones y muestre todas las creadas
                    if (pVal.ItemUID.Equals("33") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.ComboBox oStaticCombo;
                            oStaticCombo = oFormDatosPedido.Items.Item("1000001").Specific;
                            llenarCombo(oStaticCombo, "Select Code,Name from [@LIQUIDACIONES] where U_ESTADO <> '2' order by CAST(Code AS Int)", true, true);
                            oStaticCombo.Item.Refresh();
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "ResizeVisorLiquidaciones"
                    // Clic en minimizar/maximizar
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            if (oFormDatosPedido != null)
                            {
                                SAPbouiCOM.Matrix matriz = null;
                                matriz = oFormDatosPedido.Items.Item("31").Specific;

                                int altura = Convert.ToInt32(oFormDatosPedido.Height * 0.72); // Le paso como tamaño un 0,71 % del tamaño del formulario
                                if (oFormDatosPedido.Height < 650)
                                    altura = Convert.ToInt32(oFormDatosPedido.Height * 0.68); // Le paso como tamaño un 0,71 % del tamaño del formulario
                                matriz.Item.Height = altura;
                                matriz.Item.Width = 260; // Le dejo fijo el ancho a la grilla

                                matriz.LoadFromDataSource();
                                matriz.AutoResizeColumns();

                                SAPbouiCOM.StaticText texto = null;
                                texto = oFormDatosPedido.Items.Item("23").Specific;
                                if (oFormDatosPedido.Height < 650)
                                    altura = Convert.ToInt32(oFormDatosPedido.Height * 0.88); // Le paso como posicion un 0.90 % del tamaño del formulario
                                else
                                    altura = Convert.ToInt32(oFormDatosPedido.Height * 0.90); // Le paso como posicion un 0.90 % del tamaño del formulario
                                texto.Item.Top = altura;

                                SAPbouiCOM.ComboBox combo = null;
                                combo = oFormDatosPedido.Items.Item("1000001").Specific;
                                combo.Item.Top = altura;

                                SAPbouiCOM.Button button = null;
                                button = oFormDatosPedido.Items.Item("33").Specific;
                                button.Item.Top = altura;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(button);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(matriz);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(texto);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(combo);
                            }
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "ResizeVisorPedidos"
                    // Clic en minimizar/maximizar
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE && pVal.FormUID.Equals("VisualizarPedidos"))
                    {
                        try
                        {
                            if (oFormDatosPedidoVisor != null)
                            {
                                SAPbouiCOM.Matrix matriz = null;
                                matriz = oFormDatosPedidoVisor.Items.Item("31").Specific;


                                int altura = Convert.ToInt32(oFormDatosPedidoVisor.Height * 0.72); // Le paso como tamaño un 0,71 % del tamaño del formulario
                                if (oFormDatosPedidoVisor.Height < 650)
                                    altura = Convert.ToInt32(oFormDatosPedidoVisor.Height * 0.68); // Le paso como tamaño un 0,71 % del tamaño del formulario
                                matriz.Item.Height = altura;
                                matriz.Item.Width = 260; // Le dejo fijo el ancho a la grilla

                                matriz.LoadFromDataSource();
                                matriz.AutoResizeColumns();

                                SAPbouiCOM.StaticText texto = null;
                                texto = oFormDatosPedidoVisor.Items.Item("23").Specific;
                                if (oFormDatosPedidoVisor.Height < 650)
                                    altura = Convert.ToInt32(oFormDatosPedidoVisor.Height * 0.88); // Le paso como posicion un 0.90 % del tamaño del formulario
                                else
                                    altura = Convert.ToInt32(oFormDatosPedidoVisor.Height * 0.90); // Le paso como posicion un 0.90 % del tamaño del formulario
                                texto.Item.Top = altura;

                                SAPbouiCOM.ComboBox combo = null;
                                combo = oFormDatosPedidoVisor.Items.Item("1000001").Specific;
                                combo.Item.Top = altura;

                                System.Runtime.InteropServices.Marshal.ReleaseComObject(matriz);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(texto);
                                System.Runtime.InteropServices.Marshal.ReleaseComObject(combo);
                            }
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "ClicBotonesConfirmarPedido"
                    // Asigna nuevo estado de confirmacion al documento seleccionado
                    if (pVal.ItemUID.Equals("7") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarPedidos"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarPedidos", 0);
                            oStatic = oForm.Items.Item("2").Specific;

                            SAPbouiCOM.ComboBox oStaticCombo;
                            oStaticCombo = oFormDatosPedidoVisor.Items.Item("1000001").Specific; // Estado de confirmacion
                            string estadoNuevo = "";
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                estadoNuevo = oStaticCombo.Selected.Value.ToString();

                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea asignar el estado " + estadoNuevo + " a los documentos seleccionados?", 2, "Aceptar", "Cancelar");

                            if (respuestaMge == 1)
                            {
                                if (estadoNuevo.ToString().Equals("Y") || estadoNuevo.ToString().Equals("N"))
                                {
                                    int row = oStatic.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_RowOrder);

                                    if (row == -1)
                                        SBO_Application.MessageBox("Debe seleccionar una fila");
                                    else
                                    {
                                        while (row != -1)
                                        {
                                            SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row); // DocNum
                                            clsDocumento docSeleccionado = new clsDocumento();
                                            docSeleccionado.DocNum = Int32.Parse(ed.Value);
                                            ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row); // DocEntry
                                            docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                            ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row); // Tipo
                                            docSeleccionado.Tipo = ed.Value;

                                            bool res = asignarEstadoConfirmacionPedido(docSeleccionado, estadoNuevo);

                                            if (res == true)
                                            {
                                                if (estadoNuevo.ToString().Equals("Y"))
                                                {
                                                    //asignarFechaEntrega(docSeleccionado);
                                                    docSeleccionado = asignarExtraDays(docSeleccionado);
                                                    //asignarFechaEntregaTerritorioExtraDaysDiscPrcnt(docSeleccionado);
                                                    if (actualizaCentroCostos == true)
                                                        asignarCentroCostos(docSeleccionado);
                                                    //asignarDiscPrcnt(docSeleccionado);
                                                    //asignarTerritorioDocumento(docSeleccionado);
                                                }
                                            }
                                            else
                                                SBO_Application.MessageBox("Ha ocurrido un error al asignar el Estado al Pedido " + docSeleccionado.DocNum);

                                            row = oStatic.GetNextSelectedRow(row, SAPbouiCOM.BoOrderType.ot_RowOrder);
                                        }

                                        // Valor por defecto al combo
                                        SAPbouiCOM.ComboBox combo = null;
                                        combo = oForm.Items.Item("1000001").Specific;
                                        combo.Select("No filtrar", BoSearchKey.psk_ByDescription);
                                        CargarGrillaPedidosVisor();
                                    }
                                }
                                else
                                    SBO_Application.MessageBox("Debe indicar un campo de Confirmación válido (Y o N)");
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al asignar el Estado al Pedido", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }

                    // // Asigna nuevo numero de Liquidacion a todos los documentos de la grilla cuando presiona clic en "Procesar Todos"
                    if (pVal.ItemUID.Equals("15") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarPedidos"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Button oStaticText;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarPedidos", 0);
                            oStatic = oForm.Items.Item("2").Specific;
                            SAPbouiCOM.ComboBox oStaticCombo;
                            oStaticCombo = oFormDatosPedidoVisor.Items.Item("1000001").Specific; // Nro de liquidación nuevo
                            string estadoNuevo = "";
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                estadoNuevo = oStaticCombo.Selected.Value.ToString();
                            oStaticText = oForm.Items.Item("15").Specific;

                            int cantRows = oStatic.RowCount; // Saco la cantidad de registros que tiene la Grilla
                            int row = 0;

                            if (estadoNuevo.ToString().Equals("Y") || estadoNuevo.ToString().Equals("N"))
                            {
                                //if (!String.IsNullOrEmpty(nuevoNroLiquidacion.ToString()))
                                //{
                                for (int j = 0; j < cantRows; j++) // Mientras tenga registros 
                                {
                                    row = j + 1;

                                    oStaticText.Caption = "Asignando " + row + "/" + cantRows;

                                    SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row); // DocNum
                                    clsDocumento docSeleccionado = new clsDocumento();
                                    docSeleccionado.DocNum = Int32.Parse(ed.Value);
                                    ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row); // DocEntry
                                    docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                    ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row); // Tipo
                                    docSeleccionado.Tipo = ed.Value;

                                    bool res = asignarEstadoConfirmacionPedido(docSeleccionado, estadoNuevo);
                                    if (res == true)
                                    {
                                        if (estadoNuevo.ToString().Equals("Y"))
                                        {
                                            //asignarFechaEntrega(docSeleccionado);
                                            docSeleccionado = asignarExtraDays(docSeleccionado);
                                            //asignarFechaEntregaTerritorioExtraDaysDiscPrcnt(docSeleccionado);
                                            if (actualizaCentroCostos == true)
                                                asignarCentroCostos(docSeleccionado);
                                            //asignarDiscPrcnt(docSeleccionado);
                                            //asignarTerritorioDocumento(docSeleccionado);
                                        }
                                    }
                                }

                                // Valor por defecto al combo
                                SAPbouiCOM.ComboBox combo = null;
                                combo = oForm.Items.Item("1000001").Specific;
                                combo.Select("No filtrar", BoSearchKey.psk_ByDescription);
                                CargarGrillaPedidosVisor();
                            }
                            else
                                SBO_Application.MessageBox("Debe indicar un campo de Confirmación válido (Y o N)");
                            //}
                            //else
                            //    SBO_Application.MessageBox("Debe ingresar el nuevo número de liquidación para el documento");

                            oStaticText.Caption = "Asignar Todos"; // Le pongo el nombro original al Botón

                            CargarGrillaPedidosVisor(); // Recargo la grilla

                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al asignar Estado al Pedido", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }
                    #endregion

                    #region "BorrarLog"
                    // Borra el log de los documentos
                    if (pVal.ItemUID.Equals("100") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            borrarLog(); // Ejecuto la funcion que elimina el Log
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "BorrarLogPedidos"
                    // Borra el log de los documentos
                    if (pVal.ItemUID.Equals("100") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarPedidos"))
                    {
                        try
                        {
                            borrarLog(); // Ejecuto la funcion que elimina el Log
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "BuscarPedidos"
                    // Busca los documentos segun los filtros y lo muestra en la grilla de EnviarDocumento
                    if (pVal.ItemUID.Equals("8") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarPedidos"))
                    {
                        try
                        {
                            CargarGrillaPedidosVisor(); // Ejecuto la funcion que carga la grilla
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "BuscarCheques"
                    // Busca los cheques segun los filtros y lo muestra en la grilla de EnviarDocumento
                    if (pVal.ItemUID.Equals("8") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarCheques"))
                    {
                        try
                        {
                            CargarGrillaChequesVisor(false); // Ejecuto la funcion que carga la grilla
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "BuscarDocumentos"
                    // Busca los documentos segun los filtros y lo muestra en la grilla de EnviarDocumento
                    if (pVal.ItemUID.Equals("8") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            CargarGrilla(); // Ejecuto la funcion que carga la grilla
                        }
                        catch (Exception ex)
                        { }
                    }
                    #endregion

                    #region "Deposita Cheques"
                    // Acredita todos los cheques que tengan el check marcado
                    if (pVal.ItemUID.Equals("15") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarCheques"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Button oStaticText;
                            SAPbouiCOM.EditText oStaticTexto;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarCheques", 0);
                            oStatic = oForm.Items.Item("2").Specific;
                            SAPbouiCOM.ComboBox oStaticCombo;

                            DateTime fechaDeposito = DateTime.Now; string cuentaDeposito = ""; double sumaCheques = 0; string moneda = ""; string cuentaAcreditacion = "";

                            oStaticTexto = oFormDatosCheques.Items.Item("10").Specific; // Fecha
                            if (!String.IsNullOrEmpty(oStaticTexto.String))
                                fechaDeposito = Convert.ToDateTime(oStaticTexto.String);

                            oStaticCombo = oFormDatosCheques.Items.Item("28").Specific; // Cuenta de acreditacion
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                cuentaAcreditacion = oStaticCombo.Selected.Value.ToString();

                            oStaticCombo = oFormDatosCheques.Items.Item("1000001").Specific; // Cuenta de deposito

                            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea depositar los cheques seleccionados?", 2, "Aceptar", "Cancelar");

                            if (respuestaMge == 1)
                            {
                                if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()) && !String.IsNullOrEmpty(cuentaAcreditacion))
                                {
                                    cuentaDeposito = oStaticCombo.Selected.Value.ToString();
                                    oStaticText = oForm.Items.Item("15").Specific;

                                    int cantRows = oStatic.RowCount; // Saco la cantidad de registros que tiene la Grilla
                                    int row = 0;

                                    for (int j = 0; j < cantRows; j++) // Mientras tenga registros 
                                    {
                                        row = j + 1;

                                        oStaticText.Caption = "Depositando " + row + "/" + cantRows;

                                        SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row);
                                        clsCheque docSeleccionado = new clsCheque();
                                        docSeleccionado.NumSecuencia = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_12", row);
                                        docSeleccionado.NumCheque = Int32.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_3", row);
                                        docSeleccionado.Monto = Double.Parse(ed.Value);
                                        ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_2", row);
                                        docSeleccionado.Moneda = ed.Value;

                                        CheckBox oChk = (CheckBox)oStatic.GetCellSpecific("V_0", row);
                                        bool esSeleccionado = oChk.Checked;
                                        if (esSeleccionado == true) // Si es un Cheque seleccionado
                                        {
                                            depositarCheque(docSeleccionado, "Y", 999999999);
                                            sumaCheques += docSeleccionado.Monto; // Le sumo el monto
                                        }

                                        moneda = docSeleccionado.Moneda;
                                    }

                                    try
                                    {
                                        // Creo el asiento
                                        SAPbobsCOM.JournalEntries oDoc = oCompany.GetBusinessObject(BoObjectTypes.oJournalEntries);
                                        int lRetCode;
                                        oDoc.DueDate = fechaDeposito;
                                        oDoc.TaxDate = fechaDeposito;
                                        oDoc.ReferenceDate = fechaDeposito;
                                        oDoc.Memo = "Acreditación de Cheques";

                                        // Linea 1
                                        oDoc.Lines.AccountCode = cuentaAcreditacion;
                                        if (obtenerCodigoISOMoneda(moneda).ToString().Equals("UYU") || obtenerCodigoISOMoneda(moneda).ToString().Equals("CLP")) // Si es moneda local
                                        {
                                            oDoc.Lines.Credit = sumaCheques;
                                            oDoc.Lines.Debit = 0;
                                        }
                                        else
                                        {
                                            oDoc.Lines.FCCredit = sumaCheques;
                                            oDoc.Lines.FCDebit = 0;
                                        }

                                        oDoc.Lines.Add();

                                        // Linea 2
                                        oDoc.Lines.AccountCode = cuentaDeposito;
                                        if (obtenerCodigoISOMoneda(moneda).ToString().Equals("UYU") || obtenerCodigoISOMoneda(moneda).ToString().Equals("CLP")) // Si es moneda local
                                        {
                                            oDoc.Lines.Credit = 0;
                                            oDoc.Lines.Debit = sumaCheques;
                                        }
                                        else
                                        {
                                            oDoc.Lines.FCCredit = 0;
                                            oDoc.Lines.FCDebit = sumaCheques;
                                        }

                                        oDoc.Lines.Add();

                                        if (oDoc.Lines.Count != 0) // Si el Asiento tiene alguna linea
                                        {
                                            //guardaLogProceso(docEntryDocumento.ToString(), docEntryDocumento.ToString(), "Guardando pago.. ", "CardCode:" + oDoc.CardCode.ToString() + " TransferAccount: " + oDoc.TransferAccount + " TransferSum: " + oDoc.TransferSum + " DocCurrency: " + oDoc.DocCurrency + " TransferReference: " + oDoc.TransferReference); // Guarda log del Proceso

                                            lRetCode = oDoc.Add();

                                            if (lRetCode != 0)
                                            {
                                                eliminarRegistrosRecientementeDepositados();
                                                string error = oCompany.GetLastErrorDescription();
                                                SBO_Application.MessageBox(error + ". Codigo error:" + lRetCode.ToString());
                                            }
                                            else
                                            {
                                                actualizarRegistrosRecientementeDepositados();
                                                SBO_Application.MessageBox("Transacción Correcta.");
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    { }

                                    oStaticText.Caption = "Depositar seleccionados"; // Le pongo el nombro original al Botón

                                    CargarGrillaChequesVisor(false); // Recargo la grilla
                                }
                                else
                                    SBO_Application.MessageBox("Debe indicar una cuenta de Depósito");
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al querer Generar el XML", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }
                    #endregion

                    #region "VerDocumentoSeleccionado"
                    // Abre el documento seleccionado
                    if (pVal.ItemUID.Equals("1000004") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarLiquidaciones"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarLiquidaciones", 0);
                            oStatic = oForm.Items.Item("2").Specific;

                            int row = oStatic.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (row == -1)
                            {
                                SBO_Application.MessageBox("Debe seleccionar una fila");
                            }
                            else
                            {
                                SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row); // DocNum
                                clsDocumento docSeleccionado = new clsDocumento();
                                docSeleccionado.DocNum = Int32.Parse(ed.Value);
                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row); // DocEntry
                                docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row); // Tipo
                                docSeleccionado.Tipo = ed.Value;

                                bool res = abrirDocumento(docSeleccionado.DocEntry.ToString(), docSeleccionado.DocNum.ToString(), docSeleccionado.Tipo);

                                if (res == false)
                                    SBO_Application.MessageBox("Ha ocurrido un error al querer abrir el Documento");
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al Visualizar el documento", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }
                    #endregion

                    #region "VerDocumentoSeleccionadoPedidos"
                    // Abre el documento seleccionado
                    if (pVal.ItemUID.Equals("1000004") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("VisualizarPedidos"))
                    {
                        try
                        {
                            SAPbouiCOM.Matrix oStatic;
                            SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("VisualizarPedidos", 0);
                            oStatic = oForm.Items.Item("2").Specific;

                            int row = oStatic.GetNextSelectedRow(0, SAPbouiCOM.BoOrderType.ot_SelectionOrder);
                            if (row == -1)
                            {
                                SBO_Application.MessageBox("Debe seleccionar una fila");
                            }
                            else
                            {
                                SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row); // DocNum
                                clsDocumento docSeleccionado = new clsDocumento();
                                docSeleccionado.DocNum = Int32.Parse(ed.Value);
                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row); // DocEntry
                                docSeleccionado.DocEntry = Int32.Parse(ed.Value);
                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_9", row); // Tipo
                                docSeleccionado.Tipo = ed.Value;

                                bool res = abrirDocumento(docSeleccionado.DocEntry.ToString(), docSeleccionado.DocNum.ToString(), docSeleccionado.Tipo);

                                if (res == false)
                                    SBO_Application.MessageBox("Ha ocurrido un error al querer abrir el Documento");
                            }
                        }
                        catch (Exception ex)
                        {
                            guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR al Visualizar el documento", ex.Message.ToString()); // Guarda log del Proceso
                        }
                        //oStatic.GetLineData(row);
                    }
                    #endregion

                    #region "CobrarDocumentos"
                    // Lista los documentos que va a cobrar en la grilla
                    if (pVal.ItemUID.Equals("1000002") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoCobros"))
                    {
                        try
                        {
                            string numeroLiquidacion = ""; double importeCobro = 0; string monedaCobro = "";

                            SAPbouiCOM.ComboBox oStaticCombo;
                            SAPbouiCOM.EditText oStaticTexto;

                            oStaticCombo = oFormDatosIngresoCobros.Items.Item("29").Specific; // Nro de liquidación seleccionada por el usuario
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                numeroLiquidacion = oStaticCombo.Selected.Value.ToString();

                            oStaticTexto = oFormDatosIngresoCobros.Items.Item("66").Specific; // Importe
                            if (!String.IsNullOrEmpty(oStaticTexto.String))
                                importeCobro = Convert.ToDouble(oStaticTexto.Value);

                            oStaticCombo = oFormDatosIngresoCobros.Items.Item("1000001").Specific; // Moneda seleccionada
                            if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                monedaCobro = oStaticCombo.Selected.Value.ToString();

                            if (!String.IsNullOrEmpty(numeroLiquidacion.ToString()) && !String.IsNullOrEmpty(monedaCobro.ToString()) && importeCobro > 0) // Si selecciono un Número de Liquidacion y el importe de cobro es mayor a 0
                            {
                                if (CargarMatrixIngresoCobros(numeroLiquidacion, monedaCobro, importeCobro) > 0) ; //// Prueba de cargar Grilla
                                {   // Si trae algún registro habilitá el botón de Ingreso del Comprobante, de lo contrario NO.
                                    SAPbouiCOM.Button oStaticText = oFormDatosIngresoCobros.Items.Item("8").Specific;
                                    oStaticText.Item.Enabled = true;
                                }
                            }
                            else
                                SBO_Application.MessageBox("Debe seleccionar la Liquidación, moneda e importe de Cobro para listar los documentos.");
                        }
                        catch (Exception ex)
                        { }
                    }

                    // Cobro los documentos correspondientes
                    if (pVal.ItemUID.Equals("8") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoCobros"))
                    {
                        try
                        {
                            SAPbouiCOM.Button oStaticText = oFormDatosIngresoCobros.Items.Item("8").Specific;

                            if (oStaticText.Item.Enabled == true)
                            {
                                int respuestaMge = SBO_Application.MessageBox("Está seguro que desea procesar?", 2, "Procesar", "Cancelar");

                                if (respuestaMge == 1) // Quiere decir que dio clic en Procesar
                                {
                                    SAPbouiCOM.ComboBox oStaticCombo;
                                    SAPbouiCOM.EditText oStaticTexto;
                                    //SAPbouiCOM.Form oFormDatosIngresoCobros = SBO_Application.Forms.GetForm("IngresoCobros", 0);

                                    string numeroLiquidacion = ""; DateTime fechaCobro = DateTime.Now; double importeCobro = 0; double saldoImporteCobro = 0; string numeroReferencia = ""; string monedaCobro = ""; string cuentaContable = "";
                                    int docEntryLogIngreso = 0; // Aca guardo el DocEntry del proceso de la tabla [@IVZ_ODEP] con lo que voy a enlazar a [@IVZ_DEP1]

                                    oStaticCombo = oFormDatosIngresoCobros.Items.Item("29").Specific; // Nro de liquidación seleccionada por el usuario
                                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                        numeroLiquidacion = oStaticCombo.Selected.Value.ToString();

                                    oStaticTexto = oFormDatosIngresoCobros.Items.Item("10").Specific; // Fecha
                                    if (!String.IsNullOrEmpty(oStaticTexto.String))
                                        fechaCobro = Convert.ToDateTime(oStaticTexto.String);

                                    oStaticTexto = oFormDatosIngresoCobros.Items.Item("66").Specific; // Importe
                                    if (!String.IsNullOrEmpty(oStaticTexto.String))
                                        importeCobro = Convert.ToDouble(oStaticTexto.Value);

                                    oStaticTexto = oFormDatosIngresoCobros.Items.Item("16").Specific; // Número de Referencia
                                    if (!String.IsNullOrEmpty(oStaticTexto.String))
                                        numeroReferencia = oStaticTexto.String.ToString();

                                    oStaticCombo = oFormDatosIngresoCobros.Items.Item("1000001").Specific; // Moneda seleccionada
                                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                        monedaCobro = oStaticCombo.Selected.Value.ToString();

                                    oStaticCombo = oFormDatosIngresoCobros.Items.Item("51").Specific; // Cuenta seleccionada
                                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                                        cuentaContable = oStaticCombo.Selected.Value.ToString();

                                    if (!String.IsNullOrEmpty(numeroLiquidacion.ToString()) && !String.IsNullOrEmpty(cuentaContable.ToString()) && importeCobro > 0) // Si selecciono un Número de Liquidacion y el importe de cobro es mayor a 0
                                    {
                                        SAPbouiCOM.Matrix oStatic;
                                        oStatic = oFormDatosIngresoCobros.Items.Item("15").Specific;
                                        int cantRows = oStatic.RowCount; // Saco la cantidad de registros que tiene la Grilla
                                        int row = 0; bool sigoRecorriendo = true; bool estadoProceso = true; int cantRegistros = 0;

                                        for (int i = 0; i < cantRows && sigoRecorriendo == true; i++) // Mientras tenga registros y mientras no se Cancele por el usuario
                                        {
                                            row = i + 1;
                                            int objTypeDoc = 13; // Facturas
                                            SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_7", row);
                                            int docEntryDocumento = Convert.ToInt32(ed.Value.ToString());
                                            ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_6", row);
                                            int docNumDocumento = Convert.ToInt32(ed.Value.ToString());
                                            ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_5", row);
                                            string cardCodeDocumento = ed.Value.ToString();
                                            ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_4", row);
                                            string cardNameDocumento = ed.Value.ToString();
                                            string monedaDocumento = monedaCobro;
                                            double montoCobrarDocumento = 0;
                                            double montoDocumento = 0; double montoDocumentoFC = 0;
                                            try
                                            {
                                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_1", row); // Obtengo el monto a Cobrar del documento
                                                montoCobrarDocumento = Convert.ToDouble(ed.Value.ToString());
                                            }
                                            catch (Exception ex)
                                            {
                                                montoCobrarDocumento = 0;
                                            }

                                            try
                                            {
                                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_3", row); // Obtengo el monto del documento
                                                montoDocumento = Convert.ToDouble(ed.Value.ToString());
                                            }
                                            catch (Exception ex)
                                            {
                                                montoDocumento = 0;
                                            }

                                            try
                                            {
                                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_2", row); // Obtengo el montoFC del documento
                                                montoDocumentoFC = Convert.ToDouble(ed.Value.ToString());
                                            }
                                            catch (Exception ex)
                                            {
                                                montoDocumentoFC = 0;
                                            }

                                            try
                                            {
                                                ed = (SAPbouiCOM.EditText)oStatic.GetCellSpecific("V_11", row); // Obtengo el tipo de documento
                                                string tipoDocString = Convert.ToString(ed.Value.ToString());
                                                if (!tipoDocString.ToString().Contains("Factura"))
                                                    objTypeDoc = 14;
                                            }
                                            catch (Exception ex)
                                            { }

                                            string codigoISOMoneda = obtenerCodigoISOMoneda(monedaDocumento); // Obtengo el codigo ISO de la moneda
                                            string codigoISOMonedaCobro = codigoISOMoneda;
                                            try
                                            {
                                                if (montoCobrarDocumento > 0)
                                                {
                                                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                                                    int lRetCode;
                                                    oDoc.CardCode = cardCodeDocumento;
                                                    oDoc.CardName = cardNameDocumento;
                                                    oDoc.TransferAccount = cuentaContable;
                                                    oDoc.TransferSum = montoCobrarDocumento;
                                                    oDoc.TransferDate = fechaCobro;
                                                    oDoc.TransferReference = numeroReferencia;
                                                    oDoc.DocCurrency = monedaDocumento;

                                                    oDoc.Invoices.DocEntry = docEntryDocumento;
                                                    if (codigoISOMoneda.ToString().Equals("UYU") || codigoISOMoneda.ToString().Equals("CLP")) // Si el documento es en Pesos
                                                        oDoc.Invoices.SumApplied = montoCobrarDocumento;
                                                    else
                                                        oDoc.Invoices.AppliedFC = montoCobrarDocumento;

                                                    oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;

                                                    if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                                                    {
                                                        lRetCode = oDoc.Add();

                                                        if (lRetCode != 0)
                                                        {
                                                            string error = oCompany.GetLastErrorDescription();
                                                            SBO_Application.MessageBox(error + ". Codigo error:" + lRetCode.ToString());
                                                            sigoRecorriendo = false;
                                                            estadoProceso = false;
                                                        }
                                                        else
                                                        {
                                                            actualizarRegistrosIngresoCobros(numeroLiquidacion, docEntryDocumento, objTypeDoc); // Actualizo el campo U_LIQUIDACION
                                                            //actualizarAcctCode(docEntryDocumento); // Actualizo el campo AcctCode dependiendo del TaxCode
                                                            saldoImporteCobro = saldoImporteCobro - montoCobrarDocumento; // Le resto al Saldo el total del documento
                                                            cantRegistros++; // Le sumo 1 al contador de registros procesados correctamente
                                                            if (docEntryLogIngreso == 0) // Si es 0 hago el insert de la linea y me quedo con el DocEntry de la misma
                                                                docEntryLogIngreso = guardoLogIngresoComprobantes(numeroLiquidacion, cuentaContable, numeroReferencia, importeCobro, fechaCobro, codigoISOMonedaCobro, 0);
                                                            guardoLogIngresoComprobantesLinea(docEntryLogIngreso, docNumDocumento, montoDocumento, montoCobrarDocumento, codigoISOMoneda);
                                                            //SBO_Application.MessageBox("Transacción Correcta.");
                                                        }
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            { }
                                        }

                                        if (estadoProceso == true && cantRegistros > 0) // Si el proceso fue correcto y proceso correctamente al menos 1 documento
                                        {
                                            double montoPago = importeCobro; // Obtengo el total de lo que se pago  - saldoImporteCobro; 
                                            actualizoLogIngresoComprobantes(docEntryLogIngreso, montoPago); // Actualizo el total aplicado a Pagos
                                            SBO_Application.MessageBox("Transacción Correcta. " + cantRegistros + " documentos procesados.");

                                            oStaticText.Item.Enabled = false; // Inhabilito el botón nuevamente
                                            oStatic.Clear();
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        { }
                    }

                    //// Cobro los documentos correspondientes
                    //if (pVal.ItemUID.Equals("8") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.FormUID.Equals("IngresoCobros"))
                    //{
                    //    try
                    //    {
                    //        SAPbouiCOM.Button oStaticText = oFormDatosIngresoCobros.Items.Item("8").Specific;

                    //        if (oStaticText.Item.Enabled == true)
                    //        {
                    //            int respuestaMge = SBO_Application.MessageBox("Está seguro que desea procesar?", 2, "Procesar", "Cancelar");

                    //            if (respuestaMge == 1) // Quiere decir que dio clic en Procesar
                    //            {
                    //                SAPbouiCOM.ComboBox oStaticCombo;
                    //                SAPbouiCOM.EditText oStaticTexto;
                    //                //SAPbouiCOM.Form oFormDatosIngresoCobros = SBO_Application.Forms.GetForm("IngresoCobros", 0);

                    //                string numeroLiquidacion = ""; DateTime fechaCobro = DateTime.Now; double importeCobro = 0; double saldoImporteCobro = 0; string numeroReferencia = ""; string monedaCobro = ""; string cuentaContable = "";
                    //                int docEntryLogIngreso = 0; // Aca guardo el DocEntry del proceso de la tabla [@IVZ_ODEP] con lo que voy a enlazar a [@IVZ_DEP1]

                    //                oStaticCombo = oFormDatosIngresoCobros.Items.Item("29").Specific; // Nro de liquidación seleccionada por el usuario
                    //                if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                    //                    numeroLiquidacion = oStaticCombo.Selected.Value.ToString();

                    //                oStaticTexto = oFormDatosIngresoCobros.Items.Item("10").Specific; // Fecha
                    //                if (!String.IsNullOrEmpty(oStaticTexto.String))
                    //                    fechaCobro = Convert.ToDateTime(oStaticTexto.String);

                    //                oStaticTexto = oFormDatosIngresoCobros.Items.Item("66").Specific; // Importe
                    //                if (!String.IsNullOrEmpty(oStaticTexto.String))
                    //                    importeCobro = Convert.ToDouble(oStaticTexto.Value);

                    //                oStaticTexto = oFormDatosIngresoCobros.Items.Item("16").Specific; // Número de Referencia
                    //                if (!String.IsNullOrEmpty(oStaticTexto.String))
                    //                    numeroReferencia = oStaticTexto.String.ToString();

                    //                oStaticCombo = oFormDatosIngresoCobros.Items.Item("1000001").Specific; // Moneda seleccionada
                    //                if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                    //                    monedaCobro = oStaticCombo.Selected.Value.ToString();

                    //                oStaticCombo = oFormDatosIngresoCobros.Items.Item("51").Specific; // Cuenta seleccionada
                    //                if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                    //                    cuentaContable = oStaticCombo.Selected.Value.ToString();

                    //                if (!String.IsNullOrEmpty(numeroLiquidacion.ToString()) && !String.IsNullOrEmpty(cuentaContable.ToString()) && importeCobro > 0) // Si selecciono un Número de Liquidacion y el importe de cobro es mayor a 0
                    //                {
                    //                    ////// Prueba de cargar Grilla
                    //                    //CargarMatrixIngresoCobros(numeroLiquidacion, monedaCobro,importeCobro);

                    //                    bool estadoProceso = true; int cantRegistros = 0;
                    //                    SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    //                    string query = "select oin.DocEntry,oin.DocNum,PaidToDate,oin.CardCode,oin.CardName,oin.DocTotal,oin.DocTotalFC,oin.DocCur,oin.VatSum,oin.VatSumFC from OINV as oin " +
                    //                    "inner join OCRD as ocr on ocr.CardCode = oin.CardCode and ocr.CardType = 'C' " +
                    //                    "inner join OCTG as oct on oct.GroupNum = oin.GroupNum " +
                    //                    "where oin.U_LIQUIDACION = '" + numeroLiquidacion.ToString() + "' and oin.Canceled = 'N' and oct.PymntGroup like 'CONTADO%' and oin.DocCur = '" + monedaCobro + "' and oin.FolioNum is not null "; // Solo documentos que tengan como cond de pago Contado

                    //                    string codigoISOMonedaCobro = obtenerCodigoISOMoneda(monedaCobro); // Obtengo el codigo ISO de la moneda de cobro
                    //                    if (codigoISOMonedaCobro.ToString().Equals("UYU") || codigoISOMonedaCobro.ToString().Equals("CLP")) // Si el documento es en Pesos
                    //                        query += "and PaidToDate <> DocTotal "; // Acá verifica que el documento tenga saldo pendiente aún
                    //                    else
                    //                        query += "and PaidToDate <> DocTotalFC "; // Acá verifica que el documento tenga saldo pendiente aún

                    //                    query += "order by oin.DocEntry";
                    //                    oRSMyTable.DoQuery(query);

                    //                    if (oRSMyTable != null)
                    //                    {
                    //                        if (oRSMyTable.RecordCount > 0)
                    //                        {
                    //                            saldoImporteCobro = importeCobro; // Creo una variable saldo para ir restandole el total de cada documento
                    //                            bool sigoRecorriendo = true; // Bandera para saber si sigo en el while o por algun motivo debo salir

                    //                            while (!oRSMyTable.EoF && saldoImporteCobro > 0 && sigoRecorriendo == true) // Recorro los documentos mientras hay saldo para asignar
                    //                            {
                    //                                double importeYaPagoDocumento = (double)oRSMyTable.Fields.Item("PaidToDate").Value; // Importe que ya se pagó del documento
                    //                                double montoDocumento = (double)oRSMyTable.Fields.Item("DocTotal").Value; // Monto del documento
                    //                                double montoDocumentoFC = (double)oRSMyTable.Fields.Item("DocTotalFC").Value; // Monto del documento en Moneda extranjera
                    //                                int docEntryDocumento = oRSMyTable.Fields.Item("DocEntry").Value;
                    //                                int docNumDocumento = oRSMyTable.Fields.Item("DocNum").Value;
                    //                                string cardCodeDocumento = (string)oRSMyTable.Fields.Item("CardCode").Value;
                    //                                string cardNameDocumento = (string)oRSMyTable.Fields.Item("CardName").Value;
                    //                                string monedaDocumento = (string)oRSMyTable.Fields.Item("DocCur").Value;
                    //                                string codigoISOMoneda = obtenerCodigoISOMoneda(monedaDocumento); // Obtengo el codigo ISO de la moneda

                    //                                double montoCobrarDocumento = 0;
                    //                                if (codigoISOMoneda.ToString().Equals("UYU") || codigoISOMoneda.ToString().Equals("CLP")) // Si el documento es en Pesos
                    //                                    montoCobrarDocumento = montoDocumento - importeYaPagoDocumento; // Obtengo el saldo que falta pagar del documento
                    //                                else
                    //                                {
                    //                                    montoCobrarDocumento = montoDocumentoFC - importeYaPagoDocumento; // Si es en Moneda Extranjera. Obtengo el saldo que falta pagar del documento
                    //                                    montoDocumento = montoDocumentoFC; // Me guardo en esta variable el monto original del documento
                    //                                }

                    //                                if (montoCobrarDocumento > saldoImporteCobro) // Si lo que tengo para cobrar es mayor al saldo que tengo disponible
                    //                                    montoCobrarDocumento = saldoImporteCobro;

                    //                                try
                    //                                {
                    //                                    SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments);
                    //                                    int lRetCode;
                    //                                    oDoc.CardCode = cardCodeDocumento;
                    //                                    oDoc.CardName = cardNameDocumento;
                    //                                    oDoc.TransferAccount = cuentaContable;
                    //                                    oDoc.TransferSum = montoCobrarDocumento;
                    //                                    oDoc.TransferDate = fechaCobro;
                    //                                    oDoc.TransferReference = numeroReferencia;
                    //                                    oDoc.DocCurrency = monedaDocumento;

                    //                                    oDoc.Invoices.DocEntry = docEntryDocumento;
                    //                                    if (codigoISOMoneda.ToString().Equals("UYU") || codigoISOMoneda.ToString().Equals("CLP")) // Si el documento es en Pesos
                    //                                    {
                    //                                        //oDoc.Invoices.AppliedFC = 0;
                    //                                        oDoc.Invoices.SumApplied = montoCobrarDocumento;
                    //                                    }
                    //                                    else
                    //                                    {
                    //                                        oDoc.Invoices.AppliedFC = montoCobrarDocumento;
                    //                                        //oDoc.Invoices.SumApplied = 0;
                    //                                    }

                    //                                    //oDoc.Invoices.DocLine = ;
                    //                                    oDoc.Invoices.InvoiceType = BoRcptInvTypes.it_Invoice;
                    //                                    //oDoc.Invoices.Add();

                    //                                    if (oDoc.Invoices.Count != 0) // Si el Documento tiene alguna linea
                    //                                    {
                    //                                        //guardaLogProceso(docEntryDocumento.ToString(), docEntryDocumento.ToString(), "Guardando pago.. ", "CardCode:" + oDoc.CardCode.ToString() + " TransferAccount: " + oDoc.TransferAccount + " TransferSum: " + oDoc.TransferSum + " DocCurrency: " + oDoc.DocCurrency + " TransferReference: " + oDoc.TransferReference); // Guarda log del Proceso

                    //                                        lRetCode = oDoc.Add();

                    //                                        if (lRetCode != 0)
                    //                                        {
                    //                                            string error = oCompany.GetLastErrorDescription();
                    //                                            SBO_Application.MessageBox(error + ". Codigo error:" + lRetCode.ToString());
                    //                                            sigoRecorriendo = false;
                    //                                            estadoProceso = false;
                    //                                        }
                    //                                        else
                    //                                        {
                    //                                            actualizarRegistrosIngresoCobros(numeroLiquidacion, docEntryDocumento); // Actualizo el campo U_LIQUIDACION
                    //                                            //actualizarAcctCode(docEntryDocumento); // Actualizo el campo AcctCode dependiendo del TaxCode
                    //                                            saldoImporteCobro = saldoImporteCobro - montoCobrarDocumento; // Le resto al Saldo el total del documento
                    //                                            cantRegistros++; // Le sumo 1 al contador de registros procesados correctamente
                    //                                            if (docEntryLogIngreso == 0) // Si es 0 hago el insert de la linea y me quedo con el DocEntry de la misma
                    //                                                docEntryLogIngreso = guardoLogIngresoComprobantes(numeroLiquidacion, cuentaContable, numeroReferencia, importeCobro, fechaCobro, codigoISOMonedaCobro, 0);
                    //                                            guardoLogIngresoComprobantesLinea(docEntryLogIngreso, docNumDocumento, montoDocumento, montoCobrarDocumento, codigoISOMoneda);
                    //                                            //SBO_Application.MessageBox("Transacción Correcta.");
                    //                                        }
                    //                                    }
                    //                                }
                    //                                catch (Exception ex)
                    //                                { }

                    //                                oRSMyTable.MoveNext();
                    //                            }
                    //                        }
                    //                        else
                    //                            SBO_Application.MessageBox("No hay documentos para la Liquidación y Moneda seleccionada");
                    //                    }
                    //                    else
                    //                        SBO_Application.MessageBox("No hay documentos para la Liquidación y Moneda seleccionada");

                    //                    if (estadoProceso == true && cantRegistros > 0) // Si el proceso fue correcto y proceso correctamente al menos 1 documento
                    //                    {
                    //                        double montoPago = importeCobro - saldoImporteCobro; // Obtengo el total de lo que se pago
                    //                        actualizoLogIngresoComprobantes(docEntryLogIngreso, montoPago); // Actualizo el total aplicado a Pagos
                    //                        SBO_Application.MessageBox("Transacción Correcta. " + cantRegistros + " documentos procesados.");

                    //                        oStaticText.Item.Enabled = false; // Inhabilito el botón nuevamente
                    //                    }

                    //                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                    //                    oRSMyTable = null;
                    //                }
                    //                else
                    //                    SBO_Application.MessageBox("Debe seleccionar la Liquidación, cuenta contable, e ingresar Importe y Fecha como mínimo");

                    //            }
                    //        }
                    //    }
                    //    catch (Exception ex)
                    //    { }
                    //}
                    #endregion
                }
            }
            catch (Exception ex)
            {
                guardaLogProceso(pVal.FormTypeEx.ToString(), "", "ERROR Evento Principal", ex.Message.ToString()); // Guarda log del Proceso
            }
            BubbleEvent = true;
        }

        // Evento para cerar el SAP
        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            try
            {
                switch (EventType)
                {
                    case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                        System.Windows.Forms.Application.Exit();
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                        break;
                    case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                        break;
                }
            }
            catch (Exception ex)
            {
                guardaLogProceso("99999", "", "ERROR al cerrar el addOn", ex.Message.ToString()); // Guarda log del Proceso
            }
        }
        #endregion

        #region Conexion DIAPI

        // Obtengo los datos de configuracion
        public bool obtenerDatosConfiguracion()
        {
            bool res = false;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select U_DESC_EMPRESA,U_CENTRO_COSTOS from [@CONFIGLIQUIDACIONES]";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        empresa = oRSMyTable.Fields.Item("U_DESC_EMPRESA").Value;
                        int centroCostos = (int)oRSMyTable.Fields.Item("U_CENTRO_COSTOS").Value;

                        if (centroCostos == 1) // Si es 1 entonces actualiza los Centros de Costos
                            actualizaCentroCostos = true;

                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                res = true;

                return res;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso("", "", "ERROR al leer los campos de la BD", ex.Message.ToString()); // Guarda log del Proceso

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al leer los campos de la BD.");
                return res;
            }
        }

        // Obtiene el Id y el Nombre del usuario logueado
        public bool getUsuarioLogueado()
        {
            bool res = false;
            try
            {
                SAPbouiCOM.StaticText oStatic;
                SAPbouiCOM.Form oForm = SBO_Application.Forms.GetForm("169", 0);
                oStatic = (SAPbouiCOM.StaticText)oForm.Items.Item("8").Specific;
                //pasamos a nuestra variable string el nombre del usuario
                usuarioLogueado = oStatic.Caption;

                res = true;

                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso("169", "", "ERROR al buscar usuario logueado", ex.Message.ToString());// Guarda log del Proceso
            }
            return res;
        }

        // Obtengo el Id de Usuario y el Id de Sucursal
        public bool getIdUsuarioIdSucursal()
        {
            bool res = false; SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select USERID, Branch, USER_CODE, U_NAME, SUPERUSER, [U_ADD_ON_LIQUIDACION] from OUSR where U_NAME = '" + usuarioLogueado + "'";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        idSucursalUsuario = Convert.ToInt32(oRSMyTable.Fields.Item("Branch").Value);
                        idUsuarioLogueado = Convert.ToInt32(oRSMyTable.Fields.Item("USERID").Value);
                        string esSuperU = oRSMyTable.Fields.Item("SUPERUSER").Value;
                        if (esSuperU.ToString().Equals("Y"))
                            esSuperUsuario = true;
                        else
                            esSuperUsuario = false;
                        int visualizaAddOn = Convert.ToInt32(oRSMyTable.Fields.Item("U_ADD_ON_LIQUIDACION").Value);
                        if (visualizaAddOn == 2) // Si es 2 ve del addOn solo la parte de Liquidaciones
                            visualizaSoloLiquidaciones = true;
                        oRSMyTable.MoveNext();
                    }
                }

                //if (idSucursalUsuario < 0) // Si la sucursal es menor a 0 entonces se usa la 1 por defecto
                //    idSucursalUsuario = 1;

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                res = true;
                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                guardaLogProceso("", idUsuarioLogueado.ToString(), "ERROR al Buscar Sucursal", ex.Message.ToString()); // Guarda log del Proceso
                return res;
            }
        }

        // Obtengo el Id de Usuario y el Id de Sucursal pasandole el id de Usuario
        public bool getIdSucursal(int pIdUsuario)
        {
            bool res = false; SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select USERID, Branch, USER_CODE, U_NAME from OUSR where USERID = '" + pIdUsuario + "'";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        idSucursalUsuario = Convert.ToInt32(oRSMyTable.Fields.Item("Branch").Value);
                        idUsuarioLogueado = Convert.ToInt32(oRSMyTable.Fields.Item("USERID").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                //if (idSucursalUsuario < 0) // Si la sucursal es menor a 0 entonces se usa la 1 por defecto
                //    idSucursalUsuario = 1;

                res = true;
                return res;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                guardaLogProceso("", idUsuarioLogueado.ToString(), "ERROR al Buscar Sucursal", ex.Message.ToString()); // Guarda log del Proceso
                return res;
            }
        }

        // Convertir a Decimal 
        public decimal getDecimal(string pNumero)
        {
            decimal res = 0;
            try
            {
                res = decimal.Parse(pNumero, System.Globalization.CultureInfo.InvariantCulture);

                return res;
            }
            catch (Exception ex)
            {
                return res;
            }
        }

        public void CrearCampo(SAPbobsCOM.Company oCompany, String tabla, String Nombre, String descripcion, BoFieldTypes tipo, int tamanio)
        {
            SAPbobsCOM.UserFieldsMD oUFields = null;
            try
            {
                oUFields = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUFields.TableName = tabla;
                oUFields.Name = Nombre;
                oUFields.Description = descripcion;
                oUFields.Type = tipo;
                if (tipo == BoFieldTypes.db_Float)
                {
                    oUFields.SubType = BoFldSubTypes.st_Quantity;
                }
                oUFields.EditSize = tamanio;
                int iRet = oUFields.Add();
                string msg = oCompany.GetLastErrorDescription();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields);

                oUFields = null;
                GC.Collect();
            }
            catch
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oUFields);

                oUFields = null;
                GC.Collect();
            }
        }

        public int guardoLogIngresoComprobantes(string pNumeroLiquidacion, string pCuentaContable, string pNumeroReferencia, double pImporteCobro, DateTime pFechaCobro, string pCodigoISOMonedaCobro, double pSaldoImporteCobro)
        {
            int docEntry = 0;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                docEntry = obtenerDocEntryIVZ_ODEP();

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "INSERT INTO [@IVZ_ODEP] (DocEntry, DocNum, U_LIQUIDACION, U_FECHA, U_CUENTA, U_IMPORTE, U_IMPORTE_APLICADO, U_REFERENCIA, U_MONEDA ) VALUES (" + docEntry + "," + docEntry + ",'" + pNumeroLiquidacion + "','" + pFechaCobro.ToString("yyyy-MM-dd") + "','" + pCuentaContable + "','" + pImporteCobro + "','" + pSaldoImporteCobro + "','" + pNumeroReferencia + "','" + pCodigoISOMonedaCobro + "')";

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return docEntry;
            }
            catch (Exception ex)
            {
                return docEntry;
            }
        }

        public int actualizoLogIngresoComprobantes(int pDocEntryRegistro, double pSaldoImporteCobro)
        {
            int docEntry = 0;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                docEntry = obtenerDocEntryIVZ_ODEP();

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "UPDATE [@IVZ_ODEP] SET U_IMPORTE_APLICADO = '" + pSaldoImporteCobro + "' where DocEntry = " + pDocEntryRegistro + " and DocNum = " + pDocEntryRegistro;

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return docEntry;
            }
            catch (Exception ex)
            {
                return docEntry;
            }
        }

        public int guardoLogIngresoComprobantesLinea(int pDocEntryPago, int pDocNumFactura, double pImporteFactura, double pImporteCobrado, string pMonedaFactura)
        {
            int docEntry = 0;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                docEntry = obtenerDocEntryIVZ_DEP1();

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "INSERT INTO [@IVZ_DEP1] (DocEntry,LineId, U_DOCENTRYPAGO, U_DOCNUMPAGO, U_DOCNUMFACTURA, U_IMPORTEFACTURA, U_IMPORTEPAGOFACTURA, U_MONEDAISOFACTURA ) VALUES (" + docEntry + "," + pDocEntryPago + "," + pDocEntryPago + "," + pDocEntryPago + ",'" + pDocNumFactura + "','" + pImporteFactura + "','" + pImporteCobrado + "','" + pMonedaFactura + "')";

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return docEntry;
            }
            catch (Exception ex)
            {
                return docEntry;
            }
        }

        public Boolean guardaLogProceso(String pFormFactura, String pCodigoFactura, String pAccion, String pXML)
        {
            try
            {
                if (guardaLog == true)
                {
                    SAPbobsCOM.Recordset oRSMyTable = null;

                    long docEntry = obtenerDocEntryLogProceso();
                    DateTime fechaHoy = DateTime.Now;
                    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    //String query = "INSERT INTO [@LOGPROCESO] (Code, Name, U_PANTALLA, U_CODIGO,U_ACCION,U_LOGXML, U_FECHA, U_CREATE_DATE) VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + pCodigoFactura + "','" + pAccion + "','" + pXML.ToString() + "','" + fechaHoy.ToString("yyyy-MM-dd HH:mm:ss") + "','" + fechaHoy.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                    String query = "INSERT INTO [@ADDONLOGS] (Code, Name, U_PANTALLA, U_CODIGO,U_ACCION,U_LOGXML, U_FECHA, U_CREATE_DATE) VALUES (" + docEntry + ",'" + docEntry + "','" + pFormFactura + "','" + pCodigoFactura + "','" + pAccion + "','" + pXML.ToString() + "','" + fechaHoy.ToString("yyyy-MM-dd HH:mm:ss") + "','" + fechaHoy.ToString("yyyy-MM-dd HH:mm:ss") + "')";

                    oRSMyTable.DoQuery(query);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                    oRSMyTable = null;
                }

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        // Guarda en la BD un log de todo el proceso
        public Boolean borrarLog()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "delete from [@LOGPROCESO] ";

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        //Obtengo el último DocEntry de la tabla LOGPROCESO
        public long obtenerDocEntryLogProceso()
        {
            long res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //String query = "select case when MAX(CAST(Code AS nvarchar)) is null then 1 else MAX(CAST(Code AS nvarchar)) + 1 end as Prox from [@LOGPROCESO]";
                String query = "select case when MAX(CAST(Code AS bigint)) is null then 1 else MAX(CAST(Code AS bigint)) + 1 end as Prox from [@ADDONLOGS]";

                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt64(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
                return res;
            }
        }

        //// Obtengo el último DocEntry de la tabla LOGPROCESO
        //public int obtenerDocEntryLogProceso()
        //{
        //    int res = 1;
        //    SAPbobsCOM.Recordset oRSMyTable = null;
        //    try
        //    {
        //        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        String query = "select  case when MAX(CAST(DocEntry AS Int)) is null then 1 else MAX(CAST(DocEntry AS Int)) + 1 end as Prox from [@LOGPROCESO]";
        //        oRSMyTable.DoQuery(query);

        //        if (oRSMyTable != null)
        //        {
        //            while (!oRSMyTable.EoF)
        //            {
        //                res = Convert.ToInt32(oRSMyTable.Fields.Item("Prox").Value);
        //                oRSMyTable.MoveNext();
        //            }
        //        }

        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
        //        oRSMyTable = null;

        //        return res;
        //    }
        //    catch (Exception ex)
        //    {
        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
        //        oRSMyTable = null;
        //        SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
        //        return res;
        //    }
        //}

        public int obtenerDocEntryIVZ_ODEP()
        {
            int res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select  case when MAX(CAST(DocEntry AS Int)) is null then 1 else MAX(CAST(DocEntry AS Int)) + 1 end as Prox from [@IVZ_ODEP]";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt32(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
                return res;
            }
        }

        public int obtenerDocEntryIVZ_DEP1()
        {
            int res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select  case when MAX(CAST(DocEntry AS Int)) is null then 1 else MAX(CAST(DocEntry AS Int)) + 1 end as Prox from [@IVZ_DEP1]";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt32(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
                return res;
            }
        }

        public int obtenerDocEntryLiquidaciones()
        {
            int res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select  case when MAX(CAST(DocEntry AS Int)) is null then 1 else MAX(CAST(DocEntry AS Int)) + 1 end as Prox from [@LIQUIDACIONES]";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt32(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo DocEntry de LogProceso.");
                return res;
            }
        }

        public int obtenerCodeProxLiquidacion()
        {
            int res = 1;
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select case when MAX(CAST(Code AS Int)) is null then 1 else MAX(CAST(Code AS Int)) + 1 end as Prox from [@LIQUIDACIONES]";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = Convert.ToInt32(oRSMyTable.Fields.Item("Prox").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                SBO_Application.MessageBox("Ha ocurrido un error al buscar el proximo Code de Liquidaciones.");
                return res;
            }
        }

        // Cambia el estado de una Liquidación, recibiendo el codigo de la misma y el estado nuevo por parametro
        public Boolean cambiarEstadoLiquidacion(string pCodigoLiquidacion, string pEstado, DateTime pFechaCierre)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "update [@LIQUIDACIONES] set [U_ESTADO] = '" + pEstado + "' ";

                DateTime horaCierre = DateTime.Now;
                // ANTES. string ultimoCambio = pFechaCierre.ToString("yyyy-MM-dd") + " " + horaCierre.ToString("HH:mm"); // Guardo en una variable la Fecha y Hora del Cierre
                string ultimoCambio = horaCierre.ToString("yyyy-MM-dd") + " " + horaCierre.ToString("HH:mm"); // Guardo en una variable la Fecha y Hora del Cierre

                if (pEstado.ToString().Equals("2")) // Si se va a cerrar la liquidacion
                {
                    query += " , [U_FECHA_CIERRE] ='" + pFechaCierre.ToString("yyyy-MM-dd") + "', [U_ULTIMO_CAMBIO] ='" + ultimoCambio + "' ";
                    query += " , [U_HORA_CIERRE] = CASE WHEN [U_HORA_CIERRE] is null then '" + horaCierre.ToString("HH:mm") + "' ELSE [U_HORA_CIERRE] end ";
                }

                query += "where Code = '" + pCodigoLiquidacion + "'";
                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                guardaLogProceso(pCodigoLiquidacion, pCodigoLiquidacion, "ERROR al Cambiar estado liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                return false;
            }
        }

        // Obtener Datos de la Liquidacion
        public clsLiquidacion obtenerDatosLiquidacion(string pCodigoLiquidacion)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            clsLiquidacion liqui = new clsLiquidacion();
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "select Code, [U_FECHA], [U_ESTADO], [U_REPARTIDOR], case when [U_FECHA_CIERRE] is null then '01-01-1980' else [U_FECHA_CIERRE] end as FechaCierre from [@LIQUIDACIONES] where Code = '" + pCodigoLiquidacion + "'";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        liqui.Code = pCodigoLiquidacion;
                        liqui.Estado = Convert.ToInt32(oRSMyTable.Fields.Item("U_ESTADO").Value);
                        liqui.Repartidor = Convert.ToString(oRSMyTable.Fields.Item("U_REPARTIDOR").Value);
                        liqui.Fecha = Convert.ToDateTime(oRSMyTable.Fields.Item("U_FECHA").Value);
                        liqui.FechaCierre = Convert.ToDateTime(oRSMyTable.Fields.Item("FechaCierre").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return liqui;
            }
            catch (Exception ex)
            {
                guardaLogProceso(pCodigoLiquidacion, pCodigoLiquidacion, "ERROR al Obtener datos liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                return liqui;
            }
        }

        public Boolean asignarLiquidacion(clsDocumento pDocumento, string pNroLiquidacion, DateTime pFechaLiquidacion)
        {
            return asignarLiquidacionSDK(pDocumento, pNroLiquidacion, pFechaLiquidacion); // Se cambio para asignar por SDK el 04/02/19

            //SAPbobsCOM.Recordset oRSMyTable = null;
            //try
            //{
            //    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //    String query = "UPDATE OINV set U_LIQUIDACION = '" + pNroLiquidacion + "'";

            //    if (pDocumento.Tipo.ToString().Equals("NC"))
            //        query = "UPDATE ORIN set U_LIQUIDACION = '" + pNroLiquidacion + "'";

            //    if (pFechaLiquidacion != Convert.ToDateTime("01-01-1980"))
            //        query += " , U_FECHA_ENTREGA = '" + pFechaLiquidacion.ToString("yyyy-MM-dd") + "' ";

            //    query += " where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

            //    oRSMyTable.DoQuery(query);

            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
            //    oRSMyTable = null;
            //    return true;
            //}
            //catch (Exception ex)
            //{
            //    guardaLogProceso(pNroLiquidacion, pNroLiquidacion, "ERROR al Asignar liquidacion", ex.Message.ToString()); // Guarda log del Proceso
            //    return false;
            //}
        }

        public Boolean asignarLiquidacionSDK(clsDocumento pDocumento, string pNroLiquidacion, DateTime pFechaLiquidacion)
        {
            bool res = false;
            try
            {
                if (!pDocumento.Tipo.ToString().Equals("NC"))
                {
                    SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oInvoices);
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        oDoc.UserFields.Fields.Item("U_LIQUIDACION").Value = pNroLiquidacion;

                        if (pFechaLiquidacion != Convert.ToDateTime("01-01-1980"))
                            oDoc.UserFields.Fields.Item("U_FECHA_ENTREGA").Value = pFechaLiquidacion.ToString("yyyy-MM-dd");

                        int resSAP = oDoc.Update();
                        if (resSAP >= 0)
                            res = true;
                    }
                }
                else
                {
                    // Si es Nota de crédito
                    SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oCreditNotes);
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        oDoc.UserFields.Fields.Item("U_LIQUIDACION").Value = pNroLiquidacion;

                        if (pFechaLiquidacion != Convert.ToDateTime("01-01-1980"))
                            oDoc.UserFields.Fields.Item("U_FECHA_ENTREGA").Value = pFechaLiquidacion.ToString("yyyy-MM-dd");

                        int resSAP = oDoc.Update();
                        if (resSAP >= 0)
                            res = true;
                    }
                }

                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso(pNroLiquidacion, pNroLiquidacion, "ERROR al Asignar liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                return res;
            }
        }

        public Boolean depositarCheque(clsCheque pCheque, string pEstado, int pTransaccionID)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "delete from [@CHEQUES_ACREDITADOS] where Code = '" + pCheque.NumSecuencia + "' and DocEntry = '" + pCheque.NumSecuencia + "'";
                oRSMyTable.DoQuery(query); // Borro si hay registros para el actual Cheque

                query = "INSERT into [@CHEQUES_ACREDITADOS] (Code,DocEntry,U_ACREDITADO,U_TRANSID) values ('" + pCheque.NumSecuencia + "','" + pCheque.NumSecuencia + "','" + pEstado + "','" + pTransaccionID + "')";
                oRSMyTable.DoQuery(query); // Hago el insert con el Cheque

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean eliminarRegistrosRecientementeDepositados()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "delete from [@CHEQUES_ACREDITADOS] where U_TRANSID = '" + 999999999 + "'";
                oRSMyTable.DoQuery(query); // Borro si hay registros cobrados recientemente si el asiento no se guardo

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean actualizarRegistrosRecientementeDepositados()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                int transID = 0;
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "select top 1 TransId from ojdt where UserSign = '" + idUsuarioLogueado + "' and Memo = 'Acreditación de Cheques' order by TransId desc";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        transID = Convert.ToInt32(oRSMyTable.Fields.Item("TransId").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                query = "update [@CHEQUES_ACREDITADOS] set U_TRANSID = '" + transID + "' where U_TRANSID = '" + 999999999 + "'";
                oRSMyTable.DoQuery(query); // Actualizo el TransID si hay registros cobrados recientemente 

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean actualizarRegistrosIngresoCobros(string pLiquidacion, int pDocEntry, int pInvType)
        {
            //return actualizarRegistrosIngresoCobrosSDK(pLiquidacion, pDocEntry, pInvType); // Se cambio para asignar por SDK el 04/02/19
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                /*int docNum = 0;
                String query = "select top 1 DocNum from RCT2 where DocEntry = '" + pDocEntry + "' and InvType = '" + pInvType + "' order by DocEntry desc";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        docNum = Convert.ToInt32(oRSMyTable.Fields.Item("DocNum").Value);
                        oRSMyTable.MoveNext();
                    }
                }

                if (docNum != 0)
                {
                    query = "update ORCT set U_LIQUIDACION = '" + pLiquidacion + "' where DocNum = '" + docNum + "'";
                    oRSMyTable.DoQuery(query);
                }*/

                String query = "update ORCT set U_LIQUIDACION = '" + pLiquidacion + "' where DocNum = (select top 1 DocNum from RCT2 where DocEntry ='" + pDocEntry + "' and InvType = '" + pInvType + "' order by DocNum desc)";
                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                guardaLogProceso(pDocEntry.ToString() + pInvType.ToString(), pDocEntry.ToString() + pInvType.ToString(), "ERROR al actualizar el numero de liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                return false;
            }
        }
        public Boolean actualizarRegistrosIngresoCobrosSDK(string pLiquidacion, int pDocEntry, int pInvType)
        {
            bool res = false;
            try
            {
                SAPbobsCOM.Payments oDoc = oCompany.GetBusinessObject(BoObjectTypes.oIncomingPayments); // Creo el objeto Pago
                if (oDoc.GetByKey(pDocEntry))
                {
                    oDoc.UserFields.Fields.Item("U_LIQUIDACION").Value = pLiquidacion;

                    int resSAP = oDoc.Update();
                    if (resSAP >= 0)
                        res = true;
                }

                return res;
            }
            catch (Exception ex)
            {
                guardaLogProceso(pDocEntry.ToString() + pInvType.ToString(), pDocEntry.ToString() + pInvType.ToString(), "ERROR al actualizar el numero de liquidacion", ex.Message.ToString()); // Guarda log del Proceso
                return res;
            }
        }

        public Boolean asignarEstadoConfirmacionPedido(clsDocumento pDocumento, string pEstadoConfirmacion)
        {
            return asignarEstadoConfirmacionPedidoSDK(pDocumento, pEstadoConfirmacion); // Se cambio para asignar por SDK el 04/02/19
            //SAPbobsCOM.Recordset oRSMyTable = null;
            //try
            //{
            //    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //    String query = "UPDATE ORDR set Confirmed = '" + pEstadoConfirmacion + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

            //    if (pDocumento.Tipo.ToString().Equals("OF"))
            //        query = "UPDATE OQUT set Confirmed = '" + pEstadoConfirmacion + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

            //    oRSMyTable.DoQuery(query);

            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
            //    oRSMyTable = null;
            //    return true;
            //}
            //catch (Exception ex)
            //{
            //    return false;
            //}
        }

        public Boolean asignarEstadoConfirmacionPedidoSDK(clsDocumento pDocumento, string pEstadoConfirmacion)
        {
            bool res = false;
            try
            {
                if (!pDocumento.Tipo.ToString().Equals("OF"))
                {
                    SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oOrders); // Orden de Venta
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        if (pEstadoConfirmacion.Equals("Y"))
                        {
                            oDoc.Pick = BoYesNoEnum.tYES; //ASPL. 2019-09-05, cambio de campo Confirmed por Pick.
                            oDoc.Confirmed = BoYesNoEnum.tYES;
                        }
                        else
                        {
                            oDoc.Pick = BoYesNoEnum.tNO; //ASPL. 2019-09-05, cambio de campo Confirmed por Pick.
                            oDoc.Confirmed = BoYesNoEnum.tNO;
                        }

                        int resSAP = oDoc.Update();
                        if (resSAP >= 0)
                            res = true;
                    }
                }
                else
                {
                    SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oQuotations); // Oferta de Venta
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        if (pEstadoConfirmacion.Equals("Y"))
                        {
                            oDoc.Confirmed = BoYesNoEnum.tYES;
                            oDoc.Pick = BoYesNoEnum.tYES;
                        }
                        else
                        {
                            oDoc.Confirmed = BoYesNoEnum.tNO;
                            oDoc.Pick = BoYesNoEnum.tNO;
                        }
                        int resSAP = oDoc.Update();
                        if (resSAP >= 0)
                            res = true;
                    }
                }

                return res;
            }
            catch (Exception ex)
            { }
            return res;
        }

        public Boolean actualizarAcctCode(int pDocEntry)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                int docNum = 0;
                String query = "select TaxCode,AcctCode,LineNum from INV1 where DocEntry = '" + pDocEntry + "' order by LineNum";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        int lineNum = Convert.ToInt32(oRSMyTable.Fields.Item("LineNum").Value);
                        string acctCode = "4130010001"; // Codigo para IVA
                        string taxCode = oRSMyTable.Fields.Item("TaxCode").Value;
                        if (taxCode.ToString().Equals("IVA_EXE"))
                            acctCode = "4130010007"; // Codigo para IVA EXENTO
                        else if (taxCode.ToString().Equals("IVA_MIN"))
                            acctCode = "4130010002"; // Codigo para IVA MINIMO

                        query = "update INV1 set AcctCode = '" + acctCode + "' where DocEntry = '" + pDocEntry + "' and LineNum = '" + lineNum + "'";
                        SAPbobsCOM.Recordset oRSMyTableLine = null;
                        oRSMyTableLine = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        oRSMyTableLine.DoQuery(query); // Actualizo el campo
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTableLine);
                        oRSMyTableLine = null;

                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                guardaLogProceso(pDocEntry.ToString(), pDocEntry.ToString(), "ERROR al actualizar el AcctCode de documento", ex.Message.ToString()); // Guarda log del Proceso
                return false;
            }
        }

        //public Boolean asignarEstadoConfirmacionPedido(clsDocumento pDocumento, string pEstadoConfirmacion)
        //{
        //    try
        //    {
        //        if (pDocumento.Tipo.ToString().Equals("OF")) // Oferta de Ventas
        //        {
        //            SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oQuotations);

        //            if (oDoc.GetByKey(pDocumento.DocEntry))
        //            {
        //                if (pEstadoConfirmacion.ToString().Equals("Y"))
        //                    oDoc.Confirmed = BoYesNoEnum.tYES;
        //                else
        //                    oDoc.Confirmed = BoYesNoEnum.tNO;
        //                int resTrans = oDoc.Update();
        //                if (resTrans < 0)
        //                    return false;
        //            }
        //        }
        //        else
        //        {
        //            // Pedidos
        //            SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oOrders);

        //            if (oDoc.GetByKey(pDocumento.DocEntry))
        //            {
        //                if (pEstadoConfirmacion.ToString().Equals("Y"))
        //                    oDoc.Confirmed = BoYesNoEnum.tYES;
        //                else
        //                    oDoc.Confirmed = BoYesNoEnum.tNO;
        //                int resTrans = oDoc.Update();
        //                if (resTrans < 0)
        //                    return false;
        //            }
        //        }

        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}

        public clsDocumento asignarExtraDays(clsDocumento pDocumento)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                DayOfWeek dia = DateTime.Now.DayOfWeek; // Obtrngo el dia de la semana correspondiente al día de HOY
                int diaActual = Convert.ToInt32(dia); // Lo convierto a numero, 1 a 5 Lunes a Viernes
                int extraDays = 1; int extraMonth = 0;

                if (diaActual == 5) // Si es un Viernes
                    extraDays = 3; // Los ExtraDays son 3

                string tabla = "ORDR";

                if (pDocumento.Tipo.ToString().Equals("OF"))
                    tabla = "OQUT";

                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                int mesCondPago = 0; int daysCondPago = 0;

                String query = "select oct.ExtraDays,oct.ExtraMonth from OCTG as oct " +
               "inner join " + tabla + " as ord on ord.GroupNum = oct.GroupNum where ord.DocEntry = " + pDocumento.DocEntry + "  and ord.DocNum = " + pDocumento.DocNum;
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        mesCondPago = (int)oRSMyTable.Fields.Item("ExtraMonth").Value; // Obtengo el Mes Extra de la condicion de pago
                        daysCondPago = (int)oRSMyTable.Fields.Item("ExtraDays").Value; // Obtengo los dias Extra de la condicion de pago
                        oRSMyTable.MoveNext();
                    }
                }

                extraDays += daysCondPago;
                extraMonth += mesCondPago;

                pDocumento.ExtraDays = extraDays;
                pDocumento.ExtraMonth = extraMonth;

                //query = "UPDATE ORDR set ExtraDays  = '" + extraDays + "',ExtraMonth = '" + extraMonth + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

                //if (pDocumento.Tipo.ToString().Equals("OF"))
                //    query = "UPDATE OQUT set ExtraDays  = '" + extraDays + "',ExtraMonth = '" + extraMonth + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

                //oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return pDocumento;
            }
            catch (Exception ex)
            {
                return pDocumento;
            }
        }

        //public Boolean asignarExtraDays(clsDocumento pDocumento)
        //{
        //    SAPbobsCOM.Recordset oRSMyTable = null;
        //    try
        //    {
        //        DayOfWeek dia = DateTime.Now.DayOfWeek; // Obtrngo el dia de la semana correspondiente al día de HOY
        //        int diaActual = Convert.ToInt32(dia); // Lo convierto a numero, 1 a 5 Lunes a Viernes
        //        int extraDays = 1; int extraMonth = 0;

        //        if (diaActual == 5) // Si es un Viernes
        //            extraDays = 3; // Los ExtraDays son 3

        //        string tabla = "ORDR";

        //        if (pDocumento.Tipo.ToString().Equals("OF"))
        //            tabla = "OQUT";

        //        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //        int mesCondPago = 0; int daysCondPago = 0;

        //        String query = "select oct.ExtraDays,oct.ExtraMonth from OCTG as oct " +
        //       "inner join " + tabla + " as ord on ord.GroupNum = oct.GroupNum where ord.DocEntry = " + pDocumento.DocEntry + "  and ord.DocNum = " + pDocumento.DocNum;
        //        oRSMyTable.DoQuery(query);

        //        if (oRSMyTable != null)
        //        {
        //            while (!oRSMyTable.EoF)
        //            {
        //                mesCondPago = (int)oRSMyTable.Fields.Item("ExtraMonth").Value; // Obtengo el Mes Extra de la condicion de pago
        //                daysCondPago = (int)oRSMyTable.Fields.Item("ExtraDays").Value; // Obtengo los dias Extra de la condicion de pago
        //                oRSMyTable.MoveNext();
        //            }
        //        }

        //        extraDays += daysCondPago;
        //        extraMonth += mesCondPago;

        //        query = "UPDATE ORDR set ExtraDays  = '" + extraDays + "',ExtraMonth = '" + extraMonth + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

        //        if (pDocumento.Tipo.ToString().Equals("OF"))
        //            query = "UPDATE OQUT set ExtraDays  = '" + extraDays + "',ExtraMonth = '" + extraMonth + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

        //        oRSMyTable.DoQuery(query);

        //        System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
        //        oRSMyTable = null;
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}

        //public Boolean asignarDiscPrcnt(clsDocumento pDocumento)
        //{
        //    try
        //    {
        //        if (pDocumento.Tipo.ToString().Equals("OF")) // Si es una Oferta de Venta
        //        {
        //            string motVenta = obtenerMotivoVentaOfertaVenta(pDocumento);
        //            if (!String.IsNullOrEmpty(motVenta.ToString())) // Si tiene algún valor el campo U_MOTIVO_VENTA0 
        //            {
        //                SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oQuotations);

        //                if (oDoc.GetByKey(pDocumento.DocEntry))
        //                {
        //                    //oDoc.DiscountPercent = Convert.ToInt32(100); // Asigno descuento del 100 % para que la Oferta de Venta quede con valor 0
        //                    oDoc.Reference2 = "50";
        //                    int resTrans = oDoc.Update();
        //                    if (resTrans < 0)
        //                    {
        //                        SBO_Application.MessageBox(resTrans.ToString());
        //                        /*string error = oCompany.GetLastErrorDescription();
        //                        SBO_Application.MessageBox(error);*/
        //                        return false;
        //                    }
        //                }

        //            }
        //        }
        //        return true;
        //    }
        //    catch (Exception ex)
        //    {
        //        return false;
        //    }
        //}

        public Boolean asignarDiscPrcnt(clsDocumento pDocumento)
        {
            SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            try
            {
                if (pDocumento.Tipo.ToString().Equals("OF")) // Si es una Oferta de Venta
                {
                    string motVenta = obtenerMotivoVentaOfertaVenta(pDocumento);
                    if (!String.IsNullOrEmpty(motVenta.ToString())) // Si tiene algún valor el campo U_MOTIVO_VENTA0 
                    {
                        int DiscPrcnt = 100; // Asigno descuento del 100 % para que la Oferta de Venta quede con valor 0
                        decimal totales = 0; // Asigno los totales del documento y los totales de impuesto en 0

                        oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        String query = "UPDATE OQUT set DiscPrcnt  = '" + DiscPrcnt + "',VatSum = '" + totales + "', VatSumFC = '" + totales + "',DocTotal ='" + totales + "',DocTotalFC= '" + totales + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;

                        oRSMyTable.DoQuery(query);
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean asignarCentroCostos(clsDocumento pDocumento)
        {
            SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset oRSMyTable2 = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string centroCosto = "";
            try
            {
                if (pDocumento.Tipo.ToString().Equals("PE")) // Si es un Pedido
                {
                    String query = "select U_CENTRO_COSTO as CC from OCRD as ocr inner join [@CANAL] as can on can.Code = ocr.U_CANAL " +
                    "inner join ORDR as ord on ord.CardCode = ocr.CardCode where ocr.CardType = 'C' and ord.DocEntry = " + pDocumento.DocEntry + "  and ord.DocNum = " + pDocumento.DocNum;
                    oRSMyTable.DoQuery(query);

                    if (oRSMyTable != null)
                    {
                        while (!oRSMyTable.EoF)
                        {
                            centroCosto = oRSMyTable.Fields.Item("CC").Value; // Obtengo el Centro de Costo correspondiente al Canal del cliente
                            oRSMyTable.MoveNext();
                        }
                    }

                    query = "select COUNT(*) as Cant from RDR1 where DocEntry =" + pDocumento.DocEntry + " AND (Ocrcode is NULL or OcrCode = '')";
                    oRSMyTable2.DoQuery(query);
                    if (oRSMyTable2 != null)
                    {
                        if (oRSMyTable2.RecordCount != 0) // Si hay alguna línea que no tenga el Centro de Costo asignado
                        {
                            if (Convert.ToInt32(oRSMyTable2.Fields.Item("Cant").Value) > 0)
                            {
                                query = "select rdr.ItemCode, rdr.LineNum, oit.U_CC as CC,rdr.CogsOcrCod as Co, rdr.CogsOcrCo2 as Co2 from RDR1 as rdr left join OITM as oit on oit.ItemCode = rdr.ItemCode " +
                                "where rdr.DocEntry =" + pDocumento.DocEntry + "  order by rdr.LineNum"; // Recorro los articulos del pedido y ya me traigo el centro de costo de cada uno de ellos
                                oRSMyTable2.DoQuery(query);

                                string itemCode = ""; // ItemCode del articulo
                                string centroCostoArticulo = ""; // Centro de Costo de cada articulo
                                string codigoCo = ""; // Codigo CogsOcrCod
                                string codigoCo2 = ""; // Codigo CogsOcrCo2
                                int numeroLinea = 0; // Numero de linea del Pedido
                                if (oRSMyTable2 != null)
                                {
                                    while (!oRSMyTable2.EoF)
                                    {
                                        itemCode = oRSMyTable2.Fields.Item("ItemCode").Value; // Obtengo el ItemCode de la línea
                                        centroCostoArticulo = oRSMyTable2.Fields.Item("CC").Value; // Obtengo el Centro Costo de la línea
                                        numeroLinea = Convert.ToInt32(oRSMyTable2.Fields.Item("LineNum").Value);
                                        codigoCo = oRSMyTable2.Fields.Item("Co").Value; // Obtengo el CogsOcrCod de la línea
                                        codigoCo2 = oRSMyTable2.Fields.Item("Co2").Value; // Obtengo el CogsOcrCo2 de la línea

                                        actualizarLineaDocumentoPorSDK(pDocumento, centroCosto, centroCostoArticulo, itemCode, numeroLinea, codigoCo, codigoCo2); // Actualizo los campos CogsOcrCod y CogsOcrCo2
                                        oRSMyTable2.MoveNext();
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    // Si es una Oferta de Venta
                    String query = "select U_CENTRO_COSTO as CC from OCRD as ocr inner join [@CANAL] as can on can.Code = ocr.U_CANAL " +
                    "inner join OQUT as ord on ord.CardCode = ocr.CardCode where ocr.CardType = 'C' and ord.DocEntry = " + pDocumento.DocEntry + "  and ord.DocNum = " + pDocumento.DocNum;
                    oRSMyTable.DoQuery(query);

                    if (oRSMyTable != null)
                    {
                        while (!oRSMyTable.EoF)
                        {
                            centroCosto = oRSMyTable.Fields.Item("CC").Value; // Obtengo el Centro de Costo correspondiente al Canal del cliente
                            oRSMyTable.MoveNext();
                        }
                    }


                    query = "select COUNT(*) as Cant from QUT1 where DocEntry =" + pDocumento.DocEntry + " AND (Ocrcode is NULL or OcrCode = '')";
                    oRSMyTable2.DoQuery(query);
                    if (oRSMyTable2 != null)
                    {
                        if (oRSMyTable2.RecordCount != 0) // Si hay alguna línea que no tenga el Centro de Costo asignado
                        {
                            if (Convert.ToInt32(oRSMyTable2.Fields.Item("Cant").Value) > 0)
                            {
                                query = "select rdr.ItemCode, rdr.LineNum, oit.U_CC as CC, rdr.CogsOcrCod as Cod, rdr.CogsOcrCo2 as Co from QUT1 as rdr left join OITM as oit on oit.ItemCode = rdr.ItemCode " +
                                "where rdr.DocEntry =" + pDocumento.DocEntry + "  order by rdr.LineNum"; // Recorro los articulos del pedido y ya me traigo el centro de costo de cada uno de ellos
                                oRSMyTable2.DoQuery(query);

                                string itemCode = ""; // ItemCode del articulo
                                string centroCostoArticulo = ""; // Centro de Costo de cada articulo
                                int numeroLinea = 0; // Numero de linea del Pedido
                                string codigoCo = ""; // Codigo CogsOcrCod
                                string codigoCo2 = ""; // Codigo CogsOcrCo2
                                if (oRSMyTable2 != null)
                                {
                                    while (!oRSMyTable2.EoF)
                                    {
                                        itemCode = oRSMyTable2.Fields.Item("ItemCode").Value; // Obtengo el ItemCode de la línea
                                        centroCostoArticulo = oRSMyTable2.Fields.Item("CC").Value; // Obtengo el Centro Costo de la línea
                                        numeroLinea = Convert.ToInt32(oRSMyTable2.Fields.Item("LineNum").Value);
                                        codigoCo = oRSMyTable2.Fields.Item("Co").Value; // Obtengo el CogsOcrCod de la línea
                                        codigoCo2 = oRSMyTable2.Fields.Item("Co2").Value; // Obtengo el CogsOcrCo2 de la línea

                                        actualizarLineaDocumentoPorSDK(pDocumento, centroCosto, centroCostoArticulo, itemCode, numeroLinea, codigoCo, codigoCo2); // Actualizo los campos CogsOcrCod y CogsOcrCo2
                                        oRSMyTable2.MoveNext();
                                    }
                                }
                            }
                        }
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable2);
                oRSMyTable2 = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean actualizarLineaDocumento(clsDocumento pDocumento, string pCentroCosto, string pCentroCostoArticulo, string pItemCode, int pNumeroLinea, string pCogsOcrCo, string pCogsOcrCo2) // Actualizo los campos CogsOcrCod y CogsOcrCo2
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "UPDATE RDR1 set CogsOcrCod  = '" + pCentroCosto + "', CogsOcrCo2 = '" + pCentroCostoArticulo + "', OcrCode = '" + pCogsOcrCo + "', OcrCode2 = '" + pCogsOcrCo2 + "' where DocEntry = " + pDocumento.DocEntry + " and LineNum = " + pNumeroLinea + " and ItemCode ='" + pItemCode + "'";

                if (pDocumento.Tipo.ToString().Equals("OF"))
                    query = "UPDATE QUT1 set CogsOcrCod  = '" + pCentroCosto + "', CogsOcrCo2 = '" + pCentroCostoArticulo + "', OcrCode = '" + pCogsOcrCo + "', OcrCode2 = '" + pCogsOcrCo2 + "' where DocEntry = " + pDocumento.DocEntry + " and LineNum = " + pNumeroLinea + " and ItemCode ='" + pItemCode + "'";

                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean actualizarLineaDocumentoPorSDK(clsDocumento pDocumento, string pCentroCosto, string pCentroCostoArticulo, string pItemCode, int pNumeroLinea, string pCogsOcrCo, string pCogsOcrCo2) // Actualizo los campos CogsOcrCod y CogsOcrCo2
        {

            bool resultado = false;
            try
            {
                //Para los centros de costo: OcrCode = CogsOcrCode y OcrCode2 = CogsOcrCod2. 
                //El add-on asigna correctamente los campos CogsOcrCod y CogsOcrCod2 en el visor de pedidos pero pone el mismo valor en los campos OcrCode y OcrCode2.

                SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oOrders);

                if (pDocumento.DocEntry != 0)
                {
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        for (int i = oDoc.Lines.Count - 1; i >= 0; i--)
                        {
                            oDoc.Lines.SetCurrentLine(i);
                            if (oDoc.Lines.LineNum == pNumeroLinea)
                            {
                                oDoc.Lines.COGSCostingCode = pCentroCosto;
                                oDoc.Lines.COGSCostingCode2 = pCentroCostoArticulo;
                                oDoc.Lines.CostingCode = pCogsOcrCo;
                                oDoc.Lines.CostingCode2 = pCogsOcrCo2;
                                int a = oDoc.Update();
                                if (a == 0)
                                    resultado = true;
                                //else
                                //{
                                //    string error = oCompany.GetLastErrorDescription();
                                //    SBO_Application.MessageBox(error + ". Codigo error:" + a.ToString());
                                //}
                            }
                        }
                    }
                }

                if (resultado == false) // Si no lo pudo agregar por el SDK
                {
                    SAPbobsCOM.Recordset oRSMyTable = null;

                    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    String query = "UPDATE RDR1 set CogsOcrCod  = '" + pCentroCosto + "', CogsOcrCo2 = '" + pCentroCostoArticulo + "', OcrCode = '" + pCogsOcrCo + "', OcrCode2 = '" + pCogsOcrCo2 + "' where DocEntry = " + pDocumento.DocEntry + " and LineNum = " + pNumeroLinea + " and ItemCode ='" + pItemCode + "'";

                    if (pDocumento.Tipo.ToString().Equals("OF"))
                        query = "UPDATE QUT1 set CogsOcrCod  = '" + pCentroCosto + "', CogsOcrCo2 = '" + pCentroCostoArticulo + "', OcrCode = '" + pCogsOcrCo + "', OcrCode2 = '" + pCogsOcrCo2 + "' where DocEntry = " + pDocumento.DocEntry + " and LineNum = " + pNumeroLinea + " and ItemCode ='" + pItemCode + "'";

                    oRSMyTable.DoQuery(query);

                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                    oRSMyTable = null;
                }
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public Boolean asignarFechaEntrega(clsDocumento pDocumento)
        {
            return asignarFechaEntregaSDK(pDocumento);// Se cambio para asignar por SDK el 04/02/19
            //SAPbobsCOM.Recordset oRSMyTable = null;
            //try
            //{
            //    oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            //    string tabla = "ORDR";

            //    if (pDocumento.Tipo.ToString().Equals("OF"))
            //        tabla = "OQUT";

            //    string query = "UPDATE " + tabla + " set U_FECHA_ENTREGA  = DocDueDate where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;
            //    oRSMyTable.DoQuery(query);

            //    System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
            //    oRSMyTable = null;
            //    return true;
            //}
            //catch (Exception ex)
            //{
            //    return false;
            //}
        }
        public Boolean asignarFechaEntregaSDK(clsDocumento pDocumento)
        {
            bool res = false;
            try
            {
                if (!pDocumento.Tipo.ToString().Equals("OF"))
                {
                    SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oOrders); // Orden de Venta
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        oDoc.UserFields.Fields.Item("U_FECHA_ENTREGA").Value = oDoc.DocDueDate;

                        int resSAP = oDoc.Update();
                        if (resSAP >= 0)
                            res = true;
                    }
                }
                else
                {
                    SAPbobsCOM.Documents oDoc = oCompany.GetBusinessObject(BoObjectTypes.oQuotations); // Oferta de Venta
                    if (oDoc.GetByKey(pDocumento.DocEntry))
                    {
                        oDoc.UserFields.Fields.Item("U_FECHA_ENTREGA").Value = oDoc.DocDueDate;

                        int resSAP = oDoc.Update();
                        if (resSAP >= 0)
                            res = true;
                    }
                }

                return res;
            }
            catch (Exception ex)
            {
                return res;
            }
        }

        public Boolean asignarTerritorioDocumento(clsDocumento pDocumento)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string territorio = ""; string tabla = "ORDR";

                if (pDocumento.Tipo.ToString().Equals("OF"))
                    tabla = "OQUT";

                String query = "SELECT descript as Name FROM OTER as ote " +
                "inner join OCRD as ocr on ocr.Territory = ote.territryID inner join " + tabla + " as ord on ord.CardCode = ocr.CardCode " +
                "where ocr.CardType = 'C' and ord.DocEntry = " + pDocumento.DocEntry + "  and ord.DocNum = " + pDocumento.DocNum;
                oRSMyTable.DoQuery(query);
                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        territorio = oRSMyTable.Fields.Item("Name").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                if (territorio.ToString().Length > 4)
                    territorio = territorio.Substring(0, 4);

                query = "UPDATE " + tabla + " set U_TERRITORIO  = '" + territorio + "' where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;
                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        // Función mejorada para reducir los query
        public Boolean asignarFechaEntregaTerritorioExtraDaysDiscPrcnt(clsDocumento pDocumento)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string territorio = ""; string tabla = "ORDR";

                string queryOF = ""; // Query para cuando es una Oferta de Venta y tiene un Motivo de Venta
                // COMENTADO POR JUAN
                /*if (pDocumento.Tipo.ToString().Equals("OF")) // Si es una Oferta de Venta
                {
                    tabla = "OQUT";
                    string motVenta = obtenerMotivoVentaOfertaVenta(pDocumento);
                    if (!String.IsNullOrEmpty(motVenta.ToString())) // Si tiene algún valor el campo U_MOTIVO_VENTA0 
                    {
                        int DiscPrcnt = 100; // Asigno descuento del 100 % para que la Oferta de Venta quede con valor 0
                        decimal totales = 0; // Asigno los totales del documento y los totales de impuesto en 0

                        queryOF = ",DiscPrcnt  = '" + DiscPrcnt + "',VatSum = '" + totales + "', VatSumFC = '" + totales + "',DocTotal ='" + totales + "',DocTotalFC= '" + totales + "' ";
                    }
                }*/


                String query = "SELECT descript as Name FROM OTER as ote " +
                "inner join OCRD as ocr on ocr.Territory = ote.territryID inner join " + tabla + " as ord on ord.CardCode = ocr.CardCode " +
                "where ocr.CardType = 'C' and ord.DocEntry = " + pDocumento.DocEntry + "  and ord.DocNum = " + pDocumento.DocNum;
                oRSMyTable.DoQuery(query);
                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        territorio = oRSMyTable.Fields.Item("Name").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                if (territorio.ToString().Length > 4)
                    territorio = territorio.Substring(0, 4);


                query = "UPDATE " + tabla + " set U_TERRITORIO  = '" + territorio + "', U_FECHA_ENTREGA  = DocDueDate, ExtraDays  = '" + pDocumento.ExtraDays + "',ExtraMonth = '" + pDocumento.ExtraMonth + "' " + queryOF.ToString() + "  where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;
                oRSMyTable.DoQuery(query);

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }

        public string obtenerMotivoVentaOfertaVenta(clsDocumento pDocumento)
        {
            string res = "";
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select U_MOTIVO_VENTA0 as MotVent from OQUT where DocEntry = " + pDocumento.DocEntry + " and DocNum = " + pDocumento.DocNum;
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = oRSMyTable.Fields.Item("MotVent").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;

                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
        }

        // Esta funcion recibe por parámentro el CurrCode de la moneda y busca el codigo ISO de la misma
        public string obtenerCodigoISOMoneda(String pCodigoISO)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            string res = pCodigoISO;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "select ISOCurrCod from OCRN where CurrCode = '" + pCodigoISO + "'";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = (string)oRSMyTable.Fields.Item("ISOCurrCod").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
        }

        // Esta funcion obtiene la Fecha de la Liquidacion
        public DateTime obtenerFechaLiquidacion(String pCodigoLiquidacion)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            DateTime res = DateTime.Now;
            try
            {
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                String query = "SELECT U_FECHA FROM [@LIQUIDACIONES] where Code = '" + pCodigoLiquidacion + "'";
                oRSMyTable.DoQuery(query);

                if (oRSMyTable != null)
                {
                    while (!oRSMyTable.EoF)
                    {
                        res = (DateTime)oRSMyTable.Fields.Item("U_FECHA").Value;
                        oRSMyTable.MoveNext();
                    }
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return res;
            }
        }
        #endregion

        #region Formularios

        private void CargarFormulario()
        {
            try
            {
                oFormDatosPedido = SBO_Application.Forms.Item("VisualizarLiquidaciones");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisualizarLiquidaciones";
                fcp.UniqueID = "VisualizarLiquidaciones";
                try
                {
                    fcp.XmlData = LoadFromXML("VisualizarLiquidaciones.srf");

                    oFormDatosPedido = SBO_Application.Forms.AddEx(fcp);

                    oFormDatosPedido.State = BoFormStateEnum.fs_Maximized;
                }
                catch (Exception exe)
                { }
            }

            try
            {
                DateTime fechaHoy = DateTime.Now;
                DateTime fechaAntes = DateTime.Now.AddDays(-1);

                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormDatosPedido.Items.Item("29").Specific;
                llenarCombo(oStaticCombo, "Select Code,Name from [@LIQUIDACIONES] where U_ESTADO <> '2' order by CAST(Code AS Int)", true, false);
                oStaticCombo = oFormDatosPedido.Items.Item("28").Specific;
                llenarComboEstadosLiquidacion(oStaticCombo, true);
                oStaticCombo = oFormDatosPedido.Items.Item("1000001").Specific;
                llenarCombo(oStaticCombo, "Select Code,Name from [@LIQUIDACIONES] where U_ESTADO <> '2' order by CAST(Code AS Int)", true, false);
                oStaticCombo = oFormDatosPedido.Items.Item("35").Specific;
                //llenarCombo(oStaticCombo, "SELECT DISTINCT U_PZSourceId as Code, U_PZSourceId as Name FROM OINV WHERE U_PZCreated = 'Y' AND U_PZSourceType LIKE '7%' ORDER BY U_PZSourceId DESC", false, false);
                llenarCombo(oStaticCombo, "select distinct T0.U_PZSourceId as Code, T0.U_PZSourceId as Name from oinv T0 left join[@LIQUIDACIONES] T1 on T1.Code = T0.U_LIQUIDACION where T0.CANCELED = 'N' and(T0.U_LIQUIDACION is null or (T0.U_LIQUIDACION is not null and T1.U_ESTADO= 1)) and T0.U_PZSourceID <> '18201' and T0.U_PZSourceType like '7%' order by U_PZSourceId", false, false);

                SAPbouiCOM.EditText oStatic;
                oFormDatosPedido.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormDatosPedido.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oStatic = oFormDatosPedido.Items.Item("16").Specific; // Desde Fecha
                oStatic.DataBind.SetBound(true, "", "Date1");
                oStatic = oFormDatosPedido.Items.Item("10").Specific; // Hasta Fecha
                oStatic.DataBind.SetBound(true, "", "Date2");

                //CargarGrilla();
                CargarMatrixTerritoriosVisorLiquidaciones();
                oFormDatosPedido.Visible = true;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso("", "", "ERROR al Cargar Grilla", ex.Message.ToString()); // Guarda log del Proceso
            }
        }

        private void CargarFormularioPedidosVisor()
        {
            try
            {
                oFormDatosPedidoVisor = SBO_Application.Forms.Item("VisualizarPedidos");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisualizarPedidos";
                fcp.UniqueID = "VisualizarPedidos";
                try
                {
                    fcp.XmlData = LoadFromXML("VisualizarPedidos.srf");

                    oFormDatosPedidoVisor = SBO_Application.Forms.AddEx(fcp);

                    oFormDatosPedidoVisor.State = BoFormStateEnum.fs_Maximized;
                }
                catch (Exception exe)
                { }
            }

            try
            {
                DateTime fechaHoy = DateTime.Now;
                DateTime fechaAntes = DateTime.Now.AddDays(-1);

                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormDatosPedidoVisor.Items.Item("28").Specific;
                //llenarComboEstadosConfirmacionPedido(oStaticCombo, true);
                llenarComboEstadosPickeadoPedido(oStaticCombo, true);
                oStaticCombo = oFormDatosPedidoVisor.Items.Item("1000001").Specific;
                llenarComboEstadosConfirmacionPedido(oStaticCombo, true);
                oStaticCombo = oFormDatosPedidoVisor.Items.Item("1000002").Specific;
                llenarComboPedidosFreeShop(oStaticCombo);

                SAPbouiCOM.EditText oStatic;
                oFormDatosPedidoVisor.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormDatosPedidoVisor.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oStatic = oFormDatosPedidoVisor.Items.Item("16").Specific; // Desde Fecha
                oStatic.DataBind.SetBound(true, "", "Date1");
                oStatic = oFormDatosPedidoVisor.Items.Item("10").Specific; // Hasta Fecha
                oStatic.DataBind.SetBound(true, "", "Date2");

                //CargarGrillaPedidosVisor();
                CargarMatrixTerritoriosVisorPedidos();
                oFormDatosPedidoVisor.Visible = true;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso("", "", "ERROR al Cargar Grilla", ex.Message.ToString()); // Guarda log del Proceso
            }
        }

        private void CargarFormularioIngresoCobros()
        {
            try
            {
                oFormDatosIngresoCobros = SBO_Application.Forms.Item("IngresoCobros");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "IngresoCobros";
                fcp.UniqueID = "IngresoCobros";
                try
                {
                    fcp.XmlData = LoadFromXML("VisualizarCobros.srf");

                    oFormDatosIngresoCobros = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                DateTime fechaHoy = DateTime.Now;

                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormDatosIngresoCobros.Items.Item("29").Specific;
                llenarCombo(oStaticCombo, "Select Code,Name from [@LIQUIDACIONES] where U_ESTADO <> '2' order by Code", false, false);
                oStaticCombo = oFormDatosIngresoCobros.Items.Item("1000001").Specific;
                llenarCombo(oStaticCombo, "select CurrCode as Code,CurrName as Name from OCRN order by CurrName", false, false);
                oStaticCombo = oFormDatosIngresoCobros.Items.Item("51").Specific;
                llenarCombo(oStaticCombo, "select AcctCode as Code, AcctName as Name from OACT where Finanse = 'Y' order by Name", false, false);

                SAPbouiCOM.EditText oStatic;
                oFormDatosIngresoCobros.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormDatosIngresoCobros.DataSources.UserDataSources.Add("Numero1", SAPbouiCOM.BoDataType.dt_PRICE, 15);
                oStatic = oFormDatosIngresoCobros.Items.Item("10").Specific; // Fecha
                oStatic.DataBind.SetBound(true, "", "Date1");
                oStatic = oFormDatosIngresoCobros.Items.Item("66").Specific; // Importe
                oStatic.DataBind.SetBound(true, "", "Numero1");
                /*string fechaStr = fechaHoy.ToString("dd/MM/yyyy");
                oStatic.Value = fechaStr;*/

                SAPbouiCOM.Button oStaticText;
                oStaticText = oFormDatosIngresoCobros.Items.Item("8").Specific;
                oStaticText.Item.Enabled = false;

                oFormDatosIngresoCobros.Visible = true;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso("", "", "ERROR al Cargar Grilla", ex.Message.ToString()); // Guarda log del Proceso
            }
        }

        private void CargarFormularioCheques()
        {
            try
            {
                oFormDatosCheques = SBO_Application.Forms.Item("VisualizarCheques");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "VisualizarCheques";
                fcp.UniqueID = "VisualizarCheques";
                try
                {
                    fcp.XmlData = LoadFromXML("VisualizarCheques.srf");

                    oFormDatosCheques = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormDatosCheques.Items.Item("28").Specific;
                llenarCombo(oStaticCombo, "select AcctCode as Code, AcctName as Name from OACT where Finanse = 'Y' order by Name", false, false);
                oStaticCombo = oFormDatosCheques.Items.Item("1000001").Specific;
                llenarCombo(oStaticCombo, "select AcctCode as Code, AcctName as Name from OACT where Finanse = 'Y' order by Name", false, false);

                SAPbouiCOM.EditText oStatic;
                oFormDatosCheques.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oStatic = oFormDatosCheques.Items.Item("10").Specific; // Desde Fecha
                oStatic.DataBind.SetBound(true, "", "Date1");

                /*DateTime fechaHoy = DateTime.Now;

                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormDatosIngresoCobros.Items.Item("29").Specific;
                llenarCombo(oStaticCombo, "Select Code,Name from [@LIQUIDACIONES] where U_ESTADO <> '2' order by Code", false);
                oStaticCombo = oFormDatosIngresoCobros.Items.Item("1000001").Specific;
                llenarCombo(oStaticCombo, "select CurrCode as Code,CurrName as Name from OCRN order by CurrName", false);
                oStaticCombo = oFormDatosIngresoCobros.Items.Item("51").Specific;
                llenarCombo(oStaticCombo, "select AcctCode as Code, AcctName as Name from OACT where Finanse = 'Y' order by Name", false);

                SAPbouiCOM.EditText oStatic;
                oFormDatosIngresoCobros.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oFormDatosIngresoCobros.DataSources.UserDataSources.Add("Numero1", SAPbouiCOM.BoDataType.dt_PRICE, 15);
                oStatic = oFormDatosIngresoCobros.Items.Item("10").Specific; // Fecha
                oStatic.DataBind.SetBound(true, "", "Date1");
                oStatic = oFormDatosIngresoCobros.Items.Item("66").Specific; // Importe
                oStatic.DataBind.SetBound(true, "", "Numero1");
                /*string fechaStr = fechaHoy.ToString("dd/MM/yyyy");
                oStatic.Value = fechaStr;*/

                CargarGrillaChequesVisor(true);
                oFormDatosCheques.Visible = true;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso("", "", "ERROR al Cargar Grilla", ex.Message.ToString()); // Guarda log del Proceso
            }
        }

        private void CargarFormularioIngresoLiquidaciones()
        {
            try
            {
                oFormDatosIngresoLiquidaciones = SBO_Application.Forms.Item("IngresoLiquidaciones");
            }
            catch (Exception ex)
            {
                SAPbouiCOM.FormCreationParams fcp;
                fcp = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                fcp.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Sizable;
                fcp.FormType = "IngresoLiquidaciones";
                fcp.UniqueID = "IngresoLiquidaciones";
                try
                {
                    fcp.XmlData = LoadFromXML("IngresoLiquidaciones.srf");

                    oFormDatosIngresoLiquidaciones = SBO_Application.Forms.AddEx(fcp);
                }
                catch (Exception exe)
                { }
            }

            try
            {
                DateTime fechaHoy = DateTime.Now;

                // Lleno los Combos de Repartidores
                SAPbouiCOM.ComboBox oStaticCombo;
                oStaticCombo = oFormDatosIngresoLiquidaciones.Items.Item("5").Specific;
                llenarCombo(oStaticCombo, "select Code,Name from [@REPARTIDORES] order by Name", false, false);
                oStaticCombo = oFormDatosIngresoLiquidaciones.Items.Item("14").Specific;
                llenarCombo(oStaticCombo, "select Code,Name from [@REPARTIDORES] order by Name", false, false);
                // Lleno el combo de Estados de liquidaciones
                oStaticCombo = oFormDatosIngresoLiquidaciones.Items.Item("13").Specific;
                llenarComboEstadosLiquidacion(oStaticCombo, true);

                SAPbouiCOM.EditText oStatic;
                oFormDatosIngresoLiquidaciones.DataSources.UserDataSources.Add("Date1", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oStatic = oFormDatosIngresoLiquidaciones.Items.Item("6").Specific; // Desde Fecha
                oStatic.DataBind.SetBound(true, "", "Date1");

                oFormDatosIngresoLiquidaciones.DataSources.UserDataSources.Add("Date2", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oStatic = oFormDatosIngresoLiquidaciones.Items.Item("9").Specific; // Fecha de Filtro
                oStatic.DataBind.SetBound(true, "", "Date2");

                oFormDatosIngresoLiquidaciones.DataSources.UserDataSources.Add("Date3", SAPbouiCOM.BoDataType.dt_DATE, 10);
                oStatic = oFormDatosIngresoLiquidaciones.Items.Item("19").Specific; // Fecha de Cierre de liquidacion
                oStatic.DataBind.SetBound(true, "", "Date3");

                // Caja de Fecha de cierre visible solo para los superusuarios
                if (esSuperUsuario == true) //|| visualizaSoloLiquidaciones == false
                    oStatic.Item.Visible = true;
                else
                    oStatic.Item.Visible = false;

                oStatic = oFormDatosIngresoLiquidaciones.Items.Item("3").Specific; // Code proxima liquidacion
                oStatic.Value = obtenerCodeProxLiquidacion().ToString();

                //CargarMatrixIngresoLiquidaciones();
                oFormDatosIngresoLiquidaciones.Visible = true;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso("", "", "ERROR al Cargar Grilla", ex.Message.ToString()); // Guarda log del Proceso
            }
        }

        private void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus;
            SAPbouiCOM.MenuItem oMenuItem;
            oMenus = SBO_Application.Menus;

            if (!oMenus.Exists("LIQ_DOC"))
            {
                SAPbouiCOM.MenuCreationParams oCreationPackage;
                oCreationPackage = SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams);
                oMenuItem = SBO_Application.Menus.Item("43520");

                sPath = System.Windows.Forms.Application.StartupPath;
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
                oCreationPackage.UniqueID = "LIQ_DOC";
                oCreationPackage.String = "Liquidaciones";
                oCreationPackage.Position = oMenuItem.SubMenus.Count + 1;

                oMenus = oMenuItem.SubMenus;

                try
                {

                    oMenus.AddEx(oCreationPackage);
                    oMenuItem = SBO_Application.Menus.Item("LIQ_DOC");
                    if (visualizaSoloLiquidaciones == false)
                    {
                        oMenus = oMenuItem.SubMenus;
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "Liquidaciones";
                        oCreationPackage.String = "Visor de Liquidaciones";
                        oMenus.AddEx(oCreationPackage);
                    }

                    oMenus = oMenuItem.SubMenus;
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "Nuevas Liquidaciones";
                    oCreationPackage.String = "Ingreso de Liquidaciones";
                    oMenus.AddEx(oCreationPackage);

                    if (visualizaSoloLiquidaciones == false)
                    {
                        oMenuItem = SBO_Application.Menus.Item("LIQ_DOC");
                        oMenus = oMenuItem.SubMenus;
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "Pedidos";
                        oCreationPackage.String = "Visor de Pedidos";
                        oMenus.AddEx(oCreationPackage);

                        oMenuItem = SBO_Application.Menus.Item("LIQ_DOC");
                        oMenus = oMenuItem.SubMenus;
                        oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                        oCreationPackage.UniqueID = "Cobros";
                        oCreationPackage.String = "Ingreso de Comprobantes";
                        oMenus.AddEx(oCreationPackage);
                    }

                    oMenuItem = SBO_Application.Menus.Item("LIQ_DOC");
                    oMenus = oMenuItem.SubMenus;
                    oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                    oCreationPackage.UniqueID = "Cheques";
                    oCreationPackage.String = "Visor de Cheques";
                    oMenus.AddEx(oCreationPackage);
                }
                catch (Exception er)
                {
                    String msg = "";
                    if (er.Message.Equals("Menu - Already exists"))
                    {
                        msg = "Menú ya fue creado.";
                    }
                    else
                    {
                        msg = er.Message;
                    }
                }
            }
        }

        private void CargarGrilla()
        {
            try
            {
                SAPbouiCOM.EditText oStatic;
                int num = 0;
                /*oStatic = oFormDatosPedido.Items.Item("6").Specific;

                
                if (!String.IsNullOrEmpty(oStatic.String))
                    num = Convert.ToInt32(oStatic.String);*/

                string liquidacion = ""; string estado = ""; string territorio = ""; string cardCode = "";
                clsCanales clsCan = new clsCanales();
                DateTime fechaDesde = Convert.ToDateTime("01-01-1980"); //"01-01-1980"
                DateTime fechaHasta = Convert.ToDateTime("01-01-1980");
                string ordenEntrega = "";

                try
                {
                    oStatic = oFormDatosPedido.Items.Item("16").Specific; // Desde Fecha
                    /*if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.Value = fechaDesde.ToString("dd/MM/yyyy");*/

                    if (!String.IsNullOrEmpty(oStatic.String))
                        fechaDesde = Convert.ToDateTime(oStatic.String);

                    oStatic = oFormDatosPedido.Items.Item("10").Specific; // Hasta Fecha
                    /*if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.Value = fechaHasta.ToString("dd/MM/yyyy");*/

                    if (!String.IsNullOrEmpty(oStatic.String))
                        fechaHasta = Convert.ToDateTime(oStatic.String);

                    try
                    {
                        oStatic = oFormDatosPedido.Items.Item("61").Specific; // Canal
                        clsCan.Uno = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("62").Specific; // Canal
                        clsCan.Dos = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("63").Specific; // Canal
                        clsCan.Tres = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("64").Specific; // Canal
                        clsCan.Cuatro = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("65").Specific; // Canal
                        clsCan.Cinco = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("66").Specific; // Canal
                        clsCan.Seis = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("67").Specific; // Canal
                        clsCan.Siete = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("68").Specific; // Canal
                        clsCan.Ocho = oStatic.String;
                        oStatic = oFormDatosPedido.Items.Item("69").Specific; // Canal
                        clsCan.Nueve = oStatic.String;
                    }
                    catch (Exception ex)
                    { }

                    oStatic = oFormDatosPedido.Items.Item("110").Specific; // CardCode
                    cardCode = oStatic.String;

                    SAPbouiCOM.ComboBox oStaticCombo;

                    oStaticCombo = oFormDatosPedido.Items.Item("29").Specific; // Número Liquidación
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                        liquidacion = oStaticCombo.Selected.Value.ToString();

                    oStaticCombo = oFormDatosPedido.Items.Item("28").Specific; // Estado Liquidación
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                        estado = oStaticCombo.Selected.Value.ToString();

                    oStaticCombo = oFormDatosPedido.Items.Item("35").Specific; //AP. 2019-07-22, nuevo, Orden de entrega Storas.
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                        ordenEntrega = oStaticCombo.Selected.Value.ToString();

                    try
                    {
                        SAPbouiCOM.Matrix oStaticMatriz;
                        oStaticMatriz = oFormDatosPedido.Items.Item("31").Specific;

                        int cantRows = oStaticMatriz.RowCount; // Saco la cantidad de registros que tiene la Grilla
                        int row = 0;

                        for (int j = 0; j < cantRows; j++) // Mientras tenga registros 
                        {
                            row = j + 1;

                            SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStaticMatriz.GetCellSpecific("V_2", row);
                            clsTerritorio terrSeleccionado = new clsTerritorio();
                            terrSeleccionado.IdTerritorio = Int32.Parse(ed.Value);
                            //ed = (SAPbouiCOM.EditText)oStaticMatriz.GetCellSpecific("V_1", row);
                            //terrSeleccionado.NombreTerritorio = ed.Value.ToString();
                            //ed = (SAPbouiCOM.EditText)oStaticMatriz.GetCellSpecific("V_0", row);
                            CheckBox oChk = (CheckBox)oStaticMatriz.GetCellSpecific("V_4", row); // ACAAA CAMBIO V_0
                            bool esSeleccionado = oChk.Checked;
                            if (esSeleccionado == true) // Si es un territorio seleccionado
                            {
                                if (!String.IsNullOrEmpty(territorio.ToString()))
                                    territorio += " or oin.U_TERRITORIO = '" + terrSeleccionado.IdTerritorio.ToString() + "'";
                                else
                                    territorio += " oin.U_TERRITORIO = '" + terrSeleccionado.IdTerritorio.ToString() + "'";

                                //territorio = terrSeleccionado.IdTerritorio.ToString();
                            }
                        }

                        //if (!String.IsNullOrEmpty(territorio.ToString())) // Si tengo filtro de territorio le agrego los paréntesis
                            //territorio = " (" + territorio.ToString() + ") ";

                        
                    }
                    catch (Exception ex)
                    { }
                }
                catch (Exception ex)
                { }

                CargarMatrix(num, fechaDesde, fechaHasta, liquidacion, clsCan, estado, territorio, cardCode, ordenEntrega);
            }
            catch (Exception ex)
            { }
        }

        private void CargarGrillaPedidosVisor()
        {
            try
            {
                SAPbouiCOM.EditText oStatic;
                int num = 0;
                /*oStatic = oFormDatosPedido.Items.Item("6").Specific;

                
                if (!String.IsNullOrEmpty(oStatic.String))
                    num = Convert.ToInt32(oStatic.String);*/

                string estado = ""; string territorio = ""; string liquidacion = ""; string estadoPedido = ""; string cardCode = ""; string freeShop = ""; string itemCode = "";
                clsCanales clsCan = new clsCanales();
                DateTime fechaDesde = Convert.ToDateTime(DateTime.Now); //"01-01-1980"
                DateTime fechaHasta = Convert.ToDateTime(DateTime.Now);
                try
                {
                    oStatic = oFormDatosPedidoVisor.Items.Item("16").Specific; // Desde Fecha
                    /*if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.Value = fechaDesde.ToString("dd/MM/yyyy");*/

                    if (!String.IsNullOrEmpty(oStatic.String))
                        fechaDesde = Convert.ToDateTime(oStatic.String);

                    oStatic = oFormDatosPedidoVisor.Items.Item("10").Specific; // Hasta Fecha
                    /*if (String.IsNullOrEmpty(oStatic.String))
                        oStatic.Value = fechaHasta.ToString("dd/MM/yyyy");*/

                    if (!String.IsNullOrEmpty(oStatic.String))
                        fechaHasta = Convert.ToDateTime(oStatic.String);

                    try
                    {
                        oStatic = oFormDatosPedidoVisor.Items.Item("61").Specific; // Canal
                        clsCan.Uno = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("62").Specific; // Canal
                        clsCan.Dos = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("63").Specific; // Canal
                        clsCan.Tres = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("64").Specific; // Canal
                        clsCan.Cuatro = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("65").Specific; // Canal
                        clsCan.Cinco = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("66").Specific; // Canal
                        clsCan.Seis = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("67").Specific; // Canal
                        clsCan.Siete = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("68").Specific; // Canal
                        clsCan.Ocho = oStatic.String;
                        oStatic = oFormDatosPedidoVisor.Items.Item("69").Specific; // Canal
                        clsCan.Nueve = oStatic.String;
                    }
                    catch (Exception ex)
                    { }

                    oStatic = oFormDatosPedidoVisor.Items.Item("110").Specific; // CardCode
                    cardCode = oStatic.String;

                    oStatic = oFormDatosPedidoVisor.Items.Item("32").Specific; // ItemCode
                    itemCode = oStatic.String;

                    SAPbouiCOM.ComboBox oStaticCombo;

                    oStaticCombo = oFormDatosPedidoVisor.Items.Item("28").Specific; // Estado Pedido
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                        estadoPedido = oStaticCombo.Selected.Value.ToString();

                    oStaticCombo = oFormDatosPedidoVisor.Items.Item("1000002").Specific; // Pedido FreeShop
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                    {
                        freeShop = oStaticCombo.Selected.Value.ToString();
                        if (freeShop.ToString().Equals("Y"))
                            freeShop = ">"; // El signo > es para que en el where traiga todos los documentos que tienen lineas con al menos un articulo de Whisky
                        else if (freeShop.ToString().Equals("N"))
                            freeShop = "="; // El signo = es para que en el where traiga todos los documentos que NO tienen lineas con al menos un articulo de Whisky 
                        else
                            freeShop = ""; // Lo dejo vacio para que no haga el filtro
                    }
                    try
                    {
                        SAPbouiCOM.Matrix oStaticMatriz;
                        oStaticMatriz = oFormDatosPedidoVisor.Items.Item("31").Specific;

                        int cantRows = oStaticMatriz.RowCount; // Saco la cantidad de registros que tiene la Grilla
                        int row = 0;

                        for (int j = 0; j < cantRows; j++) // Mientras tenga registros 
                        {
                            row = j + 1;

                            SAPbouiCOM.EditText ed = (SAPbouiCOM.EditText)oStaticMatriz.GetCellSpecific("V_3", row);
                            clsTerritorio terrSeleccionado = new clsTerritorio();
                            terrSeleccionado.IdTerritorio = Int32.Parse(ed.Value);
                            //ed = (SAPbouiCOM.EditText)oStaticMatriz.GetCellSpecific("V_1", row);
                            //terrSeleccionado.NombreTerritorio = ed.Value.ToString();
                            //ed = (SAPbouiCOM.EditText)oStaticMatriz.GetCellSpecific("V_0", row);
                            CheckBox oChk = (CheckBox)oStaticMatriz.GetCellSpecific("V_2", row); // ACAAA CAMBIO v_0
                            bool esSeleccionado = oChk.Checked;
                            if (esSeleccionado == true) // Si es un territorio seleccionado
                            {
                                if (!String.IsNullOrEmpty(territorio.ToString()))
                                    territorio += " or U_TERRITORIO = '" + terrSeleccionado.IdTerritorio.ToString() + "'";
                                else
                                    territorio += " U_TERRITORIO = '" + terrSeleccionado.IdTerritorio.ToString() + "'";
                            }
                        }

                        if (!String.IsNullOrEmpty(territorio.ToString())) // Si tengo filtro de territorio le agrego los paréntesis
                            territorio = " (" + territorio.ToString() + ") ";
                    }
                    catch (Exception ex)
                    { }
                }
                catch (Exception ex)
                { }

                CargarMatrixPedidosVisor(num, fechaDesde, fechaHasta, liquidacion, clsCan, estado, territorio, estadoPedido, cardCode, freeShop, itemCode);
            }
            catch (Exception ex)
            { }
        }

        private void CargarGrillaChequesVisor(bool pPrimeraVez)
        {
            try
            {
                string cuentaCheque = "";

                SAPbouiCOM.ComboBox oStaticCombo;

                oStaticCombo = oFormDatosCheques.Items.Item("28").Specific;
                if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                {
                    cuentaCheque = oStaticCombo.Selected.Value.ToString();

                    CargarMatrixChequesVisor(cuentaCheque);
                }
                else
                {
                    if (pPrimeraVez == false) // Si no es la primera vez que lo abro
                        SBO_Application.MessageBox("Debe seleccionar una cuenta de cheques en custodia");
                }
            }
            catch (Exception ex)
            { }
        }

        private void CargarMatrix(int codigo, DateTime pFechaDesde, DateTime pFechaHasta, string pLiquidacion, clsCanales pCanal, string pEstadoLiquidacion, string pTerritorio, string pCardCode, string pOrdenEntrega)
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosPedido != null)
                {
                    matriz = oFormDatosPedido.Items.Item("2").Specific;
                }
                else
                {
                    oFormDatosPedido = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosPedido.Items.Item("2").Specific;
                }

                List<clsDocumento> docs = obtenerDocumentos(codigo, pFechaDesde, pFechaHasta, pLiquidacion, pCanal, pEstadoLiquidacion, pTerritorio, pCardCode, pOrdenEntrega);
                oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").Rows.Clear();
                oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").Rows.Add(docs.Count);
                int cont = 0;
                foreach (clsDocumento doc in docs)
                {
                    //oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColLink", cont, doc.DocEntry);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColDocNum", cont, doc.DocNum);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColFecha", cont, doc.Fecha);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColCliente", cont, doc.Cliente);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColDocEntry", cont, doc.DocEntry);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColMonto", cont, doc.Monto);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColLiquidacion", cont, doc.NroLiquidacion);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColCanal", cont, doc.Canal);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColTipo", cont, doc.Tipo);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColVendedor", cont, doc.Vendedor);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosDoc").SetValue("ColFolio", cont, doc.Numero);
                    cont++;
                }

                matriz.Columns.Item("V_0").DataBind.Bind("DatosDoc", "ColLiquidacion");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosDoc", "ColCanal");
                matriz.Columns.Item("V_2").DataBind.Bind("DatosDoc", "ColCliente");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosDoc", "ColMonto");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosDoc", "ColFecha");
                matriz.Columns.Item("V_6").DataBind.Bind("DatosDoc", "ColDocEntry");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosDoc", "ColDocNum");
                //matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColLink");
                matriz.Columns.Item("V_9").DataBind.Bind("DatosDoc", "ColTipo");
                matriz.Columns.Item("V_5").DataBind.Bind("DatosDoc", "ColVendedor");
                matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColFolio");
                //SAPbouiCOM.LinkedButton oLink;
                //oLink = matriz.Columns.Item("V_8").ExtendedObject;
                //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Invoice;
                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();

                SAPbouiCOM.StaticText oText;
                oText = oFormDatosPedido.Items.Item("1000005").Specific; // Cantidad de documentos
                oText.Caption = cont.ToString() + " documentos";
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
        }

        private void CargarMatrixTerritoriosVisorLiquidaciones()
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosPedido != null)
                {
                    matriz = oFormDatosPedido.Items.Item("31").Specific;
                }
                else
                {
                    oFormDatosPedido = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosPedido.Items.Item("31").Specific;
                }

                List<clsTerritorio> docs = obtenerTerritorios();
                oFormDatosPedido.DataSources.DataTables.Item("DatosTerr").Rows.Clear();
                oFormDatosPedido.DataSources.DataTables.Item("DatosTerr").Rows.Add(docs.Count);
                int cont = 0;
                foreach (clsTerritorio terr in docs)
                {
                    oFormDatosPedido.DataSources.DataTables.Item("DatosTerr").SetValue("ColCode", cont, terr.IdTerritorio);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosTerr").SetValue("ColName", cont, terr.NombreTerritorio);
                    oFormDatosPedido.DataSources.DataTables.Item("DatosTerr").SetValue("ColFiltro", cont, terr.Seleccionado.ToString());

                    cont++;
                }

                matriz.Columns.Item("V_4").DataBind.Bind("DatosTerr", "ColFiltro"); // ACAAA CAMBIO v_0
                matriz.Columns.Item("V_1").DataBind.Bind("DatosTerr", "ColName");
                matriz.Columns.Item("V_2").DataBind.Bind("DatosTerr", "ColCode");

                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
        }

        private void CargarMatrixTerritoriosVisorPedidos()
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosPedidoVisor != null)
                {
                    matriz = oFormDatosPedidoVisor.Items.Item("31").Specific;
                }
                else
                {
                    oFormDatosPedidoVisor = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosPedidoVisor.Items.Item("31").Specific;
                }

                matriz.Item.Width = 260; // Le dejo fijo el ancho a la grilla

                List<clsTerritorio> docs = obtenerTerritorios();
                oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosTerr").Rows.Clear();
                oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosTerr").Rows.Add(docs.Count);
                int cont = 0;
                foreach (clsTerritorio terr in docs)
                {
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosTerr").SetValue("ColFiltro", cont, terr.Seleccionado.ToString());
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosTerr").SetValue("ColCode", cont, terr.IdTerritorio);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosTerr").SetValue("ColName", cont, terr.NombreTerritorio);

                    cont++;
                }

                matriz.Columns.Item("V_2").DataBind.Bind("DatosTerr", "ColFiltro"); // ACAAA CAMBIO v_0
                matriz.Columns.Item("V_1").DataBind.Bind("DatosTerr", "ColName");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosTerr", "ColCode");

                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
        }

        private void CargarMatrixPedidosVisor(int codigo, DateTime pFechaDesde, DateTime pFechaHasta, string pLiquidacion, clsCanales pCanal, string pEstadoLiquidacion, string pTerritorio, string pEstadoPedido, string pCardCode, string pPedidoFreeShop, string pItemCode)
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosPedidoVisor != null)
                {
                    matriz = oFormDatosPedidoVisor.Items.Item("2").Specific;
                }
                else
                {
                    oFormDatosPedidoVisor = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosPedidoVisor.Items.Item("2").Specific;
                }

                List<clsDocumento> docs = obtenerPedidos(codigo, pFechaDesde, pFechaHasta, pLiquidacion, pCanal, pEstadoLiquidacion, pTerritorio, pEstadoPedido, pCardCode, pPedidoFreeShop, pItemCode);
                oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").Rows.Clear();
                oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").Rows.Add(docs.Count);
                int cont = 0;
                foreach (clsDocumento doc in docs)
                {
                    //oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColLink", cont, doc.DocEntry);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColDocNum", cont, doc.DocNum);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColFecha", cont, doc.Fecha);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCliente", cont, doc.Cliente);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColDocEntry", cont, doc.DocEntry);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColMonto", cont, doc.Monto);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColLiquidacion", cont, doc.NroLiquidacion);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColCanal", cont, doc.Canal);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColTipo", cont, doc.Tipo);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColConfirmed", cont, doc.Confirmado);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColComentarios", cont, doc.Comentarios);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColFreeShop", cont, doc.FreeShop);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColVendedor", cont, doc.Vendedor);
                    oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColPick", cont, doc.Pickeado);
                    cont++;
                }

                matriz.Columns.Item("V_0").DataBind.Bind("DatosDoc", "ColLiquidacion");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosDoc", "ColCanal");
                matriz.Columns.Item("V_2").DataBind.Bind("DatosDoc", "ColCliente");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosDoc", "ColMonto");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosDoc", "ColFecha");
                matriz.Columns.Item("V_6").DataBind.Bind("DatosDoc", "ColDocEntry");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosDoc", "ColDocNum");
                //matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColLink");
                matriz.Columns.Item("V_9").DataBind.Bind("DatosDoc", "ColTipo");
                matriz.Columns.Item("V_12").DataBind.Bind("DatosDoc", "ColConfirmed");
                matriz.Columns.Item("V_5").DataBind.Bind("DatosDoc", "ColComentarios");
                matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColFreeShop");
                matriz.Columns.Item("V_10").DataBind.Bind("DatosDoc", "ColVendedor");
                matriz.Columns.Item("V_11").DataBind.Bind("DatosDoc", "ColPick");

                //SAPbouiCOM.LinkedButton oLink;
                //oLink = matriz.Columns.Item("V_8").ExtendedObject;
                //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_SalesOpportunity;
                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();

                SAPbouiCOM.StaticText oText;
                oText = oFormDatosPedidoVisor.Items.Item("1000005").Specific; // Cantidad de documentos
                oText.Caption = cont.ToString() + " documentos";
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
        }

        private void CargarMatrixChequesVisor(string pCuenta)
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosCheques != null)
                {
                    matriz = oFormDatosCheques.Items.Item("2").Specific;
                }
                else
                {
                    oFormDatosCheques = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosCheques.Items.Item("2").Specific;
                }

                List<clsCheque> docs = obtenerCheques(pCuenta);
                oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").Rows.Clear();
                oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").Rows.Add(docs.Count);
                int cont = 0;
                foreach (clsCheque doc in docs)
                {
                    //oFormDatosPedidoVisor.DataSources.DataTables.Item("DatosDoc").SetValue("ColLink", cont, doc.DocEntry);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColSecuencia", cont, doc.NumSecuencia);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColNumCheque", cont, doc.NumCheque);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColMoneda", cont, doc.Moneda);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColFecha", cont, doc.Fecha);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColMonto", cont, doc.Monto);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColBanco", cont, doc.Banco);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColSucursal", cont, doc.Sucursal);
                    oFormDatosCheques.DataSources.DataTables.Item("DatosDoc").SetValue("ColAcreditado", cont, doc.Acreditado);

                    cont++;
                }

                matriz.Columns.Item("V_0").DataBind.Bind("DatosDoc", "ColAcreditado");
                matriz.Columns.Item("V_2").DataBind.Bind("DatosDoc", "ColMoneda");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosDoc", "ColMonto");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosDoc", "ColFecha");
                matriz.Columns.Item("V_6").DataBind.Bind("DatosDoc", "ColSucursal");
                matriz.Columns.Item("V_7").DataBind.Bind("DatosDoc", "ColBanco");
                //matriz.Columns.Item("V_8").DataBind.Bind("DatosDoc", "ColLink");
                matriz.Columns.Item("V_9").DataBind.Bind("DatosDoc", "ColSecuencia");
                matriz.Columns.Item("V_12").DataBind.Bind("DatosDoc", "ColNumCheque");

                //SAPbouiCOM.LinkedButton oLink;
                //oLink = matriz.Columns.Item("V_8").ExtendedObject;
                //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_SalesOpportunity;
                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();

                SAPbouiCOM.StaticText oText;
                oText = oFormDatosCheques.Items.Item("1000005").Specific; // Cantidad de documentos
                oText.Caption = cont.ToString() + " cheques";
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
        }

        private void CargarMatrixIngresoLiquidaciones()
        {
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosIngresoLiquidaciones != null)
                {
                    matriz = oFormDatosIngresoLiquidaciones.Items.Item("100").Specific;
                }
                else
                {
                    oFormDatosIngresoLiquidaciones = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosIngresoLiquidaciones.Items.Item("100").Specific;
                }

                SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                DateTime fecha = Convert.ToDateTime("01-01-1980");
                string repartidor = "";
                string estadoLiquidacion = "";
                try
                {
                    SAPbouiCOM.EditText oStatic;
                    oStatic = oFormDatosIngresoLiquidaciones.Items.Item("9").Specific; // Fecha Filtro

                    if (!String.IsNullOrEmpty(oStatic.String))
                        fecha = Convert.ToDateTime(oStatic.String);

                    SAPbouiCOM.ComboBox oStaticCombo;

                    oStaticCombo = oFormDatosIngresoLiquidaciones.Items.Item("14").Specific; // Repartidor Liquidación
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                        repartidor = oStaticCombo.Selected.Value.ToString();

                    oStaticCombo = oFormDatosIngresoLiquidaciones.Items.Item("13").Specific; // Estado Liquidación
                    if (!String.IsNullOrEmpty(oStaticCombo.Value.ToString()))
                        estadoLiquidacion = oStaticCombo.Selected.Value.ToString();
                }
                catch (Exception ex)
                { }

                string where = "";
                string query = "select liq.Code,U_FECHA as 'Fecha Despacho',rep.Name as Repartidor,liq.Name," +
                "case when U_ESTADO = 1 then 'Abierta' else 'Cerrada' end as Estado from [@LIQUIDACIONES] as liq " +
                "left join [@REPARTIDORES] as rep on rep.Code = liq.U_REPARTIDOR ";

                if (!String.IsNullOrEmpty(fecha.ToString()) && fecha != Convert.ToDateTime("01-01-1980")) // Si tiene una fecha
                {
                    if (!String.IsNullOrEmpty(where.ToString()))
                        where += " and liq.U_FECHA = '" + fecha.ToString("yyyy-MM-dd") + "' ";
                    else
                        where += " where liq.U_FECHA = '" + fecha.ToString("yyyy-MM-dd") + "' ";
                }

                if (!String.IsNullOrEmpty(repartidor.ToString())) // Si tiene un Repartidor por el cual filtrar
                {
                    if (!String.IsNullOrEmpty(where.ToString()))
                        where += " and liq.U_REPARTIDOR = '" + repartidor + "' ";
                    else
                        where += " where liq.U_REPARTIDOR = '" + repartidor + "' ";
                }

                if (!String.IsNullOrEmpty(estadoLiquidacion.ToString()) && (estadoLiquidacion.ToString().Equals("1") || estadoLiquidacion.ToString().Equals("2"))) // Si tiene un Estado de liquidacion para filtrar
                {
                    if (!String.IsNullOrEmpty(where.ToString()))
                        where += " and liq.U_ESTADO = '" + estadoLiquidacion + "' ";
                    else
                        where += " where liq.U_ESTADO = '" + estadoLiquidacion + "' ";
                }

                query += where + " order by CAST(liq.Code AS Int)";

                oRSMyTable.DoQuery(query);

                oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").Rows.Clear();
                oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").Rows.Add(oRSMyTable.RecordCount);
                int cont = 0;

                while (!oRSMyTable.EoF)
                {
                    oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").SetValue("Code", cont, Convert.ToInt32(oRSMyTable.Fields.Item("Code").Value));
                    oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").SetValue("Fecha Despacho", cont, Convert.ToDateTime(oRSMyTable.Fields.Item("Fecha Despacho").Value));
                    oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").SetValue("Repartidor", cont, oRSMyTable.Fields.Item("Repartidor").Value);
                    oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").SetValue("Name", cont, oRSMyTable.Fields.Item("Name").Value);
                    oFormDatosIngresoLiquidaciones.DataSources.DataTables.Item("DatosLiq").SetValue("Estado", cont, oRSMyTable.Fields.Item("Estado").Value);
                    cont++;
                    oRSMyTable.MoveNext();
                }

                matriz.Columns.Item("V_0").DataBind.Bind("DatosLiq", "Estado");
                matriz.Columns.Item("V_2").DataBind.Bind("DatosLiq", "Repartidor");
                matriz.Columns.Item("V_3").DataBind.Bind("DatosLiq", "Fecha Despacho");
                matriz.Columns.Item("V_4").DataBind.Bind("DatosLiq", "Code");
                matriz.Columns.Item("V_1").DataBind.Bind("DatosLiq", "Name");

                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
        }

        public int CargarMatrixIngresoCobros(string pNumeroLiquidacion, string pMonedaCobro, double pImporteCobro)
        {
            int cantRegistros = 0;
            SAPbouiCOM.Matrix matriz = null;
            try
            {
                if (oFormDatosIngresoCobros != null)
                    matriz = oFormDatosIngresoCobros.Items.Item("15").Specific;
                else
                {
                    oFormDatosIngresoCobros = SBO_Application.Forms.Item("OpenProject");
                    matriz = oFormDatosIngresoCobros.Items.Item("15").Specific;
                }

                SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                /*string query = "select oin.DocEntry,oin.DocNum,PaidToDate,oin.CardCode,oin.CardName,oin.DocTotal,oin.DocTotalFC,oin.DocCur,oin.VatSum,oin.VatSumFC from OINV as oin " +
                "inner join OCRD as ocr on ocr.CardCode = oin.CardCode and ocr.CardType = 'C' " +
                "inner join OCTG as oct on oct.GroupNum = oin.GroupNum " +
                "where oin.U_LIQUIDACION = '" + pNumeroLiquidacion.ToString() + "' and oin.Canceled = 'N' and oct.PymntGroup like 'CONTADO%' and oin.DocCur = '" + pMonedaCobro + "' and oin.FolioNum is not null "; // Solo documentos que tengan como cond de pago Contado
                */
                string codigoISOMonedaCobro = obtenerCodigoISOMoneda(pMonedaCobro); // Obtengo el codigo ISO de la moneda de cobro

                //query += " and (PaidToDate <> DocTotal and (PaidToDate <> DocTotalFC or PaidToDate = 0)) order by oin.DocEntry"; // Acá verifica que el documento tenga saldo pendiente aún

                // QUERY DE JUAN, PARA OBTENER FACTURAS Y NOTAS DE CREDITOS
                string query = "Declare @Liquidacion char(10) Declare @Moneda char(5) " +
                "Set @Liquidacion='" + pNumeroLiquidacion + "' set @Moneda='" + pMonedaCobro + "' " +
                "select oin.DocEntry,oin.DocNum, case when oin.FolioNum is null then 'Doc.Sin Folio' else ('Factura ')+rtrim(left(oct.PymntGroup,7))+' Folio: '+rtrim(oin.FolioPref)+rtrim(cast(oin.FolioNum as CHAR(10))) end Documento, " +
                "oin.CardCode, oin.CardName, oin.DocCur, /*case when @Moneda='$' then oin.DocTotal else oin.DocTotalFC end Total*/ isnull(DocTotal,0) as DocTotal, isnull(DocTotalFC,0) as DocTotalFC, case when @Moneda='$' then oin.PaidToDate else oin.PaidFC end Cancelado, case when oin.DocCur='$' then SUM(oin.DocTotal-oin.PaidToDate) else SUM(oin.DocTotalFC-oin.PaidFC) end SaldoMonedaOriginal	 " +
                "from OINV as oin  " +
                "inner join OCRD as ocr on ocr.CardCode = oin.CardCode and ocr.CardType = 'C'  " +
                "inner join OCTG as oct on oct.GroupNum = oin.GroupNum  " +
                "where	oin.U_LIQUIDACION = @Liquidacion and oin.Canceled = 'N' and oct.PymntGroup like 'CONTADO%' and oin.DocCur = @Moneda and oin.FolioNum is not null  " +
                "group by oin.DocEntry, oin.DocNum, oct.PymntGroup, oin.FolioPref, oin.FolioNum, oin.CardCode, oin.CardName, oin.DocCur, oin.DocTotal, oin.DocTotalFC, oin.PaidToDate, oin.PaidFC " +
                "having (case	when oin.DocCur='$' then SUM(oin.DocTotal-oin.PaidToDate) " +
                "else SUM(oin.DocTotalFC-oin.PaidFC)end)>0 " +
                "union all " +
                "select oin.DocEntry,oin.DocNum,case	when oin.FolioNum is null then 'Doc.Sin Folio' else ('N/Devolucion ')+rtrim(left(oct.PymntGroup,7))+' Folio: '+rtrim(oin.FolioPref)+rtrim(cast(oin.FolioNum as CHAR(10))) end Documento, " +
                "oin.CardCode,oin.CardName,oin.DocCur,/*case when @Moneda='$' then oin.DocTotal else oin.DocTotalFC end Total*/ isnull(DocTotal,0) as DocTotal, isnull(DocTotalFC,0) as DocTotalFC,case when @Moneda='$' then oin.PaidToDate else oin.PaidFC end Cancelado, case when oin.DocCur='$' then SUM(oin.DocTotal-oin.PaidToDate) else SUM(oin.DocTotalFC-oin.PaidFC) end SaldoMonedaOriginal	 " +
                "from Orin as oin  " +
                "inner join OCRD as ocr on ocr.CardCode = oin.CardCode and ocr.CardType = 'C'  " +
                "inner join OCTG as oct on oct.GroupNum = oin.GroupNum  " +
                "where	oin.U_LIQUIDACION = @Liquidacion and oin.Canceled = 'N' and oct.PymntGroup like 'CONTADO%' and oin.DocCur = @Moneda and oin.FolioNum is not null  " +
                "group by oin.DocEntry, oin.DocNum, oct.PymntGroup, oin.FolioPref, oin.FolioNum, oin.CardCode, oin.CardName, oin.DocCur, oin.DocTotal, oin.DocTotalFC, oin.PaidToDate, oin.PaidFC " +
                "having (case	when oin.DocCur='$' then SUM(oin.DocTotal-oin.PaidToDate) " +
                "else SUM(oin.DocTotalFC-oin.PaidFC)end)>0 " +
                "order by 4, 3";

                //if (codigoISOMonedaCobro.ToString().Equals("UYU") || codigoISOMonedaCobro.ToString().Equals("CLP")) // Si el documento es en Pesos
                //    query += "and PaidToDate <> DocTotal "; // Acá verifica que el documento tenga saldo pendiente aún
                //else
                //    query += "and PaidToDate <> DocTotalFC "; // Acá verifica que el documento tenga saldo pendiente aún

                oRSMyTable.DoQuery(query);

                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").Rows.Clear();
                //oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").Rows.Add(oRSMyTable.RecordCount);
                int cont = 0;

                double saldoImporteCobro = pImporteCobro; // Creo una variable saldo para ir restandole el total de cada documento

                if (oRSMyTable != null)
                {
                    if (oRSMyTable.RecordCount != 0)
                    {
                        cantRegistros = oRSMyTable.RecordCount; // Me quedo con la cantidad de Registros
                        while (!oRSMyTable.EoF)
                        {
                            try
                            {
                                //if (saldoImporteCobro > 0)
                                //{
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").Rows.Add(1);

                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("DocEntry", cont, Convert.ToInt32(oRSMyTable.Fields.Item("DocEntry").Value));
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("DocNum", cont, Convert.ToInt32(oRSMyTable.Fields.Item("DocNum").Value));
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("CardCode", cont, oRSMyTable.Fields.Item("CardCode").Value);
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("CardName", cont, oRSMyTable.Fields.Item("CardName").Value);
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("DocTotal", cont, oRSMyTable.Fields.Item("DocTotal").Value);
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("DocTotalFC", cont, oRSMyTable.Fields.Item("DocTotalFC").Value);
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("Cobrado", cont, oRSMyTable.Fields.Item("Cancelado").Value);
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("Moneda", cont, oRSMyTable.Fields.Item("DocCur").Value);
                                oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("Documento", cont, oRSMyTable.Fields.Item("Documento").Value);

                                if (saldoImporteCobro > 0) // Le asigno valor  mientras hay saldo para asignar
                                {
                                    double importeYaPagoDocumento = (double)oRSMyTable.Fields.Item("Cancelado").Value; // Importe que ya se pagó del documento
                                    double montoDocumento = (double)oRSMyTable.Fields.Item("DocTotal").Value; // Monto del documento
                                    double montoDocumentoFC = (double)oRSMyTable.Fields.Item("DocTotalFC").Value; // Monto del documento en Moneda extranjera
                                    int docEntryDocumento = oRSMyTable.Fields.Item("DocEntry").Value;
                                    int docNumDocumento = oRSMyTable.Fields.Item("DocNum").Value;
                                    string cardCodeDocumento = (string)oRSMyTable.Fields.Item("CardCode").Value;
                                    string cardNameDocumento = (string)oRSMyTable.Fields.Item("CardName").Value;
                                    string monedaDocumento = (string)oRSMyTable.Fields.Item("DocCur").Value;
                                    string codigoISOMoneda = obtenerCodigoISOMoneda(monedaDocumento); // Obtengo el codigo ISO de la moneda

                                    double montoCobrarDocumento = 0;
                                    if (importeYaPagoDocumento != montoDocumento && (importeYaPagoDocumento != montoDocumentoFC || importeYaPagoDocumento == 0))
                                    {
                                        if (codigoISOMoneda.ToString().Equals("UYU") || codigoISOMoneda.ToString().Equals("CLP")) // Si el documento es en Pesos
                                            montoCobrarDocumento = montoDocumento - importeYaPagoDocumento; // Obtengo el saldo que falta pagar del documento
                                        else
                                        {
                                            montoCobrarDocumento = montoDocumentoFC - importeYaPagoDocumento; // Si es en Moneda Extranjera. Obtengo el saldo que falta pagar del documento
                                            montoDocumento = montoDocumentoFC; // Me guardo en esta variable el monto original del documento
                                        }
                                    }

                                    if (montoCobrarDocumento > saldoImporteCobro) // Si lo que tengo para cobrar es mayor al saldo que tengo disponible
                                        montoCobrarDocumento = saldoImporteCobro;

                                    saldoImporteCobro = saldoImporteCobro - montoCobrarDocumento; // Le resto al Saldo el total del documento

                                    oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("MontoCobrar", cont, Convert.ToDouble(montoCobrarDocumento));
                                }
                                else
                                    oFormDatosIngresoCobros.DataSources.DataTables.Item("dtDoc").SetValue("MontoCobrar", cont, Convert.ToDouble(0));
                                //}
                                //else Descomentar para que liste todos los documentos por mas que el Monto no de para cubrirlos
                                //    break;
                            }
                            catch (Exception exx)
                            {
                                guardaLogProceso("", "", "ERROR al Listar documento Ingreso Cobro", exx.Message.ToString());
                            }

                            cont++;
                            oRSMyTable.MoveNext();
                        }
                    }
                }

                matriz.Columns.Item("V_11").DataBind.Bind("dtDoc", "Documento");
                matriz.Columns.Item("V_7").DataBind.Bind("dtDoc", "DocEntry");
                matriz.Columns.Item("V_6").DataBind.Bind("dtDoc", "DocNum");
                matriz.Columns.Item("V_5").DataBind.Bind("dtDoc", "CardCode");
                matriz.Columns.Item("V_4").DataBind.Bind("dtDoc", "CardName");
                matriz.Columns.Item("V_3").DataBind.Bind("dtDoc", "DocTotal");
                matriz.Columns.Item("V_2").DataBind.Bind("dtDoc", "DocTotalFC");
                matriz.Columns.Item("V_1").DataBind.Bind("dtDoc", "MontoCobrar");
                matriz.Columns.Item("V_8").DataBind.Bind("dtDoc", "Cobrado");
                matriz.Columns.Item("V_0").DataBind.Bind("dtDoc", "Moneda");

                matriz.LoadFromDataSource();
                matriz.AutoResizeColumns();

                return cantRegistros;
            }
            catch (Exception ex)
            {
                guardaLogProceso("", "", "ERROR al Cargar la grilla", ex.Message.ToString());// Guarda log del Proceso
            }
            return cantRegistros;
        }

        // Funcion que recibe un objeto ComboBox, y un String para un DataSet para llenar el objeto
        private void llenarCombo(SAPbouiCOM.ComboBox pCombo, String pQuery, bool pSinRegistro, bool pBorrarRegistros)
        {
            try
            {
                SAPbobsCOM.Recordset oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRSMyTable.DoQuery(pQuery);

                if (pBorrarRegistros == true)
                {
                    try
                    {
                        int cant = pCombo.ValidValues.Count;
                        for (int i = cant; i > 0; i--) // Elimino los datos que tenga para cargarlos nuevamente
                            pCombo.ValidValues.Remove(i - 1, SAPbouiCOM.BoSearchKey.psk_Index);
                    }
                    catch (Exception ex)
                    { }
                }

                pCombo.ValidValues.Add("", "No filtrar");
                if (pSinRegistro == true)
                    pCombo.ValidValues.Add("0", "Sin Liquidación");

                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        pCombo.ValidValues.Add(oRSMyTable.Fields.Item("Code").Value, oRSMyTable.Fields.Item("Name").Value);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
            }
            catch (Exception ex)
            { }
        }

        // Llena el objeto ComboBox con los estados de liquidacion
        private void llenarComboEstadosLiquidacion(SAPbouiCOM.ComboBox pCombo, bool pNoFiltrar)
        {
            try
            {
                if (pNoFiltrar == true)
                    pCombo.ValidValues.Add("", "No filtrar");

                pCombo.ValidValues.Add("1", "Abierto");
                pCombo.ValidValues.Add("2", "Cerrado");
            }
            catch (Exception ex)
            { }
        }

        // Llena el objeto ComboBox con los estados de confirmacion del pedido
        private void llenarComboEstadosConfirmacionPedido(SAPbouiCOM.ComboBox pCombo, bool pNoFiltrar)
        {
            try
            {
                if (pNoFiltrar == true)
                    pCombo.ValidValues.Add("", "No filtrar");

                pCombo.ValidValues.Add("Y", "Confirmado");
                pCombo.ValidValues.Add("N", "Sin Confirmar");
            }
            catch (Exception ex)
            { }
        }

        //Llena el objeto ComboBox con los estados de pickeado del pedido
        private void llenarComboEstadosPickeadoPedido(SAPbouiCOM.ComboBox pCombo, bool pNoFiltrar)
        {
            try
            {
                if (pNoFiltrar == true)
                    pCombo.ValidValues.Add("", "No filtrar");

                pCombo.ValidValues.Add("Y", "Pickeado");
                pCombo.ValidValues.Add("N", "Sin Pickear");
            }
            catch (Exception ex)
            { }
        }

        // Llena el objeto ComboBox con los pedidos de FreeShop
        private void llenarComboPedidosFreeShop(SAPbouiCOM.ComboBox pCombo)
        {
            try
            {
                pCombo.ValidValues.Add("", "No filtrar");
                pCombo.ValidValues.Add("Y", "Con Whisky");
                pCombo.ValidValues.Add("N", "Sin Whisky");
            }
            catch (Exception ex)
            { }
        }
        #endregion

        #region General

        public List<clsDocumento> obtenerDocumentos(int codigo, DateTime pFechaDesde, DateTime pFechaHasta, string pLiquidacion, clsCanales pCanal, string pEstadoLiquidacion, string pTerritorio, string pCardCode, string pOrdenEntrega)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                List<clsDocumento> docs = new List<clsDocumento>();
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                int cantidadClientes = 0; int cantidadClientesNC = 0;// Variable para mostrar la cantidad de clientes que estan presentes en los documentos filtrados
                string where = "";
                String query = "select oin.DocEntry,DocNum, DocDate, oin.CardName,case when DocCur = '$' or DocCur = 'UYU' then DocTotal else DocTotalFC end as DocTotal,U_LIQUIDACION,U_CANAL, " +
                "case when DocSubType = 'IB' then 'BO' else case when isIns = 'Y' then 'FR' else case when IsICT = 'Y' then 'FP' else case when DocSubType = 'DN' then 'ND' else 'FD' end end end end as Tipo, osl.SlpName as Vendedor, FolioNum from OINV as oin " +
                " inner join OCRD as ocr on ocr.CardCode = oin.CardCode " +
                " inner join OSLP as osl on osl.SlpCode = oin.SlpCode ";

                string queryCountClientes = "select count(oin.CardCode) from OINV as oin " + // Query para sacar la cantidad de clientes en los documentos filtrados
                " inner join OCRD as ocr on ocr.CardCode = oin.CardCode " +
                " inner join OSLP as osl on osl.SlpCode = oin.SlpCode ";

                if (!String.IsNullOrEmpty(pEstadoLiquidacion.ToString()) && (pEstadoLiquidacion.ToString().Equals("1") || pEstadoLiquidacion.ToString().Equals("2"))) // Si tiene un Estado de liquidacion para filtrar
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                        where += " inner join [@LIQUIDACIONES] as liq on liq.Code = oin.U_LIQUIDACION where liq.U_ESTADO = '" + pEstadoLiquidacion + "'";
                }

                if (!String.IsNullOrEmpty(pTerritorio.ToString()) && !pTerritorio.ToString().Equals("0")) // Si tiene un valor en el campo Territorio
                {
                    //where += " inner join OTER as ote on ote.territryID = ocr.Territory "; //Cambio Pecoy
                    if (where.ToString().Contains("where"))
                        where += " and (" + pTerritorio + ")";//" and oin.U_TERRITORIO = '" + pTerritorio + "' ";
                    else
                        where += " where (" + pTerritorio + ")";//" where oin.U_TERRITORIO = '" + pTerritorio + "' ";
                }

                if (where.ToString().Contains("where")) // Verifica que no sea un documento cancelado 
                    where += "  and oin.Canceled = 'N' ";
                else
                    where += "  where oin.Canceled = 'N' ";

                if (!String.IsNullOrEmpty(codigo.ToString()) && codigo != 0) // Si el codigo no es vacio filtra por ese campo
                {
                    if (where.ToString().Contains("where"))
                        where += " and DocNum = '" + codigo + "'";
                    else
                        where += " where DocNum = '" + codigo + "'";
                }

                if (!String.IsNullOrEmpty(pFechaDesde.ToString()) && pFechaDesde != Convert.ToDateTime("01-01-1980")) // Si las fechas no son vacias filtra por ese rango
                {
                    if (where.ToString().Contains("where"))
                        where += " and DocDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' ";
                    else
                        where += " where DocDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' ";
                }

                if (!String.IsNullOrEmpty(pFechaHasta.ToString()) && pFechaHasta != Convert.ToDateTime("01-01-1980")) // Si las fechas no son vacias filtra por ese rango
                {
                    if (where.ToString().Contains("where"))
                        where += " and DocDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "' ";
                    else
                        where += " where DocDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "' ";
                }

                //if (!String.IsNullOrEmpty(pFechaDesde.ToString()) && !String.IsNullOrEmpty(pFechaHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                //{
                //    if (query.ToString().Contains("where"))
                //        query += " and U_FECHA_ENTREGA >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and U_FECHA_ENTREGA <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                //    else
                //        query += " where U_FECHA_ENTREGA >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and U_FECHA_ENTREGA <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                //}

                if (!String.IsNullOrEmpty(pLiquidacion.ToString())) // Si tiene un valor en el campo Liquidacion
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                    {
                        if (where.ToString().Contains("where"))
                            where += " and U_LIQUIDACION = '" + pLiquidacion + "'";
                        else
                            where += " where U_LIQUIDACION = '" + pLiquidacion + "'";
                    }
                    else
                    {
                        if (where.ToString().Contains("where")) // Busca los documentos sin Liquidacion
                            where += " and (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                        else
                            where += " where (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                    }
                }

                if (!String.IsNullOrEmpty(pCardCode.ToString()) && !pCardCode.ToString().Equals("0")) // Si tiene un valor en el campo CardCode
                {
                    if (where.ToString().Contains("where"))
                        where += " and oin.CardCode = '" + pCardCode + "' ";
                    else
                        where += " where oin.CardCode = '" + pCardCode + "' ";
                }

                if (!String.IsNullOrEmpty(pOrdenEntrega.ToString()) && !pOrdenEntrega.ToString().Equals("0")) //AP. 2019-07-19 nuevo campo Orden de entrega Storas.
                {
                    if (where.ToString().Contains("where"))
                        where += " and oin.U_PZSourceId = '" + pOrdenEntrega + "' ";
                    else
                        where += " where oin.U_PZSourceId = '" + pOrdenEntrega + "' ";
                }

                string whereCanal = "";
                if (!String.IsNullOrEmpty(pCanal.Uno)) // Si tiene un valor en el campo Canal
                    whereCanal += " U_CANAL = '" + pCanal.Uno + "'";
                if (!String.IsNullOrEmpty(pCanal.Dos)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Dos + "'";
                if (!String.IsNullOrEmpty(pCanal.Tres)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Tres + "'";
                if (!String.IsNullOrEmpty(pCanal.Cuatro)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cuatro + "'";
                if (!String.IsNullOrEmpty(pCanal.Cinco)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cinco + "'";
                if (!String.IsNullOrEmpty(pCanal.Seis)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Seis + "'";
                if (!String.IsNullOrEmpty(pCanal.Siete)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Siete + "'";
                if (!String.IsNullOrEmpty(pCanal.Ocho)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Ocho + "'";
                if (!String.IsNullOrEmpty(pCanal.Nueve)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Nueve + "'";

                if (!String.IsNullOrEmpty(whereCanal.ToString()))
                    whereCanal = " and (" + whereCanal.ToString() + ") ";

                query += where + whereCanal + " order by DocNum,DocDate";
                queryCountClientes += where + whereCanal + " group by oin.CardCode order by oin.CardCode";

                oRSMyTable.DoQuery(query);
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        clsDocumento doc = new clsDocumento();
                        String cli = oRSMyTable.Fields.Item("CardName").Value;
                        int cod = oRSMyTable.Fields.Item("DocNum").Value;
                        int docEntry = oRSMyTable.Fields.Item("DocEntry").Value;
                        DateTime fecha = oRSMyTable.Fields.Item("DocDate").Value;
                        double monto = oRSMyTable.Fields.Item("DocTotal").Value;
                        string canal = oRSMyTable.Fields.Item("U_CANAL").Value;
                        string nroLiquidacion = oRSMyTable.Fields.Item("U_LIQUIDACION").Value;
                        string tipoDoc = oRSMyTable.Fields.Item("Tipo").Value;
                        string vendedor = oRSMyTable.Fields.Item("Vendedor").Value;
                        int folio = oRSMyTable.Fields.Item("FolioNum").Value;

                        doc.Cliente = cli;
                        doc.DocNum = cod;
                        doc.DocEntry = docEntry;
                        doc.Fecha = fecha.Date.ToShortDateString();
                        doc.Monto = monto;
                        doc.NroLiquidacion = nroLiquidacion;
                        doc.Canal = canal;
                        doc.Tipo = tipoDoc;
                        doc.Vendedor = vendedor;
                        doc.Numero = folio;
                        docs.Add(doc);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                // Sumo la cantidad de clientes
                oRSMyTable.DoQuery(queryCountClientes);
                if (oRSMyTable != null)
                    cantidadClientes += oRSMyTable.RecordCount;

                // Recorro las Notas de Créditos
                query = "select oin.DocEntry,DocNum, DocDate, oin.CardName,case when DocCur = '$' or DocCur = 'UYU' then DocTotal else DocTotalFC end as DocTotal,U_LIQUIDACION,U_CANAL,'NC' as Tipo, osl.SlpName as Vendedor, FolioNum from ORIN as oin " +
                " inner join OCRD as ocr on ocr.CardCode = oin.CardCode " +
                " inner join OSLP as osl on osl.SlpCode = oin.SlpCode ";
                //" where FolioNum is not null and FolioPref is not null and Ref2 is not null ";

                queryCountClientes = "select count(oin.CardCode) from ORIN as oin " + // Query para sacar la cantidad de clientes en los documentos filtrados
                " inner join OCRD as ocr on ocr.CardCode = oin.CardCode " +
                " inner join OSLP as osl on osl.SlpCode = oin.SlpCode ";

                where = "";

                if (!String.IsNullOrEmpty(pEstadoLiquidacion.ToString()) && (pEstadoLiquidacion.ToString().Equals("1") || pEstadoLiquidacion.ToString().Equals("2"))) // Si tiene un Estado de liquidacion para filtrar
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                        where += " inner join [@LIQUIDACIONES] as liq on liq.Code = oin.U_LIQUIDACION where liq.U_ESTADO = '" + pEstadoLiquidacion + "'";
                }

                if (!String.IsNullOrEmpty(pTerritorio.ToString()) && !pTerritorio.ToString().Equals("0")) // Si tiene un valor en el campo Territorio
                {
                    //where += " inner join OTER as ote on ote.territryID = ocr.Territory ";
                    if (where.ToString().Contains("where"))
                        where += " and (" + pTerritorio +")"; //" and oin.U_TERRITORIO = '" + pTerritorio + "' ";
                    else
                        where += " where (" + pTerritorio + ")";//" where oin.U_TERRITORIO = '" + pTerritorio + "' ";
                }

                if (where.ToString().Contains("where")) // Verifica que no sea un documento cancelado 
                    where += "  and oin.Canceled = 'N' ";
                else
                    where += "  where oin.Canceled = 'N' ";

                if (!String.IsNullOrEmpty(codigo.ToString()) && codigo != 0) // Si el codigo no es vacio filtra por ese campo
                {
                    if (where.ToString().Contains("where"))
                        where += " and DocNum = '" + codigo + "'";
                    else
                        where += " where DocNum = '" + codigo + "'";
                }

                if (!String.IsNullOrEmpty(pFechaDesde.ToString()) && pFechaDesde != Convert.ToDateTime("01-01-1980")) // Si las fechas no son vacias filtra por ese rango
                {
                    if (where.ToString().Contains("where"))
                        where += " and DocDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' ";
                    else
                        where += " where DocDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' ";
                }

                if (!String.IsNullOrEmpty(pFechaHasta.ToString()) && pFechaHasta != Convert.ToDateTime("01-01-1980")) // Si las fechas no son vacias filtra por ese rango
                {
                    if (where.ToString().Contains("where"))
                        where += " and DocDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "' ";
                    else
                        where += " where DocDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "' ";
                }

                //if (!String.IsNullOrEmpty(pFechaDesde.ToString()) && !String.IsNullOrEmpty(pFechaHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                //{
                //    if (query.ToString().Contains("where"))
                //        query += " and U_FECHA_ENTREGA >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and U_FECHA_ENTREGA <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                //    else
                //        query += " where U_FECHA_ENTREGA >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and U_FECHA_ENTREGA <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                //}

                if (!String.IsNullOrEmpty(pLiquidacion.ToString())) // Si tiene un valor en el campo Liquidacion
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                    {
                        if (where.ToString().Contains("where"))
                            where += " and U_LIQUIDACION = '" + pLiquidacion + "'";
                        else
                            where += " where U_LIQUIDACION = '" + pLiquidacion + "'";
                    }
                    else
                    {
                        if (where.ToString().Contains("where")) // Busca los documentos sin Liquidacion
                            where += " and (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                        else
                            where += " where (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                    }
                }

                if (!String.IsNullOrEmpty(pCardCode.ToString()) && !pCardCode.ToString().Equals("0")) // Si tiene un valor en el campo CardCode
                {
                    if (where.ToString().Contains("where"))
                        where += " and oin.CardCode = '" + pCardCode + "' ";
                    else
                        where += " where oin.CardCode = '" + pCardCode + "' ";
                }

                if (!String.IsNullOrEmpty(pOrdenEntrega.ToString()) && !pOrdenEntrega.ToString().Equals("0")) //AP. 2019-08-20 nuevo campo Orden de entrega Storas.
                {
                    if (where.ToString().Contains("where"))
                        where += " and oin.U_PZSourceId = '" + pOrdenEntrega + "' ";
                    else
                        where += " where oin.U_PZSourceId = '" + pOrdenEntrega + "' ";
                }

                whereCanal = "";
                if (!String.IsNullOrEmpty(pCanal.Uno)) // Si tiene un valor en el campo Canal
                    whereCanal += " U_CANAL = '" + pCanal.Uno + "'";
                if (!String.IsNullOrEmpty(pCanal.Dos)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Dos + "'";
                if (!String.IsNullOrEmpty(pCanal.Tres)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Tres + "'";
                if (!String.IsNullOrEmpty(pCanal.Cuatro)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cuatro + "'";
                if (!String.IsNullOrEmpty(pCanal.Cinco)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cinco + "'";
                if (!String.IsNullOrEmpty(pCanal.Seis)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Seis + "'";
                if (!String.IsNullOrEmpty(pCanal.Siete)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Siete + "'";
                if (!String.IsNullOrEmpty(pCanal.Ocho)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Ocho + "'";
                if (!String.IsNullOrEmpty(pCanal.Nueve)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Nueve + "'";

                if (!String.IsNullOrEmpty(whereCanal.ToString()))
                    whereCanal = " and (" + whereCanal.ToString() + ") ";

                query += where + whereCanal + " order by DocNum,DocDate";
                queryCountClientes += where + whereCanal + " group by oin.CardCode order by oin.CardCode";

                oRSMyTable.DoQuery(query);
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        clsDocumento doc = new clsDocumento();
                        String cli = oRSMyTable.Fields.Item("CardName").Value;
                        int cod = oRSMyTable.Fields.Item("DocNum").Value;
                        int docEntry = oRSMyTable.Fields.Item("DocEntry").Value;
                        DateTime fecha = oRSMyTable.Fields.Item("DocDate").Value;
                        double monto = oRSMyTable.Fields.Item("DocTotal").Value;
                        string canal = oRSMyTable.Fields.Item("U_CANAL").Value;
                        string nroLiquidacion = oRSMyTable.Fields.Item("U_LIQUIDACION").Value;
                        string tipoDoc = oRSMyTable.Fields.Item("Tipo").Value;
                        string vendedor = oRSMyTable.Fields.Item("Vendedor").Value;
                        int folio = oRSMyTable.Fields.Item("FolioNum").Value;

                        doc.Cliente = cli;
                        doc.DocNum = cod;
                        doc.DocEntry = docEntry;
                        doc.Fecha = fecha.Date.ToShortDateString();
                        doc.Monto = monto;
                        doc.NroLiquidacion = nroLiquidacion;
                        doc.Canal = canal;
                        doc.Tipo = tipoDoc;
                        doc.Vendedor = vendedor;
                        doc.Numero = folio;
                        docs.Add(doc);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }


                // Sumo la cantidad de clientes
                oRSMyTable.DoQuery(queryCountClientes);
                if (oRSMyTable != null)
                    cantidadClientesNC += oRSMyTable.RecordCount;

                SAPbouiCOM.StaticText oText;
                oText = oFormDatosPedido.Items.Item("32").Specific; // Cantidad de clientes presentes
                oText.Caption = "Clientes: " + cantidadClientes.ToString() + " en Fac. " + cantidadClientesNC + " en NC";

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return docs;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return null;
            }
        }

        public List<clsTerritorio> obtenerTerritorios()
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                List<clsTerritorio> territorios = new List<clsTerritorio>();
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "SELECT territryID as Code, descript as Name FROM OTER ";


                query += " order by descript";
                //String query = "SELECT DISTINCT T1.territryID as Code, T1.descript as Name FROM ORDR T0 " + //ASPL - 2020.10.01 - Nueva logica de territorios de ORDR.
                //"INNER JOIN OTER T1 ON T0.U_TERRITORIO = T1.territryID WHERE T0.U_TERRITORIO IS NOT NULL AND T0.DocStatus = 'O' ";

                oRSMyTable.DoQuery(query);
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        clsTerritorio terr = new clsTerritorio();
                        String nombre = oRSMyTable.Fields.Item("Name").Value;
                        int cod = oRSMyTable.Fields.Item("Code").Value;

                        terr.IdTerritorio = cod;
                        terr.NombreTerritorio = nombre;
                        terr.Seleccionado = false;
                        territorios.Add(terr);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return territorios;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return null;
            }
        }

        public List<clsDocumento> obtenerPedidos(int codigo, DateTime pFechaDesde, DateTime pFechaHasta, string pLiquidacion, clsCanales pCanal, string pEstadoLiquidacion, string pTerritorio, string pEstadoPedido, string pCardCode, string pPedidoFreeShop, string pItemCode)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                List<clsDocumento> docs = new List<clsDocumento>();
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "select oin.DocEntry,DocNum, DocDate, oin.CardName,case when DocCur = '$' or DocCur = 'UYU' then DocTotal else DocTotalFC end as DocTotal,U_LIQUIDACION,U_CANAL,substring (Comments , 0 ,30) as Comentarios, oin.Confirmed, oin.Pick, 'PE' as Tipo, " +
                "case when (select count(LineNum) from RDR1 as rd1 inner join OITM as oit on oit.ItemCode = rd1.ItemCode where QryGroup1 = 'Y' and rd1.DocEntry = oin.DocEntry) > 0 then 'W' else '' end as FS, osl.SlpName as Vendedor from ORDR as oin " +
                " inner join OCRD as ocr on ocr.CardCode = oin.CardCode " +
                " inner join OSLP as osl on osl.SlpCode = oin.SlpCode ";

                if (!String.IsNullOrEmpty(pEstadoLiquidacion.ToString()) && (pEstadoLiquidacion.ToString().Equals("1") || pEstadoLiquidacion.ToString().Equals("2"))) // Si tiene un Estado de liquidacion para filtrar
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                        query += " inner join [@LIQUIDACIONES] as liq on liq.Code = oin.U_LIQUIDACION where liq.U_ESTADO = '" + pEstadoLiquidacion + "'";
                }

                if (!String.IsNullOrEmpty(pTerritorio.ToString()) && !pTerritorio.ToString().Equals("0")) // Si tiene un valor en el campo Territorio
                {
                    //query += " inner join OTER as ote on ote.territryID = ocr.Territory "; //ASPL - 2021.10.18 - comentado por cambio de logica.
                    query += " inner join OTER as ote on ote.territryID = oin.U_TERRITORIO "; //ASPL - 2021.10.18 - Obtener territorios por ORDR

                    if (query.ToString().Contains("where"))
                        query += " and " + pTerritorio;
                    else
                        query += " where " + pTerritorio;
                }

                if (query.ToString().Contains("where")) // Verifica que no sea un documento cancelado y que el Estado sea "Open"
                    query += "  and oin.Canceled = 'N' and oin.DocStatus = 'O' ";
                else
                    query += "  where oin.Canceled = 'N' and oin.DocStatus = 'O' ";

                if (!String.IsNullOrEmpty(codigo.ToString()) && codigo != 0) // Si el codigo no es vacio filtra por ese campo
                {
                    if (query.ToString().Contains("where"))
                        query += " and DocNum = '" + codigo + "'";
                    else
                        query += " where DocNum = '" + codigo + "'";
                }

                if (!String.IsNullOrEmpty(pFechaDesde.ToString()) && !String.IsNullOrEmpty(pFechaHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                {
                    if (query.ToString().Contains("where"))
                        query += " and DocDueDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and DocDueDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                    else
                        query += " where DocDueDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and DocDueDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                }

                if (!String.IsNullOrEmpty(pLiquidacion.ToString())) // Si tiene un valor en el campo Liquidacion
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                    {
                        if (query.ToString().Contains("where"))
                            query += " and U_LIQUIDACION = '" + pLiquidacion + "'";
                        else
                            query += " where U_LIQUIDACION = '" + pLiquidacion + "'";
                    }
                    else
                    {
                        if (query.ToString().Contains("where")) // Busca los documentos sin Liquidacion
                            query += " and (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                        else
                            query += " where (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                    }
                }

                if (!String.IsNullOrEmpty(pEstadoPedido.ToString()) && !pEstadoPedido.ToString().Equals("0")) // Si tiene un valor en el campo Estado del pedido
                {
                    if (query.ToString().Contains("where"))
                        query += " and oin.Pick = '" + pEstadoPedido + "'";
                    else
                        query += " where oin.Pick = '" + pEstadoPedido + "'";
                }

                if (!String.IsNullOrEmpty(pCardCode.ToString()) && !pCardCode.ToString().Equals("0")) // Si tiene un valor en el campo CardCode
                {
                    if (query.ToString().Contains("where"))
                        query += " and oin.CardCode = '" + pCardCode + "' ";
                    else
                        query += " where oin.CardCode = '" + pCardCode + "' ";
                }

                if (!String.IsNullOrEmpty(pPedidoFreeShop.ToString()) && !pPedidoFreeShop.ToString().Equals("0")) // Si tiene un valor en el campo Pedido FreeShop
                {
                    if (query.ToString().Contains("where"))
                        query += " and (select count(LineNum) from RDR1 as rd1 inner join OITM as oit on oit.ItemCode = rd1.ItemCode where QryGroup1 = 'Y' and rd1.DocEntry = oin.DocEntry) " + pPedidoFreeShop + " 0 ";
                    else
                        query += " where (select count(LineNum) from RDR1 as rd1 inner join OITM as oit on oit.ItemCode = rd1.ItemCode where QryGroup1 = 'Y' and rd1.DocEntry = oin.DocEntry) " + pPedidoFreeShop + " 0 ";
                }

                if (!String.IsNullOrEmpty(pItemCode.ToString()) && !pItemCode.ToString().Equals("0")) // Si tiene un valor en el campo ItemCode, es para filtrar los documentos que tengan un determinado articulo
                {
                    if (query.ToString().Contains("where"))
                        query += " and (select count(LineNum) from RDR1 as rd1 where ItemCode = '" + pItemCode + "' and rd1.DocEntry = oin.DocEntry) > 0 ";
                    else
                        query += " where (select count(LineNum) from RDR1 as rd1 where ItemCode = '" + pItemCode + "' and rd1.DocEntry = oin.DocEntry) > 0 ";
                }

                string whereCanal = "";
                if (!String.IsNullOrEmpty(pCanal.Uno)) // Si tiene un valor en el campo Canal
                    whereCanal += " U_CANAL = '" + pCanal.Uno + "'";
                if (!String.IsNullOrEmpty(pCanal.Dos)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Dos + "'";
                if (!String.IsNullOrEmpty(pCanal.Tres)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Tres + "'";
                if (!String.IsNullOrEmpty(pCanal.Cuatro)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cuatro + "'";
                if (!String.IsNullOrEmpty(pCanal.Cinco)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cinco + "'";
                if (!String.IsNullOrEmpty(pCanal.Seis)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Seis + "'";
                if (!String.IsNullOrEmpty(pCanal.Siete)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Siete + "'";
                if (!String.IsNullOrEmpty(pCanal.Ocho)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Ocho + "'";
                if (!String.IsNullOrEmpty(pCanal.Nueve)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Nueve + "'";

                if (!String.IsNullOrEmpty(whereCanal.ToString()))
                    whereCanal = " and (" + whereCanal.ToString() + ") ";

                query += whereCanal + " order by DocNum,DocDate";

                oRSMyTable.DoQuery(query);
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        clsDocumento doc = new clsDocumento();
                        String cli = oRSMyTable.Fields.Item("CardName").Value;
                        int cod = oRSMyTable.Fields.Item("DocNum").Value;
                        int docEntry = oRSMyTable.Fields.Item("DocEntry").Value;
                        DateTime fecha = oRSMyTable.Fields.Item("DocDate").Value;
                        double monto = oRSMyTable.Fields.Item("DocTotal").Value;
                        string canal = oRSMyTable.Fields.Item("U_CANAL").Value;
                        string nroLiquidacion = oRSMyTable.Fields.Item("U_LIQUIDACION").Value;
                        string tipoDoc = oRSMyTable.Fields.Item("Tipo").Value;
                        string confirmed = oRSMyTable.Fields.Item("Confirmed").Value;
                        string pick = oRSMyTable.Fields.Item("Pick").Value; //ASPL. 2019-09-09, Nuevo req. cambio de campo Confirmed por Pick.
                        string comentarios = oRSMyTable.Fields.Item("Comentarios").Value;
                        string freeShop = oRSMyTable.Fields.Item("FS").Value;
                        string vendedor = oRSMyTable.Fields.Item("Vendedor").Value;

                        doc.Cliente = cli;
                        doc.DocNum = cod;
                        doc.DocEntry = docEntry;
                        doc.Fecha = fecha.Date.ToShortDateString();
                        doc.Monto = monto;
                        doc.NroLiquidacion = nroLiquidacion;
                        doc.Canal = canal;
                        doc.Tipo = tipoDoc;
                        doc.Confirmado = confirmed;
                        doc.Pickeado = pick;
                        doc.FreeShop = freeShop;
                        doc.Comentarios = comentarios;
                        doc.Vendedor = vendedor;
                        docs.Add(doc);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                // Recorro las Ofertas de Ventas
                query = "select oin.DocEntry,DocNum, DocDate, oin.CardName,case when DocCur = '$' or DocCur = 'UYU' then DocTotal else DocTotalFC end as DocTotal,U_LIQUIDACION,U_CANAL,substring (Comments , 0 ,30) as Comentarios, oin.Confirmed, oin.Pick,'OF' as Tipo, " +
                "case when (select count(LineNum) from QUT1 as rd1 inner join OITM as oit on oit.ItemCode = rd1.ItemCode where QryGroup1 = 'Y' and rd1.DocEntry = oin.DocEntry) > 0 then 'W' else '' end as FS, osl.SlpName as Vendedor from OQUT as oin " +
                "inner join OCRD as ocr on ocr.CardCode = oin.CardCode " +
                " inner join OSLP as osl on osl.SlpCode = oin.SlpCode ";

                if (!String.IsNullOrEmpty(pEstadoLiquidacion.ToString()) && (pEstadoLiquidacion.ToString().Equals("1") || pEstadoLiquidacion.ToString().Equals("2"))) // Si tiene un Estado de liquidacion para filtrar
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                        query += " inner join [@LIQUIDACIONES] as liq on liq.Code = oin.U_LIQUIDACION where liq.U_ESTADO = '" + pEstadoLiquidacion + "'";
                }

                if (!String.IsNullOrEmpty(pTerritorio.ToString()) && !pTerritorio.ToString().Equals("0")) // Si tiene un valor en el campo Territorio
                {
                    query += " inner join OTER as ote on ote.territryID = ocr.Territory ";
                    if (query.ToString().Contains("where"))
                        query += " and " + pTerritorio;
                    else
                        query += " where " + pTerritorio;
                }

                if (query.ToString().Contains("where")) // Verifica que no sea un documento cancelado y que el Estado sea "Open"
                    query += "  and oin.Canceled = 'N' and oin.DocStatus = 'O' ";
                else
                    query += "  where oin.Canceled = 'N' and oin.DocStatus = 'O' ";

                if (!String.IsNullOrEmpty(codigo.ToString()) && codigo != 0) // Si el codigo no es vacio filtra por ese campo
                {
                    if (query.ToString().Contains("where"))
                        query += " and DocNum = '" + codigo + "'";
                    else
                        query += " where DocNum = '" + codigo + "'";
                }

                if (!String.IsNullOrEmpty(pFechaDesde.ToString()) && !String.IsNullOrEmpty(pFechaHasta.ToString())) // Si las fechas no son vacias filtra por ese rango
                {
                    if (query.ToString().Contains("where"))
                        query += " and DocDueDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and DocDueDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                    else
                        query += " where DocDueDate >='" + pFechaDesde.ToString("yyyy-MM-dd") + "' and DocDueDate <='" + pFechaHasta.ToString("yyyy-MM-dd") + "'";
                }

                if (!String.IsNullOrEmpty(pLiquidacion.ToString())) // Si tiene un valor en el campo Liquidacion
                {
                    if (!pLiquidacion.ToString().Equals("0"))
                    {
                        if (query.ToString().Contains("where"))
                            query += " and U_LIQUIDACION = '" + pLiquidacion + "'";
                        else
                            query += " where U_LIQUIDACION = '" + pLiquidacion + "'";
                    }
                    else
                    {
                        if (query.ToString().Contains("where")) // Busca los documentos sin Liquidacion
                            query += " and (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                        else
                            query += " where (U_LIQUIDACION is null or U_LIQUIDACION = '')";
                    }
                }

                if (!String.IsNullOrEmpty(pCardCode.ToString()) && !pCardCode.ToString().Equals("0")) // Si tiene un valor en el campo CardCode
                {
                    if (query.ToString().Contains("where"))
                        query += " and oin.CardCode = '" + pCardCode + "' ";
                    else
                        query += " where oin.CardCode = '" + pCardCode + "' ";
                }

                if (!String.IsNullOrEmpty(pPedidoFreeShop.ToString()) && !pPedidoFreeShop.ToString().Equals("0")) // Si tiene un valor en el campo Pedido FreeShop
                {
                    if (query.ToString().Contains("where"))
                        query += " and (select count(LineNum) from QUT1 as rd1 inner join OITM as oit on oit.ItemCode = rd1.ItemCode where QryGroup1 = 'Y' and rd1.DocEntry = oin.DocEntry) " + pPedidoFreeShop + " 0 ";
                    else
                        query += " where (select count(LineNum) from QUT1 as rd1 inner join OITM as oit on oit.ItemCode = rd1.ItemCode where QryGroup1 = 'Y' and rd1.DocEntry = oin.DocEntry) " + pPedidoFreeShop + " 0 ";
                }

                if (!String.IsNullOrEmpty(pItemCode.ToString()) && !pItemCode.ToString().Equals("0")) // Si tiene un valor en el campo ItemCode, es para filtrar los documentos que tengan un determinado articulo
                {
                    if (query.ToString().Contains("where"))
                        query += " and (select count(LineNum) from RDR1 as rd1 where ItemCode = '" + pItemCode + "' and rd1.DocEntry = oin.DocEntry) > 0 ";
                    else
                        query += " where (select count(LineNum) from RDR1 as rd1 where ItemCode = '" + pItemCode + "' and rd1.DocEntry = oin.DocEntry) > 0 ";
                }

                whereCanal = "";
                if (!String.IsNullOrEmpty(pCanal.Uno)) // Si tiene un valor en el campo Canal
                    whereCanal += " U_CANAL = '" + pCanal.Uno + "'";
                if (!String.IsNullOrEmpty(pCanal.Dos)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Dos + "'";
                if (!String.IsNullOrEmpty(pCanal.Tres)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Tres + "'";
                if (!String.IsNullOrEmpty(pCanal.Cuatro)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cuatro + "'";
                if (!String.IsNullOrEmpty(pCanal.Cinco)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Cinco + "'";
                if (!String.IsNullOrEmpty(pCanal.Seis)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Seis + "'";
                if (!String.IsNullOrEmpty(pCanal.Siete)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Siete + "'";
                if (!String.IsNullOrEmpty(pCanal.Ocho)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Ocho + "'";
                if (!String.IsNullOrEmpty(pCanal.Nueve)) // Si tiene un valor en el campo Canal
                    whereCanal += " or U_CANAL = '" + pCanal.Nueve + "'";

                if (!String.IsNullOrEmpty(whereCanal.ToString()))
                    whereCanal = " and (" + whereCanal.ToString() + ") ";

                query += whereCanal + " order by DocNum,DocDate";

                oRSMyTable.DoQuery(query);
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        clsDocumento doc = new clsDocumento();
                        String cli = oRSMyTable.Fields.Item("CardName").Value;
                        int cod = oRSMyTable.Fields.Item("DocNum").Value;
                        int docEntry = oRSMyTable.Fields.Item("DocEntry").Value;
                        DateTime fecha = oRSMyTable.Fields.Item("DocDate").Value;
                        double monto = oRSMyTable.Fields.Item("DocTotal").Value;
                        string canal = oRSMyTable.Fields.Item("U_CANAL").Value;
                        string nroLiquidacion = oRSMyTable.Fields.Item("U_LIQUIDACION").Value;
                        string tipoDoc = oRSMyTable.Fields.Item("Tipo").Value;
                        string confirmed = oRSMyTable.Fields.Item("Confirmed").Value;
                        string pick = oRSMyTable.Fields.Item("Pick").Value; //ASPL. 2019-09-09, Nuevo req., nuevo campo.
                        string comentarios = oRSMyTable.Fields.Item("Comentarios").Value;
                        string freeShop = oRSMyTable.Fields.Item("FS").Value;
                        string vendedor = oRSMyTable.Fields.Item("Vendedor").Value;

                        doc.Cliente = cli;
                        doc.DocNum = cod;
                        doc.DocEntry = docEntry;
                        doc.Fecha = fecha.Date.ToShortDateString();
                        doc.Monto = monto;
                        doc.NroLiquidacion = nroLiquidacion;
                        doc.Canal = canal;
                        doc.Tipo = tipoDoc;
                        doc.Confirmado = confirmed;
                        doc.Pickeado = pick;
                        doc.FreeShop = freeShop;
                        doc.Comentarios = comentarios;
                        doc.Vendedor = vendedor;
                        docs.Add(doc);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return docs;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return null;
            }
        }

        public List<clsCheque> obtenerCheques(string pCuenta)
        {
            SAPbobsCOM.Recordset oRSMyTable = null;
            try
            {
                List<clsCheque> docs = new List<clsCheque>();
                oRSMyTable = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                String query = "select CheckKey as NumSecuencia, CheckNum as NumeroCheque,BankCode as Banco, Branch as Sucursal,CheckDate as Fecha,Currency, CheckSum as Monto from OCHH as och " +
                "WHERE BankAcct = '" + pCuenta.ToString() + "' and not CheckKey in (select Code from [@CHEQUES_ACREDITADOS] where U_ACREDITADO = 'Y') " +
                "order by CheckDate,NumSecuencia ";

                oRSMyTable.DoQuery(query);
                while (!oRSMyTable.EoF)
                {
                    try
                    {
                        clsCheque doc = new clsCheque();
                        String banco = oRSMyTable.Fields.Item("Banco").Value;
                        String sucursal = oRSMyTable.Fields.Item("Sucursal").Value;
                        int numSecuencia = oRSMyTable.Fields.Item("NumSecuencia").Value;
                        int numCheque = oRSMyTable.Fields.Item("NumeroCheque").Value;
                        DateTime fecha = oRSMyTable.Fields.Item("Fecha").Value;
                        double monto = oRSMyTable.Fields.Item("Monto").Value;
                        string moneda = oRSMyTable.Fields.Item("Currency").Value;

                        doc.Banco = banco;
                        doc.Moneda = moneda;
                        doc.NumCheque = numCheque;
                        doc.Fecha = fecha.Date.ToShortDateString();
                        doc.Monto = monto;
                        doc.NumSecuencia = numSecuencia;
                        doc.Sucursal = sucursal;
                        doc.Acreditado = "N";
                        docs.Add(doc);
                    }
                    catch (Exception ex)
                    { }
                    oRSMyTable.MoveNext();
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return docs;
            }
            catch (Exception ex)
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRSMyTable);
                oRSMyTable = null;
                return null;
            }
        }

        // Abre el documento que recibe por parametro
        public bool abrirDocumento(String pDocEntryDocumento, String pDocNumDocumento, String pTipoDocumento)
        {
            bool res = false;
            try
            {
                Thread.Sleep(500);

                // 3332 Pantalla de Autorizaciones
                // 2053 Factura de Deudores
                // 2054 Factura de Deudores + Pago
                // 2056 Factura de Reserva
                // 2314 Factura de Reserva Acreedores
                // 2049 Oferta de Ventas
                // 2308 Factura de Proveedores
                // 2065 Boleta

                string tipoForm = "";
                string caseSwitch = pTipoDocumento; // Numero de Formulario de documento
                switch (caseSwitch)
                {
                    case "BO":
                        tipoForm = "2065"; // Boleta
                        break;
                    case "FR":
                        tipoForm = "2056"; // Factura de Reserva
                        break;
                    case "FP":
                        tipoForm = "2054"; // Factura de Deudores + Pago
                        break;
                    case "FD":
                        tipoForm = "2053"; // Factura de Deudores
                        break;
                    case "PE":
                        tipoForm = "2050"; // Pedido
                        break;
                    case "OF":
                        tipoForm = "2049"; // Oferta de Venta
                        break;
                    case "ND":
                        tipoForm = "2064"; // Notas de Débitos
                        break;
                    case "NC":
                        tipoForm = "2055"; // Notas de Créditos
                        break;
                    default:
                        tipoForm = "2053"; // Factura de deudores
                        break;
                }


                SBO_Application.ActivateMenuItem(tipoForm); // Abro el formulario de busqueda correspondiente al documento creado
                SAPbouiCOM.Form fo = SBO_Application.Forms.ActiveForm;
                fo.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                SAPbouiCOM.EditText oStatic;
                oStatic = fo.Items.Item("8").Specific; // Numero del documento
                oStatic.Value = pDocNumDocumento; // Numero de la factura (el docEntry)

                fo.Items.Item("1").Click();
                ////SBO_Application.Menus.Item("520").Enabled = true; // PRUEBA
                ////SBO_Application.Menus.Item("520").Activate(); // Open printing dialog
                ////fo.Items.Item("1").Click(); // Cierro el formulario de la factura

                res = true;

                return res;
            }
            catch (Exception ex)
            {
                if (guardaLog == true)
                    guardaLogProceso(pTipoDocumento.ToString(), pDocNumDocumento, "ERROR al Visualizar el documento", ex.Message.ToString()); // Guarda log del Proceso
            }
            return res;
        }
        #endregion

    }
}
