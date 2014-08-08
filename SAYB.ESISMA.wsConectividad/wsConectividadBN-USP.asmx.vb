Imports System.Web.Services
Imports System.Web.Services.Protocols
Imports System.ComponentModel
Imports System.Web
Imports System.Data
Imports System.Data.SqlClient
Imports System.Math
Imports System.Array
Imports System.Diagnostics

' Para permitir que se llame a este servicio web desde un script, usando ASP.NET AJAX, quite la marca de comentario de la línea siguiente.
' <System.Web.Script.Services.ScriptService()> _
<System.Web.Services.WebService(Namespace:="http://tempuri.org/")> _
<System.Web.Services.WebServiceBinding(ConformsTo:=WsiProfiles.BasicProfile1_1)> _
<ToolboxItem(False)> _
Public Class wsConectividad_USP
    Inherits System.Web.Services.WebService

#Region "Variables"
    Dim cnnConexion As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
    Private grefTransaction As SqlTransaction
    Private gcmdComando As New SqlCommand
    Private objEntradas As New clEntradas
    Public lValoresLetraCambio(10) As Decimal
#End Region

#Region "Enumerados"

    Private Enum eCodigoConvenio
        NoAsignado = 0
        LetraCambio = 1
        Facturas = 2
        CursosLibres = 3
        Carnets = 4
        VidaEstudiantil = 5
    End Enum

    Private Enum eTipoIdentificacion
        Otros = 0
        Cedula = 1
        Id_Banco = 99
    End Enum


    Private Enum eTipoLetraCambio
        BachLic = 1
        Maestria = 0
    End Enum

    Private Enum eTipoInteresMoratorio
        Interes5 = 0
        Interes3_9 = 1
    End Enum


    Private Enum eTipoInteres
        Corrientes = 1
        Moratorios = 2
    End Enum


    Private Enum eColumnasArregloLetras
        SaldoLetra = 0
        MontoInteresOrdinario = 1
        MontoInteresMoratorio = 2
        MontoTotal = 3
        DiasOrdinarios = 4
        DiasMoratorios = 5
        TipoLetraCambio = 6
        IdLetraCambio = 7
        TotalPagado = 8
        TotalInteresOrdinario = 9
        TotalInteresMoratorio = 10

    End Enum


    Private Enum eColumnasArregloDetalleFactura
        Materias = 0
        Matricula = 1
        Seguro = 2
        MontoTotal = 3
        Descuento = 4
        TotalConDescuento = 5
        TotalSinDescuento = 6
    End Enum

    Private Enum eEstadoLetra
        Pendiente = 0
        Cancelada = 1
    End Enum

    Private Enum eMensajesErrorConsultaServicios
        ConvenioNoExiste = 1
        EstudianteNoExiste = 2
        EstudianteConBeca = 3
        NoHayPendientesPago = 4

    End Enum

#End Region

#Region "Constantes"

    Const cEstadoMatriculado As String = "Matriculado"
    Const cPorMatricular As String = "Por Matricular"
    Const cSinDefinir As String = "S/D"
    Const cNoAplica As String = "N/A"
    Const cAnulada As Integer = 2
    Const cCero As Integer = 0
    Const cSinNota As Double = 0.0
    Const cNoFacturado As Integer = 0
    Const cTotal As Double = 0.0
    Const cEfectivo As Double = 0.0
    Const cTarjeta As Double = 0.0
    Const cCancelada As Integer = 0
    Const CDescuento As Double = 0.0
    Const cTotalSinDescuento As Double = 0
    Const cCreo As String = "BN-CONECTIVIDAD"
    Const cObservaciones As String = "RECIBO REVERSADO POR BN-CONECTIVIDAD"
    Const cPorcentajeDescuento = 0.03
    Const cCreoCredomatic As String = "Pago Online"
    Const cValorMatricula As String = "M"
    Const cValorAbono As String = "A"
    Const cValorOtros As String = "S"
#End Region

#Region "Seguridad"
    Const cstrUsuario As String = "BNCR-USP"
    Const cstrClaveUsuario As String = "5@nt@9@ul@8ncr"
#End Region


#Region "Valores de Intereses"
    Const cInteresOrdinarioDiario As Decimal = 0.001
    Const cInteresMoratorioDiario_5 As Decimal = 0.0016667
    Const cInteresMoratorioDiario_3_9 As Decimal = 0.0013
    Const cDiasIncobrable As Decimal = 1461
    Const cDiasCobroJudicial As Decimal = 365
#End Region

#Region "Mensajes del Sistema"
    Const cCodMensaje00 As String = "00"
    Const cCodMensaje51 As String = "51"
    Const cCodMensaje50 As String = "50"
    Const cCodMensaje99 As String = "99"
    Const strMensaje00 As String = "Consulta Exitosa"
    Const strMensaje01 As String = "El Pago se realizó satisfactoriamente"
    Const strMensaje99 As String = "El Pago no se realizó satisfactoriamente"
    Const strMensaje02 As String = "No se ha realizado ningun pago con los datos indicados"
    Const strMensaje05 As String = "El convenio no está registrado"
    Const strMensaje06 As String = "Estudiante no esta registrado en el sistema"
    Const strMensaje07 As String = "Ud posee una BECA activa, tramiar en la USP"
    Const strMensaje08 As String = "No hay pendientes de pago en el convenio"
    Const strMensaje12 As String = "El servicio no se encuentra disponible"
    Const strMensaje13 As String = "Número de factura inexistente o inválido"
    Const strMensaje14 As String = "La transacción no se logro terminar correctamente"
    Const strMensaje15 As String = "Tipo de factura inválido"
    Const strMensaje16 As String = "Recibo inválido o inexistente"
    Const StrMensaje17 As String = "El Monto mínimo del pago es de: "
    Const strMensaje18 As String = "El convenio solo permite el monto exacto"
    Const strMensaje19 As String = "La reversión se realizó exitosamente"
    Const strMensaje20 As String = "Error en usuario o contraseña"
    Const strMensaje21 As String = "El monto a pagar es menor al mínimo a pagar en una letra de cambio"


#End Region

#Region "Procesos Publicos"
    'Funcion que devuelve en le Dataset los Datatables correspondientes a las clases de ConsultarServiciosPendientesReq,ConsultarServiciosPendientes
    <WebMethod()> _
    Public Function pfConsultaRecibosPendientes(ByVal nCodConvenio As Byte, ByVal strTipoLlave As Byte, ByVal strLLaveAcceso As String, ByVal nNumCuotas As String) As DataSet
        Dim ldtsDatos As New DataSet
        Dim lestadoMensaje As Boolean = False
        Dim lvalor As Byte = 0

        Try
            'Si el servicio esta diponible opera normalmente, si no elecuta la función publica de servicio disponible
            If lfServicioDisponible() = True Then
                'Si el convenio esta definido en los convenios de la Universidad continua
                If lfExisteConvenio(nCodConvenio) = True Then
                    'Si el Estudiante existe en la Universidad continúa
                    If lfExisteEstudiante(strLLaveAcceso) Then

                        'Primero detrmina si existen Pendientes del convenio de entrada.
                        If lfExistenPendientes(nCodConvenio, strLLaveAcceso, strTipoLlave) Then

                            Select Case nCodConvenio

                                Case eCodigoConvenio.Facturas

                                    'Si el Estudiante posee una beca no puede pagar por IB
                                    If Not pfBecaEstudiante(strLLaveAcceso) Then

                                        ldtsDatos = lfConsultaPendientes(eCodigoConvenio.Facturas, strLLaveAcceso, strTipoLlave)
                                    Else
                                        'Registrar el código de respuesta para indicar que el Estudiante posee BECA
                                        lestadoMensaje = True
                                        lvalor = eMensajesErrorConsultaServicios.EstudianteConBeca
                                    End If


                                Case eCodigoConvenio.LetraCambio
                                    ldtsDatos = lfConsultaPendientes(eCodigoConvenio.LetraCambio, strLLaveAcceso, strTipoLlave)

                                Case eCodigoConvenio.Carnets
                                    'El Convenio no esta disponible aún
                                    lestadoMensaje = True
                                    lvalor = eMensajesErrorConsultaServicios.ConvenioNoExiste
                                Case eCodigoConvenio.CursosLibres
                                    'El Convenio no esta disponible aún
                                    lestadoMensaje = True
                                    lvalor = eMensajesErrorConsultaServicios.ConvenioNoExiste
                                Case eCodigoConvenio.VidaEstudiantil
                                    'El Convenio no esta disponible aún
                                    lestadoMensaje = True
                                    lvalor = eMensajesErrorConsultaServicios.ConvenioNoExiste

                            End Select


                        Else
                            'Respuesta para no hay recibos pendientes
                            lestadoMensaje = True
                            lvalor = eMensajesErrorConsultaServicios.NoHayPendientesPago
                        End If
                        'Consulta si existen documentos pendientes de pago
                        'del estudiante que esta aplicando la solicitudo de pago.
                    Else
                        'Registrar el código de respuesta para indicar que el Estudiante no está registrado en la Universidad
                        lestadoMensaje = True
                        lvalor = eMensajesErrorConsultaServicios.EstudianteNoExiste
                    End If

                Else
                    'Mensaje de convenio no existe
                    lestadoMensaje = True
                    lvalor = eMensajesErrorConsultaServicios.ConvenioNoExiste

                End If
                'Si se ha detectado algun mensaje de error el sistema lo captura e informa
                If lvalor > 0 Then
                    ldtsDatos = lfInformacionTipoError(lvalor, strLLaveAcceso)
                End If
                'Si no hay servicio disponible invoca la funcion publica del sistema
            Else
                ldtsDatos = pfServicioDisponible()

            End If


        Catch ex As Exception

            Throw ex
        Finally


            Me.cnnConexion.Close()
        End Try
        Return ldtsDatos
    End Function

    'Funcion que ejecuta el pago de servicios
    <WebMethod()> _
    Public Function pfEjecutarPagoServicio(ByVal strUsuario As String, _
                                           ByVal strClaveUsuario As String, _
                                           ByVal strTipoLlave As String, _
                                           ByVal strLlaveAcceso As String, _
                                           ByVal nCodAgencia As Byte, _
                                           ByVal nCodConvenio As Byte, _
                                           ByVal nTipoTransaccion As Byte, _
                                           ByVal strPeriodo As String, _
                                           ByVal nMontoPagado As Decimal, _
                                           ByVal nMontoTotalRecibo As Decimal, _
                                           ByVal nNumFactura As Integer, _
                                           ByVal strRecibo As String, _
                                           ByVal nSelfVerificacion As Byte, _
                                           ByVal strCodMoneda As String, _
                                           ByVal nNumComprobante As Integer, _
                                           ByVal nNumeroCuotas As Integer, _
                                           ByVal nTipoPago As Byte, _
                                           ByVal nCodBanco As Integer, _
                                           ByVal nNumCheque As Long, _
                                           ByVal nNumCuenta As Long, _
                                           ByVal nMontoTotal As Decimal, _
                                           ByVal nNumNotaCredito As Integer) As DataSet
        Dim ldtsDatos As New DataSet
        'Procedimiento para crear las tablas del DataSet de retorno del la CLase Pago de Servicios
        'Creamos el data set que contendrá la información del Afiliado para el archivo de Envío
        Dim dtsPagoServicios As New Data.DataSet("dtsPagoServicios")
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim dtbDetalle As Data.DataTable = New DataTable("Detalle")
        Dim dtbSubDetalle As Data.DataTable = New DataTable("SubDetalle")
        Dim ldtrFila As DataRow
        Dim ldtrConceptosInformativos As DataRow
        Dim ldtrDetalle As DataRow
        Dim lstrNumeroTransaccion As String = String.Empty
        Dim lValoresFacturacion(7) As Decimal
        Dim nTotalConceptosPago As Integer = 0
        Dim nMontoTotalFacturacion As Decimal
        Dim strDescripcionConceptp As String
        Dim nEstID As Integer
        Dim lExitosa As Boolean
        Dim nEnlID As Integer = 0
        Dim nCodRespuesta As Integer = 0
        Dim strDescRespuesta As String = ""
        Dim nDescuento As Decimal = 0.0
        Dim nTotalSinDescuento As Decimal = 0.0
        Dim nTipoEntidad As Integer = 1
        Dim strNombreTarjeta As String = "N/A"
        Dim strNumeroTarjeta As String = "N/A"
        Dim pDescuento As Integer = 3
        Dim pTipoCompra As Integer = 0



        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nCantConceptosInformativos", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("nCantConceptosPago", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("strNumTransaccionPago", Type.GetType("System.String"))
        dtsPagoServicios.Tables.Add(dtbEncabezado)
        '-------------------------------------------------------------------------------------

        'Tabla del SubDetalle correspondiente a la clase conceptos Informativos
        '--------------------------------------------------------------------------------------
        dtbSubDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbSubDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbSubDetalle.Columns.Add("strValorConcepto", Type.GetType("System.String"))
        dtsPagoServicios.Tables.Add(dtbSubDetalle)
        '--------------------------------------------------------------------------------------

        'TABLA DEL DETALLE correspondiente a la clase de Conceptos de Pago
        '-------------------------------------------------------------------------------------
        dtbDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Int32"))
        dtbDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nMontoConcepto", Type.GetType("System.Decimal"))
        dtsPagoServicios.Tables.Add(dtbDetalle)
        '-------------------------------------------------------------------------------------
        Try
            'Si el servicio está diponible ejecuta todo el proceso; si no devuelve el encabezado del DTS 
            'con el código de respuesta 50

            If lfServicioDisponible() = True Then

                'Valida que el usuario  del Banco sea el Correcto
                If lfValidaUsuario(strUsuario, strClaveUsuario) = True Then
                    'Segun el tipo de convenio realiza el pago
                    Select nCodConvenio  'Select Case nCodConvenio
                        Case eCodigoConvenio.Facturas
                            'Carga los detelles de la factura antes de aplicar los cambios
                            If lfExisteFactura(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXFactura(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje13
                            End If

                            If lExitosa = True Then
                                'Obtenemos los valores del pago de la factura
                                lValoresFacturacion = pfDetalleFacturacion(nEstID)
                                nMontoTotalFacturacion = lValoresFacturacion(eColumnasArregloDetalleFactura.TotalConDescuento) + lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula) + lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro)
                                'Verificamos los conceptos de pago a pagar y realizan los procesos necesarios para realizar el pago
                                If nMontoTotalFacturacion <> 0 And nMontoTotalFacturacion = nMontoTotal Then
                                    If lValoresFacturacion(eColumnasArregloDetalleFactura.TotalConDescuento) <> 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfActualizaMateriasMatriculadas(nEstID, nTipoEntidad)
                                    End If
                                    If lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula) <> 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfIngresaPagoMatriculaEstudiante(nEstID, nNumFactura, eColumnasArregloDetalleFactura.Matricula)
                                    End If
                                    If lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro) <> 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfIngresaPagoSeguroEstudiante(nEstID, nNumFactura)
                                    End If
                                Else
                                    lExitosa = False
                                    nCodRespuesta = cCodMensaje51
                                    strDescRespuesta = "El monto introducido es inválido debe de ser igual al monto de la factura y mayor a cero."
                                End If
                                If lExitosa Then
                                    If nTotalConceptosPago > 0 Then
                                        'Se actualiza el estado a cancelada por BN-CONECTIVDAD
                                        nDescuento = lValoresFacturacion(eColumnasArregloDetalleFactura.Descuento)
                                        nTotalSinDescuento = lValoresFacturacion(eColumnasArregloDetalleFactura.MontoTotal) + lValoresFacturacion(eColumnasArregloDetalleFactura.Descuento)

                                        If pfActualizaEstadoFactura(nNumFactura, lValoresFacturacion(eColumnasArregloDetalleFactura.Descuento), nMontoTotalFacturacion, nTotalSinDescuento, cCreo, strNumeroTarjeta, nTipoEntidad, pDescuento, pTipoCompra) Then
                                            'Se genera un recibo de pago de matricula
                                            lstrNumeroTransaccion = strLlaveAcceso
                                            If lstrNumeroTransaccion > 0 Then
                                                'Se inserta el pago a la bitacora de pagos de bancos
                                                If lfInsertarBitacoraPagosBancos(0, nNumFactura, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoPagado, cCreo) = 1 Then
                                                    If lfInsertaInterfazEncabezado(nEstID, _
                                                                                   strLlaveAcceso, _
                                                                                   cValorMatricula, _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.TotalSinDescuento), _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula), _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro), _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.Materias),
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.TotalSinDescuento), _
                                                                                   pDescuento, lValoresFacturacion(eColumnasArregloDetalleFactura.Descuento), _
                                                                                   cCero, cCero, lValoresFacturacion(eColumnasArregloDetalleFactura.Descuento), _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.MontoTotal), _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.MontoTotal), _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   cCero, _
                                                                                   lValoresFacturacion(eColumnasArregloDetalleFactura.MontoTotal), _
                                                                                   cCreo, _
                                                                                   cCero, _
                                                                                   "NR", _
                                                                                   cCero, _
                                                                                   cCero) Then
                                                        Dim dtsDetalle As DataSet
                                                        Dim bs_ine_id As Integer
                                                        Dim bs_ind_cod_centro_costo As String
                                                        Dim bs_ind_mto_detalle As Decimal
                                                        Dim bs_ind_ind_tipo_detalle As String
                                                        Dim nPlanID As Integer
                                                        'Carga el Detalle de los cursos matriculados en el detalle de Matricula
                                                        dtsDetalle = pfObtenerDetallesCursosFactura(strLlaveAcceso)

                                                        'Si Existen cursos Matriculados se procede a registrar cada uno en la interfaz
                                                        If lValoresFacturacion(eColumnasArregloDetalleFactura.Materias) > 0 Then

                                                            If dtsDetalle.Tables(0).Rows.Count > 0 Then
                                                                bs_ind_ind_tipo_detalle = "C"
                                                                For i = 0 To dtsDetalle.Tables(0).Rows.Count - 1
                                                                    bs_ine_id = dtsDetalle.Tables(0).Rows(i)("bs_ine_id")
                                                                    bs_ind_cod_centro_costo = dtsDetalle.Tables(0).Rows(i)("CENTRO_COSTO")
                                                                    bs_ind_mto_detalle = dtsDetalle.Tables(0).Rows(i)("COSTO")
                                                                    lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                                Next
                                                            End If
                                                        End If
                                                        If lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula) > 0 Then
                                                            nPlanID = fnPlanAcademico(nEstID)
                                                            bs_ind_ind_tipo_detalle = "M"
                                                            bs_ine_id = dtsDetalle.Tables(0).Rows(0)("bs_ine_id")
                                                            bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlanID)
                                                            bs_ind_mto_detalle = lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula)
                                                            lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                        End If
                                                        If lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro) > 0 Then
                                                            nPlanID = fnPlanAcademico(nEstID)
                                                            bs_ind_ind_tipo_detalle = "S"
                                                            bs_ine_id = dtsDetalle.Tables(0).Rows(0)("bs_ine_id")
                                                            bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlanID)
                                                            bs_ind_mto_detalle = lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro)
                                                            lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                        End If
                                                    End If
                                                Else
                                                    lExitosa = False
                                                    nCodRespuesta = cCodMensaje51
                                                    strDescRespuesta = "No se registro el pago de recibo."
                                                End If
                                            Else
                                                lExitosa = False
                                                nCodRespuesta = cCodMensaje51
                                                strDescRespuesta = "El número de transacción del pago es inválido."

                                            End If
                                        Else
                                            lExitosa = False
                                            nCodRespuesta = cCodMensaje51
                                            strDescRespuesta = "No se logró actualizar la factura solicitada"
                                        End If
                                    Else
                                        lExitosa = False
                                        nCodRespuesta = cCodMensaje51
                                        strDescRespuesta = "No hay Conceptos de pago pendientes para la factura solicitada."
                                    End If
                                End If

                                'Else

                                'End If

                                If lExitosa Then
                                    'crea la información para el Dataset de retorno
                                    'Crea la fila del encabezado ****************************************************************
                                    ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                    ldtrFila("strCodRespuesta") = cCodMensaje00
                                    ldtrFila("strDescripcionRespuesta") = strMensaje01
                                    ldtrFila("nCantConceptosInformativos") = 1
                                    ldtrFila("nCantConceptosPago") = nTotalConceptosPago
                                    ldtrFila("strNumTransaccionPago") = lstrNumeroTransaccion
                                    dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                    '********************************************************************************************
                                    'Crea la fila del Detalle de la Factura ****************************************************************

                                    'Si existen valores en el pago de materias con descuento los agrega al table de response
                                    If lValoresFacturacion(eColumnasArregloDetalleFactura.TotalConDescuento) Then
                                        ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                        ldtrDetalle("nConsecutivo") = 1
                                        ldtrDetalle("nCodConcepto") = 1
                                        ldtrDetalle("strDescripcionConcepto") = "Materias Matriculadas "
                                        ldtrDetalle("nMontoConcepto") = lValoresFacturacion(eColumnasArregloDetalleFactura.TotalConDescuento)
                                        dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                    End If
                                    'Si existen valores en el pago de matricula los agrega al table de response
                                    If lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula) <> 0 Then
                                        ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                        ldtrDetalle("nConsecutivo") = 2
                                        ldtrDetalle("nCodConcepto") = 2
                                        ldtrDetalle("strDescripcionConcepto") = "Monto Matricula"
                                        ldtrDetalle("nMontoConcepto") = lValoresFacturacion(eColumnasArregloDetalleFactura.Matricula)
                                        dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                    End If
                                    'Si existen valores para el concepto de pago de seguro los agrega al table de response
                                    If lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro) <> 0 Then
                                        ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                        ldtrDetalle("nConsecutivo") = 3
                                        ldtrDetalle("nCodConcepto") = 3
                                        ldtrDetalle("strDescripcionConcepto") = "Seguro Estudiantil "
                                        ldtrDetalle("nMontoConcepto") = lValoresFacturacion(eColumnasArregloDetalleFactura.Seguro)
                                        dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                    End If
                                    'Crea la fila del Detalle de la Factura para conceptos informativos ****************************************************************
                                    'Crea el detalle referente a los conceptos Informativos
                                    '***************************************************************************************
                                    ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                    ldtrConceptosInformativos("nCodConcepto") = 1
                                    ldtrConceptosInformativos("strDescripcionConcepto") = "Fact. Matricula"
                                    ldtrConceptosInformativos("strValorConcepto") = nNumFactura
                                    dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)

                                Else
                                    ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                    ldtrFila("strCodRespuesta") = nCodRespuesta
                                    ldtrFila("strDescripcionRespuesta") = strDescRespuesta
                                    ldtrFila("nCantConceptosInformativos") = 0
                                    ldtrFila("nCantConceptosPago") = 1
                                    ldtrFila("strNumTransaccionPago") = 0
                                    dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                    '********************************************************************************************

                                    'Crea el detalle referente a los conceptos Informativos
                                    '***************************************************************************************
                                    ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                    ldtrConceptosInformativos("nCodConcepto") = cCodMensaje99
                                    ldtrConceptosInformativos("strDescripcionConcepto") = strMensaje02
                                    ldtrConceptosInformativos("strValorConcepto") = 0
                                    dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                    '*************************************************************************************

                                    'Crea la fila del Detalle de la Factura ****************************************************************
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 0
                                    ldtrDetalle("nCodConcepto") = 99
                                    ldtrDetalle("strDescripcionConcepto") = strMensaje02
                                    ldtrDetalle("nMontoConcepto") = 0
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle) 'Agrega la Fila al Data Table

                                    '-------------------------------------------------------------'
                                End If

                            End If
                        Case eCodigoConvenio.LetraCambio
                            'Verificamos que exista el recibo que se esta utilizando como llave de accesos y obtenemos el ID del estudiante de esa letra
                            If lfExisteRecibo(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXRecibo(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje16
                            End If

                            'Para que se pueda proceder se debe de pagar al menos el valor mínimo establecido en la 
                            'tabla de parámetros del sistema.
                            If lExitosa = True And nMontoPagado >= lfPagoMinimoLetra() Then
                                If lExitosa = True Then
                                    nEnlID = pfObtenerIDLetraCambio(strLlaveAcceso)
                                    'Verificamos que exista un comprobante para la letra por pagar, en la tabla de recibos
                                    If Not lfExisteNumComprobanteLetra(nEnlID, nNumComprobante) Then
                                        'Si el monto de pago es mayor al monto minimo admitido, se realizan los procesos de pago
                                        nDescuento = 0.0
                                        'Inserta recibo x pagar a tabla temporal en caso de reversion
                                        'If lfInsertarTemEnl(nEnlID) Then


                                        'Ingresa el Recio de PAGO en la Tabla REC_RECIBO y se obtiene el numero de transacción

                                        lstrNumeroTransaccion = pfGeneraReciboPago(strRecibo, strTipoLlave, nEstID, nNumComprobante, nMontoPagado, nCodConvenio, nDescuento, nMontoPagado, cCreo, nTipoEntidad, strNumeroTarjeta)


                                        If lstrNumeroTransaccion > 0 Then

                                            'Se crea el Abono a la Letra y se actualizan los valores de la letra de Cambio
                                            If pfCancelaLetradeCambio(strLlaveAcceso, nMontoPagado, lstrNumeroTransaccion) Then
                                                'nRetorno = lNumeroRecibo
                                                'Se inserta en la bitacora de pagos de bancos
                                                lfInsertarBitacoraPagosBancos(nEnlID, 0, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, cCreo)
                                                Dim nRecibo As String = pfObtenerRecibo(strLlaveAcceso)
                                                Dim nPlaID As Integer = fnPlanAcademico(nEstID)

                                                If lfInsertaInterfazEncabezado(nEstID, _
                                                            nRecibo, _
                                                            cValorAbono, _
                                                            nMontoPagado, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario), _
                                                            lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio), _
                                                            cCero, _
                                                            nMontoPagado, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            nMontoPagado, _
                                                            cCero, _
                                                            cCero, _
                                                            strLlaveAcceso, _
                                                            nMontoPagado, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            cCero, _
                                                            nMontoPagado, _
                                                            cCreo, _
                                                            cCero, _
                                                            "NR", _
                                                            cCero, _
                                                            cCero) Then
                                                    Dim bs_ine_id As Integer
                                                    Dim bs_ind_cod_centro_costo As String
                                                    Dim bs_ind_mto_detalle As Decimal
                                                    Dim bs_ind_ind_tipo_detalle As String
                                                    If nMontoPagado > 0 Then
                                                        bs_ind_ind_tipo_detalle = "A"
                                                        bs_ine_id = pfObtenerMaxIneID()
                                                        bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlaID)
                                                        bs_ind_mto_detalle = nMontoPagado
                                                        lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)

                                                    End If
                                                Else
                                                    lExitosa = False
                                                    nCodRespuesta = cCodMensaje99
                                                    strDescRespuesta = "No se logró crear el registro en el BEST"

                                                End If

                                            Else

                                                lExitosa = False
                                                nCodRespuesta = cCodMensaje99
                                                strDescRespuesta = "No se logró crear registro del Detalle de Abono"
                                            End If

                                        Else
                                            lExitosa = False
                                            nCodRespuesta = cCodMensaje99
                                            strDescRespuesta = "No se logró crear registro del Recibo de Dinero"
                                        End If

                                    Else
                                        lExitosa = False
                                        nCodRespuesta = cCodMensaje99
                                        strDescRespuesta = "No se logró crear registro del abono"
                                    End If
                                Else
                                    lExitosa = False
                                    nCodRespuesta = cCodMensaje51
                                    strDescRespuesta = "El Número de Comprobante ya existe para la letra de cambio solicitada."
                                End If
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje21
                            End If





                            If lExitosa = True Then

                                'Crea la fila del encabezado ****************************************************************
                                ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                ldtrFila("strCodRespuesta") = cCodMensaje00
                                ldtrFila("strDescripcionRespuesta") = strMensaje01
                                ldtrFila("nCantConceptosInformativos") = 1
                                ldtrFila("nCantConceptosPago") = 1
                                ldtrFila("strNumTransaccionPago") = lstrNumeroTransaccion
                                dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                '********************************************************************************************

                                'Crea el detalle referente a los conceptos Informativos
                                '***************************************************************************************
                                strDescripcionConceptp = "Pago/Abono L.C: " & strRecibo

                                ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                ldtrConceptosInformativos("nCodConcepto") = 1
                                ldtrConceptosInformativos("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                                ldtrConceptosInformativos("strValorConcepto") = nNumComprobante
                                dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                '*************************************************************************************

                                'Crea la fila del Detalle de la letra ****************************************************************
                                ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                ldtrDetalle("nConsecutivo") = 1
                                ldtrDetalle("nCodConcepto") = 1
                                ldtrDetalle("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                                ldtrDetalle("nMontoConcepto") = nMontoPagado
                                dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table

                            Else
                                ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                ldtrFila("strCodRespuesta") = nCodRespuesta
                                ldtrFila("strDescripcionRespuesta") = strDescRespuesta 'strMensaje21
                                ldtrFila("nCantConceptosInformativos") = 0
                                ldtrFila("nCantConceptosPago") = 0
                                ldtrFila("strNumTransaccionPago") = 0
                                dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                '********************************************************************************************

                                'Crea el detalle referente a los conceptos Informativos
                                '***************************************************************************************
                                ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                ldtrConceptosInformativos("nCodConcepto") = 0
                                ldtrConceptosInformativos("strDescripcionConcepto") = strMensaje99
                                ldtrConceptosInformativos("strValorConcepto") = 0
                                dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                '*************************************************************************************

                                'Crea la fila del Detalle  ****************************************************************
                                ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                ldtrDetalle("nConsecutivo") = 0
                                ldtrDetalle("nCodConcepto") = 0
                                ldtrDetalle("strDescripcionConcepto") = strMensaje99
                                ldtrDetalle("nMontoConcepto") = 0
                                dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table

                            End If


                        Case eCodigoConvenio.Carnets


                        Case eCodigoConvenio.CursosLibres


                        Case eCodigoConvenio.VidaEstudiantil


                    End Select

                Else

                End If
                'Si el servico no esta disponible invoca a la funcion publica para que cargue el dts con la información requerida por el BNCR.
            Else
                dtsPagoServicios = pfServicioDisponible()
            End If


            Return dtsPagoServicios
        Catch ex As Exception
            Dim lError As Boolean
            lError = lfBitacoraErrores("WS-CONECTIVIDAD", "pfEjecutarPagoServicio", Err.Number, Err.Description)
            Throw ex
        End Try
        Return ldtsDatos
    End Function

    'Funcion de Verificación del Pago de los servicios
    <WebMethod()> _
    Public Function pfVerificarPagoServicio(ByVal strEntidad As String, _
                                            ByVal strClave As String, _
                                            ByVal nNumeroFactura As Integer, _
                                            ByVal strRecibo As String, _
                                            ByVal strTipoLlave As String, _
                                            ByVal strLlaveAcceso As String, _
                                            ByVal strNumTransaccionPago As String) As DataSet

        Dim lvalor As Integer = 0
        Dim dtsVerificaPagoServicios As New Data.DataSet("dtsVerificaPagoServicios")
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim ldtrFila As DataRow
        Dim strSql As String = String.Empty
        Dim strMensaje As String = String.Empty
        Dim strCodigoMensaje As String = String.Empty


        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
        dtsVerificaPagoServicios.Tables.Add(dtbEncabezado)

        Try
            'Si el servicio está disponible opera normalemnte
            If lfServicioDisponible() = True Then
                '****************************************************************************
                If strNumTransaccionPago <> "" Then

                    strSql = "Select count(ID) " & _
                            "From REC_RECIBO " & _
                            "Where Recibo=" & strNumTransaccionPago

                    If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                    Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
                    cnnConexion.Open()
                    lvalor = cmdObtener.ExecuteScalar

                Else

                    If nNumeroFactura > 0 Then
                        strSql = "Select count(ID) " & _
                            " From ENM_ENCABEZADO_MATRICULA " & _
                            " Where FACTURA=" & nNumeroFactura & _
                            " And Cancelada=1"

                        If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                        Dim cmdConsultaFactura As New SqlCommand(strSql, cnnConexion)
                        cnnConexion.Open()
                        lvalor = cmdConsultaFactura.ExecuteScalar


                    Else
                        'Verifica que exista la letra y que tenga abonos
                        Dim lIDLetraCambio As Integer = pfObtenerIDLetraCambio(strRecibo)

                        If lIDLetraCambio > 0 Then
                            strSql = "Select count(ID) " & _
                            "From ALE_ABONO_LETRA " & _
                            "Where ENL_ID=" & lIDLetraCambio
                            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                            Dim cmdConsultaAbono As New SqlCommand(strSql, cnnConexion)
                            cnnConexion.Open()
                            lvalor = cmdConsultaAbono.ExecuteScalar
                        End If


                    End If


                End If

                If lvalor = 0 Then
                    strMensaje = strMensaje02
                    strCodigoMensaje = cCodMensaje51

                Else
                    strMensaje = strMensaje01
                    strCodigoMensaje = cCodMensaje00
                End If

                'Crea la fila del encabezado ****************************************************************
                ldtrFila = dtsVerificaPagoServicios.Tables("Encabezado").NewRow
                ldtrFila("strCodRespuesta") = strCodigoMensaje
                ldtrFila("strDescripcionRespuesta") = strMensaje
                dtsVerificaPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                '********************************************************************************************



                '****************************************************************************
                'Si el servicio no esta disponible retorna el DTS de la consulta publica
            Else
                dtsVerificaPagoServicios = pfServicioDisponible()
            End If




        Catch ex As Exception

        End Try
        Return dtsVerificaPagoServicios
    End Function

    'Función que verifica la disponibilidad del Servicio
    <WebMethod()> _
    Public Function pfServicioDisponible() As DataSet
        Dim strSql As String = String.Empty

        Dim dtsServicioDisponible As New DataSet
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim strCodigoRespuesta As String = String.Empty
        Dim ldtrFila As DataRow

        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtsServicioDisponible.Tables.Add(dtbEncabezado)
        '-------------------------------------------------------------------------------------


        Try
            If lfServicioDisponible() = True Then
                strCodigoRespuesta = "00"
            Else
                strCodigoRespuesta = "50"
            End If

            'Crea la fila del encabezado ****************************************************************
            ldtrFila = dtsServicioDisponible.Tables("Encabezado").NewRow
            ldtrFila("strCodRespuesta") = strCodigoRespuesta
            dtsServicioDisponible.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
            '********************************************************************************************
            cnnConexion.Close()


        Catch ex As Exception
            strCodigoRespuesta = "50"
            'Crea la fila del encabezado ****************************************************************
            ldtrFila = dtsServicioDisponible.Tables("Encabezado").NewRow
            ldtrFila("strCodRespuesta") = strCodigoRespuesta
            dtsServicioDisponible.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
            '********************************************************************************************
            Throw ex
        End Try
        Return dtsServicioDisponible
    End Function

    'Funcion que reversa el pago de un recibo
    <WebMethod()> _
    Public Function pfReversarPagoServicio(ByVal strEntidad As String, _
                                      ByVal strClave As String, _
                                      ByVal strTipoLlave As String, _
                                      ByVal strLlaveAcceso As String, _
                                      ByVal nCodAgencia As Integer, _
                                      ByVal nCodConvenio As Integer, _
                                      ByVal nMontoRecibo As Decimal, _
                                      ByVal nNumComprobante As Integer, _
                                      ByVal nNumFactura As Integer, _
                                      ByVal strRecibo As String, _
                                      ByVal nNumNotaCredito As Decimal, _
                                      ByVal strNumTransaccionPago As String) As DataSet

        Try
            Dim ldtsDatos As New DataSet
            Dim dtsReversaPagos As New DataSet("dtsReversaPagos")
            Dim dtbEncabezado As DataTable = New DataTable("Encabezado")
            Dim ldtrFila As DataRow
            Dim ldFechaActual As New Date
            Dim ldFechaRecibo As New Date
            ldFechaActual = FormatDateTime(Now, DateFormat.ShortDate)
            ' lstrCadenas = lfObtenerFechaCreoRecibo(nNumFactura)
            ldFechaRecibo = FormatDateTime(Now, DateFormat.ShortDate)
            Dim lExitosa As Boolean = True
            Dim lnumRecibo As Integer
            Dim lEstID As New Integer
            Dim lPerID As New Integer
            Dim strCodRespuestas As String = ""
            Dim strDescripcionRespuesta As String = ""
            Dim nIdNumLetra As New Integer
            Dim nIdFactura As Integer = 0
            Dim dstPleID As DataSet
            Dim nPleId As Integer
            Dim dtsDetalleAbono As DataSet
            Dim nEstID As Integer
            Dim EnlID As Integer = 0
            Dim nRecibo As Integer = 0
            Dim nFactura As Integer
            Dim lValoresFacturacion(3) As Decimal
            Dim strCreo As String = "BN-Conectividad"

            'TABLA DEL ENCABEZADO
            '-------------------------------------------------------------------------------------
            dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
            dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
            dtsReversaPagos.Tables.Add(dtbEncabezado)
            '-------------------------------------------------------------------------------------

            'Verifica que el servicio este disponible
            If lfServicioDisponible() = True Then
                lExitosa = True
            Else
                strCodRespuestas = cCodMensaje50
                strDescripcionRespuesta = strMensaje12
                lExitosa = False
            End If

            'Valida el usuario y la contraseña de ingreso
            If lExitosa = True Then
                If lfValidaUsuario(strEntidad, strClave) = True Then
                    lExitosa = True
                Else
                    strCodRespuestas = cCodMensaje51
                    strDescripcionRespuesta = strMensaje20
                    lExitosa = False
                End If
            End If

            'Verifica que exista el convenio
            If lExitosa Then
                If lfExisteConvenio(nCodConvenio) Then
                    lExitosa = True
                Else
                    strCodRespuestas = cCodMensaje51
                    strDescripcionRespuesta = strMensaje05
                    lExitosa = False
                End If
            End If

            If lExitosa Then
                'Se obtiene el ID del Periodo para los procesos siguientes
                lPerID = pfObtenerPeriodoActivo()
                'If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                'cnnConexion.Open()

                Select Case nCodConvenio
                    Case eCodigoConvenio.Facturas

                        nFactura = CInt(strLlaveAcceso)
                        'Verifica que la llave de acceso exista y obtiene el ID del estudiante de la factura a reversar
                        If lfObtieneIDEstudianteXFactura(strLlaveAcceso) > 0 Then
                            lExitosa = True
                            nEstID = lfObtieneIDEstudianteXFactura(strLlaveAcceso)
                        Else
                            lExitosa = False
                            strCodRespuestas = cCodMensaje51
                            strDescripcionRespuesta = strMensaje16

                        End If

                        If lExitosa = True Then

                            'Obtiene los valor de facturacion
                            'lValoresFacturacion = pfDetalleFacturacion(nEstID)

                            'Si la factura es de tipo 0 ingresa al proceso de reversión de pago
                            If lfObtieneTipoFactura(strLlaveAcceso) = 0 Then
                                'Obtiene el ID de la factura por cancelar
                                nIdFactura = pfObtenerIDFactura(strLlaveAcceso)
                                'Si la factura existe continua el proceso de reversion
                                If nIdFactura > 0 Then
                                    'Actualiza el la tabla ENM_ENCABEZADO_MATRICULA
                                    If (lfActualizaENM(nIdFactura, nEstID, strCreo) = "1") Then
                                        'Obtiene los ID de PLE_PLAN_X_ESTUDIANTE que estan relacionados con el encabezado de matricula por reversar
                                        dstPleID = lfObtienePleID(nIdFactura)
                                        If dstPleID.Tables(0).Rows.Count > 0 Then
                                            For i As Integer = 0 To dstPleID.Tables(0).Rows.Count - 1
                                                nPleId = dstPleID.Tables(0).Rows(i)("PLE_ID").ToString()
                                                If lfActualizaPLe(nPleId, nEstID) = 1 Then
                                                    lExitosa = True
                                                Else
                                                    strCodRespuestas = cCodMensaje51
                                                    strDescripcionRespuesta = "Tipo de factura inválida para reversión"
                                                End If
                                            Next
                                        Else
                                            lExitosa = False
                                            strCodRespuestas = cCodMensaje51
                                            strDescripcionRespuesta = "Tipo de factura inválida para reversión"
                                        End If

                                        If lExitosa Then
                                            If fnVerificaPagoMatriculaXDocumento(nEstID, nIdFactura) > 0 Then
                                                If lfEliminaPAM(nIdFactura, nEstID) = 0 Then
                                                    lExitosa = False
                                                End If
                                            End If
                                        End If

                                        If lExitosa Then
                                            If pfExistePagoSeguro(nEstID, nIdFactura) > 0 Then
                                                If pfEliminaPagoSeguroEstudiante(nEstID, nIdFactura, lPerID) = True Then
                                                        lExitosa = True
                                                Else
                                                    lExitosa = False
                                                End If
                                            End If

                                        End If
                                    End If
                                Else

                                    strCodRespuestas = cCodMensaje51
                                    strDescripcionRespuesta = strMensaje15
                                    lExitosa = False
                                End If
                            Else
                                strCodRespuestas = cCodMensaje51
                                strDescripcionRespuesta = strMensaje15
                                lExitosa = False
                            End If

                        Else
                            strCodRespuestas = cCodMensaje51
                            strDescripcionRespuesta = "Tipo de factura inválida para reversión"
                            lExitosa = False
                        End If


                    Case eCodigoConvenio.LetraCambio

                        EnlID = pfObtenerIDLetraCambio(strLlaveAcceso)
                        nRecibo = lfObtieneRecibo(EnlID, nNumComprobante)

                        If lfExisteRecibo(EnlID, nNumComprobante) = 1 Then
                            'Valida Llave de Acceso y obtiene ID de estudiante
                            If lfObtieneIDEstudianteXRecibo(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXRecibo(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                strCodRespuestas = cCodMensaje51
                                strDescripcionRespuesta = "LLave de Acceso Inválida"
                            End If


                            'Valida llave de Acceso y obtiene ID del Numero de letra de cambio
                            If EnlID > 0 Then
                                lExitosa = True
                            Else
                                lExitosa = False
                                strCodRespuestas = cCodMensaje51
                                strDescripcionRespuesta = "L.C inexistente"
                            End If

                            'Verifica Existencia de NumComprobrante para letra de cambio en un recibo
                            If lExitosa Then
                                If lfExisteNumComprobanteXLetra(EnlID, nNumComprobante) > 0 Then
                                    lExitosa = True
                                Else
                                    lExitosa = False
                                    strCodRespuestas = cCodMensaje51
                                    strDescripcionRespuesta = "Num.Comprobante inexistente"
                                End If
                            End If

                            If lExitosa Then
                                If lfExisteNumComprobanteXLetraXMonto(EnlID, nNumComprobante, nMontoRecibo) > 0 Then
                                    lExitosa = True
                                Else
                                    lExitosa = False
                                    strCodRespuestas = cCodMensaje51
                                    strDescripcionRespuesta = "Monto de Recibo es inválido"
                                End If
                            End If

                            If lExitosa = True Then

                                dtsDetalleAbono = lfObtieneAbonoAnterior(EnlID)

                                If lfActualizaAle(EnlID, nNumComprobante) Then
                                    lExitosa = True
                                Else
                                    strCodRespuestas = cCodMensaje51
                                    strDescripcionRespuesta = strMensaje99
                                    lExitosa = False
                                End If

                                If lExitosa Then
                                    If lfActualizaRecibo(EnlID, nNumComprobante, strCreo) Then
                                        lExitosa = True
                                    Else
                                        strCodRespuestas = cCodMensaje51
                                        strDescripcionRespuesta = strMensaje99
                                        lExitosa = False
                                    End If
                                End If

                                If lExitosa Then
                                    If lfActualizaENL(nMontoRecibo, EnlID, nNumComprobante, lnumRecibo) = "1" Then
                                        lExitosa = True
                                    Else
                                        strCodRespuestas = cCodMensaje51
                                        strDescripcionRespuesta = strMensaje99
                                        lExitosa = False
                                    End If
                                End If

                                If lExitosa Then
                                    If lfInsertarBitacoraBRA(EnlID, dtsDetalleAbono) = "1" Then
                                        lExitosa = True
                                    Else
                                        lExitosa = False
                                        strCodRespuestas = cCodMensaje51
                                        strDescripcionRespuesta = strMensaje99
                                    End If

                                End If

                            Else
                                strCodRespuestas = cCodMensaje51
                                strDescripcionRespuesta = strMensaje99
                                lExitosa = False

                            End If
                        Else
                            lExitosa = False
                            strCodRespuestas = cCodMensaje51
                            strDescripcionRespuesta = "El recibo a reversar no existe."
                        End If

                    Case Else
                        strCodRespuestas = cCodMensaje51
                        strDescripcionRespuesta = strMensaje05
                End Select
            End If

            If lExitosa Then
                If lfInsertaBitacoraReversiones(nIdFactura, EnlID, nCodConvenio, nNumComprobante, nRecibo) = "1" Then
                    lExitosa = True
                Else
                    lExitosa = False
                    strCodRespuestas = cCodMensaje51
                    strDescripcionRespuesta = "La reversión no se realizó correctamente"
                End If
            End If

            If lExitosa Then
                grefTransaction = cnnConexion.BeginTransaction
                grefTransaction.Commit()
                'Crea la fila del encabezado **************************************
                ldtrFila = dtsReversaPagos.Tables("Encabezado").NewRow
                ldtrFila("strCodRespuesta") = cCodMensaje00
                ldtrFila("strDescripcionRespuesta") = strMensaje19
                dtsReversaPagos.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                '******************************************************************
            Else
                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                cnnConexion.Open()
                grefTransaction = cnnConexion.BeginTransaction
                grefTransaction.Rollback()
                ldtrFila = dtsReversaPagos.Tables("Encabezado").NewRow
                ldtrFila("strCodRespuesta") = strCodRespuestas
                ldtrFila("strDescripcionRespuesta") = strDescripcionRespuesta
                dtsReversaPagos.Tables("Encabezado").Rows.Add(ldtrFila)
            End If

            Return dtsReversaPagos
        Catch ex As Exception
            Throw ex

        Finally
            Me.cnnConexion.Close()

            grefTransaction = Nothing
        End Try

    End Function

    'Funcion que ejecuta el pago de servicios
    <WebMethod()> _
    Public Function pfEjecutarPagoServicioCredomatic(ByVal strUsuario As String, _
                                           ByVal strClaveUsuario As String, _
                                           ByVal strTipoLlave As String, _
                                           ByVal strLlaveAcceso As String, _
                                           ByVal nCodAgencia As Byte, _
                                           ByVal nCodConvenio As Byte, _
                                           ByVal nTipoTransaccion As Byte, _
                                           ByVal strPeriodo As String, _
                                           ByVal nMontoPagado As Decimal, _
                                           ByVal nMontoTotalRecibo As Decimal, _
                                           ByVal nNumFactura As Integer, _
                                           ByVal strRecibo As String, _
                                           ByVal nSelfVerificacion As Byte, _
                                           ByVal strCodMoneda As String, _
                                           ByVal nNumComprobante As Integer, _
                                           ByVal nNumeroCuotas As Integer, _
                                           ByVal nTipoPago As Byte, _
                                           ByVal nCodBanco As Integer, _
                                           ByVal nNumCheque As Long, _
                                           ByVal nNumCuenta As Long, _
                                           ByVal nMontoTotal As Decimal, _
                                           ByVal nNumNotaCredito As Integer, _
                                           ByVal nMontoCurso As Integer, _
                                           ByVal nMontoMatricula As Integer, _
                                           ByVal nMontoSeguro As Integer, _
                                           ByVal nMontoDescuento As Decimal, _
                                           ByVal strCreo As String, _
                                           ByVal strNumeroTarjeta As String, _
                                           ByVal pDescuento As Decimal, _
                                           ByRef pTipoCompra As Decimal) As DataSet
        Dim ldtsDatos As New DataSet
        'Procedimiento para crear las tablas del DataSet de retorno del la CLase Pago de Servicios
        'Creamos el data set que contendrá la información del Afiliado para el archivo de Envío
        Dim dtsPagoServicios As New Data.DataSet("dtsPagoServicios")
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim dtbDetalle As Data.DataTable = New DataTable("Detalle")
        Dim dtbSubDetalle As Data.DataTable = New DataTable("SubDetalle")
        Dim ldtrFila As DataRow
        Dim ldtrConceptosInformativos As DataRow
        Dim ldtrDetalle As DataRow
        Dim lstrNumeroTransaccion As String = String.Empty
        Dim lValoresFacturacion(5) As Decimal
        Dim nTotalConceptosPago As Integer = 0
        Dim strDescripcionConceptp As String
        Dim nEstID As Integer
        Dim lExitosa As Boolean
        Dim nEnlID As Integer = 0
        Dim nCodRespuesta As Integer = 0
        Dim strDescRespuesta As String = ""
        Dim nDescuento As Decimal = 0.0
        Dim nTotalSinDescuento As Decimal = 0.0
        Dim nTipoEntidad = 2

        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nCantConceptosInformativos", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("nCantConceptosPago", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("strNumTransaccionPago", Type.GetType("System.String"))
        dtsPagoServicios.Tables.Add(dtbEncabezado)
        '-------------------------------------------------------------------------------------

        'Tabla del SubDetalle correspondiente a la clase conceptos Informativos
        '--------------------------------------------------------------------------------------
        dtbSubDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbSubDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbSubDetalle.Columns.Add("strValorConcepto", Type.GetType("System.String"))
        dtsPagoServicios.Tables.Add(dtbSubDetalle)
        '--------------------------------------------------------------------------------------

        'TABLA DEL DETALLE correspondiente a la clase de Conceptos de Pago
        '-------------------------------------------------------------------------------------
        dtbDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Int32"))
        dtbDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nMontoConcepto", Type.GetType("System.Decimal"))
        dtsPagoServicios.Tables.Add(dtbDetalle)
        '-------------------------------------------------------------------------------------
        Try
            'Si el servicio está diponible ejecuta todo el proceso; si no devuelve el encabezado del DTS 
            'con el código de respuesta 50

            If lfServicioDisponible() = True Then

                'Valida que el usuario este registrado en el sistema
                If lfValidaUsuario(strUsuario, strClaveUsuario) = True Then
                    'Segun el tipo de convenio realiza el pago
                    Select Case nCodConvenio

                        Case eCodigoConvenio.Facturas
                            'Carga los detelles de la factura antes de aplicar los cambios
                            If lfExisteFactura(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXFactura(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje13
                            End If

                            If lExitosa = True Then

                                If pTipoCompra = 1 Then
                                    If nMontoCurso > 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfActualizaMateriasMatriculadas(nEstID, nTipoEntidad)
                                    End If
                                    If nMontoMatricula > 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfIngresaPagoMatriculaEstudiante(nEstID, nNumFactura, nMontoMatricula)
                                    End If

                                    If nMontoSeguro > 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfIngresaPagoSeguroEstudiante(nEstID, nNumFactura)
                                    End If
                                    If lExitosa Then
                                        If nTotalConceptosPago > 0 Then
                                            'Se actualiza el estado a cancelada por BN-CONECTIVDAD
                                            If pfActualizaEstadoFactura(nNumFactura, _
                                                                        nMontoDescuento, _
                                                                        nMontoPagado, _
                                                                        nMontoTotal, _
                                                                        strCreo, _
                                                                        strNumeroTarjeta, nTipoEntidad, pDescuento, pTipoCompra) Then
                                                lstrNumeroTransaccion = strLlaveAcceso
                                                If lstrNumeroTransaccion > 0 Then
                                                    'Se inserta el pago a la bitacora de pagos de bancos
                                                    lfInsertarBitacoraPagosBancos(0, nNumFactura, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, cCreoCredomatic)
                                                    If lfInsertaInterfazEncabezado(nEstID, _
                                                        strLlaveAcceso, _
                                                        cValorMatricula, _
                                                        cCero, _
                                                        nMontoMatricula, _
                                                        nMontoSeguro, _
                                                        nMontoCurso, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        nMontoTotal, _
                                                        nMontoDescuento, _
                                                        pDescuento, _
                                                        cCero, _
                                                        cCero, _
                                                        nMontoDescuento, _
                                                        nMontoTotalRecibo, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        nMontoTotalRecibo, _
                                                        nNumComprobante, _
                                                        strNumeroTarjeta, _
                                                        nMontoTotalRecibo, _
                                                        cCreoCredomatic, _
                                                        cCero, _
                                                        "NR", _
                                                        cCero, _
                                                        cCero) = True Then
                                                        Dim dtsDetalle As New DataSet
                                                        Dim bs_ine_id As Integer
                                                        Dim bs_ind_cod_centro_costo As String
                                                        Dim bs_ind_mto_detalle As Decimal
                                                        Dim bs_ind_ind_tipo_detalle As String
                                                        Dim nPlanID As Integer
                                                        dtsDetalle = pfObtenerDetallesCursosFactura(strLlaveAcceso)
                                                        If nMontoCurso > 0 Then
                                                            bs_ind_ind_tipo_detalle = "C"
                                                            If dtsDetalle.Tables(0).Rows.Count > 0 Then
                                                                For i = 0 To dtsDetalle.Tables(0).Rows.Count - 1
                                                                    bs_ine_id = dtsDetalle.Tables(0).Rows(i)("bs_ine_id")
                                                                    bs_ind_cod_centro_costo = dtsDetalle.Tables(0).Rows(i)("CENTRO_COSTO")
                                                                    bs_ind_mto_detalle = dtsDetalle.Tables(0).Rows(i)("COSTO")
                                                                    lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                                Next
                                                            End If
                                                        End If
                                                        If nMontoMatricula > 0 Then
                                                            nPlanID = fnPlanAcademico(nEstID)
                                                            bs_ind_ind_tipo_detalle = "M"
                                                            bs_ine_id = dtsDetalle.Tables(0).Rows(0)("bs_ine_id")
                                                            bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlanID)
                                                            bs_ind_mto_detalle = nMontoMatricula
                                                            lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                        End If
                                                        If nMontoSeguro > 0 Then
                                                            nPlanID = fnPlanAcademico(nEstID)
                                                            bs_ind_ind_tipo_detalle = "S"
                                                            bs_ine_id = dtsDetalle.Tables(0).Rows(0)("bs_ine_id")
                                                            bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlanID)
                                                            bs_ind_mto_detalle = nMontoSeguro
                                                            lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                        End If
                                                    End If
                                                Else
                                                    lExitosa = False
                                                    nCodRespuesta = cCodMensaje51
                                                    strDescRespuesta = "No se creó recibo."
                                                End If
                                            Else
                                                lExitosa = False
                                                nCodRespuesta = cCodMensaje51
                                                strDescRespuesta = "No se logró actualizar la factura solicitada"
                                            End If
                                        Else
                                            lExitosa = False
                                            nCodRespuesta = cCodMensaje51
                                            strDescRespuesta = "No hay Conceptos de pago pendientes para la factura solicitada."
                                        End If
                                    End If

                                Else
                                    If pTipoCompra = 2 Then

                                        If nMontoCurso > 0 Then
                                            nTotalConceptosPago = nTotalConceptosPago + 1
                                            pfActualizaMateriasMatriculadas(nEstID, nTipoEntidad)
                                        End If

                                        If lExitosa Then
                                            If nTotalConceptosPago > 0 Then
                                                'Se actualiza el estado a cancelada por BN-CONECTIVDAD
                                                If pfActualizaEstadoFactura(nNumFactura, nMontoDescuento, nMontoPagado, nMontoTotal, strCreo, strNumeroTarjeta, nTipoEntidad, pDescuento, pTipoCompra) Then

                                                    lstrNumeroTransaccion = strLlaveAcceso
                                                    If lstrNumeroTransaccion > 0 Then
                                                        'Se inserta el pago a la bitacora de pagos de bancos
                                                        lfInsertarBitacoraPagosBancos(0, nNumFactura, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, cCreoCredomatic)
                                                        If lfInsertaInterfazEncabezado(nEstID, _
                                                   strLlaveAcceso, _
                                                   cValorMatricula, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoCurso, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoTotalRecibo, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoTotal, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoTotalRecibo, _
                                                   nNumComprobante, _
                                                   strNumeroTarjeta, _
                                                   nMontoTotalRecibo, _
                                                   cCreoCredomatic, _
                                                   cCero, _
                                                   "NR", _
                                                   cCero, _
                                                   cCero) Then
                                                            Dim dtsDetalle As DataSet
                                                            Dim bs_ine_id As Integer
                                                            Dim bs_ind_cod_centro_costo As String
                                                            Dim bs_ind_mto_detalle As Decimal
                                                            Dim bs_ind_ind_tipo_detalle As String
                                                            dtsDetalle = pfObtenerDetallesCursosFactura(strLlaveAcceso)
                                                            If nMontoCurso > 0 Then
                                                                bs_ind_ind_tipo_detalle = "C"
                                                                If dtsDetalle.Tables(0).Rows.Count > 0 Then
                                                                    For i = 0 To dtsDetalle.Tables(0).Rows.Count - 1
                                                                        bs_ine_id = dtsDetalle.Tables(0).Rows(i)("bs_ine_id")
                                                                        bs_ind_cod_centro_costo = dtsDetalle.Tables(0).Rows(i)("CENTRO_COSTO")
                                                                        bs_ind_mto_detalle = dtsDetalle.Tables(0).Rows(i)("COSTO")
                                                                        lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                                    Next
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        lExitosa = False
                                                        nCodRespuesta = cCodMensaje51
                                                        strDescRespuesta = "No se creó recibo."
                                                    End If
                                                Else
                                                    lExitosa = False
                                                    nCodRespuesta = cCodMensaje51
                                                    strDescRespuesta = "No se logró actualizar la factura solicitada"
                                                End If
                                            Else
                                                lExitosa = False
                                                nCodRespuesta = cCodMensaje51
                                                strDescRespuesta = "No hay Conceptos de pago pendientes para la factura solicitada."
                                            End If
                                        End If

                                    End If

                                End If
                            End If
                            'Else

                            'End If

                            If lExitosa Then
                                'crea la información para el Dataset de retorno
                                'Crea la fila del encabezado ****************************************************************
                                ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                ldtrFila("strCodRespuesta") = cCodMensaje00
                                ldtrFila("strDescripcionRespuesta") = strMensaje01
                                ldtrFila("nCantConceptosInformativos") = 1
                                ldtrFila("nCantConceptosPago") = nTotalConceptosPago
                                ldtrFila("strNumTransaccionPago") = lstrNumeroTransaccion
                                dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                '********************************************************************************************
                                'Crea la fila del Detalle de la Factura ****************************************************************

                                'Si existen valores en el pago de materias con descuento los agrega al table de response
                                If nMontoCurso > 0 Then
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 1
                                    ldtrDetalle("nCodConcepto") = 1
                                    ldtrDetalle("strDescripcionConcepto") = "Materias Matriculadas "
                                    ldtrDetalle("nMontoConcepto") = nMontoCurso
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                End If
                                'Si existen valores en el pago de matricula los agrega al table de response
                                If nMontoMatricula > 0 Then
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 2
                                    ldtrDetalle("nCodConcepto") = 2
                                    ldtrDetalle("strDescripcionConcepto") = "Monto Matricula"
                                    ldtrDetalle("nMontoConcepto") = nMontoMatricula
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                End If
                                'Si existen valores para el concepto de pago de seguro los agrega al table de response
                                If nMontoSeguro > 0 Then
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 3
                                    ldtrDetalle("nCodConcepto") = 3
                                    ldtrDetalle("strDescripcionConcepto") = "Seguro Estudiantil "
                                    ldtrDetalle("nMontoConcepto") = nMontoSeguro
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                End If
                                'Crea la fila del Detalle de la Factura para conceptos informativos ****************************************************************
                                'Crea el detalle referente a los conceptos Informativos
                                '***************************************************************************************
                                ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                ldtrConceptosInformativos("nCodConcepto") = 1
                                ldtrConceptosInformativos("strDescripcionConcepto") = "Fact. Matricula"
                                ldtrConceptosInformativos("strValorConcepto") = nNumFactura
                                dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)

                            Else
                                ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                ldtrFila("strCodRespuesta") = nCodRespuesta
                                ldtrFila("strDescripcionRespuesta") = strDescRespuesta
                                ldtrFila("nCantConceptosInformativos") = 0
                                ldtrFila("nCantConceptosPago") = 1
                                ldtrFila("strNumTransaccionPago") = 0
                                dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                '********************************************************************************************

                                'Crea el detalle referente a los conceptos Informativos
                                '***************************************************************************************
                                ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                ldtrConceptosInformativos("nCodConcepto") = cCodMensaje99
                                ldtrConceptosInformativos("strDescripcionConcepto") = strMensaje02
                                ldtrConceptosInformativos("strValorConcepto") = 0
                                dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                '*************************************************************************************

                                'Crea la fila del Detalle de la Factura ****************************************************************
                                ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                ldtrDetalle("nConsecutivo") = 0
                                ldtrDetalle("nCodConcepto") = 99
                                ldtrDetalle("strDescripcionConcepto") = strMensaje02
                                ldtrDetalle("nMontoConcepto") = 0
                                dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle) 'Agrega la Fila al Data Table

                                '-------------------------------------------------------------'
                            End If
                        Case eCodigoConvenio.LetraCambio
                            'Verificamos que exista el recibo que se esta utilizando como llave de accesos y obtenemos el ID del estudiante de esa letra

                            If lfExisteRecibo(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXRecibo(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje16
                            End If

                            If lExitosa = True Then
                                nEnlID = pfObtenerIDLetraCambio(strLlaveAcceso)
                                'FUNCION QUE VALIDA SI EXISTE UNA LETRA DE CAMBIO PARA REALIZAR EL PAGO.




                                'Verificamos que exista un comprobante para la letra por pagar
                                'If Not lfExisteNumComprobanteLetra(nEnlID, nNumComprobante) Then
                                'Si el monto de pago es mayor al monto minimo admitido, se realizan los procesos de pago
                                ' If nMontoPagado > lfPagoMinimoLetra() Then
                                nDescuento = 0.0
                                lstrNumeroTransaccion = pfGeneraReciboPago(strRecibo, strTipoLlave, nEstID, nNumComprobante, nMontoPagado, nCodConvenio, nDescuento, nMontoPagado, strCreo, nTipoEntidad, strNumeroTarjeta)
                                'Verificamos que se haya creado un recibo de manera exitosa
                                If lstrNumeroTransaccion > 0 Then
                                    lExitosa = True
                                    'Se inserta en la bitacora de pagos de bancos
                                    If lfInsertarBitacoraPagosBancos(nEnlID, 0, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, strCreo) = 1 Then
                                        Dim nRecibo As String = lstrNumeroTransaccion
                                        Dim nPlaID As Integer = fnPlanAcademico(nEstID)
                                        If lfInsertaInterfazEncabezado(nEstID, _
                                                    nRecibo, _
                                                    cValorAbono, _
                                                    nMontoPagado, _
                                                    cCero, _
                                                    cCero, _
                                                   cCero, _
                                                  lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario), _
                                                    lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio), _
                                                    cCero, _
                                                    nMontoPagado, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    nMontoPagado, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    nMontoPagado, _
                                                    nNumComprobante, _
                                                    strNumeroTarjeta, _
                                                    nMontoPagado, _
                                                    cCreoCredomatic, _
                                                    cCero, _
                                                    "NR", _
                                                    cCero, _
                                                    cCero) Then
                                            Dim bs_ine_id As Integer
                                            Dim bs_ind_cod_centro_costo As String
                                            Dim bs_ind_mto_detalle As Decimal
                                            Dim bs_ind_ind_tipo_detalle As String
                                            If nMontoPagado > 0 Then
                                                bs_ind_ind_tipo_detalle = "A"
                                                bs_ine_id = pfObtenerMaxIneID()
                                                bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlaID)
                                                bs_ind_mto_detalle = nMontoPagado
                                                lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)

                                            End If
                                        End If
                                    End If
                                Else
                                    lExitosa = False
                                    nCodRespuesta = cCodMensaje99
                                    strDescRespuesta = strMensaje99
                                End If

                                'Else
                                '    lExitosa = False
                                '    nCodRespuesta = cCodMensaje99
                                '    strDescRespuesta = strMensaje21
                                'End If
                                'Else
                                '    lExitosa = False
                                '    nCodRespuesta = cCodMensaje51
                                '    strDescRespuesta = "El Número de Comprobante ya existe para la de cambio solicitada."
                                'End If

                                If lExitosa = True Then

                                    'Crea la fila del encabezado ****************************************************************
                                    ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                    ldtrFila("strCodRespuesta") = cCodMensaje00
                                    ldtrFila("strDescripcionRespuesta") = strMensaje01
                                    ldtrFila("nCantConceptosInformativos") = 1
                                    ldtrFila("nCantConceptosPago") = 1
                                    ldtrFila("strNumTransaccionPago") = lstrNumeroTransaccion
                                    dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                    '********************************************************************************************

                                    'Crea el detalle referente a los conceptos Informativos
                                    '***************************************************************************************
                                    strDescripcionConceptp = "Pago/Abono L.C: " & strRecibo

                                    ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                    ldtrConceptosInformativos("nCodConcepto") = 1
                                    ldtrConceptosInformativos("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                                    ldtrConceptosInformativos("strValorConcepto") = nNumComprobante
                                    dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                    '*************************************************************************************

                                    'Crea la fila del Detalle de la letra ****************************************************************
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 1
                                    ldtrDetalle("nCodConcepto") = 1
                                    ldtrDetalle("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                                    ldtrDetalle("nMontoConcepto") = nMontoPagado
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table

                                Else
                                    ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                    ldtrFila("strCodRespuesta") = nCodRespuesta
                                    ldtrFila("strDescripcionRespuesta") = strMensaje21
                                    ldtrFila("nCantConceptosInformativos") = 0
                                    ldtrFila("nCantConceptosPago") = 0
                                    ldtrFila("strNumTransaccionPago") = 0
                                    dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                    '********************************************************************************************

                                    'Crea el detalle referente a los conceptos Informativos
                                    '***************************************************************************************
                                    ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                    ldtrConceptosInformativos("nCodConcepto") = 0
                                    ldtrConceptosInformativos("strDescripcionConcepto") = strMensaje99
                                    ldtrConceptosInformativos("strValorConcepto") = 0
                                    dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                    '*************************************************************************************

                                    'Crea la fila del Detalle  ****************************************************************
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 0
                                    ldtrDetalle("nCodConcepto") = 0
                                    ldtrDetalle("strDescripcionConcepto") = strMensaje99
                                    ldtrDetalle("nMontoConcepto") = 0
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table

                                End If
                            End If

                        Case eCodigoConvenio.Carnets


                        Case eCodigoConvenio.CursosLibres


                        Case eCodigoConvenio.VidaEstudiantil


                    End Select

                Else

                End If

                'Si el servico no esta disponible invoca a la funcion publica para que cargue el dts con la información requerida por el BNCR.
            Else
                dtsPagoServicios = pfServicioDisponible()
            End If


            Return dtsPagoServicios
        Catch ex As Exception
            Throw ex
        End Try
        Return ldtsDatos
    End Function

#End Region

#Region "Funciones Privadas"


#Region "Consultas Generales"

    Private Function lfBitacoraErrores(pSistema As String, pProcedimiento As String, pNumeroError As String, pError As String) As Boolean
        Dim strSql As String
        Try
            pError = pError.Replace("'", " ")
            strSql = " INSERT INTO [dbo].[BIE_BITACORA_ERRORES] " & _
                                "( [PANTALLA]      " & _
                                ", [PROCEDIMIENTO] " & _
                                ", [NUMERROR]      " & _
                                ", [ERROR]         " & _
                                ", [FECHAHORA]     " & _
                                ", [USUARIO] )     " & _
                        " VALUES (" & _
                               "  ' " & pSistema & "'" & _
                               " ,' " & pProcedimiento & "'" & _
                               " ,' " & pNumeroError & "'" & _
                               " ,' " & pError & "'" & _
                               " , GETDATE() " & _
                               " ,'IBANKING')"

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function


    Private Function lfObtieneTipoFactura(ByVal nNumFactura As Integer) As Integer
        Dim strSql As String
        Dim lValor As New Integer
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos2").ToString)
        Try
            cnnConexion2.Open()
            strSql = "SELECT TIPO" & _
                                " from ENM_ENCABEZADO_MATRICULA " & _
                                " Where FACTURA = " & nNumFactura & ""
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion2)
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try
    End Function

    Public Function fnFormateaFecha() As DateTime
        Dim lfFecha As DateTime = Date.Now
        Dim lcFechaFormateada As DateTime = ""
        Try
            lcFechaFormateada = Format(lfFecha, "dd/mm/yyyy")
        Catch ex As Exception
            Throw ex
        Finally

        End Try
        Return lcFechaFormateada
    End Function

    'Valida la existencia del convenio
    Private Function lfExisteConvenio(ByVal pCodConvenio As Byte) As Boolean
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            If pCodConvenio = eCodigoConvenio.LetraCambio Or _
                   pCodConvenio = eCodigoConvenio.Facturas Or _
                   pCodConvenio = eCodigoConvenio.CursosLibres Or _
                   pCodConvenio = eCodigoConvenio.Carnets Or _
                   pCodConvenio = eCodigoConvenio.VidaEstudiantil Or _
                   pCodConvenio = eCodigoConvenio.NoAsignado _
                   Then
                Return True

            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Consulta por la existencia en la Base de Datos de un Estudiante
    Private Function lfExisteEstudiante(ByVal pIdentificacion As String) As Boolean
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            strSql = "Select Count(id) " & _
                    "from EST_ESTUDIANTE " & _
                    "Where IDENTIFICACION ='" & pIdentificacion & "'"

            ' "AND TIPIOIDENTIFICACION=" & pTipoIdentificacion

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar

            If lValor = 0 Then
                Return False
            Else
                Return True
            End If
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Función que consulta si el estudiante posee Bloqueos
    Private Function lfBloqueoEstudiante(ByVal pIdentificacion As String) As String
        Dim ldtsDatos As New DataSet
        Dim strSql As String = String.Empty

        Try
            'Carga el Dataset con el detalle de los Bloqueos
            Dim adpObtener As New SqlDataAdapter(strSql, cnnConexion)
            adpObtener.Fill(ldtsDatos, "Resultado")
        Catch ex As Exception

        End Try

        Return "1"
    End Function


    'Funcion que Obtiene el ID del Estudiante
    Private Function pfObtenerIDEstudiante(ByVal pIdentificacion As String) As Integer
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            strSql = "Select ID " & _
                    "from EST_ESTUDIANTE " & _
                    "Where IDENTIFICACION ='" & pIdentificacion & "'"


            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar

            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Funcion que Obtiene el ID de la Factura
    Private Function pfObtenerIDFactura(ByVal pNumeroFactura As Integer) As Integer
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos2").ToString)
        Try
            strSql = "Select ID " & _
                    "from ENM_ENCABEZADO_MATRICULA " & _
                    "Where FACTURA =" & pNumeroFactura


            Dim cmdObtener As New SqlCommand(strSql, cnnConexion2)

            If cnnConexion2.State = ConnectionState.Open Then cnnConexion2.Close()
            cnnConexion2.Open()
            lValor = cmdObtener.ExecuteScalar

            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Funcion que Obtiene el Numero de Recibo
    Private Function pfObtenerNumeroRecibo() As Integer
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            strSql = "Select (max(RECIBO) +1) as Recibo  " & _
                    "from REC_RECIBO"
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Funcion que Obtiene el Numero de Recibo de Pago de Letra
    Private Function pfObtenerNumeroReciboLetraCambio() As Integer
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            strSql = "Select (max(RECIBO) +1) as Recibo  " & _
                    "from ALE_ABONO_LETRA"
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
            Return lValor
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Consulta si para  el servicio existen Pendientes de Pago segun los tipos de Convenio
    Private Function lfExistenPendientes(ByVal pCodConvenio As Byte, ByVal strLlaveAcceso As String, ByVal pTipoIdentificacion As Integer) As Boolean

        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0


        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()

            Select Case pCodConvenio

                Case eCodigoConvenio.Facturas
                    strSql = "SELECT COUNT(M.ID) " & _
                            "FROM  ENM_ENCABEZADO_MATRICULA M INNER JOIN EST_ESTUDIANTE E " & _
                                "ON M.EST_ID =E.ID " & _
                            "WHERE " & _
                                "M.CANCELADA = 0 " & _
                                "AND M.TIPO = 0 " & _
                                "AND M.TOTAL > 0 " & _
                                "AND E.IDENTIFICACION ='" & strLlaveAcceso & "'" & _
                                "AND M.PER_ID =" & pfObtenerPeriodoActivo()
                    ' & _ "AND E.TIPIOIDENTIFICACION=" & pTipoIdentificacion

                Case eCodigoConvenio.LetraCambio

                    strSql = "SELECT COUNT (L.ID) " & _
                            "FROM ENL_ENCABEZADO_LETRA L INNER JOIN EST_ESTUDIANTE E " & _
                                "ON E.ID=L.EST_ID " & _
                            "WHERE " & _
                                 "L.CANCELADA = 0 " & _
                                 "AND L.ESTADO =0 " & _
                                 "AND L.ACTIVA=1  " & _
                                 "AND E.IDENTIFICACION ='" & strLlaveAcceso & "' "
                    '& _ "AND E.TIPIOIDENTIFICACION=" & pTipoIdentificacion


                Case eCodigoConvenio.Carnets

                    strSql = String.Empty

                Case eCodigoConvenio.CursosLibres
                    strSql = String.Empty

                Case eCodigoConvenio.VidaEstudiantil
                    strSql = String.Empty

            End Select

            If strSql <> "" Then


                Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                lValor = cmdObtener.ExecuteScalar

                If lValor = 0 Then
                    Return False
                Else
                    Return True
                End If
                cnnConexion.Close()

            End If


        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Funcion que devuelve el Dataset con los valores cargados, para la clase de Servicios Pendientes
    Private Function lfConsultaPendientes(ByVal pCodConvenio As Byte, ByVal pIdentificacion As String, ByVal pTipoIdentificacion As Integer) As DataSet

        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Dim ldtsDatos As New DataSet
        Dim ldtsSubdetalle As New DataSet
        Dim lConsecutivo As Byte = 0
        Dim lConsecutivo_Subdetalle As Byte = 0
        Dim strDescripcionConceptp As String
        Dim nEstId As Integer = pfObtenerIDEstudiante(pIdentificacion)

        'Procedimiento para crear las tablas del DataSet de retorno. 
        'Creamos el data set que contendrá la información del Afiliado para el archivo de Envío
        Dim dtsPendientesPago As New Data.DataSet("dtsPendientesPago")
        'Se crea un DataTable para cada tabla y se asocia al Dataset de Pendietnes de Pago
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim dtbDetalle As Data.DataTable = New DataTable("Detalle")
        Dim dtbSubDetalle As Data.DataTable = New DataTable("SubDetalle")
        Dim ldtrFila As DataRow
        Dim ldtrSubdetalle As DataRow
        Dim nTotalConceptosPago As Integer = 0
        Dim nFactura As Integer
        Dim nDescuento As Decimal = 0.0
        Dim nTotalConDescuento As Decimal = 0.0
        'Crea las columnas que componen cada DataTable

        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nTipoIdentificacion", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("strIdentificacion", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strNombreCliente", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strCodMoneda", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionMoneda", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nCantidadSevicios", Type.GetType("System.Int32"))
        dtsPendientesPago.Tables.Add(dtbEncabezado)
        '-------------------------------------------------------------------------------------


        'TABLA DEL DETALLE
        '-------------------------------------------------------------------------------------
        dtbDetalle.Columns.Add("strValorServicio", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strDescripcionValorServicio", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strPeriodoRecibo", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nMontoTotal", Type.GetType("System.Decimal"))
        dtbDetalle.Columns.Add("nNumfactura", Type.GetType("System.Int32"))
        dtbDetalle.Columns.Add("strRecibo", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nSelfVerificacion", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nMontoInteres", Type.GetType("System.Decimal"))
        dtbDetalle.Columns.Add("nMontoMora", Type.GetType("System.Decimal"))
        dtbDetalle.Columns.Add("strFechaVencimiento", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strFechaVigenciaDesde", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strFechaVigenciaHasta", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nPrioridadPago", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nNumCuota", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nCantRubros", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Byte"))
        dtsPendientesPago.Tables.Add(dtbDetalle)
        '-------------------------------------------------------------------------------------



        'TABLA DEL SUBDETALLE
        '-------------------------------------------------------------------------------------
        dtbSubDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Int32"))
        dtbSubDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbSubDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbSubDetalle.Columns.Add("nMontoConcepto", Type.GetType("System.Decimal"))
        dtsPendientesPago.Tables.Add(dtbSubDetalle)
        '-------------------------------------------------------------------------------------

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()


            'Carga la lista de Servicios pendientes según el convenio y estudiante
            ldtsDatos = lfObtenerListaServiciosPendientes(pIdentificacion, nEstId, pCodConvenio)


            'Si el Dataset trae servicios pendientes procede, si no envia un mensaje de error en el encabezado
            If ldtsDatos.Tables(0).Rows.Count > 0 Then

                Select Case pCodConvenio


                    Case eCodigoConvenio.Facturas
                        'Recorre la lista de Servicios Pendientes y va asignando los valores al DataTable Detalle

                        nFactura = ldtsDatos.Tables(0).Rows(0)("nNumfactura")
                        lValor = ldtsDatos.Tables(0).Rows.Count
                        If lfConsultaTipoFactura(nFactura) = 0 Then
                            ldtrFila = dtsPendientesPago.Tables("Encabezado").NewRow
                            ldtrFila("strCodRespuesta") = cCodMensaje00
                            ldtrFila("strDescripcionRespuesta") = strMensaje00
                            ldtrFila("nTipoIdentificacion") = pTipoIdentificacion
                            ldtrFila("strIdentificacion") = pIdentificacion
                            ldtrFila("strNombreCliente") = lfNombreEstudiante(pTipoIdentificacion, nEstId)
                            ldtrFila("strCodMoneda") = "01"
                            ldtrFila("strDescripcionMoneda") = "Colones"
                            ldtrFila("nCantidadSevicios") = lValor
                            dtsPendientesPago.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table
                        Else
                            ldtrFila = dtsPendientesPago.Tables("Encabezado").NewRow
                            ldtrFila("strCodRespuesta") = cCodMensaje51
                            ldtrFila("strDescripcionRespuesta") = "Tipo de Factura inválida."
                            ldtrFila("nTipoIdentificacion") = pTipoIdentificacion
                            ldtrFila("strIdentificacion") = pIdentificacion
                            ldtrFila("strNombreCliente") = lfNombreEstudiante(pTipoIdentificacion, nEstId)
                            ldtrFila("strCodMoneda") = "01"
                            ldtrFila("strDescripcionMoneda") = "Colones"
                            ldtrFila("nCantidadSevicios") = lValor
                            dtsPendientesPago.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table
                        End If

                        lConsecutivo = lConsecutivo + 1
                        Dim lValoresFacturacion(3) As Decimal
                        For i As Integer = 0 To ldtsDatos.Tables(0).Rows.Count - 1
                            lValoresFacturacion = pfDetalleFacturacion(nEstId)

                            'Crea un DataRow del Tipo de la Tabla de Detalle
                            ldtrFila = dtsPendientesPago.Tables("Detalle").NewRow
                            ldtrFila("strValorServicio") = ldtsDatos.Tables(0).Rows(i)("strValorServicio").ToString
                            ldtrFila("strDescripcionValorServicio") = ldtsDatos.Tables(0).Rows(i)("strDescripcionValorServicio").ToString
                            ldtrFila("strPeriodoRecibo") = ldtsDatos.Tables(0).Rows(i)("strPeriodoRecibo").ToString
                            ldtrFila("nMontoTotal") = lValoresFacturacion(eColumnasArregloDetalleFactura.MontoTotal)
                            ldtrFila("nNumfactura") = ldtsDatos.Tables(0).Rows(i)("nNumfactura")
                            ldtrFila("strRecibo") = ldtsDatos.Tables(0).Rows(i)("strRecibo").ToString
                            ldtrFila("nSelfVerificacion") = ldtsDatos.Tables(0).Rows(i)("nSelfVerificacion")
                            ldtrFila("nMontoInteres") = ldtsDatos.Tables(0).Rows(i)("nMontoInteres")
                            ldtrFila("nMontoMora") = ldtsDatos.Tables(0).Rows(i)("nMontoMora")
                            ldtrFila("strFechaVencimiento") = ldtsDatos.Tables(0).Rows(i)("strFechaVencimiento").ToString
                            ldtrFila("strFechaVigenciaDesde") = ldtsDatos.Tables(0).Rows(i)("strFechaVigenciaDesde").ToString
                            ldtrFila("nPrioridadPago") = ldtsDatos.Tables(0).Rows(i)("nPrioridadPago")
                            ldtrFila("nNumCuota") = ldtsDatos.Tables(0).Rows(i)("nNumCuota")
                            ldtrFila("nCantRubros") = 0 'ldtsDatos.Tables(0).Rows(i)("nCantRubros")
                            ldtrFila("nConsecutivo") = lConsecutivo
                            dtsPendientesPago.Tables("Detalle").Rows.Add(ldtrFila)

                            'Carga los Subdetalles del servicio Pendiente

                            ldtrSubdetalle = dtsPendientesPago.Tables("SubDetalle").NewRow
                            ldtrSubdetalle("nConsecutivo") = 1
                            ldtrSubdetalle("nCodConcepto") = "01"
                            ldtrSubdetalle("strDescripcionConcepto") = "Pago de Factura de Matrícula "
                            ldtrSubdetalle("nMontoConcepto") = lValoresFacturacion(eColumnasArregloDetalleFactura.MontoTotal)
                            dtsPendientesPago.Tables("Subdetalle").Rows.Add(ldtrSubdetalle)  'Agrega la Fila al Data Table

                            '********************************************************************************************
                            lConsecutivo = lConsecutivo + 1


                        Next





                    Case eCodigoConvenio.LetraCambio

                        'Recorre la lista de Servicios Pendientes y va asignando los valores al DataTable Detalle
                        Dim lNumeroLetraCambio As String = String.Empty
                        Dim lValoresLetraCambio(10) As Decimal
                        Dim lConsecutivoDetalle As Byte = 0
                        'lValoresLetraCambio = pfDetalleLetraCambio(lNumeroLetraCambio)
                        lConsecutivo = lConsecutivo + 1
                        lValor = ldtsDatos.Tables(0).Rows.Count

                        ' For Each ldtr_ServiciosPendientes As DataRow In ldtsDatos.Tables(0).Rows

                        For i As Integer = 0 To ldtsDatos.Tables(0).Rows.Count - 1
                            'Carga los valores de la letra de cambio
                            '******************************************************************
                            lNumeroLetraCambio = ldtsDatos.Tables(0).Rows(i)("strRecibo")
                            lValoresLetraCambio = pfDetalleLetraCambio(lNumeroLetraCambio)
                            '*******************************************************************
                            ldtrFila = dtsPendientesPago.Tables("Encabezado").NewRow
                            ldtrFila("strCodRespuesta") = cCodMensaje00
                            ldtrFila("strDescripcionRespuesta") = strMensaje00
                            ldtrFila("nTipoIdentificacion") = pTipoIdentificacion
                            ldtrFila("strIdentificacion") = pIdentificacion
                            ldtrFila("strNombreCliente") = lfNombreEstudiante(pTipoIdentificacion, nEstId)
                            ldtrFila("strCodMoneda") = "01"
                            ldtrFila("strDescripcionMoneda") = "Colones"
                            ldtrFila("nCantidadSevicios") = lValor
                            dtsPendientesPago.Tables("Encabezado").Rows.Add(ldtrFila)


                            'Crea un DataRow del Tipo de la Tabla de Detalle
                            ldtrFila = dtsPendientesPago.Tables("Detalle").NewRow
                            ldtrFila("strValorServicio") = ldtsDatos.Tables(0).Rows(i)("strValorServicio").ToString
                            ldtrFila("strDescripcionValorServicio") = ldtsDatos.Tables(0).Rows(i)("strDescripcionValorServicio").ToString
                            ldtrFila("strPeriodoRecibo") = ldtsDatos.Tables(0).Rows(i)("strPeriodoRecibo").ToString
                            'ldtrFila("nMontoTotal") = lValoresLetraCambio(eColumnasArregloLetras.SaldoLetra)
                            ldtrFila("nMontoTotal") = lValoresLetraCambio(eColumnasArregloLetras.MontoTotal)
                            ldtrFila("nNumfactura") = lfMaxRecibo() 'ldtsDatos.Tables(0).Rows(0)("nNumfactura")
                            ldtrFila("strRecibo") = ldtsDatos.Tables(0).Rows(i)("strValorServicio").ToString
                            ldtrFila("nSelfVerificacion") = ldtsDatos.Tables(0).Rows(i)("nSelfVerificacion")
                            ldtrFila("nMontoInteres") = lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario)
                            ldtrFila("nMontoMora") = lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio)
                            ldtrFila("strFechaVencimiento") = ldtsDatos.Tables(0).Rows(i)("strFechaVencimiento").ToString
                            ldtrFila("strFechaVigenciaDesde") = ldtsDatos.Tables(0).Rows(i)("strFechaVigenciaDesde").ToString
                            ldtrFila("nPrioridadPago") = ldtsDatos.Tables(0).Rows(i)("nPrioridadPago")
                            ldtrFila("nNumCuota") = ldtsDatos.Tables(0).Rows(i)("nNumCuota")
                            ldtrFila("nCantRubros") = ldtsDatos.Tables(0).Rows(i)("nCantRubros")
                            ldtrFila("nConsecutivo") = lConsecutivo
                            dtsPendientesPago.Tables("Detalle").Rows.Add(ldtrFila)
                            'Incrementa el Consecutivo
                            lConsecutivo = lConsecutivo + 1

                            'Carga los Subdetalles del servicio Pendiente
                            'Para la letra de cambio solo se envía en el Subdetalle el Numero de Letra a la que se le realizará el abono
                            '******************************************************************************************

                            strDescripcionConceptp = "Pago/Abono L.C: " & lNumeroLetraCambio
                            ldtrSubdetalle = dtsPendientesPago.Tables("SubDetalle").NewRow
                            ldtrSubdetalle("nConsecutivo") = 1
                            ldtrSubdetalle("nCodConcepto") = "01"
                            ldtrSubdetalle("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                            'ldtrSubdetalle("nMontoConcepto") = lValoresLetraCambio(eColumnasArregloLetras.SaldoLetra)
                            ldtrSubdetalle("nMontoConcepto") = lValoresLetraCambio(eColumnasArregloLetras.MontoTotal)
                            dtsPendientesPago.Tables("Subdetalle").Rows.Add(ldtrSubdetalle)  'Agrega la Fila al Data Table


                            '********************************************************************************************


                        Next i


                    Case eCodigoConvenio.Carnets

                        strSql = String.Empty

                    Case eCodigoConvenio.CursosLibres
                        strSql = String.Empty

                    Case eCodigoConvenio.VidaEstudiantil
                        strSql = String.Empty

                End Select

            Else

                'Crea la fila del encabezado ****************************************************************
                ldtrFila = dtsPendientesPago.Tables("Encabezado").NewRow
                ldtrFila("strCodRespuesta") = cCodMensaje51
                ldtrFila("strDescripcionRespuesta") = "No existen pendientes actualmente."
                ldtrFila("nTipoIdentificacion") = pTipoIdentificacion
                ldtrFila("strIdentificacion") = pIdentificacion
                ldtrFila("strNombreCliente") = lfNombreEstudiante(pTipoIdentificacion, nEstId)
                ldtrFila("strCodMoneda") = "01"
                ldtrFila("strDescripcionMoneda") = "Colones"
                ldtrFila("nCantidadSevicios") = 0
                dtsPendientesPago.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table
                '********************************************************************************************

                'Crea un DataRow del Tipo de la Tabla de Detalle
                ldtrFila = dtsPendientesPago.Tables("Detalle").NewRow
                ldtrFila("strValorServicio") = 0
                ldtrFila("strDescripcionValorServicio") = 0
                ldtrFila("strPeriodoRecibo") = 0
                ldtrFila("nMontoTotal") = 0
                ldtrFila("nNumfactura") = 0
                ldtrFila("strRecibo") = 0
                ldtrFila("nSelfVerificacion") = 0
                ldtrFila("nMontoInteres") = 0
                ldtrFila("nMontoMora") = 0
                ldtrFila("strFechaVencimiento") = 0 / 0 / 0
                ldtrFila("strFechaVigenciaDesde") = 0 / 0 / 0
                ldtrFila("nPrioridadPago") = 0
                ldtrFila("nNumCuota") = 0
                ldtrFila("nCantRubros") = 0
                ldtrFila("nConsecutivo") = 0
                dtsPendientesPago.Tables("Detalle").Rows.Add(ldtrFila)
                'Incrementa el Consecutivo
                lConsecutivo = lConsecutivo + 1

                'Carga los Subdetalles del servicio Pendiente
                'Para la letra de cambio solo se envía en el Subdetalle el Numero de Letra a la que se le realizará el abono
                '******************************************************************************************

                strDescripcionConceptp = "Pago/Abono L.C: " & 0
                ldtrSubdetalle = dtsPendientesPago.Tables("SubDetalle").NewRow
                ldtrSubdetalle("nConsecutivo") = 1
                ldtrSubdetalle("nCodConcepto") = 0
                ldtrSubdetalle("strDescripcionConcepto") = ""
                ldtrSubdetalle("nMontoConcepto") = 0
                dtsPendientesPago.Tables("Subdetalle").Rows.Add(ldtrSubdetalle)  'Agrega la Fila al Data Table


                '********************************************************************************************


            End If



        Catch ex As Exception
            Throw ex

        End Try

        Return dtsPendientesPago
    End Function

    'Obtiene el Nombre del Estudiante
    Private Function lfNombreEstudiante(ByVal pTipoIdentificacion As Byte, ByVal nEstID As Integer) As String
        Dim strSql As String = String.Empty
        Dim lValor As String = String.Empty
        Try
            strSql = "Select NombreCompleto " & _
                    "from EST_ESTUDIANTE " & _
                    "Where ID = " & nEstID & " "
            '& _ "AND TIPIOIDENTIFICACION=" & pTipoIdentificacion
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()

            lValor = cmdObtener.ExecuteScalar

            If lValor = String.Empty Then
                lValor = "ESTUDIANTE NO REGISTRADO"
            End If

            Return lValor

            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Obtiene el Nombre del Estudiante
    Private Function lfNombreEstudianteXIdentificacion(ByVal strIdentificacion As String) As String
        Dim strSql As String = String.Empty
        Dim lValor As String = String.Empty
        Try
            strSql = "Select NombreCompleto " & _
                    "from EST_ESTUDIANTE " & _
                    "Where IDENTIFICACION = '" & strIdentificacion & "' "
            '& _ "AND TIPIOIDENTIFICACION=" & pTipoIdentificacion
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()

            lValor = cmdObtener.ExecuteScalar

            If lValor = String.Empty Then
                lValor = "ESTUDIANTE NO REGISTRADO"
            End If

            Return lValor

            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Función que devuelve el dataset de los servicios pendientes de pago, según el convenio
    Private Function lfObtenerListaServiciosPendientes(ByVal pIdentificacion As String, ByVal ptipoIdentificacion As String, ByVal pCodigoConvenio As Integer) As DataSet
        Dim ldtsServiciosPendientes As New DataSet
        Dim strSql As String = String.Empty
        Dim strDescripcionValorServicio As String = "Servicios Pendientes"

        Try
            Select Case pCodigoConvenio
                Case eCodigoConvenio.Facturas
                    strDescripcionValorServicio = Mid("Pago Matricula-" & pfObtenerNombrePeriodo(), 1, 49)
                    Dim lPeriodoActivo As Integer = pfObtenerPeriodoActivo()
                    strSql = _
                                  "Select DISTINCT TOP 1 " & _
                                  "M.FACTURA as strValorServicio, " & _
                                  "'" & strDescripcionValorServicio & "' as strDescripcionValorServicio," & _
                                  "convert(varchar,M.FECHAINCLUSION,103) AS strPeriodoRecibo, " & _
                                  "M.TOTAL  as  nMontoTotal, " & _
                                  "M.FACTURA AS nNumfactura, " & _
                                  "M.ID AS strRecibo, " & _
                                  "0 AS nSelfVerificacion, " & _
                                  "0 AS nMontoInteres, " & _
                                  "0 AS nMontoMora, " & _
                                  "convert(varchar,M.FECHAINCLUSION,103) AS  strFechaVencimiento, " & _
                                  "convert(varchar,M.FECHAINCLUSION,103) AS strFechaVigenciaDesde, " & _
                                  "'01/01/1800' AS strFechaVigenciaHasta, " & _
                                  "0 AS nPrioridadPago," & _
                                  "0 AS nNumCuota, " & _
                                  "0 AS nCantRubros, " & _
                                  "0 AS nConsecutivo, " & _
                                  "M.FECHACREO " & _
                                  "From  ENM_ENCABEZADO_MATRICULA M " & _
                                      "inner join EST_Estudiante E " & _
                                          "On E.id=M.Est_id " & _
                                    "inner join DEM_DETALLE_MATRICULA D " & _
                                        "ON D.ENM_ID=M.ID " & _
                                    "INNER JOIN PLE_PLAN_x_ESTUDIANTE PE " & _
                                        "ON PE.ID=D.PLE_ID " & _
                                  "where Cancelada = 0 " & _
                                  "And M.Tipo=0 " & _
                                  "AND TOTAL > 0 " & _
                                  "And E.Identificacion='" & pIdentificacion & "' " & _
                                  " And PE.PERIODO_ID=" & lPeriodoActivo & " " & _
                                  " ORDER BY M.FECHACREO desc "

                Case eCodigoConvenio.LetraCambio
                    strDescripcionValorServicio = "Letra(s) de cambio pendientes al " & Date.Today.ToString("dd / MM / yyyy")
                    strSql = _
                                  "Select DISTINCT TOP 1 " & _
                                  "L.NUMERO as strValorServicio," & _
                                  "'" & strDescripcionValorServicio & "' as strDescripcionValorServicio," & _
                                  "convert(varchar,L.FECHAULTIMOPAGO,103) AS strPeriodoRecibo, " & _
                                  "L.SALDO  as  nMontoTotal, " & _
                                  "L.ID AS nNumfactura, " & _
                                  "L.NUMERO AS strRecibo, " & _
                                  "0 AS nSelfVerificacion, " & _
                                  "0 AS nMontoInteres, " & _
                                  "0 AS nMontoMora, " & _
                                  "convert(varchar,l.FECHAFINAL,103) AS  strFechaVencimiento, " & _
                                  "convert(varchar,l.FECHAINCLUSION,103) AS strFechaVigenciaDesde, " & _
                                  "'01/01/1800' AS strFechaVigenciaHasta, " & _
                                  "0 AS nPrioridadPago," & _
                                  "0 AS nNumCuota, " & _
                                  "1 AS nCantRubros, " & _
                                  "0 AS nConsecutivo, " & _
                                  "L.FECHA " & _
                                  "From  ENL_ENCABEZADO_LETRA L " & _
                                      "inner join EST_Estudiante E " & _
                                          "On E.id=L.Est_id " & _
                                  "where Cancelada = 0 " & _
                                  "And L.Activa=1  " & _
                                  "And E.Identificacion='" & pIdentificacion & "'" & _
                                " ORDER BY L.FECHA desc "
                    '& _ "And E.TIPIOIDENTIFICACION=" & ptipoIdentificacion

                    'Se envía un 1 en Cantidad de Rubros
                    'por solo detallar el numero de Letra y monto pagado en el abono




                Case eCodigoConvenio.Carnets

                    strSql = String.Empty

                Case eCodigoConvenio.CursosLibres
                    strSql = String.Empty

                Case eCodigoConvenio.VidaEstudiantil
                    strSql = String.Empty
            End Select

            'Carga el Dataset col los servicios pendientes, pude ser mas de un servicio
            Dim adpObtener As New SqlDataAdapter(strSql, cnnConexion)
            adpObtener.Fill(ldtsServiciosPendientes, "Resultado")


        Catch ex As Exception
            Throw ex
        End Try

        Return ldtsServiciosPendientes
    End Function

    Private Function lfObtenerFechaCreoRecibo(ByVal strRecibo As String) As Date
        Dim StrSql As String
        Dim lValor As Date
        Try
            StrSql = "SELECT FECHACREO FROM ENM_ENCABEZADO_MATRICULA WHERE NUMEROLETRA =  '" & strRecibo & "' "

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(StrSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
        Catch ex As Exception
            Throw ex
        Finally

        End Try
        Return lValor
    End Function

    Private Function lfObtenerFechaCreoFactura(ByVal pNumFactura As Integer) As Date
        Dim StrSql As String
        Dim lValor As Date
        Try
            StrSql = "SELECT ALE.FECHA FROM ALE_ABONO_LETRA " & _
                     "INNER JOIN ENL_ENCABEZADO_LETRA enl " & _
                     "WHERE ENL.ID =  " & pNumFactura & " "

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(StrSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
        Catch ex As Exception
            Throw ex
        Finally

        End Try
        Return lValor
    End Function

    Private Function lfExisteFactura(ByVal nNumFactura As Integer) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Try
            strSql = "SELECT count(FACTURA) " & _
                    " from ENM_ENCABEZADO_MATRICULA " & _
                    " Where FACTURA = " & nNumFactura & ""
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteLetra(ByVal nEnlID As Integer) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Try
            strSql = "  SELECT COUNT(ID) " & _
                     "from ENL_ENCABEZADO_LETRA " & _
                     "WHERE ID = " & nEnlID & ""
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfObtieneRecibo(ByVal nEnlID As Integer, ByVal nNumeroComprobante As Integer) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        'Dim nENL_ID As Integer = pfObtenerIDLetraCambio(strNumero)
        Try
            strSql = "SELECT rec.RECIBO " & _
                     "from ALE_ABONO_LETRA ale " & _
                     "LEFT JOIN REC_RECIBO rec " & _
                     "on ale.RECIBO = rec.RECIBO " & _
                     "WHERE ale.ENL_ID = " & nEnlID & " AND rec.NUMCOMPROBANTE = " & nNumeroComprobante & ""
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteNumComprobanteLetra(ByVal nEnlID As Integer, ByVal nNumComprobante As Integer) As Boolean
        Dim lValor As Integer
        Dim bRetorno As Boolean
        Dim strSql As String = String.Empty
        'Dim nENL_ID As Integer = pfObtenerIDLetraCambio(strNumero)
        Try
            strSql = "SELECT COUNT(rec.NUMCOMPROBANTE) " & _
                     "from ALE_ABONO_LETRA ale " & _
                     "LEFT JOIN REC_RECIBO rec " & _
                     "on ale.RECIBO = rec.RECIBO " & _
                     "WHERE ale.ENL_ID = " & nEnlID & " AND rec.NUMCOMPROBANTE = " & nNumComprobante & ""
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            If lValor > 0 Then
                bRetorno = True
            Else
                bRetorno = False
            End If
            Return bRetorno
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteNumComprobanteXLetra(ByVal nEnlID As Integer, ByVal nNumeroComprobante As Integer) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        'Dim nENL_ID As Integer = pfObtenerIDLetraCambio(strNumero)
        Try
            strSql = "SELECT COUNT(rec.ID) " & _
                     "from ALE_ABONO_LETRA ale " & _
                     "LEFT JOIN REC_RECIBO rec " & _
                     "on ale.RECIBO = rec.RECIBO " & _
                     "WHERE ale.ENL_ID = " & nEnlID & " AND rec.NUMCOMPROBANTE = " & nNumeroComprobante & ""
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteNumComprobanteXLetraXMonto(ByVal nENL_ID As String, ByVal nNumeroComprobante As String, ByVal nMonto As Decimal) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        'Dim nENL_ID As Integer = pfObtenerIDLetraCambio(strNumero)
        Try
            strSql = "SELECT COUNT(rec.ID) " & _
                     "from ALE_ABONO_LETRA ale " & _
                     "INNER JOIN REC_RECIBO rec " & _
                     "on ale.RECIBO = rec.RECIBO " & _
                     "WHERE ale.ENL_ID = " & nENL_ID & " AND rec.NUMCOMPROBANTE = " & nNumeroComprobante & " and rec.TOTAL = " & nMonto & ""
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteNumFactura(ByVal nFactura As Integer) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Try
            strSql = "  SELECT COUNT(ID) " & _
                     "from ENL_ENCABEZADO_LETRA " & _
                     "WHERE ID = " & nFactura & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteFactura(ByVal strFactura As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Dim nFactura As Integer
        Try
            nFactura = lfObtieneIDFactura(strFactura)
            strSql = "  SELECT COUNT(ID) " & _
                     "from ENM_ENCABEZADO_MATRICULA " & _
                     "WHERE ID = " & nFactura & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfObtieneIDFactura(ByVal strFactura As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Dim nFactura As Integer = CInt(strFactura)
        Try

            strSql = "  SELECT ID " & _
                     "from ENM_ENCABEZADO_MATRICULA " & _
                     "WHERE FACTURA = " & nFactura & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfConsultaTipoFactura(ByVal strFactura As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Dim nFactura As String
        Try
            nFactura = CInt(strFactura)
            strSql = "  SELECT TIPO " & _
                     "from ENM_ENCABEZADO_MATRICULA " & _
                     "WHERE FACTURA = " & nFactura & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function



    Private Function lfObtieneIDEstudianteXFactura(ByVal strLlaveAcceso As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Dim nFactura As String
        Try
            nFactura = CInt(strLlaveAcceso)
            strSql = "  SELECT TOP 1 EST_ID " & _
                     "from ENM_ENCABEZADO_MATRICULA " & _
                     "WHERE Factura = " & nFactura & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function


    Private Function lfExisteENL(ByVal strRecibo As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty

        Try
            'nRecibo = CInt(strRecibo)
            strSql = "  SELECT COUNT(ID) " & _
                     "from ENL_ENCABEZADO_LETRA " & _
                     "WHERE NUMERO = '" & strRecibo & "' "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteRecibo(ByVal strRecibo As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty

        Try
            'nRecibo = CInt(strRecibo)
            strSql = "  SELECT COUNT(ID) " & _
                     "from ENL_ENCABEZADO_LETRA " & _
                     "WHERE NUMERO = '" & strRecibo & "' "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfObtieneIDEstudianteXRecibo(ByVal strRecibo As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Try
            ' nRecibo = CInt(strRecibo)
            strSql = "SELECT TOP 1 EST_ID " & _
                     "from ENL_ENCABEZADO_LETRA " & _
                     "WHERE NUMERO = '" & strRecibo & "' "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function

    Private Function lfExisteRecibo(ByVal nEnlID As Integer, ByVal nNumComprobante As Integer) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty

        Try
            ' nRecibo = CInt(strRecibo)
            strSql = " select COUNT(rec.ID) from REC_RECIBO rec where " & _
                     "  rec.ID = " & _
                     "  (select top 1 rec.id from ALE_ABONO_LETRA ale " & _
                     "  inner join REC_RECIBO rec " & _
                     "  on ale.RECIBO = rec.RECIBO " & _
                     "  where rec.NUMCOMPROBANTE = " & nNumComprobante & " and ale.ENL_ID = " & nEnlID & " ) "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try
    End Function




    Private Function lfObtenerIDNumeroLetra(ByVal strRecibo As String) As Integer
        Dim lValor As Integer
        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try
            cnnConexion2.Open()
            strSql = "SELECT ID FROM ENL_ENCABEZADO_LETRA " & _
                     " WHERE NUMERO = '" & strRecibo & "'"
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComando As New SqlCommand(strSql, cnnConexion)
            lValor = cmdComando.ExecuteScalar
            Return lValor

        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try

    End Function

    Private Function lfInsertarBitacoraBRP(ByVal nNumFactura As Integer, ByVal lstrNumeroLetra As String) As String
        Dim strSql As String
        Try
            strSql = "INSERT INTO BRA_BITACORA_REVERSIONES_ABONO_LETRA " & _
                            "( ID " & _
                             " ,ALE_ID " & _
                             " ,ENL_ID " & _
                             " ,RECIBO " & _
                             " ,DIASACUMULADO " & _
                             " ,TOTALPAGO " & _
                             " ,ABONO " & _
                             " ,SALDO " & _
                             " ,OBSERVACION " & _
                             " ,FECHA " & _
                             " ,DIASMORATORIOS " & _
                             " ,INTERESORDINARIO " & _
                             " ,ABONOINTERESORDINARIO " & _
                             " ,SALDOINTERESORDINARIO " & _
                             " ,INTERESMORATORIO " & _
                             " ,ABONOINTERESMORATORIO " & _
                             " ,SALDOINTERESMORATORIO " & _
                           " VALUES " & _
                                "(" & _
                         "  " & nNumFactura & "," & _
                         "  '" & lstrNumeroLetra & "'," & _
                         " getDate(),  " & _
                         " '" & cCreo & "' )"
            Dim cmdInsertaBitacora As New SqlCommand(strSql, cnnConexion)
            cmdInsertaBitacora.Transaction = grefTransaction
            cmdInsertaBitacora.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfActualizaENM(ByVal nNumFactura As Integer) As String
        Dim strSql As String
        Try
            strSql = "UPDATE ENM_ENCABEZADO_MATRICULA SET " & _
                            " FACTURADO = " & cNoFacturado & " " & _
                              " ,TOTAL = " & cTotal & " " & _
                              " ,EFECTIVO = " & cEfectivo & " " & _
                              " ,TARJETA = " & cTarjeta & " " & _
                              " ,CANCELADA = " & cCancelada & " " & _
                              " ,DESCUENTO = " & CDescuento & " " & _
                              " ,TOTALSINDESCUENTO = " & cTotalSinDescuento & " " & _
                              " WHERE FACTURA = " & nNumFactura

            Dim cmdIngresaMatricula As New SqlCommand(strSql, cnnConexion)
            cmdIngresaMatricula.Transaction = grefTransaction
            cmdIngresaMatricula.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfActualizaDEF(ByVal nNumFactura As Integer) As String
        Dim strSql As String
        Try
            strSql = " DECLARE @FACTURA int = " & _
                     " ( SELECT FACTURA FROM DEF_DETALLE_FACTURA def " & _
                        "INNER JOIN ENM_ENCABEZADO_MATRICULA enm " & _
                        "ON def.ENM_ID = enm.ID " & _
                        "WHERE enm.FACTURA = " & nNumFactura & ")" & _
                    "UPDATE DEF_DETALLE_FACTURA SET " & _
                              " FACTURADO = " & cNoFacturado & " " & _
                              " ,TOTAL = " & cTotal & " " & _
                              " ,EFECTIVO = " & cEfectivo & " " & _
                              " ,TARJETA = " & cTarjeta & " " & _
                              " ,CANCELADA = " & cCancelada & " " & _
                              " ,DESCUENTO = " & CDescuento & " " & _
                              " ,TOTALSINDESCUENTO = " & cTotalSinDescuento & " " & _
                              " WHERE FACTURA = @FACTURA "

            Dim cmdActualizaDEF As New SqlCommand(strSql, cnnConexion)
            cmdActualizaDEF.Transaction = grefTransaction
            cmdActualizaDEF.ExecuteNonQuery()
            cnnConexion.Close()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfActualizaENL1(ByVal nNumeroFactura As Integer, ByVal strEstIdentificacion As String, ByVal nMontoRecibo As Double, ByVal nEnlID As Integer) As String
        Dim strSql As String
        Try
            strSql = " DECLARE @EST_ID int " & _
                     "SET @EST_ID = (  SELECT est.ID FROM EST_ESTUDIANTE est " & _
                     "INNER JOIN ENM_ENCABEZADO_MATRICULA enm " & _
                     "ON Est.ID = enm.EST_ID " & _
                     "WHERE (est.IDENTIFICACION = ' " & strEstIdentificacion & "')" & _
                     " AND (enm.FACTURA = " & nNumeroFactura & ") )" & _
                    "UPDATE ENL_ENCABEZADO_LETRA SET " & _
                              " TOTALPAGADO = TOTALPAGADO - " & nMontoRecibo & " " & _
                              " ,SALDO =  SALDO + " & nMontoRecibo & " " & _
                              " ,ESTADO =  3 " & _
                              " ,CUOTAS = CUOTAS + 1" & _
                              " WHERE EST_ID = @EST_ID "
            Dim cmdActualizaENL As New SqlCommand(strSql, cnnConexion)
            cmdActualizaENL.Transaction = grefTransaction
            cmdActualizaENL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function



    Private Function lfEliminaRecibo(ByVal nNumFactura As Integer) As String
        Dim strSql As String
        strSql = "DECLARE @Recibo int = " & _
                 "(SELECT top 1 rec.RECIBO FROM ENM_ENCABEZADO_MATRICULA enm " & _
                 " INNER JOIN REC_RECIBO rec " & _
                 " ON enm.EST_ID = rec.EST_ID" & _
                 " WHERE(enm.FACTURA = " & nNumFactura & " )" & _
                 " ORDER BY rec.FECHACREO desc )" & _
                 "DELETE FROM REC_RECIBO WHERE RECIBO = @Recibo"

        Try

            Dim cmdEliminaRecibo As New SqlCommand(strSql, cnnConexion)
            cmdEliminaRecibo.Transaction = grefTransaction
            cmdEliminaRecibo.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
        cnnConexion.Close()
    End Function

    'Función que devuelve la cantidad de servicios pendientes de pago
    Private Function lfCantidadServiciosPendientes(ByVal pTipoIdentificacion As String, ByVal pIdentificacion As String, ByVal pCodConvenio As Integer) As Integer

        Dim strSql As String = String.Empty
        Dim lvalor As Integer = 0

        If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
        Try
            Select Case pCodConvenio
                Case eCodigoConvenio.Facturas

                    strSql = "SELECT COUNT(M.ID) " & _
                           "FROM  ENM_ENCABEZADO_MATRICULA M INNER JOIN EST_ESTUDIANTE E " & _
                               "ON M.EST_ID =E.ID " & _
                           "WHERE " & _
                               "M.CANCELADA = 0 " & _
                               "AND M.TIPO=0 " & _
                               "AND E.IDENTIFICACION ='" & pIdentificacion & "'"
                    '& _"AND E.TIPIOIDENTIFICACION=" & pTipoIdentificacion


                Case eCodigoConvenio.LetraCambio

                    strSql = "SELECT COUNT(L.ID) " & _
                           "FROM  ENL_ENCABEZADO_LETRA L INNER JOIN EST_ESTUDIANTE E " & _
                               "ON L.EST_ID =E.ID " & _
                           "WHERE " & _
                               "L.CANCELADA = 0 AND L.ACTIVA =1 AND  L.ESTADO =0 " & _
                               "AND E.IDENTIFICACION ='" & pIdentificacion & "' "
                    '& _ "AND E.TIPIOIDENTIFICACION=" & pTipoIdentificacion

                Case eCodigoConvenio.Carnets

                    strSql = String.Empty

                Case eCodigoConvenio.CursosLibres
                    strSql = String.Empty

                Case eCodigoConvenio.VidaEstudiantil
                    strSql = String.Empty
            End Select


            If strSql <> String.Empty Then
                Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                lvalor = cmdObtener.ExecuteScalar
                cnnConexion.Close()
            End If
            Return lvalor

        Catch ex As Exception

        End Try
    End Function

    'Funcion que determina si el estudiante posee algun bloqueo
    Private Function pfBloqueoEstudiante(ByVal pIdentificacion As String) As Boolean
        Dim lValor As Integer
        Dim strSql = "SELECT COUNT(B.ID) " & _
                    "FROM BLE_BLOQUEO_ESTUDIANTE B INNER JOIN EST_ESTUDIANTE E " & _
                        "ON E.ID=B.EST_ID " & _
                    "WHERE ESTADO = 1 " & _
                        "AND IDENTIFICACION ='" & pIdentificacion & "'"

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()

            If lValor > 0 Then
                Return True
            Else
                Return False
            End If


        Catch ex As Exception
            Return False
        End Try

    End Function


    'Funcion que determina si el estudiante posee Becas
    Private Function pfBecaEstudiante(ByVal pIdentificacion As String) As Boolean
        Dim lValor As Integer
        Dim strSql = "Select count(B.id) From Bee_Beca_X_Estudiante B inner Join " & _
                                "EST_estudiante E on E.ID=B.Est_ID Where E.Identificacion='" & pIdentificacion & "'"


        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()

            If lValor > 0 Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return False
        End Try

    End Function


    'Funcion que valida el usuario y clave de acceso
    Private Function lfValidaUsuario(ByVal strUsuario As String, ByVal strClave As String) As Boolean

        If strUsuario = cstrUsuario And strClave = cstrClaveUsuario Then
            Return True
        Else
            Return False
        End If

    End Function


    'Función que retorna un DTS con la información del error para la consulta de servicios pendientes
    Private Function lfInformacionTipoError(ByVal pCodigoError As Byte, ByVal strIdentificacion As String) As DataSet
        Dim strMensaje As String = String.Empty
        Dim strCodMensaje As String = String.Empty
        Dim ldtrFila As DataRow
        'Creamos el data set que contendrá la información del Estudiante
        Dim dtsPendientesPago As New Data.DataSet("dtsPendientesPago")
        'Procedimiento para crear las tablas del DataSet de retorno
        ''Se crea un DataTable para cada tabla y se asocia al Dataset de Pendietnes de Pago
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim dtbDetalle As Data.DataTable = New DataTable("Detalle")
        Dim dtbSubDetalle As Data.DataTable = New DataTable("SubDetalle")
        Dim strEstNombre As String = lfNombreEstudianteXIdentificacion(strIdentificacion)
        'Crea las columnas que componen cada DataTable
        '-------------------------------------------------------------------------------------
        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nTipoIdentificacion", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("strIdentificacion", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strNombreCliente", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strCodMoneda", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionMoneda", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nCantidadSevicios", Type.GetType("System.Int32"))
        dtsPendientesPago.Tables.Add(dtbEncabezado)
        '-------------------------------------------------------------------------------------
        'TABLA DEL DETALLE
        '-------------------------------------------------------------------------------------
        dtbDetalle.Columns.Add("strValorServicio", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strDescripcionValorServicio", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strPeriodoRecibo", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nMontoTotal", Type.GetType("System.Decimal"))
        dtbDetalle.Columns.Add("nNumfactura", Type.GetType("System.Int32"))
        dtbDetalle.Columns.Add("strRecibo", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nSelfVerificacion", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nMontoInteres", Type.GetType("System.Decimal"))
        dtbDetalle.Columns.Add("nMontoMora", Type.GetType("System.Decimal"))
        dtbDetalle.Columns.Add("strFechaVencimiento", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strFechaVigenciaDesde", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("strFechaVigenciaHasta", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nPrioridadPago", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nNumCuota", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nCantRubros", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Byte"))
        dtsPendientesPago.Tables.Add(dtbDetalle)
        '-------------------------------------------------------------------------------------
        'TABLA DEL SUBDETALLE
        '-------------------------------------------------------------------------------------
        dtbSubDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Int32"))
        dtbSubDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbSubDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbSubDetalle.Columns.Add("nMontoConcepto", Type.GetType("System.Decimal"))
        dtsPendientesPago.Tables.Add(dtbSubDetalle)
        '-------------------------------------------------------------------------------------

        Try

            Select Case pCodigoError
                Case eMensajesErrorConsultaServicios.ConvenioNoExiste
                    strCodMensaje = cCodMensaje51
                    strMensaje = strMensaje05

                Case eMensajesErrorConsultaServicios.EstudianteConBeca
                    strCodMensaje = cCodMensaje51
                    strMensaje = strMensaje07

                Case eMensajesErrorConsultaServicios.EstudianteNoExiste
                    strCodMensaje = cCodMensaje51
                    strMensaje = strMensaje06


                Case eMensajesErrorConsultaServicios.NoHayPendientesPago
                    strCodMensaje = cCodMensaje51
                    strMensaje = strMensaje08

            End Select
            'Crea la fila del encabezado ****************************************************************
            ldtrFila = dtsPendientesPago.Tables("Encabezado").NewRow
            ldtrFila("strCodRespuesta") = strCodMensaje
            ldtrFila("strDescripcionRespuesta") = strMensaje
            ldtrFila("nTipoIdentificacion") = 99
            ldtrFila("strIdentificacion") = strIdentificacion
            ldtrFila("strNombreCliente") = strEstNombre
            ldtrFila("strCodMoneda") = "01"
            ldtrFila("strDescripcionMoneda") = "Colones"
            ldtrFila("nCantidadSevicios") = 0
            dtsPendientesPago.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table
            '********************************************************************************************
            'Crea Fila de Detalle ***********************************************************************
            ldtrFila = dtsPendientesPago.Tables("Detalle").NewRow
            ldtrFila("strValorServicio") = 0
            ldtrFila("strDescripcionValorServicio") = "N/R"
            ldtrFila("strPeriodoRecibo") = "N/R"
            ldtrFila("nMontoTotal") = 0
            ldtrFila("nNumfactura") = 0
            ldtrFila("strRecibo") = 0
            ldtrFila("nSelfVerificacion") = 0
            ldtrFila("nMontoInteres") = 0
            ldtrFila("nMontoMora") = 0
            ldtrFila("strFechaVencimiento") = "N/R"
            ldtrFila("strFechaVigenciaDesde") = "N/R"
            ldtrFila("strFechaVigenciaHasta") = "N/R"
            ldtrFila("nPrioridadPago") = 0
            ldtrFila("nNumCuota") = 0
            ldtrFila("nCantRubros") = 0
            ldtrFila("nConsecutivo") = 0
            dtsPendientesPago.Tables("Detalle").Rows.Add(ldtrFila)
            '***************************************************************
            'Crea Fila de Subdetalle****************************************
            ldtrFila = dtsPendientesPago.Tables("SubDetalle").NewRow
            ldtrFila("nConsecutivo") = 0
            ldtrFila("nCodConcepto") = 0
            ldtrFila("strDescripcionConcepto") = "N/R"
            ldtrFila("nMontoConcepto") = 0
            dtsPendientesPago.Tables("SubDetalle").Rows.Add(ldtrFila)
        Catch ex As Exception
        Finally

        End Try

        Return dtsPendientesPago




    End Function

    'Funcion que consutlta si el servicio esta diponible 
    Private Function lfServicioDisponible() As Boolean
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            strSql = "Select SERVICIO_DISPONIBLE " & _
                    "from PAR_PARAMETRO "

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar

            If lValor = 0 Then
                Return False 'El servicio no esta disponible
            Else
                Return True 'El servicio si está disponible
            End If
        Catch ex As Exception
            cnnConexion.Close()
            Return False
            Throw ex
        End Try
        cnnConexion.Close()
    End Function


    'Funcion que consutlta el pago mínimo de una letra de cambio
    Private Function lfPagoMinimoLetra() As Decimal
        Dim strSql As String = String.Empty
        Dim lValor As Decimal = 0.0
        Try
            strSql = "Select PAGO_MINIMO_LETRA " & _
                    "from PAR_PARAMETRO "

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar

            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            cnnConexion.Close()
            Return lValor
            Throw ex
        End Try
    End Function

#End Region


#Region "Procesos de Facturas"

#Region "Consultas de valores de Factura"
    'Función que obtiene los detalles del pago para la facturación
    Private Function pfDetalleFacturacion(ByVal nEstId As Integer) As Decimal()
        Dim ArregloDetalleFactura(6) As Decimal
        Dim lPerID As Integer = pfObtenerPeriodoActivo()

        'Inicializa el arreglo del Detalle de Facturas
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Descuento) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.MontoTotal) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalConDescuento) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalSinDescuento) = 0
        '********************************************************************

        Try

            'Actualiza el arreglo del Detalle de Facturas
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) = pfMontoMaterias(nEstId)
            'Se aplica descuento sobre total en materias
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Descuento) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) * cPorcentajeDescuento
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalConDescuento) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) - ArregloDetalleFactura(eColumnasArregloDetalleFactura.Descuento)
            '-------------------------------------------------------------------------------------'
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) = pfMontoMatricula(nEstId)
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro) = pfMontoSeguroPorPagar(nEstId, lPerID)
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.MontoTotal) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalConDescuento) + _
                                                                                ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) + _
                                                                                ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro)

            ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalSinDescuento) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) + _
                                                                                ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) + _
                                                                                ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro)
            '*******************************************************************

        Catch ex As Exception
            Throw ex
        End Try
        Return ArregloDetalleFactura
    End Function

    'Función que obtiene los detalles del pago para la facturación
    Private Function pfDetalleFacturacionCredomatic(ByVal nEstId As Integer) As Decimal()
        Dim ArregloDetalleFactura(5) As Decimal
        Dim lPerID As Integer = pfObtenerPeriodoActivo()
        'Inicializa el arreglo del Detalle de Facturas
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Descuento) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.MontoTotal) = 0
        ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalConDescuento) = 0
        '********************************************************************

        Try

            'Actualiza el arreglo del Detalle de Facturas
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) = pfMontoMaterias(nEstId)
            'Se aplica descuento sobre total en materias
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Descuento) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) * cPorcentajeDescuento
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalConDescuento) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.Materias) - ArregloDetalleFactura(eColumnasArregloDetalleFactura.Descuento)
            '-------------------------------------------------------------------------------------'
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) = pfMontoMatriculaCredomatic(nEstId)
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro) = pfMontoSeguroPorPagar(nEstId, lPerID)
            ArregloDetalleFactura(eColumnasArregloDetalleFactura.MontoTotal) = ArregloDetalleFactura(eColumnasArregloDetalleFactura.TotalConDescuento) + _
                                                                                ArregloDetalleFactura(eColumnasArregloDetalleFactura.Matricula) + _
                                                                                ArregloDetalleFactura(eColumnasArregloDetalleFactura.Seguro)
            '*******************************************************************

        Catch ex As Exception
            Throw ex
        End Try
        Return ArregloDetalleFactura
    End Function


    'Funcion que obtiene el valor de las materias matriculadas
    Private Function pfMontoMaterias(ByVal nEstId As Integer) As Decimal
        Dim lvalor As Decimal = 0
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        Dim strSql As String = "Select  ISNULL(sum(Precio),0.0) as Precio from Ple_Plan_X_Estudiante P Inner join EST_Estudiante E " & _
                                " On E.id=P.Est_ID " & _
                                " where Temporal = 1 " & _
                                " And E.ID= " & nEstId & "" & _
                                " And P.Periodo_ID=" & lID_Periodo
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lvalor = cmdObtener.ExecuteScalar
            cnnConexion.Close()

        Catch ex As Exception
            Throw ex
        End Try


        Return Round(lvalor, 2, MidpointRounding.AwayFromZero)


    End Function


    'Funcion que determina el pago de seguro de un estudiante
    Private Function pfMontoSeguroPorPagar(ByVal nEstID As Integer, ByRef nPer_ID As Integer) As Decimal
        Dim lSeguro As Decimal = 0
        Dim lValor As Decimal
        Dim strSQL As String = String.Empty

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            strSQL = "Select dbo.fn_ObtenerSeguroEstudiantil (" & nEstID & "," & nPer_ID & ")"

            Dim cmdSeguroEstudiantil As New SqlCommand(strSQL, cnnConexion)
            cnnConexion.Open()
            lSeguro = cmdSeguroEstudiantil.ExecuteScalar
            cnnConexion.Close()
            If lSeguro > 0 Then
                lValor = lSeguro
            Else
                lValor = 0
            End If
            Return lValor
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Funcion que determina si el seguro esta cancelado
    Private Function pfExistePagoSeguro(ByVal nEstID As Integer, ByRef nEnm_ID As Integer) As Byte
        Dim lSeguro As Decimal = 0
        Dim lValor As Decimal
        Dim strSQL As String = String.Empty

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            strSQL = "Select Count(ID) from PSE_PAGO_SEGURO_ESTUDIANTIL where est_id=" & nEstID & "And Enm_ID=" & nEnm_ID
            Dim cmdSeguroEstudiantil As New SqlCommand(strSQL, cnnConexion)
            cnnConexion.Open()
            lSeguro = cmdSeguroEstudiantil.ExecuteScalar
            cnnConexion.Close()
            lValor = lSeguro
            Return lValor
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'Funcion que determina el pago de seguro de un estudiante
    Private Function pfMontoSeguroPorPagarEStId(ByVal nEstId As Integer) As Decimal
        Dim lSeguro As Byte = 0
        Dim lValor As Decimal
        Dim strSQL As String = String.Empty

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            strSQL = "Select ISNULL(count(id),0) as ID " & _
                    " from est_estudiante " & _
                    " where DateAdd(YY, 1, FECHASEGURO) < getdate() " & _
                    " And ID = " & nEstId & ""

            Dim cmdSeguroEstudiantil As New SqlCommand(strSQL, cnnConexion)
            cnnConexion.Open()
            lSeguro = cmdSeguroEstudiantil.ExecuteScalar
            cnnConexion.Close()
            If lSeguro > 0 Then
                lValor = pfPrecioSeguro()
            Else
                lValor = 0
            End If

        Catch ex As Exception
            Throw ex
        End Try


        Return Round(lValor, 2, MidpointRounding.AwayFromZero)

        'Public Function Seguro_estudiantil(ByVal pEst_ID As Long, ByVal pOperacion As Byte)
        '    Dim lRstConsulta As New ADODB.Recordset
        '    Dim tmpArreglo(2) As String
        '    With lRstConsulta
        '        .CacheSize = 1024 ' = 128
        '        .CursorLocation = adUseServer
        '        strSQL = "Exec SP_SEGURO_ESTUDIANTIL " & pEst_ID & "," & pOperacion
        '        .Open(strSQL, Strconexion, adOpenForwardOnly, adLockReadOnly)
        '        If pOperacion = 1 Then
        '            tmpArreglo(1) = LibSistema.FECHASERVIDOR(Strconexion)
        '            tmpArreglo(0) = True
        '        Else
        '            tmpArreglo(0) = lRstConsulta!SEGURO
        '            tmpArreglo(1) = lRstConsulta!Fecha
        '        End If

        '        If .State = 1 Then .Close()
        '    End With
        '    Seguro_estudiantil = tmpArreglo
    End Function


    'Funcion que Determina si el Estudiante debe de cancelar la matricula y el monto
    Private Function pfMontoMatricula(ByVal nEstId As Integer) As Decimal

        Dim lValor As Decimal = 0
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        Dim nGraId As Integer = fnGradoPlan(nEstId)
        'Dim strIdentificacion As String

        Try
            'strIdentificacion = pfObtenerIdentificacion(nEstId)
            If pfEsCursoLibre(nEstId) = False Then

                If fnVerificaPagoMatricula(nEstId) = 0 Then
                    lValor = fnObtieneCostoMatricula(nGraId)
                Else
                    lValor = 0.0

                End If
            End If
            Return Round(lValor, 2, MidpointRounding.AwayFromZero)
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Funcion que Determina si el Estudiante debe de cancelar la matricula y el monto
    Private Function pfMontoMatriculaCredomatic(ByVal nEstId As Integer) As Decimal

        Dim lValor As Decimal = 0
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        Dim nRetorno As Decimal
        'Dim strIdentificacion As String

        Try
            'strIdentificacion = pfObtenerIdentificacion(nEstId)
            If pfEsCursoLibre(nEstId) = False Then
                Dim strSQL As String = "Select top 1 ISNULL(MONTO,0.0) " & _
                            " From Pam_Pago_Matricula P inner Join Est_Estudiante E " & _
                            " On E.ID=P.Est_ID " & _
                            " Where P.Per_Id = " & lID_Periodo

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                Dim cmdObtener As New SqlCommand(strSQL, cnnConexion)
                cnnConexion.Open()
                lValor = cmdObtener.ExecuteScalar
                cnnConexion.Close()

                Return lValor


            End If

            Return Round(nRetorno, 2, MidpointRounding.AwayFromZero)

        Catch ex As Exception
            Throw ex
        End Try
    End Function


    Private Function pfObtenerIdentificacion(ByVal nEstId As Integer) As String

        Dim lValor As Decimal = 0
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()


        Try
            ' strIdentificacion = pfObtenerIdentificacion(nEstId)
            If pfEsCursoLibre(nEstId) = False Then
                Dim strSQL As String = "Select IDENTIFICACION " & _
                            " From Est_Estudiante  " & _
                            " Where ID = " & nEstId
                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                Dim cmdObtener As New SqlCommand(strSQL, cnnConexion)
                cnnConexion.Open()
                lValor = cmdObtener.ExecuteScalar
                cnnConexion.Close()
            End If

            Return Round(lValor, 2, MidpointRounding.AwayFromZero)

        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Funcion que obtiene el Periodo Actual al cobro, el cual es el periodo de Matricula
    Private Function pfObtenerPeriodoActivo() As Integer
        Dim lValor As Integer
        Dim strSql As String = "SELECT PER_ID FROM PAR_PARAMETRO"
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            'lValor = 138
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try

        Return lValor
    End Function

    'Funcion que devuelve el nombre del Periodo Actual
    Private Function pfObtenerNombrePeriodo() As String
        Dim lValor As String
        Dim strSql As String = "SELECT NOMBRE FROM PER_PERIODO WHERE ID=" & pfObtenerPeriodoActivo()
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
            Return lValor
        Catch ex As Exception
            Throw ex

        End Try

    End Function

    'Función que obtiene el Precio de la matricula
    Private Function pfPrecioMatricula() As Decimal
        Dim strSQL As String = "Select top 1 Costo from Com_Costo_Matricula order by fechaModifico desc"
        Dim lValor As Decimal = 0


        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSQL, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
        Return lValor

    End Function


    'Función que obtiene el Precio del Seguro
    Private Function pfPrecioSeguro() As Decimal

        Dim lValor As Integer
        Dim strSql As String = "SELECT MONTOSEGURO FROM PAR_PARAMETRO"
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try

        Return lValor

    End Function

    'Función que determina si el Plan Académico es Curso Libre
    Private Function pfEsCursoLibre(ByVal pId As Integer) As Boolean
        Dim lValor As Integer
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        Dim bRetorno As Boolean

        Dim strSql As String = "SELECT ISNULL(COUNT(P.PLA_ID),0) " & _
                                " FROM PLE_PLAN_X_ESTUDIANTE P INNER JOIN EST_ESTUDIANTE E " & _
                                " ON E.ID=P.EST_ID WHERE PLA_ID=2 AND PERIODO_ID=" & lID_Periodo & _
                                " AND E.ID = " & pId & " " & _
                                " AND CONVALIDADO =0"
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()

            If lValor = 0 Then
                bRetorno = False
            Else
                bRetorno = True
            End If

        Catch ex As Exception
            Throw ex
        End Try

        Return bRetorno
    End Function


    'Verifica si el estudiante ha pagado la matricula al periodo actual
    Private Function fnVerificaPagoMatricula(ByVal nEstID As Integer) As Integer
        Dim lValor As Integer
        Dim nPeriodoActual As Integer = pfObtenerPeriodoActivo()
        Dim strSql As String = "select count(ID) from PAM_PAGO_MATRICULA " & _
                                "where MONTO > 0 AND EST_ID = " & nEstID & " and PER_ID = " & nPeriodoActual & " "
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try

        Return lValor
    End Function



    'Verifica si el estudiante ha pagado la matricula al periodo actual
    Private Function fnVerificaPagoMatriculaXDocumento(ByVal nEstID As Integer, ByVal nIDFactura As Integer) As Integer
        Dim lValor As Integer
        Dim nPeriodoActual As Integer = pfObtenerPeriodoActivo()
        Dim strSql As String = "select count(ID) from PAM_PAGO_MATRICULA " & _
                                "where EST_ID = " & nEstID & " and PER_ID = " & nPeriodoActual & " and DOCUMENTO = " & nIDFactura & " "
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try

        Return lValor
    End Function

    'Obtiene el costo de la matricula

    Private Function fnObtieneCostoMatricula(ByVal nGraID As Integer) As Integer
        Dim lValor As Integer
        Dim nPeriodoActual As Integer = pfObtenerPeriodoActivo()
        Dim strSql As String = "select COSTO from COM_COSTO_MATRICULA " & _
                                "where GRA_ID = " & nGraID & " "
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try

        Return lValor
    End Function

    Private Function fnGradoPlan(ByVal nEstID As Integer) As Integer
        Dim GRADO As Integer = 0
        Dim nPeriodoActual As Integer = pfObtenerPeriodoActivo()
        Try
            Dim sql As String = " select TOP 1 GRA.ID  " & _
                                " from PLE_PLAN_X_ESTUDIANTE PLE " & _
                                " INNER Join " & _
                                " PLA_PLAN PLA " & _
                                " ON PLE.PLA_ID = PLA.ID " & _
                                " INNER Join " & _
                                " GRA_GRADO_ACADEMICO GRA " & _
                                " ON PLA.GRA_ID = GRA.ID  " & _
                                " WHERE(PLE.EST_ID = " & nEstID & " And PLE.PERIODO_ID = " & nPeriodoActual & ") " & _
                                " ORDER BY GRA.ID desc "
            Dim cmdObtener As New SqlCommand(sql, cnnConexion)
            Me.cnnConexion.Open()
            GRADO = cmdObtener.ExecuteScalar
            Me.cnnConexion.Close()
            Return GRADO

        Catch ex As Exception
            Throw ex
        Finally
            Me.cnnConexion.Close()
        End Try
    End Function

    Private Function fnPlanAcademicoanterior(ByVal nEstID As Integer) As Integer
        Dim GRADO As Integer = 0
        Dim nPeriodoActual As Integer = pfObtenerPeriodoActivo()
        Try
            Dim sql As String = " select TOP 1 PLA.ID  " & _
                                " from PLE_PLAN_X_ESTUDIANTE PLE " & _
                                " INNER Join " & _
                                " PLA_PLAN PLA " & _
                                " ON PLE.PLA_ID = PLA.ID " & _
                                " INNER Join " & _
                                " GRA_GRADO_ACADEMICO GRA " & _
                                " ON PLA.GRA_ID = GRA.ID  " & _
                                " WHERE(PLE.EST_ID = " & nEstID & " And PLE.PERIODO_ID = " & nPeriodoActual & ") " & _
                                " ORDER BY GRA.ID desc "
            Dim cmdObtener As New SqlCommand(sql, cnnConexion)
            Me.cnnConexion.Open()
            GRADO = cmdObtener.ExecuteScalar
            Me.cnnConexion.Close()
            Return GRADO

        Catch ex As Exception
            Throw ex
        Finally
            Me.cnnConexion.Close()
        End Try
    End Function
#End Region


    Private Function fnPlanAcademico(ByVal nEstID As Integer) As Integer
        Dim lValor As Integer = 0
        Dim nPeriodoActual As Integer = pfObtenerPeriodoActivo()
        Try
            Dim sql As String = " select DISTINCT TOP 1 PLE.PLA_ID  " & _
                                " from PLE_PLAN_X_ESTUDIANTE PLE " & _
                                " WHERE PLE.EST_ID = " & nEstID & " And PLE.PERIODO_ID = " & nPeriodoActual & _
                                " ORDER BY PLE.PLA_ID desc "
            Dim cmdObtener As New SqlCommand(sql, cnnConexion)
            Me.cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Me.cnnConexion.Close()
            Return lValor

        Catch ex As Exception
            Throw ex
        Finally
            Me.cnnConexion.Close()
        End Try
    End Function
#End Region

#Region "Ejecuta el pago de laFactura"

    Private Function lfActualizaRecibo(ByVal strRecibo As String) As String
        Dim strSql As String
        Try
            strSql = "UPDATE REC_RECIBO SET " & _
                              " OBSERVACIONES = '" & cObservaciones & "' " & _
                              " ,FECHACREO = getDate() " & _
                              " ,FECHAMODIFICO = getDate() " & _
                              " ,TOTAL = " & cTotal & " " & _
                              " ,EFECTIVO = " & cEfectivo & " " & _
                              " ,TARJETA =  " & cTarjeta & " " & _
                              " ,MONTOTRANSACCION = " & cTotalSinDescuento & " " & _
                              " WHERE RECIBO = '" & strRecibo & "'"
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdActualizaAlE As New SqlCommand(strSql, cnnConexion)
            cmdActualizaAlE.Transaction = grefTransaction
            cmdActualizaAlE.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfInsertaBitacoraReversiones(ByVal nENM_ID As Integer, ByVal nEnl_ID As String, ByVal nConvenio As Integer, ByVal nNumComprobante As Integer, ByVal nRecibo As Integer) As String
        Dim strSql As String
        Try
            strSql = "INSERT INTO BR_BITACORA_REVERSIONES " & _
                             "( ENM_ID " & _
                             " ,ENL_ID " & _
                             " ,FECHACREO " & _
                             " ,CREO " & _
                             " ,COD_CONVENIO" & _
                             " ,NUMCOMPROBANTE" & _
                             " ,RECIBO )" & _
                           " VALUES " & _
                                " (" & nENM_ID & "," & _
                                " '" & nEnl_ID & "'," & _
                                " getDate()," & _
                                " '" & cCreo & "', " & _
                                " " & nConvenio & "," & _
                                " " & nNumComprobante & ", " & _
                                " " & nRecibo & ") "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfEliminaAbono(ByVal nRecibo As Integer) As String
        Dim strSql As String
        strSql = "DELETE FROM ALE_ABONO_LETRA WHERE RECIBO = " & nRecibo

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfEliminaPAM(ByVal nIDFactura As Integer, ByVal nEstID As Integer) As String
        Dim strSql As String
        strSql = "DELETE FROM PAM_PAGO_MATRICULA WHERE DOCUMENTO = " & nIDFactura & " AND EST_ID = " & nEstID & ""

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfEliminaPSE(ByVal nIDFactura As Integer) As String
        Dim strSql As String
        strSql = "DELETE FROM PSE_PAGO_SEGURO_ESTUDIANTIL WHERE ENM_ID = " & nIDFactura

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfActualizaFechaSeguroEstudiante(ByVal strIdentificacion As Integer) As Integer
        Dim strSql As String
        Dim lValor As String
        Dim dtFechaDefault As String = "01-01-1800"
        Dim Retorno As Integer
        Dim dtFecha As String
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos2").ToString)
        Try


            strSql = "SELECT ISNULL(convert(varchar,FECHASEGURO,103), '1800-01-01 00:00:00.000') " & _
                                    " FROM EST_ESTUDIANTE " & _
                                    " WHERE IDENTIFICACION = '" & strIdentificacion & "'"
            Dim cmdObtener2 As New SqlCommand(strSql, cnnConexion)
            cnnConexion2.Open()
            lValor = cmdObtener2.ExecuteScalar
            If cnnConexion.State = ConnectionState.Closed Then cnnConexion.Open()
            If lValor <> dtFechaDefault Then
                dtFecha = lValor
            Else
                dtFecha = lValor
            End If
            If dtFecha = Nothing Then
                dtFecha = dtFechaDefault
            End If
            strSql = "UPDATE EST_ESTUDIANTE SET " & _
                         " FECHASEGURO = convert(datetime," & dtFecha & ",103) " & _
                         " WHERE IDENTIFICACION = '" & strIdentificacion & "'"
            Dim cmdActualizaAlE As New SqlCommand(strSql, cnnConexion)
            cmdActualizaAlE.Transaction = grefTransaction
            cmdActualizaAlE.ExecuteNonQuery()
            Retorno = 1
            Return Retorno
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
            cnnConexion.Close()
            cnnConexion2.Close()
        End Try
    End Function

    Private Function lfObtieneAbonoAnterior(ByVal nEnlID As Integer) As DataSet

        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos2").ToString)
        Try

            strSql = "SELECT ID " & _
                     " ,ENL_ID " & _
                     " ,RECIBO " & _
                     " ,DIASACUMULADO " & _
                     " ,TOTALPAGO " & _
                     " ,ABONO " & _
                     " ,SALDO " & _
                     " ,OBSERVACION " & _
                     " ,convert(datetime,FECHA,120) as FECHA " & _
                     " ,DIASMORATORIOS " & _
                     " ,INTERESORDINARIO " & _
                     " ,ABONOINTERESORDINARIO " & _
                     " ,SALDOINTERESORDINARIO " & _
                     " ,INTERESMORATORIO " & _
                     " ,ABONOINTERESMORATORIO " & _
                     " ,SALDOINTERESMORATORIO " & _
                     " FROM ALE_ABONO_LETRA " & _
                     " WHERE ENL_ID = " & nEnlID & " "
            cnnConexion2.Open()
            Dim adpObtener2 As New SqlDataAdapter(strSql, cnnConexion2)
            Dim dstObtener2 As New DataSet
            adpObtener2.Fill(dstObtener2, "Resultado")
            Return dstObtener2
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try

    End Function

    Private Function lfInsertarBitacoraBRA(ByVal ENL_ID As Integer, ByVal dstAbono As DataSet) As String
        Dim strSql As String
        Dim nRetorno As Integer
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try


            If dstAbono.Tables(0).Rows.Count > 0 Then
                Dim dtFechas As DateTime = dstAbono.Tables(0).Rows(0)("FECHA")
                strSql = "INSERT INTO BRA_BITACORA_REVERSIONES_ABONO_LETRA " & _
                                             "( ALE_ID " & _
                                             " ,ENL_ID " & _
                                             " ,RECIBO " & _
                                             " ,DIASACUMULADO " & _
                                             " ,TOTALPAGO " & _
                                             " ,ABONO " & _
                                             " ,SALDO " & _
                                             " ,OBSERVACION " & _
                                             " ,FECHA " & _
                                             " ,DIASMORATORIOS " & _
                                             " ,INTERESORDINARIO " & _
                                             " ,ABONOINTERESORDINARIO " & _
                                             " ,SALDOINTERESORDINARIO " & _
                                             " ,INTERESMORATORIO " & _
                                             " ,ABONOINTERESMORATORIO " & _
                                             " ,SALDOINTERESMORATORIO )" & _
                                                "VALUES ( " & _
                                                 " " & dstAbono.Tables(0).Rows(0)("ID") & " " & _
                                                 " ," & dstAbono.Tables(0).Rows(0)("ENL_ID") & " " & _
                                                 " , '" & dstAbono.Tables(0).Rows(0)("RECIBO").ToString & "' " & _
                                                 " , " & dstAbono.Tables(0).Rows(0)("DIASACUMULADO") & " " & _
                                                 " , " & dstAbono.Tables(0).Rows(0)("TOTALPAGO") & " " & _
                                                 " , " & dstAbono.Tables(0).Rows(0)("ABONO") & " " & _
                                                 " , " & dstAbono.Tables(0).Rows(0)("SALDO") & " " & _
                                                 " , '" & dstAbono.Tables(0).Rows(0)("OBSERVACION").ToString & "' " & _
                                                 " , convert(datetime,'" & dtFechas & "',120) " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("DIASMORATORIOS") & " " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("INTERESORDINARIO") & " " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("ABONOINTERESORDINARIO") & " " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("SALDOINTERESORDINARIO") & " " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("INTERESMORATORIO") & " " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("ABONOINTERESMORATORIO") & " " & _
                                                " , " & dstAbono.Tables(0).Rows(0)("SALDOINTERESMORATORIO") & " " & _
                                                            ")"
                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                cnnConexion.Open()
                Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
                cmdComandoSQL.Transaction = grefTransaction
                cmdComandoSQL.ExecuteNonQuery()
                nRetorno = 1
            Else
                nRetorno = 0
            End If
            Return nRetorno

        Catch ex As Exception
            Throw ex
        Finally
        End Try
    End Function

    Private Function lfObtieneMaxID() As Integer
        Dim lValor As Integer
        Dim strSql As String = "select ISNULL(MAX(ID),0) + 1 from bpb_BITACORA_PAGOS_BANCOS"
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            cnnConexion.Close()

        Catch ex As Exception
            Throw ex
        End Try
        Return lValor
    End Function

    Private Function lfInsertarBitacoraPagosBancos(ByVal nEnl As Integer, ByVal nEnmId As Integer, ByVal nNumComprobante As Integer, ByVal nRecibo As Integer, ByVal nTipoConvenio As Integer, ByVal monto As Decimal, ByVal strCreo As String) As Integer
        Dim strSql As String
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        '  Dim nId As Integer = lfObtieneMaxID()
        Try
            strSql = "INSERT INTO BIT_PAGOS_BANCOS " & _
                     "  (RECIBO " & _
                      " ,ENL_ID " & _
                      " ,ENM_ID " & _
                      " ,TIPOCONVENIO " & _
                      " ,NUMCOMPROBANTE " & _
                      " ,MONTO " & _
                      " ,CREO " & _
                      " ,FECHACREO) " & _
                      " VALUES " & _
                      " ( " & nRecibo & "  " & _
                      " , " & nEnl & " " & _
                      " , " & nEnmId & " " & _
                      " , " & nTipoConvenio & "  " & _
                      " , " & nNumComprobante & " " & _
                      " , " & monto & " " & _
                      " , '" & strCreo & "' " & _
                      " ,getdate()) "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return 1
        Catch ex As Exception
            Throw ex
            Return 0
        Finally
        End Try
    End Function



    Private Function lfActualizaENL(ByVal nMontoRecibo As Double, ByVal nEnlID As Integer, ByVal nNumComprobante As Integer, ByVal nRecibo As Integer) As String
        Dim strSql As String
        Try
            strSql = "UPDATE ENL_ENCABEZADO_LETRA SET " & _
                        " FECHA = (SELECT FECHA from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & ") " & _
                        " ,TOTALINTERES = (SELECT TOTALINTERES from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & ") " & _
                        " ,INTERESES = (SELECT INTERESES from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & ") " & _
                        " ,SALDO = (SELECT SALDO from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " ) " & _
                        " ,TOTALPAGADO = (SELECT TOTALPAGADO from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " ) " & _
                        " ,CANCELADA = (SELECT CANCELADA from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " ) " & _
                        " ,ESTADO = (SELECT ESTADO from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " )  " & _
                        " WHERE (ID = " & nEnlID & ") "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfEliminaAbono(ByVal nMontoRecibo As Double, ByVal nEnlID As Integer, ByVal nNumComprobante As Integer, ByVal nRecibo As Integer) As String
        Dim strSql As String
        Try

            strSql = " DECLARE @FACTURA int, @FECHA datetime,@TOTALPAGO decimal,@ABONOINTERESORDINARIO decimal,@ABONOINTERESMORATORIO decimal " & _
                        "SET @TOTALPAGO = (SELECT TOP 1 TOTALPAGO FROM ALE_ABONO_LETRA ALE INNER JOIN REC_RECIBO rec on ale.RECIBO = rec.RECIBO WHERE ALE.ENL_ID = " & nEnlID & " AND REC.NUMCOMPROBANTE = " & nNumComprobante & ")" & _
                        "SET @ABONOINTERESORDINARIO = (SELECT TOP 1 ABONOINTERESORDINARIO  FROM ALE_ABONO_LETRA ALE INNER JOIN REC_RECIBO rec on ale.RECIBO = rec.RECIBO WHERE ENL_ID = " & nEnlID & " AND REC.NUMCOMPROBANTE = " & nNumComprobante & ")" & _
                        "SET @ABONOINTERESMORATORIO = (SELECT TOP 1 ABONOINTERESMORATORIO FROM ALE_ABONO_LETRA ALE INNER JOIN REC_RECIBO rec on ale.RECIBO = rec.RECIBO WHERE ENL_ID = " & nEnlID & " AND REC.NUMCOMPROBANTE = " & nNumComprobante & ")" & _
                        "SET @FECHA = ( SELECT TOP 1 ISNULL(FECHA,'1800-01-01 00:00:00.000') FROM ALE_ABONO_LETRA ALE INNER JOIN REC_RECIBO rec on ale.RECIBO = rec.RECIBO WHERE ENL_ID = " & nEnlID & " AND REC.NUMCOMPROBANTE = " & nNumComprobante & ")" & _
                        "UPDATE ENL_ENCABEZADO_LETRA SET " & _
                              " FECHA = @FECHA " & _
                              " ,TOTALINTERES = TOTALINTERES - @ABONOINTERESORDINARIO " & _
                              " ,INTERESES = INTERESES - @ABONOINTERESMORATORIO " & _
                              " ,SALDO = (" & nMontoRecibo & " + SALDO) - (@ABONOINTERESORDINARIO - @ABONOINTERESORDINARIO) " & _
                              " ,TOTALPAGADO = TOTALPAGADO - " & nMontoRecibo & " " & _
                              " ,FECHAULTIMOPAGO = @FECHA " & _
                              " ,CANCELADA = 0 " & _
                              " ,ESTADO = 0  " & _
                              " WHERE (ID = " & nEnlID & ") "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function


    Private Function lfEliminaReciboXLetra(ByVal nEnlID As Integer, ByVal nNumComprobante As Integer) As Integer
        Dim strSql As String
        Try

            strSql = " delete from REC_RECIBO where " & _
                       " ID = " & _
                       " (select top 1 rec.id from ALE_ABONO_LETRA ale " & _
                       " inner join REC_RECIBO rec " & _
                       " on ale.RECIBO = rec.RECIBO " & _
                       " where rec.NUMCOMPROBANTE = " & nNumComprobante & "and ale.ENL_ID = " & nEnlID & ")"
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return 1
        Catch ex As Exception
            Throw ex
            Return 1
        Finally
        End Try
    End Function

    Private Function lfObtieneRecibo(ByVal nNUMEROLETRA As String, ByVal nEstID As Integer) As Integer
        Dim lValor As String
        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try

            strSql = "SELECT enl.ID FROM ENM_ENCABEZADO_MATRICULA enm " & _
                                       "INNER JOIN ENL_ENCABEZADO_LETRA enl " & _
                                       "ON enm.ID = enl.ENM_ID " & _
                                       "INNER JOIN EST_ESTUDIANTE est " & _
                                       "ON enl.EST_ID = est.ID " & _
                                       "WHERE (enl.NUMERO = '" & nNUMEROLETRA & "' ) "
            cnnConexion2.Open()
            Dim cmdComando2 As New SqlCommand(strSql, cnnConexion2)
            cmdComando2.Transaction = grefTransaction
            lValor = cmdComando2.ExecuteScalar
            Return lValor

        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try

    End Function

    Private Function lfActualizaPLe(ByVal pPle_ID As Integer, ByVal nEstID As Integer) As String
        Dim strSql As String
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        Try
            strSql = "UPDATE PLE_PLAN_X_ESTUDIANTE SET " & _
                     " CG = '" & cSinDefinir & "' " & _
                     " ,NOTA = " & cSinNota & " " & _
                     " ,ESTADO = '" & cPorMatricular & "'" & _
                     " ,EXAMEN = " & cSinNota & " " & _
                     " ,TEMPORAL =1 " & _
                     " ,EST_TRANSACCION = '" & cSinDefinir & "'" & _
                     " ,ESTADO_CURSO = '" & cSinDefinir & "'" & _
                     " ,NOTA_TEMPORAL = " & cSinNota & "" & _
                     " ,ESTADO_TEMPORAL = 'Por Cursar' " & _
                     " ,ACTIVO = 0 " & _
                     " ,TIPO = 1 " & _
                     " WHERE EST_ID = " & nEstID & " " & _
                     " AND TEMPORAL = 2 " & _
                     " AND PERIODO_ID = " & lID_Periodo & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfActualizaENM(ByVal nNumFactura As Integer, ByVal nEstID As Integer, ByVal strCreo As String) As String
        Dim strSql As String
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        Try
            strSql = "UPDATE ENM_ENCABEZADO_MATRICULA SET " & _
                    "FECHAMODIFICO = getdate() " & _
                              " ,MODIFICO = '" & strCreo & "' " & _
                              " ,CANCELADA = 0 " & _
                              " ,OBSERVACIONES = 'PREFACTURADO' " & _
                              " ,TOTALSINDESCUENTO = " & cTotalSinDescuento & " " & _
                              " ,FACTURADO = " & cCero & " " & _
                              " ,TIPO = 0 " & _
                              " ,MONTOTRANSACCION=0 " & _
                              " ,MONTO_DESCUENTO=0 " & _
                              " WHERE ID = " & nNumFactura & " "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    Private Function lfObtienePleID(ByVal nEnmID As Integer) As DataSet
        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try

            strSql = "select PLE_ID from DEM_DETALLE_MATRICULA " & _
                                       "WHERE (ENM_ID = " & nEnmID & ")"
            cnnConexion2.Open()
            Dim adpObtener As New SqlDataAdapter(strSql, cnnConexion2)
            Dim dstObtener As New DataSet
            adpObtener.Fill(dstObtener, "Resultado")
            Return dstObtener
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try

    End Function

    'Private Function pfEjecutaPagoFactura(ByVal pNumeroFactura As Integer, _
    '                                      ByVal pIdentificacion As String, _
    '                                      ByVal pTipoIdentificacion As String, _
    '                                      ByVal pMontoPagado As Decimal, _
    '                                      ByVal pNumComprobante As Integer, _
    '                                      ByVal pTipoPago As Byte, _
    '                                      ByVal pCodBanco As Integer, _
    '                                      ByVal nNumCheque As Long, _
    '                                      ByVal nNumCuenta As Long, _
    '                                      ByVal nMontoTotal As Decimal) As DataSet
    '    Dim ldtsDatos As New DataSet

    '    Try
    '        pfActualizaMateriasMatriculadas(pIdentificacion)
    '        pfActualizaEstadoFactura(pNumeroFactura)
    '        pfIngresaPagoMatriculaEstudiante(pIdentificacion, pNumeroFactura)
    '        pfIngresaPagoSeguroEstudiante(pIdentificacion, pNumeroFactura)



    '    Catch ex As Exception
    '        Throw ex
    '    End Try

    '    Return ldtsDatos
    'End Function

    Private Function lfObtieneDatosAle(ByVal nEnlID As Integer, ByVal nNumComprobante As Integer) As DataSet

        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try

            strSql = "SELECT ale.ID " & _
                         " ,ale.TOTALPAGO as TOTALPAGO " & _
                         " ,ale.ABONO as ABONO" & _
                         " ,ale.SALDO as SALDO" & _
                         " ,ale.INTERESORDINARIO as INTERESORDINARIO " & _
                         " ,ale.INTERESMORATORIO as INTERESMORATORIO " & _
                         " ,ale.ABONOINTERESMORATORIO as ABONOINTERESMORATORIO " & _
                         " ,enl.INTERESES as INTERES " & _
                         " ,rec.NUMCOMPROBANTE " & _
                         " FROM ALE_ABONO_LETRA ale" & _
                         " INNER JOIN ENL_ENCABEZADO_LETRA enl " & _
                         " ON ale.ENL_ID = enl.ID " & _
                         " INNER JOIN REC_RECIBO REC " & _
                         " ON ale.RECIBO = rec.RECIBO " & _
                         " WHERE ale.ENL_ID = " & nEnlID & " AND rec.NUMCOMPROBANTE = " & nNumComprobante & "  "
            cnnConexion2.Open()
            Dim adpObtener2 As New SqlDataAdapter(strSql, cnnConexion2)
            Dim dstObtener2 As New DataSet
            adpObtener2.Fill(dstObtener2, "Resultado")
            Return dstObtener2
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try

    End Function

    Private Function lfObtieneFechasAle(ByVal nEnlID As Integer) As String
        Dim lValor As String
        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try

            strSql = "SELECT TOP 1 convert(datetime,ISNULL(ale.FECHA,'1800-01-01 00:00:00.000'),120) " & _
                          " FROM ALE_ABONO_LETRA ale" & _
                          " INNER JOIN ENL_ENCABEZADO_LETRA enl " & _
                          " ON ale.ENL_ID = enl.ID " & _
                          " WHERE ENL_ID = " & nEnlID & " " & _
                          " ORDER BY ale.DESC "
            cnnConexion2.Open()
            Dim cmdComando2 As New SqlCommand(strSql, cnnConexion2)
            cmdComando2.Transaction = grefTransaction
            lValor = cmdComando2.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion2.Close()
        End Try

    End Function

    Private Function lfActualizaAle(ByVal nEnlID As Integer, ByVal nNumComprobante As Integer) As Integer
        Dim strSql As String = String.Empty
        Dim nIntMoratorios As Decimal
        Dim nSaldo As Decimal
        Dim dtDatosAle As DataSet
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try
            dtDatosAle = lfObtieneDatosAle(nEnlID, nNumComprobante)
            If dtDatosAle.Tables(0).Rows.Count > 0 Then
                nIntMoratorios = CDec(dtDatosAle.Tables(0).Rows(0)("INTERESMORATORIO")) + CDec(dtDatosAle.Tables(0).Rows(0)("INTERESORDINARIO"))
                nSaldo = CDec(dtDatosAle.Tables(0).Rows(0)("SALDO")) + CDec(dtDatosAle.Tables(0).Rows(0)("ABONO"))
                'strFechaAle = lfObtieneFechasAle(nEnlID)
                ''Si existe un abono anterior al que se desea reversar toma la 
                'If strFechaAle <> "" Then
                '    dtFecha = strFechaAle
                'Else
                '    dtFecha = CStr(Now)
                'End If
                strSql = "UPDATE ALE_ABONO_LETRA SET " & _
                    " TotalPago = " & cCero & "" & _
                    " ,ABONO = " & cCero & "" & _
                    " ,OBSERVACION = 'Recibo Reversado por: BNConectividad'" & _
                    " WHERE ID = (SELECT TOP 1 ALE.ID FROM REC_RECIBO REC INNER JOIN ALE_ABONO_LETRA ALE ON ALE.RECIBO = REC.RECIBO  WHERE ALE.ENL_ID = " & nEnlID & " AND REC.NUMCOMPROBANTE = " & nNumComprobante & ")"
                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                cnnConexion.Open()
                Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
                cmdComandoSQL.Transaction = grefTransaction
                cmdComandoSQL.ExecuteNonQuery()
                Return 1
            Else
                Return 0
            End If
        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try

    End Function

    Private Function lfActualizaRecibo(ByVal nEnlID As Integer, ByVal nNumComprobante As Integer, ByVal strCreo As String) As Integer
        Dim strSql As String = String.Empty
        Dim cnnConexion2 As New SqlConnection(ConfigurationManager.ConnectionStrings("BaseDatos").ToString)
        Try

            strSql = "update rec_recibo " & _
                     "set  " & _
                    "Total = 0 ," & _
                    "Observaciones = 'Recibo reversado por: BN-CONECTIVIDAD' ," & _
                    "Efectivo = 0 ," & _
                    "Tarjeta = 0 ," & _
                    "NumeroTarjeta = 'N/A' ," & _
                    "Autorizacion = 'N/A' ," & _
                    "Descuento = 0 ," & _
                    "Transaccion = 'N/A' ," & _
                    "TotalsinDescuento = 0 ," & _
                    "Emitido = 'N/A' ," & _
                    "Informacion = 'N/A' ," & _
                    "Modifico = '" & strCreo & "' ," & _
                    "FechaModifico = getdate() " & _
                    "WHERE ID = " & _
                    "(select top 1 rec.ID from REC_RECIBO rec " & _
                    "inner join ALE_ABONO_LETRA ale " & _
                    "on rec.RECIBO = ale.RECIBO " & _
                    "where " & _
                    "ale.ENL_ID = " & nEnlID & "  and rec.NUMCOMPROBANTE = " & nNumComprobante & "  )"
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return 1

        Catch ex As Exception
            Throw ex
        Finally
            cnnConexion.Close()
        End Try

    End Function

    Private Function pfActualizaMateriasMatriculadas(ByVal nEstId As Integer, ByVal nTipo As Integer) As Boolean
        Dim strSql As String = String.Empty
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        If nTipo = 1 Then
            strSql = "Update Ple_Plan_X_Estudiante " & _
                     "Set Matriculo='" & cCreo & "', Temporal=2, Estado='" & cEstadoMatriculado & _
                     "' Where Est_Id=" & nEstId & _
                     " and Temporal = 1" & _
                     " and Periodo_ID=" & lID_Periodo
        ElseIf nTipo = 2 Then
            strSql = "Update Ple_Plan_X_Estudiante " & _
                     "Set Matriculo='PAGOS EN LÍNEA', Temporal=2, Estado='" & cEstadoMatriculado & _
                     "',FECHAMATRICULO = getdate() Where Est_Id=" & nEstId & _
                     " and Temporal = 1" & _
                     " and Periodo_ID=" & lID_Periodo
        End If
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            cmdObtener.ExecuteNonQuery()
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function




    'Funcion que registra el pago de la Matricula del Estudiante en el Periodo
    Private Function pfIngresaPagoMatriculaEstudiante(ByVal nEstID As Integer, ByVal pNumeroFactura As Integer, ByVal nMontoMatricula As Decimal) As Boolean
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0

        Try
            Dim lPeriodoID As Integer = pfObtenerPeriodoActivo()
            Dim lMontoMatricula As Decimal = pfMontoMatricula(nEstID)
            Dim lIDFacturas As Integer = pfObtenerIDFactura(pNumeroFactura)

            'strSql = "Select ISNULL(MONTO,0.0) " & _
            '                " From Pam_Pago_Matricula P inner Join Est_Estudiante E " & _
            '                " On E.ID=P.Est_ID " & _
            '                " Where E.ID = " & nEstID & " " & _
            '                " and P.Per_Id = " & lPeriodoID
            'If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            'Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            'cnnConexion.Open()
            'lValor = cmdObtener.ExecuteScalar
            ''Si el estudiante no ha pagado la matricula

            If nMontoMatricula > 0 Then
                strSql = "INSERT INTO PAM_PAGO_MATRICULA " & _
                            "(EST_ID, PER_ID, DOCUMENTO, MONTO) " & _
                        "VALUES (" & nEstID & "," & _
                                lPeriodoID & "," & _
                                lIDFacturas & "," & _
                                lMontoMatricula & ")"

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                Dim cmdIngresaMatricula As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                cmdIngresaMatricula.ExecuteNonQuery()
                cnnConexion.Close()
                Return True

            Else
                Return False
            End If


        Catch ex As Exception
            Throw ex
        End Try

    End Function

    'Funcion que registra y/o actualiza el seguro del Estudiante
    Private Function pfIngresaPagoSeguroEstudiante(ByVal nEstID As Integer, ByVal pNumeroFactura As Integer) As Boolean
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Dim lPerID As Integer = pfObtenerPeriodoActivo()

        Try

            Dim lMontoSeguro As Decimal = pfMontoSeguroPorPagar(nEstID, lPerID)
            Dim lIDFacturas As Integer = pfObtenerIDFactura(pNumeroFactura)

            If lMontoSeguro > 0 Then
                strSql = "INSERT INTO PSE_PAGO_SEGURO_ESTUDIANTIL " & _
                                        "(EST_ID, PER_ID, ENM_ID, MONTO,FECHAPAGO ) " & _
                                    "VALUES (" & nEstID & "," & _
                                            lPerID & "," & _
                                            lIDFacturas & "," & _
                                            lMontoSeguro & "," & _
                                            "Getdate() )"

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                Dim cmdIngresar As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                cmdIngresar.ExecuteNonQuery()
                cnnConexion.Close()

                strSql = "Update Est_estudiante " & _
                       "Set FechaSeguro=GetDate()," & _
                        "Seguro=1" & _
                        "Where ID=" & nEstID
                Dim cmdActualizar As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                cmdActualizar.ExecuteNonQuery()
                cnnConexion.Close()
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try

    End Function

    'Funcion que Elimina el seguro del Estudiante
    Private Function pfEliminaPagoSeguroEstudiante(ByVal nEstID As Integer, ByVal pEnmID As Integer, ByVal nPerID As Integer) As Boolean
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0

        Try
            strSql = "DELETE FROM PSE_PAGO_SEGURO_ESTUDIANTIL WHERE EST_ID = " & nEstID & " and PER_ID = " & nPerID & " and ENM_ID = " & pEnmID

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdIngresar As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            cmdIngresar.ExecuteNonQuery()
            cnnConexion.Close()

            strSql = "Update Est_estudiante " & _
                   "Set FechaSeguro=Convert(Datetime,'01/01/1900',103) ," & _
                    "Seguro= 0 " & _
                    "Where ID= " & nEstID
            Dim cmdActualizar As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            cmdActualizar.ExecuteNonQuery()
            cnnConexion.Close()

            Return True

        Catch ex As Exception
            Throw ex
            Return False
        End Try

    End Function

    'Función que Actualiza el Estado de la factura
    Private Function pfActualizaEstadoFactura(ByVal pNumeroFactura As Integer, _
                                              ByVal nMontoDescuento As Integer, _
                                              ByVal Total As Integer, _
                                              ByVal TotalSinDescuento As Integer, _
                                              ByVal strCreo As String, _
                                              ByVal pstrNumeroTarjeta As String, _
                                              ByVal nTipoTransaccion As Integer, _
                                              ByVal pDescuento As Integer, _
                                              ByVal pTipoCompra As Integer) As Boolean
        Dim strSql As String = String.Empty
        Dim lIDFactura = pfObtenerIDFactura(pNumeroFactura)
        Dim nCXB_ID As Integer = 4
        Dim lPeriodoID As Integer = pfObtenerPeriodoActivo()
        If nTipoTransaccion = 1 Then
            strSql = "Update ENM_ENCABEZADO_MATRICULA " & _
                "Set Cancelada = 1, " & _
                "OBSERVACIONES = 'PAGO CON BNCR-USP CONECTIVIDAD', " & _
                "MODIFICO = '" & strCreo & "', " & _
                "FECHAMODIFICO = getdate(), " & _
                "DESCUENTO =  " & pDescuento & " , " & _
                "TOTAL = " & Total & ", " & _
                "TOTALSINDESCUENTO = " & TotalSinDescuento & ", " & _
                "EFECTIVO = 0 , " & _
                "TARJETA = 0 , " & _
                "MONTO_DESCUENTO = " & nMontoDescuento & ", " & _
                "CXB_ID = " & nCXB_ID & ", " & _
                "PER_ID = " & lPeriodoID & ", " & _
                "MONTOTRANSACCION = " & Total & ", " & _
                "FECHA_DEPOSITO = getdate() " & _
                " Where ID =" & lIDFactura

        ElseIf nTipoTransaccion = 2 Then
            If pTipoCompra = 1 Then
                strSql = "Update ENM_ENCABEZADO_MATRICULA " & _
             "Set Cancelada= 1, " & _
             "OBSERVACIONES = 'PAGO EN LINEA BN', " & _
             "MODIFICO = '" & strCreo & "', " & _
             "FECHAMODIFICO = getdate(), " & _
             "DESCUENTO = " & pDescuento & " , " & _
             "TOTAL = " & Total & ", " & _
             "TOTALSINDESCUENTO = " & TotalSinDescuento & ", " & _
             "NUMEROTARJETA = '" & pstrNumeroTarjeta & "', " & _
             "NOMBRETARJETA = 'CONVENIO BN', " & _
             "MONTO_DESCUENTO = " & nMontoDescuento & ", " & _
             "TARJETA = " & Total & ", " & _
             "PER_ID = " & lPeriodoID & ", " & _
             "MONTOTRANSACCION = 0, " & _
             "CXB_ID = 0 " & _
             " Where ID =" & lIDFactura
            ElseIf pTipoCompra = 2 Then
                strSql = "Update ENM_ENCABEZADO_MATRICULA " & _
             "Set Cancelada= 1, " & _
             "OBSERVACIONES = 'PAGO EN LINEA BN', " & _
             "MODIFICO = '" & strCreo & "', " & _
             "FECHAMODIFICO = getdate(), " & _
             "DESCUENTO = 0 , " & _
             "TOTAL = " & TotalSinDescuento & ", " & _
             "TOTALSINDESCUENTO = 0, " & _
             "NUMEROTARJETA = '" & pstrNumeroTarjeta & "', " & _
             "NOMBRETARJETA = 'CONVENIO BN', " & _
             "MONTO_DESCUENTO = 0, " & _
             "TARJETA = " & TotalSinDescuento & ", " & _
             "PER_ID = " & lPeriodoID & ", " & _
             "MONTOTRANSACCION = 0, " & _
             "CXB_ID = 0 " & _
             " Where ID = " & lIDFactura
            End If
        End If

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            cmdObtener.ExecuteNonQuery()
            cnnConexion.Close()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Funcion que genera el recibo de Pago 
    Private Function pfGeneraReciboPago(ByVal pNumeroLetraCambio As String, ByVal pTipoIdentificacion As String, ByVal nEstID As Integer, ByVal pNumComprobante As Integer, ByVal pMontoPagado As Decimal, ByVal pnCodConvenio As Byte, ByVal nDescuento As Decimal, ByVal nTotalSinDescuento As Decimal, ByVal cCreo As String, ByVal TipoPago As Integer, ByVal strNumeroTarjeta As String) As Integer
        Dim strSql As String = String.Empty
        Dim strMensajeRecibo As String = String.Empty
        Dim lNumeroRecibo As Integer = pfObtenerNumeroRecibo()
        Dim lIDLetraCambio As Integer = pfObtenerIDLetraCambio(pNumeroLetraCambio)
        Dim nRetorno As Integer = 0

        Select Case pnCodConvenio
            Case eCodigoConvenio.LetraCambio
                strMensajeRecibo = "Recibo que cancela LETRA DE CAMBIO ID#: " & lIDLetraCambio
            Case eCodigoConvenio.Facturas
                strMensajeRecibo = "Recibo cancelado con el comprobante#: " & lNumeroRecibo
            Case eCodigoConvenio.CursosLibres

            Case eCodigoConvenio.Carnets

        End Select

        Try

            'SI ES UNA LETRA DE CAMBIO CON CALCULO DE INTERESES
            If TipoPago = 1 Then

                strSql = "INSERT INTO REC_RECIBO " & _
                            "(RECIBO, " & _
                            "OBSERVACIONES, " & _
                            "CREO, " & _
                            "MODIFICO, " & _
                            "FECHACREO, " & _
                            "FECHAMODIFICO, " & _
                            "TOTAL,  " & _
                            "EFECTIVO, " & _
                            "TARJETA, " & _
                            "NUMEROTARJETA,  " & _
                            "NOMBRETARJETA, " & _
                            "AUTORIZACION, " & _
                            "DESCUENTO, " & _
                            "TOTALSINDESCUENTO, " & _
                            "TRANSACCION, " & _
                            "EMITIDO, " & _
                            "INFORMACION, " & _
                            "MONTOTRANSACCION, " & _
                            "FACTURAANOMBRE, " & _
                            "EST_ID, " & _
                            "NUMCOMPROBANTE, " & _
                            "MONTO_DESCUENTO, " & _
                            "FECHA_DEPOSITO,  " & _
                            "CXB_ID) " & _
                        "VALUES (" & _
                            lNumeroRecibo & ", " & _
                            "'" & strMensajeRecibo & "', '" & _
                            cCreo & "', '" & _
                            cCreo & "', " & _
                            "GetDate(), " & _
                            "GetDate(), " & _
                            pMontoPagado & ", " & _
                            " 0," & _
                            " 0," & _
                            "'N/A'," & _
                            "'N/A'," & _
                            pNumComprobante & ", " & _
                            nDescuento & ", " & _
                            nTotalSinDescuento & ", " & _
                            lNumeroRecibo & ", " & _
                            "'Recibo Generado mediante el Convenio USP-BN', " & _
                            "'Recibo Generado mediante el Convenio USP-BN', " & _
                            pMontoPagado & " , '" & _
                           lfNombreEstudiante(pTipoIdentificacion, nEstID) & "', " & _
                           nEstID & " , " & _
                           pNumComprobante & _
                          " , 0, " & _
                          " getdate(),  " & _
                          " 4)"

            ElseIf TipoPago = 2 Then
                strSql = "INSERT INTO REC_RECIBO " & _
                            "(RECIBO, " & _
                            "OBSERVACIONES, " & _
                            "CREO, " & _
                            "MODIFICO, " & _
                            "FECHACREO, " & _
                            "FECHAMODIFICO, " & _
                            "TOTAL,  " & _
                            "EFECTIVO, " & _
                            "TARJETA, " & _
                            "NUMEROTARJETA,  " & _
                            "NOMBRETARJETA, " & _
                            "AUTORIZACION, " & _
                            "DESCUENTO, " & _
                            "TOTALSINDESCUENTO, " & _
                            "TRANSACCION, " & _
                            "EMITIDO, " & _
                            "INFORMACION, " & _
                            "MONTOTRANSACCION, " & _
                            "FACTURAANOMBRE, " & _
                            "EST_ID, " & _
                            "NUMCOMPROBANTE, " & _
                            "MONTO_DESCUENTO, " & _
                            "CXB_ID) " & _
                        "VALUES (" & _
                            lNumeroRecibo & ", " & _
                            "'" & strMensajeRecibo & "', '" & _
                            cCreo & "', '" & _
                            cCreo & "', " & _
                            "GetDate(), " & _
                            "GetDate(), " & _
                            pMontoPagado & ", " & _
                            " 0, " & _
                            pMontoPagado & " ," & _
                            "'" & strNumeroTarjeta & "', " & _
                            "'CONVENIO BN', " & _
                            pNumComprobante & ", " & _
                            nDescuento & ", " & _
                            nTotalSinDescuento & ", " & _
                            lNumeroRecibo & ", " & _
                            "'Recibo Generado mediante pagos en línea', " & _
                            "'', " & _
                            "0 , '" & _
                           lfNombreEstudiante(pTipoIdentificacion, nEstID) & "', " & _
                           nEstID & " , " & _
                            pNumComprobante & ", " & _
                          " 0, " & _
                          " 0)"


            End If

            'Ejecuta el ingreso del recibo de pago
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            cmdObtener.ExecuteNonQuery()
            cnnConexion.Close()

            Return lNumeroRecibo

            'Select Case pnCodConvenio
            '    Case eCodigoConvenio.LetraCambio
            '        If pfCancelaLetradeCambio(pNumeroLetraCambio, pMontoPagado, lNumeroRecibo) Then
            '            nRetorno = lNumeroRecibo
            '        Else

            '            nRetorno = 0
            '        End If
            '    Case eCodigoConvenio.Facturas
            '        nRetorno = lNumeroRecibo

            '    Case eCodigoConvenio.CursosLibres

            '    Case eCodigoConvenio.Carnets

            'End Select

        Catch ex As Exception
            Dim lError As Boolean
            lError = lfBitacoraErrores("WS-CONECTIVIDAD", "pfGeneraReciboPago", Err.Number, Err.Description)
            Return 0
            Throw ex
            Return 0

        End Try
    End Function

    'Función que realiza la cancelación de la letra de cambio
    Private Function pfCancelaLetradeCambio(ByVal pNumeroLetraCambio As String, ByVal pMontoPagado As Decimal, ByVal pReciboPago As Integer) As Boolean

        Dim strSql As String = String.Empty

        lValoresLetraCambio = pfDetalleLetraCambio(pNumeroLetraCambio)
        Dim lTipoLetraCambio As Byte = lValoresLetraCambio(eColumnasArregloLetras.TipoLetraCambio)
        Dim lIDLetraCambio As Integer = lValoresLetraCambio(eColumnasArregloLetras.IdLetraCambio)
        Dim lRecibo As Integer = pReciboPago
        Dim lDiasAcumuados As Integer = lValoresLetraCambio(eColumnasArregloLetras.DiasOrdinarios)
        Dim lTotalPago As Decimal = lValoresLetraCambio(eColumnasArregloLetras.MontoTotal)
        Dim lAbono As Decimal = pMontoPagado
        Dim lSaldo As Decimal = lValoresLetraCambio(eColumnasArregloLetras.SaldoLetra)
        Dim lObservaciones As String = String.Empty
        Dim lFecha As Date = Date.Now
        Dim lDiasMoratorios As Integer = lValoresLetraCambio(eColumnasArregloLetras.DiasMoratorios)
        Dim lInteresOrdinario As Decimal = lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario)
        Dim lAbonoInteresOrdinario As Decimal = 0
        Dim lSaldoInteresOrdinario As Decimal = 0
        Dim lInteresMoratorio As Decimal = lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio)
        Dim lAbonoInteresMoratorio As Decimal = 0
        Dim lSaldoInteresMoratorio As Decimal = 0
        Dim lSaldoRemante As Decimal = 0
        Dim lEstadoLetra As Byte = 0
        Dim lFechaUltimoPago As String = pfFechaCalculoIntereses(pNumeroLetraCambio)


        Try

            'Las letras de Bach-Lic tienen cálculo de intereses
            If lTipoLetraCambio = eTipoLetraCambio.BachLic Then

                'El Saldo remanente es igual al monto del abono
                lSaldoRemante = lAbono

                'INICIA CALCULANDO EL INTERES MORATORIO
                'SI ES SUFICIENTE CANCELA EL TOTAL DEL INTERES MORATORIO
                If lSaldoRemante >= lInteresMoratorio Then
                    lAbonoInteresMoratorio = lInteresMoratorio
                    lSaldoInteresMoratorio = 0
                    lSaldoRemante = lSaldoRemante - lInteresMoratorio
                Else
                    lAbonoInteresMoratorio = lSaldoRemante
                    lSaldoInteresMoratorio = lInteresMoratorio - lAbonoInteresMoratorio
                    lSaldoRemante = 0
                End If
                'SI ES SUFICIENTE CANCELA EL TOTAL DEL INTERES ORDINARIO
                If lSaldoRemante >= lInteresOrdinario Then
                    lAbonoInteresOrdinario = lInteresOrdinario
                    lSaldoInteresOrdinario = 0
                    lSaldoRemante = lSaldoRemante - lInteresOrdinario
                Else
                    lAbonoInteresOrdinario = lSaldoRemante
                    lSaldoInteresOrdinario = lInteresOrdinario - lSaldoRemante
                    lSaldoRemante = 0
                End If
                'DEL SALDO DEL ABONO, AMORTIZA EL CAPITAL
                If lSaldoRemante >= lSaldo Then
                    lSaldo = 0
                Else
                    lSaldo = lSaldo - lSaldoRemante
                End If

                'INGRESA EL REGISTRO A LA TABLA DE ABONOS DE LETRA DE CAMBIO.

                strSql = "INSERT INTO ALE_ABONO_LETRA " & _
                               "([ENL_ID] " & _
                               ",[RECIBO] " & _
                               ",[DIASACUMULADO] " & _
                               ",[TOTALPAGO]  " & _
                               ",[ABONO] " & _
                               ",[SALDO] " & _
                               ",[OBSERVACION] " & _
                               ",[FECHA] " & _
                               ",[DIASMORATORIOS] " & _
                               ",[INTERESORDINARIO] " & _
                               ",[ABONOINTERESORDINARIO] " & _
                               ",[SALDOINTERESORDINARIO] " & _
                               ",[INTERESMORATORIO] " & _
                               ",[ABONOINTERESMORATORIO] " & _
                               ",[SALDOINTERESMORATORIO]) " & _
                        "VALUES " & _
                                "(" & lIDLetraCambio & _
                                "," & lRecibo & _
                                "," & lDiasAcumuados & _
                                "," & lTotalPago & _
                                "," & lAbono & _
                                "," & lSaldo & _
                                ", 'CANCELA LETRA DE CAMBIO MEDIANTE CONVENIO USP-BN'" & _
                                ", GETDATE()" & _
                                "," & lDiasMoratorios & _
                                "," & lInteresOrdinario & _
                                "," & lAbonoInteresOrdinario & _
                                "," & lSaldoInteresOrdinario & _
                                "," & lInteresMoratorio & _
                                "," & lAbonoInteresMoratorio & _
                                "," & lSaldoInteresMoratorio & ")"


                'strSql = "INSERT INTO ALE_ABONO_LETRA " & _
                '            "(ENL_ID,RECIBO,DIASACUMULADO,TOTALPAGO,ABONO,SALDO,OBSERVACION " & _
                '            ",FECHA,DIASMORATORIOS,INTERESORDINARIO,ABONOINTERESORDINARIO " & _
                '            ",SALDOINTERESORDINARIO,INTERESMORATORIO,ABONOINTERESMORATORIO " & _
                '            ",SALDOINTERESMORATORIO) " & _
                '        "VALUES(" & _
                '            lIDLetraCambio & "," & _
                '            lRecibo & "," & _
                '            lDiasAcumuados & "," & _
                '            lTotalPago & "," & _
                '            lAbono & "," & _
                '            lSaldo & "," & _
                '            "'CANCELA LETRA DE CAMBIO MEDIANTE CONVENIO USP-BN'" & "," & _
                '            "getdate(), " & _
                '            lDiasMoratorios & "," & _
                '            lInteresOrdinario & "," & _
                '            lAbonoInteresOrdinario & "," & _
                '            lSaldoInteresOrdinario & "," & _
                '            lInteresMoratorio & "," & _
                '            lAbonoInteresMoratorio & "," & _
                '            lSaldoInteresMoratorio & ")"

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                cmdObtener.ExecuteNonQuery()
                cnnConexion.Close()


                'ACTUALIZA LOS VALORES EN EL ENCABEZADO DE LA LETRA
                If lSaldo <= 0 Then lEstadoLetra = eEstadoLetra.Cancelada

                Dim dFechaPago As Date = lFechaUltimoPago

                strSql = "Update ENL_ENCABEZADO_LETRA " & _
                        "Set FechaultimoPago = convert(datetime,'" & dFechaPago.Date & "',103) " & _
                        ",Fecha= Getdate() " & _
                ",Saldo=" & lSaldo & _
                ",TotalPagado=" & (lValoresLetraCambio(eColumnasArregloLetras.TotalPagado) + lAbono) & _
                ",TotalInteres=" & (lValoresLetraCambio(eColumnasArregloLetras.TotalInteresOrdinario) + lInteresOrdinario) & _
                ",TotalInteresMoratorios=" & (lValoresLetraCambio(eColumnasArregloLetras.TotalInteresMoratorio) + lInteresMoratorio) & _
                ",Cancelada =" & lEstadoLetra & _
                " Where NUMERO = '" & pNumeroLetraCambio & "'"

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()

                cnnConexion.Open()

                Dim cmdActualizar As New SqlCommand(strSql, cnnConexion)
                cmdActualizar.ExecuteNonQuery()
                cnnConexion.Close()


                'Las letras de Cambio de Maestria no tienen Intereses
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            Dim lError As Boolean
            lError = lfBitacoraErrores("WS-CONECTIVIDAD", "pfCancelaLetradeCambio", Err.Number, Err.Description)
            Return False
        End Try
    End Function

    'Obtiene Maximo Numero de Recibo
    Private Function lfMaxRecibo() As Integer
        Dim lValor As Integer
        Dim strSql As String = "select MAX(RECIBO) + 1 from REC_RECIBO"
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            'lValor = 138
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try

        Return lValor
    End Function


#End Region


#End Region


#Region "Calculos de Letras de Cambio"
    'Funcion para el calculo de los interes de una letra de cambio Devuelve el resultado de en un arreglo del tipo Decimal
    Private Function pfDetalleLetraCambio(ByVal NumeroLetraCambio As String) As Decimal()

        Dim ldst_LetrasCambio As New DataSet
        Dim LonLetraId As Long
        Dim FechaInicioCalculo As Date
        Dim FechaPago As Date
        Dim FechaAbono As Date
        Dim SaldoLetra As Decimal
        Dim strSql As String = String.Empty
        Dim lContador As Integer = 0
        Dim lintTipoLetaCambio As Integer = -1
        Dim lintTipoInteresMoratorio As Integer = 0
        Dim lstrNumeroLetra As String = String.Empty
        Dim lintDias As Integer = 0
        Dim lMontoInteresOrdinario As Decimal = 0.0
        Dim lMontoInteresMoratorio As Decimal = 0.0
        Dim lMontoSaldoLetra As Decimal = 0.0
        Dim lMontoTotalPagar As Decimal = 0.0
        Dim lArregloLetras(10) As Decimal


        'Inicializa el Arreglo de letrras
        lArregloLetras(eColumnasArregloLetras.SaldoLetra) = 0
        lArregloLetras(eColumnasArregloLetras.MontoInteresOrdinario) = 0
        lArregloLetras(eColumnasArregloLetras.MontoInteresMoratorio) = 0
        lArregloLetras(eColumnasArregloLetras.MontoTotal) = 0
        lArregloLetras(eColumnasArregloLetras.DiasOrdinarios) = 0
        lArregloLetras(eColumnasArregloLetras.DiasMoratorios) = 0
        lArregloLetras(eColumnasArregloLetras.TipoLetraCambio) = 0
        lArregloLetras(eColumnasArregloLetras.IdLetraCambio) = 0
        lArregloLetras(eColumnasArregloLetras.TotalInteresOrdinario) = 0
        lArregloLetras(eColumnasArregloLetras.TotalInteresMoratorio) = 0
        lArregloLetras(eColumnasArregloLetras.TotalPagado) = 0


        '**********************************

        Try
            strSql = "SELECT L.ID,NUMERO AS NUMERO_LETRA, " & _
                    " CASE WHEN MONTOLETRA =0 THEN TOTALPAGAR ELSE MONTOLETRA END AS [MONTO_LETRA], " & _
                    "L.Tipo,SALDO AS [Saldo_Letra], " & _
                    "CASE WHEN L.FECHAINCLUSION > L.FECHA THEN L.FECHA ELSE L.FECHAINCLUSION END AS FECHA_LETRA, " & _
                    "FECHAFINAL AS FECHA_PAGO, " & _
                    "TIPOINTERESMORATORIO, NOMBRE, " & _
                    "TOTALPAGADO,TOTALINTERES,TOTALINTERESMORATORIOS " & _
                    "FROM ENL_ENCABEZADO_LETRA L INNER JOIN ENM_ENCABEZADO_MATRICULA M " & _
                    "ON M.ID=L.ENM_ID " & _
                    "WHERE L.CANCELADA=0 " & _
                    "AND NUMERO = '" & NumeroLetraCambio & "'"

            Dim adpObtener As New SqlDataAdapter(strSql, cnnConexion)
            adpObtener.Fill(ldst_LetrasCambio, "Resultado")


            'Si la consulta de letras devuelve información se continúa con el proceso de calculo

            If ldst_LetrasCambio.Tables("Resultado").Rows.Count > 0 Then

                LonLetraId = ldst_LetrasCambio.Tables("Resultado").Rows(0)("ID") ' ID DE LA LETRA DE CAMBIO
                lstrNumeroLetra = ldst_LetrasCambio.Tables("Resultado").Rows(0)("NUMERO_LETRA") ' NUMERO DE LA LETRA DE CAMBIO
                lintTipoLetaCambio = ldst_LetrasCambio.Tables("Resultado").Rows(0)("TIPO") 'TIPO DE LETRA DE CAMBIO
                lintTipoInteresMoratorio = ldst_LetrasCambio.Tables("Resultado").Rows(0)("TipoInteresMoratorio") 'TIPO DE INTERES MORATORIO
                SaldoLetra = Round(ldst_LetrasCambio.Tables("Resultado").Rows(0)("Saldo_Letra"), 2, MidpointRounding.AwayFromZero)  'SALDO DE CAPITAL DE LA LETRA
                'FechaInicioCalculo = ldst_LetrasCambio.Tables("Resultado").Rows(0)("FECHA_LETRA") 'FECHA DE LA LETRA
                FechaPago = ldst_LetrasCambio.Tables("Resultado").Rows(0)("FECHA_PAGO") 'FECHA DE VENCIMIENTO DE LA LETRA
                FechaAbono = Date.Now 'OBTIENE LA FECHA DE HOY
                'Carga la función de Fecha del último pago realizado
                FechaInicioCalculo = CDate(pfFechaCalculoIntereses(lstrNumeroLetra))

                'Dependiendo del Tipo de Letra así realiza los cálculos de los saldos
                If lintTipoLetaCambio = eTipoLetraCambio.BachLic Then
                    '***********************************************
                    'LETRAS CON INTERESES 
                    '***********************************************

                    '1. DETERMINA SI HAY INTERESES ORDINARIOS
                    lintDias = DateDiff("d", FechaInicioCalculo, FechaAbono)

                    'Basados en la Fecha de Vencimiento de la Letra y  con la fecha del abono determina si
                    'la letra está vencida

                    If FechaAbono <= FechaPago Then
                        'SI LA FECHA DE PAGO ES POSTERIOR AL VENCIMIENTO DE LA LETRA
                        'Si la letra está vencida y existieron intereses ordinarios, toma la fecha de vencimiento
                        'para el calculo del interes moratorio
                        lMontoInteresOrdinario = pfCalculoInteresOrdinario(FechaInicioCalculo, FechaAbono, SaldoLetra) _
                                                         + pfSaldosInteresAbonos(LonLetraId, eTipoInteres.Corrientes)
                        lArregloLetras(eColumnasArregloLetras.DiasOrdinarios) = lintDias
                    Else
                        lintDias = DateDiff("d", FechaInicioCalculo, FechaPago, Microsoft.VisualBasic.FirstDayOfWeek.Sunday)


                        If lintDias > 0 Then
                            lMontoInteresOrdinario = pfCalculoInteresOrdinario(FechaInicioCalculo, FechaPago, SaldoLetra) _
                                                                                     + pfSaldosInteresAbonos(LonLetraId, eTipoInteres.Corrientes)

                        Else
                            lintDias = 0
                            lMontoInteresOrdinario = pfSaldosInteresAbonos(LonLetraId, eTipoInteres.Corrientes)


                        End If
                        lArregloLetras(eColumnasArregloLetras.DiasOrdinarios) = lintDias

                        If FechaPago > FechaInicioCalculo Then FechaInicioCalculo = FechaPago
                        lintDias = DateDiff("d", FechaInicioCalculo, FechaAbono)
                        If (lintDias) > 365 Then lintDias = 365
                        lArregloLetras(eColumnasArregloLetras.DiasMoratorios) = lintDias
                        lMontoInteresMoratorio = pfCalculoInteresMoratorio(FechaInicioCalculo, FechaAbono, SaldoLetra, lintTipoInteresMoratorio) _
                                        + pfSaldosInteresAbonos(LonLetraId, eTipoInteres.Moratorios) 'Funcion de Suma de Intereses Moratorios Cancelados
                    End If

                Else
                    '***********************************************
                    'LETRAS SIN INTERESES 
                    '***********************************************
                    lMontoInteresOrdinario = 0.0
                    lMontoInteresMoratorio = 0.0

                End If

                lMontoTotalPagar = SaldoLetra + Round(lMontoInteresOrdinario, 2, MidpointRounding.AwayFromZero) + Round(lMontoInteresMoratorio, 2, MidpointRounding.AwayFromZero)

                'Asigna los valores al Arreglo de letrras
                lArregloLetras(eColumnasArregloLetras.SaldoLetra) = SaldoLetra
                lArregloLetras(eColumnasArregloLetras.MontoInteresOrdinario) = lMontoInteresOrdinario
                lArregloLetras(eColumnasArregloLetras.MontoInteresMoratorio) = lMontoInteresMoratorio
                lArregloLetras(eColumnasArregloLetras.MontoTotal) = lMontoTotalPagar
                lArregloLetras(eColumnasArregloLetras.IdLetraCambio) = LonLetraId
                lArregloLetras(eColumnasArregloLetras.TipoLetraCambio) = lintTipoLetaCambio
                lArregloLetras(eColumnasArregloLetras.TotalInteresOrdinario) = ldst_LetrasCambio.Tables("Resultado").Rows(0)("TOTALINTERES")
                lArregloLetras(eColumnasArregloLetras.TotalInteresMoratorio) = ldst_LetrasCambio.Tables("Resultado").Rows(0)("TOTALINTERESMORATORIOS")
                lArregloLetras(eColumnasArregloLetras.TotalPagado) = ldst_LetrasCambio.Tables("Resultado").Rows(0)("TOTALPAGADO")

                '**********************************
            End If

        Catch ex As Exception
            Dim lError As Boolean
            lError = lfBitacoraErrores("WS-CONECTIVIDAD", "pfDetalleLetradeCambio", Err.Number, Err.Description)
        End Try

        Return lArregloLetras

    End Function

    'Devuelve la fecha del último pago realizado a una letra de Bach -Li
    Private Function pfFechaCalculoIntereses(ByVal pNumeroLetra As String) As Date
        Dim ldtsDatos As New DataSet
        Dim strSQL As String = String.Empty
        Dim lvalor, lFechaAbono As Date


        Try
            'Carga la Letra de cambio
            strSQL = "Select ID, Fecha as Fecha " & _
                    "From ENL_encabezado_Letra " & _
                    "Where numero= '" & pNumeroLetra & "'"
            Dim adpObtener As New SqlDataAdapter(strSQL, cnnConexion)
            adpObtener.Fill(ldtsDatos, "Resultado")
            lvalor = ldtsDatos.Tables("Resultado").Rows(0)("Fecha")

            'Busca si la letra de cambio presentó abonos.
            strSQL = "Select isNull(max(fecha),'1800-01-01') as Fecha " & _
                                "From ALE_Abono_letra " & _
                                "Where Abono > 0 and enl_id =" & ldtsDatos.Tables("Resultado").Rows(0)("ID")

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSQL, cnnConexion)
            cnnConexion.Open()
            lFechaAbono = cmdObtener.ExecuteScalar


            If lFechaAbono > lvalor Then
                Return lFechaAbono
            Else
                Return lvalor
            End If

            cnnConexion.Close()

        Catch ex As Exception
        Finally
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
        End Try


    End Function

    'Funcion que calcula el monto de Intereses Ordinarios
    Private Function pfCalculoInteresOrdinario(ByVal pFechaInicio As Date, ByVal pFechaFinal As Date, ByVal pMontoCapital As Decimal) As Decimal
        Dim lDias As Integer
        Dim lvalor As Decimal

        Try
            lDias = DateDiff("d", pFechaInicio, pFechaFinal)
            If lDias <= 0 Then
                lvalor = 0.0
            Else
                lvalor = pMontoCapital * lDias * cInteresOrdinarioDiario
            End If
        Catch ex As Exception

        End Try


        Return Round(lvalor, 2, MidpointRounding.AwayFromZero)
    End Function

    'Función que determina el monto de interés moratorio, para un máximo de 365 días
    Private Function pfCalculoInteresMoratorio(ByVal pFechaInicio As Date, ByVal pFechaFinal As Date, ByVal pMontoCapital As Decimal, ByVal pTipo As Integer) As Decimal
        Dim lDias As Integer
        Dim lValor As Decimal = 0

        Try
            lDias = DateDiff("d", pFechaInicio, pFechaFinal)
            If lDias > 365 Then lDias = 365

            'Las letas anteriores que estan pactadas a un 5% las calcula con este interés
            If pTipo = eTipoInteresMoratorio.Interes5 Then
                lValor = pMontoCapital * lDias * cInteresMoratorioDiario_5
            Else
                lValor = pMontoCapital * lDias * cInteresMoratorioDiario_3_9
            End If


        Catch ex As Exception

        Finally

        End Try
        Return Round(lValor, 2, MidpointRounding.AwayFromZero)

    End Function

    'Función que deermina si existen saldos de intereses no cancelados
    Private Function pfSaldosInteresAbonos(ByVal pEnl_ID As Long, ByVal TipoInteres As Byte) As Decimal
        Dim lSaldoInteresOrdinario As Decimal = 0
        Dim lSaldoInteresMoratorio As Decimal = 0
        Dim strSql As String = String.Empty
        Dim ldtsDatos As New DataSet
        Dim lvalor As Decimal = 0

        strSql = "Select TOP 1 ID,SALDOINTERESORDINARIO,SALDOINTERESMORATORIO " & _
                    "From ALE_ABONO_LETRA WHERE ENL_ID=" & pEnl_ID & " ORDER BY ID DESC"
        Dim adpObtener As New SqlDataAdapter(strSql, cnnConexion)
        adpObtener.Fill(ldtsDatos, "Resultado")

        If ldtsDatos.Tables("Resultado").Rows.Count > 0 Then
            lSaldoInteresOrdinario = ldtsDatos.Tables("Resultado").Rows(0)("SaldoInteresOrdinario")
            lSaldoInteresMoratorio = ldtsDatos.Tables("Resultado").Rows(0)("SaldoInteresMoratorio")

            If TipoInteres = eTipoInteres.Corrientes Then

                If lSaldoInteresOrdinario > 0 Then
                    lvalor = lSaldoInteresOrdinario
                Else
                    lvalor = 0
                End If
            Else
                If lSaldoInteresMoratorio > 0 Then
                    lvalor = lSaldoInteresMoratorio
                Else
                    lvalor = 0
                End If
            End If
        Else

        End If

        Return Round(lvalor, 2, MidpointRounding.AwayFromZero)

    End Function

    'Función que obtiene el tipo de letra de cambio
    Private Function pfTipoLetraCambio(ByVal pNumeroLetraCambio As String) As Byte
        Dim lValor As Byte = 0
        Dim strSql As String = String.Empty
        Try
            strSql = "Select TIPO " & _
                    "from ENL_ENCABEZADO_LETRA " & _
                    "Where NUMERO ='" & pNumeroLetraCambio & "'"
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'Función que obtiene el ID de la Letra de Cambio
    Private Function pfObtenerIDLetraCambio(ByVal pNumeroLetraCambio As String) As Integer
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try
            strSql = "Select ID " & _
                    "from ENL_ENCABEZADO_LETRA " & _
                    "Where NUMERO = '" & pNumeroLetraCambio & "'"


            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar

            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#End Region




    'Obtener Tipo Cambio Diego
    Private Function pfObtenerTipoCambio() As Integer
        Dim strSql As String = String.Empty
        Dim lValor As Integer = 0
        Try

            'select * from [vis_tipo_cambio_best]
            strSql = "Select ge_tcv_val_compra from vis_tipo_cambio_best"


            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar

            Return lValor
            cnnConexion.Close()
        Catch ex As Exception
            Throw ex
        End Try
    End Function

#Region "Interfaz BEST"
    Private Function lfInsertaInterfazEncabezado(ByVal bs_ine_est_id As String, ByVal bs_ine_cod_documento As String, ByVal bs_ine_ind_tipo_documento As String, ByVal bs_ine_mto_documento As Decimal, ByVal bs_ine_mto_matricula As Decimal, ByVal bs_ine_mto_seguro As Decimal, ByVal bs_ine_mto_materias As Decimal, ByVal bs_ine_mto_int_ordinario As Decimal, ByVal bs_ine_mto_int_moratorio As Decimal, ByVal bs_ine_mto_impuesto As Decimal, ByVal bs_ine_mto_total_factura As Decimal, ByVal bs_ine_val_porc_descuento As Decimal, ByVal bs_ine_mto_descuento As Decimal, ByVal bs_ine_val_porc_exoneracion As Decimal, ByVal bs_ine_mto_exoneracion As Decimal, ByVal bs_ine_mto_total_deduccion As Decimal, ByVal bs_ine_mto_total_neto As Decimal, ByVal bs_ine_mto_efectivo As Decimal, ByVal bs_ine_mto_letra_cambio As Decimal, ByVal bs_ine_numero_letra_cambio As String, ByVal bs_ine_mto_transaccion As Decimal, ByVal bs_ine_cod_transaccion As String, ByVal bs_ine_cod_cta_corriente As String, ByVal bs_ine_mto_tarjeta As Decimal, ByVal bs_ine_cod_autorizacion As String, ByVal bs_ine_cod_cta_tarjeta As String, ByVal bs_ine_mto_total_forma_pago As Decimal, ByVal bs_ine_cod_usr_creacion As String, ByVal bs_ine_mto_total_vuelto As Decimal, ByVal bs_ine_cod_usr_aplica As String, ByVal bs_ine_mto_dinero_estudiante As Decimal, ByVal bs_ine_forma_pago As Integer) As Boolean


        Dim strSql As String
        Dim nPerId As Integer
        Dim nTipoCambio As Integer

        Try
            nPerId = pfObtenerPeriodoActivo()
            nTipoCambio = pfObtenerTipoCambio()
            strSql = "INSERT INTO bs_interfaz_encabezado " & _
               "(bs_ine_cod_compania " & _
               ",bs_ine_per_id " & _
               ",bs_ine_est_id " & _
               ",bs_ine_fec_documento " & _
              ",bs_ine_ind_tipo_documento " & _
              ",bs_ine_cod_documento " & _
              ",bs_ine_cod_moneda " & _
              ",bs_ine_val_tipo_cambio " & _
              ",bs_ine_ind_estado " & _
              ",bs_ine_num_asiento " & _
              ",bs_ine_mto_documento " & _
              ",bs_ine_mto_matricula " & _
              ",bs_ine_mto_seguro " & _
              ",bs_ine_mto_materias " & _
              ",bs_ine_mto_int_ordinario " & _
              ",bs_ine_mto_int_moratorio " & _
              ",bs_ine_mto_impuesto " & _
              ",bs_ine_mto_total_factura " & _
              ",bs_ine_val_porc_beca " & _
              ",bs_ine_mto_beca " & _
              ",bs_ine_val_porc_descuento " & _
              ",bs_ine_mto_descuento " & _
              ",bs_ine_val_porc_exoneracion " & _
              ",bs_ine_mto_exoneracion " & _
              ",bs_ine_mto_total_deduccion " & _
              ",bs_ine_mto_total_neto " & _
              ",bs_ine_mto_efectivo " & _
              ",bs_ine_mto_letra_cambio " & _
              ",bs_ine_numero_letra_cambio " & _
              ",bs_ine_mto_transaccion " & _
              ",bs_ine_cod_transaccion " & _
              ",bs_ine_cod_cta_corriente " & _
              ",bs_ine_mto_tarjeta " & _
              ",bs_ine_cod_autorizacion " & _
              ",bs_ine_cod_cta_tarjeta " & _
              ",bs_ine_mto_total_forma_pago " & _
              ",bs_ine_mto_total_vuelto " & _
              ",bs_ine_cod_usr_creacion " & _
              ",bs_ine_fec_creacion " & _
              ",bs_ine_cod_usr_aplica " & _
              ",bs_ine_fec_aplicacion " & _
              ",bs_ine_anuado " & _
              ",bs_ine_fecha_deposito " & _
              ",bs_ine_mto_dinero_estudiante " & _
              ",bs_ine_forma_pago)  " & _
              " VALUES " & _
              "( 'SP01' ," & _
              " '" & nPerId & "' ," & _
              " '" & bs_ine_est_id & "' , " & _
              "  getdate() , " & _
              " '" & bs_ine_ind_tipo_documento & "' , " & _
              " '" & bs_ine_cod_documento & "' , " & _
              " 'CRC', " & _
              " '" & nTipoCambio & "' ," & _
              " 'P', " & _
              " 0 ," & _
              " " & bs_ine_mto_documento & " , " & _
              " " & bs_ine_mto_matricula & ", " & _
              " " & bs_ine_mto_seguro & ", " & _
              " " & bs_ine_mto_materias & ", " & _
              " " & bs_ine_mto_int_ordinario & ", " & _
              " " & bs_ine_mto_int_moratorio & " , " & _
              " 0, " & _
              " " & bs_ine_mto_total_factura & ", " & _
              " 0, " & _
              " 0," & _
              " " & bs_ine_val_porc_descuento & ", " & _
              " " & bs_ine_mto_descuento & "  , " & _
              " 0," & _
              " 0, " & _
              " " & bs_ine_mto_total_deduccion & ", " & _
              " " & bs_ine_mto_total_neto & ", " & _
              " " & bs_ine_mto_efectivo & ", " & _
              " " & bs_ine_mto_letra_cambio & ", " & _
              " '" & bs_ine_numero_letra_cambio & "'  , " & _
              " " & bs_ine_mto_transaccion & ", " & _
              " '" & bs_ine_cod_transaccion & "' , " & _
              " '" & bs_ine_cod_cta_corriente & "'  , " & _
              " " & bs_ine_mto_tarjeta & ", " & _
              " '" & bs_ine_cod_autorizacion & "' , " & _
              " '" & bs_ine_cod_cta_tarjeta & "' , " & _
              " " & bs_ine_mto_total_forma_pago & " ," & _
              " 0, " & _
              " '" & bs_ine_cod_usr_creacion & " ', " & _
              " getdate() , " & _
              " 'NR' , " & _
              " getdate() , " & _
              " 0, " & _
              " getdate(), " & _
              " 0,  " & _
              " 0 ) "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
            Throw ex
        End Try

    End Function

    Private Function lfInsertaInterfazDetalle(ByVal bs_ine_id As Integer, ByVal bs_ind_cod_centro_costo As String, ByVal bs_ind_mto_detalle As Decimal, ByVal bs_ind_ind_tipo_detalle As String) As String
        Dim strSql As String
        Try
            strSql = "INSERT INTO bs_interfaz_detalle " & _
                   "(bs_ine_cod_compania " & _
                   ",bs_ine_id " & _
                   ",bs_ind_cod_centro_costo " & _
                   ",bs_ind_mto_detalle " & _
                   ",bs_ind_mto_impuesto " & _
                   ",bs_ind_ind_tipo_detalle ) " & _
                   " VALUES " & _
                   "('SP01' " & _
                   "," & bs_ine_id & " " & _
                   ",'" & bs_ind_cod_centro_costo & "' " & _
                   "," & bs_ind_mto_detalle & " " & _
                   ",0 " & _
                   ",'" & bs_ind_ind_tipo_detalle & "') "

            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
            Throw ex
        End Try

    End Function

    Public Function pfObtenerDetallesCursosFactura(ByVal pFactura As String) As DataSet
        Dim strSql As String = String.Empty
        Try
            strSql = "select bse.bs_ine_id,dem.ENM_ID, dem.PLE_ID,dem.COSTO,isnull(dem.CENTRO_COSTO,'0') as CENTRO_COSTO " & _
                     "from DEM_DETALLE_MATRICULA dem " & _
                     "inner join ENM_ENCABEZADO_MATRICULA enm " & _
                     "on dem.ENM_ID = enm.ID " & _
                     "inner join bs_interfaz_encabezado bse " & _
                     "on enm.FACTURA = bse.bs_ine_cod_documento " & _
                     "where(enm.FACTURA = '" & pFactura & "') "

            'SE ELIMINA LA REFERENCIA A LA TABLA DE DETALLE DE FACTURA POR NO TENER NINGUNA RELEVANCIA

            '   strSql = "select bse.bs_ine_id,dem.ENM_ID, dem.PLE_ID,dem.COSTO,isnull(dem.CENTRO_COSTO,'0') as CENTRO_COSTO " & _
            '"from DEM_DETALLE_MATRICULA dem " & _
            '"inner join ENM_ENCABEZADO_MATRICULA enm " & _
            '"on dem.ENM_ID = enm.ID " & _
            '"inner join bs_interfaz_encabezado bse " & _
            '"on enm.FACTURA = bse.bs_ine_cod_documento " & _
            '"inner join DEM_DETALLE_MATRICULA def " & _
            '"on enm.ID = def.ENM_ID " & _
            '"where(enm.FACTURA = '" & pFactura & "') "
            Dim adpObtener As New SqlDataAdapter(strSql, cnnConexion)
            Dim dstObtener As New DataSet
            adpObtener.Fill(dstObtener, "resultado")
            Return dstObtener
        Catch ex As Exception
            Throw ex
        End Try


    End Function

    <WebMethod()> _
    Public Function pfObtenerCentroCostoMatricula(ByVal pPlaID As Integer) As String
        Dim lRstConsulta As New DataSet
        Dim strSql As String = String.Empty
        Dim lValor As String
        Try
            strSql = "select CENTRO_COSTO from PLa_Plan " & _
                     "where(ID = " & pPlaID & ") "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    Public Function pfObtenerRecibo(ByVal Letra As String) As String
        Dim lRstConsulta As New DataSet
        Dim strSql As String = String.Empty
        Dim lValor As String
        Try
            strSql = "select TOP 1 REC.RECIBO from REC_RECIBO REC " & _
                     "inner join ALE_ABONO_LETRA ALE " & _
                     "on rec.RECIBO = ALE.RECIBO " & _
                     "inner join ENL_ENCABEZADO_LETRA enl " & _
                     "on ale.ENL_ID = enl.ID " & _
                    "where enl.NUMERO = '" & Letra & "'" & _
                    "order by rec.recibo desc"
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        End Try


    End Function


    Public Function pfObtenerMaxIneID() As Integer
        Dim lRstConsulta As New DataSet
        Dim strSql As String = String.Empty
        Dim lValor As Integer
        Try
            strSql = "select max(bs_ine_id) from bs_interfaz_encabezado"
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            lValor = cmdObtener.ExecuteScalar
            Return lValor
        Catch ex As Exception
            Throw ex
        End Try


    End Function








#End Region



#Region "FUNCIONES BNCOMERCIOELECTRONICO"

    'Función que Actualiza el Estado de la factura BNCOMERCIOELECTRONICO
    <WebMethod()> _
    Public Function pfActualizaEstadoFacturaBNComercioElectronico(ByVal pNumeroFactura As Integer, ByVal nMontoDescuento As Integer, ByVal Total As Integer, ByVal TotalSinDescuento As Integer, ByVal strCreo As String, ByVal pstrNumeroTarjeta As String, ByVal nTipoTransaccion As Integer, ByVal pDescuento As Integer, ByVal pTipoCompra As Integer) As Boolean
        Dim strSql As String = String.Empty
        Dim lIDFactura = pfObtenerIDFactura(pNumeroFactura)
        Dim nCXB_ID As Integer = 4
        Dim lPeriodoID As Integer = pfObtenerPeriodoActivo()
        If nTipoTransaccion = 1 Then
            strSql = "Update ENM_ENCABEZADO_MATRICULA " & _
                "Set Cancelada = 1, " & _
                "OBSERVACIONES = 'PAGO CON BNCR-USP CONECTIVIDAD', " & _
                "MODIFICO = '" & strCreo & "', " & _
                "FECHAMODIFICO = getdate(), " & _
                "DESCUENTO =  " & pDescuento & " , " & _
                "TOTAL = " & Total & ", " & _
                "TOTALSINDESCUENTO = " & TotalSinDescuento & ", " & _
                "EFECTIVO = 0 , " & _
                "TARJETA = 0 , " & _
                "MONTO_DESCUENTO = " & nMontoDescuento & ", " & _
                "CXB_ID = " & nCXB_ID & ", " & _
                "PER_ID = " & lPeriodoID & ", " & _
                "MONTOTRANSACCION = " & Total & ", " & _
                "FECHA_DEPOSITO = getdate() " & _
                " Where ID =" & lIDFactura

        ElseIf nTipoTransaccion = 2 Then
            If pTipoCompra = 1 Then
                strSql = "Update ENM_ENCABEZADO_MATRICULA " & _
             "Set Cancelada= 1, " & _
             "OBSERVACIONES = 'PAGO EN BN COMERCIO ELECTRONICO', " & _
             "MODIFICO = '" & strCreo & "', " & _
             "FECHAMODIFICO = getdate(), " & _
             "DESCUENTO = " & pDescuento & " , " & _
             "TOTAL = " & Total & ", " & _
             "TOTALSINDESCUENTO = " & TotalSinDescuento & ", " & _
             "NUMEROTARJETA = '" & pstrNumeroTarjeta & "', " & _
             "NOMBRETARJETA = 'BN', " & _
             "MONTO_DESCUENTO = " & nMontoDescuento & ", " & _
             "TARJETA = " & Total & ", " & _
             "PER_ID = " & lPeriodoID & ", " & _
             "MONTOTRANSACCION = 0, " & _
             "CXB_ID = 0 " & _
             " Where ID =" & lIDFactura
            ElseIf pTipoCompra = 2 Then
                strSql = "Update ENM_ENCABEZADO_MATRICULA " & _
             "Set Cancelada= 1, " & _
             "OBSERVACIONES = 'PAGO EN LINEA BN COMERCIO ELECTRONICO', " & _
             "MODIFICO = '" & strCreo & "', " & _
             "FECHAMODIFICO = getdate(), " & _
             "DESCUENTO = 0 , " & _
             "TOTAL = " & TotalSinDescuento & ", " & _
             "TOTALSINDESCUENTO = 0, " & _
             "NUMEROTARJETA = '" & pstrNumeroTarjeta & "', " & _
             "NOMBRETARJETA = 'BN', " & _
             "MONTO_DESCUENTO = 0, " & _
             "TARJETA = " & TotalSinDescuento & ", " & _
             "PER_ID = " & lPeriodoID & ", " & _
             "MONTOTRANSACCION = 0, " & _
             "CXB_ID = 0 " & _
             " Where ID = " & lIDFactura
            End If
        End If

        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            cmdObtener.ExecuteNonQuery()
            cnnConexion.Close()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function

    'BEST 
    <WebMethod()> _
    Public Function lfInsertaInterfazEncabezadoBNComercioElectronico(ByVal bs_ine_est_id As String, ByVal bs_ine_cod_documento As String, ByVal bs_ine_ind_tipo_documento As String, ByVal bs_ine_mto_documento As Decimal, ByVal bs_ine_mto_matricula As Decimal, ByVal bs_ine_mto_seguro As Decimal, ByVal bs_ine_mto_materias As Decimal, ByVal bs_ine_mto_int_ordinario As Decimal, ByVal bs_ine_mto_int_moratorio As Decimal, ByVal bs_ine_mto_impuesto As Decimal, ByVal bs_ine_mto_total_factura As Decimal, ByVal bs_ine_val_porc_descuento As Decimal, ByVal bs_ine_mto_descuento As Decimal, ByVal bs_ine_val_porc_exoneracion As Decimal, ByVal bs_ine_mto_exoneracion As Decimal, ByVal bs_ine_mto_total_deduccion As Decimal, ByVal bs_ine_mto_total_neto As Decimal, ByVal bs_ine_mto_efectivo As Decimal, ByVal bs_ine_mto_letra_cambio As Decimal, ByVal bs_ine_numero_letra_cambio As String, ByVal bs_ine_mto_transaccion As Decimal, ByVal bs_ine_cod_transaccion As String, ByVal bs_ine_cod_cta_corriente As String, ByVal bs_ine_mto_tarjeta As Decimal, ByVal bs_ine_cod_autorizacion As String, ByVal bs_ine_cod_cta_tarjeta As String, ByVal bs_ine_mto_total_forma_pago As Decimal, ByVal bs_ine_cod_usr_creacion As String, ByVal bs_ine_mto_total_vuelto As Decimal, ByVal bs_ine_cod_usr_aplica As String, ByVal bs_ine_mto_dinero_estudiante As Decimal, ByVal bs_ine_forma_pago As Integer) As Boolean


        Dim strSql As String
        Dim nPerId As Integer
        Try
            nPerId = pfObtenerPeriodoActivo()

            strSql = "INSERT INTO bs_interfaz_encabezado " & _
               "(bs_ine_cod_compania " & _
               ",bs_ine_per_id " & _
               ",bs_ine_est_id " & _
               ",bs_ine_fec_documento " & _
              ",bs_ine_ind_tipo_documento " & _
              ",bs_ine_cod_documento " & _
              ",bs_ine_cod_moneda " & _
              ",bs_ine_val_tipo_cambio " & _
              ",bs_ine_ind_estado " & _
              ",bs_ine_num_asiento " & _
              ",bs_ine_mto_documento " & _
              ",bs_ine_mto_matricula " & _
              ",bs_ine_mto_seguro " & _
              ",bs_ine_mto_materias " & _
              ",bs_ine_mto_int_ordinario " & _
              ",bs_ine_mto_int_moratorio " & _
              ",bs_ine_mto_impuesto " & _
              ",bs_ine_mto_total_factura " & _
              ",bs_ine_val_porc_beca " & _
              ",bs_ine_mto_beca " & _
              ",bs_ine_val_porc_descuento " & _
              ",bs_ine_mto_descuento " & _
              ",bs_ine_val_porc_exoneracion " & _
              ",bs_ine_mto_exoneracion " & _
              ",bs_ine_mto_total_deduccion " & _
              ",bs_ine_mto_total_neto " & _
              ",bs_ine_mto_efectivo " & _
              ",bs_ine_mto_letra_cambio " & _
              ",bs_ine_numero_letra_cambio " & _
              ",bs_ine_mto_transaccion " & _
              ",bs_ine_cod_transaccion " & _
              ",bs_ine_cod_cta_corriente " & _
              ",bs_ine_mto_tarjeta " & _
              ",bs_ine_cod_autorizacion " & _
              ",bs_ine_cod_cta_tarjeta " & _
              ",bs_ine_mto_total_forma_pago " & _
              ",bs_ine_mto_total_vuelto " & _
              ",bs_ine_cod_usr_creacion " & _
              ",bs_ine_fec_creacion " & _
              ",bs_ine_cod_usr_aplica " & _
              ",bs_ine_fec_aplicacion " & _
              ",bs_ine_anuado " & _
              ",bs_ine_fecha_deposito " & _
              ",bs_ine_mto_dinero_estudiante " & _
              ",bs_ine_forma_pago)  " & _
              " VALUES " & _
              "( 'SP01' ," & _
              " '" & nPerId & "' ," & _
              " '" & bs_ine_est_id & "' , " & _
              "  getdate() , " & _
              " '" & bs_ine_ind_tipo_documento & "' , " & _
              " '" & bs_ine_cod_documento & "' , " & _
              " 'CRC', " & _
              " 1 , " & _
              " 'P', " & _
              " 0 ," & _
              " " & bs_ine_mto_documento & " , " & _
              " " & bs_ine_mto_matricula & ", " & _
              " " & bs_ine_mto_seguro & ", " & _
              " " & bs_ine_mto_materias & ", " & _
              " " & bs_ine_mto_int_ordinario & ", " & _
              " " & bs_ine_mto_int_moratorio & " , " & _
              " 0, " & _
              " " & bs_ine_mto_total_factura & ", " & _
              " 0, " & _
              " 0," & _
              " " & bs_ine_mto_descuento & "  , " & _
              " " & bs_ine_val_porc_descuento & ", " & _
              " 0," & _
              " 0, " & _
              " " & bs_ine_mto_total_deduccion & ", " & _
              " " & bs_ine_mto_total_neto & ", " & _
              " " & bs_ine_mto_efectivo & ", " & _
              " " & bs_ine_mto_letra_cambio & ", " & _
              " '" & bs_ine_numero_letra_cambio & "'  , " & _
              " " & bs_ine_mto_transaccion & ", " & _
              " '" & bs_ine_cod_transaccion & "' , " & _
              " '" & bs_ine_cod_cta_corriente & "'  , " & _
              " " & bs_ine_mto_tarjeta & ", " & _
              " '" & bs_ine_cod_autorizacion & "' , " & _
              " '" & bs_ine_cod_cta_tarjeta & "' , " & _
              " " & bs_ine_mto_total_forma_pago & " ," & _
              " 0, " & _
              " '" & bs_ine_cod_usr_creacion & " ', " & _
              " getdate() , " & _
              " 'NR' , " & _
              " getdate() , " & _
              " 0, " & _
              " getdate(), " & _
              " 0,  " & _
              " 0 ) "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return True
        Catch ex As Exception
            Return False
            Throw ex
        End Try

    End Function



    'Actualiza Matricula
    <WebMethod()> _
    Public Function pfActualizaMateriasMatriculadasBNComercioElectronico(ByVal nEstId As Integer, ByVal nTipo As Integer) As Boolean
        Dim strSql As String = String.Empty
        Dim lID_Periodo As Integer = pfObtenerPeriodoActivo()
        If nTipo = 1 Then 'letra
            strSql = "Update Ple_Plan_X_Estudiante " & _
                     "Set Matriculo='" & cCreo & "', Temporal=2, Estado='" & cEstadoMatriculado & _
                     "' Where Est_Id=" & nEstId & _
                     " and Temporal = 1" & _
                     " and Periodo_ID=" & lID_Periodo
        ElseIf nTipo = 2 Then 'matri
            strSql = "Update Ple_Plan_X_Estudiante " & _
                     "Set Matriculo='PAGOS EN LÍNEA', Temporal=2, Estado='" & cEstadoMatriculado & _
                     "',FECHAMATRICULO = getdate() Where Est_Id=" & nEstId & _
                     " and Temporal = 1" & _
                     " and Periodo_ID=" & lID_Periodo
        End If
        Try
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
            cnnConexion.Open()
            cmdObtener.ExecuteNonQuery()
            cnnConexion.Close()
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function


    'LETRAS



    'ACTUALIZA ENL
    <WebMethod()> _
    Public Function lfActualizaENLBNComercioElectronico(ByVal nMontoRecibo As Double, ByVal nEnlID As Integer, ByVal nNumComprobante As Integer, ByVal nRecibo As Integer) As String
        Dim strSql As String
        Try
            strSql = "UPDATE ENL_ENCABEZADO_LETRA SET " & _
                        " FECHA = (SELECT FECHA from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & ") " & _
                        " ,TOTALINTERES = (SELECT TOTALINTERES from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & ") " & _
                        " ,INTERESES = (SELECT INTERESES from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & ") " & _
                        " ,SALDO = (SELECT SALDO from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " ) " & _
                        " ,TOTALPAGADO = (SELECT TOTALPAGADO from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " ) " & _
                        " ,CANCELADA = (SELECT CANCELADA from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " ) " & _
                        " ,ESTADO = (SELECT ESTADO from tmpEnl_Encabezado_Letra where ENL_ID = " & nEnlID & " )  " & _
                        " WHERE (ID = " & nEnlID & ") "
            If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
            cnnConexion.Open()
            Dim cmdComandoSQL As New SqlCommand(strSql, cnnConexion)
            cmdComandoSQL.Transaction = grefTransaction
            cmdComandoSQL.ExecuteNonQuery()
            Return "1"
        Catch ex As Exception
            Throw ex
            Return "0"
        Finally
        End Try
    End Function

    ''ALE ABONO LETRA
    <WebMethod()> _
    Public Function pfCancelaLetradeCambioBNComercioElectronico(ByVal pNumeroLetraCambio As String, ByVal pMontoPagado As Decimal, ByVal pReciboPago As Integer) As Boolean

        Dim strSql As String = String.Empty

        lValoresLetraCambio = pfDetalleLetraCambio(pNumeroLetraCambio)
        Dim lTipoLetraCambio As Byte = lValoresLetraCambio(eColumnasArregloLetras.TipoLetraCambio)
        Dim lIDLetraCambio As Integer = lValoresLetraCambio(eColumnasArregloLetras.IdLetraCambio)
        Dim lRecibo As Integer = pReciboPago
        Dim lDiasAcumuados As Integer = lValoresLetraCambio(eColumnasArregloLetras.DiasOrdinarios)
        Dim lTotalPago As Decimal = lValoresLetraCambio(eColumnasArregloLetras.MontoTotal)
        Dim lAbono As Decimal = pMontoPagado
        Dim lSaldo As Decimal = lValoresLetraCambio(eColumnasArregloLetras.SaldoLetra)
        Dim lObservaciones As String = String.Empty
        Dim lFecha As Date = Date.Now
        Dim lDiasMoratorios As Integer = lValoresLetraCambio(eColumnasArregloLetras.DiasMoratorios)
        Dim lInteresOrdinario As Decimal = lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario)
        Dim lAbonoInteresOrdinario As Decimal = 0
        Dim lSaldoInteresOrdinario As Decimal = 0
        Dim lInteresMoratorio As Decimal = lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio)
        Dim lAbonoInteresMoratorio As Decimal = 0
        Dim lSaldoInteresMoratorio As Decimal = 0
        Dim lSaldoRemante As Decimal = 0
        Dim lEstadoLetra As Byte = 0
        Dim lFechaUltimoPago As String = pfFechaCalculoIntereses(pNumeroLetraCambio)


        Try

            'Las letras de Bach-Lic tienen cálculo de intereses
            If lTipoLetraCambio = eTipoLetraCambio.BachLic Then

                'El Saldo remanente es igual al monto del abono
                lSaldoRemante = lAbono

                'INICIA CALCULANDO EL INTERES MORATORIO
                'SI ES SUFICIENTE CANCELA EL TOTAL DEL INTERES MORATORIO
                If lSaldoRemante >= lInteresMoratorio Then
                    lAbonoInteresMoratorio = lInteresMoratorio
                    lSaldoInteresMoratorio = 0
                    lSaldoRemante = lSaldoRemante - lInteresMoratorio
                Else
                    lAbonoInteresMoratorio = lSaldoRemante
                    lSaldoInteresMoratorio = lInteresMoratorio - lSaldoInteresMoratorio
                    lSaldoRemante = lAbono
                End If
                'SI ES SUFICIENTE CANCELA EL TOTAL DEL INTERES ORDINARIO
                If lSaldoRemante >= lInteresOrdinario Then
                    lAbonoInteresOrdinario = lInteresOrdinario
                    lSaldoInteresOrdinario = 0
                    lSaldoRemante = lSaldoRemante - lInteresOrdinario
                Else
                    lAbonoInteresOrdinario = lSaldoRemante
                    lSaldoInteresOrdinario = lInteresOrdinario - lSaldoRemante
                    lSaldoRemante = 0
                End If
                'DEL SALDO DEL ABONO, AMORTIZA EL CAPITAL
                If lSaldoRemante >= lSaldo Then
                    lSaldo = 0
                Else
                    lSaldo = lSaldo - lSaldoRemante
                End If

                'INGRESA EL REGISTRO A LA TABLA DE ABONOS DE LETRA DE CAMBIO.

                strSql = "INSERT INTO ALE_ABONO_LETRA " & _
                            "(ENL_ID,RECIBO,DIASACUMULADO,TOTALPAGO,ABONO,SALDO,OBSERVACION " & _
                            ",FECHA,DIASMORATORIOS,INTERESORDINARIO,ABONOINTERESORDINARIO " & _
                            ",SALDOINTERESORDINARIO,INTERESMORATORIO,ABONOINTERESMORATORIO " & _
                            ",SALDOINTERESMORATORIO) " & _
                        "VALUES(" & _
                            lIDLetraCambio & "," & _
                            lRecibo & "," & _
                            lDiasAcumuados & "," & _
                            lTotalPago & "," & _
                            lAbono & "," & _
                            lSaldo & "," & _
                            "'CANCELA LETRA DE CAMBIO MEDIANTE CONVENIO USP-BN'" & "," & _
                            "getdate(), " & _
                            lDiasMoratorios & "," & _
                            lInteresOrdinario & "," & _
                            lAbonoInteresOrdinario & "," & _
                            lSaldoInteresOrdinario & "," & _
                            lInteresMoratorio & "," & _
                            lAbonoInteresMoratorio & "," & _
                            lSaldoInteresMoratorio & ")"

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()
                Dim cmdObtener As New SqlCommand(strSql, cnnConexion)
                cnnConexion.Open()
                cmdObtener.ExecuteNonQuery()
                cnnConexion.Close()


                'ACTUALIZA LOS VALORES EN EL ENCABEZADO DE LA LETRA
                If lSaldo <= 0 Then lEstadoLetra = eEstadoLetra.Cancelada

                Dim dFechaPago As Date = lFechaUltimoPago

                strSql = "Update ENL_ENCABEZADO_LETRA " & _
                        "Set FechaultimoPago = convert(datetime,'" & dFechaPago & "',101) " & _
                        ",Fecha= Getdate() " & _
                ",Saldo=" & lSaldo & _
                ",TotalPagado=" & (lValoresLetraCambio(eColumnasArregloLetras.TotalPagado) + lAbono) & _
                ",TotalInteres=" & (lValoresLetraCambio(eColumnasArregloLetras.TotalInteresOrdinario) + lInteresOrdinario) & _
                ",TotalInteresMoratorios=" & (lValoresLetraCambio(eColumnasArregloLetras.TotalInteresMoratorio) + lInteresMoratorio) & _
                ",Cancelada =" & lEstadoLetra & _
                " Where NUMERO = '" & pNumeroLetraCambio & "'"

                If cnnConexion.State = ConnectionState.Open Then cnnConexion.Close()

                cnnConexion.Open()

                Dim cmdActualizar As New SqlCommand(strSql, cnnConexion)
                cmdActualizar.ExecuteNonQuery()
                cnnConexion.Close()


                'Las letras de Cambio de Maestria no tienen Intereses
            Else
                Return False
            End If
            Return True
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    'LETRAS, BEST ENCABEZADO, DETALLE Y CREA RECIBO
    <WebMethod()> _
    Public Function lfInsertaLetracambioBNComercioElectronico(ByVal strUsuario As String, _
                                           ByVal strClaveUsuario As String, _
                                           ByVal strTipoLlave As String, _
                                           ByVal strLlaveAcceso As String, _
                                           ByVal nCodAgencia As Byte, _
                                           ByVal nCodConvenio As Byte, _
                                           ByVal nTipoTransaccion As Byte, _
                                           ByVal strPeriodo As String, _
                                           ByVal nMontoPagado As Decimal, _
                                           ByVal nMontoTotalRecibo As Decimal, _
                                           ByVal nNumFactura As Integer, _
                                           ByVal strRecibo As String, _
                                           ByVal nSelfVerificacion As Byte, _
                                           ByVal strCodMoneda As String, _
                                           ByVal nNumComprobante As Integer, _
                                           ByVal nNumeroCuotas As Integer, _
                                           ByVal nTipoPago As Byte, _
                                           ByVal nCodBanco As Integer, _
                                           ByVal nNumCheque As Long, _
                                           ByVal nNumCuenta As Long, _
                                           ByVal nMontoTotal As Decimal, _
                                           ByVal nNumNotaCredito As Integer, _
                                           ByVal nMontoCurso As Integer, _
                                           ByVal nMontoMatricula As Integer, _
                                           ByVal nMontoSeguro As Integer, _
                                           ByVal nMontoDescuento As Decimal, _
                                           ByVal strCreo As String, _
                                           ByVal strNumeroTarjeta As String, _
                                           ByVal pDescuento As Decimal, _
                                           ByRef pTipoCompra As Decimal) As DataSet



        Dim ldtsDatos As New DataSet
        'Procedimiento para crear las tablas del DataSet de retorno del la CLase Pago de Servicios
        'Creamos el data set que contendrá la información del Afiliado para el archivo de Envío
        Dim dtsPagoServicios As New Data.DataSet("dtsPagoServicios")
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim dtbDetalle As Data.DataTable = New DataTable("Detalle")
        Dim dtbSubDetalle As Data.DataTable = New DataTable("SubDetalle")
        Dim lstrNumeroTransaccion As String = String.Empty
        Dim lValoresFacturacion(5) As Decimal
        Dim nTotalConceptosPago As Integer = 0
        Dim nEstID As Integer
        Dim lExitosa As Boolean
        Dim nEnlID As Integer = 0
        Dim nCodRespuesta As Integer = 0
        Dim strDescRespuesta As String = ""
        Dim nDescuento As Decimal = 0.0
        Dim nTotalSinDescuento As Decimal = 0.0
        Dim nTipoEntidad = 2

        If lfExisteRecibo(strLlaveAcceso) > 0 Then
            nEstID = (strLlaveAcceso)
            lExitosa = True
        Else
            lExitosa = False
            nCodRespuesta = cCodMensaje51
            strDescRespuesta = strMensaje16
        End If

        If lExitosa = True Then
            nEnlID = pfObtenerIDLetraCambio(strLlaveAcceso)
            nDescuento = 0.0
            lstrNumeroTransaccion = pfGeneraReciboPago(strRecibo, strTipoLlave, nEstID, nNumComprobante, nMontoPagado, nCodConvenio, nDescuento, nMontoPagado, strCreo, nTipoEntidad, strNumeroTarjeta)
            'Verificamos que se haya creado un recibo de manera exitosa
            If lstrNumeroTransaccion > 0 Then
                lExitosa = True
                'Se inserta en la bitacora de pagos de bancos
                Dim nRecibo As String = lstrNumeroTransaccion
                Dim nPlaID As Integer = fnPlanAcademico(nEstID)
                If lfInsertaInterfazEncabezado(nEstID, _
                            nRecibo, _
                            cValorAbono, _
                            nMontoPagado, _
                            cCero, _
                            cCero, _
                           cCero, _
                          lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario), _
                            lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio), _
                            cCero, _
                            nMontoPagado, _
                            cCero, _
                            cCero, _
                            cCero, _
                            cCero, _
                            cCero, _
                            nMontoPagado, _
                            cCero, _
                            cCero, _
                            cCero, _
                            cCero, _
                            cCero, _
                            cCero, _
                            nMontoPagado, _
                            nNumComprobante, _
                            strNumeroTarjeta, _
                            nMontoPagado, _
                            cCreoCredomatic, _
                            cCero, _
                            "NR", _
                            cCero, _
                            cCero) Then
                    Dim bs_ine_id As Integer
                    Dim bs_ind_cod_centro_costo As String
                    Dim bs_ind_mto_detalle As Decimal
                    Dim bs_ind_ind_tipo_detalle As String
                    If nMontoPagado > 0 Then
                        bs_ind_ind_tipo_detalle = "A"
                        bs_ine_id = pfObtenerMaxIneID()
                        bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlaID)
                        bs_ind_mto_detalle = nMontoPagado
                        lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)

                    End If
                End If
            End If
        Else
            lExitosa = False
            nCodRespuesta = cCodMensaje99
            strDescRespuesta = strMensaje99
        End If



    End Function

    'Funcion que ejecuta el pago de servicios
    <WebMethod()> _
    Public Function pfEjecutarPagoServicioBNComercioElectronico(ByVal strUsuario As String, _
                                              ByVal strClaveUsuario As String, _
                                              ByVal strTipoLlave As String, _
                                              ByVal strLlaveAcceso As String, _
                                              ByVal nCodAgencia As Byte, _
                                              ByVal nCodConvenio As Byte, _
                                              ByVal nTipoTransaccion As Byte, _
                                              ByVal strPeriodo As String, _
                                              ByVal nMontoPagado As Decimal, _
                                              ByVal nMontoTotalRecibo As Decimal, _
                                              ByVal nNumFactura As Integer, _
                                              ByVal strRecibo As String, _
                                              ByVal nSelfVerificacion As Byte, _
                                              ByVal strCodMoneda As String, _
                                              ByVal nNumComprobante As Integer, _
                                              ByVal nNumeroCuotas As Integer, _
                                              ByVal nTipoPago As Byte, _
                                              ByVal nCodBanco As Integer, _
                                              ByVal nNumCheque As Long, _
                                              ByVal nNumCuenta As Long, _
                                              ByVal nMontoTotal As Decimal, _
                                              ByVal nNumNotaCredito As Integer, _
                                              ByVal nMontoCurso As Integer, _
                                              ByVal nMontoMatricula As Integer, _
                                              ByVal nMontoSeguro As Integer, _
                                              ByVal nMontoDescuento As Decimal, _
                                              ByVal strCreo As String, _
                                              ByVal strNumeroTarjeta As String, _
                                              ByVal pDescuento As Decimal, _
                                              ByRef pTipoCompra As Decimal) As DataSet
        Dim ldtsDatos As New DataSet
        'Procedimiento para crear las tablas del DataSet de retorno del la CLase Pago de Servicios
        'Creamos el data set que contendrá la información del Afiliado para el archivo de Envío
        Dim dtsPagoServicios As New Data.DataSet("dtsPagoServicios")
        Dim dtbEncabezado As Data.DataTable = New DataTable("Encabezado")
        Dim dtbDetalle As Data.DataTable = New DataTable("Detalle")
        Dim dtbSubDetalle As Data.DataTable = New DataTable("SubDetalle")
        Dim ldtrFila As DataRow
        Dim ldtrConceptosInformativos As DataRow
        Dim ldtrDetalle As DataRow
        Dim lstrNumeroTransaccion As String = String.Empty
        Dim lValoresFacturacion(5) As Decimal
        Dim nTotalConceptosPago As Integer = 0
        Dim strDescripcionConceptp As String
        Dim nEstID As Integer
        Dim lExitosa As Boolean
        Dim nEnlID As Integer = 0
        Dim nCodRespuesta As Integer = 0
        Dim strDescRespuesta As String = ""
        Dim nDescuento As Decimal = 0.0
        Dim nTotalSinDescuento As Decimal = 0.0
        Dim nTipoEntidad = 2

        'TABLA DEL ENCABEZADO
        '-------------------------------------------------------------------------------------
        dtbEncabezado.Columns.Add("strCodRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("strDescripcionRespuesta", Type.GetType("System.String"))
        dtbEncabezado.Columns.Add("nCantConceptosInformativos", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("nCantConceptosPago", Type.GetType("System.Int32"))
        dtbEncabezado.Columns.Add("strNumTransaccionPago", Type.GetType("System.String"))
        dtsPagoServicios.Tables.Add(dtbEncabezado)
        '-------------------------------------------------------------------------------------

        'Tabla del SubDetalle correspondiente a la clase conceptos Informativos
        '--------------------------------------------------------------------------------------
        dtbSubDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbSubDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbSubDetalle.Columns.Add("strValorConcepto", Type.GetType("System.String"))
        dtsPagoServicios.Tables.Add(dtbSubDetalle)
        '--------------------------------------------------------------------------------------

        'TABLA DEL DETALLE correspondiente a la clase de Conceptos de Pago
        '-------------------------------------------------------------------------------------
        dtbDetalle.Columns.Add("nConsecutivo", Type.GetType("System.Int32"))
        dtbDetalle.Columns.Add("nCodConcepto", Type.GetType("System.Byte"))
        dtbDetalle.Columns.Add("strDescripcionConcepto", Type.GetType("System.String"))
        dtbDetalle.Columns.Add("nMontoConcepto", Type.GetType("System.Decimal"))
        dtsPagoServicios.Tables.Add(dtbDetalle)
        '-------------------------------------------------------------------------------------
        Try
            'Si el servicio está diponible ejecuta todo el proceso; si no devuelve el encabezado del DTS 
            'con el código de respuesta 50

            If lfServicioDisponible() = True Then

                'Valida que el usuario este registrado en el sistema
                If lfValidaUsuario(strUsuario, strClaveUsuario) = True Then
                    'Segun el tipo de convenio realiza el pago
                    Select Case nCodConvenio

                        Case eCodigoConvenio.Facturas
                            'Carga los detelles de la factura antes de aplicar los cambios
                            If lfExisteFactura(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXFactura(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje13
                            End If

                            If lExitosa = True Then

                                If pTipoCompra = 1 Then
                                    If nMontoCurso > 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfActualizaMateriasMatriculadas(nEstID, nTipoEntidad)
                                    End If
                                    If nMontoMatricula > 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfIngresaPagoMatriculaEstudiante(nEstID, nNumFactura, nMontoMatricula)
                                    End If

                                    If nMontoSeguro > 0 Then
                                        nTotalConceptosPago = nTotalConceptosPago + 1
                                        pfIngresaPagoSeguroEstudiante(nEstID, nNumFactura)
                                    End If
                                    If lExitosa Then
                                        If nTotalConceptosPago > 0 Then
                                            'Se actualiza el estado a cancelada por BN-CONECTIVDAD
                                            If pfActualizaEstadoFactura(nNumFactura, _
                                                                        nMontoDescuento, _
                                                                        nMontoPagado, _
                                                                        nMontoTotal, _
                                                                        strCreo, _
                                                                        strNumeroTarjeta, nTipoEntidad, pDescuento, pTipoCompra) Then
                                                lstrNumeroTransaccion = strLlaveAcceso
                                                If lstrNumeroTransaccion > 0 Then
                                                    'Se inserta el pago a la bitacora de pagos de bancos
                                                    lfInsertarBitacoraPagosBancos(0, nNumFactura, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, cCreoCredomatic)
                                                    If lfInsertaInterfazEncabezado(nEstID, _
                                                        strLlaveAcceso, _
                                                        cValorMatricula, _
                                                        cCero, _
                                                        nMontoMatricula, _
                                                        nMontoSeguro, _
                                                        nMontoCurso, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        nMontoTotal, _
                                                        nMontoDescuento, _
                                                        pDescuento, _
                                                        cCero, _
                                                        cCero, _
                                                        nMontoDescuento, _
                                                        nMontoTotalRecibo, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        cCero, _
                                                        nMontoTotalRecibo, _
                                                        nNumComprobante, _
                                                        strNumeroTarjeta, _
                                                        nMontoTotalRecibo, _
                                                        cCreoCredomatic, _
                                                        cCero, _
                                                        "NR", _
                                                        cCero, _
                                                        cCero) = True Then
                                                        Dim dtsDetalle As New DataSet
                                                        Dim bs_ine_id As Integer
                                                        Dim bs_ind_cod_centro_costo As String
                                                        Dim bs_ind_mto_detalle As Decimal
                                                        Dim bs_ind_ind_tipo_detalle As String
                                                        Dim nPlanID As Integer
                                                        dtsDetalle = pfObtenerDetallesCursosFactura(strLlaveAcceso)
                                                        If nMontoCurso > 0 Then
                                                            bs_ind_ind_tipo_detalle = "C"
                                                            If dtsDetalle.Tables(0).Rows.Count > 0 Then
                                                                For i = 0 To dtsDetalle.Tables(0).Rows.Count - 1
                                                                    bs_ine_id = dtsDetalle.Tables(0).Rows(i)("bs_ine_id")
                                                                    bs_ind_cod_centro_costo = dtsDetalle.Tables(0).Rows(i)("CENTRO_COSTO")
                                                                    bs_ind_mto_detalle = dtsDetalle.Tables(0).Rows(i)("COSTO")
                                                                    lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                                Next
                                                            End If
                                                        End If
                                                        If nMontoMatricula > 0 Then
                                                            nPlanID = fnPlanAcademico(nEstID)
                                                            bs_ind_ind_tipo_detalle = "M"
                                                            bs_ine_id = dtsDetalle.Tables(0).Rows(0)("bs_ine_id")
                                                            bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlanID)
                                                            bs_ind_mto_detalle = nMontoMatricula
                                                            lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                        End If
                                                        If nMontoSeguro > 0 Then
                                                            nPlanID = fnPlanAcademico(nEstID)
                                                            bs_ind_ind_tipo_detalle = "S"
                                                            bs_ine_id = dtsDetalle.Tables(0).Rows(0)("bs_ine_id")
                                                            bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlanID)
                                                            bs_ind_mto_detalle = nMontoSeguro
                                                            lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                        End If
                                                    End If
                                                Else
                                                    lExitosa = False
                                                    nCodRespuesta = cCodMensaje51
                                                    strDescRespuesta = "No se creó recibo."
                                                End If
                                            Else
                                                lExitosa = False
                                                nCodRespuesta = cCodMensaje51
                                                strDescRespuesta = "No se logró actualizar la factura solicitada"
                                            End If
                                        Else
                                            lExitosa = False
                                            nCodRespuesta = cCodMensaje51
                                            strDescRespuesta = "No hay Conceptos de pago pendientes para la factura solicitada."
                                        End If
                                    End If

                                Else
                                    If pTipoCompra = 2 Then

                                        If nMontoCurso > 0 Then
                                            nTotalConceptosPago = nTotalConceptosPago + 1
                                            pfActualizaMateriasMatriculadas(nEstID, nTipoEntidad)
                                        End If

                                        If lExitosa Then
                                            If nTotalConceptosPago > 0 Then
                                                'Se actualiza el estado a cancelada por BN-CONECTIVDAD
                                                If pfActualizaEstadoFactura(nNumFactura, nMontoDescuento, nMontoPagado, nMontoTotal, strCreo, strNumeroTarjeta, nTipoEntidad, pDescuento, pTipoCompra) Then

                                                    lstrNumeroTransaccion = strLlaveAcceso
                                                    If lstrNumeroTransaccion > 0 Then
                                                        'Se inserta el pago a la bitacora de pagos de bancos
                                                        lfInsertarBitacoraPagosBancos(0, nNumFactura, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, cCreoCredomatic)
                                                        If lfInsertaInterfazEncabezado(nEstID, _
                                                   strLlaveAcceso, _
                                                   cValorMatricula, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoCurso, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoTotalRecibo, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoTotal, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   cCero, _
                                                   nMontoTotalRecibo, _
                                                   nNumComprobante, _
                                                   strNumeroTarjeta, _
                                                   nMontoTotalRecibo, _
                                                   cCreoCredomatic, _
                                                   cCero, _
                                                   "NR", _
                                                   cCero, _
                                                   cCero) Then
                                                            Dim dtsDetalle As DataSet
                                                            Dim bs_ine_id As Integer
                                                            Dim bs_ind_cod_centro_costo As String
                                                            Dim bs_ind_mto_detalle As Decimal
                                                            Dim bs_ind_ind_tipo_detalle As String
                                                            dtsDetalle = pfObtenerDetallesCursosFactura(strLlaveAcceso)
                                                            If nMontoCurso > 0 Then
                                                                bs_ind_ind_tipo_detalle = "C"
                                                                If dtsDetalle.Tables(0).Rows.Count > 0 Then
                                                                    For i = 0 To dtsDetalle.Tables(0).Rows.Count - 1
                                                                        bs_ine_id = dtsDetalle.Tables(0).Rows(i)("bs_ine_id")
                                                                        bs_ind_cod_centro_costo = dtsDetalle.Tables(0).Rows(i)("CENTRO_COSTO")
                                                                        bs_ind_mto_detalle = dtsDetalle.Tables(0).Rows(i)("COSTO")
                                                                        lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)
                                                                    Next
                                                                End If
                                                            End If
                                                        End If
                                                    Else
                                                        lExitosa = False
                                                        nCodRespuesta = cCodMensaje51
                                                        strDescRespuesta = "No se creó recibo."
                                                    End If
                                                Else
                                                    lExitosa = False
                                                    nCodRespuesta = cCodMensaje51
                                                    strDescRespuesta = "No se logró actualizar la factura solicitada"
                                                End If
                                            Else
                                                lExitosa = False
                                                nCodRespuesta = cCodMensaje51
                                                strDescRespuesta = "No hay Conceptos de pago pendientes para la factura solicitada."
                                            End If
                                        End If

                                    End If

                                End If
                            End If
                            'Else

                            'End If

                            If lExitosa Then
                                'crea la información para el Dataset de retorno
                                'Crea la fila del encabezado ****************************************************************
                                ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                ldtrFila("strCodRespuesta") = cCodMensaje00
                                ldtrFila("strDescripcionRespuesta") = strMensaje01
                                ldtrFila("nCantConceptosInformativos") = 1
                                ldtrFila("nCantConceptosPago") = nTotalConceptosPago
                                ldtrFila("strNumTransaccionPago") = lstrNumeroTransaccion
                                dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                '********************************************************************************************
                                'Crea la fila del Detalle de la Factura ****************************************************************

                                'Si existen valores en el pago de materias con descuento los agrega al table de response
                                If nMontoCurso > 0 Then
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 1
                                    ldtrDetalle("nCodConcepto") = 1
                                    ldtrDetalle("strDescripcionConcepto") = "Materias Matriculadas "
                                    ldtrDetalle("nMontoConcepto") = nMontoCurso
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                End If
                                'Si existen valores en el pago de matricula los agrega al table de response
                                If nMontoMatricula > 0 Then
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 2
                                    ldtrDetalle("nCodConcepto") = 2
                                    ldtrDetalle("strDescripcionConcepto") = "Monto Matricula"
                                    ldtrDetalle("nMontoConcepto") = nMontoMatricula
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                End If
                                'Si existen valores para el concepto de pago de seguro los agrega al table de response
                                If nMontoSeguro > 0 Then
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 3
                                    ldtrDetalle("nCodConcepto") = 3
                                    ldtrDetalle("strDescripcionConcepto") = "Seguro Estudiantil "
                                    ldtrDetalle("nMontoConcepto") = nMontoSeguro
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table
                                End If
                                'Crea la fila del Detalle de la Factura para conceptos informativos ****************************************************************
                                'Crea el detalle referente a los conceptos Informativos
                                '***************************************************************************************
                                ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                ldtrConceptosInformativos("nCodConcepto") = 1
                                ldtrConceptosInformativos("strDescripcionConcepto") = "Fact. Matricula"
                                ldtrConceptosInformativos("strValorConcepto") = nNumFactura
                                dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)

                            Else
                                ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                ldtrFila("strCodRespuesta") = nCodRespuesta
                                ldtrFila("strDescripcionRespuesta") = strDescRespuesta
                                ldtrFila("nCantConceptosInformativos") = 0
                                ldtrFila("nCantConceptosPago") = 1
                                ldtrFila("strNumTransaccionPago") = 0
                                dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                '********************************************************************************************

                                'Crea el detalle referente a los conceptos Informativos
                                '***************************************************************************************
                                ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                ldtrConceptosInformativos("nCodConcepto") = cCodMensaje99
                                ldtrConceptosInformativos("strDescripcionConcepto") = strMensaje02
                                ldtrConceptosInformativos("strValorConcepto") = 0
                                dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                '*************************************************************************************

                                'Crea la fila del Detalle de la Factura ****************************************************************
                                ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                ldtrDetalle("nConsecutivo") = 0
                                ldtrDetalle("nCodConcepto") = 99
                                ldtrDetalle("strDescripcionConcepto") = strMensaje02
                                ldtrDetalle("nMontoConcepto") = 0
                                dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle) 'Agrega la Fila al Data Table

                                '-------------------------------------------------------------'
                            End If
                        Case eCodigoConvenio.LetraCambio
                            'Verificamos que exista el recibo que se esta utilizando como llave de accesos y obtenemos el ID del estudiante de esa letra
                            If lfExisteRecibo(strLlaveAcceso) > 0 Then
                                nEstID = lfObtieneIDEstudianteXRecibo(strLlaveAcceso)
                                lExitosa = True
                            Else
                                lExitosa = False
                                nCodRespuesta = cCodMensaje51
                                strDescRespuesta = strMensaje16
                            End If

                            If lExitosa = True Then
                                nEnlID = pfObtenerIDLetraCambio(strLlaveAcceso)
                                'Verificamos que exista un comprobante para la letra por pagar
                                'If Not lfExisteNumComprobanteLetra(nEnlID, nNumComprobante) Then
                                'Si el monto de pago es mayor al monto minimo admitido, se realizan los procesos de pago
                                ' If nMontoPagado > lfPagoMinimoLetra() Then
                                nDescuento = 0.0
                                lstrNumeroTransaccion = pfGeneraReciboPago(strRecibo, strTipoLlave, nEstID, nNumComprobante, nMontoPagado, nCodConvenio, nDescuento, nMontoPagado, strCreo, nTipoEntidad, strNumeroTarjeta)
                                'Verificamos que se haya creado un recibo de manera exitosa
                                If lstrNumeroTransaccion > 0 Then
                                    lExitosa = True
                                    'Se inserta en la bitacora de pagos de bancos
                                    If lfInsertarBitacoraPagosBancos(nEnlID, 0, nNumComprobante, lstrNumeroTransaccion, nCodConvenio, nMontoTotal, strCreo) = 1 Then
                                        Dim nRecibo As String = lstrNumeroTransaccion
                                        Dim nPlaID As Integer = fnPlanAcademico(nEstID)
                                        If lfInsertaInterfazEncabezado(nEstID, _
                                                    nRecibo, _
                                                    cValorAbono, _
                                                    nMontoPagado, _
                                                    cCero, _
                                                    cCero, _
                                                   cCero, _
                                                  lValoresLetraCambio(eColumnasArregloLetras.MontoInteresOrdinario), _
                                                    lValoresLetraCambio(eColumnasArregloLetras.MontoInteresMoratorio), _
                                                    cCero, _
                                                    nMontoPagado, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    nMontoPagado, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    cCero, _
                                                    nMontoPagado, _
                                                    nNumComprobante, _
                                                    strNumeroTarjeta, _
                                                    nMontoPagado, _
                                                    cCreoCredomatic, _
                                                    cCero, _
                                                    "NR", _
                                                    cCero, _
                                                    cCero) Then
                                            Dim bs_ine_id As Integer
                                            Dim bs_ind_cod_centro_costo As String
                                            Dim bs_ind_mto_detalle As Decimal
                                            Dim bs_ind_ind_tipo_detalle As String
                                            If nMontoPagado > 0 Then
                                                bs_ind_ind_tipo_detalle = "A"
                                                bs_ine_id = pfObtenerMaxIneID()
                                                bs_ind_cod_centro_costo = pfObtenerCentroCostoMatricula(nPlaID)
                                                bs_ind_mto_detalle = nMontoPagado
                                                lfInsertaInterfazDetalle(bs_ine_id, bs_ind_cod_centro_costo, bs_ind_mto_detalle, bs_ind_ind_tipo_detalle)

                                            End If
                                        End If
                                    End If
                                Else
                                    lExitosa = False
                                    nCodRespuesta = cCodMensaje99
                                    strDescRespuesta = strMensaje99
                                End If

                                'Else
                                '    lExitosa = False
                                '    nCodRespuesta = cCodMensaje99
                                '    strDescRespuesta = strMensaje21
                                'End If
                                'Else
                                '    lExitosa = False
                                '    nCodRespuesta = cCodMensaje51
                                '    strDescRespuesta = "El Número de Comprobante ya existe para la de cambio solicitada."
                                'End If

                                If lExitosa = True Then

                                    'Crea la fila del encabezado ****************************************************************
                                    ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                    ldtrFila("strCodRespuesta") = cCodMensaje00
                                    ldtrFila("strDescripcionRespuesta") = strMensaje01
                                    ldtrFila("nCantConceptosInformativos") = 1
                                    ldtrFila("nCantConceptosPago") = 1
                                    ldtrFila("strNumTransaccionPago") = lstrNumeroTransaccion
                                    dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                    '********************************************************************************************

                                    'Crea el detalle referente a los conceptos Informativos
                                    '***************************************************************************************
                                    strDescripcionConceptp = "Pago/Abono L.C: " & strRecibo

                                    ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                    ldtrConceptosInformativos("nCodConcepto") = 1
                                    ldtrConceptosInformativos("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                                    ldtrConceptosInformativos("strValorConcepto") = nNumComprobante
                                    dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                    '*************************************************************************************

                                    'Crea la fila del Detalle de la letra ****************************************************************
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 1
                                    ldtrDetalle("nCodConcepto") = 1
                                    ldtrDetalle("strDescripcionConcepto") = strDescripcionConceptp.Substring(0, 20)
                                    ldtrDetalle("nMontoConcepto") = nMontoPagado
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table

                                Else
                                    ldtrFila = dtsPagoServicios.Tables("Encabezado").NewRow
                                    ldtrFila("strCodRespuesta") = nCodRespuesta
                                    ldtrFila("strDescripcionRespuesta") = strMensaje21
                                    ldtrFila("nCantConceptosInformativos") = 0
                                    ldtrFila("nCantConceptosPago") = 0
                                    ldtrFila("strNumTransaccionPago") = 0
                                    dtsPagoServicios.Tables("Encabezado").Rows.Add(ldtrFila)  'Agrega la Fila al Data Table 
                                    '********************************************************************************************

                                    'Crea el detalle referente a los conceptos Informativos
                                    '***************************************************************************************
                                    ldtrConceptosInformativos = dtsPagoServicios.Tables("SubDetalle").NewRow
                                    ldtrConceptosInformativos("nCodConcepto") = 0
                                    ldtrConceptosInformativos("strDescripcionConcepto") = strMensaje99
                                    ldtrConceptosInformativos("strValorConcepto") = 0
                                    dtsPagoServicios.Tables("SubDetalle").Rows.Add(ldtrConceptosInformativos)
                                    '*************************************************************************************

                                    'Crea la fila del Detalle  ****************************************************************
                                    ldtrDetalle = dtsPagoServicios.Tables("Detalle").NewRow
                                    ldtrDetalle("nConsecutivo") = 0
                                    ldtrDetalle("nCodConcepto") = 0
                                    ldtrDetalle("strDescripcionConcepto") = strMensaje99
                                    ldtrDetalle("nMontoConcepto") = 0
                                    dtsPagoServicios.Tables("Detalle").Rows.Add(ldtrDetalle)  'Agrega la Fila al Data Table

                                End If
                            End If

                        Case eCodigoConvenio.Carnets


                        Case eCodigoConvenio.CursosLibres


                        Case eCodigoConvenio.VidaEstudiantil


                    End Select

                Else

                End If

                'Si el servico no esta disponible invoca a la funcion publica para que cargue el dts con la información requerida por el BNCR.
            Else
                dtsPagoServicios = pfServicioDisponible()
            End If


            Return dtsPagoServicios
        Catch ex As Exception
            Throw ex
        End Try
        Return ldtsDatos
    End Function


#End Region




End Class