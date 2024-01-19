using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SAP.Middleware;
using SAP.Middleware.Connector;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Net;

namespace MB51_TO_SQLHrxHr
{
    public class SAP30
    {
        private string Entorno = "PRD";
        public string FindHU(string h_hu)
        {
            // Coloca ceros al principio del HU
            int hu_len = h_hu.Length;
            int i = 0;
            string Zeros = "";
            if (hu_len < 20)
            {
                for (i = hu_len; i <= 19; i++)
                {
                    Zeros = Zeros + "0";
                }
            }

            h_hu = Zeros + h_hu;

            string valor = "Listo";
            try
            {


                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_BAPI_FINDHU_FMX");
                Bapi.SetValue("MATNR", " ");
                Bapi.SetValue("EXTHU", h_hu);

                Bapi.Invoke(SapRfcDestination);
                valor = Bapi.GetValue("RHU").ToString();

                if (valor == "1") { valor = "true"; } else { valor = "false"; }
                return valor;
            }
            catch (Exception ex)
            {
                return ex.ToString();
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }



        }
        public void GeneracionDocumentosSAP(DateTime Doc_Date,DateTime Plan_Date,int NoInventario,string Referencia,string Almacen,string Usuario)
        {
            try
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZFM_CREATE_POST_INV_DOCUMENT");
                Bapi.SetValue("I_PLANT", "1841");
                Bapi.SetValue("I_DOC_DATE", Doc_Date);
                Bapi.SetValue("I_PLAN_DATE", Plan_Date);
                Bapi.SetValue("I_INVENTORY_NO", NoInventario.ToString());
                Bapi.SetValue("I_INVENTORY_REF", Referencia);
                DataTable DT = GetDiferenciasByAlmacen(Almacen);
                IRfcTable IT_LIST = Bapi.GetTable("IT_LIST");
                for (int i=0;i<=DT.Rows.Count-1;i++)
                {
                    string Mov = TipoMovimiento(Convert.ToInt32(DT.Rows[i]["Diferencia"]));
                    int Dif = Convert.ToInt32(DT.Rows[i]["Diferencia"]);
                    if(Dif<0)
                    {
                        Dif = Dif * -1;
                    }
                    IT_LIST.Append();
                    IT_LIST.SetValue("STGE_LOC", Almacen);
                    IT_LIST.SetValue("MATNR", DT.Rows[i]["NoParte"].ToString());
                    IT_LIST.SetValue("ERFMG", Convert.ToInt64(Dif));
                    IT_LIST.SetValue("TIPO_MOVIMIENTO", Convert.ToInt32(Mov));
                }
                Bapi.Invoke(SapRfcDestination);


                //IRfcTable BWWART = Bapi.GetTable("IT_WERKS");
                //BWWART.Append();
                //BUDAT.SetValue("WERKS", "1841");
                //101

                IRfcTable RESULTADOS = Bapi.GetTable("ET_RETURN");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                for (int i = 0; DTSalida.Rows.Count - 1 >= i; i++)
                {
                    this.InsertDocInventariosLogs(DateTime.Now.Year,Almacen, Convert.ToInt32(DT.Rows[i]["Diferencia"]),0,DTSalida.Rows[i]["MESSAGE"].ToString(),Usuario,"", 
                        TipoMovimiento(Convert.ToInt32(DT.Rows[i]["Diferencia"])),DTSalida.Rows[i]["TYPE"].ToString(), DTSalida.Rows[i]["ID"].ToString(), Convert.ToInt32(DTSalida.Rows[i]["NUMBER"].ToString()), 
                        DTSalida.Rows[i]["MESSAGE"].ToString(), DTSalida.Rows[i]["LOG_NO"].ToString(), DTSalida.Rows[i]["LOG_MSG_NO"].ToString(), DTSalida.Rows[i]["MESSAGE_V1"].ToString(),
                        DTSalida.Rows[i]["MESSAGE_V2"].ToString(), DTSalida.Rows[i]["MESSAGE_V3"].ToString(), DTSalida.Rows[i]["MESSAGE_V4"].ToString(), DTSalida.Rows[i]["PARAMETER"].ToString(), 
                        DTSalida.Rows[i]["ROW"].ToString(), DTSalida.Rows[i]["FIELD"].ToString(), DTSalida.Rows[i]["SYSTEM"].ToString());
                }

                //return DTSalida;
            }
            catch(Exception ex)
            {
                throw new Exception(ex.ToString());
                //return null;
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
        }
        public void InsertDocInventariosLogs (int Periodo, string Almacen, int Diferencia, int NoDocumento, string Mensajes, 
                                                string Usuario, string NoParte, string Movimiento,string TYPE, string IDSAP, 
                                                int NUMBER,string MESSAGE, string LOG_NO, string LOG_MSG_NO, string MESSAGEV1,
                                                string MESSAGEV2, string MESSAGEV3, string MESSAGEV4, string PARAMETERS, 
                                                string ROW, string FIELD,string SYSTEM)
        {
            try
            {
                SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["InventarioDB"].ToString());
                SqlCommand oCmd = new SqlCommand("dbo.InsertDocInventarios", oConn);
                SqlDataAdapter oAdapt = new SqlDataAdapter(oCmd);
                DataTable DT = new DataTable();
                oCmd.CommandType = CommandType.StoredProcedure;
                oCmd.Parameters.Add(new SqlParameter("@Periodo", Periodo));
                oCmd.Parameters.Add(new SqlParameter("@Almacen", Almacen));
                oCmd.Parameters.Add(new SqlParameter("@Diferencia", Diferencia));
                oCmd.Parameters.Add(new SqlParameter("@NoDocumento", NoDocumento));
                oCmd.Parameters.Add(new SqlParameter("@Mensajes", MESSAGE));
                oCmd.Parameters.Add(new SqlParameter("@Usuario", Usuario));
                oCmd.Parameters.Add(new SqlParameter("@NoParte", NoParte));
                oCmd.Parameters.Add(new SqlParameter("@Movimiento", Movimiento));
                oCmd.Parameters.Add(new SqlParameter("@TYPE", TYPE));
                oCmd.Parameters.Add(new SqlParameter("@IDSAP", IDSAP));
                oCmd.Parameters.Add(new SqlParameter("@NUMBER", NUMBER));
                oCmd.Parameters.Add(new SqlParameter("@MESSAGE", MESSAGE));
                oCmd.Parameters.Add(new SqlParameter("@LOG_NO", LOG_NO));
                oCmd.Parameters.Add(new SqlParameter("@LOG_MSG_NO", LOG_MSG_NO));
                oCmd.Parameters.Add(new SqlParameter("@MESSAGE_V1", MESSAGEV1));
                oCmd.Parameters.Add(new SqlParameter("@MESSAGE_V2", MESSAGEV2));
                oCmd.Parameters.Add(new SqlParameter("@MESSAGE_V3", MESSAGEV3));
                oCmd.Parameters.Add(new SqlParameter("@MESSAGE_V4", MESSAGEV4));
                oCmd.Parameters.Add(new SqlParameter("@PARAMETER", PARAMETERS));
                oCmd.Parameters.Add(new SqlParameter("@ROW", ROW));
                oCmd.Parameters.Add(new SqlParameter("@FIELD", FIELD));
                oCmd.Parameters.Add(new SqlParameter("@SYSTEM", SYSTEM)); 
                oConn.Open();
                oCmd.ExecuteNonQuery();
                oConn.Close();
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }
        public string TipoMovimiento(int Diferencia)
        {
            string Movimiento = "";
            if(Diferencia<0)
            {
                Movimiento = "702";
            }
            else
            {
                Movimiento = "701";
            }
            return Movimiento;
        }
        public DataTable GetDiferenciasByAlmacen(string Almacen)
        {
            //GetAjustesByAlmacen
            try
            {
                SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["InventarioDB"].ToString());
                SqlCommand oCmd = new SqlCommand("dbo.GetAjustesByAlmacen", oConn);
                SqlDataAdapter oAdapt = new SqlDataAdapter(oCmd);
                DataTable DT = new DataTable();
                oCmd.CommandType = CommandType.StoredProcedure;
                oCmd.Parameters.Add(new SqlParameter("@Almacen", Almacen));
                oAdapt.Fill(DT);
                if (DT != null)
                {
                    if (DT.Rows.Count != 0)
                    {
                        return DT;
                    }
                    else return null;
                }
                else return null;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }

           
        }
        public void InsertMD16(int evento,string MATNR,string PERTR,string PSTTR, string PEDTR,double GSMNG, string MEINS, string PLAFX, string BESKZ, string ESOBS, string PLNUM, string PAART, string KNTTP, string KDAUF, string KDPOS)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
            SqlCommand oCmd= new SqlCommand("dbo.InsertMD16",oConn);
            oCmd.CommandType=CommandType.StoredProcedure;
            oCmd.Parameters.Add(new SqlParameter("@NoEvento",evento));
            oCmd.Parameters.Add(new SqlParameter("@MATNR",MATNR));
            oCmd.Parameters.Add(new SqlParameter("@PERTR",PERTR));
            oCmd.Parameters.Add(new SqlParameter("@PSTTR",PSTTR));
            oCmd.Parameters.Add(new SqlParameter("@PEDTR",PEDTR));
            oCmd.Parameters.Add(new SqlParameter("@GSMNG",GSMNG));
            oCmd.Parameters.Add(new SqlParameter("@MEINS", MEINS));
            oCmd.Parameters.Add(new SqlParameter("@PLAFX", PLAFX));
            oCmd.Parameters.Add(new SqlParameter("@BESKZ", BESKZ));
            oCmd.Parameters.Add(new SqlParameter("@ESOBS", ESOBS));
            oCmd.Parameters.Add(new SqlParameter("@PLNUM", PLNUM));
            oCmd.Parameters.Add(new SqlParameter("@PAART", PAART));
            oCmd.Parameters.Add(new SqlParameter("@KNTTP", KNTTP));
            oCmd.Parameters.Add(new SqlParameter("@KDAUF", KDAUF));
            oCmd.Parameters.Add(new SqlParameter("@KDPOS", KDPOS));
            oConn.Open();
            oCmd.ExecuteNonQuery();
            oConn.Close();
        }
        public void  Ejecuta_MD16()
        {
            try
            {

                int Evento= GetUltimoEvento();
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                DateTime Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZMF_MD16");
                Bapi.SetValue("I_WERKS", "1841");
                Bapi.SetValue("I_DISPO_FLAG", "X");
                Bapi.SetValue("I_MATNR_FLAG", "");
                Bapi.SetValue("I_DISPO", "PPP");
                Bapi.SetValue("I_MATNR", "");
                Bapi.SetValue("I_BISDT", Fecha.Date);
        
                //101
                Bapi.Invoke(SapRfcDestination);
                IRfcTable RESULTADOS = Bapi.GetTable("T_DATA");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                for ( int i = 0; DTSalida.Rows.Count - 1 > i; i++)
                {
                    InsertMD16(Evento, DTSalida.Rows[i]["MATNR"].ToString(), DTSalida.Rows[i]["PERTR"].ToString(), DTSalida.Rows[i]["PSTTR"].ToString(), DTSalida.Rows[i]["PEDTR"].ToString(),
                        Convert.ToDouble(DTSalida.Rows[i]["GSMNG"].ToString()), DTSalida.Rows[i]["MEINS"].ToString(), DTSalida.Rows[i]["PLAFX"].ToString(), DTSalida.Rows[i]["BESKZ"].ToString(),
                        DTSalida.Rows[i]["ESOBS"].ToString(), DTSalida.Rows[i]["PLNUM"].ToString(), DTSalida.Rows[i]["PAART"].ToString(), DTSalida.Rows[i]["KNTTP"].ToString(),
                        DTSalida.Rows[i]["KDAUF"].ToString(), DTSalida.Rows[i]["KDPOS"].ToString());
                }

                IRfcTable MENSAJES = Bapi.GetTable("T_RESULT");
                DataTable DTMensajes = this.ConvertToDT(MENSAJES);
                for (int i = 0; DTMensajes.Rows.Count - 1 > i;i++ )
                {
                    InsertaMensajeLogMD16(DTMensajes.Rows[i][0].ToString(), DTMensajes.Rows[i][1].ToString(), System.Net.Dns.GetHostName().ToString(),Evento);
                }
                   // return DTSalida;
            }
            catch(Exception ex) 
            {
                
            }
        }
        public int GetUltimoEvento()
        {
            int Evento = 0;
            try
            {
                
                SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
                SqlCommand command = new SqlCommand("Select MAX(NoEvento) from OrdenesPlaneadasSAP", oConn);
                oConn.Open();
                SqlDataReader reader = command.ExecuteReader();
                if (reader.HasRows)
                {
                    while (reader.Read())
                    {
                        if(reader.IsDBNull(0))
                        {
                            Evento = 1;
                        }
                        else
                        {
                            Evento = reader.GetInt32(0) + 1;
                        }
                        
                    }
                }
                else
                {
                    Evento = 1;
                }
                oConn.Close();
            }
            catch (Exception ex)
            {

            }
         
            return Evento;
        }
        public void InsertaMensajeLogMD16(string Tipo,string Mensaje,string Host, int evento)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
            SqlCommand command = new SqlCommand("dbo.InsertLogMD16", oConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@Evento", evento);
            command.Parameters.AddWithValue("@Tipo", Tipo);
            command.Parameters.AddWithValue("@Mensaje",Mensaje);
            command.Parameters.AddWithValue("@Host",Host);
            oConn.Open();
            command.ExecuteNonQuery();
            oConn.Close();
        }
        public DataTable EjecutaZ_ShippingHeader1(int Dias,DateTime FechaIni,DateTime FechaFin)
        {
            try
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_MM_EXTRAE_MOVIMIENTOS");
                Bapi.SetValue("I_DIAS", Dias);
                Bapi.SetValue("I_CRIT", "");
                Bapi.SetValue("I_EXCL", "");
                Bapi.SetValue("I_SUMA", "");
                IRfcTable BUDAT = Bapi.GetTable("IT_DATUM");
                BUDAT.Append();
                BUDAT.SetValue("DATUM", FechaIni);
                BUDAT.SetValue("DATUM", FechaFin);

                IRfcTable BWWART = Bapi.GetTable("IT_WERKS");
                BWWART.Append();
                BUDAT.SetValue("WERKS", "1841");
                //101

                IRfcTable RESULTADOS = Bapi.GetTable("IT_HEADER1");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                return DTSalida;
            }
            catch
            {
                return null;
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
           
        }

        public string CancelHU(string h_hu, string h_version)
        {
            // Coloca ceros al principio del HU
            int Count;
            string valor = "Listo";


            try
            {
                // string valor = "";

                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_Bapi_Findhu_Fmx");
                Bapi.SetValue("V_HUCHAR", h_hu);
                Bapi.SetValue("V_VERID", h_version);

                Bapi.Invoke(SapRfcDestination);
                valor = Bapi.GetValue("SUBRC").ToString();
                if (valor == "1") { valor = "true"; } else { valor = "false"; }
                return valor;
            }
            catch (Exception ex)
            {
                return ex.ToString();
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }







            if (valor == "1") { valor = "true"; }
            else { valor = "false"; }
            return valor;

        }
        public DataTable CargaInventario(DateTime Fechaini, DateTime FechaFin)
        {
            try
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_MM_EXTRAE_MOVIMIENTOS");
                Bapi.SetValue("IN_WERKS", "1841");
                IRfcTable BUDAT = Bapi.GetTable("IN_T_BUDAT");
                BUDAT.Append();
                BUDAT.SetValue("SIGN_R", "I");
                BUDAT.SetValue("OPTION_R", "EQ");
                BUDAT.SetValue("LOW", Fechaini);
                BUDAT.SetValue("HIGH", FechaFin);

                IRfcTable BWWART = Bapi.GetTable("IN_T_BWART");
                BWWART.Append();

                //101
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "101");
                BWWART.SetValue("HIGH", " ");
                //102
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "102");
                BWWART.SetValue("HIGH", " ");
                //131
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "131");
                BWWART.SetValue("HIGH", " ");
                //132
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "132");
                BWWART.SetValue("HIGH", " ");
                //201
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "201");
                BWWART.SetValue("HIGH", " ");
                //202
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "202");
                BWWART.SetValue("HIGH", " ");

                //261
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "261");
                BWWART.SetValue("HIGH", " ");
                //262
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "262");
                BWWART.SetValue("HIGH", " ");
                //551
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "551");
                BWWART.SetValue("HIGH", " ");
                //552
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "552");
                BWWART.SetValue("HIGH", " ");
                //553
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "553");
                BWWART.SetValue("HIGH", " ");
                //554
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "554");
                BWWART.SetValue("HIGH", " ");
                //601
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "601");
                BWWART.SetValue("HIGH", " ");
                //602
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "602");
                BWWART.SetValue("HIGH", " ");
                //701
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "701");
                BWWART.SetValue("HIGH", " ");
                //702
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "702");
                BWWART.SetValue("HIGH", " ");
                //711
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "711");
                BWWART.SetValue("HIGH", " ");
                //712
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "702");
                BWWART.SetValue("HIGH", " ");

                Bapi.SetValue("IN_KOKRS", "CEMM");

                Bapi.Invoke(SapRfcDestination);
                IRfcTable RESULTADOS = Bapi.GetTable("EX_T_MOVS");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                string MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTM, USNAM, BKTXT, XBLNR, MATNR, GRUND, BWART, SAKTO, LGORT, SHKZG;
                double DMBTR;
                string WAERS, MEINS, SGTXT, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH;
                int MENGE;
                for (int i = 0; i <= DTSalida.Rows.Count - 1; i++)
                {
                    MBLNR = DTSalida.Rows[i]["MBLNR"].ToString();
                    ZEILE = DTSalida.Rows[i]["ZEILE"].ToString();
                    BUDATv = DTSalida.Rows[i]["BUDAT"].ToString();
                    WERKS = DTSalida.Rows[i]["WERKS"].ToString();
                    BLDAT = DTSalida.Rows[i]["BLDAT"].ToString();
                    CPUTM = DTSalida.Rows[i]["CPUTM"].ToString();
                    USNAM = DTSalida.Rows[i]["USNAM"].ToString();
                    BKTXT = DTSalida.Rows[i]["BKTXT"].ToString();
                    XBLNR = DTSalida.Rows[i]["XBLNR"].ToString();
                    MATNR = DTSalida.Rows[i]["MATNR"].ToString();
                    GRUND = DTSalida.Rows[i]["GRUND"].ToString();
                    BWART = DTSalida.Rows[i]["BWART"].ToString();
                    SAKTO = DTSalida.Rows[i]["SAKTO"].ToString();
                    LGORT = DTSalida.Rows[i]["LGORT"].ToString();
                    SHKZG = DTSalida.Rows[i]["SHKZG"].ToString();
                    DMBTR = Convert.ToDouble(DTSalida.Rows[i]["DMBTR"].ToString());
                    WAERS = DTSalida.Rows[i]["WAERS"].ToString();
                    MENGE = Convert.ToInt32(Math.Truncate(Convert.ToDecimal(DTSalida.Rows[i]["MENGE"])));
                    MEINS = DTSalida.Rows[i]["MEINS"].ToString();
                    SGTXT = DTSalida.Rows[i]["SGTXT"].ToString();
                    KOSTL = DTSalida.Rows[i]["KOSTL"].ToString();
                    MJAHR = DTSalida.Rows[i]["MJAHR"].ToString();
                    GRTXT = DTSalida.Rows[i]["GRTXT"].ToString();
                    ZCCNAME = DTSalida.Rows[i]["ZCCNAME"].ToString();
                    PRCTR = DTSalida.Rows[i]["PRCTR"].ToString();
                    ZPCNAME = DTSalida.Rows[i]["ZPCNAME"].ToString();
                    MDV01 = DTSalida.Rows[i]["MDV01"].ToString();
                    ZWCNAME = DTSalida.Rows[i]["ZWCNAME"].ToString();
                    KTSCH = DTSalida.Rows[i]["KTSCH"].ToString();
                    this.InsertaMovInventarioSQL(FechaFin, MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTM, USNAM, BKTXT, XBLNR, MATNR,
                       GRUND, BWART, SAKTO, LGORT, SHKZG, DMBTR, WAERS, MENGE, MEINS, SGTXT, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH);
                }
                return DTSalida;
            }
            catch (Exception ex)
            {
                return null;
            }

        }
        //[dbo].[InsertInventarioMB51HrxHrTEST] 
        public void InsertaMovInventarioSQLHrxHrTEST(DateTime FechaIni, string MBLNR, string ZEIL, string BUDAT, string WERKS, string BLDAT, string CPUTM, string USNAM, string BKTXT, string XBLNR, string MATNR, string GRUND,
  string BWART, string SAKTO, string LGORT, string SHKZG, double DMBTR, string WAERS, double MENGE, string MEINS, string SGTXT, string WEMPF, string KOSTL, string MAJHR, string GRTXT, string ZCCNAME,
  string PRCTR, string ZPNAME, string MDV01, string ZWCNAME, string KTSCH)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["Reporte"].ToString());
            SqlCommand oCmd = new SqlCommand("dbo.InsertInventarioMB51HrxHrTEST", oConn);
            // SqlDataAdapter oAdap = new SqlDataAdapter(oCmd);
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.Add(new SqlParameter("@FechaInventario", FechaIni));
            oCmd.Parameters.Add(new SqlParameter("@MBLNR", MBLNR));
            oCmd.Parameters.Add(new SqlParameter("@ZEIL", ZEIL));
            oCmd.Parameters.Add(new SqlParameter("@BUDAT", BUDAT));
            oCmd.Parameters.Add(new SqlParameter("@WERKS", WERKS));
            oCmd.Parameters.Add(new SqlParameter("@BLDAT", BLDAT));
            oCmd.Parameters.Add(new SqlParameter("@CPUTM", CPUTM));
            oCmd.Parameters.Add(new SqlParameter("@USNAM", USNAM));
            oCmd.Parameters.Add(new SqlParameter("@BKTXT", BKTXT));
            oCmd.Parameters.Add(new SqlParameter("@XBLNR", XBLNR));
            oCmd.Parameters.Add(new SqlParameter("@MATNR", MATNR));
            oCmd.Parameters.Add(new SqlParameter("@GRUND", GRUND));
            oCmd.Parameters.Add(new SqlParameter("@BWART", BWART));
            oCmd.Parameters.Add(new SqlParameter("@SAKTO", SAKTO));
            oCmd.Parameters.Add(new SqlParameter("@LGORT", LGORT));
            oCmd.Parameters.Add(new SqlParameter("@SHKZG", SHKZG));
            oCmd.Parameters.Add(new SqlParameter("@DMBTR", DMBTR));
            oCmd.Parameters.Add(new SqlParameter("@WAERS", WAERS));
            oCmd.Parameters.Add(new SqlParameter("@MENGE", MENGE));
            oCmd.Parameters.Add(new SqlParameter("@MEINS", MEINS));
            oCmd.Parameters.Add(new SqlParameter("@SGTXT", SGTXT));
            oCmd.Parameters.Add(new SqlParameter("@WEMPF", WEMPF));
            oCmd.Parameters.Add(new SqlParameter("@KOSTL", KOSTL));
            oCmd.Parameters.Add(new SqlParameter("@MAJHR", MAJHR));
            oCmd.Parameters.Add(new SqlParameter("@GRTXT", GRTXT));
            oCmd.Parameters.Add(new SqlParameter("@ZCCNAME", ZCCNAME));
            oCmd.Parameters.Add(new SqlParameter("@PRCTR", PRCTR));
            oCmd.Parameters.Add(new SqlParameter("@ZPNAME", ZPNAME));
            oCmd.Parameters.Add(new SqlParameter("@MDV01", MDV01));
            oCmd.Parameters.Add(new SqlParameter("@ZWCNAME", ZWCNAME));
            oCmd.Parameters.Add(new SqlParameter("@KTSCH", KTSCH));
            oConn.Open();
            oCmd.ExecuteNonQuery();
            oConn.Close();
        }

        //[dbo].[InsertInventarioMB51HrxHr] 
        public void InsertaMovInventarioSQLHrxHr(DateTime FechaIni, string MBLNR, string ZEIL, string BUDAT, string WERKS, string BLDAT, string CPUTM, string USNAM, string BKTXT, string XBLNR, string MATNR, string GRUND,
    string BWART, string SAKTO, string LGORT, string SHKZG, double DMBTR, string WAERS, double MENGE, string MEINS, string SGTXT,string WEMPF, string KOSTL, string MAJHR, string GRTXT, string ZCCNAME,
    string PRCTR, string ZPNAME, string MDV01, string ZWCNAME, string KTSCH, string AUFNR,double STPRS,double PEINH)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["Reporte"].ToString());
            SqlCommand oCmd = new SqlCommand("dbo.InsertInventarioMB51HrxHr", oConn);
            // SqlDataAdapter oAdap = new SqlDataAdapter(oCmd);
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.Add(new SqlParameter("@FechaInventario", FechaIni));
            oCmd.Parameters.Add(new SqlParameter("@MBLNR", MBLNR));
            oCmd.Parameters.Add(new SqlParameter("@ZEIL", ZEIL));
            oCmd.Parameters.Add(new SqlParameter("@BUDAT", BUDAT));
            oCmd.Parameters.Add(new SqlParameter("@WERKS", WERKS));
            oCmd.Parameters.Add(new SqlParameter("@BLDAT", BLDAT));
            oCmd.Parameters.Add(new SqlParameter("@CPUTM", CPUTM));
            oCmd.Parameters.Add(new SqlParameter("@USNAM", USNAM));
            oCmd.Parameters.Add(new SqlParameter("@BKTXT", BKTXT));
            oCmd.Parameters.Add(new SqlParameter("@XBLNR", XBLNR));
            oCmd.Parameters.Add(new SqlParameter("@MATNR", MATNR));
            oCmd.Parameters.Add(new SqlParameter("@GRUND", GRUND));
            oCmd.Parameters.Add(new SqlParameter("@BWART", BWART));
            oCmd.Parameters.Add(new SqlParameter("@SAKTO", SAKTO));
            oCmd.Parameters.Add(new SqlParameter("@LGORT", LGORT));
            oCmd.Parameters.Add(new SqlParameter("@SHKZG", SHKZG));
            oCmd.Parameters.Add(new SqlParameter("@DMBTR", DMBTR));
            oCmd.Parameters.Add(new SqlParameter("@WAERS", WAERS));
            oCmd.Parameters.Add(new SqlParameter("@MENGE", MENGE));
            oCmd.Parameters.Add(new SqlParameter("@MEINS", MEINS));
            oCmd.Parameters.Add(new SqlParameter("@SGTXT", SGTXT));
            oCmd.Parameters.Add(new SqlParameter("@WEMPF", WEMPF));
            oCmd.Parameters.Add(new SqlParameter("@KOSTL", KOSTL));
            oCmd.Parameters.Add(new SqlParameter("@MAJHR", MAJHR));
            oCmd.Parameters.Add(new SqlParameter("@GRTXT", GRTXT));
            oCmd.Parameters.Add(new SqlParameter("@ZCCNAME", ZCCNAME));
            oCmd.Parameters.Add(new SqlParameter("@PRCTR", PRCTR));
            oCmd.Parameters.Add(new SqlParameter("@ZPNAME", ZPNAME));
            oCmd.Parameters.Add(new SqlParameter("@MDV01", MDV01));
            oCmd.Parameters.Add(new SqlParameter("@ZWCNAME", ZWCNAME));
            oCmd.Parameters.Add(new SqlParameter("@KTSCH", KTSCH));
            oCmd.Parameters.Add(new SqlParameter("@AUFNR", AUFNR));
            oCmd.Parameters.Add(new SqlParameter("@STPRS", STPRS));
            oCmd.Parameters.Add(new SqlParameter("@PEINH", PEINH));
            oConn.Open();
            oCmd.ExecuteNonQuery();
            oConn.Close();
        }
        public void InsertaMovInventarioSQL(DateTime FechaIni, string MBLNR, string ZEIL, string BUDAT, string WERKS, string BLDAT, string CPUTM, string USNAM, string BKTXT, string XBLNR, string MATNR, string GRUND,
            string BWART, string SAKTO, string LGORT, string SHKZG, double DMBTR, string WAERS, int MENGE, string MEINS, string SGTXT, string KOSTL, string MAJHR, string GRTXT, string ZCCNAME,
            string PRCTR, string ZPNAME, string MDV01, string ZWCNAME, string KTSCH)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["Reporte"].ToString());
            SqlCommand oCmd = new SqlCommand("dbo.InsertInventarioMB51_HOY", oConn);
            // SqlDataAdapter oAdap = new SqlDataAdapter(oCmd);
            oCmd.CommandType = CommandType.StoredProcedure;
            oCmd.Parameters.Add(new SqlParameter("@FechaInventario", FechaIni));
            oCmd.Parameters.Add(new SqlParameter("@MBLNR", MBLNR));
            oCmd.Parameters.Add(new SqlParameter("@ZEIL", ZEIL));
            oCmd.Parameters.Add(new SqlParameter("@BUDAT", BUDAT));
            oCmd.Parameters.Add(new SqlParameter("@WERKS", WERKS));
            oCmd.Parameters.Add(new SqlParameter("@BLDAT", BLDAT));
            oCmd.Parameters.Add(new SqlParameter("@CPUTM", CPUTM));
            oCmd.Parameters.Add(new SqlParameter("@USNAM", USNAM));
            oCmd.Parameters.Add(new SqlParameter("@BKTXT", BKTXT));
            oCmd.Parameters.Add(new SqlParameter("@XBLNR", XBLNR));
            oCmd.Parameters.Add(new SqlParameter("@MATNR", MATNR));
            oCmd.Parameters.Add(new SqlParameter("@GRUND", GRUND));
            oCmd.Parameters.Add(new SqlParameter("@BWART", BWART));
            oCmd.Parameters.Add(new SqlParameter("@SAKTO", SAKTO));
            oCmd.Parameters.Add(new SqlParameter("@LGORT", LGORT));
            oCmd.Parameters.Add(new SqlParameter("@SHKZG", SHKZG));
            oCmd.Parameters.Add(new SqlParameter("@DMBTR", DMBTR));
            oCmd.Parameters.Add(new SqlParameter("@WAERS", WAERS));
            oCmd.Parameters.Add(new SqlParameter("@MENGE", MENGE));
            oCmd.Parameters.Add(new SqlParameter("@MEINS", MEINS));
            oCmd.Parameters.Add(new SqlParameter("@SGTXT", SGTXT));
            oCmd.Parameters.Add(new SqlParameter("@KOSTL", KOSTL));
            oCmd.Parameters.Add(new SqlParameter("@MAJHR", MAJHR));
            oCmd.Parameters.Add(new SqlParameter("@GRTXT", GRTXT));
            oCmd.Parameters.Add(new SqlParameter("@ZCCNAME", ZCCNAME));
            oCmd.Parameters.Add(new SqlParameter("@PRCTR", PRCTR));
            oCmd.Parameters.Add(new SqlParameter("@ZPNAME", ZPNAME));
            oCmd.Parameters.Add(new SqlParameter("@MDV01", MDV01));
            oCmd.Parameters.Add(new SqlParameter("@ZWCNAME", ZWCNAME));
            oCmd.Parameters.Add(new SqlParameter("@KTSCH", KTSCH));
            oConn.Open();
            oCmd.ExecuteNonQuery();
            oConn.Close();
        }
        public DataTable ConvertToDT(IRfcTable TablaSAP)
        {
            DataTable DT = new DataTable();
            for (int i = 0; i < TablaSAP.ElementCount; i++)
            {
                RfcElementMetadata MetaDato = TablaSAP.GetElementMetadata(i);
                DT.Columns.Add(MetaDato.Name);
            }
            foreach (IRfcStructure row in TablaSAP)
            {
                DataRow dr = DT.NewRow();
                for (int i = 0; i < TablaSAP.ElementCount; i++)
                {
                    RfcElementMetadata MetaDatos = TablaSAP.GetElementMetadata(i);
                    if (MetaDatos.DataType == RfcDataType.BCD && MetaDatos.Name == "ABC")
                    {
                        dr[i] = row.GetInt(MetaDatos.Name);
                    }
                    else
                    {
                        dr[i] = row.GetString(MetaDatos.Name);
                    }
                }
                DT.Rows.Add(dr);
            }
            return DT;
        }
        public string Empaque(int cantidad, string HU)
        {
            string erroresfinales = "";
            try
            {


                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_MMIM_HU_REETIQUETADO");
                Bapi.SetValue("LV_EXIDV", HU);
                Bapi.SetValue("LV_QTY", cantidad);

                Bapi.Invoke(SapRfcDestination);

                IRfcTable strReturn = Bapi.GetTable("EX_IT_ERRORMSGS");
                string strmsgv1 = "";
                string strmsgv2 = "";
                string strmsgv3 = "";
                string strmsgv4 = "";
                if (strReturn.RowCount != 0)
                {
                    string strtcode = strReturn.GetString("TCODE");
                    string strdyname = strReturn.GetString("DYNAME");
                    string intNumber = strReturn.GetString("DYNUMB");
                    string strm = strReturn.GetString("MSGTYP");
                    string strms = strReturn.GetString("MSGSPRA");
                    string strmsid = strReturn.GetString("MSGID");
                    string strmsg = strReturn.GetString("MSGNR");
                    strmsgv1 = strReturn.GetString("MSGV1");
                    strmsgv2 = strReturn.GetString("MSGV2");
                    strmsgv3 = strReturn.GetString("MSGV3");
                    strmsgv4 = strReturn.GetString("MSGV4");
                    string strENV = strReturn.GetString("ENV");
                    string strFLDNAME = strReturn.GetString("FLDNAME");

                }

                if (strReturn != null)
                {
                    if (strReturn.Count > 0)
                    {
                        for (int i = 0; i < strReturn.Count; i++)
                        {
                            if (strmsgv1 != "") { CL_proc.WriteData.Create_Error(strmsgv1, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv2 != "") { CL_proc.WriteData.Create_Error(strmsgv2, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv3 != "") { CL_proc.WriteData.Create_Error(strmsgv3, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv4 != "") { CL_proc.WriteData.Create_Error(strmsgv4, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            erroresfinales = "Error al Planear la HU en SAP: " + strmsgv1 + "; " + strmsgv2 + "; " + strmsgv3 + "; " + strmsgv4;
                        }
                    }
                }
                return erroresfinales;
            }
            catch (Exception ex)
            {
                return ex.ToString();
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }

        }
        public string Procesa(string h_material, string h_planta, string h_cantidad, string h_hu, string h_version)
        {

            string ErrorT = "";
            try
            {
                string valor = "Listo";
                string StrERROR = "";
                string g_material = "P" + h_material;
                string g_planta = h_planta;
                string g_cantidad = "Q" + h_cantidad;
                string g_hu = h_hu;
                string g_version = h_version;
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Zsdsh_Deliverycheck");
                Bapi.SetValue("ML_MATCHAR", g_material);
                Bapi.SetValue("ML_CANTCHAR", g_cantidad);
                Bapi.SetValue("ML_HUCHAR", g_hu);
                Bapi.SetValue("ML_VERCHAR", g_version);
                Bapi.Invoke(SapRfcDestination);
                IRfcTable strReturn = Bapi.GetTable("ERRORS_TABLE");
                string strmsgv1 = "";
                string strmsgv2 = "";
                string strmsgv3 = "";
                StrERROR = Bapi.GetValue("MESSTAB").ToString();
                if (strReturn.RowCount != 0)
                {
                    string strtcode = strReturn.GetString("TCODE");
                    string strdyname = strReturn.GetString("DYNAME");
                    string intNumber = strReturn.GetString("DYNUMB");
                    string strm = strReturn.GetString("MSGTYP");
                    string strms = strReturn.GetString("MSGSPRA");
                    string strmsid = strReturn.GetString("MSGID");
                    string strmsg = strReturn.GetString("MSGNR");
                    strmsgv1 = strReturn.GetString("MSGV1");
                    strmsgv2 = strReturn.GetString("MSGV2");
                    strmsgv3 = strReturn.GetString("MSGV3");
                    ErrorT = strmsgv2;
                    //MessageBox.Show (ErrorT);
                    // this.Mensajes1.Text = ErrorT;
                }
                return valor;
            }
            catch (Exception ex)
            {
                return ex.ToString();
                //   MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
        }
        public string ReportaSPO_CompradoSAP(string HU, string NoParte, int Cantidad)
        {
            string erroresfinales = "";
            try
            {
                if (Cantidad > 0 & NoParte != "")
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                    // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                    IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZBAPI_HU02_EMP_COM");
                    Bapi.SetValue("I_MATNR", NoParte);
                    Bapi.SetValue("I_QUANT", Cantidad);
                    Bapi.SetValue("I_STORAGE", "EM05");
                    Bapi.SetValue("I_HU", HU);
                    Bapi.SetValue("I_PLANT", "1841");
                    Bapi.Invoke(SapRfcDestination);
                    IRfcTable strReturn = Bapi.GetTable("ET_MESSAGES");
                    string strmsgv1 = "";
                    string strmsgv2 = "";
                    string strmsgv3 = "";
                    string strmsgv4 = "";
                    if (strReturn.RowCount != 0)
                    {
                        string strtcode = strReturn.GetString("TCODE");
                        string strdyname = strReturn.GetString("DYNAME");
                        string intNumber = strReturn.GetString("DYNUMB");
                        string strm = strReturn.GetString("MSGTYP");
                        string strms = strReturn.GetString("MSGSPRA");
                        string strmsid = strReturn.GetString("MSGID");
                        string strmsg = strReturn.GetString("MSGNR");
                        strmsgv1 = strReturn.GetString("MSGV1");
                        strmsgv2 = strReturn.GetString("MSGV2");
                        strmsgv3 = strReturn.GetString("MSGV3");
                        strmsgv4 = strReturn.GetString("MSGV4");
                    }
                    if (strReturn.Count > 0)
                    {
                        for (int i = 0; i < strReturn.Count; i++)
                        {
                            if (strmsgv1 != "") { CL_proc.WriteData.Create_Error(strmsgv1, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv2 != "") { CL_proc.WriteData.Create_Error(strmsgv2, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv3 != "") { CL_proc.WriteData.Create_Error(strmsgv3, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv4 != "") { CL_proc.WriteData.Create_Error(strmsgv4, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            erroresfinales = "Error al Planear la HU en SAP: " + strmsgv1 + "; " + strmsgv2 + "; " + strmsgv3 + "; " + strmsgv4;
                        }
                    }

                }
                if (erroresfinales.Equals("Error al Planear la HU en SAP: Handling units saved; ; ; ")) { erroresfinales = ""; }
                return erroresfinales;
            }
            catch (Exception ex)
            {

                // string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().Replace("FMX\\", "");
                CL_proc.WriteData.Create_Error(ex.Message.ToString().Trim(), "ImprimirEtiquetaM.PrintingAuthorization-" + ex.Source.ToString().Trim(), System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()));
                return erroresfinales;
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
            return erroresfinales;
        }
        public string PlannedHU(int Cantidad, string HU, string NoParte, string Planta, string VersionProduccion)
        {
            string excantidad = "";
            string exhu = "";
            string exnoparte = "";
            string explanta = "";
            string exversion = "";
            string Ret = "";
            string cantidad = "Q" + Cantidad.ToString();
            string nopartefinal = "P" + NoParte.ToString();
            string erroresfinales = "";
            try
            {

                //BAPI.Z_Bapi_Repprod_Labels_Planned2(cantidad, HU, nopartefinal, Planta, VersionProduccion, out excantidad
                //    , out exhu, out exnoparte, out exversion, out explanta, out  Ret, ref SAPTbl2,  ref SAPTbl  );
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_BAPI_REPPROD_LABELS_PLANNED2");
                Bapi.SetValue("ML_MATCHAR", nopartefinal);
                Bapi.SetValue("ML_CANTCHAR", cantidad.ToString());
                Bapi.SetValue("ML_HUCHAR", HU);
                Bapi.SetValue("ML_VERCHAR", VersionProduccion);
                Bapi.SetValue("ML_PLANTACHAR", Planta);
                Bapi.Invoke(SapRfcDestination);
                IRfcTable SAPTbl2 = Bapi.GetTable("EX_IT_ERRORMSGS");
                string strmsgv1 = "";
                string strmsgv2 = "";
                string strmsgv3 = "";
                string strmsgv4 = "";
                if (SAPTbl2.RowCount != 0)
                {
                    string strtcode = SAPTbl2.GetString("TCODE");
                    string strdyname = SAPTbl2.GetString("DYNAME");
                    string intNumber = SAPTbl2.GetString("DYNUMB");
                    string strm = SAPTbl2.GetString("MSGTYP");
                    string strms = SAPTbl2.GetString("MSGSPRA");
                    string strmsid = SAPTbl2.GetString("MSGID");
                    string strmsg = SAPTbl2.GetString("MSGNR");
                    strmsgv1 = SAPTbl2.GetString("MSGV1");
                    strmsgv2 = SAPTbl2.GetString("MSGV2");
                    strmsgv3 = SAPTbl2.GetString("MSGV3");
                    strmsgv4 = SAPTbl2.GetString("MSGV4");
                }
                if (SAPTbl2 != null)
                {
                    if (SAPTbl2.Count > 0)
                    {
                        for (int i = 0; i < SAPTbl2.Count; i++)
                        {
                            if (strmsgv1 != "") { CL_proc.WriteData.Create_Error(strmsgv1, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv2 != "") { CL_proc.WriteData.Create_Error(strmsgv2, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv3 != "") { CL_proc.WriteData.Create_Error(strmsgv3, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv4 != "") { CL_proc.WriteData.Create_Error(strmsgv4, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            erroresfinales = "Error al Planear la HU en SAP: " + strmsgv1 + "; " + strmsgv2 + "; " + strmsgv3 + "; " + strmsgv4;
                        }
                    }
                }
            }


            catch (Exception ex)
            {
                //Mesajes
                string error = ex.Message.ToString();
                CL_proc.WriteData.Create_Error(ex.Message.ToString().Trim(), "ImprimirEtiquetaM.PlannedHU-" + ex.Source.ToString().Trim(), System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()));

            }
            return erroresfinales;

            #region Ejemplo Web
            ////			Private Sub Button1_Click(ByVal sender As System.Object, _
            ////    ByVal e As System.EventArgs) Handles Button1.Click
            ////
            ////    Dim BAPI As New SAPProxy1
            ////    Dim conStr As String = "ASHOST=AppserverIP SYSNR=sysnumber CLIENT=clientnumber _
            ////                           USER=user PASSWD=password"
            //// 
            ////
            ////
            ////    BAPI = New SAPProxy1(conStr)
            ////
            ////    Dim SAPtbl As New ZSDUS_INVOICETable
            ////    Dim Ret As New VS2003_SAP_Connect.BAPIRETURNTable
            ////
            ////    Try
            ////        'params are contract # and line item
            ////        BAPI.Zsdus_Get_Invoice_Number("0045001126", "000010", SAPtbl, Ret)
            ////    Catch ex As Exception
            ////        MsgBox(ex.Message)
            ////    End Try
            ////
            ////    'ret table of success or failure
            ////    'ret table is always empty upon success if ret contains a
            ////    'value that indicates failure
            ////    If Ret.Count > 0
            ////    Then
            ////        MsgBox(Ret.Item(0).Message.ToString)
            ////    End If
            ////
            ////    TextBox1.Text = (SAPtbl.Item(0).Vbeln.ToString)
            ////End Sub
            #endregion

        }
        //public string ReportaSAPasPU()
        ////{
        ////    string Valor = "";


        //}

        public string ReportaSAP(string HU, string Material1, int Cant1, string Material2, int Cant2)
        {
            try
            {
                if (HU != "")
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                    //   RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                    IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_BAPI_REPPROD_LABELS_PLANNED3");
                    IRfcTable tblImport = Bapi.GetTable("T_MATERIALES");
                    tblImport.Clear();
                    int ren1 = tblImport.RowCount;
                    tblImport.Append();

                    tblImport.CurrentRow.SetValue(0, "P" + Material1);
                    tblImport.CurrentRow.SetValue(1, Cant1);
                    tblImport.CurrentRow.SetValue(2, "0001");

                    if (Cant2 != 0)
                    {
                        tblImport.Append();
                        tblImport.SetValue(0, Material2);
                        tblImport.SetValue(1, Cant2);
                        tblImport.SetValue(2, "0001");
                        //tblImport.Insert();
                    }
                    // tblImport.Insert();
                    Bapi.SetValue("ML_HUCHAR", HU);
                    // Bapi.SetValue("ML_HUCHAR", "302231494");
                    Bapi.SetValue("ML_PLANTACHAR", "1841");
                    int renglones = tblImport.RowCount;
                    Bapi.Invoke(SapRfcDestination);

                    IRfcTable strReturn = Bapi.GetTable("EX_IT_ERRORMSGS");
                    string strmsgv1 = "";
                    string strmsgv2 = "";
                    string strmsgv3 = "";
                    if (strReturn.RowCount != 0)
                    {
                        string strtcode = strReturn.GetString("TCODE");
                        string strdyname = strReturn.GetString("DYNAME");
                        string intNumber = strReturn.GetString("DYNUMB");
                        string strm = strReturn.GetString("MSGTYP");
                        string strms = strReturn.GetString("MSGSPRA");
                        string strmsid = strReturn.GetString("MSGID");
                        string strmsg = strReturn.GetString("MSGNR");
                        strmsgv1 = strReturn.GetString("MSGV1");
                        strmsgv2 = strReturn.GetString("MSGV2");
                        strmsgv3 = strReturn.GetString("MSGV3");
                    }


                    string strmsgvtot = strmsgv1 + "____" + strmsgv2 + "____" + strmsgv3;
                    bool flagerr = true;
                    if (strmsgvtot.Equals("________"))
                    {
                        flagerr = false;
                        /// aqui se agrega log de reporteo
                        // Datos.WriteData.SaveSAPReportHistory(IntIdEstacion, strModel, strmaterial, IntQty, IntTurno, proddatesql, datevalue);
                        if (Cant2 != 0)
                        {
                            //Cl_p.InsertLogSAP(Material1, Material2, Cant1, Cant2, "HU REPORTADO SATISFACTORIAMENTE", true, HU);
                            //CL_proc.WriteData.InsertLogSAP(Material1, Cant1, Material2, Cant2, "HU REPORTADO SATISFACTORIAMENTE", true, HU);
                            return "";
                        }
                        else
                        {
                            // CL_proc.WriteData.InsertLogSAP(Material1, Cant1, "", 0, "HU REPORTADO SATISFACTORIAMENTE", true, HU);
                            return "";
                        }

                    }
                    else
                    {
                        //label2.Content = strmessage.Trim();
                        // aqui se agrega el status en la columna de la tabla secuenciado_daimler.
                        //  CL_proc.WriteData.InsertLogSAP("", 0, "", 0, strmsgvtot, false, "");
                        return "Error:" + strmsgvtot;

                        //Datos.WriteData.SaveStatusSAP(IntIdEstacion, strModel, strmaterial, IntQty, IntTurno, proddatesql, datevalue, strmessage);
                    }
                }
                else
                {
                    return "Error: Falta HU";
                }
            }
            catch (Exception ex)
            {
                //throw new Exception(ex.Message.ToString());
                //label2.Content = ex.Message.ToString();




                string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().Replace("FMX\\", "");
                CL_proc.WriteData.Create_Error(ex.Message.ToString().Trim(), "SecuenciadoFR.ReportaSAPEmpacado-" + ex.Source.ToString().Trim(), System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()));
                return "Error:" + ex.Message.ToString();
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
        }
        public string NotificaHU(string HU)
        {
            string erroresfinales = "";
            try
            {
                if (HU != "")
                {
                    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                    // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                    RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                    IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZMMIM_HUBACKFLUSH");
                    Bapi.SetValue("HUCHAR", HU);
                    Bapi.SetValue("PLANTACHAR", "1841");
                    Bapi.Invoke(SapRfcDestination);
                    IRfcTable strReturn = Bapi.GetTable("EX_ERRORS_TABLE");
                    string strmsgv1 = "";
                    string strmsgv2 = "";
                    string strmsgv3 = "";
                    string strmsgv4 = "";
                    string strenv = "";
                    string strfldname = "";
                    string outError = "";
                    outError = Bapi.GetValue("EX_ERROR_FOUND").ToString();

                    if (strReturn.RowCount != 0)
                    {
                        string strtcode = strReturn.GetString("TCODE");
                        string strdyname = strReturn.GetString("DYNAME");
                        string intNumber = strReturn.GetString("DYNUMB");
                        string strm = strReturn.GetString("MSGTYP");
                        string strms = strReturn.GetString("MSGSPRA");
                        string strmsid = strReturn.GetString("MSGID");
                        string strmsg = strReturn.GetString("MSGNR");
                        strmsgv1 = strReturn.GetString("MSGV1");
                        strmsgv2 = strReturn.GetString("MSGV2");
                        strmsgv3 = strReturn.GetString("MSGV3");
                        strmsgv4 = strReturn.GetString("MSGV4");
                        strenv = strReturn.GetString("ENV");
                        strfldname = strReturn.GetString("FLDNAME");
                    }
                    if (strReturn.Count > 0)
                    {
                        for (int i = 0; i < strReturn.Count; i++)
                        {
                            if (strmsgv1 != "") { CL_proc.WriteData.Create_Error(strmsgv1, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv2 != "") { CL_proc.WriteData.Create_Error(strmsgv2, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv3 != "") { CL_proc.WriteData.Create_Error(strmsgv3, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            if (strmsgv4 != "") { CL_proc.WriteData.Create_Error(strmsgv4, "ImprimirEtiquetaM.PlannedHU.Errors", System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString())); };
                            erroresfinales = "Error al Planear la HU en SAP: " + strmsgv1 + "; " + strmsgv2 + "; " + strmsgv3 + "; " + strmsgv4;
                        }
                    }

                }
                if (erroresfinales.Equals("Error al Planear la HU en SAP: Handling units saved; ; ; ")) { erroresfinales = ""; }
                return erroresfinales;
            }
            catch (Exception ex)
            {

                // string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().Replace("FMX\\", "");
                CL_proc.WriteData.Create_Error(ex.Message.ToString().Trim(), "ImprimirEtiquetaM.PlannedHU-" + ex.Source.ToString().Trim(), System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()));
                return erroresfinales;
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
            return erroresfinales;

            //string conStr = ConfigurationSettings.AppSettings["ConSAP"];
            //SAPLabels.SAPFMX BAPI = new SAPFMX(conStr);
            //SAPLabels.BDCMSGCOLLTable SAPTbl = new BDCMSGCOLLTable();
            //SAPLabels.BDCMSGCOLLTable SAPTbl2 = new BDCMSGCOLLTable();
            //SAP.Connector.BAPIRET1 Ret1 = new SAP.Connector.BAPIRET1();

            //string excantidad = "";
            //string exhu = "";
            //string exnoparte = "";
            //string explanta = "";
            //string exversion = "";
            //string Ret = "";
            //string cantidad = "Q" + Cantidad.ToString();
            //string nopartefinal = "P" + NoParte.ToString();
            //string erroresfinales = "";
            //try
            //{
            //    BAPI.Zmmim_Hubackflush(HU, Planta, out Cantidad, out erroresfinales, out exhu, out nopartefinal, out VersionProduccion, out Planta, ref SAPTbl);
            //}
            //catch (Exception ex)
            //{
            //    string error = ex.Message.ToString();
            //    //CL_proc.WriteData.Create_Error(nopartefinal.ToString() + ";" + HU.ToString() + ";" + Planta.ToString() + ";"  +  ex.Message.ToString().Trim(), "ImprimirEtiquetaM.PlannedHU-" + ex.Source.ToString().Trim(), System.Net.Dns.GetHostName().ToString().Trim(), CL_proc.ReadData.GetUsername(System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString()), Planta);
            //}
            //return erroresfinales;

        }
        //public void CargaMB52(string Almacen, int NoEvento)
        //{
        //    RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
        //    // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
        //    RfcRepository SapRfcRepository = SapRfcDestination.Repository;
        //    IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_BAPI_NA_REEMP");
        //    Bapi.SetValue("GV_WERKS", "1841");
        //    Bapi.SetValue("GV_LGORT", Almacen);
        //    Bapi.Invoke(SapRfcDestination);
        //    IRfcTable strReturn = Bapi.GetTable("IT_ZTSD_INVENTA");
        //    DataTable DT = this.ConvertToDT(strReturn);
        //    foreach (DataRow row in DT.Rows)
        //    {
        //        InsertInventarioSAP_MB52(row["MANDT"].ToString(), row["MATNR"].ToString(), row["WERKS"].ToString(), row["LGORT"].ToString(), row["PSTAT"].ToString(), row["LVORM"].ToString(), row["LFGJA"].ToString(), row["LFMON"].ToString(), Double.Parse(row["LABST"].ToString()), Double.Parse(row["UMLME"].ToString()), Double.Parse(row["INSME"].ToString()), Double.Parse(row["EINME"].ToString()), Double.Parse(row["SPEME"].ToString()), Double.Parse(row["RETME"].ToString()), Double.Parse(row["VMLAB"].ToString()), Double.Parse(row["VMUML"].ToString()), Double.Parse(row["VMINS"].ToString()), Double.Parse(row["VMEIN"].ToString()), Double.Parse(row["VMSPE"].ToString()), row["KZILL"].ToString(), NoEvento);
        //    }
        //}

        public DataTable CargaInventarioHRxHR(DateTime Fecha, DateTime HoraIni, DateTime HoraFin, string Planta)
        {
            try
            {
                DateTime HIni = new DateTime(HoraIni.Year, HoraIni.Month, HoraIni.Day, HoraIni.Hour, 0, 0);
                if (Fecha.Hour.Equals(0))
                {
                    if (Fecha.Minute.Equals(0))
                    {
                        DateTime HI = HIni.AddDays(-1);
                        DateTime HF = HoraFin.AddDays(-1);
                        HIni = new DateTime(HI.Year, HI.Month, HI.Day, 23, 0, 0);
                        HoraFin = new DateTime(HF.Year, HF.Month, HF.Day, 23, 59, 59);
                        Fecha = Fecha.AddDays(-1);
                    }

                }

                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZMM_EXTRAE_MOVIMIENTOS_MB51");
                Bapi.SetValue("IN_WERKS", Planta);
                IRfcTable BUDAT = Bapi.GetTable("IN_T_BUDAT");
                BUDAT.Append();
                BUDAT.SetValue("SIGN_R", "I");
                BUDAT.SetValue("OPTION_R", "EQ");
                BUDAT.SetValue("LOW", Fecha);
                BUDAT.SetValue("HIGH", Fecha);

                IRfcTable CPUTM = Bapi.GetTable("IN_T_CPUTM");
                CPUTM.Append();
                CPUTM.SetValue("SIGN_R", "I");
                CPUTM.SetValue("OPTION_R", "BT");
                CPUTM.SetValue("LOW", HIni);
                CPUTM.SetValue("HIGH", HoraFin);


                IRfcTable BWWART = Bapi.GetTable("IN_T_BWART");
                BWWART.Append();

                //101
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "101");
                BWWART.SetValue("HIGH", " ");
                //102
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "102");
                BWWART.SetValue("HIGH", " ");
                //131
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "131");
                BWWART.SetValue("HIGH", " ");
                //132
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "132");
                BWWART.SetValue("HIGH", " ");
                //201
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "201");
                BWWART.SetValue("HIGH", " ");
                //202
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "202");
                BWWART.SetValue("HIGH", " ");

                //261
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "261");
                BWWART.SetValue("HIGH", " ");

                //262
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "262");
                BWWART.SetValue("HIGH", " ");

                //321
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "321");
                BWWART.SetValue("HIGH", " ");
                //322
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "322");
                BWWART.SetValue("HIGH", " ");

                //343
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "343");
                BWWART.SetValue("HIGH", " ");

                //344
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "344");
                BWWART.SetValue("HIGH", " ");

                //551
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "551");
                BWWART.SetValue("HIGH", " ");
                //552
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "552");
                BWWART.SetValue("HIGH", " ");
                //553
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "553");
                BWWART.SetValue("HIGH", " ");
                //554
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "554");
                BWWART.SetValue("HIGH", " ");
                //601
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "601");
                BWWART.SetValue("HIGH", " ");
                //602
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "602");
                BWWART.SetValue("HIGH", " ");
                //701
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "701");
                BWWART.SetValue("HIGH", " ");
                //702
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "702");
                BWWART.SetValue("HIGH", " ");
                //711
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "711");
                BWWART.SetValue("HIGH", " ");
                //712
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "702");
                BWWART.SetValue("HIGH", " ");

                Bapi.SetValue("IN_KOKRS", "CEMM");

                Bapi.Invoke(SapRfcDestination);
                IRfcTable RESULTADOS = Bapi.GetTable("EX_T_MOVS");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                string MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTMv, USNAM, BKTXT, XBLNR, MATNR, GRUND, BWART, SAKTO, LGORT, SHKZG, WEMPF;
                double DMBTR;
                string WAERS, MEINS, SGTXT, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH, AUFNR;
                double MENGE, STPRS, PEINH;
                for (int i = 0; i <= DTSalida.Rows.Count - 1; i++)
                {
                    MBLNR = DTSalida.Rows[i]["MBLNR"].ToString();
                    ZEILE = DTSalida.Rows[i]["ZEILE"].ToString();
                    BUDATv = DTSalida.Rows[i]["BUDAT"].ToString();
                    WERKS = DTSalida.Rows[i]["WERKS"].ToString();
                    BLDAT = DTSalida.Rows[i]["BLDAT"].ToString();
                    CPUTMv = DTSalida.Rows[i]["CPUTM"].ToString();
                    USNAM = DTSalida.Rows[i]["USNAM"].ToString();
                    BKTXT = DTSalida.Rows[i]["BKTXT"].ToString();
                    XBLNR = DTSalida.Rows[i]["XBLNR"].ToString();
                    MATNR = DTSalida.Rows[i]["MATNR"].ToString();
                    GRUND = DTSalida.Rows[i]["GRUND"].ToString();
                    BWART = DTSalida.Rows[i]["BWART"].ToString();
                    SAKTO = DTSalida.Rows[i]["SAKTO"].ToString();
                    LGORT = DTSalida.Rows[i]["LGORT"].ToString();
                    SHKZG = DTSalida.Rows[i]["SHKZG"].ToString();
                    DMBTR = Convert.ToDouble(DTSalida.Rows[i]["DMBTR"].ToString());
                    WAERS = DTSalida.Rows[i]["WAERS"].ToString();
                    MENGE = Convert.ToDouble(DTSalida.Rows[i]["MENGE"]);
                    MEINS = DTSalida.Rows[i]["MEINS"].ToString();
                    SGTXT = DTSalida.Rows[i]["SGTXT"].ToString();
                    WEMPF = DTSalida.Rows[i]["WEMPF"].ToString();
                    KOSTL = DTSalida.Rows[i]["KOSTL"].ToString();

                    MJAHR = DTSalida.Rows[i]["MJAHR"].ToString();
                    GRTXT = DTSalida.Rows[i]["GRTXT"].ToString();
                    ZCCNAME = DTSalida.Rows[i]["ZCCNAME"].ToString();
                    PRCTR = DTSalida.Rows[i]["PRCTR"].ToString();
                    ZPCNAME = DTSalida.Rows[i]["ZPCNAME"].ToString();
                    MDV01 = DTSalida.Rows[i]["MDV01"].ToString();
                    ZWCNAME = DTSalida.Rows[i]["ZWCNAME"].ToString();
                    KTSCH = DTSalida.Rows[i]["KTSCH"].ToString();
                    AUFNR = DTSalida.Rows[i]["AUFNR"].ToString();
                    STPRS = Convert.ToDouble(DTSalida.Rows[i]["STPRS"].ToString());
                    PEINH = Convert.ToDouble(DTSalida.Rows[i]["PEINH"].ToString());
                    this.InsertaMovInventarioSQLHrxHr(Fecha, MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTMv, USNAM, BKTXT, XBLNR, MATNR,
                       GRUND, BWART, SAKTO, LGORT, SHKZG, DMBTR, WAERS, MENGE, MEINS, SGTXT, WEMPF, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH,AUFNR,STPRS,PEINH);
                    //this.InsertaMovInventarioSQLHrxHrTEST(Fecha, MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTMv, USNAM, BKTXT, XBLNR, MATNR,
                    //   GRUND, BWART, SAKTO, LGORT, SHKZG, DMBTR, WAERS, MENGE, MEINS, SGTXT, WEMPF, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH);
                }
                return DTSalida;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public int GetUltimoEventoMB52()
        {
            int NoEvento = 0;
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["Reporte"].ToString());
            try
            {
                SqlCommand oCmd = new SqlCommand("SELECT MAX(NoEvento) from [dbo].InventarioSAP_MB52", oConn);
                SqlDataAdapter oadap = new SqlDataAdapter(oCmd);
                DataTable DT = new DataTable();
                oadap.Fill(DT);
                if (DT != null)
                {
                    if (DT.Rows.Count != 0)
                    {
                        if (DT.Rows[0][0].ToString().Equals(""))
                        {
                            NoEvento = 1;
                        }
                        else
                        {
                            NoEvento = Convert.ToInt32(DT.Rows[0][0].ToString());
                        }

                    }
                    else
                    {
                        NoEvento = 1;
                    }
                }
                else
                {
                    NoEvento = 1;
                }
            }

            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                NoEvento = 0;
            }
            finally
            {
                if (oConn.State == ConnectionState.Open) oConn.Close();
                oConn = null;

            }
            return NoEvento + 1;
        }
        public void CargaMB52(string Almacen, int NoEvento)
        {
            RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
            // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
            RfcRepository SapRfcRepository = SapRfcDestination.Repository;
            IRfcFunction Bapi = SapRfcRepository.CreateFunction("Z_BAPI_NA_REEMP");
            Bapi.SetValue("GV_WERKS", "1841");
            Bapi.SetValue("GV_LGORT", Almacen);
            Bapi.Invoke(SapRfcDestination);
            IRfcTable strReturn = Bapi.GetTable("IT_ZTSD_INVENTA");
            DataTable DT = this.ConvertToDT(strReturn);
            foreach (DataRow row in DT.Rows)
            {
                InsertInventarioSAP_MB52(row["MANDT"].ToString(), row["MATNR"].ToString(), row["WERKS"].ToString(), row["LGORT"].ToString(), row["PSTAT"].ToString(), row["LVORM"].ToString(), row["LFGJA"].ToString(), row["LFMON"].ToString(), Double.Parse(row["LABST"].ToString()), Double.Parse(row["UMLME"].ToString()), Double.Parse(row["INSME"].ToString()), Double.Parse(row["EINME"].ToString()), Double.Parse(row["SPEME"].ToString()), Double.Parse(row["RETME"].ToString()), Double.Parse(row["VMLAB"].ToString()), Double.Parse(row["VMUML"].ToString()), Double.Parse(row["VMINS"].ToString()), Double.Parse(row["VMEIN"].ToString()), Double.Parse(row["VMSPE"].ToString()), row["KZILL"].ToString(), NoEvento);
            }
        }
        public DataTable CargaInventarioHRxHR(DateTime Fecha, DateTime HoraIni, DateTime HoraFin)
        {
            try
            {
                DateTime HIni = new DateTime(HoraIni.Year, HoraIni.Month, HoraIni.Day, HoraIni.Hour,0, 0);
                if (Fecha.Hour.Equals(0))
                {
                    if(Fecha.Minute.Equals(0))
                    {
                        DateTime HI = HIni.AddDays(-1);
                        DateTime HF = HoraFin.AddDays(-1);
                        HIni = new DateTime(HI.Year, HI.Month, HI.Day, 23, 0, 0);
                        HoraFin = new DateTime(HF.Year, HF.Month, HF.Day, 23, 59, 59);
                        Fecha = Fecha.AddDays(-1);
                    }

                }
                
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZMM_EXTRAE_MOVIMIENTOS_MB51");
                Bapi.SetValue("IN_WERKS", "1841");
                IRfcTable BUDAT = Bapi.GetTable("IN_T_BUDAT");
                BUDAT.Append();
                BUDAT.SetValue("SIGN_R", "I");
                BUDAT.SetValue("OPTION_R", "EQ");
                BUDAT.SetValue("LOW", Fecha);
                BUDAT.SetValue("HIGH", Fecha);

                IRfcTable CPUTM = Bapi.GetTable("IN_T_CPUTM");
                CPUTM.Append();
                CPUTM.SetValue("SIGN_R", "I");
                CPUTM.SetValue("OPTION_R", "BT");
                CPUTM.SetValue("LOW", HIni);
                CPUTM.SetValue("HIGH", HoraFin);
               

                IRfcTable BWWART = Bapi.GetTable("IN_T_BWART");
                BWWART.Append();

                //101
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "101");
                BWWART.SetValue("HIGH", " ");
                //102
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "102");
                BWWART.SetValue("HIGH", " ");
                //131
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "131");
                BWWART.SetValue("HIGH", " ");
                //132
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "132");
                BWWART.SetValue("HIGH", " ");
                //201
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "201");
                BWWART.SetValue("HIGH", " ");
                //202
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "202");
                BWWART.SetValue("HIGH", " ");

                //261
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "261");
                BWWART.SetValue("HIGH", " ");

                //262
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "262");
                BWWART.SetValue("HIGH", " ");

                //343
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "343");
                BWWART.SetValue("HIGH", " ");

                //344
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "344");
                BWWART.SetValue("HIGH", " ");

                //541
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "541");
                BWWART.SetValue("HIGH", " ");

                //542
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "542");
                BWWART.SetValue("HIGH", " ");

                //551
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "551");
                BWWART.SetValue("HIGH", " ");
                //552
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "552");
                BWWART.SetValue("HIGH", " ");
                //553
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "553");
                BWWART.SetValue("HIGH", " ");
                //554
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "554");
                BWWART.SetValue("HIGH", " ");
                //601
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "601");
                BWWART.SetValue("HIGH", " ");
                //602
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "602");
                BWWART.SetValue("HIGH", " ");
                //701
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "701");
                BWWART.SetValue("HIGH", " ");
                //702
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "702");
                BWWART.SetValue("HIGH", " ");
                //711
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "711");
                BWWART.SetValue("HIGH", " ");
                //712
                BWWART.Append();
                BWWART.SetValue("SIGN_R", "I");
                BWWART.SetValue("OPTION_R", "EQ");
                BWWART.SetValue("LOW", "702");
                BWWART.SetValue("HIGH", " ");

                Bapi.SetValue("IN_KOKRS", "CEMM");

                Bapi.Invoke(SapRfcDestination);
                IRfcTable RESULTADOS = Bapi.GetTable("EX_T_MOVS");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                string MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTMv, USNAM, BKTXT, XBLNR, MATNR, GRUND, BWART, SAKTO, LGORT, SHKZG, WEMPF;
                double DMBTR;
                string WAERS, MEINS, SGTXT, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH,AUFNR;
                double MENGE, STPRS, PEINH;
                for (int i = 0; i <= DTSalida.Rows.Count - 1; i++)
                {
                    MBLNR = DTSalida.Rows[i]["MBLNR"].ToString();
                    ZEILE = DTSalida.Rows[i]["ZEILE"].ToString();
                    BUDATv = DTSalida.Rows[i]["BUDAT"].ToString();
                    WERKS = DTSalida.Rows[i]["WERKS"].ToString();
                    BLDAT = DTSalida.Rows[i]["BLDAT"].ToString();
                    CPUTMv = DTSalida.Rows[i]["CPUTM"].ToString();
                    USNAM = DTSalida.Rows[i]["USNAM"].ToString();
                    BKTXT = DTSalida.Rows[i]["BKTXT"].ToString();
                    XBLNR = DTSalida.Rows[i]["XBLNR"].ToString();
                    MATNR = DTSalida.Rows[i]["MATNR"].ToString();
                    GRUND = DTSalida.Rows[i]["GRUND"].ToString();
                    BWART = DTSalida.Rows[i]["BWART"].ToString();
                    SAKTO = DTSalida.Rows[i]["SAKTO"].ToString();
                    LGORT = DTSalida.Rows[i]["LGORT"].ToString();
                    SHKZG = DTSalida.Rows[i]["SHKZG"].ToString();
                    DMBTR = Convert.ToDouble(DTSalida.Rows[i]["DMBTR"].ToString());
                    WAERS = DTSalida.Rows[i]["WAERS"].ToString();
                    MENGE = Convert.ToDouble(DTSalida.Rows[i]["MENGE"]);
                    MEINS = DTSalida.Rows[i]["MEINS"].ToString();
                    SGTXT = DTSalida.Rows[i]["SGTXT"].ToString();
                    WEMPF = DTSalida.Rows[i]["WEMPF"].ToString();
                    KOSTL = DTSalida.Rows[i]["KOSTL"].ToString();
                   
                    MJAHR = DTSalida.Rows[i]["MJAHR"].ToString();
                    GRTXT = DTSalida.Rows[i]["GRTXT"].ToString();
                    ZCCNAME = DTSalida.Rows[i]["ZCCNAME"].ToString();
                    PRCTR = DTSalida.Rows[i]["PRCTR"].ToString();
                    ZPCNAME = DTSalida.Rows[i]["ZPCNAME"].ToString();
                    MDV01 = DTSalida.Rows[i]["MDV01"].ToString();
                    ZWCNAME = DTSalida.Rows[i]["ZWCNAME"].ToString();
                    KTSCH = DTSalida.Rows[i]["KTSCH"].ToString();
                    AUFNR = DTSalida.Rows[i]["AUFNR"].ToString();
                    STPRS = Convert.ToDouble(DTSalida.Rows[i]["STPRS"].ToString());
                    PEINH = Convert.ToDouble(DTSalida.Rows[i]["PEINH"].ToString());

                    this.InsertaMovInventarioSQLHrxHr(Fecha, MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTMv, USNAM, BKTXT, XBLNR, MATNR,
                       GRUND, BWART, SAKTO, LGORT, SHKZG, DMBTR, WAERS, MENGE, MEINS, SGTXT, WEMPF, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH, AUFNR, STPRS, PEINH);
                    //this.InsertaMovInventarioSQLHrxHrTEST(Fecha, MBLNR, ZEILE, BUDATv, WERKS, BLDAT, CPUTMv, USNAM, BKTXT, XBLNR, MATNR,
                    //   GRUND, BWART, SAKTO, LGORT, SHKZG, DMBTR, WAERS, MENGE, MEINS, SGTXT, WEMPF, KOSTL, MJAHR, GRTXT, ZCCNAME, PRCTR, ZPCNAME, MDV01, ZWCNAME, KTSCH);
                }
                return DTSalida;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        public void ejecuta_MD16FG()
        {
            try
            {

                int Evento = GetUltimoEvento();
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                DateTime Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZMF_MD16");
                Bapi.SetValue("I_WERKS", "1841");
                Bapi.SetValue("I_DISPO_FLAG", "X");
                Bapi.SetValue("I_MATNR_FLAG", "");
                Bapi.SetValue("I_DISPO", "PPE");
                Bapi.SetValue("I_MATNR", "");
                Bapi.SetValue("I_BISDT", Fecha.Date);

                //101
                Bapi.Invoke(SapRfcDestination);
                IRfcTable RESULTADOS = Bapi.GetTable("T_DATA");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                for (int i = 0; DTSalida.Rows.Count - 1 > i; i++)
                {
                    InsertMD16(Evento, DTSalida.Rows[i]["MATNR"].ToString(), DTSalida.Rows[i]["PERTR"].ToString(), DTSalida.Rows[i]["PSTTR"].ToString(), DTSalida.Rows[i]["PEDTR"].ToString(),
                        Convert.ToDouble(DTSalida.Rows[i]["GSMNG"].ToString()), DTSalida.Rows[i]["MEINS"].ToString(), DTSalida.Rows[i]["PLAFX"].ToString(), DTSalida.Rows[i]["BESKZ"].ToString(),
                        DTSalida.Rows[i]["ESOBS"].ToString(), DTSalida.Rows[i]["PLNUM"].ToString(), DTSalida.Rows[i]["PAART"].ToString(), DTSalida.Rows[i]["KNTTP"].ToString(),
                        DTSalida.Rows[i]["KDAUF"].ToString(), DTSalida.Rows[i]["KDPOS"].ToString());
                }

                IRfcTable MENSAJES = Bapi.GetTable("T_RESULT");
                DataTable DTMensajes = this.ConvertToDT(MENSAJES);
                for (int i = 0; DTMensajes.Rows.Count - 1 > i; i++)
                {
                    InsertaMensajeLogMD16(DTMensajes.Rows[i][0].ToString(), DTMensajes.Rows[i][1].ToString(), System.Net.Dns.GetHostName().ToString(), Evento);
                }
                // return DTSalida;
            }
            catch (Exception ex)
            {

            }
        }
        public void Ejecuta_MD16Blanks()
        {
            try
            {

                int Evento = GetUltimoEvento();
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                DateTime Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZMF_MD16");
                Bapi.SetValue("I_WERKS", "1841");
                Bapi.SetValue("I_DISPO_FLAG", "X");
                Bapi.SetValue("I_MATNR_FLAG", "");
                Bapi.SetValue("I_DISPO", "PPB");
                Bapi.SetValue("I_MATNR", "");
                Bapi.SetValue("I_BISDT", Fecha.Date);

                //101
                Bapi.Invoke(SapRfcDestination);
                IRfcTable RESULTADOS = Bapi.GetTable("T_DATA");
                DataTable DTSalida = this.ConvertToDT(RESULTADOS);
                for (int i = 0; DTSalida.Rows.Count - 1 > i; i++)
                {
                    InsertMD16(Evento, DTSalida.Rows[i]["MATNR"].ToString(), DTSalida.Rows[i]["PERTR"].ToString(), DTSalida.Rows[i]["PSTTR"].ToString(), DTSalida.Rows[i]["PEDTR"].ToString(),
                        Convert.ToDouble(DTSalida.Rows[i]["GSMNG"].ToString()), DTSalida.Rows[i]["MEINS"].ToString(), DTSalida.Rows[i]["PLAFX"].ToString(), DTSalida.Rows[i]["BESKZ"].ToString(),
                        DTSalida.Rows[i]["ESOBS"].ToString(), DTSalida.Rows[i]["PLNUM"].ToString(), DTSalida.Rows[i]["PAART"].ToString(), DTSalida.Rows[i]["KNTTP"].ToString(),
                        DTSalida.Rows[i]["KDAUF"].ToString(), DTSalida.Rows[i]["KDPOS"].ToString());
                }

                IRfcTable MENSAJES = Bapi.GetTable("T_RESULT");
                DataTable DTMensajes = this.ConvertToDT(MENSAJES);
                for (int i = 0; DTMensajes.Rows.Count - 1 > i; i++)
                {
                    InsertaMensajeLogMD16(DTMensajes.Rows[i][0].ToString(), DTMensajes.Rows[i][1].ToString(), System.Net.Dns.GetHostName().ToString(), Evento);
                }
                // return DTSalida;
            }
            catch (Exception ex)
            {

            }
        }
        public string deliverycheck(string strpartnumber, string strDelivery, string strplant)
        {
            string ErrorT = "false";
            //bool ErrorT = "false";
            try
            {
                string StrERROR = "";
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZSDSH_DELIVERYCHECK");
                Bapi.SetValue("DELIVERY", strDelivery);
                Bapi.SetValue("MATERIAL", strpartnumber);
                Bapi.SetValue("PLANT", strplant);

                Bapi.Invoke(SapRfcDestination);
                IRfcTable strReturn = Bapi.GetTable("ERRORS_TABLE");
                string strmsgv1 = "";
                string strmsgv2 = "";
                string strmsgv3 = "";
                StrERROR = Bapi.GetValue("ERROR_FOUND").ToString();
                if (strReturn.RowCount != 0)
                {
                    string strtcode = strReturn.GetString("TCODE");
                    string strdyname = strReturn.GetString("DYNAME");
                    string intNumber = strReturn.GetString("DYNUMB");
                    string strm = strReturn.GetString("MSGTYP");
                    string strms = strReturn.GetString("MSGSPRA");
                    string strmsid = strReturn.GetString("MSGID");
                    string strmsg = strReturn.GetString("MSGNR");
                    strmsgv1 = strReturn.GetString("MSGV1");
                    strmsgv2 = strReturn.GetString("MSGV2");
                    strmsgv3 = strReturn.GetString("MSGV3");

                    if (StrERROR.Equals("X"))
                    {

                        //deliverycheck=false;

                        ErrorT = strmsgv2;
                        //MessageBox.Show (ErrorT);
                        //this.Mensajes1.Text = ErrorT;
                        //ErrorT = "false";
                    }
                    else
                    {
                        ErrorT = "true";
                        //deliverycheck=true;
                    }
                }
                else
                {
                    ErrorT = "true";
                }
                return ErrorT;
            }
            catch (Exception ex)
            {

                string username = System.Security.Principal.WindowsIdentity.GetCurrent().Name.ToString().Replace("FMX\\", "");
                return ErrorT;
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
        }

        public void Carga_Z_SHIPING(int Dias, string I_CRIT, string I_EXCL, string I_SUMA, DateTime FechaIni, DateTime FechaFin)
        {
            String Host = Dns.GetHostName();
            DataTable dtHeader1 = new DataTable();
            DataTable dtHeader2 = new DataTable();
            DataTable dtDetail = new DataTable();
            try
            {
                if (EjecutaZ_Shipping(Dias, I_CRIT, I_EXCL, I_SUMA, FechaIni, FechaFin, ref dtHeader1, ref dtHeader2, ref dtDetail))
                {
                    Int32 NoEvento = Consulta_Z_ShippingH1_NoEvento_Actual() + 1;
                    if (NoEvento > 0)
                    {
                        foreach (DataRow row in dtHeader1.Rows)
                        {
                            InsertZShippingH1(NoEvento, Host, row["KUNNR"].ToString(), row["PRCTR"].ToString(), row["NAME1"].ToString(), int.Parse(row["ITEM"].ToString()), row["MATNR"].ToString(), row["KDMAT"].ToString(), row["ARKTX"].ToString(), double.Parse(row["SLXX"].ToString()), double.Parse(row["LABST"].ToString()), double.Parse(row["INSME"].ToString()), double.Parse(row["PASTD"].ToString()), int.Parse(row["ENVIO"].ToString()), row["KUNWE"].ToString());
                        }
                        foreach (DataRow row in dtHeader2.Rows)
                        {
                            InsertZShippingH2(NoEvento, Host, row["MATNR"].ToString(), row["KDMAT"].ToString(), row["ARCTX"].ToString(), double.Parse(row["SLXX"].ToString()), double.Parse(row["LABST"].ToString()), double.Parse(row["INSME"].ToString()), double.Parse(row["PASTD"].ToString()), int.Parse(row["ENVIO"].ToString()), true, row["KUNWE"].ToString());
                        }
                        foreach (DataRow row in dtDetail.Rows)
                        {
                            InsertZShippingDetail(NoEvento, Host, row["MATNR"].ToString(), row["KDMAT"].ToString(), row["ARKTX"].ToString(), DateTime.Parse(row["DATUM"].ToString()), int.Parse(row["QUANTITY"].ToString()), true, row["KUNWE"].ToString());
                        }
                    }
                    else
                    {
                        Console.WriteLine("Z_SHIPING: No se encontro el ultimó Numero de Evento.");
                    }
                    if (EjecutaZ_ShippingH2(Dias, I_CRIT, I_EXCL, "X", FechaIni, FechaFin, ref dtHeader1, ref dtHeader2, ref dtDetail))
                    {
                        foreach (DataRow row in dtHeader2.Rows)
                        {
                            InsertZShippingH2(NoEvento, Host, row["MATNR"].ToString(), row["KDMAT"].ToString(), row["ARKTX"].ToString(), double.Parse(row["SLXX"].ToString()), double.Parse(row["LABST"].ToString()), double.Parse(row["INSME"].ToString()), double.Parse(row["PASTD"].ToString()), int.Parse(row["ENVIO"].ToString()), true, row["KUNWE"].ToString());
                        }
                    }
                }
                else
                {
                    Console.WriteLine("Z_SHIPING: Ocurrio un error al obtener los datos.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Z_SHIPING: Ocurrio un error al almacenar los datos.");
            }
       }

        public bool EjecutaZ_ShippingH2(int Dias, string I_CRIT, string I_EXCL, string I_SUMA, DateTime FechaIni, DateTime FechaFin, ref DataTable dtHeader1, ref DataTable dtHeader2, ref DataTable dtDetail)
        {
            try
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZFM_SHIPPING");
                if (Dias > 0) Bapi.SetValue("I_DIAS", Dias);
                if (I_CRIT != "") Bapi.SetValue("I_CRIT", I_CRIT);
                if (I_EXCL != "") Bapi.SetValue("I_EXCL", I_EXCL);
                if (I_SUMA != "") Bapi.SetValue("I_SUMA", I_SUMA);

                IRfcTable BWWART = Bapi.GetTable("IT_WERKS");
                BWWART.Append();
                BWWART.SetValue("WERKS", "1841");

                IRfcTable BUDAT = Bapi.GetTable("IT_DATUM");
                BUDAT.Append();
                BUDAT.SetValue("DATUM", FechaIni);
                BUDAT.Append();
                BUDAT.SetValue("DATUM", FechaFin);

                Bapi.Invoke(SapRfcDestination);

                IRfcTable header1 = Bapi.GetTable("IT_HEADER1");
                IRfcTable header2 = Bapi.GetTable("IT_HEADER2");
                IRfcTable detail = Bapi.GetTable("IT_DETAIL");

                dtHeader1 = this.ConvertToDT(header1);
                dtHeader2 = this.ConvertToDT(header2);
                dtDetail = this.ConvertToDT(detail);

                return true;
            }
            catch (Exception ex)
            {
                return false;
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
        }



        public bool EjecutaZ_Shipping(int Dias, string I_CRIT, string I_EXCL, string I_SUMA, DateTime FechaIni, DateTime FechaFin, ref DataTable dtHeader1, ref DataTable dtHeader2, ref DataTable dtDetail)
        {
            try
            {
                RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination(Entorno);
                // RfcDestination SapRfcDestination = RfcDestinationManager.GetDestination("QAS");
                RfcRepository SapRfcRepository = SapRfcDestination.Repository;
                IRfcFunction Bapi = SapRfcRepository.CreateFunction("ZFM_SHIPPING");
                if (Dias > 0) Bapi.SetValue("I_DIAS", Dias);
                if (I_CRIT != "") Bapi.SetValue("I_CRIT", I_CRIT);
                if (I_EXCL != "") Bapi.SetValue("I_EXCL", I_EXCL);
                if (I_SUMA != "") Bapi.SetValue("I_SUMA", I_SUMA);

                IRfcTable BWWART = Bapi.GetTable("IT_WERKS");
                BWWART.Append();
                BWWART.SetValue("WERKS", "1841");

                IRfcTable BUDAT = Bapi.GetTable("IT_DATUM");
                BUDAT.Append();
                BUDAT.SetValue("DATUM", FechaIni);
                BUDAT.Append();
                BUDAT.SetValue("DATUM", FechaFin);

                Bapi.Invoke(SapRfcDestination);

                IRfcTable header1 = Bapi.GetTable("IT_HEADER1");
                IRfcTable header2 = Bapi.GetTable("IT_HEADER2");
                IRfcTable detail = Bapi.GetTable("IT_DETAIL");

                dtHeader1 = this.ConvertToDT(header1);
                dtHeader2 = this.ConvertToDT(header2);
                dtDetail = this.ConvertToDT(detail);

                return true;
            }
            catch(Exception ex)
            {
                return false;
                // MessageBox.Show("Error calling SAP RFC \n" + ex.ToString(), "Problem with SAP Search Synch");
                //BLogics.Error.CreateErrorLog(ex.Message.ToString(), "public static void SapJsns(int IntIdEstacion, string strmaterial, int intQty))", "BLogics.SAPReport", username);
            }
        }

        public void InsertZShippingH1(int NoEvento,string Host, string KUNNR, string PRCTR, string NAME1, int ITEM, string MATNR, string KDMAT, string ARKTX, double SLXX, double LABST, double INSME, double PASTD, int ENVIO, String KUNWE)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
            try
            {
                SqlCommand oCmd = new SqlCommand("dbo.Insert_ZshippingH1", oConn);
                oCmd.CommandType = CommandType.StoredProcedure;
                oCmd.Parameters.Add(new SqlParameter("@NoEvento", NoEvento));
                oCmd.Parameters.Add(new SqlParameter("@Host", Host));
                oCmd.Parameters.Add(new SqlParameter("@KUNNR", KUNNR));
                oCmd.Parameters.Add(new SqlParameter("@PRCTR", PRCTR));
                oCmd.Parameters.Add(new SqlParameter("@NAME1", NAME1));
                oCmd.Parameters.Add(new SqlParameter("@ITEM", ITEM));
                oCmd.Parameters.Add(new SqlParameter("@MATNR", MATNR));
                oCmd.Parameters.Add(new SqlParameter("@KDMAT", KDMAT));
                oCmd.Parameters.Add(new SqlParameter("@ARKTX", ARKTX));
                oCmd.Parameters.Add(new SqlParameter("@SLXX", SLXX));
                oCmd.Parameters.Add(new SqlParameter("@LABST", LABST));
                oCmd.Parameters.Add(new SqlParameter("@INSME", INSME));
                oCmd.Parameters.Add(new SqlParameter("@PASTD", PASTD));
                oCmd.Parameters.Add(new SqlParameter("@ENVIO", ENVIO));
                oCmd.Parameters.Add(new SqlParameter("@KUNWE", KUNWE));
                oConn.Open();
                oCmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (oConn.State == ConnectionState.Open) oConn.Close();
                oConn = null;
            }
        }

        public void InsertZShippingH2(Int32 NoEvento, String Host, String MATNR, String KDMAT, String ARKTX, double SLXX, double LABST, double INSME, double PASTD, Int32 ENVIO, Boolean Estado, String KUNWE)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
            SqlCommand command = new SqlCommand("dbo.Insert_Z_ShippingH2", oConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@NoEvento", NoEvento);
            command.Parameters.AddWithValue("@Host", Host);
            command.Parameters.AddWithValue("@MATNR", MATNR);
            command.Parameters.AddWithValue("@KDMAT", KDMAT);
            command.Parameters.AddWithValue("@ARKTX", ARKTX);
            command.Parameters.AddWithValue("@SLXX", SLXX);
            command.Parameters.AddWithValue("@LABST", LABST);
            command.Parameters.AddWithValue("@INSME", INSME);
            command.Parameters.AddWithValue("@PASTD", PASTD);
            command.Parameters.AddWithValue("@ENVIO", ENVIO);
            command.Parameters.AddWithValue("@Estado", Estado);
            command.Parameters.AddWithValue("@KUNWE", KUNWE);
            try
            {
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (command.Connection.State == ConnectionState.Open) command.Connection.Close();
                oConn = null;
            }
        }

        public void InsertZShippingDetail(Int32 NoEvento, String Host, String MATNR, String KDMAT, String ARKTX, DateTime DATUM, Int32 QUANTITY, Boolean Estado, String KUNWE)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
            SqlCommand command = new SqlCommand("dbo.Insert_Z_ShippingDetail", oConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@NoEvento", NoEvento);
            command.Parameters.AddWithValue("@Host", Host);
            command.Parameters.AddWithValue("@MATNR", MATNR);
            command.Parameters.AddWithValue("@KDMAT", KDMAT);
            command.Parameters.AddWithValue("@ARKTX", ARKTX);
            command.Parameters.AddWithValue("@DATUM", DATUM);
            command.Parameters.AddWithValue("@QUANTITY", QUANTITY);
            command.Parameters.AddWithValue("@Estado", Estado);
            command.Parameters.AddWithValue("@KUNWE", KUNWE);
            try
            {
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (command.Connection.State == ConnectionState.Open) command.Connection.Close();
                oConn = null;
            }
        }

        public Int32 Consulta_Z_ShippingH1_NoEvento_Actual()
        {
            Int32 evento_actual = 0;
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["ReportesDB"].ToString());
            try
            {
                SqlCommand oCmd = new SqlCommand("SELECT MAX(NoEvento) from Z_ShippingH1 where Estado=1", oConn);
                oCmd.CommandType = CommandType.Text;
                oConn.Open();
                Int32.TryParse(oCmd.ExecuteScalar().ToString(), out evento_actual);
                return evento_actual;
            }
            catch (Exception ex)
            {
                return -1;
            }
            finally
            {
                if (oConn.State == ConnectionState.Open) oConn.Close();
            }
        }

        public void InsertInventarioSAP_MB52(String MAN, String MATNR, String WERK, String LGORT, String PSTAT, String LVORM, String LFGJA, String LFMON, Double LABST, Double UMLME, Double INSME, Double EINME, Double SPEME, Double RETME, Double VMLAB, Double VMUML, Double VMINS, Double VMEIN, Double VMSPE, String KZILL, int NoEvento)
        {
            SqlConnection oConn = new SqlConnection(ConfigurationSettings.AppSettings["Reporte"].ToString());
            SqlCommand command = new SqlCommand("dbo.Insert_InventarioSAP_MB52", oConn);
            command.CommandType = System.Data.CommandType.StoredProcedure;
            command.Parameters.AddWithValue("@MAN", MAN);
            command.Parameters.AddWithValue("@MATNR", MATNR);
            command.Parameters.AddWithValue("@WERK", WERK);
            command.Parameters.AddWithValue("@LGORT", LGORT);
            command.Parameters.AddWithValue("@PSTAT", PSTAT);
            command.Parameters.AddWithValue("@LVORM", LVORM);
            command.Parameters.AddWithValue("@LFGJA", LFGJA);
            command.Parameters.AddWithValue("@LFMON", LFMON);
            command.Parameters.AddWithValue("@LABST", LABST);
            command.Parameters.AddWithValue("@UMLME", UMLME);
            command.Parameters.AddWithValue("@INSME", INSME);
            command.Parameters.AddWithValue("@EINME", EINME);
            command.Parameters.AddWithValue("@SPEME", SPEME);
            command.Parameters.AddWithValue("@RETME", RETME);
            command.Parameters.AddWithValue("@VMLAB", VMLAB);
            command.Parameters.AddWithValue("@VMUML", VMUML);
            command.Parameters.AddWithValue("@VMINS", VMINS);
            command.Parameters.AddWithValue("@VMEIN", VMEIN);
            command.Parameters.AddWithValue("@VMSPE", VMSPE);
            command.Parameters.AddWithValue("@KZILL", KZILL);
            command.Parameters.AddWithValue("@NoEvento", NoEvento);
            try
            {
                command.Connection.Open();
                command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (command.Connection.State == ConnectionState.Open) command.Connection.Close();
                oConn = null;
            }
        }
    }
}

