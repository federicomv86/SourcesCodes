using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MB51_TO_SQLHrxHr
{
    class Program
    {
        
        static void Main(string[] args)
        {
            SAP30 SP = new SAP30();
            //DateTime HoraFin= new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day,DateTime.Now.Hour, 0, 0);
            ////  DateTime HoraFin = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0); 
            //DateTime Fecha = new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day, 0, 0, 0);
            ////SP.CargaInventarioHRxHR(DateTime.Now, DateTime.Now.AddHours(-2),HoraFin );
           
            //SP.CargaMB52("EM08", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM08");

            //SP.GeneracionDocumentosSAP(DateTime.Now, DateTime.Now, 1, "TEST", "WH05","");
            //SP.GeneracionDocumentosSAP(DateTime.Now, DateTime.Now, 1, "TEST", "EM01", "");
            //SP.GeneracionDocumentosSAP(DateTime.Now, DateTime.Now, 1, "TEST", "B001", "");
            //SP.GeneracionDocumentosSAP(DateTime.Now, DateTime.Now, 1, "TEST", "B002", "");

            Console.WriteLine("Inicio de inventario actualizacion por minutos!!!");
            SP.CargaInventarioHRxHR(DateTime.Now, DateTime.Now.AddHours(-1), DateTime.Now, "0547");

            //Console.WriteLine("Termina de ejecuarse transaccion Inventario");

            //SP.Carga_Z_SHIPING(0, "", "X", "", DateTime.Now, DateTime.Now.AddDays(7));
            //SP.Carga_Z_SHIPING(0, "", "", "", DateTime.Now, DateTime.Now.AddDays(13));




            //Console.WriteLine("Ejecutando MD16");
            //SP.Ejecuta_MD16();
            //SP.ejecuta_MD16FG();
            //SP.Ejecuta_MD16Blanks();
            //Console.WriteLine("Termina de ejecuarse transaccion MD16");
            //Console.WriteLine("...");
            //Console.WriteLine("Inicio de inventario actualizacion por minutos!!!");
            ////SP.CargaInventarioHRxHR(DateTime.Now, DateTime.Now.AddHours(-1), DateTime.Now);
            //Console.WriteLine("Termina de ejecuarse transaccion Inventario");
            //Console.WriteLine("....");
            //Console.WriteLine("Inicio de ejecucion de ZSHIPPING Excluyendo Partes SPO");
            //SP.Carga_Z_SHIPING(0, "", "X", "", DateTime.Now, DateTime.Now.AddDays(7));
            //Console.WriteLine("Termina de ejecuarse  ZSHIPPING Excluyendo Partes SPO");
            //Console.WriteLine("....");
            //Console.WriteLine("Inicio de MB52");

            //int NoEventoMB52 = SP.GetUltimoEventoMB52();
            //// SP.CargaMB52("EM17");
            ////       int NoEventoMB52 = SP.GetUltimoEventoMB52();
            //SP.CargaMB52("CIMS", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: CIMS");
            //SP.CargaMB52("E101", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E101");
            //SP.CargaMB52("E102", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E102");
            //SP.CargaMB52("E103", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E103");
            //SP.CargaMB52("E104", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E104");
            //SP.CargaMB52("E105", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E105");
            //SP.CargaMB52("E106", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E106");
            //SP.CargaMB52("E107", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E107");
            //SP.CargaMB52("E108", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E108");
            //SP.CargaMB52("E109", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E109");
            //SP.CargaMB52("E110", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E110");
            //SP.CargaMB52("E111", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: E111");
            //SP.CargaMB52("EM01", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM01");
            //SP.CargaMB52("EM02", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM02");
            //SP.CargaMB52("EM03", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM03");
            //SP.CargaMB52("EM04", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM04");
            //SP.CargaMB52("EM05", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM05");
            //SP.CargaMB52("EM06", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM06");
            //SP.CargaMB52("EM07", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM07");
            //SP.CargaMB52("EM08", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM08");
            //SP.CargaMB52("EM09", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM09");
            //SP.CargaMB52("EM10", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM10");
            //SP.CargaMB52("EM11", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM11");
            //SP.CargaMB52("EM12", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM12");
            //SP.CargaMB52("EM13", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM13");
            //SP.CargaMB52("EM14", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM14");
            //SP.CargaMB52("EM15", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM15");
            //SP.CargaMB52("EM17", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM17");
            //SP.CargaMB52("EM21", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM21");
            //SP.CargaMB52("EM23", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM23");
            //SP.CargaMB52("EM24", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM24");
            //SP.CargaMB52("EM25", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM25");
            //SP.CargaMB52("EM26", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM26");
            //SP.CargaMB52("EM27", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM27");
            //SP.CargaMB52("EM33", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM33");
            //SP.CargaMB52("EM35", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM35");
            //SP.CargaMB52("EM36", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM36");
            //SP.CargaMB52("EM37", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM37");
            //SP.CargaMB52("EM43", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM43");
            //SP.CargaMB52("EM45", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM45");
            //SP.CargaMB52("EM46", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM46");
            //SP.CargaMB52("EM47", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM47");
            //SP.CargaMB52("EM90", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM90");
            //SP.CargaMB52("EM91", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM91");
            //SP.CargaMB52("EM92", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: EM92");
            //SP.CargaMB52("WH05", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: WH05");
            //SP.CargaMB52("B001", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: B001");
            //SP.CargaMB52("B002", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: B002");
            //SP.CargaMB52("B003", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: B003");
            //SP.CargaMB52("Y001", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y001");
            //SP.CargaMB52("Y002", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y002");
            //SP.CargaMB52("Y003", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y003");
            //SP.CargaMB52("Y004", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y004");
            //SP.CargaMB52("Y005", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y005");
            //SP.CargaMB52("Y006", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y006");
            //SP.CargaMB52("Y007", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y007");
            //SP.CargaMB52("Y008", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y008");
            //SP.CargaMB52("Y009", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y009");
            //SP.CargaMB52("Y010", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y010");
            //SP.CargaMB52("Y011", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y011");
            //SP.CargaMB52("Y012", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y012");
            //SP.CargaMB52("Y013", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y013");
            //SP.CargaMB52("Y014", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y014");
            //SP.CargaMB52("Y015", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y015");
            //SP.CargaMB52("Y016", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y016");
            //SP.CargaMB52("Y017", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y017");
            //SP.CargaMB52("Y018", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y018");
            //SP.CargaMB52("Y019", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y019");
            //SP.CargaMB52("Y020", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y020");
            //SP.CargaMB52("Y021", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y021");
            //SP.CargaMB52("Y022", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y022");
            //SP.CargaMB52("Y023", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y023");
            //SP.CargaMB52("Y024", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y024");
            //SP.CargaMB52("Y025", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y025");
            //SP.CargaMB52("Y026", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y026");
            //SP.CargaMB52("Y027", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y027");
            //SP.CargaMB52("Y028", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y028");
            //SP.CargaMB52("Y029", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y029");
            //SP.CargaMB52("Y030", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y030");
            //SP.CargaMB52("Y031", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y031");
            //SP.CargaMB52("Y032", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y032");
            //SP.CargaMB52("Y033", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y033");
            //SP.CargaMB52("Y034", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y034");
            //SP.CargaMB52("Y035", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y035");
            //SP.CargaMB52("Y036", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y036");
            //SP.CargaMB52("Y037", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y037");
            //SP.CargaMB52("Y038", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y038");
            //SP.CargaMB52("Y039", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y039");
            //SP.CargaMB52("Y040", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y040");
            //SP.CargaMB52("Y041", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y041");
            //SP.CargaMB52("Y042", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y042");
            //SP.CargaMB52("Y043", NoEventoMB52);
            //Console.WriteLine("Cargado Almacen: Y043");

            //Console.WriteLine("Termina de ejecuarse transaccion MB52");
            Console.WriteLine("Termino de programa!!!");

            //   SP.CargaInventarioHRxHR(Fecha, DateTime.Now.AddHours(-1), HoraFin);

        }
    }
}
