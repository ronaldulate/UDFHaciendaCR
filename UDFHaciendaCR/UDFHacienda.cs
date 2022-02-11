using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace UDFHaciendaCR
{
    static class DefaultValue{

        //Tramos de renta Año 2022
        public const decimal TramoMin = 0.0m;
        public const decimal Tramo1 = 5286000.00m;
        public const decimal Tramo2 = 7930000.00m;
        public const decimal Tramo3 = 10573000.00m;
        public const decimal TramoMax = 112070000.00m;
        public const string TramosdeRenta = "Tramos de renta para el 2022";
        

        public const decimal PorcentajeTramo1 = 0.05m;
        public const decimal PorcentajeTramo2 = 0.1m;
        public const decimal PorcentajeTramo3 = 0.15m;
        public const decimal PorcentajeSobreTramo3 = 0.20m;
        public const decimal PorcentajeTramoMax = 0.30m;

        public const string  nombreTramoMin =  "TramoMin";
        public const string nombreTramo1 = "Tramo1";
        public const string nombreTramo2 = "Tramo2";
        public const string nombreTramo3 = "Tramo3";
        public const string nombreSobreTramo3 = "SobreTramo3";
        public const string nombreTramoMax = "TramoMax";

    }

    public static class MyFunctions
    {
        
        //Funciones de Excel
   
        [ExcelFunction(Description = "Calcula el impuesto de renta para personas jurídicas de acuerdo a los tramos de renta y tasa de impuesto ingreso.")]
        public static decimal CalcularImpuestoRentaPersonaJuridica(
            decimal RentaNetaAnual,
            decimal MontoTramo1,
            decimal MontoTramo2,
            decimal MontoTramo3,
            decimal MontoTramoMax,
            decimal PorcentajeTramo1,
            decimal PorcentajeTramo2,
            decimal PorcentajeTramo3,
            decimal PorcentajeSobreTramo3,
            decimal PorcentajeTramoMax
            ) {

            decimal impuestoRenta = 0m;
            decimal saldoRenta = 0m;
            decimal rentaNeta = 0m;
            decimal excesoTramo = 0m;

            decimal tramoRenta1;
            decimal tramoRenta2;
            decimal tramoRenta3;
            decimal tramoRentaMax;

            decimal porcentajeTramo1;
            decimal porcentajeTramo2;
            decimal porcentajeTramo3;
            decimal porcentajeSobreTramo3;
            decimal porcentajeTramoMax;

            //Validando datos:

            rentaNeta = Convert.ToDecimal(RentaNetaAnual);

            tramoRenta1 = Convert.ToDecimal(MontoTramo1);
            tramoRenta2 = Convert.ToDecimal(MontoTramo2);
            tramoRenta3 = Convert.ToDecimal(MontoTramo3);
            tramoRentaMax = Convert.ToDecimal(MontoTramoMax);

            tramoRenta1 = tramoRenta1 > 0 ? tramoRenta1 : DefaultValue.Tramo1;
            tramoRenta2 = tramoRenta2 > 0 ? tramoRenta2 : DefaultValue.Tramo2;
            tramoRenta3 = tramoRenta3 > 0 ? tramoRenta3 : DefaultValue.Tramo3;
            tramoRentaMax = tramoRentaMax > 0 ? tramoRentaMax : DefaultValue.TramoMax;


            porcentajeTramo1 = Convert.ToDecimal(PorcentajeTramo1);
            porcentajeTramo2 = Convert.ToDecimal(PorcentajeTramo2);
            porcentajeTramo3 = Convert.ToDecimal(PorcentajeTramo3);
            porcentajeSobreTramo3 = Convert.ToDecimal(PorcentajeSobreTramo3);
            porcentajeTramoMax = Convert.ToDecimal(PorcentajeTramoMax);

            porcentajeTramo1 = porcentajeTramo1 > 0 ? porcentajeTramo1 : DefaultValue.PorcentajeTramo1;
            porcentajeTramo2 = porcentajeTramo2 > 0 ? porcentajeTramo2 : DefaultValue.PorcentajeTramo2;
            porcentajeTramo3 = porcentajeTramo3 > 0 ? porcentajeTramo3 : DefaultValue.PorcentajeTramo3;
            porcentajeSobreTramo3 = porcentajeSobreTramo3 > 0 ? porcentajeSobreTramo3 : DefaultValue.PorcentajeSobreTramo3;
            porcentajeTramoMax = porcentajeTramoMax > 0 ? porcentajeTramoMax : DefaultValue.PorcentajeTramoMax;

            //Calculo de Impuesto:

            //Validando condición 1:
            if (rentaNeta > tramoRentaMax)
            {
                impuestoRenta = rentaNeta * porcentajeTramoMax;
            }
            else {

                //Tramos de renta: Tabla

                saldoRenta = rentaNeta;

                //Exceso Tramo 3
                if (saldoRenta > tramoRenta3) {
                    excesoTramo = saldoRenta - tramoRenta3;
                    impuestoRenta = impuestoRenta + (excesoTramo * porcentajeSobreTramo3);
                    saldoRenta = tramoRenta3;
                }

                //Exceso Tramo 2
                if (saldoRenta > tramoRenta2)
                {
                    excesoTramo = saldoRenta - tramoRenta2;
                    impuestoRenta = impuestoRenta + (excesoTramo * porcentajeTramo3);
                    saldoRenta = tramoRenta2;
                }

                //Exceso Tramo 1
                if (saldoRenta > tramoRenta1)
                {
                    excesoTramo = saldoRenta - tramoRenta1;
                    impuestoRenta = impuestoRenta + (excesoTramo * porcentajeTramo2);
                    saldoRenta = tramoRenta1;
                }

                if (saldoRenta <= tramoRenta1) {
                    impuestoRenta = impuestoRenta + (saldoRenta * porcentajeTramo1);
                    saldoRenta = 0m;
                }


            }

            return impuestoRenta;
        }

        [ExcelFunction(Description = "Para una renta neta anual seleccionada, calcula el exceso del Tramo indicado. Nombres de tramos requeridos: Tramo1 | Tramo2 | Tramo3 | SobreTramo3 | TramoMax")]
        public static decimal CalculaExcesoTramo(decimal rentaNetaAnual, string Tramo) {

            decimal exceso = 0.0m;
            decimal rentaNeta = Convert.ToDecimal(rentaNetaAnual);


            switch (Tramo) {

                case DefaultValue.nombreTramoMax:

                    if (rentaNeta > DefaultValue.TramoMax) {
                        exceso = rentaNeta;
                    }
                    break;
                case DefaultValue.nombreSobreTramo3:

                    if (rentaNeta <= DefaultValue.TramoMax)
                    {
                        exceso = ExcesoEntreTramos(rentaNeta, DefaultValue.TramoMax, DefaultValue.Tramo3);
                    }

                    break;
                case DefaultValue.nombreTramo3:

                    if (rentaNeta <= DefaultValue.TramoMax) {
                        exceso = ExcesoEntreTramos(rentaNeta, DefaultValue.Tramo3, DefaultValue.Tramo2);
                    }

                    break;
                case DefaultValue.nombreTramo2:
                    if (rentaNeta <= DefaultValue.TramoMax)
                    {
                        exceso = ExcesoEntreTramos(rentaNeta, DefaultValue.Tramo2, DefaultValue.Tramo1);
                    }

                    break;
                case DefaultValue.nombreTramo1:
                    if (rentaNeta <= DefaultValue.TramoMax)
                    {
                        exceso = ExcesoEntreTramos(rentaNeta, DefaultValue.Tramo1, 0m);
                    }

                    break;

            }
            
            return exceso;
        }


        //Utilidades:
        public static bool IsDecimal(decimal Number) {

            return Convert.ToDecimal(Number) > 0;
        }

        public static decimal ExcesoEntreTramos(decimal renta, decimal tramosuperior, decimal tramoinferior) {

            decimal exceso=0.0m;

            if (renta > tramoinferior) {

                if (renta > tramosuperior) {
                    exceso = tramosuperior - tramoinferior;
                }
                else {
                    exceso = renta - tramoinferior;
                }
            
            }
             return exceso;
        }

    }
}
