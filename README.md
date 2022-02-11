# UDFHaciendaCR



CalcularImpuestoRentaPersonaJuridica(RentaNetaAnual):  Calcula el impuesto sobre la renta enviando como parámetro la celda que contiene la renta anual.

  Utiliza por defecto los valores del periodo fiscal 2022.
  
  
  
  
 Las personas jurídicas, cuya renta bruta no supere la suma de ciento doce millones
setenta mil colones (¢112.070.000,00) durante el periodo fiscal:

i. Cinco por ciento (5%), sobre los primeros cinco millones doscientos ochenta y
seis mil colones (¢5.286.000,00) de renta neta anual.
ii. Diez por ciento (10%), sobre el exceso de cinco millones doscientos ochenta y
seis mil colones (¢5.286.000,00) y hasta siete millones novecientos treinta mil
colones (¢7.930.000,00) de renta neta anual.
iii. Quince por ciento (15%), sobre el exceso de siete millones novecientos treinta
mil colones (¢7.930.000,00) y hasta diez millones quinientos setenta y tres mil
colones (¢10.573.000,00) de renta neta anual.
iv. Veinte por ciento (20%), sobre el exceso de diez millones quinientos setenta y
tres mil colones (¢10.573.000,00) de renta neta anual.


Se pueden indicar otros valores personalizados de la siguiente manera:

CalcularImpuestoRentaPersonaJuridica(RentaNetaAnual, tramo1, tramo2, tramo3, tramoMaximo, porcentajeTramo1,  porcentajeTramo2, porcentajeTramo3, porcentajeSobreTramo3, porcentajeTramoMax)

  por ejemplo para los valores del 2021:
  
  CalcularImpuestoRentaPersonaJuridica(RentaNetaAnual, 5157000, 7737000, 10315000, 109337000, 0.05,  0.10, 0.15, 0.20, 0.30)
  
    
