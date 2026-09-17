import pandas as pd
import numpy as np

# Generamos el Excel desordenado que te entrega el cliente
datos = {
    'Fecha': ['2026-08-01', '2026-08-01', '2026-08-02', '2026-08-02', '2026-08-03', '2026-08-03'],
    'Sucursal': [' centro ', 'NORTE', 'Centro', np.nan, 'norte', 'CENTRO'],
    'Cliente': [' Juan Perez ', 'MARIA GOMEZ', 'juan perez', 'Carlos Ruiz', 'MARIA GOMEZ', np.nan],
    'Monto': [15000, 23000, np.nan, 12000, 23000, 8000],
    'Metodo_Pago': ['Efectivo', 'Tarjeta', 'Efectivo', 'Transferencia', 'Tarjeta', 'Efectivo']
}

df_suculento = pd.DataFrame(datos)
df_suculento.to_excel('ventas_sucursales.xlsx', index=False)
print("¡Archivo 'ventas_sucursales.xlsx' generado con éxito!")