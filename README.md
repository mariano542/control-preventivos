# Control de Mantenimiento Preventivo

App web para procesar los reportes de CONSUMAN y generar el Excel de control.

## Cómo usar

1. Exportá desde CONSUMAN: **Consultar → Planes de Mantenimiento por Activo → Exportar**
2. Guardá el archivo como CSV
3. Abrí la app y subí el CSV
4. Revisá el tablero con semáforos y filtros
5. Descargá el Excel con formato completo

## Columnas calculadas

- **DIFERENCIA FRECUENCIA**: Frecuencia - Frecuencia acumulada
- **PEDIR MATERIAL**: Si la diferencia es ≤ 180 hs → PEDIR MATERIAL
- **DIAS PARA TAREA**: Días desde hoy hasta la fecha planificada
- **PRÓXIMO**: REVISAR (vencido) / PRÓXIMO (< 45 días) / OK (≥ 45 días)
