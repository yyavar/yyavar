# Arquitectura propuesta para un clon ThinkCell-like en Excel

## 1. Capacidades clave

- Ingesta de rangos y tablas de Excel.
- Motor de transformación semántica (series, categorías, unidades).
- Motor de estilo corporativo (paletas, fuentes, líneas guía, etiquetas).
- Persistencia de plantillas y presets.
- Exportación a PowerPoint (fase posterior).

## 2. Diseño por capas

1. **UI Layer (Task Pane)**
   - Configuración del gráfico, selección de plantilla, validaciones.
2. **Domain Layer**
   - Modelado de datos (`DataSet`, `Series`, `AxisRules`, `LabelRules`).
3. **Formatting Engine**
   - Reglas deterministas para color, leyendas, decimales, orden.
4. **Excel Adapter**
   - Lectura/escritura con Office.js.
5. **Template Repository**
   - JSON con presets por industria/cliente.

## 3. Roadmap

### Fase 1 (2-4 semanas)
- Sideload estable + lectura de datos.
- Normalización y tabla de salida estandarizada.
- 3 presets visuales básicos.

### Fase 2 (4-8 semanas)
- Constructor de reglas de etiquetas.
- Actualización incremental (reaplicar formato tras cambio de datos).
- Auditoría de errores de calidad de datos.

### Fase 3 (8-12 semanas)
- Integración con PowerPoint.
- Librería de objetos visuales corporativos.
- Métricas de uso y licenciamiento.

## 4. Consideraciones técnicas

- Office.js impone límites en operaciones masivas; agrupar lecturas/escrituras.
- Trabajar con nombres de tablas/rangos para trazabilidad.
- Agregar telemetry opt-in para entender qué reglas son más usadas.
