# Excel Add-in estilo ThinkCell (MVP)

Este repositorio contiene una base para construir un **add-in de Excel** inspirado en ThinkCell.

## Objetivo del MVP

Implementar un panel lateral (task pane) en Excel que permita:

1. Leer rangos de datos seleccionados.
2. Generar gráficas con formato corporativo.
3. Aplicar reglas automáticas de etiquetado y colores.
4. Insertar y actualizar objetos visuales de forma consistente.

> Nota: ThinkCell incluye funcionalidades avanzadas de PowerPoint y layout inteligente. Este MVP se enfoca en el lado de Excel como primer paso.

## Estructura

- `manifest.xml`: manifiesto de Office Add-in.
- `src/taskpane/taskpane.html`: UI del panel.
- `src/taskpane/taskpane.js`: lógica Office.js para leer datos y crear tabla de salida.
- `docs/arquitectura.md`: propuesta técnica por fases para acercarse a un clon funcional.

## Cómo probar rápidamente

1. Sirve el contenido con HTTPS (recomendado para Office Add-ins).
2. Sideload del `manifest.xml` en Excel.
3. Abre el panel y selecciona un rango de celdas.
4. Haz clic en **Aplicar estilo ThinkCell-like** para generar una tabla de salida normalizada.

## Siguientes pasos sugeridos

- Migrar a TypeScript + bundler (Vite/Webpack).
- Añadir un motor de reglas de formato declarativo.
- Integrar una librería de gráficas (o usar Office Charts API cuando aplique).
- Añadir telemetría, licenciamiento y actualizador.
