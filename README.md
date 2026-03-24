# PVGIS Monthly Exporter

Repositorio para generar y versionar exportaciones mensuales de PVGIS a partir de una latitud y una longitud.

## Qué hace

Para una ubicación (`latitude`, `longitude`), consulta PVGIS y genera promedios mensuales del período 2005–2023 usando la base solar `PVGIS-SARAH3` en tres formatos:

- CSV
- JSON
- XLSX

## Variables incluidas

Para cada mes (1–12), se exportan:

- `global_horizontal_irradiation_kwh_m2_month`
- `direct_normal_irradiation_kwh_m2_month`
- `global_irradiation_optimum_angle_kwh_m2_month`
- `diffuse_to_global_ratio`
- `average_temperature_c`

Además, cada fila incluye:

- `month_number`
- `month_name`
- `latitude`
- `longitude`
- `solar_radiation_database`
- `start_year`
- `end_year`

## Estructura de salida

Los archivos generados se guardan en:

```text
data/<location-slug>/
