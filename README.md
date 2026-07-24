# Observatorio de servicios públicos

Página estática interactiva con indicadores oficiales de España y sus comunidades autónomas. La interfaz está en `index.html`; los datos normalizados, las fuentes y el estado de cada actualización se guardan en `data/dashboard.json`.

## Cobertura actual

La versión incluida contiene **176 series conectadas** dentro de un catálogo de **197 indicadores**. Las series pueden ser:

- **Oficiales directas**: reproducen una tabla, fichero o informe de un organismo público.
- **Calculadas a partir de datos oficiales**: totales, cuotas o tasas construidas únicamente con valores oficiales y descritas como tales en la ficha metodológica.

### Demografía, economía, empleo, renta y vivienda

- Población residente.
- PIB regional y PIB por habitante.
- Renta neta media por persona.
- Personas paradas y tasa de paro según la EPA.
- Afiliación media a la Seguridad Social.
- Paro registrado.
- Tasa de riesgo de pobreza.
- Viviendas de uso turístico y peso sobre el parque residencial.
- Valor tasado medio de la vivienda en euros por m².
- Variación trimestral e interanual del valor tasado.

### Enseñanza no universitaria

Para nueve etapas se ofrecen alumnado público, privado, total y cuotas pública/privada:

- Enseñanzas de régimen general.
- Educación Infantil.
- Primer ciclo de Infantil, 0–3 años.
- Educación Primaria.
- Educación Secundaria Obligatoria.
- Bachillerato.
- FP de Grado Básico.
- FP de Grado Medio.
- FP de Grado Superior.

Las tablas históricas de EDUCAbase agrupan dentro de **centros privados** los concertados y los no concertados. El observatorio no inventa una separación que la fuente histórica no ofrece.

### Universidad

- Estudiantes universitarios totales, de grado, máster y doctorado.
- Tasa neta de escolarización universitaria.
- Evolución nacional de estudiantes de grado y máster.
- Universidades totales, públicas y privadas por comunidad autónoma.
- Porcentaje de universidades públicas y privadas.
- Alumnado en universidades públicas y privadas por nivel.
- Cuotas pública y privada del alumnado universitario.

El último curso universitario incorporado es **2024-2025** y las cifras de alumnado de ese curso tienen carácter provisional en la publicación del SIIU. Ceuta y Melilla aparecen integradas en Andalucía en la tabla regional de alumnado universitario.

### Atención Primaria

- Áreas de salud y zonas básicas.
- Centros de Atención Primaria, centros de salud y consultorios locales.
- Centros con acreditación docente.
- Gestión pública directa, otras modalidades públicas y gestión privada.
- Desglose de esas modalidades para centros de salud y consultorios.

El Catálogo de Centros de Atención Primaria 2026 refleja la situación a **31 de diciembre de 2025**. En esta fuente Ceuta y Melilla se ofrecen conjuntamente.

### Atención hospitalaria

- Pacientes en lista de espera quirúrgica, tasa, tiempo medio y esperas superiores a seis meses.
- Listas quirúrgicas de cirugía general, oftalmología y traumatología.
- Pacientes pendientes de primera consulta, tasa, tiempo medio y espera superior a 60 días.
- Primeras consultas de dermatología, oftalmología y traumatología.
- Hospitales totales, públicos y privados.
- Camas hospitalarias totales, públicas y privadas.
- Hospitales generales, especializados, de media y larga estancia, de salud mental y otros centros con internamiento.
- Equipos TAC, resonancia magnética, PET, mamografía y puestos de hemodiálisis.
- Altas de hospitales en el catálogo, con desglose público/privado, y bajas totales.

El Catálogo Nacional de Hospitales 2025 refleja la situación a **31 de diciembre de 2024**. La titularidad se agrupa usando los códigos de dependencia funcional del propio catálogo.

### Dependencia

- Solicitudes y resoluciones de grado.
- Personas con derecho a prestación.
- Resoluciones de PIA.
- Beneficiarios con prestación efectiva.
- Personas con PIA sin prestación efectiva.
- Esperas superiores a seis meses.
- Prestaciones reconocidas.
- Tiempos desde solicitud a grado, de grado a prestación y tiempo total.

El último corte integrado es **30 de junio de 2026**.

### Pensiones e Ingreso Mínimo Vital

- Pensionistas y pensiones contributivas.
- Pensiones de jubilación.
- Pensiones por pensionista.
- Pensión media del sistema y de jubilación.
- Nómina mensual total y de jubilación.
- Fondo de Reserva de la Seguridad Social, 2000–marzo de 2026.
- Hogares y personas beneficiarias del IMV.
- Nómina mensual y cuantías medias por hogar y beneficiario.

El último corte mensual de pensiones e IMV incorporado es **junio de 2026**.

### Residencias para personas mayores

- Centros residenciales totales, públicos y privados.
- Plazas totales y plazas según financiación pública o privada.
- Cobertura de plazas y personas usuarias.
- Precio público anual y precio de concertación.
- Aportación de la persona usuaria en plaza pública y concertada.

El último año disponible integrado es **2024**.

## Indicadores que siguen sin una serie nacional homogénea

El catálogo conserva también indicadores todavía no conectados. Entre los principales límites estadísticos están:

- Derivaciones individualizadas de pacientes de la sanidad pública a la privada.
- Coste comparable por paciente en centros públicos y privados.
- Dinero público recibido por cada proveedor sanitario privado.
- Listas de espera administrativas homogéneas de Atención Primaria.
- Personal sanitario público y privado con la misma definición territorial y temporal.
- Número real de millonarios y población que concentra exactamente el 80 % de la riqueza.
- Viviendas vacías con actualización anual; el dato completo sigue siendo fundamentalmente censal.
- Tramos exactos de renta para toda la población, no solo declarantes del IRPF o estimaciones muestrales.
- Separación histórica uniforme entre enseñanza privada concertada y privada no concertada en todas las etapas.

Estas fichas permanecen visibles para que el observatorio muestre también los huecos de información pública y no solo aquello que ya puede medirse.

## Abrir la página en el Mac

No hace falta instalar dependencias para visualizar los datos incluidos:

```bash
cd ~/Downloads/observatorio_servicios_publicos_html
python3 -m http.server 8000
```

Abrir después:

```text
http://localhost:8000
```

Para detener el servidor, pulsar `Control + C` en Terminal.

## Actualizar manualmente los datos

El recolector utiliza Python, `openpyxl`, `pypdf` y LibreOffice para convertir algunos ficheros XLS heredados del SEPE.

```bash
cd ~/Downloads/observatorio_servicios_publicos_html
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python scripts/update_data.py
```

Cada fuente se ejecuta de forma independiente. Si una falla o cambia de formato, se conserva la última versión válida de sus series y el error se registra en `meta.source_checks`; las demás fuentes continúan actualizándose.

Algunas publicaciones anuales integradas mediante tablas oficiales de síntesis —por ejemplo, el informe universitario o el Fondo de Reserva— requieren revisar el recolector cuando el organismo publique una nueva edición. El control diario permite detectar cambios, pero no convierte una publicación anual en una serie diaria.

## Publicación y actualización automática en GitHub Pages

El archivo `.github/workflows/daily-update.yml` realiza dos tareas:

1. Publica la web en GitHub Pages cada vez que se envían cambios a la rama `main`.
2. A las 09:00, hora de Madrid, consulta las fuentes, actualiza `data/dashboard.json`, guarda los cambios y despliega inmediatamente la nueva versión.

Para activarlo:

1. Subir toda la carpeta a un repositorio público de GitHub, incluida la carpeta oculta `.github`.
2. En **Settings → Pages**, seleccionar **GitHub Actions** como fuente de publicación.
3. En **Settings → Actions → General → Workflow permissions**, activar **Read and write permissions**.
4. Abrir **Actions → Actualizar y publicar observatorio** y ejecutar una vez **Run workflow**.

El flujo programa dos comprobaciones UTC para cubrir tanto CET como CEST, pero solo actualiza cuando en `Europe/Madrid` son las 09:00. Una ejecución manual actualiza y publica de inmediato.

## Estructura

```text
observatorio_servicios_publicos_html/
├── index.html
├── requirements.txt
├── README.md
├── data/
│   └── dashboard.json
├── scripts/
│   ├── update_data.py
│   └── extended_collectors.py
└── .github/
    └── workflows/
        └── daily-update.yml
```

## Precisión metodológica

La aplicación distingue entre:

- periodo de referencia del dato;
- fecha de publicación;
- fecha de consulta;
- dato oficial directo o cálculo derivado;
- carácter definitivo, provisional, muestral o experimental.

Una comprobación diaria no implica que todas las series cambien diariamente: las frecuencias reales son mensuales, trimestrales, semestrales, anuales o por curso académico.
