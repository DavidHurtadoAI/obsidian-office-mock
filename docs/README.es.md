# Office Mock 1.0.1

Tema independiente para Obsidian con documentos azules, correo azul, reuniones moradas y hojas de cálculo verdes. Tiene interfaz propia en claro y oscuro y respeta las fuentes y el tamaño de texto configurados en Obsidian.

## Instalación

Descarga la última release del repositorio, descomprime el ZIP y copia su carpeta `Office Mock` en `.obsidian/themes/` de tu bóveda. Dentro deben estar directamente `manifest.json` y `theme.css`. Después selecciona **Office Mock** en Ajustes → Apariencia → Temas. Si no aparece todavía, reinicia Obsidian.

Desactiva el snippet `office-clean` si utilizabas la versión anterior. El tema ya incluye sus vistas y funciona por sí solo. Los ejemplos existentes de `Office Mock CSS` sirven tal cual.

## Notas y edición

Las notas normales se presentan como un documento. `cssclasses` permite elegir `office-document`, `office-email`, `office-meeting` u `office-spreadsheet`. Usa `office-plain` para conservar la disposición nativa de una nota. Otras clases y una propiedad vacía ya no desactivan automáticamente el documento: el tema evita depender del DOM de Propiedades.

En Live Preview puedes escribir en el texto y en las celdas de las tablas. `Tab` confirma y avanza; `Mayús+Tab` vuelve a la celda anterior. Los cambios siguen guardándose como Markdown. El modo Fuente conserva el editor nativo.

El asunto o título va en un encabezado Markdown. Los datos de correo y convocatoria van en una cita Markdown inmediatamente posterior. La UI está definida en CSS; los archivos no contienen barras, HTML ni filas de relleno.

Las cintas, búsqueda, pestañas y barra de fórmulas son decorativas. La edición y las tablas son las de Obsidian. No hay cálculos, envío de correo ni creación de reuniones. Los datos y adjuntos escritos en los ejemplos son ficticios.

## Desarrollo en Obsidian-dev

Fuente: `Development/office-mock-theme`. Desde la raíz del proyecto:

```sh
node scripts/obsidian-dev.mjs install office-mock
node scripts/obsidian-dev.mjs verify office-mock
```

La instalación conserva ejemplos y plantillas existentes y guarda copias de los archivos reemplazados fuera de la bóveda, en `.dev-backups/`.

La publicación en GitHub y la aceptación en Community Themes son pasos distintos. Consulta el estado actualizado en `docs/COMMUNITY.md` del repositorio.
