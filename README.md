# SLC OPS

PWA modular para control de operaciones.

## Módulos
1. **Hourly Reports** — funcional en esta primera versión.
2. **Employees** — reservado.
3. **Estimado de Operaciones** — reservado.

## Funciones actuales
- Iniciar / continuar / finalizar operaciones.
- Importar Excel Hourly.
- Detectar barco y viaje por separado.
- Asociar cada Hourly a un auditor.
- Mantener cada importación como snapshot histórico.
- Editar variables sin borrar el valor original.
- KPI de movimientos, productividad y ETC.
- Tabla de productividad por hora.
- Gangs, movimientos por tipo, vacíos y remarks.
- IndexedDB local.
- Exportar/importar backup JSON.
- PWA instalable.

## Auditores iniciales
1. Leonardo Espino (Leo)
2. Darlin Tapia
3. Ridy Gonzalez

Supervisor de Operadores: Pedro Amador.

## GitHub Pages
Sube el contenido de este proyecto a un repositorio y activa **Settings → Pages → Deploy from a branch**.

> La lectura de Excel usa SheetJS desde CDN. La primera carga con conexión permite que el service worker lo almacene en caché del navegador.
