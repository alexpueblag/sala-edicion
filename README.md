# Sala de Edición · Yo Desarrollo

El portal diario de producción de contenido. Cada mañana a las 7:00 llega un correo con la
liga; el editor palomea, tacha y escribe al margen; lo aprobado se vuelve video y todo queda
registrado día por día.

**En línea:** https://yodesarrollo.github.io/sala-edicion/

## Cómo viaja el dato

```
Mac (produce)                    Sheet "Sala de Edición · YOD"           Editor
sala_publicar.py ── láminas ──▶  repo (GitHub Pages)                    abre el portal
                 ── manifiesto ─▶ GAS · pestaña PROPUESTAS   ◀── lee ──  palomea / tacha / nota
sala_publicar.py ◀── cosecha ──  GAS · pestaña DECISIONES    ◀── POST ── envía su revisión
                 ── avances ───▶ GAS · PRODUCCION + BITACORA             ve la retro de ayer
                                 GAS · correo 7:00 am ────────────────▶  bandeja de entrada
```

- Sin datos sensibles en el repo (es público): las decisiones viven en el Sheet.
- La clave de la sala está en la pestaña CONFIG del Sheet, nunca aquí.
- Si el Sheet no contesta, el portal cae a `datos/manifiesto.json` y las revisiones
  quedan firmadas en el navegador hasta reconectar. No se pierde nada.

## Piezas

| Qué | Dónde |
|---|---|
| Portal (una sola página) | `index.html` |
| Manifiesto de respaldo | `datos/manifiesto.json` |
| Backend (se pega a mano en Apps Script) | `gas/Code.gs` — las instrucciones van adentro |
| Publicador del lado de la Mac | `~/yod_audit/sala_publicar.py` |
