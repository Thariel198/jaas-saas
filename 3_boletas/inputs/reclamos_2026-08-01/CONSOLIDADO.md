# Consolidado de correcciones — 01/08/2026

Resumen de todo lo parchado hoy en `3_boletas/inputs/DATA_boletas.xlsx`.
Detalle completo (motivo, verificación contra el ledger, pendiente exacto)
de cada fila en `README.md` de esta misma carpeta. Ninguna de estas
correcciones tocó el ledger real (`shared/seguimiento_pueblo.xlsx`) — son
cosméticas para la boleta de hoy, pendientes de reconciliar el mes que
viene.

## Bloque A — 13 reclamos dictados de las fotos

| # | Predio | Nombre | Recibo | Corregido | Total antes → después |
|---|---|---|---|---|---|
| 1 | G-14 | Margarita Gomez Bonifacio | 18114 | Convenio 38→0, Multa 0→50, Techado y campo 21→50 | 83 → 124 |
| 2 | E-14B | Juan Saavedra Saavedra | 18084 | Techado y campo 0→75 | 11 → 86 |
| 3 | J-1 | Comedor Popular Club de Madres | 18170 | Multa 20→0, Techado y campo 75→0 | 216 → 121 |
| 4 | K-9 | Fortunato Vargas Cabello | 18192 | Convenio 75→0 | 87 → 12 |
| 5 | K-8 | Victor Teodoro Flores Durand | 18191 | Techado y campo 30→0 | 45 → 15 |
| 6 | T-14 | Pedro Candacho Huarac | 18342 | Convenio 50→0 | 187 → 137 |
| 7 | V-14 | Leonardo Huamani Sotelo | 18368 | Corte y reconexión 40→0 | 84 → 44 |
| 8 | B-8 | Rosalina Olimpia Ciriaco Sotelo | 17992 | Convenio 75→0, Techado y campo 5→0, Mes anterior 46→0 (atendido en persona por el usuario) | 150 → 24 |
| 9 | F-10 | Herminio Lucero Trujillo | 18096 | Techado y campo 50→0 | 80 → 30 |
| 10 | F-1 | Maria Godo Sifuentes | 18086 | Multa 20→0 | 37 → 17 |
| 11 | F-7 | Victor Laurencio Valladares | 18093 | Techado y campo 25→0 | 126 → 101 |
| 12 | D-6 | Hermelinda Jara Trujillo | 18056 | Techado y campo 50→0 | 239 → 189 |
| 13 | O-2 | Carmen Ingaruca Julca | 18245 | Multa 30→0 | 38 → 8 |
| 14 | G-4 | Natalia Chinchay Collas | 18104 | Convenio 75→0, Multa 25→50, Techado y campo 50→50 | 158 → 108 |

## Bloque B — 11 predios del lote de SALDO negativo (investigación, no reclamo)

Distinto origen: no lo pidió la directiva, salió de investigar por qué
salían boletas con Convenio/Techado y campo/Multa en negativo (pago
fantasma de julio sin respaldo real — ver
`4b_reclamos/outputs/reporte_lote_saldo_negativo_2026-07.pdf`). Acá se
restauró el SALDO a la deuda real de junio (verificada evento por evento),
no se exoneró nada.

| # | Predio | Nombre | Recibo | Corregido | Total antes → después |
|---|---|---|---|---|---|
| 1 | A-8 | Victor Melgarejo Corcino | 17979 | Convenio -50→50 | 0 → 90 |
| 2 | **B-5** | Pompeyo Celestino Lliuya | 17989 | Convenio -50→50, Techado y campo -25→25 | 12 → 162 |
| 3 | C-1 | Odilon Cerna Romero | 18007 | Convenio -50→50, Techado y campo -25→25 | 0 → 94 |
| 4 | C-7 | Victor Lopez Trujillo | 18014 | Convenio -25→25 | 16 → 66 |
| 5 | E-12 | Teofila Fernandez Reyes | 18080 | Convenio -16→26 | 5 → 47 |
| 6 | I-11 | Dominga Chacara Lopez | 18162 | Convenio -25→25 | 1 → 51 |
| 7 | I-16 | Adolfo Rosario Rojas | 18167 | Multa -18→18, Techado y campo 47→75 | 137 → 201 |
| 8 | J-3 | Vilma Celestino Villafana | 18172 | Convenio -30→50 | 0 → 66 |
| 9 | K-17 | Marcial Sanchez Araoz | 18200 | Convenio -25→25 | 0 → 40 |
| 10 | K-2 | Antonio Espinoza Sifuentes | 18185 | Convenio -25→25 | 0 → 39 |
| 11 | P-12 | Judith Venturo Rosales | 18285 | Convenio -50→50 | 0 → 86 |

**B-5 está en los dos bloques** — apareció en la foto de reclamo (Pompeyo
decía que debía techado y campo, no convenio) pero se resolvió con el
método del Bloque B (restaurar tal cual, versión A), no con lo que él
declaró — queda pendiente de confirmar con él (ver `README.md`).

## Bloque C — H-16, mismo patrón que el Bloque B pero encontrado después

Se escapó del filtro original del lote de 11 porque su SALDO no quedó
negativo (47, no -X) — mismo mecanismo (pago fantasma de julio + AJUSTE que
revierte mal), detectado al revisar el pedido puntual de la directiva.

| Predio | Nombre | Recibo | Corregido | Total antes → después |
|---|---|---|---|---|
| H-16 | Gregorio Tolentino Sanchez | 18144 | Techado y campo 47→75 | 66 → 94 |

## Totales

- **25 boletas** corregidas hoy (13 + 11 + 1, sin contar el duplicado de B-5).
- **0 tocan el ledger real** — todo pendiente de `registrar_ajuste` /
  `registrar_pago` en `seguimiento_pueblo.xlsx` el mes que viene, detalle
  predio por predio en `README.md`.
- Backups de cada paso en `3_boletas/inputs/backups/DATA_boletas_pre_*.xlsx`.
