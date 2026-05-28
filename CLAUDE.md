# CLAUDE.md — ADU Materials Estimate Project

## Проєкт

**Garage Conversion to ADU — Menlo Park, CA 94025**
- ADU footprint: 15' × 21' = 315 sq ft
- Vaulted ceiling: 8 ft edges / 12 ft center
- Raised floor: 10 inches
- Electric-only (no gas)
- Type: Materials-only estimate per ТЗ §32 (no labor / tax / permits / delivery)

## Final Result

**Grand Total: $43,668.11** materials only

| Reference | $ |
|---|---:|
| + CA Sales Tax 9.375% | $4,094 |
| + 15% Material Buffer | $6,550 |
| **Planning Budget (materials + tax + buffer)** | **$54,312** |
| + Labor + Permits + Markup | ~$45,000-70,000 |
| **All-in ADU project** | **$90,000-115,000** |

Cost per sq ft: $139/sf materials only, $285-365/sf all-in (typical Bay Area ADU).

## Файли в папці

| File | Purpose |
|---|---|
| `ADU_Materials_Estimate.xlsx` | **Main deliverable** — 20 листів, 465 items |
| `build_estimate.py` | Python generator (creates file from scratch) |
| `update_prices.py` | Apply prices to existing file |
| `Garage Conversion to ADU.md` | Original ТЗ (34 sections) |
| `ChatGPT Image 23 трав. 2026 р., 12_43_00.png` | PNG concept layout |
| `ADU_Materials_Estimate_v1_zones_backup.xlsx` | Backup zone-based V1 |
| `CLAUDE.md` | This file |

## Workflow для майбутніх сесій

### Стандартний rebuild

```bash
cd "/Users/anatoliilisak/claude-workspace/documents/ADU — Menlo Park, CA 94025"
python3 build_estimate.py    # creates fresh xlsx з нуля
python3 update_prices.py     # applies all prices
```

**Чому два кроки:** build_estimate.py створює структуру з нуля. update_prices.py вставляє ціни в конкретні клітинки. Якщо запустити update двічі, Notes можуть подвоїтись — тому workflow завжди build → update.

### Як змінювати

| Що треба змінити | Файл |
|---|---|
| Додати/видалити item у фазі | `build_estimate.py` — list у відповідній phase |
| Змінити ціну existing item | `update_prices.py` — corresponding tuple |
| Змінити quantity Low/Mid/High | `build_estimate.py` (item tuple) або `QTY_OVERRIDES` |
| Додати нову phase | `build_estimate.py` — `SHEETS` list + new data list + write logic |

### Перевірка результату

```bash
python3 -c "
from openpyxl import load_workbook
wb = load_workbook('ADU_Materials_Estimate.xlsx', data_only=False)
ws = wb['01_Summary']
# Grand Total знаходиться на row 37, column F
print('Grand Total:', ws.cell(row=37, column=6).value)
"
```

## Структура Excel (20 листів)

**01_Summary** — зведений з Reference Totals (tax + buffers), Top Cost Drivers, High-Risk Verification, Included/Out-of-scope blocks.

**Phases в порядку construction (shell-to-finish):**

| # | Phase | $ | Notes |
|---|---|---:|---|
| 02 | Site & Exterior Utilities | $2,566 | Sewer 100 ft + Water 100 ft + Electric 50 ft |
| 03 | Raised Floor System | $2,923 | 2x10 joists, 10" rise |
| 04 | Framing | $1,546 | Left wall +21 ft + interior partitions |
| 05 | Doors & Windows | $7,925 | **Top cost driver** — 3 windows + 5 doors + folding + French |
| 06 | Plumbing Rough-in | $1,490 | PEX + drain + vent + shut-offs |
| 07 | Electrical Rough-in | $3,522 | 100A subpanel + AFCI breakers (CA Code) |
| 08 | HVAC Mini-Split | $3,660 | MRCOOL DIY 5th Gen 18k 9+9 kit |
| 09 | Insulation | $1,898 | R-15 walls + R-30 vaulted ceiling |
| 10 | Drywall | $1,923 | 67 sheets + waterproof shower zone |
| 11 | Paint | $998 | 12 gal wall (2 coats) |
| 12 | Flooring & Trim | $1,458 | LifeProof LVP 22 MIL 350 sf |
| 13 | Kitchen Cabinets | $2,890 | IKEA SEKTION one-wall + countertop |
| 14 | Bathroom Fixtures | $1,856 | Prefab shower 32×40 + 24" vanity |
| 15 | Closet System | $637 | IKEA BOAXEL 49" combo |
| 16 | Water Heater | $675 | Rheem 30 gal short electric |
| 17 | Appliances | $4,715 | 5 kitchen + Bosch washer/dryer (Special Order) |
| 18 | Furniture IKEA | $2,987 | Living (FRIHETEN sleeper) + Bedroom (MALM) + Dining |

**Cross-reference sheets (qty 0, not added to Grand Total):**
- 19_Special_Orders — список 6 Special Order items з lead times
- 20_Verification_Items — 26 пунктів §33 must-verify

## Колонки в робочих листах (17 колонок)

| # | Col | Назва |
|---|---|---|
| A | 1 | # |
| B | 2 | Phase |
| C | 3 | Категорія |
| D | 4 | Item / Material |
| E | 5 | Spec / Опис |
| F | 6 | Unit |
| G | 7 | Qty (Low) |
| H | 8 | Qty (Mid) |
| I | 9 | Qty (High) |
| J | 10 | Unit Price ($) |
| K | 11 | Total ($) — формула =H × J |
| L | 12 | Source |
| M | 13 | SKU / Model |
| N | 14 | URL |
| O | 15 | Status |
| P | 16 | Confidence |
| Q | 17 | Notes |

## Critical cross-references (avoid duplicates!)

Деякі items могли бути дублюваними між фазами — вони cross-referenced (qty 0 у одній фазі):

| Item | Real location | Cross-ref в |
|---|---|---|
| Floor insulation | **Phase 2 r12** | Phase 8 r7 |
| Shower drain assembly | **Phase 5 r33** | Phase 13 r7 |
| Shower rough-in valve body | **Phase 5 r32** | Phase 13 r8 |
| Toilet flange | **Phase 5 r19** | (Phase 13) |
| Bathroom P-trap | **Phase 5 r14** (4 total) | Phase 13 r17 |
| Bathroom shut-offs | **Phase 5 r8** (19 total) | Phase 13 r18 |
| Cement board + RedGard | **Phase 9 r7 + r16** | Phase 13 |
| PVA drywall primer | **Phase 9 r12** | Phase 10 (finish primer separate) |
| Closet sliding door | **Phase 4 r8** | Phase 14 r19 |
| Closet motion light | **Phase 6 r25** | Phase 14 r20 |
| Utility closet door | **Phase 4 r11** | Phase 15 r13 |
| WH connection kit | **Phase 5 r20** | Phase 15 r8 |
| WH shut-offs | **Phase 5 r8** | Phase 15 r9 |
| WH electrical | **Phase 6 r9** (240V cable) | Phase 15 r21 |
| Mini-split indoor units (2) | **Phase 7 r5 kit** | Phase 7 r6 (qty 0) |
| Mini-split line sets (2) | **Phase 7 r5 kit** | Phase 7 r7 (qty 0) |
| Mini-split electrical disconnect | **Phase 6 r19** | Phase 7 r28 |
| Dishwasher air gap | **Phase 5 r37** | Phase 12 + 16 cross-ref |
| Dishwasher anti-tip mounting | **Phase 12 r33** | (Phase 16) |
| Fridge water line kit | **Phase 5 r27** | Phase 12 + 16 cross-ref |

## Special Orders (lead 4-12 weeks)

| Item | Phase | $ | Lead |
|---|---|---:|---|
| French Door | Phase 4 r9 | $1,500 | 4-8 weeks |
| Folding Partition System 8 ft (INTERIOR) | Phase 4 r13 | $3,000 | 4-12 weeks |
| French Door Hardware | Phase 4 r24 | $300 | 2-4 weeks |
| MRCOOL DIY Mini-split kit | Phase 7 r5 | $3,150 | 4-8 weeks |
| Bosch Compact Washer 24" | Phase 16 r10 | $1,000 | 4-6 weeks |
| Bosch Ventless Dryer 24" | Phase 16 r11 | $1,400 | 4-6 weeks |
| **TOTAL Special Orders** | | **$10,350** | |

## Process — як вели по фазах

Per кожна phase (стандартний workflow):

1. **Pre-Takeoff Check** — cross-check ТЗ §X + PNG plan + Owner Confirmed §5, виявити пропущене, технічні рішення
2. User погоджує
3. **Збір цін** через WebSearch (Home Depot, IKEA, Amazon, Best Buy, ikea.com)
4. Edit `build_estimate.py` (новi items, qty corrections) + `update_prices.py` (prices)
5. Rebuild + verify
6. **Звіт** з top-3, cumulative, scenarios
7. User робить **detailed review** (5-10 правок з своїми професійними experience)
8. Імплементую V2 (новi items + notes updates + corrections)
9. Rebuild + show V2 total

**Iterations:** Кожна фаза дала ~3 review cycles (V1 → V2 → V3). Final estimate after 17 phases priced + 17 reviews implemented.

## Sourcing strategy

| Category | Source |
|---|---|
| Construction materials | Home Depot (Charlotte, Owens Corning, USG, Southwire, Cantex, Carlon, Quikrete) |
| Cabinets + Furniture | IKEA (SEKTION kitchen, BOAXEL closet, MALM bed, FRIHETEN sleeper, LACK, BESTÅ, KALLAX, LOHALS, VIDJA, RITVA, TRETUR) |
| Appliances | Best Buy / Amazon / Home Depot (Frigidaire, GE, Bosch, MRCOOL DIY) |
| Plumbing | SharkBite PEX-A, Apollo valves, Charlotte/IPEX ABS DWV |
| Electrical | Square D QO + AFCI breakers, Leviton TR-rated outlets, Southwire Romex |
| Lighting | Halo/Lithonia LED wafer, First Alert smoke/CO |
| Bathroom | MAAX/Sterling shower, Home Decorators vanity, Glacier Bay toilet, Behr paint |
| HVAC | MRCOOL DIY 5th Gen kit (Home Depot) |
| Water Heater | Rheem Performance |

## CA Code compliance pinned items

Items required by CA Code 2019+:
- AFCI/Dual-function breakers (Phase 6) — bedroom/living/kitchen/dining
- Tamper-resistant outlets (Phase 6)
- Hardwired interconnected Smoke/CO (Phase 6)
- Seismic strap kit для water heater (Phase 15)
- Pressure regulator якщо street >80 PSI (Phase 1, verify)
- Tracer wire underground utilities (Phase 1, verify)
- 5/8" Type X drywall ceiling якщо fire-rating req'd (Phase 9, optional)

## 26 Verification Items (§33) — High Risk

Critical items to verify before final ordering:

1. Field dimensions of garage
2. Bathroom size after adding laundry + utility
3. Washer/dryer exact location
4. Utility closet exact size
5. Electric tank water heater exact size (20 vs 30 gal)
6. Washer/dryer exact model
7. Mini-split BTU sizing (Manual J load calc)
8. IKEA Kitchen Planner exact layout
9. Exact appliance dimensions
10. Window sizes (3)
11. French door size
12. Main entry door size
13. Folding system product + rough opening
14. Closet sliding door size
15. Sewer route + slope verification
16. Water supply route
17. Electrical route
18. Existing main panel capacity
19. Required ADU subpanel amperage
20. Floor slab leveling needed?
21. Moisture mitigation over slab?
22. Additional waterproofing у wet zones?
23. Structural headers for exterior openings?
24. CA / Menlo Park code requirements (fire, energy, vent, safety)

Phase 20 в Excel містить повний список — qty 0 priced, just checklist.

## Що далі (Optional next steps)

1. **Export PDF** — final delivery для estimator
2. **Verification Items deep-dive** — review кожен з 26 пунктів окремо
3. **Special Orders sheet update** з точними lead times і supplier контактами
4. **README** для естіматора з процесом використання
5. **Site visit** — польові заміри для зменшення TBD positions
6. **Licensed contractor consultation** — MEP plan, structural calc, Title 24

## Контакт інфо проєкту

Project type: Materials-only estimate (pre-permit)
Date created: 2026-05-23 (first version)
Date completed: 2026-05-27 (final V2 з усіма reviews)
Tools used: Python 3 + openpyxl

---

*Цей file — guide для майбутніх сесій з проєкту. Він описує final state, не процес.*
