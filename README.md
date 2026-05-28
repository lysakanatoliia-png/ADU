# ADU Materials Estimate · Menlo Park

**Project:** Garage conversion to ADU
**Location:** 311 Lexington Drive · Menlo Park · CA 94025
**Footprint:** 315 sq ft (15' × 21')

## Grand Total

**$43,833** materials only (per spec §32 — no labor, tax, permits, delivery, contingency).

| Reference | $ |
|---|---:|
| + CA Sales Tax 9.375% | $4,109 |
| **Materials + Tax** | **$47,942** |
| + 15% material buffer | $6,575 |
| **Planning budget** | **$54,517** |
| + Labor + permits + markup + contingency | ~$45,000–70,000 |
| **All-in project estimate** | **$90,000–115,000** |

## Live dashboard

Open [`index.html`](./index.html) in any modern browser — single-file dashboard with:

- Interactive floor plan (7 functional zones, click → detail breakdown)
- 17 construction phases with color-coded stages (Shell → MEP → Envelope → Finish)
- Mobile-responsive (iPhone 14/15/17 optimized)

To publish as live URL: enable **GitHub Pages** in repo settings → Source: `main` branch root.

## Files

| File | Purpose |
|---|---|
| `index.html` | Interactive dashboard (single file, no build) |
| `adu_plan_en.svg` | English floor plan (used as background) |
| `ADU_Materials_Estimate.xlsx` | Full estimate · 20 sheets · 474 items |
| `build_estimate.py` | Python generator (rebuild xlsx from scratch) |
| `update_prices.py` | Python prices applier |
| `Garage Conversion to ADU.md` | Original project spec (34 sections) |
| `CLAUDE.md` | Project handover guide |

## Construction phases

| # | Phase | $ |
|---|---|---:|
| 1 | Site & Exterior Utilities | $2,566 |
| 2 | Raised Floor System | $2,923 |
| 3 | Framing | $1,546 |
| 4 | Doors & Windows | $7,925 |
| 5 | Plumbing Rough-in | $1,490 |
| 6 | Electrical Rough-in | $3,522 |
| 7 | HVAC Mini-Split | $3,660 |
| 8 | Insulation | $1,898 |
| 9 | Drywall | $1,923 |
| 10 | Paint | $998 |
| 11 | Flooring & Trim | $1,458 |
| 12 | Kitchen Cabinets | $2,890 |
| 13 | Bathroom Fixtures + Tile | $2,021 |
| 14 | Closet System | $637 |
| 15 | Water Heater | $675 |
| 16 | Appliances | $4,715 |
| 17 | Furniture (IKEA) | $2,987 |

## Functional zones (interactive on plan)

| Zone | $ |
|---|---:|
| Kitchen | $5,145 |
| Bathroom | $2,895 |
| Laundry | $2,545 |
| Living Room | $1,850 |
| Bedroom | $1,450 |
| Entry | $700 |
| Walk-in Closet | $637 |
| Dining | $360 |

## Special Orders (lead 4–12 weeks)

| Item | $ |
|---|---:|
| French Door | $1,500 |
| Folding Partition System 8' (interior) | $3,000 |
| MRCOOL DIY Mini-split kit | $3,150 |
| Bosch compact washer 24" | $1,000 |
| Bosch ventless dryer 24" | $1,400 |
| French Door Hardware | $300 |
| **Total Special Orders** | **$10,350** |

## Rebuild estimate from scratch

```bash
python3 build_estimate.py    # creates fresh xlsx
python3 update_prices.py     # applies all prices
```

Edit `build_estimate.py` to add/remove items per phase.
Edit `update_prices.py` to change prices.

## Out of scope (per spec §32)

Not included in Grand Total:
- Labor / installation pricing
- Permits + plan check + inspections
- Architectural / engineering / Title 24
- Demolition + trenching labor
- Contractor markup + contingency
- Delivery / shipping / disposal fees

---

*Materials-only estimate · pre-permit budget planning · 474 items across 17 construction phases.*
