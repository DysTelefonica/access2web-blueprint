# Design system

Source: `docs/uat/assets/style-base.css` (baseline custom properties for planning and status reports; the file itself says they are not Telefónica brand tokens).

## Colors

| Token | Value | Usage | Source |
|---|---|---|---|
| `--color-brand` | `#0066CC` | Primary brand, active bars, hero gradient start | `docs/uat/assets/style-base.css` |
| `--color-brand-dark` | `#003D7A` | Headings in cards, hero gradient end, table headers | `docs/uat/assets/style-base.css` |
| `--color-text-primary` | `#1A1A1A` | Body text | `docs/uat/assets/style-base.css` |
| `--color-text-secondary` | `#5C5C5C` | Secondary text, metadata | `docs/uat/assets/style-base.css` |
| `--color-background` | `#FFFFFF` | Card background | `docs/uat/assets/style-base.css` |
| `--color-surface` | `#F5F7FA` | Page background, table header | `docs/uat/assets/style-base.css` |
| `--color-surface-strong` | `#E8EEF5` | Code chips, grey pill | `docs/uat/assets/style-base.css` |
| `--color-border` | `#D4DBE3` | Card and table borders | `docs/uat/assets/style-base.css` |
| `--color-success` | `#2E7D32` | Done, green pill text | `docs/uat/assets/style-base.css` |
| `--color-warning` | `#ED6C02` | Warning, amber pill text | `docs/uat/assets/style-base.css` |
| `--color-danger` | `#C62828` | Blocked, red pill text | `docs/uat/assets/style-base.css` |
| `--color-info` | `#0277BD` | Info callout border | `docs/uat/assets/style-base.css` |
| `--color-violet` | `#6A1B9A` | Violet pill and callout | `docs/uat/assets/style-base.css` |
| `--color-grey` | `#757575` | Planned bars | `docs/uat/assets/style-base.css` |

Pill backgrounds: green `#E8F5E9`, amber `#FFF3E0`, red `#FFEBEE`, violet `#F3E5F5`. Callout backgrounds: info `#E3F2FD`, warn `#FFF8E1`, danger `#FFEBEE`, violet `#F3E5F5`.

## Typography

- Family: `-apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif`; line-height 1.55.
- Monospace: `ui-monospace, SFMono-Regular, Menlo, Consolas, monospace`, 12px.
- Hero title: `clamp(22px, 3vw, 30px)`, 700. Eyebrow: 12px, 600, uppercase, letter-spacing .08em.
- Card h2: 18px, 700, brand-dark. Card h3: 15px, 700. Lead: 16px. Table: 13px; table header 12px uppercase.

## Spacing

`--space-1` 4px, `--space-2` 8px, `--space-3` 12px, `--space-4` 16px, `--space-5` 24px, `--space-6` 32px, `--space-7` 48px.

## Radius and shadow

- `--radius-card` 12px, `--radius-chip` 24px, `--radius-input` 8px.
- `--shadow-card`: `0 1px 3px rgba(0,0,0,.06), 0 4px 12px rgba(0,0,0,.04)`.

## Components

- **Page:** max-width 1180px, centered, padding 48px 24px.
- **Hero:** 135° gradient brand → brand-dark, white text, card radius and shadow.
- **Card (`section.card`):** white, 1px border, card radius, 32px padding, card shadow.
- **RAG pill:** 3px 10px, chip radius, 12px bold; green, amber, red, grey, violet variants.
- **Data table:** full width, 13px, row borders, surface header in brand-dark uppercase.
- **Callout:** 4px left border in info color on `#E3F2FD`; `warn`, `danger`, `violet` variants.
- **Stat grid:** auto-fit columns of at least 160px, 12px gap.
