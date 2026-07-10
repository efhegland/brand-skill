---
description: "M1 brand font (Inter Variable). Use when the user mentions: font, brand font, M1 font, typography, Inter, typeface, font install, font check, font usage, or any typography guidance for M1 materials."
disable-model-invocation: false
---

# M1 brand font: Inter Variable

Inter is the M1 brand font. M1 has standardized company-wide on the **Inter
variable font** — file `InterVariable.ttf`, family name **`Inter Variable`**. It
is required for all external communications and PowerPoint presentations.

## The variable-font named instances (what PowerPoint uses)

PowerPoint does not consume the variable axis directly — it asks for a font by
its **family / Full Name** string. The installed `InterVariable.ttf` exposes
weight named-instances as the Full Names below (verify with
`system_profiler SPFontsDataType | grep -i "inter variable"`). These four are the
only font names that should appear in a brand deck:

| PowerPoint family string    | Role                          |
| --------------------------- | ----------------------------- |
| `Inter Variable Thin`       | titles / headlines            |
| `Inter Variable`            | body copy / default           |
| `Inter Variable SemiBold`   | subheads / emphasis           |
| `Inter Variable Medium`     | medium emphasis               |

> Note: do **not** use `Inter Variable Text *` (or the legacy static `Inter Thin`
> / `Inter`). Those names are not exposed by the installed font and PowerPoint
> will fall back to a serif. The string must match an installed Full Name exactly.

These names are the single source of truth, mirrored in
`scripts/pptx_helpers.py` as `BRAND_FONT_TITLE` / `BRAND_FONT_BODY` /
`BRAND_FONT_SEMIBOLD` / `BRAND_FONT_MEDIUM`. If you ever need to reference a font
name in a build, use those constants — never a literal string.

## Default behavior: silently check, then offer to install only if missing

Whenever a powerpoint script is invoked — or when building a presentation —
**silently check whether Inter Variable is installed** before doing anything
else. Do not ask the user if they have it. Just check.

### macOS detection

`fc-list` is **not** installed on stock macOS — do not use it. Check for the
installed variable font directly, either by file or via CoreText:

```bash
# File check (fastest, authoritative for our install path)
test -f ~/Library/Fonts/InterVariable.ttf && echo "Inter Variable installed"

# Or confirm the named instances CoreText exposes (what PowerPoint actually sees)
system_profiler SPFontsDataType | grep -i "inter variable"
```

If `InterVariable.ttf` is present (or `system_profiler` lists the "Inter
Variable" instances), the font is installed. Briefly note "Inter Variable
detected" and move on.

### Windows detection

```powershell
Get-ChildItem "$env:LOCALAPPDATA\Microsoft\Windows\Fonts\*inter*" -ErrorAction SilentlyContinue
Get-ChildItem "$env:WINDIR\Fonts\*inter*" -ErrorAction SilentlyContinue
```

If either returns results, the font is installed.

### If Inter Variable is NOT installed

Only then, offer to install it:

1. **macOS:** Copy the font file directly to the user font directory:

   ```bash
   cp ~/.claude/skills/m1-brand/assets/InterVariable.ttf ~/Library/Fonts/
   ```

   Then confirm installation by re-running the file/`system_profiler` check above.

2. **Windows:** Copy to the user font directory:

   ```powershell
   Copy-Item "$env:USERPROFILE\.claude\skills\m1-brand\assets\InterVariable.ttf" "$env:LOCALAPPDATA\Microsoft\Windows\Fonts\"
   ```

3. **Fallback:** If the copy fails due to permissions, tell the user to
   double-click the font file at `~/.claude/skills/m1-brand/assets/InterVariable.ttf`
   to open the system font installer, or download from
   https://rsms.me/inter/download/.

**Do not prompt the user about the font if it is already installed.** Just move on.

## Font specs: use these sizes & weights whenever possible

Weights below are the variable-axis values. In the deck they resolve to the
named instances: **Thin ≈ 100/200** (`Inter Variable Thin`, titles),
**Regular ≈ 400** (`Inter Variable`, body), **Medium ≈ 500**
(`Inter Variable Medium`), **SemiBold ≈ 600** (`Inter Variable SemiBold`).

### Headline sizes

| Style | Size | Line-height | Preferred Weights       |
|:-----:|:----:|:-----------:| ----------------------- |
| HXL   | 52   | 64          | 200, 300, 400, 500, 600 |
| HL    | 42   | 48          | 300, 500, 600           |
| HM    | 32   | 40          | 300, 500, 600           |
| HS    | 28   | 32          | 300, 500, 600           |
| HXS   | 24   | 28          | 300, 500, 600           |
| HXXS  | 20   | 28          | 300, 500, 600           |

### Body sizes

| Style | Size | Line-height | Preferred Weights  |
|:-----:|:----:|:-----------:| ------------------ |
| PXL   | 18   | 24          | 300, 400, 500, 600 |
| PL    | 16   | 24          | 400, 600           |
| PM    | 14   | 20          | 400, 600           |
| PS    | 12   | 16          | 400, 600           |
| PXS   | 11   | 14          | 400, 600           |

### Rules

- **Line-height:** Always keep line-height a few points larger than font-size.
- **PowerPoint minimum:** Never shrink body text below 16pt. If content doesn't fit, create a new slide.
- **Color:** Headlines use text color by default. Marketing callout headlines can optionally use navy (`--light-blue-05` / `#152B56`).
