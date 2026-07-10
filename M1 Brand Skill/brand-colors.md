---
description: "M1 brand colors and palette. Use when the user mentions: colors, brand colors, M1 colors, color palette, color codes, hex codes, color usage, accessible colors, on-brand colors, or any color guidance for M1 materials."
disable-model-invocation: false
---

# M1 brand colors

## Default behavior: always open the interactive swatches first

For ANY general color question — "show me the colors," "what are our brand colors," "I need to pick a color," "show me the palette," or anything similar — **immediately run this command without asking**:

```bash
open ~/.claude/skills/m1-brand/assets/ColorKeeper-m1-2026.html
```

This opens a clickable, visual color swatch tool in the browser. It is always the right first answer for color questions. Never dump the color list as text for general questions.

## When to use the text palette instead

Only reference the color list below when the user asks for a **specific value** — e.g., "what hex is the primary teal?" or "give me the CSS variable for the navy." For everything else, open the swatches.

#### Color general usage:

For rich colors, dark backgrounds or anything that needs a dark color: --light-blue-05;

For buttons use the primary teal: --light-teal-o4; Text for the button is white.

For light backgrounds dafault to white or --light-blue-01; for a section on white;

For general text use: --light-grey-06;

# All of our colors are listed below

Colors can be queried. "Light" denotes a light mode color and "dark" denotes a dark mode color.

#### CSS Variable Swatches

**Light Mode Primitives**

--light-teal-01
rgb(229 239 244)
#e5eff4

--light-teal-02
rgb(166 201 218)
#a6c9da

--light-teal-03
rgb(111 167 195)
#6fa7c3

--light-teal-04 (primary)
rgb(42 125 167)
#2a7da7

--light-teal-05
rgb(31 91 122)
#1f5b7a

--light-blue-01
rgb(232 237 245)
#e8edf5

--light-blue-02
rgb(204 214 234)
#ccd6ea

--light-blue-03
rgb(154 174 214)
#9aaed6

--light-blue-04
rgb(61 90 148)
#3d5a94

--light-blue-05
rgb(21 43 86)
#152b56

--light-violet-01
rgb(235 237 247)
#ebedf7

--light-violet-02
rgb(205 210 234)
#cdd2ea

--light-violet-03
rgb(169 177 219)
#a9b1db

--light-violet-04
rgb(85 102 185)
#5566b9

--light-violet-05
rgb(36 54 143)
#24368f

--light-red-01
rgb(246 233 235)
#f6e9eb

--light-red-02
rgb(240 220 223)
#f0dcdf

--light-red-03
rgb(208 142 153)
#d08e99

--light-red-04
rgb(179 72 90)
#b3485a

--light-red-05
rgb(119 30 45)
#771e2d

--light-orange-01
rgb(245 236 231)
#f5ece7

--light-orange-02
rgb(230 206 194)
#e6cec2

--light-orange-03
rgb(200 148 122)
#c8947a

--light-orange-04
rgb(168 85 44)
#a8552c

--light-orange-05
rgb(99 50 25)
#633219

--light-yellow-01
rgb(251 244 223)
#fbf4df

--light-yellow-02
rgb(244 224 162)
#f4e0a2

--light-yellow-03
rgb(237 209 123)
#edd17b

--light-yellow-04
rgb(200 162 50)
#c8a232

--light-yellow-05
rgb(122 99 31)
#7a631f

--light-green-01
rgb(233 246 238)
#e9f6ee

--light-green-02
rgb(219 240 226)
#dbf0e2

--light-green-03
rgb(146 210 168)
#92d2a8

--light-green-04
rgb(39 157 79)
#279d4f

--light-green-05
rgb(31 124 62)
#1f7c3e

--light-jade-01
rgb(229 245 242)
#e5f5f2

--light-jade-02
rgb(204 234 229)
#cceae5

--light-jade-03
rgb(125 201 188)
#7dc9bc

--light-jade-04
rgb(39 153 134)
#279986

--light-jade-05
rgb(30 121 106)
#1e796a

--light-grey-00
rgb(255 255 255)
#ffffff

--light-grey-01
rgb(243 245 248)
#f3f5f8

--light-grey-02
rgb(232 236 242)
#e8ecf2

--light-grey-03
rgb(212 219 230)
#d4dbe6

--light-grey-04
rgb(143 155 174)
#8f9bae

--light-grey-05
rgb(84 96 115)
#546073

--light-grey-06
rgb(28 32 38)
#1c2026

--light-gold-01
rgb(247 237 222)
#f7edde

--light-gold-02
rgb(235 210 173)
#ebd2ad

--light-gold-03
rgb(214 165 92)
#d6a55c

--light-gold-04
rgb(187 131 48)
#bb8330

--light-gold-05
rgb(106 74 27)
#6a4a1b

--light-orc-01
rgb(96 108 130)
#606c82

--light-orc-02
rgb(35 38 45)
#23262d

--light-orc-03
rgb(0 0 0)
#000000

**Light Mode Semantic**

--light-fg-neutral-main
var(--light-grey-06)

--light-fg-neutral-secondary
var(--light-grey-05)

--light-fg-neutral-tertiary
var(--light-grey-04)

--light-fg-neutral-inverse
var(--light-grey-00)

--light-fg-neutral-on-dark
var(--light-grey-00)

--light-fg-neutral-on-light
var(--light-grey-06)

--light-fg-primary
var(--light-teal-04)

--light-fg-primary-tint
var(--light-teal-02)

--light-fg-primary-shade
var(--light-teal-05)

--light-fg-secondary
var(--light-blue-05)

--light-fg-secondary-tint
var(--light-blue-03)

--light-fg-promotion
var(--light-gold-03)

--light-fg-promotion-shade
var(--light-gold-05)

--light-fg-orc
var(--light-orc-01)

--light-fg-success
var(--light-green-05)

--light-fg-success-tint
var(--light-green-04)

--light-fg-warning
var(--light-yellow-03)

--light-fg-warning-shade
var(--light-yellow-05)

--light-fg-critical
var(--light-red-04)

--light-bg-neutral-primary
var(--light-grey-00)

--light-bg-neutral-secondary
var(--light-grey-01)

--light-bg-neutral-tertiary
var(--light-grey-02)

--light-bg-elevation-1
var(--light-grey-00)

--light-bg-elevation-2
var(--light-grey-00)

--light-bg-feature-flat
var(--light-blue-05)

--light-bg-shadow
var(--light-grey-05)

--light-bg-primary-subtle
var(--light-teal-01)

--light-bg-info-subtle
var(--light-blue-01)

--light-bg-success-subtle
var(--light-green-01)

--light-bg-warning-subtle
var(--light-yellow-01)

--light-bg-critical-subtle
var(--light-red-01)

--light-bg-promotion-subtle
var(--light-gold-01)

--light-bg-elevation-1
var(--light-grey-00)

--light-bg-elevation-2
var(--light-grey-00)

--light-bg-feature-highlight
var(--light-jade-03)

--light-border-promotion
var(--light-gold-03)

--light-border-main
var(--light-grey-03)

--light-border-active
var(--light-teal-04)

--light-border-critical-active
var(--light-red-04)

--light-border-info
var(--light-blue-03)

--light-border-success
var(--light-green-03)

--light-border-warning
var(--light-yellow-03)

--light-border-critical
var(--light-red-03)

--light-chart-datapoint
var(--light-blue-04)

--light-chart-datapoint-feature
var(--light-blue-03)

--light-chart-bg-plot
var(--light-grey-03)

--light-chart-bg-plot-feature
var(--light-blue-04)

--light-chart-datapoint-1
var(--light-blue-04)

--light-chart-datapoint-2
var(--light-blue-03)

--light-chart-datapoint-3
var(--light-blue-02)

--light-chart-datapoint-4
var(--light-blue-05)

**Dark Mode Primitives**

--dark-teal-01
rgb(13 37 49)
#0d2531

--dark-teal-02
rgb(22 64 85)
#164055

--dark-teal-03
var(--light-teal-04)

--dark-teal-04
var(--light-teal-03)

--dark-teal-05
var(--light-teal-02)

--dark-blue-01
rgb(12 25 49)
#0c1931

--dark-blue-02
var(--light-blue-05)

--dark-blue-03
var(--light-blue-04)

--dark-blue-04
rgb(110 139 196)
#6e8bc4

--dark-blue-05
rgb(183 197 225)
#b7c5e1

--dark-violet-01
rgb(16 25 65)
#101941

--dark-violet-02
rgb(26 38 102)
#1a2666

--dark-violet-03
var(--light-violet-04)

--dark-violet-04
rgb(120 133 199)
#7885c7

--dark-violet-05
rgb(188 194 227)
#bcc2e3

--dark-red-01
rgb(49 12 19)
#310c13

--dark-red-02
var(--light-red-05)

--dark-red-03
var(--light-red-04)

--dark-red-04
rgb(196 113 127)
#c4717f

--dark-red-05
rgb(218 165 174)
#daa5ae

--dark-orange-01
rgb(49 25 12)
#31190c

--dark-orange-02
var(--light-orange-05)

--dark-orange-03
var(--light-orange-04)

--dark-orange-04
rgb(186 122 89)
#ba7a59

--dark-orange-05
rgb(221 189 173)
#ddbdad

--dark-yellow-01
rgb(40 36 21)
#282415

--dark-yellow-02
rgb(71 63 37)
#473f25

--dark-yellow-03
rgb(122 107 63)
#7a6b3f

--dark-yellow-04
rgb(220 194 114)
#dcc272

--dark-yellow-05
var(--light-yellow-02)

--dark-green-01
rgb(12 49 25)
#0c3119

--dark-green-02
rgb(18 73 37)
#124925

--dark-green-03
var(--light-green-05)

--dark-green-04
var(--light-green-04)

--dark-green-05
rgb(205 234 214)
#cdead6

--dark-jade-01
rgb(12 49 43)
#0c312b

--dark-jade-02
rgb(18 71 62)
#12473e

--dark-jade-03
var(--light-jade-05)

--dark-jade-04
rgb(66 177 158)
#42b19e

--dark-jade-05
rgb(172 221 212)
#acddd4

--dark-grey-00
rgb(15 17 21)
#0f1115

--dark-grey-01
rgb(26 30 35)
#1a1e23

--dark-grey-02
rgb(39 44 53)
#272c35

--dark-grey-03
rgb(47 54 65)
#2f3641

--dark-grey-04
rgb(94 109 130)
#5e6d82

--dark-grey-05
var(--light-grey-04)

--dark-grey-06
rgb(236 239 244)
#eceff4

--dark-gold-01
rgb(54 43 28)
#362b1c

--dark-gold-02
var(--light-gold-05)

--dark-gold-03
var(--light-gold-04)

--dark-gold-04
var(--light-gold-03)

--dark-gold-05
var(--light-gold-02)

--dark-orc-01
var(--light-orc-03)

--dark-orc-02
var(--light-orc-02)

--dark-orc-03
var(--light-orc-01)
Dark Mode Semantic

--dark-teal-03
var(--light-teal-04)

--dark-teal-04
var(--light-teal-03)

--dark-teal-05
var(--light-teal-02)

--dark-blue-02
var(--light-blue-05)

--dark-blue-03
var(--light-blue-04)

--dark-violet-03
var(--light-violet-04)

--dark-red-02
var(--light-red-05)

--dark-red-03
var(--light-red-04)

--dark-orange-02
var(--light-orange-05)

--dark-orange-03
var(--light-orange-04)

--dark-yellow-05
var(--light-yellow-02)

--dark-green-03
var(--light-green-05)

--dark-green-04
var(--light-green-04)

--dark-jade-03
var(--light-jade-05)

--dark-grey-05
var(--light-grey-04)

--dark-gold-02
var(--light-gold-05)

--dark-gold-03
var(--light-gold-04)

--dark-gold-04
var(--light-gold-03)

--dark-gold-05
var(--light-gold-02)

--dark-orc-01
var(--light-orc-03)

--dark-orc-02
var(--light-orc-02)

--dark-orc-03
var(--light-orc-01)

--dark-achromatic-pure-black
var(--light-orc-03)

--dark-achromatic-pure-white
var(--light-grey-00)

--dark-fg-neutral-main
var(--dark-grey-06)

--dark-fg-neutral-secondary
var(--light-grey-04)

--dark-fg-neutral-tertiary
var(--dark-grey-04)

--dark-fg-neutral-inverse
var(--dark-grey-00)

--dark-fg-neutral-on-dark
var(--dark-grey-06)

--dark-fg-neutral-on-light
var(--dark-grey-00)

--dark-fg-primary
var(--light-teal-03)

--dark-fg-primary-tint
var(--light-teal-03)

--dark-fg-primary-shade
var(--light-teal-04)

--dark-fg-secondary
var(--light-blue-05)

--dark-fg-secondary-tint
var(--light-blue-03)

--dark-fg-promotion
var(--light-gold-03)

--dark-fg-promotion-shade
var(--light-gold-05)

--dark-fg-success
var(--light-green-05)

--dark-fg-success-tint
var(--light-green-04)

--dark-fg-warning
var(--light-yellow-03)

--dark-fg-warning-shade
var(--light-yellow-05)

--dark-fg-critical
var(--light-red-04)

--dark-fg-orc
var(--light-orc-01)

--dark-bg-neutral-primary
var(--dark-grey-00)

--dark-bg-neutral-secondary
var(--dark-grey-01)

--dark-bg-neutral-tertiary
var(--dark-grey-02)

--dark-bg-feature-flat
var(--dark-blue-01)

--dark-bg-primary-subtle
var(--dark-teal-01)

--dark-bg-info-subtle
var(--dark-blue-01)

--dark-bg-success-subtle
var(--dark-green-01)

--dark-bg-warning-subtle
var(--dark-yellow-01)

--dark-bg-critical-subtle
var(--dark-red-01)

--dark-bg-shadow
var(--dark-grey-00)

--dark-bg-promotion-subtle
var(--dark-gold-01)

--dark-bg-feature-highlight
var(--dark-jade-05)

--dark-border-main
var(--dark-grey-03)

--dark-border-active
var(--light-teal-04)

--dark-border-critical-active
var(--light-red-04)

--dark-border-info
var(--light-blue-03)

--dark-border-success
var(--light-green-03)

--dark-border-warning
var(--light-yellow-03)

--dark-border-critical
var(--light-red-03)

--dark-border-m1-plus
var(--light-gold-03)

--dark-chart-datapoint
var(--dark-blue-04)

--dark-chart-datapoint-feature
var(--dark-blue-04)

--dark-chart-bg-plot
var(--dark-grey-03)

--dark-chart-datapoint-1
var(--dark-blue-05)

--dark-chart-datapoint-2
var(--dark-blue-04)

--dark-chart-datapoint-3
var(--light-blue-04)

--dark-chart-datapoint-4
var(--dark-blue-01)