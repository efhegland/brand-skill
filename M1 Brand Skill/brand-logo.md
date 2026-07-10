---
description: "M1 brand logo(s), including sub-brand logos. Use when the user mentions: logo, brand logo, M1 logo, logomark, wordmark, brand mark, M1 Intelligence logo, M1 Advisor logo, or any request for logo files or logo usage."
disable-model-invocation: false
---

# M1 brand logo

The M1 logo files live in SharePoint, not in this skill. **All M1 logo variants and sub-brand logos live at the same SharePoint location** — this is not limited to the primary M1 mark. That includes (but isn't limited to) the added/sub-brand logos:

- **M1 Intelligence**
- **M1 Advisor**

For ANY logo question — the primary M1 mark, a sub-brand mark like M1 Intelligence or M1 Advisor, "where's the logo," "send me the logo," "I need the M1 mark," or anything similar — **open or suggest this link**:

```
https://m1financial.sharepoint.com/:f:/r/marketing/Shared%20Documents/Creative%20Assets/_Brand%20Assets/0_Logo?csf=1&web=1&e=hRc1ZN
```

On macOS, open it directly:

```bash
open "https://m1financial.sharepoint.com/:f:/r/marketing/Shared%20Documents/Creative%20Assets/_Brand%20Assets/0_Logo?csf=1&web=1&e=hRc1ZN"
```

If the environment can't open a browser, share the link as text instead.

## Local fallback

If SharePoint is unreachable, or the user just needs a logo file to drop straight into a deck or doc, use the bundled fallback instead of blocking on the link:

```
~/.claude/skills/m1-brand/assets/M1logo-RGB-Blu.svg
```

This is the primary M1 logomark in navy (`#152b56`), RGB color space. Mention it's a fallback covering only this one variant — for other colorways (white/reversed, black, other formats), sub-brand marks (M1 Intelligence, M1 Advisor), or official source files, the SharePoint folder above is still the authority.
