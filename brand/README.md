# PExL brand assets

Logos and brand guidance for PExL. Everything here is vector (SVG) so it scales to
any size; rasterize to PNG/ICO as needed (see [Converting](#converting-to-pngicojpg)).

## The mark

The logo is the language's signature operator — the pipe **`|>`** ("feed this into the
next step") — set in a rounded square, the universal "this is an app/add-in" cue. The
green→blue gradient is the same one used for the site's hero title, so the logo and the
product read as one brand.

## Files

| File | What it is | Use it for |
| ---- | ---------- | ---------- |
| `pexl-icon.svg` | Gradient tile, white `\|>` | Primary app icon, ribbon button, store/listing icon, light surfaces |
| `pexl-icon-dark.svg` | Dark tile, gradient `\|>` | Favicons, dark UI, anywhere the bright tile is too loud |
| `pexl-mark.svg` | Bare gradient `\|>` glyph (transparent) | Inline glyph, watermarks, monochrome contexts |
| `pexl-wordmark.svg` | Icon + “PExL” + tagline (horizontal) | README headers, docs, slides, email signatures |
| `pexl-banner.svg` | 1280×640 dark hero card | GitHub social preview, release headers, slide backgrounds |

## Palette

| Token | Hex | Role |
| ----- | --- | ---- |
| Green | `#4ec9b0` | Gradient start / primary |
| Blue  | `#569cd6` | Gradient end / accent |
| Ink   | `#0d1117` | Dark background |
| Panel border | `#2a3240` | Hairline on dark tiles |
| Muted | `#8b98a9` | Secondary text |
| Light | `#e6edf3` | Primary text on dark |

Gradient: `linear-gradient(135deg, #4ec9b0 → #569cd6)`.

## Typography

- **Wordmark / UI:** Segoe UI (system-ui fallback), weight 800 for "PExL".
- **Code / motto:** Cascadia Code / JetBrains Mono / Consolas, ligatures off.

## Clear space & sizing

- Keep padding of at least the tile's corner radius around the logo.
- The icon stays legible down to **16px**; below that, prefer `pexl-mark.svg`.
- Don't recolor the gradient, stretch the tile, or rotate the `|>`.

## Converting to PNG/ICO/JPG

The SVGs are the source of truth. To produce raster assets for the .xll release,
Windows icons, or web favicons, use any one of these.

**Inkscape** (best fidelity):

```bash
inkscape pexl-icon.svg -w 256 -h 256 -o pexl-icon-256.png
inkscape pexl-icon.svg -w 32  -h 32  -o pexl-icon-32.png
```

**ImageMagick** (quick):

```bash
magick -background none pexl-icon.svg -resize 256x256 pexl-icon-256.png
# multi-resolution .ico for Windows:
magick -background none pexl-icon.svg -define icon:auto-resize=16,32,48,256 pexl-icon.ico
```

**Node (no global install):**

```bash
npx svgexport pexl-icon.svg pexl-icon-256.png 256:256
```

Recommended raster set if you add one: `16, 32, 48, 64, 128, 256` PNGs plus a
combined `pexl-icon.ico`.

> Note: the in-product **ribbon icon is drawn at runtime** by `BrandAssets.cs`
> (GDI+) to match `pexl-icon.svg`, so the add-in needs no bundled PNG. These files
> are for distribution, the website, GitHub, and any external tooling.
