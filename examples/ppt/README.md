# PowerPoint (.pptx) Examples & Templates

Examples and professional style templates for PowerPoint presentation automation.

## 📂 Structure

```
ppt/
├── README.md                          # This file
├── gen-beautiful-pptx.sh              # Basic examples
├── gen-animations-pptx.sh
├── gen-video-pptx.py
├── outputs/                           # Generated examples
└── templates/                         # 35 professional style templates ⭐
    ├── README.md                      # Style index and guide
    └── styles/                        # Individual style directories
        ├── dark--*/                   (14 dark styles, 8 available)
        ├── light--*/                  (8 light styles, 6 available)
        ├── warm--*/                   (5 warm styles)
        ├── vivid--*/                  (2 vivid styles)
        ├── bw--*/                     (3 black & white, 1 available)
        └── mixed--*/                  (1 mixed style)
```

---

## 🚀 Quick Start

### Basic Examples

```bash
# Beautiful presentation with morph transitions
bash gen-beautiful-pptx.sh

# Animation effects
bash gen-animations-pptx.sh

# Video embedding (Python)
python gen-video-pptx.py
```

### Professional Style Templates

```bash
cd templates/styles/dark--investor-pitch
# View pre-generated PPT
open template.pptx

# Or regenerate
bash build.sh
```

👉 **[Browse all 35 styles →](templates/)** (15 with pre-generated PPTs)

---

## 🎨 Basic Scripts

### [gen-beautiful-pptx.sh](gen-beautiful-pptx.sh)
**Create a beautiful presentation with morph transitions**

```bash
bash gen-beautiful-pptx.sh
```

**Demonstrates:**
- Morph transitions between slides
- Shape creation and positioning
- Text styling and alignment
- Color palettes and gradients
- Layout design patterns

**Output:** [`outputs/beautiful_presentation.pptx`](outputs/beautiful_presentation.pptx)

---

### [gen-animations-pptx.sh](gen-animations-pptx.sh)
**Comprehensive animation examples**

```bash
bash gen-animations-pptx.sh
```

**Demonstrates:**
- Entrance animations (fade, fly, zoom, etc.)
- Emphasis animations (pulse, grow, spin)
- Exit animations (disappear, fly out)
- Animation timing and sequencing
- Multiple animations per object

**Output:** [`outputs/gen-animations-pptx.pptx`](outputs/gen-animations-pptx.pptx)

---

### [gen-video-pptx.py](gen-video-pptx.py)
**Embed video in PowerPoint (Python)**

```bash
python gen-video-pptx.py
```

**Demonstrates:**
- Video embedding
- Media positioning
- Python integration with OfficeCLI

**Output:** [`outputs/gen-video-pptx.pptx`](outputs/gen-video-pptx.pptx)

---

## 📈 Sample Outputs

Pre-generated examples in [`outputs/`](outputs/):
- `beautiful_presentation.pptx` - Professional presentation with morph
- `data_presentation.pptx` - Data visualization deck
- `gen-animations-pptx.pptx` - Animation showcase
- `gen-video-pptx.pptx` - Video embedding example

---

## 🎨 Professional Style Templates

**35 design styles organized by color palette:**

### 🌑 Dark Palette (14 styles, 8 available)
Perfect for tech, corporate, and futuristic themes.

**Available (✅):**
- `dark--investor-pitch` - Investor pitches, fundraising decks
- `dark--cosmic-neon` - Science talks, futuristic topics
- `dark--editorial-story` - Brand storytelling, editorial magazines
- `dark--tech-cosmos` - Tech talks, architecture reviews
- `dark--cyber-future` - Futuristic topics, cyberpunk, AI
- `dark--luxury-minimal` - Luxury brands, premium products
- `dark--space-odyssey` - Space/astronomy, science education
- `dark--neon-productivity` - Productivity talks, motivation

**Reference-Only (⚙️):**
- `dark--liquid-flow`, `dark--premium-navy`, `dark--blueprint-grid`,
- `dark--diagonal-cut`, `dark--spotlight-stage`, `dark--circle-digital`

### ☀️ Light Palette (8 styles, 6 available)
Clean and professional for business and product showcases.

**Available (✅):**
- `light--minimal-corporate` - Annual reports, business proposals
- `light--minimal-product` - Product launches, brand introductions
- `light--project-proposal` - Project kickoffs, bid presentations
- `light--spring-launch` - Spring launches, seasonal marketing
- `light--training-interactive` - Corporate training, online courses

### 🧡 Warm Palette (5 styles)
Warm and friendly for lifestyle and organic brands. (Reference-only)

### 🌈 Vivid Palette (2 styles)
Energetic and youthful for marketing campaigns. (Reference-only)

### ⬛ Black & White (3 styles, 1 available)
Minimalist and sophisticated.

**Available (✅):**
- `bw--swiss-bauhaus` - Design agencies, architecture firms

### 🎨 Mixed Palette (1 style)
Bold architectural designs. (Reference-only)

👉 **[Full style index with mood, use cases →](templates/README.md)**

---

## 📖 Quick Lookup by Use Case

| Use Case | Recommended Styles |
|----------|-------------------|
| **Tech / AI / SaaS** | ✅ dark--tech-cosmos, ✅ dark--cyber-future |
| **Investment / Pitch** | ✅ dark--investor-pitch, ✅ light--project-proposal |
| **Corporate / Business** | ✅ light--minimal-corporate, ✅ light--minimal-product |
| **Education / Training** | ✅ light--training-interactive |
| **Sci-Fi / Space / Future** | ✅ dark--space-odyssey, ✅ dark--cosmic-neon |
| **Luxury / Premium** | ✅ dark--luxury-minimal |
| **Design / Architecture** | ✅ bw--swiss-bauhaus |

---

## 🎓 Key Concepts

### Presentation Structure
```
/Presentation
  /slide[1]               # First slide
    /shape[1]             # First shape
    /shape[2]
  /slide[2]
  /master[1]              # Slide master
```

### Common Commands

**Add a slide:**
```bash
officecli add deck.pptx / --type slide \
  --prop layout=blank \
  --prop background=1A1A2E
```

**Add a shape:**
```bash
officecli add deck.pptx /slide[1] --type shape \
  --prop text="Hello World" \
  --prop x=5cm \
  --prop y=5cm \
  --prop width=10cm \
  --prop height=3cm \
  --prop size=48 \
  --prop bold=true \
  --prop color=FFFFFF
```

**Set transition:**
```bash
officecli set deck.pptx /slide[1] \
  --prop transition=morph \
  --prop advanceTime=3000
```

**Copy slide:**
```bash
officecli add deck.pptx / --from /slide[1]
```

---

## 🎨 Shape Types

### Available Presets

| Preset | Description |
|--------|-------------|
| `rect` | Rectangle |
| `roundRect` | Rounded rectangle |
| `ellipse` | Circle/Ellipse |
| `triangle` | Triangle |
| `diamond` | Diamond |
| `pentagon` | Pentagon |
| `hexagon` | Hexagon |
| `star5` | 5-point star |
| `arrow` | Arrow |
| `callout` | Callout bubble |

**View all presets:**
```bash
officecli pptx add
```

---

## 📊 Available Properties

### Slide
- `layout` - Slide layout (blank, title, titleContent, etc.)
- `background` - Background color (hex)
- `transition` - Transition effect (fade, push, wipe, morph, etc.)
- `advanceTime` - Auto-advance time in milliseconds
- `notes` - Speaker notes

### Shape
- `name` - Shape name/identifier
- `preset` - Shape preset (rect, ellipse, arrow, etc.)
- `text` - Text content
- `x`, `y` - Position (cm, in, pt, px, EMU)
- `width`, `height` - Size
- `rotation` - Rotation angle (degrees)
- `fill` - Fill color (hex)
- `line` - Line color (hex or "none")
- `opacity` - Opacity (0.0 to 1.0)

### Text Formatting
- `font` - Font name
- `size` - Font size in points
- `bold` - true/false
- `italic` - true/false
- `color` - Text color (hex)
- `align` - left, center, right, justify
- `valign` - top, middle, bottom

### Animations
- `animation` - Animation effect (fade, fly, zoom, etc.)
- `animDelay` - Delay before animation starts (ms)
- `animDuration` - Animation duration (ms)
- `animTrigger` - click, afterPrev, withPrev

**For complete property list:**
```bash
officecli pptx set
officecli pptx set slide
officecli pptx set shape
```

---

## 🎬 Transitions & Animations

### Popular Transitions
- `morph` - Seamless object morphing
- `fade` - Fade in/out
- `push` - Push from side
- `wipe` - Wipe across
- `zoom` - Zoom in/out
- `cube` - 3D cube rotation

### Animation Types
- **Entrance:** fade, fly, zoom, appear, split
- **Emphasis:** pulse, grow, spin, teeter
- **Exit:** disappear, fly, fade, zoom

**Morph Transition Tips:**
- Name objects consistently across slides (e.g., `name="!!title"`)
- Keep object hierarchy the same
- Change position, size, or color for smooth morphing

---

## 🔧 Tips

1. **View presentation structure:**
   ```bash
   officecli view deck.pptx outline
   ```

2. **Check statistics:**
   ```bash
   officecli view deck.pptx stats
   ```

3. **Query shapes:**
   ```bash
   officecli query deck.pptx "shape[fill=FF0000]"
   ```

4. **Batch slide building:**
   ```bash
   cat << EOF | officecli batch deck.pptx
   [
     {"command":"add","parent":"/","type":"slide","props":{"background":"000000"}},
     {"command":"add","parent":"/slide[1]","type":"shape","props":{"text":"Title","x":"5cm","y":"5cm"}}
   ]
   EOF
   ```

5. **Resident mode for multi-slide decks:**
   ```bash
   officecli open deck.pptx
   officecli add deck.pptx / --type slide
   officecli add deck.pptx / --type slide
   officecli close deck.pptx
   ```

6. **Position units:**
   - `cm` - Centimeters (recommended)
   - `in` - Inches
   - `pt` - Points
   - `px` - Pixels
   - EMU - Raw units (914400 = 1 inch)

---

## 📚 More Resources

- **[Style Templates](templates/)** - 35 professional styles (19 ready-to-use)
- **[PowerPoint documentation](../../SKILL.md#powerpoint-pptx)** - Complete reference
- **[All examples](../)** - Word, Excel, PowerPoint
- **[Word examples](../word/)** - Document automation
- **[Excel examples](../excel/)** - Spreadsheet automation
