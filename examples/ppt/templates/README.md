# PowerPoint Style Templates

Professional presentation style templates for OfficeCLI. Each style includes design guidelines (`style.md`) and reference implementation (`build.sh`).

## ✅ Available Templates (17 with pre-generated PPTs)

Templates marked with ✅ include working build scripts and pre-generated `.pptx` files ready to use.

---

## 🌑 Dark Palette (8 available / 14 total)

| Status | Directory | Style Name | Best For | Mood |
|--------|-----------|------------|----------|------|
| ✅ | dark--investor-pitch | Investor Pitch Pro | Investor pitches, fundraising decks, business plans | Professional, trustworthy |
| ✅ | dark--cosmic-neon | Cosmic Neon | Science talks, futuristic topics, physics, cosmic themes | Sci-fi, mysterious, futuristic |
| ✅ | dark--editorial-story | Editorial Magazine Story | Brand storytelling, editorial magazines, content releases | Narrative, artistic, premium |
| ✅ | dark--tech-cosmos | Tech Cosmos | Tech talks, architecture reviews, scientific presentations | Futuristic, scientific, cosmic |
| ✅ | dark--cyber-future | Cyber Future | Futuristic topics, tech vision, cyberpunk, AI/robotics | Futuristic, cyberpunk |
| ✅ | dark--luxury-minimal | Luxury Minimal | Luxury brands, premium products, high-end corporate | Luxurious, minimalist |
| ✅ | dark--space-odyssey | Space Odyssey | Space/astronomy, science education, exploration narratives | Cosmic, inspiring, epic |
| ✅ | dark--neon-productivity | Neon Productivity | Productivity talks, tech workshops, motivation, startups | Energetic, modern, vibrant |
| ⚙️ | dark--liquid-flow | Liquid Light | Brand upgrades, creative launches, fashion showcases | Fluid, dreamy, avant-garde |
| ⚙️ | dark--premium-navy | Premium Navy & Gold | High-end corporate, annual strategy, board presentations | Authoritative, refined |
| ⚙️ | dark--blueprint-grid | Blueprint Grid | Technical planning, engineering blueprints, system architecture | Precise, professional |
| ⚙️ | dark--diagonal-cut | Diagonal Industrial Cut | Industrial, engineering, construction, manufacturing | Rugged, powerful, bold |
| ⚙️ | dark--spotlight-stage | Spotlight Stage | Keynotes, launch events, TED-style talks, galas | Dramatic, focused |
| ⚙️ | dark--circle-digital | Dark Digital Agency | Digital marketing, creative agencies, tech companies | Modern, dark-cool, digital |
| ⚙️ | dark--architectural-plan | Architectural Plan | Architectural design, business plans, real estate development | Professional, structured |

---

## ☀️ Light Palette (6 available / 8 total)

| Status | Directory | Style Name | Best For | Mood |
|--------|-----------|------------|----------|------|
| ✅ | light--minimal-corporate | Minimal Corporate Report | Annual reports, work summaries, business proposals | Professional, clean |
| ✅ | light--minimal-product | Minimal Product Showcase | Product launches, tech showcases, brand introductions | Modern, minimalist |
| ✅ | light--project-proposal | Project Proposal | Project kickoffs, business proposals, bid presentations | Professional, rigorous |
| ✅ | light--spring-launch | Spring Launch Fresh | Spring launches, new product releases, seasonal marketing | Fresh, natural, vibrant |
| ✅ | light--training-interactive | Interactive Training | Corporate training, online courses, knowledge sharing | Educational, friendly |
| ⚙️ | light--bold-type | Bold Typography | Editorial layouts, magazine-style, brand manuals | Bold, modern, editorial |
| ⚙️ | light--isometric-clean | Isometric Clean Tech | Tech products, SaaS platforms, data presentations | Fresh, modern, techy |
| ⚙️ | light--watercolor-wash | Watercolor Wash | Art, cultural creative, tea ceremony, weddings | Soft, poetic, artistic |

---

## 🧡 Warm Palette (4 available / 5 total)

| Status | Directory | Style Name | Best For | Mood |
|--------|-----------|------------|----------|------|
| ✅ | warm--minimal-brand | Minimal Brand | Brand introductions, product launches, premium brand showcases | Warm, refined, minimalist |
| ✅ | warm--playful-organic | Playful Organic | Lifestyle, pet/animal topics, children's education | Warm, playful, friendly |
| ⚙️ | warm--earth-organic | Earth & Sage | Eco-friendly, sustainability, organic brands | Warm, sincere, natural |
| ✅ | warm--brand-refresh | Brand Refresh | Brand launches, corporate image updates, creative proposals | Fashionable, colorful |
| ✅ | warm--creative-marketing | Creative Marketing | Marketing campaigns, ad creatives, poster-style PPTs | Bold, impactful |

---

## 🌈 Vivid Palette (1 available / 2 total)

| Status | Directory | Style Name | Best For | Mood |
|--------|-----------|------------|----------|------|
| ✅ | vivid--playful-marketing | Vibrant Youth Marketing | Marketing campaigns, new product promos, sales events | Youthful, energetic |
| ⚙️ | vivid--candy-stripe | Rainbow Candy Stripe | Event celebrations, holidays, children's education | Joyful, lively, rainbow |

---

## ⬛ Black & White (1 available / 3 total)

| Status | Directory | Style Name | Best For | Mood |
|--------|-----------|------------|----------|------|
| ✅ | bw--swiss-bauhaus | Swiss Bauhaus | Design agencies, architecture firms, art exhibitions | Rational, rigorous, classic |
| ⚙️ | bw--mono-line | Minimal Line | Minimalist corporate, academic reports, consulting proposals | Calm, restrained |
| ⚙️ | bw--brutalist-raw | Brutalist Raw | Avant-garde art shows, experimental design, indie brands | Rebellious, rugged |

---

## 🎨 Mixed Palette (0 available / 1 total)

| Status | Directory | Style Name | Best For | Mood |
|--------|-----------|------------|----------|------|
| ⚙️ | mixed--duotone-split | Duotone Split | Brand launches, architectural design, premium showcases | Bold, architectural |

---

## 📊 Summary

| Palette | Total Styles | Available (✅) | Reference (⚙️) |
|---------|--------------|----------------|----------------|
| 🌑 Dark | 14 | 8 | 6 |
| ☀️ Light | 8 | 6 | 2 |
| 🧡 Warm | 5 | 4 | 1 |
| 🌈 Vivid | 2 | 1 | 1 |
| ⬛ Black & White | 3 | 1 | 2 |
| 🎨 Mixed | 1 | 0 | 1 |
| **Total** | **34** | **19** | **15** |

---

## 🚀 Quick Start

### Use a Pre-Generated Template

```bash
cd styles/dark--investor-pitch
# View the generated PPT
open template.pptx

# Or regenerate it
bash build.sh
```

### Browse by Use Case

| Use Case | Recommended Styles (✅ = Ready to Use) |
|----------|----------------------------------------|
| **Tech / AI / SaaS** | ✅ dark--tech-cosmos, ✅ dark--cyber-future, ⚙️ light--isometric-clean |
| **Investment / Pitch** | ✅ dark--investor-pitch, ⚙️ dark--premium-navy, ✅ light--project-proposal |
| **Corporate / Business** | ✅ light--minimal-corporate, ✅ light--minimal-product |
| **Brand / Launch / Marketing** | ✅ warm--minimal-brand, ✅ vivid--playful-marketing, ⚙️ warm--brand-refresh |
| **Design / Architecture / Art** | ✅ bw--swiss-bauhaus, ⚙️ bw--brutalist-raw |
| **Education / Training** | ✅ light--training-interactive, ✅ warm--playful-organic |
| **Keynotes / Events** | ⚙️ dark--spotlight-stage, ⚙️ dark--liquid-flow |
| **Sci-Fi / Space / Future** | ✅ dark--space-odyssey, ✅ dark--cosmic-neon, ✅ dark--cyber-future |
| **Luxury / Premium** | ✅ dark--luxury-minimal, ✅ warm--minimal-brand |
| **Productivity / Motivation** | ✅ dark--neon-productivity |

---

## 📖 How to Use

### 1. Browse Styles

Each style directory contains:
- `style.md` - Design philosophy, color palette, techniques
- `build.sh` - Reference implementation script
- `*.pptx` - Pre-generated presentation (✅ styles only)

### 2. Generate from Script

```bash
cd styles/dark--investor-pitch
bash build.sh
```

### 3. Learn from Examples

The build scripts demonstrate:
- Color scheme application
- Shape positioning and morphing
- Layout patterns
- Animation choreography

**Important**: These are **reference examples**, not production-ready code. Some scripts may have layout issues. Use them to understand techniques, then adapt for your needs.

---

## 🔧 About Build Scripts

### Working Scripts (✅)
These scripts successfully generate PPTs and serve as good references for:
- Proper officecli command syntax
- Morph transition techniques
- Color palette implementation
- Shape and layout design

### Reference-Only Scripts (⚙️)
These scripts have path errors or outdated command syntax but are still valuable for:
- Understanding design philosophy
- Color scheme reference
- Layout concept inspiration

**Common errors:**
- Missing template file paths
- References to non-existent directories
- Outdated command syntax

To fix these scripts:
1. Update file paths to current directory
2. Fix command syntax (`--name=` → `--prop name=`)
3. Remove references to missing dependencies

---

## 📚 More Resources

- **[style.md files](styles/)** - Detailed design guidelines for each style
- **[PowerPoint examples](../)** - Basic PPT examples
- **[Complete documentation](../../../SKILL.md)** - Full OfficeCLI reference
- **[All examples](../../)** - Word, Excel, PowerPoint examples

---

## 🤝 Contributing

Want to fix a ⚙️ script or add a new style?

1. Test the script thoroughly
2. Ensure proper error handling
3. Update this README with ✅ status
4. Submit a PR

---

**Legend:**
- ✅ **Available** - Pre-generated PPT included, build script works
- ⚙️ **Reference-Only** - Design reference available, build script needs fixes
