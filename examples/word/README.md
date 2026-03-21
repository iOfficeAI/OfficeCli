# Word (.docx) Examples

Examples demonstrating OfficeCLI capabilities for Word document automation.

## 📄 Scripts

### [gen-formulas.sh](gen-formulas.sh)
**Insert mathematical formulas and equations**

```bash
bash gen-formulas.sh
```

**Demonstrates:**
- LaTeX math formula support
- Equation insertion
- Formula formatting

**Output:** [`outputs/complex_formulas.docx`](outputs/complex_formulas.docx)

---

### [gen-complex-tables.sh](gen-complex-tables.sh)
**Generate complex tables with styling**

```bash
bash gen-complex-tables.sh
```

**Demonstrates:**
- Table creation and formatting
- Cell styling (borders, shading, alignment)
- Row and column manipulation
- Table properties (width, height, spacing)

**Output:** [`outputs/complex_tables.docx`](outputs/complex_tables.docx)

---

### [gen-complex-textbox.sh](gen-complex-textbox.sh)
**Create styled text boxes**

```bash
bash gen-complex-textbox.sh
```

**Demonstrates:**
- Text box creation
- Font styling (bold, italic, size, color)
- Text alignment and formatting
- Paragraph properties

**Output:** Generated dynamically

---

## 🎓 Key Concepts

### Document Structure
```
/document
  /body
    /p[1]           # Paragraph 1
      /r[1]         # Run 1
    /p[2]
    /tbl[1]         # Table 1
      /tr[1]        # Row 1
        /tc[1]      # Cell 1
```

### Common Commands

**Create a paragraph:**
```bash
officecli add report.docx /body --type paragraph \
  --prop text="Hello World" \
  --prop style=Heading1
```

**Modify text formatting:**
```bash
officecli set report.docx /body/p[1]/r[1] \
  --prop bold=true \
  --prop color=FF0000 \
  --prop size=24
```

**Add a table:**
```bash
officecli add report.docx /body --type table \
  --prop rows=3 \
  --prop cols=4
```

---

## 📊 Available Properties

### Paragraph
- `text` - Paragraph text content
- `style` - Paragraph style (Normal, Heading1-9, etc.)
- `alignment` - left, center, right, justify
- `lineSpacing` - Line spacing (e.g., 1.5, 2.0)
- `indent` - Indentation in points

### Run (Text Formatting)
- `text` - Text content
- `bold` - true/false
- `italic` - true/false
- `underline` - true/false
- `strike` - true/false
- `font` - Font name
- `size` - Font size in points
- `color` - Hex color (e.g., FF0000)
- `highlight` - Highlight color

### Table
- `rows` - Number of rows
- `cols` - Number of columns
- `width` - Table width
- `border.color` - Border color
- `border.width` - Border width
- `border.style` - Border style

**For complete property list:**
```bash
officecli docx set
officecli docx set paragraph
officecli docx set run
officecli docx set table
```

---

## 🔧 Tips

1. **View structure first:**
   ```bash
   officecli view report.docx outline
   ```

2. **Check content:**
   ```bash
   officecli view report.docx text
   ```

3. **Query elements:**
   ```bash
   officecli query report.docx "paragraph[style=Heading1]"
   ```

4. **Batch operations:**
   ```bash
   cat << EOF | officecli batch report.docx
   [
     {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Para 1"}},
     {"command":"add","parent":"/body","type":"paragraph","props":{"text":"Para 2"}}
   ]
   EOF
   ```

5. **Validate after changes:**
   ```bash
   officecli validate report.docx
   ```

---

## 📚 More Resources

- [Complete Word documentation](../../SKILL.md#word-docx)
- [All examples](../)
- [PowerPoint examples](../powerpoint/)
- [Excel examples](../excel/)
