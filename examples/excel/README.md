# Excel (.xlsx) Examples

Examples demonstrating OfficeCLI capabilities for Excel spreadsheet automation.

## 📊 Scripts

### [gen-beautiful-charts.sh](gen-beautiful-charts.sh)
**Create professional charts with custom styling**

```bash
bash gen-beautiful-charts.sh
```

**Demonstrates:**
- Multiple chart types (bar, line, pie, scatter, area)
- Chart styling and colors
- Data series configuration
- Legend and axis formatting
- Chart positioning

**Output:** [`outputs/beautiful_charts.xlsx`](outputs/beautiful_charts.xlsx)

---

### [gen-charts-demo.sh](gen-charts-demo.sh)
**Comprehensive chart examples**

```bash
bash gen-charts-demo.sh
```

**Demonstrates:**
- 14+ chart types
- Chart variations (stacked, clustered, 3D)
- Data range selection
- Title and label configuration
- Chart layout

**Output:** [`outputs/charts_demo.xlsx`](outputs/charts_demo.xlsx)

---

## 📈 Sample Output

[`outputs/sales_report.xlsx`](outputs/sales_report.xlsx) - Pre-generated sales report example

---

## 🎓 Key Concepts

### Workbook Structure
```
/Workbook
  /Sheet1
    /A1           # Cell A1
    /B2           # Cell B2
    /A1:C10       # Range A1 to C10
  /Sheet2
  /Chart1         # Chart objects
```

### Common Commands

**Set cell value:**
```bash
officecli set data.xlsx /Sheet1/A1 \
  --prop value="Revenue" \
  --prop bold=true \
  --prop size=14
```

**Add formula:**
```bash
officecli set data.xlsx /Sheet1/B10 \
  --prop formula="=SUM(B2:B9)" \
  --prop numFmt="$#,##0.00"
```

**Create chart:**
```bash
officecli add data.xlsx /Sheet1 --type chart \
  --prop chartType=column \
  --prop dataRange="A1:B10" \
  --prop title="Sales Report"
```

**Add sheet:**
```bash
officecli add data.xlsx / --type sheet \
  --prop name="Q2 Data"
```

---

## 📊 Chart Types

### Supported Chart Types

| Type | Description | Usage |
|------|-------------|-------|
| `column` | Vertical bar chart | Comparing values |
| `bar` | Horizontal bar chart | Ranking data |
| `line` | Line chart | Trends over time |
| `pie` | Pie chart | Part-to-whole |
| `scatter` | Scatter plot | Correlation |
| `area` | Area chart | Volume over time |
| `doughnut` | Doughnut chart | Part-to-whole with center |
| `radar` | Radar chart | Multivariate data |
| `combo` | Combination chart | Multiple series types |

**Variations:**
- `columnStacked` - Stacked columns
- `columnClustered` - Grouped columns
- `column3D` - 3D columns
- `lineMarkers` - Line with data points
- `areaStacked` - Stacked areas

**View all chart types:**
```bash
officecli xlsx add
```

---

## 📊 Available Properties

### Cell
- `value` - Cell value (text or number)
- `formula` - Excel formula (e.g., =SUM(A1:A10))
- `bold` - true/false
- `italic` - true/false
- `size` - Font size
- `font` - Font name
- `color` - Text color (hex)
- `fill` - Background color (hex)
- `numFmt` - Number format (e.g., "0.00%", "$#,##0.00")
- `alignment` - horizontal, vertical
- `border` - Border style

### Chart
- `chartType` - Chart type (column, bar, line, pie, etc.)
- `dataRange` - Data range (e.g., "A1:B10")
- `title` - Chart title
- `x` - X position
- `y` - Y position
- `width` - Chart width
- `height` - Chart height
- `legend` - Legend position (top, bottom, left, right, none)

### Sheet
- `name` - Sheet name
- `tabColor` - Tab color (hex)
- `hidden` - true/false

**For complete property list:**
```bash
officecli xlsx set
officecli xlsx set cell
officecli xlsx set chart
```

---

## 🔧 Tips

1. **View data:**
   ```bash
   officecli view data.xlsx text --cols A,B,C --max-lines 50
   ```

2. **Check formulas:**
   ```bash
   officecli view data.xlsx issues --type content
   ```

3. **Query cells:**
   ```bash
   officecli query data.xlsx "cell[formula~=SUM]"
   ```

4. **Batch cell updates:**
   ```bash
   cat << EOF | officecli batch data.xlsx
   [
     {"command":"set","path":"/Sheet1/A1","props":{"value":"Name","bold":"true"}},
     {"command":"set","path":"/Sheet1/B1","props":{"value":"Score","bold":"true"}},
     {"command":"set","path":"/Sheet1/A2","props":{"value":"Alice"}},
     {"command":"set","path":"/Sheet1/B2","props":{"value":"95"}}
   ]
   EOF
   ```

5. **Number formats:**
   - Currency: `"$#,##0.00"`
   - Percentage: `"0.00%"`
   - Date: `"yyyy-mm-dd"`
   - Custom: `"#,##0.00;[Red]-#,##0.00"`

---

## 📚 More Resources

- [Complete Excel documentation](../../SKILL.md#excel-xlsx)
- [All examples](../)
- [Word examples](../word/)
- [PowerPoint examples](../powerpoint/)
