# Hebrew DOS Formatting Codes - Analysis Results

## Pattern Analysis from Actual Files

### Code Frequency

| Code | Occurrences | Likely Meaning |
|------|-------------|----------------|
| `>99<` | 121 times | **Footnote/Reference marker** - Most common citation |
| `>55<` | 128 times | **Footnote/Reference marker** - Most common citation |
| `>77<` | 21 times | **Footnote/Reference marker** - Frequent citation |
| `>11<` | 17 times | **Footnote/Reference marker** - Section/chapter reference |
| `>31<` | 9 times | **Cross-reference group** (appears with 34, 44, 66, 77) |
| `>34<` | 9 times | **Cross-reference group** |
| `>44<` | 9 times | **Cross-reference group** |
| `>66<` | 9 times | **Cross-reference group** |
| `>42<` | 1 time | **Rare reference** |
| `>62<` | 1 time | **Rare reference** |

## What These Codes Mean

### These are **CONTENT**, not formatting codes!

The `>number<` patterns are **footnote and reference markers** embedded in the Torah text. They indicate:

1. **Source citations** - References to Talmud pages, other texts
2. **Footnotes** - Notes and commentary references
3. **Cross-references** - Links to other parts of the document

### Examples from Analysis:

```
Line 35: >62<>31<]שי"א[ >66<>34<הכל >44<>99<)שכמש ד"ה( >77<חייב לכל
```
Translation: Multiple references (62, 31, 66, 34, 44, 99, 77) marking different citations

```
Line 40: >99<)ש' ד"ה( >55<דוקא בראשון, לפירוש משנה עבדו
```
Translation: References 99 and 55 marking citations within the text

## Coding Decision

### ✅ KEEP These Codes

These are **essential content** that should be preserved in the output! They are:
- Part of the scholarly apparatus
- Critical for understanding the sources
- Meaningful reference numbers

### Current Implementation

The cleaning function now:
1. **Protects** `>number<` patterns during cleaning
2. **Removes** actual formatting codes like:
   - `.` prefix lines (formatting commands)
   - `BNARF B XX*`, `OISAR M XX*` (reference system markers)
   - `<>number` or `number<>` (coordinate/position codes)
   - Standalone decimal numbers (550.0, 6.31, etc.)

## Hebrew DOS Word Processor Context

These files likely came from:
- **Dagesh** word processor
- **ChiWriter** with Hebrew support
- **Hebrew WinWord** DOS version
- Other Hebrew DOS text processors

The `>number<` notation was a common way to embed:
- Superscript references
- Footnote markers
- Cross-references

Without the actual footnote/reference table, we preserve the markers as-is so readers know where citations belong.

## Recommendation

**Keep all `>number<` patterns in the output** - they are scholarly content, not noise!

Users can:
1. Keep them as reference markers
2. Convert them to superscript in final formatting
3. Link them to an actual reference table if available

